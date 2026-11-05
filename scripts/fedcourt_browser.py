"""Fetch Federal Court judgment documents through a real Chrome.

judgments.fedcourt.gov.au (and the www.fedcourt.gov.au file store its
"Original Word Document" links redirect to) sit behind Cloudflare's managed
challenge, so a plain HTTP client only ever gets the "Just a moment..."
interstitial. This drives a real Chrome via nodriver, headful under Xvfb, to
clear the challenge and then fetches the files from inside the page.

Adapted from nwbort/tribunal-tracker (scrape.py, parse.py), which uses the
same method to mirror the Competition Tribunal's decisions:

  www.judgments.fedcourt.gov.au/judgments/Judgments/fca/single/2023/2023fca0665
    -> "Original Word Document" -> .../2023fca0665/2023fca0665
    -> 302 to www.fedcourt.gov.au/file-store/Judgments/.../2023FCA0665.docx

Both hosts have their own challenge and the redirect crosses between them,
so a fetch() from the judgment page is refused (CORS); the destination is
read from the browser's network events, its host's challenge is cleared by
loading its home page, and the file is fetched from there.
"""

import asyncio
import base64
import binascii
import json
import os
import re
import socket
import subprocess
import tempfile
import time
import urllib.request
from html import unescape
from html.parser import HTMLParser
from shutil import which
from urllib.parse import unquote, urljoin, urlparse

import nodriver as uc

JUDGMENT_DOMAIN = "fedcourt.gov.au"

# Markers that indicate we're still looking at the Cloudflare challenge rather
# than the real page.
CHALLENGE_MARKERS = (
    "Just a moment",
    "challenge-platform",
    "cf_chl_opt",
    "Enable JavaScript and cookies to continue",
    "Verifying you are human",
)

MAX_WAIT_SECONDS = 90
POLL_SECONDS = 0.5
TURNSTILE_CLICK_INTERVAL = 3
# Ceiling on any single nodriver/CDP call, so a stalled round-trip can't hang
# the run until Actions kills it.
CDP_CALL_TIMEOUT_SECONDS = 30
TURNSTILE_FIND_TIMEOUT_SECONDS = 5
REQUEST_TIMEOUT = 60

# Cloudflare decides per request: a 403 (cf-mitigated: challenge) one moment
# often goes straight through a few seconds later. 404/410 aren't retried.
BROWSER_FETCH_ATTEMPTS = 3
BROWSER_FETCH_RETRY_SECONDS = 5
RETRYABLE_FETCH_STATUSES = frozenset({403, 408, 425, 429, 500, 502, 503, 504})

# How long to wait for a judgment page to render its file link, and for the
# browser to report where that link redirects.
JUDGMENT_LINK_WAIT_SECONDS = 15
JUDGMENT_REDIRECT_WAIT_SECONDS = 20

# What a judgment's file link says on its page. The Word document is the
# original; anything else offered is a rendering of it.
_JUDGMENT_FILE_LINK_RE = re.compile(r"\b(?:word document|pdf|rtf)\b", re.IGNORECASE)

_SNIFFED_EXTENSIONS = (
    (b"PK\x03\x04", ".docx"),
    (b"%PDF", ".pdf"),
    (b"{\\rtf", ".rtf"),
    (b"\xd0\xcf\x11\xe0", ".doc"),
)
JUDGMENT_EXTENSIONS = {".docx", ".doc", ".pdf", ".rtf"}

# nodriver's default args, which help the browser look like a normal user
# session rather than an automated one.
CHROME_ARGS = [
    "--remote-allow-origins=*",
    "--no-first-run",
    "--no-service-autorun",
    "--no-default-browser-check",
    "--homepage=about:blank",
    "--no-pings",
    "--password-store=basic",
    "--disable-infobars",
    "--disable-breakpad",
    "--disable-dev-shm-usage",
    "--disable-session-crashed-bubble",
    "--disable-search-engine-choice-screen",
    "--disable-features=IsolateOrigins,site-per-process",
    "--disable-gpu",
    "--window-size=1920,1080",
    "--no-sandbox",  # CI runs as root
]


def log(msg: str) -> None:
    print(msg, flush=True)


# --- URL / HTML helpers -------------------------------------------------------


def is_judgment_host_url(url: str) -> bool:
    """True if url is https on the Federal Court's domain or a subdomain."""
    parsed = urlparse(url)
    host = parsed.hostname
    return (
        parsed.scheme == "https"
        and host is not None
        and (host == JUDGMENT_DOMAIN or host.endswith("." + JUDGMENT_DOMAIN))
    )


def safe_filename(name: str) -> str | None:
    name = re.sub(r"[^\w.()-]+", "-", name).strip("-.")
    return name or None


class _LinkParser(HTMLParser):
    """Collect (href, text) for every link on a page."""

    def __init__(self):
        super().__init__()
        self.links: list[tuple[str, str]] = []
        self._href: str | None = None
        self._text: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self._href = dict(attrs).get("href")
            self._text = []

    def handle_endtag(self, tag):
        if tag == "a" and self._href is not None:
            text = re.sub(r"\s+", " ", unescape("".join(self._text))).strip()
            self.links.append((self._href, text))
            self._href = None

    def handle_data(self, data):
        if self._href is not None:
            self._text.append(data)


def judgment_file_link(html: str, page_url: str) -> str | None:
    """The judgment's file link on its page, preferring Word. The file sits
    beneath the judgment's own path (.../2023fca0665/2023fca0665)."""
    parser = _LinkParser()
    parser.feed(html)
    base_path = urlparse(page_url).path.rstrip("/").lower() + "/"
    candidates = []
    for href, text in parser.links:
        href = urljoin(page_url, href)
        if not is_judgment_host_url(href):
            continue
        if not urlparse(href).path.lower().startswith(base_path):
            continue
        if _JUDGMENT_FILE_LINK_RE.search(text):
            candidates.append((0 if "word" in text.lower() else 1, href))
    return min(candidates)[1] if candidates else None


def looks_like_challenge(html: str) -> bool:
    return any(marker in html for marker in CHALLENGE_MARKERS)


def is_challenge_payload(data: bytes) -> bool:
    return b"Just a moment" in data[:2048] or b"challenge-platform" in data[:8192]


def sniff_extension(data: bytes) -> str | None:
    return next((e for magic, e in _SNIFFED_EXTENSIONS if data.startswith(magic)), None)


def judgment_filename(final_url: str, data: bytes) -> str | None:
    """Local filename: the file-store basename (2023FCA0665.docx), with an
    extension sniffed from the bytes when the URL doesn't carry one."""
    name = os.path.basename(unquote(urlparse(final_url).path)).strip()
    stem, ext = os.path.splitext(name)
    if ext.lower() not in JUDGMENT_EXTENSIONS:
        ext = sniff_extension(data)
        if not ext:
            return None
        stem = name
    return safe_filename(stem + ext)


def follow_redirects(url: str, hops: dict[str, str]) -> str:
    seen = {url}
    while url in hops and hops[url] not in seen:
        url = hops[url]
        seen.add(url)
    return url


# --- Browser plumbing ---------------------------------------------------------


async def with_timeout(coro, what: str, seconds: float = CDP_CALL_TIMEOUT_SECONDS):
    try:
        return await asyncio.wait_for(coro, timeout=seconds)
    except asyncio.TimeoutError:
        raise TimeoutError(f"{what} did not respond within {seconds}s")


def find_chrome() -> str:
    env = os.environ.get("CHROME_PATH")
    if env and os.path.exists(env):
        return env
    for candidate in ("google-chrome", "google-chrome-stable", "chromium-browser", "chromium"):
        path = which(candidate)
        if path:
            return path
    raise FileNotFoundError("Could not find a Chrome/Chromium binary")


def free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def wait_for_devtools(port: int, timeout: float = 30.0) -> bool:
    """nodriver's own launcher only waits ~2.5s for the DevTools port, which
    loses a race against Chrome's cold start on CI runners."""
    deadline = time.time() + timeout
    url = f"http://127.0.0.1:{port}/json/version"
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=2) as r:
                log(f"DevTools ready: {json.load(r).get('Browser')}")
                return True
        except Exception:
            time.sleep(0.5)
    return False


class Browser:
    """async context manager: launches Chrome and attaches nodriver."""

    def __init__(self):
        self.proc = None
        self.browser = None

    async def __aenter__(self):
        chrome_path = find_chrome()
        port = free_port()
        args = [
            chrome_path,
            *CHROME_ARGS,
            f"--user-data-dir={tempfile.mkdtemp(prefix='cf-judgments-')}",
            "--remote-debugging-host=127.0.0.1",
            f"--remote-debugging-port={port}",
        ]
        # Chrome ignores the proxy environment variables (not needed on
        # GitHub's runners).
        proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
        if proxy:
            args.append(f"--proxy-server={proxy}")
        args.append("about:blank")
        log(f"Launching Chrome: {chrome_path} (port {port})")
        self.proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if not wait_for_devtools(port):
            self.proc.terminate()
            raise RuntimeError("Chrome DevTools endpoint never became ready")
        self.browser = await uc.start(
            host="127.0.0.1", port=port, browser_executable_path=chrome_path
        )
        return self.browser

    async def __aexit__(self, *exc):
        for stop in (lambda: self.browser and self.browser.stop(),
                     lambda: self.proc and self.proc.terminate()):
            try:
                stop()
            except Exception:
                pass


async def try_click_turnstile(tab):
    """Best-effort click of the Turnstile / 'Verify you are human' checkbox."""
    for text in ("Verify you are human", "Verify you are a human", "human",
                 "challenges.cloudflare.com"):
        try:
            el = await with_timeout(
                tab.find(text, best_match=True, timeout=3),
                "tab.find",
                TURNSTILE_FIND_TIMEOUT_SECONDS,
            )
            if el:
                await with_timeout(el.mouse_click(), "mouse_click")
                log(f"  clicked element matching '{text}'")
                return True
        except Exception:
            pass
    return False


async def fetch_page(browser, url: str, tab=None, same_host: bool = False):
    """Load url and wait for the Cloudflare challenge to clear. Returns
    (tab, html), or (tab, None) if the challenge never cleared. With
    same_host, any page on url's host counts as arrived."""
    log(f"Navigating to {url}")
    try:
        if tab is None:
            tab = await with_timeout(browser.get(url), "browser.get")
        else:
            await with_timeout(tab.get(url), "tab.get")
    except Exception as e:
        log(f"  navigation failed: {e}")
        return tab, None

    target = url.split("?")[0].split("#")[0].rstrip("/").lower()
    if same_host:
        parsed = urlparse(url)
        target = f"{parsed.scheme}://{parsed.netloc}".lower()
    deadline = time.time() + MAX_WAIT_SECONDS
    next_click = time.time() + TURNSTILE_CLICK_INTERVAL
    attempt = 0
    last_title = None
    while time.time() < deadline:
        attempt += 1
        try:
            state = await with_timeout(
                tab.evaluate("document.readyState + ' ' + location.origin + location.pathname"),
                "tab.evaluate",
            )
            ready, _, href = str(state).partition(" ")
            href = href.rstrip("/").lower()
            on_page = href.startswith(target) if same_host else href == target
            html = (
                await with_timeout(tab.get_content(), "tab.get_content")
                if on_page and ready == "complete"
                else None
            )
        except Exception as e:
            log(f"  page check failed: {e}")
            await tab.sleep(POLL_SECONDS)
            continue

        if html is not None:
            m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
            title = m.group(1).strip() if m else ""
            challenge = looks_like_challenge(html)
            if not challenge or title != last_title:
                log(f"  attempt {attempt}: {len(html)} bytes, title={title!r}")
                last_title = title
            if not challenge:
                return tab, html

        if html is not None and time.time() >= next_click:
            await try_click_turnstile(tab)
            next_click = time.time() + TURNSTILE_CLICK_INTERVAL

        await tab.sleep(POLL_SECONDS)

    log(f"WARNING: page did not load or challenge did not clear for {url}")
    return tab, None


async def browser_fetch_once(tab, url: str) -> tuple[bytes | None, str]:
    """Fetch url with the page's own fetch(). Returns (bytes, "") or
    (None, what went wrong)."""
    js = (
        "(async () => {"
        "  try {"
        f"    const r = await fetch({json.dumps(url)}, {{credentials: 'include'}});"
        "    if (!r.ok) {"
        "      const m = r.headers.get('cf-mitigated');"
        "      return 'ERR:HTTP ' + r.status + (m ? ' (cf-mitigated: ' + m + ')' : '');"
        "    }"
        "    const b = new Uint8Array(await r.arrayBuffer());"
        "    let s = '';"
        "    for (let i = 0; i < b.length; i += 0x8000)"
        "      s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));"
        "    return 'OK:' + btoa(s);"
        "  } catch (e) {"
        "    return 'ERR:' + String(e);"
        "  }"
        "})()"
    )
    try:
        result = await with_timeout(
            tab.evaluate(js, await_promise=True, return_by_value=True),
            "tab.evaluate (fetch)",
            REQUEST_TIMEOUT,
        )
    except Exception as e:
        return None, f"evaluate raised {e}"
    if not isinstance(result, str) or not result.startswith("OK:"):
        if isinstance(result, str) and result.startswith("ERR:"):
            return None, result[4:]
        return None, f"unexpected result {type(result).__name__}"
    try:
        data = base64.b64decode(result[3:], validate=True)
    except (ValueError, binascii.Error) as e:
        return None, f"undecodable payload: {e}"
    if not data:
        return None, "no bytes returned"
    if is_challenge_payload(data):
        return None, "Cloudflare challenge instead of the file"
    return data, ""


def retryable_fetch_failure(detail: str) -> bool:
    status = re.match(r"HTTP (\d{3})", detail)
    return int(status.group(1)) in RETRYABLE_FETCH_STATUSES if status else True


async def browser_fetch(tab, url: str) -> bytes | None:
    """browser_fetch_once, repeated while the failure looks transient."""
    for attempt in range(1, BROWSER_FETCH_ATTEMPTS + 1):
        data, detail = await browser_fetch_once(tab, url)
        if data is not None:
            return data
        log(f"  browser fetch of {url} failed (attempt {attempt}/{BROWSER_FETCH_ATTEMPTS}): {detail}")
        if attempt == BROWSER_FETCH_ATTEMPTS or not retryable_fetch_failure(detail):
            return None
        await asyncio.sleep(BROWSER_FETCH_RETRY_SECONDS)
    return None


async def wait_for_judgment_link(tab, page_url: str, html: str) -> str | None:
    deadline = time.time() + JUDGMENT_LINK_WAIT_SECONDS
    while True:
        link = judgment_file_link(html, page_url)
        if link or time.time() >= deadline:
            return link
        await tab.sleep(1)
        try:
            html = await with_timeout(tab.get_content(), "tab.get_content")
        except Exception:
            pass


async def resolve_redirect(tab, link: str) -> str | None:
    """Where link redirects to, read from the browser's network events (a
    fetch() is refused by CORS once the redirect leaves the host, but only
    after Chrome has issued the redirected request)."""
    hops: dict[str, str] = {}

    def on_request(event):
        if event.redirect_response is not None:
            hops[event.redirect_response.url] = event.request.url

    tab.add_handler(uc.cdp.network.RequestWillBeSent, on_request)
    await with_timeout(tab.send(uc.cdp.network.enable()), "network.enable")
    await browser_fetch_once(tab, link)
    deadline = time.time() + JUDGMENT_REDIRECT_WAIT_SECONDS
    while link not in hops and time.time() < deadline:
        await tab.sleep(0.5)
    try:
        await with_timeout(tab.send(uc.cdp.network.disable()), "network.disable")
    except Exception:
        pass
    try:
        tab.remove_handler(uc.cdp.network.RequestWillBeSent, on_request)
    except Exception:
        pass
    final_url = follow_redirects(link, hops)
    return final_url if final_url != link else None


async def fetch_by_visiting(tab, url: str) -> bytes | None:
    """Load url in tab so Cloudflare can show its challenge, then fetch it.
    Once cleared the navigation becomes a download and the tab stays put, so
    keep retrying the fetch rather than waiting for the page to change."""
    log(f"  visiting {url} to clear its challenge")
    try:
        await with_timeout(tab.get(url), "tab.get")
    except Exception as e:
        log(f"  navigation to {url} ended: {e}")
    deadline = time.time() + MAX_WAIT_SECONDS
    next_click = time.time() + TURNSTILE_CLICK_INTERVAL
    detail = ""
    while time.time() < deadline:
        await tab.sleep(2)
        data, detail = await browser_fetch_once(tab, url)
        if data is not None:
            return data
        if time.time() >= next_click:
            await try_click_turnstile(tab)
            next_click = time.time() + TURNSTILE_CLICK_INTERVAL
    log(f"  WARNING: could not fetch {url}: {detail}")
    return None


class NotPublished(Exception):
    """The judgment page loaded but has no file link (not yet published, or
    suppressed)."""


async def download_judgment(browser, page_url: str) -> tuple[str, bytes, str] | None:
    """Fetch the file behind a Federal Court judgment page, in a tab of its
    own. Returns (filename, bytes, file_url), or None if it couldn't be
    fetched. Raises NotPublished if the page has no file link."""
    log(f"Fetching judgment {page_url}")
    side_tab = None
    try:
        side_tab = await with_timeout(browser.get("about:blank", new_tab=True), "browser.get")
        tab, html = await fetch_page(browser, page_url, tab=side_tab)
        if html is None:
            return None

        link = await wait_for_judgment_link(tab, page_url, html)
        if link is None:
            raise NotPublished(page_url)

        final_url = link
        data, _ = await browser_fetch_once(tab, link)
        if data is None or sniff_extension(data) is None:
            final_url = await resolve_redirect(tab, link)
            if final_url is None:
                log(f"  WARNING: {link} is neither a file nor a redirect to one")
                return None
            if not is_judgment_host_url(final_url):
                log(f"  WARNING: {link} redirects to {final_url}, not a Federal Court host")
                return None
            origin = f"https://{urlparse(final_url).hostname}/"
            tab, home_html = await fetch_page(browser, origin, tab=tab, same_host=True)
            if home_html is None:
                return None
            data = await browser_fetch(tab, final_url)
            if data is None:
                data = await fetch_by_visiting(tab, final_url)
            if data is None:
                return None
            if sniff_extension(data) is None:
                log(f"  WARNING: {final_url} is not a document")
                return None

        filename = judgment_filename(final_url, data)
        if filename is None:
            log(f"  WARNING: no safe filename for {final_url}")
            return None
        return filename, data, final_url
    except NotPublished:
        raise
    except Exception as e:
        log(f"  WARNING: fetching {page_url} failed: {e!r}")
        return None
    finally:
        if side_tab is not None:
            try:
                await with_timeout(side_tab.close(), "tab.close")
            except Exception:
                pass

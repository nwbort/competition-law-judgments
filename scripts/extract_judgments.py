#!/usr/bin/env python3
"""Turn the downloaded Word judgments into structured JSON and Markdown.

For every document under data/judgments/<year>/ this writes

    data/extracted/<year>/<name>.json   metadata + structured text
    data/extracted/<year>/<name>.md     the full text as Markdown

and an index of the metadata alone:

    data/extracted/index.json
    data/extracted/index.csv

What is extracted (see README for the JSON layout):

* The cover sheet ("Citation:", "File number:", "Judge:", "Catchwords:",
  "Legislation:", "Cases cited:", "Counsel for the Applicant:" ...), which
  the Federal Court's template has carried, in slightly different forms,
  since 2011. Labels are normalised to fixed keys; every row is also kept
  verbatim in ``cover_sheet``.
* The orders: one block per file number, with its parties, judge(s), date
  and the numbered orders.
* The reasons: the numbered paragraphs (the [n] used to cite a judgment),
  each with any quotes, lists and tables that belong to it, its footnotes,
  and the headings it falls under.
* Anything after the Associate's certificate (schedules of parties,
  annexures).
* The full text, in reading order, as Markdown.

Usage:
    python scripts/extract_judgments.py                 # every judgment
    python scripts/extract_judgments.py --only 2023FCA0665 2020FCAFC0145
"""
from __future__ import annotations

import argparse
import csv
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_reader import Para, Table, blocks_text, read_docx  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JUDGMENTS_DIR = os.path.join(ROOT, "data", "judgments")
JUDGMENTS_JSON = os.path.join(ROOT, "data", "judgments.json")
OUT_DIR = os.path.join(ROOT, "data", "extracted")

# Bump when the output format changes.
EXTRACTOR_VERSION = 2

FILE_NAME_RE = re.compile(r"^(\d{4})(FCAFC|FCA)(\d{4})$")
FILE_NUMBER_RE = re.compile(r"\b([A-Z]{3})\s*(\d{1,5})\s*(?:of|/)\s*(\d{4})\b")
CITATION_RE = re.compile(r"\[(\d{4})\]\s*(FCAFC|FCA)\s*(\d+)")
MONTHS = {m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july", "august",
     "september", "october", "november", "december"], 1)}
MONTH_RE = "|".join(MONTHS)

REASONS_RE = re.compile(
    r"^(REASONS\s+FOR\s+(JUDG(E)?MENT|DECISION|RULINGS?|ORDERS?)|REASONS|JUDGMENT)\b"
    r"(\s*\(.*\))?\s*$", re.I)
ORDERS_HEADING_RE = re.compile(r"^(ORDERS?|THE COURT (ORDERS|DECLARES|NOTES)( THAT)?:?)$", re.I)
CERTIFICATE_RE = re.compile(r"^(\d+\.?\s+)?I certify that the preceding", re.I)
# A number typed in front of the text rather than produced by Word's
# numbering: "12\t", "1.\t", "(a)\t", "(iv)\t", "[4.10]\t".
LITERAL_LABEL_RE = re.compile(
    r"^(\(?\d{1,4}[.)]?|\(?[A-Za-z]{1,2}[.)]|\(?[ivxlcdmIVXLCDM]{1,7}[.)]|\[\d+(\.\d+)*\]) *\t")
JUDGE_LINE_RE = re.compile(r"^(THE COURT|[A-Z][A-Z' ’.-]+ (C?J|JJ|P|ACJ|DCJ)):?$")
ROLE_RE = re.compile(
    r"^((First|Second|Third|Fourth|Fifth|Sixth|Seventh|Eighth|Ninth|Tenth|Eleventh|Twelfth|"
    r"[A-Z][a-z]+(-[A-Z]?[a-z]+)?)\s+)?"
    r"(Cross[- ])?(Applicant|Respondent|Appellant|Plaintiff|Defendant|Claimant|Intervener|"
    r"Intervenor|Interested Person|Prospective Applicant|Prospective Respondent|Cross-Claimant|"
    r"Cross-Respondent|Non-Party|Third Party)s?\b.*$", re.I)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


# --------------------------------------------------------------------------
# Small parsers


def clean(s: str) -> str:
    s = s.replace("\u00a0", " ").replace("\u2011", "-")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r" *\n *", "\n", s)
    return s.strip()


def lines(s: str) -> list[str]:
    return [l.strip() for l in clean(s).split("\n") if l.strip()]


def parse_dates(text: str) -> list[str]:
    """ISO dates in strings like "7 September 2011", "1-2 December 2025",
    "1, 2 & 3 July 2020", "28 February, 1 and 2 March 2016". A day range
    ("3-6 May") is expanded to every day in it."""
    if not text:
        return []
    t = text.replace("–", "-").replace("—", "-")
    # Group "<days> <Month> [Year]" and attach a missing year from the next
    # group that has one ("28 February, 1 March 2016").
    groups = []
    for m in re.finditer(
            rf"((?:\d{{1,2}}(?:st|nd|rd|th)?\s*(?:-|,|&|and|to)?\s*)+)\s*({MONTH_RE})\b,?\s*(\d{{4}})?",
            t, re.I):
        groups.append([m.group(1), MONTHS[m.group(2).lower()], int(m.group(3)) if m.group(3) else None])
    year = None
    for g in reversed(groups):
        if g[2] is None:
            g[2] = year
        year = g[2]
    out = []
    for days, month, yr in groups:
        if yr is None:
            continue
        nums = [(int(n), sep) for n, sep in re.findall(r"(\d{1,2})(?:st|nd|rd|th)?\s*(-|to)?", days)]
        i = 0
        while i < len(nums):
            n, sep = nums[i]
            if sep and i + 1 < len(nums):
                end = nums[i + 1][0]
                for d in range(n, max(n, end) + 1):
                    out.append((yr, month, d))
                i += 2
                continue
            out.append((yr, month, n))
            i += 1
    iso = []
    for y, mo, d in out:
        try:
            iso.append(date(y, mo, d).isoformat())
        except ValueError:
            pass
    return sorted(set(iso))


def file_numbers(text: str) -> list[str]:
    """"NSD 702 of 2006" -> "NSD702/2006" (the form data/judgments.json uses)."""
    seen = []
    for reg, n, y in FILE_NUMBER_RE.findall(text or ""):
        fn = f"{reg}{int(n)}/{y}"
        if fn not in seen:
            seen.append(fn)
    return seen


def split_judges(text: str) -> list[str]:
    """"ALLSOP CJ, BEACH AND COLVIN JJ" -> ["ALLSOP CJ", "BEACH J", "COLVIN J"]."""
    text = (text or "").upper().replace("’", "'")
    out = []
    for line in lines(text):
        line = re.sub(r"\s+", " ", line).strip().rstrip(":")
        m = re.match(r"^(.*?)\s+JJ\.?$", line, re.I)
        if m:
            names = re.split(r",\s*|\s+and\s+|\s*&\s*", m.group(1), flags=re.I)
            for nm in names:
                nm = nm.strip()
                if not nm:
                    continue
                # "ALLSOP CJ" inside a JJ list keeps its own title.
                if re.search(r"\b(CJ|ACJ|J|P|DCJ)$", nm):
                    out.append(nm)
                else:
                    out.append(f"{nm} J")
        else:
            for nm in re.split(r",\s*(?=[A-Z])|\s+and\s+", line, flags=re.I):
                if nm.strip():
                    out.append(nm.strip())
    return out


REGISTRIES = {"NSW": "New South Wales", "VIC": "Victoria", "QLD": "Queensland", "SA": "South Australia",
              "WA": "Western Australia", "TAS": "Tasmania", "ACT": "Australian Capital Territory",
              "NT": "Northern Territory"}


def normalise_registry(r):
    if not r:
        return None
    r = clean(r)
    r = REGISTRIES.get(r.upper().replace(".", ""), r)
    return r.title().replace(" Of ", " of ") if r.isupper() or r.islower() else r


def normalise_division(d):
    if not d:
        return None
    d = clean(d).title()
    return d if d.endswith("Division") else d + " Division"


def to_int(text: str):
    m = re.search(r"\d[\d,]*", text or "")
    return int(m.group(0).replace(",", "")) if m else None


def split_label(p: Para) -> tuple[str, str]:
    """(label, text) for a paragraph whose number was typed in."""
    if p.label:
        return p.label, p.text
    m = LITERAL_LABEL_RE.match(p.text)
    if m:
        return m.group(1), p.text[m.end():].strip()
    return "", p.text


def is_heading(p: Para) -> int | None:
    """Heading level (1-6) for paragraphs styled as headings."""
    s = p.style.lower()
    m = re.match(r"^heading (\d)", s)
    if m:
        return min(int(m.group(1)), 6)
    if s in ("normal headings", "style bold", "subs heading 1", "orders heading 1",
             "pleadings heading 1"):
        return 1
    m = re.match(r"^(subs|orders|pleadings) heading (\d)", s)
    if m:
        return min(int(m.group(2)), 6)
    if p.outline is not None and p.outline < 6 and not s.startswith("toc"):
        return p.outline + 1
    return None


def is_title_style(p: Para) -> bool:
    return p.style.lower() in ("mnc", "media neutral style", "courttitle")


def is_quote(p: Para) -> bool:
    return p.style.lower().startswith("quote") or p.style.lower().startswith("block")


# --------------------------------------------------------------------------
# Cover sheet


COVER_KEYS = [
    # (regex on the normalised label, key)
    (r"^citation$", "citation"),
    (r"^parties$", "parties"),
    (r"^file numbers?$|^file number\(s\)$", "file_numbers"),
    (r"^judges?$|^judge\(s\)$|^judgment of$|^before$", "judges"),
    (r"^date of judgment$|^date of (decision|orders?)$", "date_of_judgment"),
    (r"^date of publication of reasons$", "date_of_publication_of_reasons"),
    (r"^catchwords$", "catchwords"),
    (r"^legislation( cited)?$", "legislation"),
    (r"^cases cited$", "cases_cited"),
    (r"^texts cited$", "texts_cited"),
    (r"^dates? of hearings?$", "hearing"),
    (r"^dates? of (last|further) submissions?(/s)?$", "submissions"),
    (r"^place$", "place"),
    (r"^division$", "division"),
    (r"^registry$", "registry"),
    (r"^national practice area$", "national_practice_area"),
    (r"^sub-?area$", "sub_area"),
    (r"^category$", "category"),
    (r"^number of paragraphs$", "number_of_paragraphs"),
    (r"^corrigendum$|^date of corrigendum$", "corrigendum"),
    (r"^appeal from$", "appeal_from"),
    (r"^review of$", "review_of"),
]
REP_RE = re.compile(r"^(counsel|solicitors?)\s+for\s+(.+)$", re.I)
REP_GROUP_RE = re.compile(r"^for\s+(.+)$", re.I)


def norm_label(label: str) -> str:
    s = label.replace("–", "-").replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip().rstrip(":").strip().lower()
    return s


def cover_rows(tables: list[Table]) -> list[tuple[str, str]]:
    rows = []
    for t in tables:
        for row in t.rows:
            nested = [b for c in row for b in c if isinstance(b, Table)]
            if nested:  # a cover sheet laid out as a table inside a table
                rows.extend(cover_rows(nested))
                continue
            cells = [clean(blocks_text(c)) for c in row]
            cells = [c for c in cells if c]
            if not cells:
                continue
            if len(cells) == 1:
                # A label with no value ("For NSD 1986 of 2019:")
                if cells[0].endswith(":") or REP_GROUP_RE.match(cells[0]):
                    rows.append((cells[0], ""))
                continue
            rows.append((cells[0], "\n".join(cells[1:])))
    return rows


def parse_cover(rows) -> dict:
    meta = {}
    reps = []
    group = None
    for label, value in rows:
        nl = norm_label(label)
        m = REP_RE.match(nl)
        if m:
            role = "counsel" if m.group(1).startswith("counsel") else "solicitor"
            reps.append({
                "role": role,
                "for": re.sub(r"\s+", " ", label.strip().rstrip(":")).split(" for ", 1)[1].strip(),
                "name": clean(value),
                **({"proceeding": group} if group else {}),
            })
            continue
        g = REP_GROUP_RE.match(nl)
        if g and not value:
            fns = file_numbers(label)
            group = fns[0] if len(fns) == 1 else label.strip().rstrip(":")[4:].strip()
            continue
        for rx, key in COVER_KEYS:
            if re.match(rx, nl):
                if key in meta and value:
                    meta[key] += "\n" + value
                elif value:
                    meta[key] = value
                break
    meta["_representation"] = reps
    return meta


# --------------------------------------------------------------------------
# Court header tables (orders and older reasons blocks)


def table_kind(t: Table) -> str | None:
    texts = [" ".join(r).strip() for r in t.cell_texts()]
    first = next((x for x in texts if x), "")
    joined = "\n".join(texts)
    if CERTIFICATE_RE.match(first):
        return "certificate"
    if re.search(r"^(Catchwords|Citation|Date of judgment|Number of paragraphs)\b", joined, re.I | re.M):
        return None  # cover sheet
    filled = [r for r in t.cell_texts() if any(c.strip() for c in r)]
    if filled and sum(1 for r in filled if re.fullmatch(r"\[\d+\]", r[-1].strip())) >= max(2, len(filled) // 2):
        return "contents"
    if re.search(r"^\s*(BETWEEN|AND)\s*:", joined, re.M) and re.search(r"BETWEEN\s*:", joined):
        return "parties"
    if re.match(r"^\s*AND\s*:", first):
        return "more_parties"  # the AND: half, when it is a table of its own
    if re.match(r"^IN THE FEDERAL COURT", first, re.I):
        return "court"
    if re.search(r"^(JUDGES?|ORDER MADE BY|DATE OF ORDER|DATE|WHERE MADE|PLACE)\s*:", joined, re.M | re.I) \
            and len(texts) <= 5:
        return "order_details"
    if re.fullmatch(r"\s*[A-Z]{3}\s*\d+\s+of\s+\d{4}\s*", first or "") and len([x for x in texts if x]) <= 2:
        return "file_number"
    return None


def parse_contents(t: Table) -> list[dict]:
    """Table of contents rows: [{"heading": "2.1 Background", "paragraph": 13}]."""
    out = []
    for row in t.cell_texts():
        cells = [clean(c).replace("\n", " ") for c in row if clean(c)]
        if not cells:
            continue
        m = re.fullmatch(r"\[(\d+)\]", cells[-1])
        heading = re.sub(r"\s+", " ", " ".join(cells[:-1] if m else cells)).strip()
        if heading:
            out.append({"heading": heading, "paragraph": int(m.group(1)) if m else None})
    return out


def parse_parties(t: Table) -> list[dict]:
    """BETWEEN:/AND: table -> [{"name": ..., "role": ...}]."""
    parties = []
    for row in t.rows:
        cells = [blocks_text(c) for c in row]
        if not cells:
            continue
        head = clean(cells[0]).upper().rstrip(":")
        if head not in ("BETWEEN", "AND") and not any(head.startswith(x) for x in ("BETWEEN", "AND")):
            continue
        body = "\n".join(c for c in cells[1:] if c.strip())
        pending = []
        for line in lines(body):
            if ROLE_RE.match(line) and not line.isupper():
                for name in pending:
                    parties.append({"name": name, "role": line})
                pending = []
            else:
                pending.append(line)
        for name in pending:
            parties.append({"name": name, "role": None})
    return parties


def order_details(t: Table) -> dict:
    out = {}
    for row in t.cell_texts():
        cells = [clean(c) for c in row if clean(c)]
        if len(cells) < 2:
            continue
        k = norm_label(cells[0])
        v = cells[1]
        if k in ("judge", "judges", "order made by", "orders made by"):
            out["judges"] = split_judges(v)
        elif k in ("date of order", "date", "date of orders"):
            out["date_raw"] = v
            d = parse_dates(v)
            out["date"] = d[0] if d else None
        elif k in ("where made", "place"):
            out["place"] = v
    return out


# --------------------------------------------------------------------------
# Rendering blocks


class Footnotes:
    """Collects footnote references in reading order and numbers them the
    way Word displays them (1, 2, 3 ...)."""

    REF_RE = re.compile(r"\[\^(e?-?\d+)\]")

    def __init__(self, footnotes: dict, endnotes: dict):
        self.notes = {**footnotes, **{f"e{k}": v for k, v in endnotes.items()}}
        self.display = {}

    def number(self, text: str) -> tuple[str, list]:
        refs = []

        def sub(m):
            nid = m.group(1)
            if nid not in self.display:
                self.display[nid] = len(self.display) + 1
            n = self.display[nid]
            refs.append(n)
            return f"[^{n}]"

        return self.REF_RE.sub(sub, text), refs

    def text_of(self, n: int) -> str:
        for nid, d in self.display.items():
            if d == n:
                return clean(self.notes.get(nid, ""))
        return ""

    def all(self) -> list[dict]:
        return [{"n": d, "text": clean(self.notes.get(nid, ""))}
                for nid, d in sorted(self.display.items(), key=lambda x: x[1])]


def md_escape_cell(s: str) -> str:
    return clean(s).replace("|", "\\|").replace("\n", "<br>")


def table_markdown(t: Table, fn: Footnotes) -> str:
    rows = []
    for row in t.rows:
        cells = []
        for c in row:
            txt = blocks_text(c, "\n")
            txt, _ = fn.number(txt)
            cells.append(md_escape_cell(txt))
        rows.append(cells)
    rows = [r for r in rows if any(c for c in r)]
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    if width == 1:
        return "\n".join(r[0] for r in rows)
    out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(out)


def para_markdown(p: Para, fn: Footnotes, para_number: int | None = None) -> str:
    label, text = split_label(p)
    text, _ = fn.number(clean(text))
    if not text and not label:
        return ""
    h = is_heading(p)
    if para_number is not None:
        return f"[{para_number}] {text}"
    if (h and len(text) < 300 and not is_quote(p)) or JUDGE_LINE_RE.match(text):
        heading = f"{label} {text}" if label else text
        return "#" * min((h or 1) + 2, 6) + " " + heading.replace("\n", " ")
    indent = ""
    if p.level and p.level > 0:
        indent = "  " * p.level
    body = f"{label} {text}".strip() if label else text
    if is_quote(p):
        return "\n".join("> " + indent + l if l else ">" for l in body.split("\n"))
    return indent + body


def block_markdown(b, fn: Footnotes, para_number=None) -> str:
    if isinstance(b, Table):
        return table_markdown(b, fn)
    return para_markdown(b, fn, para_number)


def block_plain(b, fn: Footnotes) -> str:
    """Plain text of a block (footnote markers kept as [^n])."""
    if isinstance(b, Table):
        rows = []
        for row in b.rows:
            cells = [clean(blocks_text(c, " ")) for c in row]
            line = " | ".join(c for c in cells if c)
            if line:
                rows.append(line)
        text = "\n".join(rows)
    else:
        label, text = split_label(b)
        text = f"{label} {clean(text)}".strip() if label else clean(text)
    return fn.number(text)[0]


# --------------------------------------------------------------------------
# Segmentation


def segment(blocks):
    """Split the body into title, cover, orders part, reasons, after-matter."""
    i = 0
    n = len(blocks)
    title_paras = []
    # Title: paragraphs before the first table.
    while i < n and isinstance(blocks[i], Para):
        title_paras.append(blocks[i])
        i += 1
    cover = []
    while i < n:
        b = blocks[i]
        if isinstance(b, Table):
            k = table_kind(b)
            if k in ("parties", "court", "order_details", "file_number", "certificate"):
                break
            cover.append(b)
            i += 1
        elif not b.text.strip():
            i += 1
        else:
            break
    rest = blocks[i:]
    # Reasons heading
    r = None
    for j, b in enumerate(rest):
        if isinstance(b, Para) and REASONS_RE.match(b.text.strip()) and (is_heading(b) or b.bold):
            r = j
            break
    if r is None:
        # No heading: reasons start at the first ParaNumbering paragraph.
        for j, b in enumerate(rest):
            if isinstance(b, Para) and b.style == "ParaNumbering" and b.text:
                r = j
                break
    if r is None:
        r = len(rest)
    front = rest[:r]
    reasons_all = rest[r:]
    cert = None
    for j, b in enumerate(reasons_all):
        if isinstance(b, Table) and table_kind(b) == "certificate":
            cert = j
            break
        if isinstance(b, Para) and CERTIFICATE_RE.match(split_label(b)[1].strip()):
            cert = j
            break
    if cert is None:
        reasons, after = reasons_all, []
    else:
        reasons, after = reasons_all[:cert], reasons_all[cert:]
    return title_paras, cover, front, reasons, after


def parse_orders(front, fn: Footnotes):
    """The orders part: one block per header (BETWEEN: table). Also returns
    any table of contents and corrigenda found among them."""
    blocks_out = []
    contents = []
    corrigenda = []
    corr = None
    cur = None

    def new_block():
        return {"file_number": None, "parties": [], "judges": [], "date": None,
                "date_raw": None, "place": None, "orders": [], "_md": []}

    for b in front:
        if isinstance(b, Table) and table_kind(b) == "contents":
            contents.extend(parse_contents(b))
            continue
        if isinstance(b, Para) and re.match(r"^CORRIGEND(UM|A)$", clean(b.text), re.I):
            corr = []
            corrigenda.append(corr)
            continue
        if corr is not None:
            # A corrigendum runs until the judgment starts again.
            if (isinstance(b, Table) and table_kind(b) not in (None, "certificate")) or (
                    isinstance(b, Para) and re.match(r"^FEDERAL COURT OF AUSTRALIA$", clean(b.text), re.I)):
                corr = None
            else:
                t = block_plain(b, fn)
                if t:
                    corr.append(t)
                continue
        if isinstance(b, Para) and (CITATION_RE.search(b.text) or is_title_style(b) or re.match(
                r"^FEDERAL COURT OF AUSTRALIA$", clean(b.text), re.I)) and (cur is None or not cur["orders"]):
            continue  # the title, repeated after a corrigendum
        if isinstance(b, Table):
            k = table_kind(b)
            if k in ("court", "file_number") or (k == "parties" and (cur is None or cur["parties"] or cur["orders"])):
                if cur is None or cur["orders"] or cur["parties"] or (k == "court" and cur.get("file_number")):
                    cur = new_block()
                    blocks_out.append(cur)
            if cur is None:
                cur = new_block()
                blocks_out.append(cur)
            if k in ("court", "file_number", "parties"):
                fns = file_numbers(blocks_text(b.rows[0][-1] if b.rows and b.rows[0] else []) + "\n" +
                                   "\n".join(" ".join(r) for r in b.cell_texts() if not
                                             re.match(r"^\s*(BETWEEN|AND)", " ".join(r))))
                if fns and not cur["file_number"]:
                    cur["file_number"] = fns[0]
            if k == "court":
                for row in b.cell_texts():
                    t = clean(" ".join(row))
                    m = re.search(r"^(.*?)\s+(DISTRICT\s+)?REGISTRY", t, re.I)
                    if m:
                        cur["registry"] = m.group(1).title()
                    if re.search(r"DIVISION", t, re.I):
                        cur["division"] = re.sub(FILE_NUMBER_RE, "", t).strip().title()
            if k == "parties":
                cur["parties"] = parse_parties(b)
            elif k == "more_parties":
                cur["parties"] += parse_parties(b)
            elif k == "order_details":
                cur.update(order_details(b))
            elif k is None:
                cur["orders"].append({"label": "", "text": block_plain(b, fn), "level": 0, "table": True})
            continue
        if cur is None:
            cur = new_block()
            blocks_out.append(cur)
        text = clean(b.text)
        if not text:
            continue
        if ORDERS_HEADING_RE.match(text) or (is_heading(b) and not b.label):
            cur.setdefault("headings", []).append(text)
            continue
        label, body = split_label(b)
        body, _ = fn.number(clean(body))
        if re.match(r"^Note:", body):
            cur["note"] = body
            continue
        cur["orders"].append({"label": label, "text": body, "level": b.level or 0})
    out = []
    header_parties = []
    header = {}
    for blk in blocks_out:
        for k in ("registry", "division"):
            if blk.get(k) and k not in header:
                header[k] = blk[k]
        for p in blk["parties"]:
            if p not in header_parties:
                header_parties.append(p)
        if not blk["orders"]:
            continue  # the pre-2016 repeat of the header above the reasons
        blk.pop("_md", None)
        blk = {k: v for k, v in blk.items() if v not in (None, [], "")}
        out.append(blk)
    corrigenda = ["\n".join(c) for c in corrigenda if c]
    header["parties"] = header_parties
    return out, contents, corrigenda, header


def table_first_text(t: Table) -> str:
    for row in t.rows:
        for cell in row:
            for b in cell:
                if isinstance(b, Para) and (b.text or b.label):
                    return (b.label + "\t" + b.text) if b.label else b.text
            return ""
    return ""


def parse_reasons(reasons, fn: Footnotes, expected_total: int | None, contents: list):
    """Numbered paragraphs, each with the blocks that follow it. Material
    between a heading and the next numbered paragraph (a table, a quote) is
    kept in order as an entry with "n": null."""
    paras = []
    preamble = []
    headings = []  # stack of (level, text)
    expected = 1
    cur = None
    numbered_label_styles = {}
    # Which label source is the judgment's own paragraph numbering? Look at
    # styles: ParaNumbering (all templates since 2011), else the most common
    # style among integer-labelled paragraphs.
    for b in reasons:
        if isinstance(b, Para) and not is_quote(b):
            label, _ = split_label(b)
            if re.fullmatch(r"\d+\.?", label):
                numbered_label_styles[b.style] = numbered_label_styles.get(b.style, 0) + 1
    main_styles = {"ParaNumbering"} if "ParaNumbering" in numbered_label_styles else (
        {max(numbered_label_styles, key=numbered_label_styles.get)} if numbered_label_styles else set())
    # Word's own list for those paragraphs: a paragraph in the same list is a
    # numbered paragraph whatever its style (e.g. a last paragraph left in
    # "Body Text" but still numbered 156).
    main_lists = {b.num_id for b in reasons if isinstance(b, Para) and b.style in main_styles
                  and b.num_id is not None and b.label}

    judge_heading = None
    for b in reasons:
        if isinstance(b, Para):
            text = clean(b.text)
            if not text:
                continue
            label, body = split_label(b)
            n = None
            in_main = b.style in main_styles or (b.num_id is not None and b.num_id in main_lists)
            if in_main and re.fullmatch(r"\d+\.?", label or "") and (b.level in (None, 0)):
                cand = int(label.rstrip("."))
                if cand == expected or b.style == "ParaNumbering" or (expected < cand <= expected + 5):
                    n = cand
            if n is not None:
                expected = n + 1
                body_text, refs = fn.number(clean(body))
                cur = {
                    "n": n,
                    "headings": [h for _, h in headings],
                    "text": body_text,
                    "_footnotes": list(refs),
                }
                if judge_heading:
                    cur["judge"] = judge_heading
                paras.append(cur)
                continue
            if REASONS_RE.match(text):
                continue
            if JUDGE_LINE_RE.match(text) and not label:
                # "BEACH J:" / "THE COURT:" above the reasons of each judge.
                judge_heading = text.rstrip(":").strip()
                headings = []
                cur = None
                continue
            h = is_heading(b)
            if h and len(text) < 300 and not is_quote(b) and b.style not in main_styles:
                heading = f"{label} {clean(body)}" if label else text
                while headings and headings[-1][0] >= h:
                    headings.pop()
                headings.append((h, fn.number(heading)[0].replace("\t", " ")))
                cur = None
                continue
        if isinstance(b, Table) and table_kind(b) == "contents":
            contents.extend(parse_contents(b))
            continue
        piece = block_plain(b, fn)
        if not piece:
            continue
        refs = [int(x) for x in re.findall(r"\[\^(\d+)\]", piece)]
        # A paragraph that is a table, numbered in its first cell.
        m = re.match(r"^(\d+)\.? *\t", table_first_text(b)) if isinstance(b, Table) else None
        if m and int(m.group(1)) == expected and expected > 1:  # (not a table of contents)
            cur = {"n": expected, "headings": [h for _, h in headings],
                   "text": re.sub(r"^\d+\.? *\t?\s*", "", piece), "_footnotes": refs}
            if judge_heading:
                cur["judge"] = judge_heading
            paras.append(cur)
            expected += 1
        elif cur is not None:
            cur["text"] += "\n" + piece
            cur["_footnotes"].extend(refs)
        elif paras:
            cur = {"n": None, "headings": [h for _, h in headings], "text": piece, "_footnotes": refs}
            if judge_heading:
                cur["judge"] = judge_heading
            paras.append(cur)
        else:
            preamble.append(piece)
    for p in paras:
        ids = sorted(set(p.pop("_footnotes")))
        if ids:
            p["footnotes"] = [{"n": i, "text": fn.text_of(i)} for i in ids]
    warnings = []
    last = max((p["n"] for p in paras if p["n"] is not None), default=None)
    if expected_total is not None and last != expected_total:
        warnings.append(f"cover sheet says {expected_total} paragraphs, last numbered paragraph is {last}")
    nums = [p["n"] for p in paras if p["n"] is not None]
    gaps = [b for a, b in zip(nums, nums[1:]) if b != a + 1]
    if gaps:
        warnings.append(f"paragraph numbering is not continuous at {gaps[:10]}")
    return preamble, paras, warnings


def parse_after(after, fn: Footnotes) -> dict:
    out = {"certificate": None, "associate": None, "dated": None, "schedules": []}
    rest = []
    for b in after:
        t = block_plain(b, fn)
        if not t:
            continue
        if out["certificate"] is None and CERTIFICATE_RE.match(t):
            out["certificate"] = t
            continue
        m = re.match(r"^Associate:\s*(.*)$", t, re.S)
        if m and not out["associate"]:
            out["associate"] = clean(m.group(1)) or None
            continue
        m = re.match(r"^Dated:\s*(.*)$", t)
        if m and not out["dated"]:
            d = parse_dates(m.group(1))
            out["dated"] = d[0] if d else clean(m.group(1))
            continue
        rest.append(b)
    # Group whatever follows into schedules/annexures by heading.
    cur = None
    for b in rest:
        if isinstance(b, Para) and (is_heading(b) or b.bold) and len(b.text) < 200 and \
                re.match(r"^(SCHEDULE|ANNEXURE|ANNEX|APPENDIX|ATTACHMENT)", b.text.strip(), re.I):
            cur = {"title": clean(b.text), "text": ""}
            out["schedules"].append(cur)
            continue
        if cur is None:
            cur = {"title": None, "text": ""}
            out["schedules"].append(cur)
        piece = block_plain(b, fn)
        cur["text"] = (cur["text"] + "\n" + piece).strip()
    return out


# --------------------------------------------------------------------------
# Whole document


def load_index() -> dict:
    """judgments.json, keyed by the document's file stem."""
    try:
        with open(JUDGMENTS_JSON) as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    by_stem = {}
    for file_number, items in data.get("judgments", {}).items():
        for j in items:
            if not j.get("file"):
                continue
            stem = os.path.splitext(os.path.basename(j["file"]))[0]
            e = by_stem.setdefault(stem, {**j, "portal_file_numbers": []})
            e["portal_file_numbers"].append(file_number)
    return by_stem


def convert_doc(path: str, tmpdir: str) -> str | None:
    """Convert a legacy .doc to .docx with LibreOffice, if it is installed."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    subprocess.run([soffice, "--headless", "--convert-to", "docx", "--outdir", tmpdir, path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    out = os.path.join(tmpdir, os.path.splitext(os.path.basename(path))[0] + ".docx")
    return out if os.path.exists(out) else None


def extract(path: str, index_entry: dict | None) -> tuple[dict, str]:
    stem = os.path.splitext(os.path.basename(path))[0]
    with tempfile.TemporaryDirectory() as tmp:
        src = path
        if path.lower().endswith(".doc"):
            src = convert_doc(path, tmp)
            if not src:
                raise RuntimeError("legacy .doc and LibreOffice (soffice) is not installed")
        doc = read_docx(src)
    fn = Footnotes(doc.footnotes, doc.endnotes)
    title_paras, cover_tables, front, reasons, after = segment(doc.blocks)

    rows = cover_rows(cover_tables)
    cover = parse_cover(rows)
    reps = cover.pop("_representation")

    m = FILE_NAME_RE.match(stem)
    court = m.group(2) if m else None
    year = int(m.group(1)) if m else None
    number = int(m.group(3)) if m else None
    citation = f"[{year}] {court} {number}" if m else None

    # Case name: the "MNC" (media neutral citation) line under the court name.
    title_line = " ".join(clean(p.text) for p in title_paras
                          if p.text.strip() and not re.match(r"^FEDERAL COURT OF AUSTRALIA$", clean(p.text), re.I))
    if not CITATION_RE.search(title_line):
        title_line = ""
    if not title_line and cover.get("citation"):
        title_line = clean(cover["citation"])
    name = clean(CITATION_RE.sub("", title_line)).rstrip(" ,;") if title_line else None

    orders, contents, corrigenda, header = parse_orders(front, fn)
    expected_total = to_int(cover.get("number_of_paragraphs"))
    preamble, paragraphs, warnings = parse_reasons(reasons, fn, expected_total, contents)
    after_info = parse_after(after, fn)

    # Registry/division: cover sheet (2016+) or the court header (earlier).
    registry = cover.get("registry")
    division = cover.get("division")
    registry = registry or header.get("registry")
    division = division or header.get("division")
    hearing_raw = cover.get("hearing")
    judges = split_judges(cover.get("judges", "")) or next(
        (o["judges"] for o in orders if o.get("judges")), [])
    fns = file_numbers(cover.get("file_numbers", "")) or [o["file_number"] for o in orders if o.get("file_number")]
    djudg = parse_dates(cover.get("date_of_judgment", ""))

    parties = list(header["parties"])
    if not parties:
        # No orders page (e.g. reasons delivered first): use the BETWEEN:
        # table at the head of the reasons.
        for b in reasons[:15]:
            if isinstance(b, Table) and table_kind(b) == "parties":
                parties = parse_parties(b)
                break

    meta = {
        "citation": citation,
        "name": name,
        "court": court,
        "year": year,
        "number": number,
        "date_of_judgment": djudg[0] if djudg else None,
        "file_numbers": fns,
        "judges": judges,
        "registry": normalise_registry(registry),
        "division": normalise_division(division),
        "national_practice_area": cover.get("national_practice_area"),
        "sub_area": cover.get("sub_area"),
        "category": cover.get("category"),
        "place": cover.get("place") or next((o.get("place") for o in orders if o.get("place")), None),
        "hearing_dates": parse_dates(hearing_raw),
        "hearing_raw": hearing_raw,
        "last_submissions": (parse_dates(cover.get("submissions")) or [None])[-1],
        "number_of_paragraphs": expected_total,
        "catchwords": lines(cover.get("catchwords", "")),
        "legislation": lines(cover.get("legislation", "")),
        "cases_cited": lines(cover.get("cases_cited", "")),
        "texts_cited": lines(cover.get("texts_cited", "")),
        "appeal_from": cover.get("appeal_from"),
        "review_of": cover.get("review_of"),
        "corrigendum": cover.get("corrigendum"),
        "parties": parties,
        "parties_summary": cover.get("parties"),
        "representation": reps,
    }
    if index_entry:
        extra = sorted(set(index_entry.get("portal_file_numbers") or []) - set(fns))
        if fns and extra:
            warnings.append(f"listed on the portal under {', '.join(extra)}, which the judgment itself "
                            f"does not name (it names {', '.join(fns)})")
        meta["url"] = index_entry.get("url")
        meta["portal_file_numbers"] = index_entry.get("portal_file_numbers")
        meta["portal_event"] = index_entry.get("event")
        meta["portal_event_date"] = index_entry.get("event_date")
    meta = {k: v for k, v in meta.items() if v not in (None, "", [])}

    # Full text as Markdown, in reading order.
    md = [f"# {title_line or stem}", ""]
    if rows:
        md.append("## Cover sheet\n")
        for label, value in rows:
            v = clean(fn.number(value)[0])
            md.append(f"**{clean(label).rstrip(':')}:** " + (v.replace("\n", "  \n") if v else ""))
            md.append("")
    body_md = []
    # Re-render the front matter and reasons block by block.
    for b in front:
        piece = block_markdown(b, fn)
        if piece:
            body_md.append(piece)
    expected_numbers = {p["n"] for p in paragraphs}
    seen = set()
    for b in reasons:
        num = None
        if isinstance(b, Para):
            label, _ = split_label(b)
            if re.fullmatch(r"\d+\.?", label or "") and b.level in (None, 0) and not is_quote(b):
                cand = int(label.rstrip("."))
                if cand in expected_numbers and cand not in seen:
                    num = cand
                    seen.add(cand)
        piece = block_markdown(b, fn, num)
        if piece:
            body_md.append(piece)
    for b in after:
        piece = block_markdown(b, fn)
        if piece:
            body_md.append(piece)
    md.append("## Judgment\n")
    md.append("\n\n".join(body_md))
    notes = fn.all()
    if notes:
        md.append("\n## Footnotes\n")
        for note in notes:
            md.append(f"[^{note['n']}]: " + note["text"].replace("\n", "\n    "))
    markdown = "\n".join(md).rstrip() + "\n"

    record = {
        "extractor_version": EXTRACTOR_VERSION,
        "source_file": os.path.relpath(path, ROOT),
        "metadata": meta,
        "cover_sheet": [{"label": clean(l).rstrip(":"), "value": clean(fn.number(v)[0])} for l, v in rows],
        "contents": contents,
        "corrigenda": corrigenda,
        "orders": orders,
        "reasons": {
            "preamble": preamble,
            "paragraphs": paragraphs,
        },
        "certificate": after_info["certificate"],
        "associate": after_info["associate"],
        "dated": after_info["dated"],
        "schedules": after_info["schedules"],
        "footnotes": notes,
        "warnings": warnings,
    }
    return record, markdown


INDEX_FIELDS = ["citation", "name", "court", "year", "number", "date_of_judgment", "file_numbers",
                "judges", "registry", "division", "national_practice_area", "sub_area", "category",
                "hearing_dates", "number_of_paragraphs", "catchwords", "legislation", "url"]


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def summarise(record: dict, json_path: str, md_path: str) -> dict:
    meta = record["metadata"]
    summary = {k: meta.get(k) for k in INDEX_FIELDS if meta.get(k) is not None}
    summary["paragraphs_found"] = sum(1 for p in record["reasons"]["paragraphs"] if p["n"] is not None)
    summary["source_file"] = record["source_file"]
    summary["json"] = os.path.relpath(json_path, ROOT)
    summary["markdown"] = os.path.relpath(md_path, ROOT)
    if record["warnings"]:
        summary["warnings"] = record["warnings"]
    return summary


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--only", nargs="+", help="file stems to extract, e.g. 2023FCA0665")
    ap.add_argument("--force", action="store_true",
                    help="re-extract even when the output is up to date with the document")
    ap.add_argument("--out", default=OUT_DIR)
    args = ap.parse_args(argv)

    paths = sorted(glob.glob(os.path.join(JUDGMENTS_DIR, "*", "*.doc*")))
    paths = [p for p in paths if not os.path.basename(p).startswith("~$")]
    if args.only:
        want = {s.upper() for s in args.only}
        paths = [p for p in paths if os.path.splitext(os.path.basename(p))[0].upper() in want]
    index = load_index()
    summaries = []
    failures = done = skipped = 0
    for path in paths:
        stem = os.path.splitext(os.path.basename(path))[0]
        out_dir = os.path.join(args.out, os.path.basename(os.path.dirname(path)))
        json_path = os.path.join(out_dir, stem + ".json")
        md_path = os.path.join(out_dir, stem + ".md")
        digest = sha256(path)
        record = None
        if not args.force and os.path.exists(json_path) and os.path.exists(md_path):
            try:
                with open(json_path) as f:
                    old = json.load(f)
                if old.get("extractor_version") == EXTRACTOR_VERSION and old.get("source_sha256") == digest:
                    record = old
            except (OSError, ValueError):
                pass
        if record is not None:
            skipped += 1
            summaries.append(summarise(record, json_path, md_path))
            continue
        try:
            record, markdown = extract(path, index.get(stem))
        except Exception as e:  # keep going; report at the end
            failures += 1
            log(f"{stem}: FAILED: {e}")
            continue
        record["source_sha256"] = digest
        record["markdown_file"] = os.path.relpath(md_path, ROOT)
        os.makedirs(out_dir, exist_ok=True)
        with open(json_path, "w") as f:
            json.dump(record, f, ensure_ascii=False, indent=1)
            f.write("\n")
        with open(md_path, "w") as f:
            f.write(markdown)
        done += 1
        summaries.append(summarise(record, json_path, md_path))
        w = f" ({'; '.join(record['warnings'])})" if record["warnings"] else ""
        numbered = sum(1 for p in record["reasons"]["paragraphs"] if p["n"] is not None)
        log(f"{stem}: {numbered} paragraphs, {len(record['orders'])} order block(s), "
            f"{len(record['footnotes'])} footnotes{w}")

    if not args.only:
        summaries.sort(key=lambda s: (s.get("year") or 0, s.get("court") or "", s.get("number") or 0))
        os.makedirs(args.out, exist_ok=True)
        with open(os.path.join(args.out, "index.json"), "w") as f:
            json.dump({"extractor_version": EXTRACTOR_VERSION, "count": len(summaries),
                       "judgments": summaries}, f, ensure_ascii=False, indent=1)
            f.write("\n")
        with open(os.path.join(args.out, "index.csv"), "w", newline="") as f:
            cols = INDEX_FIELDS + ["paragraphs_found", "json", "markdown"]
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            for s in summaries:
                row = dict(s)
                for k, v in row.items():
                    if isinstance(v, list):
                        row[k] = "; ".join(str(x) for x in v)
                w.writerow(row)
    log(f"Extracted {done}, up to date {skipped}, failed {failures} (of {len(paths)} documents)")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())

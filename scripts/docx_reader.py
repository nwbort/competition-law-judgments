"""Read a Word (.docx) judgment into an ordered list of blocks.

python-docx's ``Paragraph.text`` only sees runs that are direct children of a
paragraph, so it silently drops text inside fields (``w:fldSimple``), smart
tags (``w:smartTag``), content controls (``w:sdt``) and the like. Federal
Court judgments use all of these: the cover sheet's file number and judge are
``REF`` fields, and older judgments wrap place names in smart tags ("NEW SOUTH
WALES DISTRICT REGISTRY" comes out as "NEW SOUTH  DISTRICT REGISTRY"). So this
module walks the XML itself.

It also works out Word's automatic numbering (numbering.xml), because the
paragraph numbers of a judgment -- the [12] that everyone cites -- and the
numbers of its orders are not in the text at all.

The result of :func:`read_docx` is a :class:`Document`:

* ``blocks``: the body in order, each either a :class:`Para` or a
  :class:`Table` (rows of cells, each cell a list of blocks).
* ``footnotes`` / ``endnotes``: ``{id: text}``. References in the body are
  written inline as ``[^id]``.
* ``core``: docProps/core.xml (title, creator, created, modified ...).
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def wval(el, attr: str = "val"):
    return None if el is None else el.get(q(attr))


@dataclass
class Para:
    text: str
    style: str = ""  # style name (not id), e.g. "ParaNumbering"
    label: str = ""  # automatic number, e.g. "12", "(a)", "1."
    level: int | None = None  # numbering level (ilvl), when numbered
    num_id: str | None = None
    outline: int | None = None  # outline level (0 = Heading 1), when known
    bold: bool = False  # every visible run is bold
    page_break_before: bool = False

    kind = "p"


@dataclass
class Table:
    rows: list  # list[list[list[Para | Table]]]

    kind = "table"

    def cell_texts(self):
        """Rows as lists of cell strings, with horizontally merged cells
        (repeated by gridSpan) collapsed."""
        out = []
        for row in self.rows:
            out.append([blocks_text(cell) for cell in row])
        return out


@dataclass
class Document:
    blocks: list
    footnotes: dict = field(default_factory=dict)
    endnotes: dict = field(default_factory=dict)
    core: dict = field(default_factory=dict)


def blocks_text(blocks, sep: str = "\n") -> str:
    parts = []
    for b in blocks:
        if isinstance(b, Para):
            t = (b.label + " " + b.text).strip() if b.label else b.text
            if t.strip():
                parts.append(t)
        else:
            for row in b.rows:
                cells = [blocks_text(c, " ") for c in row]
                line = " | ".join(c for c in cells if c)
                if line:
                    parts.append(line)
    return sep.join(parts)


# --------------------------------------------------------------------------
# Styles


class Styles:
    def __init__(self, root):
        self.by_id = {}
        self.default_para = None
        if root is None:
            return
        for s in root.iter(q("style")):
            sid = s.get(q("styleId"))
            name_el = s.find(q("name"))
            based = s.find(q("basedOn"))
            ppr = s.find(q("pPr"))
            rpr = s.find(q("rPr"))
            num = None
            outline = None
            if ppr is not None:
                npr = ppr.find(q("numPr"))
                if npr is not None:
                    num = (wval(npr.find(q("numId"))), wval(npr.find(q("ilvl"))))
                ol = ppr.find(q("outlineLvl"))
                if ol is not None:
                    outline = int(wval(ol))
            self.by_id[sid] = {
                "name": wval(name_el) or sid,
                "type": s.get(q("type")),
                "based_on": wval(based),
                "num": num,
                "outline": outline,
                "bold": _on(rpr.find(q("b"))) if rpr is not None and rpr.find(q("b")) is not None else None,
                "vanish": _on(rpr.find(q("vanish"))) if rpr is not None and rpr.find(q("vanish")) is not None else None,
            }
            if s.get(q("type")) == "paragraph" and s.get(q("default")) in ("1", "true"):
                self.default_para = sid

    def name(self, sid):
        st = self.by_id.get(sid)
        return st["name"] if st else (sid or "")

    def _chain(self, sid):
        seen = set()
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            yield self.by_id[sid]
            sid = self.by_id[sid]["based_on"]

    def inherited(self, sid, key):
        for st in self._chain(sid):
            if st[key] is not None:
                return st[key]
        return None


def _on(el) -> bool:
    """A toggle property (w:b, w:vanish...) is on unless val says otherwise."""
    if el is None:
        return False
    v = el.get(q("val"))
    return v is None or v.lower() not in ("0", "false", "off")


# --------------------------------------------------------------------------
# Numbering


ROMAN = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
         (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]


def to_roman(n: int) -> str:
    out = ""
    for v, s in ROMAN:
        while n >= v:
            out += s
            n -= v
    return out


def to_letter(n: int) -> str:
    # Word: a..z, aa..zz, aaa..
    if n <= 0:
        return ""
    return chr(ord("a") + (n - 1) % 26) * ((n - 1) // 26 + 1)


def format_number(n: int, fmt: str) -> str:
    if fmt in ("decimal", None, "decimalEnclosedParen", "decimalEnclosedFullstop"):
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "lowerLetter":
        return to_letter(n)
    if fmt == "upperLetter":
        return to_letter(n).upper()
    if fmt == "lowerRoman":
        return to_roman(n)
    if fmt == "upperRoman":
        return to_roman(n).upper()
    if fmt == "ordinal":
        suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
        return f"{n}{suffix}"
    if fmt in ("none", "bullet"):
        return ""
    return str(n)


class Numbering:
    """Tracks list counters the way Word does: counters belong to the
    abstract numbering definition, so separate w:num instances that point at
    the same abstractNum keep counting unless a lvlOverride/startOverride
    restarts them."""

    def __init__(self, root):
        self.abstract = {}  # abstractNumId -> {ilvl: lvl dict}
        self.nums = {}  # numId -> (abstractNumId, {ilvl: override dict})
        self.counters = {}  # abstractNumId -> {ilvl: count}
        self.started = set()  # numIds whose startOverrides have been applied
        if root is None:
            return
        for an in root.findall(q("abstractNum")):
            self.abstract[an.get(q("abstractNumId"))] = {
                int(l.get(q("ilvl"))): self._lvl(l) for l in an.findall(q("lvl"))
            }
            link = an.find(q("numStyleLink"))
            if link is not None:
                self.abstract[an.get(q("abstractNumId"))]["_style_link"] = wval(link)
        for n in root.findall(q("num")):
            nid = n.get(q("numId"))
            aid = wval(n.find(q("abstractNumId")))
            overrides = {}
            for o in n.findall(q("lvlOverride")):
                ilvl = int(o.get(q("ilvl")))
                d = {}
                so = o.find(q("startOverride"))
                if so is not None:
                    d["start"] = int(wval(so))
                lvl = o.find(q("lvl"))
                if lvl is not None:
                    d["lvl"] = self._lvl(lvl)
                overrides[ilvl] = d
            self.nums[nid] = (aid, overrides)

    @staticmethod
    def _lvl(l):
        start = l.find(q("start"))
        restart = l.find(q("lvlRestart"))
        return {
            "start": int(wval(start)) if start is not None else 0,
            "fmt": wval(l.find(q("numFmt"))) or "decimal",
            "text": wval(l.find(q("lvlText"))),
            "restart": int(wval(restart)) if restart is not None else None,
            "legal": l.find(q("isLgl")) is not None,
        }

    def _levels(self, num_id):
        aid, overrides = self.nums.get(num_id, (None, {}))
        levels = dict(self.abstract.get(aid, {}))
        levels.pop("_style_link", None)
        for ilvl, o in overrides.items():
            if "lvl" in o:
                levels[ilvl] = o["lvl"]
        return aid, overrides, levels

    def next_label(self, num_id, ilvl) -> str | None:
        if num_id in (None, "0") or num_id not in self.nums:
            return None
        aid, overrides, levels = self._levels(num_id)
        lvl = levels.get(ilvl)
        if lvl is None:
            return None
        counters = self.counters.setdefault(aid, {})
        if num_id not in self.started:
            self.started.add(num_id)
            for ol, o in overrides.items():
                if "start" in o:
                    # The override's first use starts the level at `start`.
                    counters[ol] = o["start"] - 1
                    for deeper in [k for k in counters if k > ol]:
                        del counters[deeper]
        start = lvl["start"]
        counters[ilvl] = counters.get(ilvl, start - 1) + 1
        # Restart deeper levels (lvlRestart is rarely set to anything but the
        # default "restart after any higher level").
        for deeper in [k for k in counters if k > ilvl]:
            del counters[deeper]
        text = lvl["text"]
        if text is None:
            return ""

        def sub(m):
            k = int(m.group(1)) - 1
            fmt = levels.get(k, {}).get("fmt", "decimal")
            if lvl["legal"]:
                fmt = "decimal"
            n = counters.get(k, levels.get(k, {}).get("start", 1))
            return format_number(n, fmt)

        if lvl["fmt"] == "bullet":
            return "•"
        return re.sub(r"%(\d)", sub, text).strip()


# --------------------------------------------------------------------------
# Text


SKIP = {q("del"), q("moveFrom"), q("instrText"), q("delText"), q("delInstrText"),
        q("rPr"), q("pPr"), q("sectPr"), q("tblPr"), q("trPr"), q("tcPr"),
        q("proofErr"), q("bookmarkStart"), q("bookmarkEnd"), q("commentRangeStart"),
        q("commentRangeEnd"), q("commentReference"), q("drawing"), q("pict"),
        q("object"), q("lastRenderedPageBreak"), q("fldChar"), q("sdtPr"), q("sdtEndPr")}


class TextCollector:
    def __init__(self, styles: Styles, para_style: str | None):
        self.styles = styles
        self.para_style = para_style
        self.parts = []
        self.bold_chars = 0
        self.plain_chars = 0

    def run_props(self, r):
        rpr = r.find(q("rPr"))
        rstyle = wval(rpr.find(q("rStyle"))) if rpr is not None else None
        b = rpr.find(q("b")) if rpr is not None else None
        if b is not None:
            bold = _on(b)
        else:
            bold = bool(self.styles.inherited(rstyle, "bold") if rstyle else None) or bool(
                self.styles.inherited(self.para_style, "bold"))
        v = rpr.find(q("vanish")) if rpr is not None else None
        if v is not None:
            hidden = _on(v)
        else:
            hidden = bool(self.styles.inherited(rstyle, "vanish")) if rstyle else False
        return bold, hidden

    def add(self, s: str, bold: bool):
        self.parts.append(s)
        n = len(s.strip())
        if bold:
            self.bold_chars += n
        else:
            self.plain_chars += n

    def walk(self, el):
        for child in el:
            tag = child.tag
            if not isinstance(tag, str) or tag in SKIP:
                continue
            if tag == q("r"):
                self.run(child)
            elif tag == f"{{{MC}}}AlternateContent":
                # Prefer the Fallback rendering (usually VML) over Choice.
                fb = child.find(f"{{{MC}}}Fallback")
                if fb is not None:
                    self.walk(fb)
            else:
                # hyperlink, smartTag, fldSimple, ins, sdt/sdtContent,
                # customXml, moveTo, bdo, dir, oMath...
                self.walk(child)

    def run(self, r):
        bold, hidden = self.run_props(r)
        if hidden:
            return
        for c in r:
            tag = c.tag
            if tag == q("t"):
                self.add(c.text or "", bold)
            elif tag == q("tab") or tag == q("ptab"):
                self.parts.append("\t")
            elif tag in (q("br"), q("cr")):
                if c.get(q("type")) != "page":
                    self.parts.append("\n")
            elif tag == q("noBreakHyphen"):
                self.add("-", bold)
            elif tag == q("sym"):
                ch = c.get(q("char")) or ""
                try:
                    code = int(ch, 16)
                    # Symbol-font private-use codes (F0xx) map to Latin-1.
                    if 0xF000 <= code <= 0xF0FF:
                        code -= 0xF000
                    self.add(chr(code), bold)
                except ValueError:
                    pass
            elif tag == q("footnoteReference"):
                self.parts.append(f"[^{c.get(q('id'))}]")
            elif tag == q("endnoteReference"):
                self.parts.append(f"[^e{c.get(q('id'))}]")
            elif tag in (q("drawing"), q("pict"), q("object")):
                continue
            elif tag == f"{{{MC}}}AlternateContent":
                continue

    def text(self) -> str:
        s = "".join(self.parts).replace(" ", " ")
        s = re.sub(r"[ \t]*\n[ \t]*", "\n", s)
        return s.strip()


# --------------------------------------------------------------------------
# Body


class Reader:
    def __init__(self, styles: Styles, numbering: Numbering):
        self.styles = styles
        self.numbering = numbering

    def para(self, p) -> Para:
        ppr = p.find(q("pPr"))
        sid = None
        num_id = ilvl = None
        outline = None
        pbb = False
        if ppr is not None:
            sid = wval(ppr.find(q("pStyle")))
            npr = ppr.find(q("numPr"))
            if npr is not None:
                num_id = wval(npr.find(q("numId")))
                il = wval(npr.find(q("ilvl")))
                ilvl = int(il) if il is not None else None
            ol = ppr.find(q("outlineLvl"))
            if ol is not None:
                outline = int(wval(ol))
            pbb = _on(ppr.find(q("pageBreakBefore"))) if ppr.find(q("pageBreakBefore")) is not None else False
        sid = sid or self.styles.default_para
        if num_id is None or ilvl is None:
            snum = self.styles.inherited(sid, "num")
            if snum:
                if num_id is None:
                    num_id = snum[0]
                if ilvl is None and snum[1] is not None:
                    ilvl = int(snum[1])
        if num_id is not None and ilvl is None:
            ilvl = 0
        if outline is None:
            outline = self.styles.inherited(sid, "outline")
        tc = TextCollector(self.styles, sid)
        tc.walk(p)
        text = tc.text()
        label = None
        # Word counts every paragraph that carries numbering, empty or not,
        # so the counters advance here even when the text is blank.
        if num_id not in (None, "0"):
            label = self.numbering.next_label(num_id, ilvl)
        return Para(
            text=text,
            style=self.styles.name(sid),
            label=label or "",
            level=ilvl if label is not None else None,
            num_id=num_id if label is not None else None,
            outline=outline,
            bold=tc.bold_chars > 0 and tc.plain_chars == 0,
            page_break_before=pbb,
        )

    def table(self, tbl) -> Table:
        rows = []
        for tr in tbl.findall(q("tr")):
            cells = []
            for tc in tr.findall(q("tc")):
                tcpr = tc.find(q("tcPr"))
                vmerge = tcpr.find(q("vMerge")) if tcpr is not None else None
                if vmerge is not None and wval(vmerge) in (None, "continue"):
                    cells.append([])  # continuation of a vertically merged cell
                    continue
                cells.append(self.blocks(tc))
            # rows can also be wrapped in sdt
            rows.append(cells)
        return Table(rows=rows)

    def blocks(self, container) -> list:
        out = []
        for child in container:
            tag = child.tag
            if tag == q("p"):
                out.append(self.para(child))
            elif tag == q("tbl"):
                out.append(self.table(child))
            elif tag in (q("sdt"), q("customXml"), q("ins")):
                content = child.find(q("sdtContent")) if tag == q("sdt") else child
                if content is not None:
                    out.extend(self.blocks(content))
            elif tag == q("sdtContent"):
                out.extend(self.blocks(child))
        return out


def _xml(z: zipfile.ZipFile, name: str):
    try:
        data = z.read(name)
    except KeyError:
        return None
    return etree.fromstring(data, parser=etree.XMLParser(huge_tree=True, recover=True))


def _notes(root, reader: Reader, tag: str) -> dict:
    out = {}
    if root is None:
        return out
    for n in root.findall(q(tag)):
        if n.get(q("type")) in ("separator", "continuationSeparator", "continuationNotice"):
            continue
        nid = n.get(q("id"))
        paras = [b for b in reader.blocks(n)]
        text = blocks_text(paras, "\n").strip()
        out[nid] = text
    return out


def _core(z) -> dict:
    root = _xml(z, "docProps/core.xml")
    out = {}
    if root is None:
        return out
    for el in root:
        if isinstance(el.tag, str) and el.text and el.text.strip():
            out[etree.QName(el).localname] = el.text.strip()
    return out


def read_docx(path: str) -> Document:
    with zipfile.ZipFile(path) as z:
        styles = Styles(_xml(z, "word/styles.xml"))
        numbering = Numbering(_xml(z, "word/numbering.xml"))
        reader = Reader(styles, numbering)
        doc = _xml(z, "word/document.xml")
        body = doc.find(q("body"))
        blocks = reader.blocks(body)
        # Notes use their own numbering instances rarely; a fresh reader keeps
        # the body's counters untouched.
        note_reader = Reader(styles, Numbering(_xml(z, "word/numbering.xml")))
        footnotes = _notes(_xml(z, "word/footnotes.xml"), note_reader, "footnote")
        endnotes = _notes(_xml(z, "word/endnotes.xml"), note_reader, "endnote")
        core = _core(z)
    return Document(blocks=blocks, footnotes=footnotes, endnotes=endnotes, core=core)

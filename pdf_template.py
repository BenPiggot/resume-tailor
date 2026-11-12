"""
Use an existing resume PDF as a layout template.

The PDF is split into text "slots" (a bullet, a paragraph, a skills line...). Claude rewrites the
text of some slots; this module removes ONLY the changed text from the original PDF (lines, shapes,
colors, images and all unchanged text stay exactly as they were) and writes the new text into the
same slots with the same fonts, sizes, colors and line positions.

Slots can mix two styles, e.g. black bullet text with colored "(React, TypeScript)" tech tags.
The secondary style is written as ⟦...⟧ in the text Claude sees and returns.
"""
from __future__ import annotations

import json
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None

ALT_OPEN, ALT_CLOSE = "⟦", "⟧"
BULLETS = set("•●▪◦‣∙·■□➢►▶✓✔-–—*○◆◇")
LABEL_END = re.compile(r"[:|]\s*$")
CONTACT = re.compile(r"@|https?://|www\.|linkedin\.com|github\.com|\(\d{3}\)\s*\d{3}|\b\d{3}[-.\s]\d{3}[-.\s]\d{4}\b", re.I)
DATE = re.compile(r"\b(19|20)\d{2}\b|\bPresent\b|\bCurrent\b", re.I)


def require():
    if fitz is None:
        sys.exit("Matching your PDF's layout needs PyMuPDF:  pip install pymupdf")


def strip_marks(t: str) -> str:
    return t.replace(ALT_OPEN, "").replace(ALT_CLOSE, "")


# ---------------------------------------------------------------- data model


@dataclass(frozen=True)
class Style:
    font: str
    size: float
    color: tuple
    flags: int

    @property
    def bold(self):
        return bool(self.flags & 16) or bool(re.search(r"bold|black|heavy|semibold", self.font, re.I))

    @property
    def italic(self):
        return bool(self.flags & 2) or bool(re.search(r"italic|oblique", self.font, re.I))

    @property
    def serif(self):
        return bool(self.flags & 4) and not re.search(r"sans", self.font, re.I)

    @property
    def mono(self):
        return bool(self.flags & 8) or bool(re.search(r"mono|courier|consol", self.font, re.I))


@dataclass
class Line:
    baseline: float
    x_start: float      # where the editable text starts on this line
    x1: float
    redact: list        # rects covering the editable text
    runs: list          # [(text, Style)]

    @property
    def text(self):
        return "".join(t for t, _ in self.runs)


@dataclass
class Unit:
    id: str
    page: int
    lines: list = field(default_factory=list)
    prefix: str = ""        # bullet glyph or "Label:" (kept as is)
    bullet: bool = False
    right: float = 0        # right edge available for text
    align: str = "left"
    protected: bool = False
    main: Style | None = None
    alt: Style | None = None
    text: str = ""          # with ⟦⟧ marking the alt style
    plain: str = ""

    @property
    def x0(self):
        return min(l.x_start for l in self.lines)

    @property
    def style(self):
        return self.main

    def finalize(self):
        counts = {}
        for ln in self.lines:
            for t, st in ln.runs:
                n = len(t.strip())
                counts[st] = counts.get(st, 0) + n
        ranked = sorted(counts, key=counts.get, reverse=True)
        self.main = ranked[0]
        others = [s for s in ranked[1:] if counts[s] > 0]
        self.alt = others[0] if others else None
        parts = []
        for i, ln in enumerate(self.lines):
            if i:
                parts.append((" ", self.main))
            parts.extend(ln.runs)
        text, in_alt = "", False
        for t, st in parts:
            is_alt = self.alt is not None and st != self.main and t.strip() != ""
            if not t.strip():
                text += t
                continue
            if is_alt and not in_alt:
                text += ALT_OPEN; in_alt = True
            elif not is_alt and in_alt:
                text += ALT_CLOSE; in_alt = False
            text += t
        if in_alt:
            text += ALT_CLOSE
        text = re.sub(r"\s+", " ", text).strip()
        text = re.sub(rf"(\s+){re.escape(ALT_CLOSE)}", rf"{ALT_CLOSE}\1", text)
        self.text = text.replace(f"{ALT_OPEN} ", f" {ALT_OPEN}")
        self.plain = strip_marks(self.text)


def _color(c: int) -> tuple:
    return ((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255


def _style(span) -> Style:
    return Style(re.sub(r"^[A-Z]{6}\+", "", span["font"]), round(span["size"], 1), _color(span["color"]),
                 span["flags"])


def _is_bullet(items) -> bool:
    c = items[0]["c"]
    if len(items) < 2 or c.isalnum():
        return False
    nxt = items[1]
    gap = nxt["bbox"][0] - items[0]["bbox"][2]
    spaced = not nxt["c"].strip() or gap > 0.25 * items[0]["st"].size
    if c in "-–—*":
        return not nxt["c"].strip()
    symbolic = (c in BULLETS or 0xE000 <= ord(c) <= 0xF8FF
                or unicodedata.category(c) in ("So", "Sm", "Po", "Pd")
                or items[0]["st"].font.lower().startswith(("symbol", "opensymbol", "wingding", "zapf")))
    return symbolic and spaced and c not in "([{\"'#$%&@"


# ---------------------------------------------------------------- reading the template


class Template:
    def __init__(self, path: str | Path):
        require()
        self.path = str(Path(path).expanduser())
        self.doc = fitz.open(self.path)
        self.units: list[Unit] = []
        self.missing_fonts: list[str] = []
        self._extract()

    def _segments(self, pno):
        """Runs of characters on one baseline with no large horizontal gap."""
        page = self.doc[pno]
        raw = page.get_text("rawdict", flags=0)
        spans = []
        for b in raw["blocks"]:
            if b.get("type") != 0:
                continue
            for ln in b["lines"]:
                if abs(ln["dir"][1]) > 0.01:
                    continue
                for sp in ln["spans"]:
                    if sp["chars"]:
                        spans.append(sp)
        rows = {}
        for sp in spans:
            base = sp["origin"][1]
            key = next((k for k in rows if abs(k - base) < 1.2), round(base, 1))
            rows.setdefault(key, []).append(sp)
        segs = []
        for base, sps in rows.items():
            chars = []
            for sp in sorted(sps, key=lambda s: s["chars"][0]["origin"][0]):
                st = _style(sp)
                for c in sp["chars"]:  # natural order (sorting glyphs by x scrambles ligatures)
                    chars.append({"c": c["c"], "bbox": c["bbox"], "x": c["origin"][0], "st": st})
            cur = []
            for ch in chars:
                if cur and cur[-1]["c"].strip() and ch["c"].strip():
                    gap = ch["bbox"][0] - cur[-1]["bbox"][2]
                    if gap > max(2.2 * ch["st"].size, 16):
                        segs.append((base, cur)); cur = []
                elif cur and not ch["c"].strip():
                    pass
                cur.append(ch)
                # also split at a run of spaces wider than the gap threshold
                if len(cur) > 2 and not cur[-1]["c"].strip():
                    vis = [c for c in cur if c["c"].strip()]
                    if vis and ch["bbox"][2] - vis[-1]["bbox"][2] > max(2.2 * ch["st"].size, 16):
                        segs.append((base, [c for c in cur if c is not ch])); cur = []
            if cur:
                segs.append((base, cur))
        out = []
        for base, items in segs:
            while items and not items[0]["c"].strip():
                items = items[1:]
            while items and not items[-1]["c"].strip():
                items = items[:-1]
            if items:
                out.append({"base": base, "items": items, "x0": items[0]["bbox"][0], "x1": items[-1]["bbox"][2]})
        out.sort(key=lambda s: (s["base"], s["x0"]))
        return out

    @staticmethod
    def _split_prefix(items):
        if _is_bullet(items):
            k = 1
            while k < len(items) and not items[k]["c"].strip():
                k += 1
            if k < len(items):
                return items[:k], items[k:], True
        st0, k = items[0]["st"], 0
        while k < len(items) and items[k]["st"] == st0:
            k += 1
        if 0 < k < len(items):
            label = "".join(i["c"] for i in items[:k])
            if LABEL_END.search(label) and len(label) < 40:
                while k < len(items) and not items[k]["c"].strip():
                    k += 1
                return items[:k], items[k:], False
        return [], items, False

    @staticmethod
    def _runs(items):
        runs = []
        for it in items:
            if runs and runs[-1][1] == it["st"]:
                runs[-1][0] += it["c"]
            elif runs and not it["c"].strip():
                runs[-1][0] += it["c"]
            else:
                runs.append([it["c"], it["st"]])
        return [(t, s) for t, s in runs]

    @staticmethod
    def _rects(items):
        r = None
        for it in items:
            if it["c"].strip():
                bb = fitz.Rect(it["bbox"])
                r = bb if r is None else r | bb
        if r is None:
            return []
        h = r.height
        return [fitz.Rect(r.x0 - 0.3, r.y0 + 0.22 * h, r.x1 + 0.3, r.y1 - 0.22 * h)]

    def _extract(self):
        for pno, page in enumerate(self.doc):
            segs = self._segments(pno)
            pw = page.rect.width
            units = []
            for s in segs:
                prefix, body, is_bullet = self._split_prefix(s["items"])
                if not body:
                    continue
                line = Line(s["base"], body[0]["x"], body[-1]["bbox"][2], self._rects(body), self._runs(body))
                sizes = {it["st"].size for it in body if it["c"].strip()}
                fonts = {it["st"].font for it in body if it["c"].strip()}
                target = None
                if not prefix:
                    for u in reversed(units):
                        last = u.lines[-1]
                        size = max(st.size for _, st in last.runs)
                        gap = s["base"] - last.baseline
                        starts = [u.lines[0].x_start, u._seg_x0] + ([u.lines[1].x_start] if len(u.lines) > 1 else [])
                        u_sizes = {st.size for ln in u.lines for _, st in ln.runs}
                        u_fonts = {st.font for ln in u.lines for _, st in ln.runs}
                        if (0.6 * size < gap < 1.75 * size and any(abs(line.x_start - x) < 2.5 for x in starts)
                                and sizes & u_sizes and fonts & u_fonts
                                and body[0]["st"].font in u_fonts):
                            target = u
                            break
                if target:
                    target.lines.append(line)
                    continue
                u = Unit("", pno, [line], prefix="".join(i["c"] for i in prefix), bullet=is_bullet)
                u._seg_x0 = s["x0"]
                units.append(u)

            units = self._split_short_lines(units)
            for u in units:
                u.finalize()
                self._right_edge(u, segs, pw)
            self.units.extend(units)
        self._lock()
        self.units = self._reading_order(self.units)
        for i, u in enumerate(self.units, 1):
            u.id = f"u{i}"

    @staticmethod
    def _split_short_lines(units):
        """A line that ends well before the column edge ends its paragraph (e.g. one-line skills rows)."""
        out = []
        for u in units:
            right = max(l.x1 for l in u.lines)
            cur = Unit("", u.page, [u.lines[0]], u.prefix, u.bullet)
            cur._seg_x0 = u._seg_x0
            for prev, ln in zip(u.lines, u.lines[1:]):
                size = max(st.size for _, st in prev.runs)
                if prev.x1 < right - max(6 * size, 0.18 * (right - u.lines[0].x_start)):
                    out.append(cur)
                    cur = Unit("", u.page, [ln])
                    cur._seg_x0 = ln.x_start
                else:
                    cur.lines.append(ln)
            out.append(cur)
        return out

    @staticmethod
    def _right_edge(u, segs, pw):
        left = u.lines[0].x_start
        size = u.main.size
        same_col = [s for s in segs if abs(s["x0"] - left) < 3 * size or abs(s["x0"] - u._seg_x0) < 3 * size]
        right = max([s["x1"] for s in same_col] + [l.x1 for l in u.lines])
        top, bottom = u.lines[0].baseline - size, u.lines[-1].baseline + 2
        mine = max(l.x1 for l in u.lines)
        neighbours = [s["x0"] for s in segs if top <= s["base"] <= bottom and s["x0"] > mine + 4]
        if neighbours:
            right = min(right, min(neighbours) - 8)
        u.right = max(right, mine)
        ln = u.lines[0]
        c = (ln.x_start + ln.x1) / 2
        if len(u.lines) == 1 and abs(c - pw / 2) < 4 and ln.x_start > pw * 0.2:
            u.align = "center"
            content_left = min(s["x0"] for s in segs)
            content_right = max(s["x1"] for s in segs)
            half = max(min(c - content_left, content_right - c), (ln.x1 - ln.x_start) / 2)
            u.right = c + half
            u.center_left = c - half

    def _lock(self):
        """Never change: contact details, dates, names, job titles, companies, locations, headings."""
        if not self.units:
            return
        for u in self.units:
            p = u.plain
            u.protected = bool(
                CONTACT.search(p)
                or (DATE.search(p) and len(p) < 60)
                or (DATE.search(p) and not u.bullet and not u.prefix.strip() and len(p) < 250)  # education
                or (not u.bullet and not u.prefix.strip() and len(p) < 45))
        first = [u for u in self.units if u.page == 0] or self.units
        max(first, key=lambda u: u.main.size).protected = True
        rows = {(u.page, round(l.baseline)) for u in self.units if u.protected for l in u.lines}
        for u in self.units:
            if not u.bullet and len(u.lines) == 1 and any(
                    (u.page, round(u.lines[0].baseline) + d) in rows for d in (-1, 0, 1)) and len(u.plain) < 90:
                u.protected = True

    @staticmethod
    def _reading_order(units):
        out = []
        for pno in sorted({u.page for u in units}):
            us = [u for u in units if u.page == pno]
            xs = sorted({round(u.x0) for u in us})
            clusters = []
            for x in xs:
                if clusters and x - clusters[-1][-1] < 60:
                    clusters[-1].append(x)
                else:
                    clusters.append([x])

            def col(u):
                return next((i for i, c in enumerate(clusters) if c[0] - 1 <= round(u.x0) <= c[-1] + 1), 0)

            counts = {}
            for u in us:
                counts[col(u)] = counts.get(col(u), 0) + 1
            # a real column has several multi-word paragraphs; right-aligned dates etc. are not columns
            real = {c for c in counts if sum(1 for u in us if col(u) == c and not u.protected) >= 3} | {0}

            def key(u):
                c = col(u)
                while c not in real and c > 0:
                    c -= 1
                return (c, u.lines[0].baseline, u.x0)

            out += sorted(us, key=key)
        return out

    # ---------------------------------------------------------------- info for Claude

    def _cont_x(self, u: Unit):
        return u.lines[1].x_start if len(u.lines) > 1 else u.lines[0].x_start

    def capacity(self, u: Unit) -> int:
        used = sum(l.x1 - l.x_start for l in u.lines)
        avg = used / max(len(u.plain), 1) if used > 0 else u.main.size * 0.5
        if u.align == "center":
            return max(int((u.right - u.center_left) / avg * 0.97), len(u.plain))
        avail = sum(u.right - (l.x_start if i == 0 else self._cont_x(u)) for i, l in enumerate(u.lines))
        return max(int(avail / avg * 0.97), len(u.plain))

    def listing(self) -> str:
        rows = []
        for u in self.units:
            kind = "bullet" if u.bullet else (f"after label '{u.prefix.strip()}'" if u.prefix.strip()
                                              else ("bold" if u.main.bold else "text"))
            cap = self.capacity(u)
            if u.protected:
                size = "LOCKED"
            elif len(u.lines) > 1:
                lo = int(cap * (len(u.lines) - 1) / len(u.lines)) + 3
                size = f"{len(u.lines)} lines: {lo}-{cap} chars"
            else:
                size = f"1 line: max {cap} chars"
            alt = " ⟦⟧=" + ("colored" if u.alt and u.alt.color != u.main.color else "styled") if u.alt else ""
            rows.append(f"[{u.id}] ({kind}, {size}{alt}) {u.text}")
        return "\n".join(rows)

    def header(self) -> tuple[str, str]:
        first = [u for u in self.units if u.page == 0] or self.units
        name = max(first, key=lambda u: u.main.size).plain if first else ""
        bits = [u.plain for u in first if CONTACT.search(u.plain)][:4]
        return name, " | ".join(bits)

    def as_markdown(self, edits: dict | None = None) -> str:
        edits = edits or {}
        name, contact = self.header()
        rows = [f"# {name}", contact, ""]
        for u in self.units:
            if u.plain in (name, contact):
                continue
            t = strip_marks(edits.get(u.id, u.text) if not u.protected else u.text)
            if t.strip():
                rows.append(("- " if u.bullet else (u.prefix.strip() + " " if u.prefix.strip() else "")) + t)
        return "\n".join(rows) + "\n"

    def body_style(self) -> dict:
        weights = {}
        for u in self.units:
            k = (u.main.font, u.main.size)
            weights[k] = weights.get(k, 0) + len(u.plain)
        font, size = max(weights, key=weights.get) if weights else ("Calibri", 10.5)
        biggest = max(self.units, key=lambda u: u.main.size) if self.units else None
        family = re.split(r"[-,]", font)[0]
        family = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", family).replace("MT", "").strip()
        return {"font": family, "font_size": round(size * 2) / 2,
                "accent": biggest.main.color if biggest else (0, 0, 0)}

    # ---------------------------------------------------------------- writing

    def render(self, edits: dict, out_path, min_scale=0.9) -> dict:
        """Apply {slot id: new text}. Returns {slot id: capacity} for texts that didn't fit."""
        doc = fitz.open(self.path)
        fonts = FontResolver(doc)
        overflow, changed = {}, []
        for u in self.units:
            if u.id not in edits or u.protected:
                continue
            new = re.sub(r"\s+", " ", str(edits[u.id] or "")).strip()
            if new == u.text or strip_marks(new) == u.plain and new.count(ALT_OPEN) == u.text.count(ALT_OPEN):
                continue
            changed.append((u, new))
        for pno in {u.page for u, _ in changed}:
            page = doc[pno]
            for u, _ in changed:
                if u.page == pno:
                    for ln in u.lines:
                        for r in ln.redact:
                            page.add_redact_annot(r, fill=False)
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                  graphics=getattr(fitz, "PDF_REDACT_LINE_ART_NONE", 0),
                                  text=getattr(fitz, "PDF_REDACT_TEXT_REMOVE", 0))
        for u, new in changed:
            page = doc[u.page]
            scale, fit = 1.0, None
            while scale >= min_scale - 1e-6 and fit is None:
                fit = self._wrap(u, new, fonts, scale)
                if fit is None:
                    scale -= 0.02
            if fit is None:
                overflow[u.id] = self.capacity(u)
                while fit is None and scale > 0.6:
                    fit = self._wrap(u, new, fonts, scale)
                    if fit is None:
                        scale -= 0.02
                if fit is None:
                    fit = self._wrap(u, new, fonts, scale, force=True)
            for x, y, tokens in fit:
                for pieces in tokens:
                    for t, st in pieces:
                        fonts.draw(page, x, y, t, st, st.size * scale)
                        x += fonts.width(t, st, st.size * scale)
                    x += fonts.width(" ", u.main, u.main.size * scale)
        doc.save(str(out_path), garbage=3, deflate=True)
        self.missing_fonts = sorted(fonts.missing)
        return overflow

    def _tokens(self, u: Unit, text: str):
        """Words as lists of (text, style) pieces, honoring ⟦alt⟧ markers."""
        alt = u.alt or u.main
        words, cur, style = [], [], u.main
        buf = ""

        def flush_piece():
            nonlocal buf
            if buf:
                cur.append((buf, style))
                buf = ""

        for ch in text:
            if ch == ALT_OPEN:
                flush_piece(); style = alt
            elif ch == ALT_CLOSE:
                flush_piece(); style = u.main
            elif ch == " ":
                flush_piece()
                if cur:
                    words.append(cur); cur = []
            else:
                buf += ch
        flush_piece()
        if cur:
            words.append(cur)
        return words

    def _wrap(self, u: Unit, text: str, fonts, scale, force=False):
        tokens = self._tokens(u, text)
        if not tokens:
            return []
        wid = lambda tok: sum(fonts.width(t, st, st.size * scale) for t, st in tok)
        space = fonts.width(" ", u.main, u.main.size * scale)
        slots = [(u.lines[0].x_start, u.lines[0].baseline)] + [(self._cont_x(u), l.baseline) for l in u.lines[1:]]
        out, i = [], 0
        for k, (x, y) in enumerate(slots):
            width = u.right - (u.center_left if u.align == "center" else x)
            line, used = [], 0.0
            while i < len(tokens):
                w = wid(tokens[i]) + (space if line else 0)
                if used + w <= width + 0.5:
                    line.append(tokens[i]); used += w; i += 1
                elif not line and not force:
                    return None
                elif not line:
                    line.append(tokens[i]); used += w; i += 1
                else:
                    break
            if k == len(slots) - 1 and i < len(tokens):
                if not force:
                    return None
                line += tokens[i:]; i = len(tokens)
            if line:
                if u.align == "center":
                    x = (u.lines[0].x_start + u.lines[0].x1) / 2 - used / 2
                out.append((x, y, line))
        return out


# ---------------------------------------------------------------- fonts


def _norm(name: str) -> str:
    n = re.sub(r"^[A-Z]{6}\+", "", name)
    n = re.sub(r"(PSMT|MT|PS)$", "", n.replace(",", "-"))
    return re.sub(r"[^a-z0-9]", "", n.lower())


# metric-compatible equivalents: same character widths, near-identical look
EQUIVALENTS = {
    "liberationsans": ["arial", "helvetica", "arimo", "helveticaneue"],
    "arimo": ["arial", "liberationsans", "helvetica"],
    "arial": ["liberationsans", "arimo", "helvetica"],
    "helvetica": ["arial", "liberationsans", "helveticaneue"],
    "liberationserif": ["timesnewroman", "times", "tinos"],
    "tinos": ["timesnewroman", "liberationserif", "times"],
    "timesnewroman": ["liberationserif", "tinos", "times"],
    "liberationmono": ["couriernew", "cousine", "courier"],
    "couriernew": ["liberationmono", "cousine", "courier"],
    "carlito": ["calibri"], "calibri": ["carlito"],
    "caladea": ["cambria"], "cambria": ["caladea"],
}
STYLE_WORDS = r"(bold|italic|oblique|regular|semibold|medium|light|roman|book|bolditalic|boldoblique)+$"

FONT_DIRS = ["~/Library/Fonts", "/Library/Fonts", "/System/Library/Fonts", "/System/Library/Fonts/Supplemental",
             "C:/Windows/Fonts", "~/AppData/Local/Microsoft/Windows/Fonts", "/usr/share/fonts",
             "/usr/local/share/fonts", "~/.fonts", "~/.local/share/fonts",
             "/Applications/Microsoft Word.app/Contents/Resources/DFonts"]
_CACHE = Path(__file__).resolve().parent / ".font_index.json"


def system_fonts() -> dict:
    """{normalized font name: font file}, cached after the first scan."""
    if _CACHE.exists():
        try:
            return json.loads(_CACHE.read_text())
        except Exception:
            pass
    index = {}
    for d in FONT_DIRS:
        root = Path(d).expanduser()
        if not root.exists():
            continue
        for f in root.rglob("*"):
            if f.suffix.lower() not in (".ttf", ".otf"):
                continue
            try:
                fo = fitz.Font(fontfile=str(f))
                index.setdefault(_norm(fo.name), str(f))
                index.setdefault(_norm(f.stem), str(f))
            except Exception:
                continue
    try:
        _CACHE.write_text(json.dumps(index))
    except Exception:
        pass
    return index


class FontResolver:
    """Best available font for a style: the same font installed on this computer; else a
    metric-compatible equivalent (e.g. Arial for Liberation Sans); else the copy embedded in the
    PDF (only for characters it contains); else a similar standard font."""

    def __init__(self, doc):
        self.doc = doc
        self.embedded = {}
        for page in doc:
            for f in page.get_fonts(full=True):
                key = _norm(f[3])
                if key in self.embedded:
                    continue
                try:
                    _, _, _, buf = doc.extract_font(f[0])
                    if buf:
                        self.embedded[key] = fitz.Font(fontbuffer=buf)
                except Exception:
                    pass
        self._sys = None
        self._cache = {}
        self.missing = set()
        self._aliases = {}

    def _system(self, style: Style):
        if self._sys is None:
            self._sys = system_fonts()
        key = _norm(style.font)
        family = re.sub(STYLE_WORDS, "", key)
        suffix = ("bold" if style.bold else "") + ("italic" if style.italic else "")
        names = [key] if key != family else []
        for fam in [family] + EQUIVALENTS.get(family, []):
            if suffix:
                names += [fam + suffix, fam + suffix.replace("italic", "oblique"),
                          fam + {"bold": "bd", "italic": "i", "bolditalic": "bi"}.get(suffix, "")]
            else:
                names += [fam, fam + "regular", fam + "roman"]
        for n in names:
            path = self._sys.get(n)
            if path:
                try:
                    return fitz.Font(fontfile=path), path
                except Exception:
                    pass
        return None, None

    def _embedded_for(self, style: Style):
        n = _norm(style.font)
        if n in self.embedded:
            return self.embedded[n]
        base = re.sub(STYLE_WORDS, "", n)
        cands = [k for k in self.embedded if k.startswith(base) or base.startswith(re.sub(STYLE_WORDS, "", k))]
        if not cands:
            return None
        score = lambda k: (("bold" in k) == style.bold) + (("italic" in k or "oblique" in k) == style.italic)
        return self.embedded[max(cands, key=score)]

    @staticmethod
    def _fallback(style: Style):
        fam = "co" if style.mono else ("ti" if style.serif else "he")
        suf = {(1, 1): "bi", (1, 0): "bo", (0, 1): "it"}.get((int(style.bold), int(style.italic)))
        code = {"he": "helv", "ti": "tiro", "co": "cour"}[fam] if not suf else fam + suf
        return fitz.Font(code), code

    def chain(self, style: Style):
        key = (style.font, style.flags)
        if key not in self._cache:
            c = []
            f, path = self._system(style)
            if f:
                c.append((f, "file", path))
            emb = self._embedded_for(style)
            if emb:
                c.append((emb, "buffer", emb))
            fb, code = self._fallback(style)
            c.append((fb, "base14", code))
            self._cache[key] = c
        return self._cache[key]

    def _runs(self, text, style):
        chain = self.chain(style)
        runs = []
        for ch in text:
            idx = next((i for i, (f, _, _) in enumerate(chain) if f.has_glyph(ord(ch))), len(chain) - 1)
            if chain[idx][1] == "base14" and ch.strip():
                self.missing.add(re.sub(r"^[A-Z]{6}\+", "", style.font))
            if runs and runs[-1][0] == idx:
                runs[-1][1] += ch
            else:
                runs.append([idx, ch])
        return [(chain[i], t) for i, t in runs]

    def width(self, text, style, size):
        return sum(f.text_length(t, fontsize=size) for (f, _, _), t in self._runs(text, style))

    def draw(self, page, x, y, text, style, size):
        for (font, kind, ref), t in self._runs(text, style):
            if kind == "base14":
                alias = ref
            else:
                k = (id(page.parent), page.number, id(font))
                alias = self._aliases.get(k)
                if not alias:
                    alias = f"RT{len(self._aliases)}"
                    if kind == "file":
                        page.insert_font(fontname=alias, fontfile=ref)
                    else:
                        page.insert_font(fontname=alias, fontbuffer=font.buffer)
                    self._aliases[k] = alias
            if t.strip():
                page.insert_text((x, y), t, fontname=alias, fontsize=size, color=style.color)
            x += font.text_length(t, fontsize=size)


# ---------------------------------------------------------------- cover letter PDF (no LibreOffice needed)


def letter_styles(tpl: "Template | None", cfg: dict) -> tuple[Style, Style]:
    """(body style, name style) for the letter: taken from the resume template when there is one."""
    if tpl and tpl.units:
        weights = {}
        for u in tpl.units:
            if not u.style.bold:
                weights[id(u.style)] = (weights.get(id(u.style), (0, u.style))[0] + len(u.text), u.style)
        body = max(weights.values(), key=lambda t: t[0])[1] if weights else tpl.units[0].style
        first = [u for u in tpl.units if u.page == 0] or tpl.units
        name = max(first, key=lambda u: u.style.size).style
        return body, name
    font = cfg.get("font", "Calibri")
    accent = tuple(cfg.get("accent") or (0, 0, 0))
    return (Style(font, float(cfg.get("font_size", 10.5)), (0, 0, 0), 0),
            Style(font + " Bold", float(cfg.get("font_size", 10.5)) + 8, accent, 16))


def write_letter_pdf(path, letter: str, name: str, contact: str, company: str, date_text: str,
                     tpl: "Template | None", cfg: dict):
    require()
    body, name_st = letter_styles(tpl, cfg)
    src = fitz.open(tpl.path) if tpl else fitz.open()
    fonts = FontResolver(src)
    pw, ph = (tpl.doc[0].rect.width, tpl.doc[0].rect.height) if tpl else (612, 792)
    out = fitz.open()
    margin_x, top, bottom = 72, 64, 64
    size = min(max(body.size + 0.5, 10.5), 12)
    lead = size * 1.42
    body = Style(body.font, size, body.color, body.flags & ~16 & ~2)
    contact_st = Style(body.font, size - 1.5, (0.35, 0.35, 0.35), body.flags)
    name_size = min(max(name_st.size, 18), 26)
    accent = name_st.color if any(name_st.color) else (0, 0, 0)

    page = out.new_page(width=pw, height=ph)
    y = top + name_size

    def centered(text, st, sz, y):
        w = fonts.width(text, st, sz)
        fonts.draw(page, (pw - w) / 2, y, text, st, sz)

    centered(name, Style(name_st.font, name_size, accent, name_st.flags), name_size, y)
    y += name_size * 0.55 + contact_st.size * 1.3
    if contact:
        centered(contact, contact_st, contact_st.size, y)
        y += contact_st.size * 0.9
    page.draw_line((margin_x, y + 6), (pw - margin_x, y + 6), color=accent, width=0.8)
    y += 6 + lead * 1.6

    def new_page():
        nonlocal page, y
        page = out.new_page(width=pw, height=ph)
        y = top + size

    def para(text, after=0.6):
        nonlocal y
        width = pw - 2 * margin_x
        words, line = text.split(), []
        for w in words:
            trial = " ".join(line + [w])
            if line and fonts.width(trial, body, size) > width:
                if y > ph - bottom:
                    new_page()
                fonts.draw(page, margin_x, y, " ".join(line), body, size)
                y += lead
                line = [w]
            else:
                line.append(w)
        if line:
            if y > ph - bottom:
                new_page()
            fonts.draw(page, margin_x, y, " ".join(line), body, size)
            y += lead
        y += lead * after

    para(date_text, 0)
    if company and company.lower() != "unknown":
        para(company, 0)
    y += lead * 0.8
    for block in re.split(r"\n\s*\n", letter.strip()):
        lines = [l.strip() for l in block.splitlines() if l.strip()]
        if not lines:
            continue
        if re.match(r"(?i)(sincerely|best|regards|kind regards|thank you|warmly|respectfully)", lines[0]):
            for l in lines:  # sign-off keeps its line breaks
                para(l, 0)
        else:
            para(" ".join(lines))
    out.save(str(path), garbage=3, deflate=True)
    return sorted(fonts.missing)


def markdown_resume_pdf(md: str, path, cfg: dict):
    """Standard-layout resume straight to PDF (used when LibreOffice/Word aren't available)."""
    require()
    import html as H

    def inline(t):
        t = H.escape(re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", t))
        t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
        return re.sub(r"\*(.+?)\*", r"<i>\1</i>", t)

    lines, out, in_list, first = md.strip().splitlines(), [], False, True
    for raw in lines:
        l = raw.rstrip()
        bullet = re.match(r"^\s*[-*•] (.*)", l)
        if in_list and not bullet:
            out.append("</ul>"); in_list = False
        if not l.strip():
            continue
        if l.startswith("# "):
            out.append(f"<h1>{inline(l[2:])}</h1>")
        elif first and out and out[-1].startswith("<h1>"):
            out.append(f"<p class='contact'>{inline(l)}</p>")
        elif l.startswith("## "):
            out.append(f"<h2>{inline(l[3:].upper())}</h2>")
        elif l.startswith("### "):
            parts = l[4:].rsplit(" | ", 1)
            left = inline(parts[0].replace(" | ", "  ·  "))
            right = f"<span class='date'>{inline(parts[1])}</span>" if len(parts) == 2 else ""
            out.append(f"<table class='job'><tr><td><b>{left}</b></td><td class='r'>{right}</td></tr></table>")
        elif bullet:
            if not in_list:
                out.append("<ul>"); in_list = True
            out.append(f"<li>{inline(bullet.group(1))}</li>")
        else:
            out.append(f"<p>{inline(l)}</p>")
        first = False
    if in_list:
        out.append("</ul>")
    size = float(cfg.get("font_size", 10.5))
    css = f"""* {{font-family: sans-serif; font-size: {size}pt; color: #222;}}
h1 {{font-size: {size + 8}pt; text-align: center; margin: 0 0 2pt 0; color:#111;}}
p.contact {{text-align: center; font-size: {size - 1}pt; color: #444; margin: 0 0 8pt 0;}}
h2 {{font-size: {size + 1}pt; border-bottom: 0.8pt solid #444; margin: 9pt 0 3pt 0; padding-bottom: 1pt;}}
p {{margin: 0 0 2pt 0;}} ul {{margin: 1pt 0 2pt 0; padding-left: 14pt;}} li {{margin: 0 0 1pt 0;}}
table.job {{width: 100%; margin: 4pt 0 1pt 0;}} td.r {{text-align: right;}} .date {{font-style: italic;}}"""
    story = fitz.Story(html="".join(out), user_css=css)
    writer = fitz.DocumentWriter(str(path))
    rect, where = fitz.paper_rect("letter"), fitz.paper_rect("letter") + (50, 43, -50, -43)
    more = 1
    while more:
        dev = writer.begin_page(rect)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()

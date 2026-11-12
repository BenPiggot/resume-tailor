#!/usr/bin/env python3
"""
resume-tailor: tailor a base resume (and optionally a cover letter) to a job
description with Claude, then save .docx/.pdf/.md versions into a folder.

Usage:
  python tailor.py import my_resume.docx            # one-time: convert Word resume -> base_resume.md
  python tailor.py job posting.txt                  # tailor resume, auto-named output folder
  python tailor.py job https://jobs.example.com/123 --cover-letter --out ~/Jobs/Acme
  python tailor.py job - --cover-letter             # paste the description (end with Ctrl-D)
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import os
import re
import shutil
import subprocess
import sys
import tomllib
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_CONFIG = {
    "base_resume": str(HERE / "base_resume.md"),
    "output_root": "~/Job Applications",
    "model": "claude-sonnet-5-5",
    "formats": ["pdf"],
    "cover_letter": "auto",
    "match_layout": True,
    "template_pdf": "",
    "folder_name": "{date}_{company}_{role}",
    "file_name": "{name}_{company}_{doc}",
    "date_format": "%Y-%m-%d",
    "space": "_",
    "resume_label": "Resume",
    "cover_letter_label": "Cover Letter",
    "font": "Calibri",
    "font_size": 10.5,
    "extra_instructions": "",
}

# ---------------------------------------------------------------- config


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    path = HERE / "config.toml"
    if path.exists():
        with open(path, "rb") as f:
            cfg.update(tomllib.load(f))
    bp = Path(cfg["base_resume"]).expanduser()
    cfg["base_resume"] = str(bp if bp.is_absolute() else HERE / bp)  # relative to the tool's folder
    return cfg


# ---------------------------------------------------------------- job input


def read_job(source: str) -> str:
    if source == "-":
        if sys.stdin.isatty():
            print("Paste the job description, then press Ctrl-D (Ctrl-Z then Enter on Windows):")
        return sys.stdin.read().strip()
    if re.match(r"https?://", source):
        return fetch_job(source)
    p = Path(source).expanduser()
    if not p.exists():
        sys.exit(f"Job description not found: {p}")
    if p.suffix.lower() == ".docx":
        return docx_to_markdown(p)
    if p.suffix.lower() == ".pdf":
        return pdf_to_text(p)
    return p.read_text(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------- fetching listings from the web

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")


def http_get(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode(r.headers.get_content_charset() or "utf-8", "replace")


def html_to_text(raw: str) -> str:
    raw = re.sub(r"(?is)<(script|style|noscript|svg|nav|footer|header).*?</\1>", " ", raw)
    raw = re.sub(r"(?i)<li[^>]*>", "\n- ", raw)
    raw = re.sub(r"(?i)<br\s*/?>|</(p|div|li|ul|ol|h\d|tr|section)>", "\n", raw)
    text = html.unescape(re.sub(r"<[^>]+>", " ", raw))
    text = re.sub(r"[ \t\xa0]+", " ", text)
    text = "\n".join(l.strip() for l in text.splitlines())
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _fmt_listing(title="", company="", location="", body="", url="") -> str:
    head = [f"Job title: {title}" if title else "", f"Company: {company}" if company else "",
            f"Location: {location}" if location else "", f"Source: {url}" if url else ""]
    return "\n".join(h for h in head if h) + "\n\n" + body.strip()


def fetch_from_ats(url: str) -> str | None:
    """Use the public posting APIs of the common applicant-tracking systems when the URL is one of theirs."""
    import json
    from urllib.parse import parse_qs, urlparse

    u = urlparse(url)
    host, parts, q = u.netloc.lower(), [p for p in u.path.split("/") if p], parse_qs(u.query)

    # Greenhouse: boards.greenhouse.io/{board}/jobs/{id}, job-boards.greenhouse.io/{board}/jobs/{id},
    #             .../embed/job_app?for={board}&token={id}
    if "greenhouse.io" in host:
        board = q.get("for", [None])[0] or (parts[0] if parts else None)
        job_id = q.get("token", [None])[0] or (parts[parts.index("jobs") + 1] if "jobs" in parts[:-1] else None)
        if board and job_id:
            j = json.loads(http_get(f"https://boards-api.greenhouse.io/v1/boards/{board}/jobs/{job_id}"))
            try:
                company = json.loads(http_get(f"https://boards-api.greenhouse.io/v1/boards/{board}")).get("name", board)
            except Exception:
                company = board
            body = html_to_text(html.unescape(j.get("content", "")))  # content is HTML-escaped HTML
            return _fmt_listing(j.get("title", ""), company, (j.get("location") or {}).get("name", ""),
                                body, j.get("absolute_url", url))

    # Lever: jobs.lever.co/{company}/{id}  (EU: jobs.eu.lever.co)
    if "lever.co" in host and len(parts) >= 2:
        api = "api.eu.lever.co" if ".eu." in host else "api.lever.co"
        j = json.loads(http_get(f"https://{api}/v0/postings/{parts[0]}/{parts[1]}?mode=json"))
        sections = [j.get("descriptionPlain") or html_to_text(j.get("description", ""))]
        for lst in j.get("lists", []):
            sections.append(f"{lst.get('text', '')}\n{html_to_text(lst.get('content', ''))}")
        sections.append(j.get("additionalPlain") or html_to_text(j.get("additional", "")))
        cats = j.get("categories") or {}
        return _fmt_listing(j.get("text", ""), parts[0], cats.get("location", ""),
                            "\n\n".join(s for s in sections if s), j.get("hostedUrl", url))

    # Ashby: jobs.ashbyhq.com/{board}/{id}
    if "ashbyhq.com" in host and len(parts) >= 2:
        data = json.loads(http_get(f"https://api.ashbyhq.com/posting-api/job-board/{parts[0]}?includeCompensation=true"))
        for j in data.get("jobs", []):
            if j.get("id") == parts[1] or parts[1] in (j.get("jobUrl") or ""):
                body = j.get("descriptionPlain") or html_to_text(j.get("descriptionHtml", ""))
                comp = (j.get("compensation") or {}).get("compensationTierSummary") or ""
                if comp:
                    body += f"\n\nCompensation: {comp}"
                return _fmt_listing(j.get("title", ""), parts[0], j.get("location", ""), body, j.get("jobUrl", url))
    return None


def extract_jsonld_job(raw_html: str, url: str) -> str | None:
    """Most career sites (Workday, iCIMS, SmartRecruiters, LinkedIn public pages, Indeed…) embed a
    schema.org JobPosting block for Google Jobs. It's the cleanest source when present."""
    import json

    for block in re.findall(r'(?is)<script[^>]+application/ld\+json[^>]*>(.*?)</script>', raw_html):
        try:
            data = json.loads(block.strip())
        except Exception:
            continue
        stack = data if isinstance(data, list) else [data]
        while stack:
            d = stack.pop()
            if isinstance(d, list):
                stack.extend(d); continue
            if not isinstance(d, dict):
                continue
            if "@graph" in d:
                stack.extend(d["@graph"])
            t = d.get("@type")
            if t == "JobPosting" or (isinstance(t, list) and "JobPosting" in t):
                org = d.get("hiringOrganization") or {}
                org = org.get("name", "") if isinstance(org, dict) else str(org)
                loc = d.get("jobLocation") or {}
                loc = loc[0] if isinstance(loc, list) and loc else loc
                addr = (loc.get("address") or {}) if isinstance(loc, dict) else {}
                place = ", ".join(x for x in [addr.get("addressLocality"), addr.get("addressRegion")] if x) \
                    if isinstance(addr, dict) else ""
                if d.get("jobLocationType") == "TELECOMMUTE":
                    place = (place + " (Remote)").strip()
                body = html_to_text(html.unescape(d.get("description", "")))
                if len(body) > 200:
                    return _fmt_listing(d.get("title", ""), org, place, body, url)
    return None


def render_with_browser(url: str) -> str | None:
    """Optional: for pages that only load the listing with JavaScript.
    Enable with:  pip install playwright && playwright install chromium"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None
    print("  Page needs JavaScript; rendering it in a headless browser…", flush=True)
    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page(user_agent=UA)
        page.goto(url, wait_until="networkidle", timeout=45000)
        content = page.content()
        b.close()
    return content


def fetch_job(url: str) -> str:
    print(f"Fetching listing: {url}", flush=True)
    try:
        if (t := fetch_from_ats(url)) and len(t) > 300:
            return t[:40000]
    except Exception as e:
        print(f"  (job-board API lookup failed: {e}; reading the page instead)")
    try:
        raw = http_get(url)
    except Exception as e:
        raw = ""
        print(f"  Couldn't download the page directly ({e}).")
    if raw and (t := extract_jsonld_job(raw, url)):
        return t[:40000]
    text = html_to_text(raw) if raw else ""
    if len(text) < 800:  # little or no listing text: probably JavaScript-rendered or blocked
        try:
            rendered = render_with_browser(url)
        except Exception as e:
            rendered = None
            print(f"  (headless browser failed: {str(e).splitlines()[0]})")
        if rendered:
            if t := extract_jsonld_job(rendered, url):
                return t[:40000]
            text = html_to_text(rendered)
    if len(text) < 300:
        sys.exit("Couldn't read the listing from that page. It may need a login (e.g. LinkedIn while "
                 "signed in) or JavaScript.\n  • For JavaScript pages:  pip install playwright && "
                 "playwright install chromium\n  • Or copy the listing text into a .txt file, or use  job -  to paste it.")
    return _fmt_listing(body=text, url=url)[:40000]


def pdf_to_text(p: Path) -> str:
    if shutil.which("pdftotext"):
        return subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True).stdout
    try:
        from pypdf import PdfReader
    except ImportError:
        sys.exit("To read PDFs run: pip install pypdf")
    return "\n".join(page.extract_text() or "" for page in PdfReader(str(p)).pages)


# ---------------------------------------------------------------- Claude

SYSTEM_PROMPT = """You are an expert resume writer and career coach. You tailor a candidate's \
base resume to a specific job posting, and optionally write a cover letter.

Hard rules (never break these):
- Use ONLY facts present in the base resume. Never invent employers, titles, dates, degrees, \
certifications, tools, metrics or accomplishments. You may rephrase, reorder, merge, trim, and \
choose emphasis. If the posting asks for something the candidate lacks, do not claim it; list it \
under gaps in the notes instead.
- Keep names, contact details, employers, job titles, dates and education exactly as written.
- Mirror the posting's terminology where it truthfully applies (helps applicant tracking systems).
- Put the most relevant experience and bullets first within each section; cut or shorten the \
least relevant. Aim for one page unless the base resume is clearly senior/multi-page.
- Strong action verbs, quantified results where the base resume has numbers, no first person, \
no buzzword filler.

Resume output format (Markdown, follow exactly so it can be rendered to Word):
# Full Name
contact line (email | phone | city | links), copied from the base resume
## Section Heading            (e.g. Summary, Skills, Experience, Education)
### Job Title, Company | Location | Dates      (the LAST " | " part is right-aligned, so put dates last)
- bullet
Plain lines are paragraphs. You may use **bold** and *italic*. No tables, no images, no HTML.

Whether to write a cover letter is given as YES, NO or AUTO. For AUTO, write one only if the \
posting asks for or accepts one (mentions a cover letter, letter of interest, motivation letter, \
or asks applicants to explain why they're interested); otherwise leave it empty. Always explain \
the decision in <cover_letter_decision>.

Cover letter (when written): 250-400 words, plain paragraphs, no header block or \
address lines (those are added automatically), start with "Dear Hiring Manager," unless a \
name is given in the posting, end with "Sincerely," then the candidate's name. Specific, warm, \
confident; connect 2-3 concrete achievements to the role's top needs. Never invent facts.

Reply with exactly these XML blocks and nothing else:
<company>Company name from the posting, or Unknown</company>
<role>Job title from the posting</role>
<resume>...tailored resume markdown...</resume>
<cover_letter_decision>"written" or "skipped", then a short reason</cover_letter_decision>
<cover_letter>...cover letter text, or empty if not written...</cover_letter>
<notes>Markdown: top requirements from the posting, what you emphasized and why, \
keywords matched, and honest gaps the candidate may want to address.</notes>"""


def _ask(system: str, user: str, cfg: dict, label: str) -> str:
    try:
        import anthropic
    except ImportError:
        sys.exit("Missing dependency: pip install -r requirements.txt")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        sys.exit("Set your API key first:  export ANTHROPIC_API_KEY=sk-ant-...  "
                 "(get one at https://console.anthropic.com)")
    print(f"{label} with {cfg['model']}…", flush=True)
    with anthropic.Anthropic().messages.stream(
        model=cfg["model"], max_tokens=16000, system=system,
        messages=[{"role": "user", "content": user}],
    ) as stream:
        msg = stream.get_final_message()
    return "".join(b.text for b in msg.content if b.type == "text")


LAYOUT_PROMPT = """You are an expert resume writer. The candidate's resume is a designed PDF whose \
layout must stay exactly as it is. It has been split into numbered text slots. You tailor the \
resume to a job posting by rewriting the text of some slots; the new text is typeset into the \
same place, font and size.

Hard rules:
- Use ONLY facts in the resume (and the master resume, if one is given). Never invent employers, \
titles, dates, degrees, tools, metrics or accomplishments. Rephrase, reorder and choose emphasis.
- Never change LOCKED slots (name, contact details, job titles, companies, dates). Leave section \
headings unchanged.
- Respect each slot's length: stay within its max characters, and for multi-line slots use at \
least the minimum so the lines are filled (an empty line looks broken).
- A slot's text replaces only that slot. To reorder bullets within a job, write bullet A's text \
into bullet B's slot and vice versa, keeping each within its slot's limits. Keep bullets under \
the job they belong to; never move achievements between employers.
- Don't empty a slot unless it's irrelevant filler; if you do, return "" for it.
- Mirror the posting's terminology where it truthfully applies. No first person, no filler.
- Text inside ⟦ ⟧ is printed in the slot's secondary style (usually colored tech tags such as \
⟦(React, TypeScript)⟧). Keep that convention: when a slot has ⟦ ⟧ text, keep the tag in ⟦ ⟧ at \
the same place (usually the end, before the final period), listing only technologies actually \
used in that work. Never put ⟦ ⟧ in slots that don't have it. The tag counts toward the length.

Whether to write a cover letter is given as YES, NO or AUTO. For AUTO, write one only if the \
posting asks for or accepts one (mentions a cover letter, letter of interest, motivation letter, \
or asks why the applicant is interested). Cover letter: 250-400 words, plain paragraphs, no \
address block, start "Dear Hiring Manager," unless a name is given, end "Sincerely," and the \
candidate's name. Never invent facts.

Reply with exactly these XML blocks and nothing else:
<company>Company name from the posting, or Unknown</company>
<role>Job title from the posting</role>
<edits>JSON object mapping slot ids to new text, ONLY for slots you change, e.g. {"u7": "..."}</edits>
<cover_letter_decision>"written" or "skipped", then a short reason</cover_letter_decision>
<cover_letter>...or empty...</cover_letter>
<notes>Markdown: top requirements from the posting, what you emphasized and why, keywords \
matched, and honest gaps.</notes>"""


def _json_obj(text: str) -> dict:
    import json
    text = text.strip()
    text = re.sub(r"^```(?:json)?|```$", "", text).strip()
    try:
        d = json.loads(text)
        return {str(k): str(v) for k, v in d.items()} if isinstance(d, dict) else {}
    except Exception:
        m = re.search(r"\{.*\}", text, re.S)
        if m:
            try:
                return {str(k): str(v) for k, v in json.loads(m.group(0)).items()}
            except Exception:
                pass
    return {}


def call_claude_layout(tpl, master: str | None, job: str, cover_letter: str, cfg: dict, extra: str) -> dict:
    instructions = "\n".join(x for x in [cfg.get("extra_instructions", ""), extra] if x)
    user = (
        (f"<master_resume>\n{master}\n</master_resume>\n(Extra material you may draw from, "
         "e.g. bullets not on the PDF.)\n\n" if master else "")
        + f"<resume_slots>\n{tpl.listing()}\n</resume_slots>\n\n"
        f"<job_posting>\n{job}\n</job_posting>\n\n"
        f"Write a cover letter: {cover_letter}\n"
        + (f"\nAdditional instructions from the candidate:\n{instructions}\n" if instructions else "")
    )
    text = _ask(LAYOUT_PROMPT, user, cfg, "Tailoring (keeping your PDF layout)")
    res = parse_response(text, require_resume=False)
    m = re.search(r"<edits>(.*?)</edits>", text, re.S)
    res["edits"] = _json_obj(m.group(1)) if m else {}
    if not m:
        sys.exit("Claude's reply didn't contain the edits. Raw reply:\n\n" + text[:2000])
    known = {u.id for u in tpl.units}
    locked = {u.id for u in tpl.units if u.protected}
    dropped = [k for k in res["edits"] if k in locked]
    res["edits"] = {k: v for k, v in res["edits"].items() if k in known and k not in locked}
    if dropped:
        print(f"  (kept locked text unchanged: {', '.join(dropped)})")
    res["resume"] = tpl.as_markdown(res["edits"])
    return res


def shorten_overflow(tpl, edits: dict, overflow: dict, cfg: dict) -> dict:
    by_id = {u.id: u for u in tpl.units}
    items = "\n".join(f"[{k}] (max {cap} chars) {edits[k]}" for k, cap in overflow.items())
    text = _ask("You shorten resume lines to fit a fixed layout. Keep every fact and keyword that "
                "matters, never add facts. Keep any ⟦...⟧ tag markers (they mark colored text). "
                "Reply with only a JSON object mapping each id to its shortened text.",
                f"These texts are too long for their slots. Rewrite each to at most its max "
                f"characters:\n{items}", cfg, "Fitting text to your layout")
    fixed = _json_obj(text)
    return {**edits, **{k: v for k, v in fixed.items() if k in by_id and k in overflow}}


def call_claude(base_resume: str, job: str, cover_letter: str, cfg: dict, extra: str,
                letter_only: bool = False) -> dict:
    """cover_letter is "YES", "NO" or "AUTO"."""
    instructions = "\n".join(x for x in [cfg.get("extra_instructions", ""), extra] if x)
    user = (
        f"<base_resume>\n{base_resume}\n</base_resume>\n\n"
        f"<job_posting>\n{job}\n</job_posting>\n\n"
        f"Write a cover letter: {cover_letter}\n"
        + ("\nOnly the cover letter is needed this time: leave <resume> empty. The resume above is "
           "already tailored to this job.\n" if letter_only else "")
        + (f"\nAdditional instructions from the candidate:\n{instructions}\n" if instructions else "")
    )
    text = _ask(SYSTEM_PROMPT, user, cfg, "Writing cover letter" if letter_only else "Tailoring")
    return parse_response(text, require_resume=not letter_only)


def parse_response(text: str, require_resume: bool = True) -> dict:
    def tag(name):
        m = re.search(rf"<{name}>(.*?)</{name}>", text, re.S)
        return m.group(1).strip() if m else ""

    out = {k: tag(k) for k in ["company", "role", "resume", "cover_letter", "cover_letter_decision", "notes"]}
    if require_resume and not out["resume"]:
        sys.exit("Claude's reply didn't contain a resume. Raw reply:\n\n" + text[:2000])
    return out


# ---------------------------------------------------------------- Markdown -> docx


def _add_runs(par, text: str, base_bold=False, base_italic=False):
    # **bold**, *italic*, [text](url) -> text
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1", text)
    for part in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            r = par.add_run(part[2:-2]); r.bold = True; r.italic = base_italic
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            r = par.add_run(part[1:-1]); r.italic = True; r.bold = base_bold
        else:
            r = par.add_run(part); r.bold = base_bold; r.italic = base_italic


def _new_document(cfg: dict):
    from docx import Document
    from docx.shared import Inches, Pt

    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Inches(0.6)
        s.left_margin = s.right_margin = Inches(0.7)
    st = doc.styles["Normal"]
    st.font.name = cfg["font"]
    st.font.size = Pt(cfg["font_size"])
    st.element.rPr.rFonts.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", cfg["font"])
    st.paragraph_format.space_after = Pt(2)
    st.paragraph_format.space_before = Pt(0)
    return doc


def _bottom_border(par):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    pPr = par._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in {"w:val": "single", "w:sz": "6", "w:space": "1", "w:color": "444444"}.items():
        b.set(qn(k), v)
    bdr.append(b)
    pPr.append(bdr)


def _header_block(doc, name: str, contact: str, cfg: dict):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(name); r.bold = True; r.font.size = Pt(cfg["font_size"] + 8)
    if cfg.get("accent"):
        from docx.shared import RGBColor
        r.font.color.rgb = RGBColor(*[int(round(c * 255)) for c in cfg["accent"]])
    p.paragraph_format.space_after = Pt(1)
    if contact:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_runs(p, contact)
        for run in p.runs:
            run.font.size = Pt(cfg["font_size"] - 1)
        p.paragraph_format.space_after = Pt(6)


def split_header(md: str) -> tuple[str, str, list[str]]:
    lines = md.strip().splitlines()
    name, contact, i = "", "", 0
    if lines and lines[0].startswith("# "):
        name, i = lines[0][2:].strip(), 1
        while i < len(lines) and not lines[i].strip():
            i += 1
        if i < len(lines) and not lines[i].startswith("#") and not lines[i].lstrip().startswith("- "):
            contact, i = lines[i].strip(), i + 1
    return name, contact, lines[i:]


def resume_md_to_docx(md: str, path: Path, cfg: dict):
    from docx.enum.text import WD_TAB_ALIGNMENT
    from docx.shared import Inches, Pt

    doc = _new_document(cfg)
    name, contact, body = split_header(md)
    if name:
        _header_block(doc, name, contact, cfg)
    width = doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin

    for raw in body:
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("## "):
            p = doc.add_paragraph()
            r = p.add_run(line[3:].strip().upper()); r.bold = True
            r.font.size = Pt(cfg["font_size"] + 1)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            _bottom_border(p)
        elif line.startswith("### "):
            text = line[4:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.keep_with_next = True
            if " | " in text:
                left, right = text.rsplit(" | ", 1)
                left = left.replace(" | ", "  ·  ")
                p.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
                _add_runs(p, left, base_bold=True)
                p.add_run("\t")
                _add_runs(p, right, base_italic=True)
            else:
                _add_runs(p, text, base_bold=True)
        elif re.match(r"^\s*[-*•] ", line):
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.space_after = Pt(1)
            _add_runs(p, re.sub(r"^\s*[-*•] ", "", line))
        elif line.startswith("# "):
            _add_runs(doc.add_paragraph(), line[2:], base_bold=True)
        else:
            _add_runs(doc.add_paragraph(), line.strip())
    doc.save(path)


def letter_to_docx(letter: str, resume_md: str, path: Path, cfg: dict, company: str):
    from docx.shared import Pt

    doc = _new_document(cfg)
    for st in [doc.styles["Normal"]]:
        st.font.size = Pt(cfg["font_size"] + 0.5)
    name, contact, _ = split_header(resume_md)
    if name:
        _header_block(doc, name, contact, cfg)
    doc.add_paragraph(dt.date.today().strftime("%B %-d, %Y") if os.name != "nt"
                      else dt.date.today().strftime("%B %#d, %Y"))
    if company and company.lower() != "unknown":
        doc.add_paragraph(company)
    doc.add_paragraph()
    for para in re.split(r"\n\s*\n", letter.strip()):
        p = doc.add_paragraph()
        lines = [l.strip() for l in para.strip().splitlines() if l.strip()]
        if re.match(r"(?i)(sincerely|best|regards|kind regards|thank you|warmly|respectfully)", lines[0]):
            for k, l in enumerate(lines):  # keep sign-off on separate lines
                if k:
                    p.add_run().add_break()
                _add_runs(p, l)
        else:
            _add_runs(p, " ".join(lines))
        p.paragraph_format.space_after = Pt(9)
    doc.save(path)


# ---------------------------------------------------------------- docx -> pdf / markdown


def find_soffice() -> str | None:
    for c in ["soffice", "libreoffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"]:
        if shutil.which(c) or Path(c).exists():
            return shutil.which(c) or c
    return None


def docx_to_pdf(docx_path: Path) -> Path | None:
    soffice = find_soffice()
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        str(docx_path.parent), str(docx_path)],
                       check=False, capture_output=True, timeout=120)
        pdf = docx_path.with_suffix(".pdf")
        if pdf.exists():
            return pdf
    try:  # uses Microsoft Word on Windows/macOS
        from docx2pdf import convert
        convert(str(docx_path), str(docx_path.with_suffix(".pdf")))
        return docx_path.with_suffix(".pdf")
    except Exception:
        return None


def docx_to_markdown(path: Path) -> str:
    from docx import Document

    doc = Document(str(path))
    out, first = [], True
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style = (p.style.name or "").lower()
        if first:
            out.append(f"# {text}"); first = False
        elif style.startswith("heading 1") or style == "title" or (text.isupper() and len(text) < 40):
            out.append(f"\n## {text.title() if text.isupper() else text}")
        elif style.startswith("heading"):
            out.append(f"### {text}")
        elif "list" in style or text[:1] in "•▪●–-":
            out.append(f"- {text.lstrip('•▪●–- ').strip()}")
        elif len(text) < 140 and re.search(r"\b(19|20)\d\d\b|\bPresent\b", text) and not text.endswith("."):
            # looks like "Title, Company   Location   2020 – Present"
            parts = [x.strip() for x in re.split(r"\t+|\s{3,}|\s+[·|]\s+", text) if x.strip()]
            out.append("### " + " | ".join(parts))
        else:
            out.append(text)
    for t in doc.tables:  # some templates put content in tables
        for row in t.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                out.append(" | ".join(dict.fromkeys(cells)))
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------- commands


def name_fields(cfg: dict, resume_md: str, company: str, role: str) -> dict:
    full = split_header(resume_md)[0] or "Resume"
    if full.isupper():  # "BEN PIGGOT" on the resume -> "Ben_Piggot" in file names
        full = full.title()
    parts = full.split()
    return {"name": full, "first": parts[0] if parts else full, "last": parts[-1] if len(parts) > 1 else "",
            "company": company, "role": role, "date": dt.date.today().strftime(cfg["date_format"])}


def render_name(pattern: str, fields: dict, cfg: dict, **extra) -> str:
    """Fill a pattern like "{first}_{last}_{doc}_{company}". Field values have spaces replaced by
    cfg["space"]; characters that aren't allowed in file names are removed."""
    sp = cfg.get("space", "_")
    vals = {k: re.sub(r"\s+", sp, re.sub(r'[,.\'*?"<>|]+', "", re.sub(r"[\\/:&+]+", " ", str(v))).strip())[:60]
            for k, v in {**fields, **extra}.items()}
    try:
        name = pattern.format(**vals)
    except KeyError as e:
        sys.exit(f"Unknown placeholder {e} in name pattern {pattern!r}. "
                 f"Use: {{name}} {{first}} {{last}} {{company}} {{role}} {{date}} {{doc}}")
    name = re.sub(r'[\\/:*?"<>|]+', "", name)
    if sp.strip():
        name = re.sub(rf"(?:{re.escape(sp)}){{2,}}", sp, name)
    return name.strip(" _-.") or "Application"


def safe(s: str) -> str:
    return re.sub(r"[^\w\- ]+", "", s).strip().replace(" ", "_")[:60] or "Unknown"


def read_base(p: Path) -> str:
    """Base resume may be .md/.txt, .pdf, or .docx; text is extracted and Claude restructures it."""
    suf = p.suffix.lower()
    if suf == ".pdf":
        text = pdf_to_text(p)
        if len(text.strip()) < 200:
            sys.exit(f"Couldn't extract text from {p.name} (it may be a scanned image). "
                     "Export it from Word/Google Docs as a text-based PDF, or use a .docx/.md.")
        return text
    if suf == ".docx":
        return docx_to_markdown(p)
    return p.read_text(encoding="utf-8")


def write_doc(out: Path, stem: str, md: str, build, formats: list[str], cfg: dict | None = None) -> list[str]:
    """Build stem.docx (and .pdf), keep or tuck away the .md source. Returns file names written."""
    written = []
    docx_path = out / f"{stem}.docx"
    build(docx_path)
    if "docx" in formats:
        written.append(docx_path.name)
    pdf_ok = False
    if "pdf" in formats:
        pdf = docx_to_pdf(docx_path)
        if not pdf:  # no LibreOffice/Word: build the PDF directly
            try:
                import pdf_template
                pdf_template.markdown_resume_pdf(md, out / f"{stem}.pdf", cfg or DEFAULT_CONFIG)
                pdf = out / f"{stem}.pdf"
            except Exception as e:
                print(f"  ! Couldn't make a PDF ({e}).")
        if pdf:
            written.append(pdf.name); pdf_ok = True
    if "docx" not in formats:
        if pdf_ok or "pdf" not in formats:
            docx_path.unlink(missing_ok=True)
        else:  # never leave you with nothing: keep the Word file when PDF conversion fails
            written.append(docx_path.name)
            print(f"    Kept {docx_path.name} instead.")
    if "md" in formats:
        (out / f"{stem}.md").write_text(md + "\n", encoding="utf-8")
        written.append(f"{stem}.md")
    else:  # keep an editable copy out of the way for `render`
        (out / ".source").mkdir(exist_ok=True)
        (out / ".source" / f"{stem}.md").write_text(md + "\n", encoding="utf-8")
    return written


def cover_mode(args, cfg) -> str:
    if args.cover_letter:
        return "YES"
    if args.no_cover_letter:
        return "NO"
    mode = str(cfg.get("cover_letter", "auto")).upper()
    return mode if mode in ("YES", "NO", "AUTO") else "AUTO"


def load_layout_template(args, cfg, base_path: Path):
    """The PDF whose layout the tailored resume should copy, or None for the simple layout."""
    if getattr(args, "simple", False) or not cfg.get("match_layout", True):
        return None
    path = getattr(args, "template", None) or cfg.get("template_pdf") or (
        str(base_path) if base_path.suffix.lower() == ".pdf" else "")
    if not path:
        return None
    path = Path(path).expanduser().resolve()
    if not path.exists():
        sys.exit(f"Layout template not found: {path}")
    import pdf_template
    tpl = pdf_template.Template(path)
    editable = [u for u in tpl.units if not u.protected]
    if len(editable) < 4 or sum(len(u.text) for u in tpl.units) < 300:
        print(f"  ! Couldn't read enough text from {path.name} to reuse its layout (is it a scanned "
              f"image?). Using the standard layout instead.")
        return None
    return tpl


def write_layout_pdf(out: Path, stem: str, tpl, edits: dict, cfg: dict) -> list[str]:
    import json
    pdf = out / f"{stem}.pdf"
    overflow = tpl.render(edits, pdf)
    if overflow:
        edits = shorten_overflow(tpl, edits, overflow, cfg)
        overflow = tpl.render(edits, pdf)
        if overflow:
            print(f"  ! {len(overflow)} line(s) were still long and were set slightly smaller to fit.")
    if getattr(tpl, "missing_fonts", None):
        print(f"  ! Font(s) not installed on this computer: {', '.join(tpl.missing_fonts)}. Changed text "
              f"used the font embedded in your PDF or a close look-alike. Install the font(s) for an exact match.")
    (out / ".source").mkdir(exist_ok=True)
    (out / ".source" / f"{stem}.edits.json").write_text(json.dumps(
        {"template": tpl.path, "edits": edits,
         "original": {u.id: u.text for u in tpl.units if u.id in edits}}, indent=2, ensure_ascii=False),
        encoding="utf-8")
    return [pdf.name]


def cmd_job(args, cfg):
    base_path = Path(args.base or cfg["base_resume"]).expanduser()
    if not base_path.exists():
        sys.exit(f"Base resume not found at {base_path}.\n"
                 f"Create it, or run:  python tailor.py import your_resume.docx")
    tpl = load_layout_template(args, cfg, base_path)
    base = None if (tpl and Path(tpl.path) == base_path.resolve()) else read_base(base_path)
    job = read_job(args.source)
    if len(job.strip()) < 50:
        sys.exit("The job description looks empty.")

    mode = cover_mode(args, cfg)
    if tpl:
        res = call_claude_layout(tpl, base, job, mode, cfg, args.instructions or "")
        letter_cfg = {**cfg, **tpl.body_style()}  # cover letter in the same font/size/accent color
    else:
        res = call_claude(base, job, mode, cfg, args.instructions or "")
        letter_cfg = cfg
    company, role = res["company"] or "Unknown", res["role"] or "Role"

    fields = name_fields(cfg, res["resume"], company, role)
    if args.out:
        out = Path(args.out).expanduser()
    else:
        root = Path(cfg["output_root"]).expanduser()
        pattern = args.folder_name or cfg["folder_name"]  # "/" makes subfolders, e.g. "{company}/{role}"
        out = root.joinpath(*[render_name(part, fields, cfg) for part in pattern.split("/") if part])
        if out.exists():  # applied to the same role twice today: don't overwrite
            k = 2
            while (alt := out.with_name(f"{out.name}_{k}")).exists():
                k += 1
            out = alt
    out.mkdir(parents=True, exist_ok=True)

    file_pat = args.name or cfg["file_name"]
    resume_stem = render_name(file_pat, fields, cfg, doc=cfg["resume_label"])
    letter_stem = render_name(file_pat, fields, cfg, doc=cfg["cover_letter_label"])
    if resume_stem == letter_stem:
        sys.exit("Your file_name pattern needs {doc} so the resume and cover letter get different names.")
    formats = [f.lower() for f in (args.formats.split(",") if args.formats else cfg["formats"])]

    (out / "job_description.txt").write_text(job, encoding="utf-8")
    decision = res.get("cover_letter_decision", "")
    (out / "tailoring_notes.md").write_text(
        f"# {role} at {company}\n\n{res['notes']}\n"
        + (f"\n## Cover letter\n{decision}\n" if decision else ""), encoding="utf-8")
    written = ["job_description.txt", "tailoring_notes.md"]

    import json
    if tpl:
        written += write_layout_pdf(out, resume_stem, tpl, res["edits"], cfg)
        (out / ".source").mkdir(exist_ok=True)
        (out / ".source" / f"{resume_stem}.md").write_text(res["resume"], encoding="utf-8")
    else:
        written += write_doc(out, resume_stem, res["resume"],
                             lambda p: resume_md_to_docx(res["resume"], p, cfg), formats, cfg)
    (out / ".source").mkdir(exist_ok=True)
    (out / ".source" / "application.json").write_text(json.dumps(
        {"company": company, "role": role, "resume_md": f"{resume_stem}.md", "letter_stem": letter_stem,
         "letter_style": {k: letter_cfg[k] for k in ("font", "font_size", "accent") if k in letter_cfg},
         "template": tpl.path if tpl else "", "formats": formats,
         "created": dt.datetime.now().isoformat(timespec="seconds")}, indent=2), encoding="utf-8")
    if "md" in formats:  # keep a copy in .source too so `letter` always finds the tailored resume
        (out / ".source" / f"{resume_stem}.md").write_text(res["resume"] + "\n", encoding="utf-8")
    letter_written = mode != "NO" and bool(res["cover_letter"])
    if letter_written:
        written += save_letter(out, res["cover_letter"], cfg, note="first draft")

    print(f"\n✓ {role} at {company}\n  Saved to: {out}")
    for w in written:
        print(f"   - {w}")
    if mode == "AUTO":
        reason = re.sub(r"^(written|skipped)\W*", "", decision, flags=re.I) or "-"
        print(f"\n  Cover letter {'written' if letter_written else 'skipped'}: {reason}")
        if not letter_written:
            print(f"  Want one anyway?  python tailor.py letter \"{out}\"")
    if args.open:
        opener = {"darwin": "open", "win32": "explorer"}.get(sys.platform, "xdg-open")
        subprocess.run([opener, str(out)], check=False)


# ---------------------------------------------------------------- cover letters: versions & revisions


def find_app(arg: str, cfg: dict) -> Path:
    """Accept a full path, a folder name, part of a name ("atlassian"), or "latest"."""
    p = Path(arg).expanduser()
    if (p / "job_description.txt").exists():
        return p
    root = Path(cfg["output_root"]).expanduser()
    if (root / arg / "job_description.txt").exists():
        return root / arg
    apps = sorted({j.parent for j in list(root.glob("*/job_description.txt")) + list(root.glob("*/*/job_description.txt"))},
                  key=lambda d: (d / "job_description.txt").stat().st_mtime, reverse=True) if root.exists() else []
    if arg.lower() in ("latest", "last", "recent"):
        if apps:
            print(f"Using the most recent application: {apps[0].name}")
            return apps[0]
    else:
        words = [w for w in re.split(r"[\s_/-]+", arg.lower()) if w]
        hits = [d for d in apps if all(w in str(d.relative_to(root)).lower() for w in words)]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            print(f"Several applications match \"{arg}\"; using the most recent: {hits[0].name}")
            for h in hits[1:6]:
                print(f"   also: {h.name}")
            return hits[0]
    msg = f"Couldn't find an application folder called \"{arg}\" (looked in {root})."
    if apps:
        msg += "\nRecent applications:\n" + "\n".join(f"   {d.relative_to(root)}" for d in apps[:8])
        msg += "\nYou can pass part of the name, e.g.  revise atlassian   or   revise latest"
    sys.exit(msg)


def load_app(out: Path, cfg: dict) -> tuple[dict, str, str]:
    """(meta, tailored resume markdown, job description) for an application folder."""
    import json
    if not (out / "job_description.txt").exists():
        sys.exit(f"No job_description.txt in {out}. Point this at a folder created by `tailor.py job`.")
    meta_file = out / ".source" / "application.json"
    meta = json.loads(meta_file.read_text(encoding="utf-8")) if meta_file.exists() else {}
    src = out / ".source" / meta.get("resume_md", "-")
    if not src.exists():  # older folders: find the saved tailored resume
        label = re.sub(r"\s+", cfg.get("space", "_"), cfg["resume_label"]).lower()
        found = [f for f in sorted((out / ".source").glob("*.md")) + sorted(out.glob("*.md"))
                 if label in f.stem.lower() and "cover" not in f.stem.lower()]
        src = found[0] if found else src
    resume_md = src.read_text(encoding="utf-8") if src.exists() else read_base(Path(cfg["base_resume"]).expanduser())
    return meta, resume_md, (out / "job_description.txt").read_text(encoding="utf-8")


def _save_meta(out: Path, meta: dict):
    import json
    (out / ".source").mkdir(exist_ok=True)
    (out / ".source" / "application.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")


def _versions_dir(out: Path) -> Path:
    d = out / ".source" / "cover_letter_versions"
    d.mkdir(parents=True, exist_ok=True)
    return d


def letter_versions(out: Path) -> list[Path]:
    return sorted(_versions_dir(out).glob("v*.md"), key=lambda p: int(p.stem[1:]))


def letter_history(out: Path) -> list[dict]:
    import json
    f = _versions_dir(out) / "history.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else []


def save_letter(out: Path, text: str, cfg: dict, note: str = "", feedback: list | None = None) -> list[str]:
    """Save a new version of the cover letter and (re)build the PDF."""
    import json
    meta, resume_md, _ = load_app(out, cfg)
    company = meta.get("company", "")
    stem = meta.get("letter_stem") or render_name(
        cfg["file_name"], name_fields(cfg, resume_md, company, meta.get("role", "")), cfg,
        doc=cfg["cover_letter_label"])
    if not meta.get("letter_stem"):
        meta["letter_stem"] = stem
        (out / ".source" / "application.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")

    vdir = _versions_dir(out)
    n = len(letter_versions(out)) + 1
    (vdir / f"v{n}.md").write_text(text.strip() + "\n", encoding="utf-8")
    hist = letter_history(out) + [{"version": n, "change": note, "feedback": feedback or [],
                                   "time": dt.datetime.now().isoformat(timespec="seconds")}]
    (vdir / "history.json").write_text(json.dumps(hist, indent=2), encoding="utf-8")
    (out / ".source" / f"{stem}.md").write_text(text.strip() + "\n", encoding="utf-8")
    return build_letter_files(out, stem, text, meta, resume_md, cfg)


def build_letter_files(out: Path, stem: str, text: str, meta: dict, resume_md: str, cfg: dict) -> list[str]:
    import pdf_template
    formats = [f.lower() for f in meta.get("formats") or cfg["formats"]]
    lcfg = {**cfg, **meta.get("letter_style", {})}
    tpl = None
    tpl_path = meta.get("template")
    if tpl_path is None and cfg.get("match_layout", True):  # folder from an older version of the tool
        base = Path(cfg.get("template_pdf") or cfg["base_resume"]).expanduser()
        tpl_path = str(base) if base.suffix.lower() == ".pdf" else ""
    if tpl_path and Path(tpl_path).exists():
        tpl = pdf_template.Template(tpl_path)
    name, contact, _ = split_header(resume_md)
    if tpl:
        name, contact = tpl.header()
    date_text = dt.date.today().strftime("%B %d, %Y").replace(" 0", " ")
    written = []
    if "pdf" in formats or not formats:
        missing = pdf_template.write_letter_pdf(out / f"{stem}.pdf", text, name, contact,
                                                meta.get("company", ""), date_text, tpl, lcfg)
        written.append(f"{stem}.pdf")
        if missing:
            print(f"  ! Font not installed: {', '.join(missing)}. The letter used a close look-alike; "
                  f"install the font for an exact match.")
    if "docx" in formats:
        letter_to_docx(text, resume_md, out / f"{stem}.docx", lcfg, meta.get("company", ""))
        written.append(f"{stem}.docx")
    return written


REVISE_PROMPT = """You revise a job-application cover letter based on the candidate's feedback.
- Apply the new feedback precisely, and keep honoring the earlier feedback listed (don't undo it).
- Change only what the feedback calls for; keep everything else as it is.
- Facts must come from the resume, the job posting, or the candidate's feedback (the candidate may \
give you new facts about themselves; use them). Never invent anything else.
- Keep it 250-400 words unless the feedback asks for a different length. Plain paragraphs, no \
address block; keep the greeting and the sign-off with the candidate's name.
Reply with exactly:
<cover_letter>...the full revised letter...</cover_letter>
<changes>One or two short sentences on what you changed.</changes>"""


def revise_letter(current: str, feedback: str, out: Path, resume_md: str, job: str, cfg: dict) -> tuple[str, str]:
    hist = letter_history(out)
    earlier = hist[-1].get("feedback", []) if hist else []  # feedback that led to the current version
    user = (f"<resume>\n{resume_md}\n</resume>\n\n<job_posting>\n{job}\n</job_posting>\n\n"
            f"<current_letter>\n{current}\n</current_letter>\n\n"
            + ("<earlier_feedback>\n" + "\n".join(f"- {e}" for e in earlier) + "\n</earlier_feedback>\n\n"
               if earlier else "")
            + f"<new_feedback>\n{feedback}\n</new_feedback>")
    text = _ask(REVISE_PROMPT, user, cfg, "Revising cover letter")
    m = re.search(r"<cover_letter>(.*?)</cover_letter>", text, re.S)
    if not m or not m.group(1).strip():
        sys.exit("Claude didn't return a revised letter. Try rewording the feedback.")
    c = re.search(r"<changes>(.*?)</changes>", text, re.S)
    return m.group(1).strip(), (c.group(1).strip() if c else "")


def _open(path: Path):
    opener = {"darwin": "open", "win32": "explorer"}.get(sys.platform, "xdg-open")
    subprocess.run([opener, str(path)], check=False)


def cmd_letter(args, cfg):
    """Write a cover letter for an application folder made earlier (or start over with --new)."""
    out = find_app(args.folder, cfg)
    meta, resume_md, job = load_app(out, cfg)
    if args.model:
        cfg["model"] = args.model
    if letter_versions(out) and not args.new:
        sys.exit(f"This application already has a cover letter.\n"
                 f"  To improve it:      python tailor.py revise \"{out}\"\n"
                 f"  To start over:      python tailor.py letter \"{out}\" --new")
    res = call_claude(resume_md, job, "YES", {**cfg, **meta.get("letter_style", {})},
                      args.instructions or "", letter_only=True)
    if not res["cover_letter"]:
        sys.exit("Claude didn't return a cover letter. Try again, or add -i with more guidance.")
    note = "rewritten from scratch" if letter_versions(out) else "first draft"
    written = save_letter(out, res["cover_letter"], cfg, note=note)
    print(f"\n✓ Cover letter (version {len(letter_versions(out))}) saved to: {out}")
    for w in written:
        print(f"   - {w}")
    print(f"\n  To refine it:  python tailor.py revise \"{out}\"")


def cmd_revise(args, cfg):
    """Improve the cover letter with feedback; one-shot or interactive. Every version is kept."""
    out = find_app(args.folder, cfg)
    meta, resume_md, job = load_app(out, cfg)
    if args.model:
        cfg["model"] = args.model
    if args.feedback and " ".join(args.feedback).strip().lower() in ("revise", "letter", "interactive"):
        args.feedback = []  # e.g. `revise <folder> revise --open`
    versions = letter_versions(out)
    if not versions:  # folder made by an earlier version of the tool: adopt its existing letter
        label = re.sub(r"\s+", cfg.get("space", "_"), cfg["cover_letter_label"])
        old = [f for f in sorted((out / ".source").glob("*.md")) + sorted(out.glob("*.md"))
               if "cover" in f.stem.lower() or label.lower() in f.stem.lower()]
        if old:
            if not meta.get("letter_stem"):
                meta["letter_stem"] = old[0].stem
                _save_meta(out, meta)
            save_letter(out, old[0].read_text(encoding="utf-8"), cfg, note="first draft")
            versions = letter_versions(out)
            meta, resume_md, job = load_app(out, cfg)
    if not versions:
        sys.exit(f"No cover letter yet. Create one first:  python tailor.py letter \"{out.name}\"")
    hist = {h["version"]: h for h in letter_history(out)}

    def show_list():
        for v in versions:
            n = int(v.stem[1:])
            h = hist.get(n, {})
            print(f"  v{n}  {h.get('time', '')[:16].replace('T', ' ')}  {h.get('change', '')}")

    if args.list:
        show_list(); return
    if args.use:
        src = _versions_dir(out) / f"v{args.use}.md"
        if not src.exists():
            sys.exit(f"No version {args.use}. Versions:"); 
        old = next((h for h in letter_history(out) if h["version"] == args.use), {})
        written = save_letter(out, src.read_text(encoding="utf-8"), cfg, note=f"restored v{args.use}",
                              feedback=old.get("feedback", []))
        print(f"✓ Restored version {args.use} as v{len(letter_versions(out))}: {', '.join(written)}")
        return

    current = versions[-1].read_text(encoding="utf-8")
    pdf = out / f"{meta.get('letter_stem', '')}.pdf"

    def one_round(feedback):
        nonlocal current
        new, changes = revise_letter(current, feedback, out, resume_md, job, cfg)
        prior = (letter_history(out) or [{}])[-1].get("feedback", [])
        save_letter(out, new, cfg, note=feedback, feedback=prior + [feedback])
        current = new
        print(f"\n✓ Saved version {len(letter_versions(out))}" + (f": {changes}" if changes else ""))
        if args.open and pdf.exists():
            _open(pdf)

    if args.feedback:
        one_round(" ".join(args.feedback))
        print(f"  {pdf}")
        return

    # interactive
    print(f"\n{'─' * 70}\n{current.strip()}\n{'─' * 70}")
    print("Tell me what to change, e.g. \"warmer tone\", \"mention I led the 2024 migration\",\n"
          "\"shorter, 3 paragraphs\". Commands: show, undo, versions, done (or just press Enter).")
    while True:
        try:
            fb = input("\nfeedback> ").strip()
        except (EOFError, KeyboardInterrupt):
            print(); break
        if fb.lower() in ("", "done", "q", "quit", "exit"):
            break
        if fb.lower() == "show":
            print(f"\n{current.strip()}"); continue
        if fb.lower() in ("versions", "list"):
            versions = letter_versions(out); hist = {h["version"]: h for h in letter_history(out)}
            show_list(); continue
        if fb.lower() == "undo":
            vs = letter_versions(out)
            if len(vs) < 2:
                print("  Nothing to undo."); continue
            prev = vs[-2].read_text(encoding="utf-8")
            old = next((h for h in letter_history(out) if h["version"] == int(vs[-2].stem[1:])), {})
            save_letter(out, prev, cfg, note=f"restored v{vs[-2].stem[1:]}", feedback=old.get("feedback", []))
            current = prev
            print(f"  Back to the previous version (saved as v{len(letter_versions(out))}).")
            continue
        one_round(fb)
        print(f"\n{'─' * 70}\n{current.strip()}\n{'─' * 70}")
    print(f"\nFinal letter: {pdf}\nAll versions: {_versions_dir(out)}")


def cmd_fetch(args, cfg):
    """Preview what the tool reads from a listing URL (no API call)."""
    text = fetch_job(args.url)
    if args.save:
        Path(args.save).expanduser().write_text(text, encoding="utf-8")
        print(f"✓ Saved {len(text)} characters to {args.save}")
    else:
        print("\n" + text)


def cmd_import(args, cfg):
    src = Path(args.file).expanduser()
    dest = Path(args.dest or cfg["base_resume"]).expanduser()
    if src.suffix.lower() == ".docx":
        md = docx_to_markdown(src)
    elif src.suffix.lower() == ".pdf":
        md = pdf_to_text(src)
    else:
        md = src.read_text(encoding="utf-8")
    if dest.exists() and not args.force:
        sys.exit(f"{dest} already exists. Use --force to overwrite.")
    dest.write_text(md, encoding="utf-8")
    print(f"✓ Wrote {dest}\n  Open it and tidy the structure (see README: '# Name', contact line, "
          f"'## Section', '### Title, Company | Location | Dates', '- bullets').")


def cmd_render(args, cfg):
    """Re-render an edited .md (resume or letter) to docx/pdf without calling Claude."""
    src = Path(args.file).expanduser()
    if src.suffix.lower() == ".json":  # layout-matched resume: re-apply (edited) slot texts to the PDF
        import json
        import pdf_template
        data = json.loads(src.read_text(encoding="utf-8"))
        tpl = pdf_template.Template(data["template"])
        dest = (src.parent.parent if src.parent.name == ".source" else src.parent) / \
            (src.name.replace(".edits.json", "") + ".pdf")
        over = tpl.render(data["edits"], dest)
        print(f"✓ {dest}" + (f"\n  ! Set smaller to fit: {', '.join(over)}" if over else ""))
        return
    md = src.read_text(encoding="utf-8")
    dest_dir = src.parent.parent if src.parent.name == ".source" else src.parent
    docx_path = dest_dir / (src.stem + ".docx")
    if md.lstrip().startswith("# "):
        resume_md_to_docx(md, docx_path, cfg)
    else:
        base = Path(cfg["base_resume"]).expanduser()
        letter_to_docx(md, read_base(base) if base.exists() else "", docx_path, cfg, "")
    pdf = docx_to_pdf(docx_path)
    if pdf and "docx" not in cfg["formats"]:
        docx_path.unlink(missing_ok=True)
        print(f"✓ {pdf}")
    else:
        print(f"✓ {docx_path}" + (f"\n✓ {pdf}" if pdf else "\n  ! Couldn't make a PDF (install LibreOffice)."))


def main():
    cfg = load_config()
    ap = argparse.ArgumentParser(description="Tailor your resume and cover letter to a job posting.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    j = sub.add_parser("job", help="tailor to a job posting")
    j.add_argument("source", help="job description: .txt/.md/.pdf/.docx file, a URL, or - to paste")
    j.add_argument("-o", "--out", help="output folder (default: output_root/DATE_Company_Role)")
    cl = j.add_mutually_exclusive_group()
    cl.add_argument("-c", "--cover-letter", action="store_true", help="always write a cover letter")
    cl.add_argument("-n", "--no-cover-letter", action="store_true", help="never write a cover letter")
    j.add_argument("-i", "--instructions", help='extra guidance, e.g. "emphasize leadership"')
    j.add_argument("--base", help="base resume .md (default from config)")
    j.add_argument("--template", help="PDF whose layout to copy (default: your base resume, if it's a PDF)")
    j.add_argument("--simple", action="store_true", help="use the standard layout instead of your PDF's")
    j.add_argument("--model", help="Claude model id (default from config)")
    j.add_argument("--formats", help="comma list of docx,pdf,md (default from config)")
    j.add_argument("--open", action="store_true", help="open the output folder when done")
    j.add_argument("--name", help='file name pattern, e.g. "{name} - {doc} - {company}"')
    j.add_argument("--folder-name", help='folder name pattern, e.g. "{company}/{role}"')

    im = sub.add_parser("import", help="convert an existing .docx/.pdf resume into base_resume.md")
    im.add_argument("file")
    im.add_argument("--dest")
    im.add_argument("--force", action="store_true")

    r = sub.add_parser("render", help="re-render an edited .md to .docx/.pdf (no API call)")
    r.add_argument("file")

    lt = sub.add_parser("letter", help="write a cover letter for an existing application folder")
    lt.add_argument("folder")
    lt.add_argument("-i", "--instructions", help="extra guidance for the letter")
    lt.add_argument("--new", action="store_true", help="start over even if a letter exists")
    lt.add_argument("--model")

    rv = sub.add_parser("revise", help="improve the cover letter with your feedback (keeps every version)")
    rv.add_argument("folder", help="the application folder")
    rv.add_argument("feedback", nargs="*", help='what to change; omit for an interactive session')
    rv.add_argument("--list", action="store_true", help="list saved versions")
    rv.add_argument("--use", type=int, help="restore version N")
    rv.add_argument("--open", action="store_true", help="open the PDF after each revision")
    rv.add_argument("--model")

    f = sub.add_parser("fetch", help="preview the text read from a listing URL (no API call)")
    f.add_argument("url")
    f.add_argument("--save", help="write it to a file instead of printing")

    args = ap.parse_args()
    if getattr(args, "model", None):
        cfg["model"] = args.model
    {"job": cmd_job, "import": cmd_import, "render": cmd_render,
     "letter": cmd_letter, "revise": cmd_revise, "fetch": cmd_fetch}[args.cmd](args, cfg)


if __name__ == "__main__":
    main()

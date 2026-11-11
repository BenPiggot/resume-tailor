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
    "formats": ["docx", "pdf"],
    "cover_letter": "auto",
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


def call_claude(base_resume: str, job: str, cover_letter: str, cfg: dict, extra: str,
                letter_only: bool = False) -> dict:
    """cover_letter is "YES", "NO" or "AUTO"."""
    try:
        import anthropic
    except ImportError:
        sys.exit("Missing dependency: pip install -r requirements.txt")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        sys.exit("Set your API key first:  export ANTHROPIC_API_KEY=sk-ant-...  "
                 "(get one at https://console.anthropic.com)")
    instructions = "\n".join(x for x in [cfg.get("extra_instructions", ""), extra] if x)
    user = (
        f"<base_resume>\n{base_resume}\n</base_resume>\n\n"
        f"<job_posting>\n{job}\n</job_posting>\n\n"
        f"Write a cover letter: {cover_letter}\n"
        + ("\nOnly the cover letter is needed this time: leave <resume> empty. The resume above is "
           "already tailored to this job.\n" if letter_only else "")
        + (f"\nAdditional instructions from the candidate:\n{instructions}\n" if instructions else "")
    )
    client = anthropic.Anthropic()
    print(f"{'Writing cover letter' if letter_only else 'Tailoring'} with {cfg['model']}…", flush=True)
    with client.messages.stream(
        model=cfg["model"],
        max_tokens=16000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user}],
    ) as stream:
        msg = stream.get_final_message()
    text = "".join(b.text for b in msg.content if b.type == "text")
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


def write_doc(out: Path, stem: str, md: str, build, formats: list[str]) -> list[str]:
    """Build stem.docx (and .pdf), keep or tuck away the .md source. Returns file names written."""
    written = []
    docx_path = out / f"{stem}.docx"
    build(docx_path)
    if "docx" in formats:
        written.append(docx_path.name)
    if "pdf" in formats:
        pdf = docx_to_pdf(docx_path)
        if pdf:
            written.append(pdf.name)
        else:
            print("  ! PDF skipped: install LibreOffice (or `pip install docx2pdf` with MS Word) to enable it.")
    if "docx" not in formats:
        docx_path.unlink(missing_ok=True)
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


def cmd_job(args, cfg):
    base_path = Path(args.base or cfg["base_resume"]).expanduser()
    if not base_path.exists():
        sys.exit(f"Base resume not found at {base_path}.\n"
                 f"Create it, or run:  python tailor.py import your_resume.docx")
    base = read_base(base_path)
    job = read_job(args.source)
    if len(job.strip()) < 50:
        sys.exit("The job description looks empty.")

    mode = cover_mode(args, cfg)
    res = call_claude(base, job, mode, cfg, args.instructions or "")
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

    written += write_doc(out, resume_stem, res["resume"],
                         lambda p: resume_md_to_docx(res["resume"], p, cfg), formats)
    import json
    (out / ".source").mkdir(exist_ok=True)
    (out / ".source" / "application.json").write_text(json.dumps(
        {"company": company, "role": role, "resume_md": f"{resume_stem}.md", "letter_stem": letter_stem,
         "created": dt.datetime.now().isoformat(timespec="seconds")}, indent=2), encoding="utf-8")
    if "md" in formats:  # keep a copy in .source too so `letter` always finds the tailored resume
        (out / ".source" / f"{resume_stem}.md").write_text(res["resume"] + "\n", encoding="utf-8")
    letter_written = mode != "NO" and bool(res["cover_letter"])
    if letter_written:
        written += write_doc(out, letter_stem, res["cover_letter"],
                             lambda p: letter_to_docx(res["cover_letter"], res["resume"], p, cfg, company),
                             formats)

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


def cmd_letter(args, cfg):
    """Add a cover letter to an application folder made earlier, reusing its tailored resume."""
    out = Path(args.folder).expanduser()
    job_file = out / "job_description.txt"
    if not job_file.exists():
        sys.exit(f"No job_description.txt in {out}. Point this at a folder created by `tailor.py job`.")
    import json
    meta_file = out / ".source" / "application.json"
    meta = json.loads(meta_file.read_text()) if meta_file.exists() else {}
    src = out / ".source" / meta.get("resume_md", "-")
    resume_md = src.read_text(encoding="utf-8") if src.exists() else read_base(Path(cfg["base_resume"]).expanduser())
    res = call_claude(resume_md, job_file.read_text(encoding="utf-8"), "YES", cfg,
                      args.instructions or "", letter_only=True)
    if not res["cover_letter"]:
        sys.exit("Claude didn't return a cover letter. Try again, or add -i with more guidance.")
    company = meta.get("company") or res["company"] or "Unknown"
    stem = args.name or meta.get("letter_stem") or render_name(
        cfg["file_name"], name_fields(cfg, resume_md, company, meta.get("role", "")), cfg,
        doc=cfg["cover_letter_label"])
    formats = [f.lower() for f in cfg["formats"]]
    written = write_doc(out, stem, res["cover_letter"],
                        lambda p: letter_to_docx(res["cover_letter"], resume_md, p, cfg, company), formats)
    print(f"\n✓ Cover letter saved to: {out}")
    for w in written:
        print(f"   - {w}")


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
    md = src.read_text(encoding="utf-8")
    dest_dir = src.parent.parent if src.parent.name == ".source" else src.parent
    docx_path = dest_dir / (src.stem + ".docx")
    if md.lstrip().startswith("# "):
        resume_md_to_docx(md, docx_path, cfg)
    else:
        base = Path(cfg["base_resume"]).expanduser()
        letter_to_docx(md, read_base(base) if base.exists() else "", docx_path, cfg, "")
    pdf = docx_to_pdf(docx_path)
    print(f"✓ {docx_path}" + (f"\n✓ {pdf}" if pdf else ""))


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

    lt = sub.add_parser("letter", help="add a cover letter to an existing application folder")
    lt.add_argument("folder")
    lt.add_argument("-i", "--instructions", help="extra guidance for the letter")
    lt.add_argument("--model")
    lt.add_argument("--name", help="exact file name (without extension) for the letter")

    f = sub.add_parser("fetch", help="preview the text read from a listing URL (no API call)")
    f.add_argument("url")
    f.add_argument("--save", help="write it to a file instead of printing")

    args = ap.parse_args()
    if getattr(args, "model", None):
        cfg["model"] = args.model
    {"job": cmd_job, "import": cmd_import, "render": cmd_render,
     "letter": cmd_letter, "fetch": cmd_fetch}[args.cmd](args, cfg)


if __name__ == "__main__":
    main()

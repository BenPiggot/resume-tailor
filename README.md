# resume-tailor

This tool tailors your master resume to a specific job posting, and can write a cover letter too. It uses Claude and saves everything into one folder for each application:

```
~/Job Applications/2026-11-11_Northwind_Logistics_Senior_Backend_Engineer/
├── Jordan_Rivera_Northwind_Logistics_Resume.docx
├── Jordan_Rivera_Northwind_Logistics_Resume.pdf
├── Jordan_Rivera_Northwind_Logistics_Cover_Letter.docx   (when a cover letter is written)
├── Jordan_Rivera_Northwind_Logistics_Cover_Letter.pdf
├── tailoring_notes.md      ← what was emphasized, keywords matched, honest gaps
├── job_description.txt     ← a copy of the posting, in case it's taken down later
└── .source/               ← editable .md sources (see "Tweak and re-render")
```

You can change all of these names. See [Naming](#naming).

The tool **never invents experience**. It reorders, rephrases and trims what is already in your base resume, and it lists any requirements you don't meet under "gaps" in the notes.

## Setup (one time)

1. Install Python 3.11 or newer, then run:
   ```bash
   cd resume-tailor
   pip install -r requirements.txt
   ```
2. Get an API key at <https://console.anthropic.com>, then set it:
   ```bash
   export ANTHROPIC_API_KEY=sk-ant-...        # macOS/Linux: add this to ~/.zshrc or ~/.bashrc
   setx ANTHROPIC_API_KEY "sk-ant-..."        # Windows: run once, then open a new terminal
   ```
3. **For PDF output**, install [LibreOffice](https://www.libreoffice.org/download/), which is free. Without LibreOffice, run `pip install docx2pdf` to use Microsoft Word for the conversion instead. If neither is installed, you still get the .docx files.
4. **Add your resume.** Either edit `base_resume.md` directly (it holds a sample), or import your Word resume:
   ```bash
   python tailor.py import ~/Documents/My_Resume.docx --force
   ```
   Then open `base_resume.md` and check that its structure matches the format below.

   **Or just use your PDF or Word file as is.** Set `base_resume = "~/Documents/My_Resume.pdf"` in `config.toml`, or pass `--base My_Resume.pdf` on a run. The tool extracts the text on every run, and Claude puts it back into the resume layout. This works well for ordinary text-based PDFs. It can go wrong with two-column or table-heavy designs, and it doesn't work with scanned images. If dates or sections come out jumbled, run the `import` command once and use the Markdown file instead.

### Base resume format

```markdown
# Your Name
email | phone | city | linkedin.com/in/you

## Summary
One or two sentences.

## Experience
### Job Title, Company | City, ST | 2022 – Present
- Accomplishment with a number
```
The last ` | ` part of a `###` line is right-aligned, so put dates there. Use `**bold**` for labels like `**Languages:**`.

**Tip:** your base resume can be *longer* than one page. Include every project and bullet you might ever want. The tailoring step picks the most relevant ones for each job.

## Use

```bash
# straight from a job listing link
python tailor.py job https://boards.greenhouse.io/acme/jobs/123

# from a saved text/PDF/Word file of the posting, into a folder you choose
python tailor.py job posting.txt --out ~/Jobs/Acme

# paste the description into the terminal (finish with Ctrl-D)
python tailor.py job -

# steer it
python tailor.py job posting.txt -i "Emphasize people leadership; I was the de-facto tech lead at Cascade"
```

| Flag | Meaning |
|---|---|
| `-c` / `-n` | always / never write a cover letter (default: decide automatically, see below) |
| `-o`, `--out DIR` | save into exactly this folder |
| `--name "..."` / `--folder-name "..."` | naming patterns for this run (see [Naming](#naming)) |
| `-i "..."` | extra instructions for this application |
| `--model` | e.g. `claude-opus-5-5` for the most capable model |
| `--formats docx,pdf,md` | which files to keep |
| `--open` | open the folder when done |

### Reading listings from a link

When you give the tool a link, it gets the listing in the most reliable way available:

1. **Greenhouse, Lever and Ashby** links (`boards.greenhouse.io`, `job-boards.greenhouse.io`, `jobs.lever.co`, `jobs.ashbyhq.com`): the tool reads the posting from that platform's public API, so it gets clean, complete text.
2. **Most other career sites** (Workday, iCIMS, SmartRecruiters, many company sites): the tool reads the structured job data these pages include for Google Jobs.
3. **Anything else**: the tool reads the visible text of the page.
4. **Pages that only load with JavaScript**: install the optional headless browser and the tool will use it automatically:
   ```bash
   pip install playwright && playwright install chromium
   ```

Pages that need you to be logged in (for example LinkedIn) can't be read from a link. For those, copy the listing into a `.txt` file, or run `job -` and paste it in.

To check what the tool reads from a link without calling Claude, run:
```bash
python tailor.py fetch https://jobs.lever.co/acme/1234            # print it
python tailor.py fetch https://jobs.lever.co/acme/1234 --save job.txt
```

### Cover letters: written only when needed

By default (`cover_letter = "auto"` in `config.toml`), Claude writes a cover letter only when the listing asks for or accepts one. That means it mentions a cover letter, a letter of interest or a motivation letter, or asks why you're interested. The terminal shows the decision and the reason, and they are also saved in `tailoring_notes.md`.

- To always get one, use `-c`; to never get one, use `-n`. Or set `cover_letter = "yes"` or `"no"` in `config.toml`.
- **To add one later** to an application that skipped it:
  ```bash
  python tailor.py letter "~/Job Applications/2026-11-11_Acme_Corp_Backend_Engineer"
  ```
  This uses that folder's tailored resume and saved job description, so nothing gets fetched again.

## Naming

Folder and file names are patterns in `config.toml`:

```toml
folder_name = "{date}_{company}_{role}"         # 2026-11-11_Acme_Corp_Backend_Engineer
file_name   = "{name}_{company}_{doc}"          # Jordan_Rivera_Acme_Corp_Resume.pdf
```

The placeholders are `{name}` (your full name, from the top of your resume), `{first}`, `{last}`, `{company}`, `{role}`, `{date}`, and `{doc}`. `{doc}` becomes "Resume" or "Cover Letter", and the file name pattern must include it. Some examples:

| Setting | Result |
|---|---|
| `file_name = "{name} - {doc} - {company}"` | `Jordan_Rivera - Resume - Acme_Corp.pdf` |
| `file_name = "{last}_{doc}"` | `Rivera_Resume.pdf` |
| `folder_name = "{company}/{date}_{role}"` | the folder `Acme_Corp/2026-11-11_Backend_Engineer/`, one subfolder per company |
| `space = " "` | keeps spaces in names: `Acme Corp`, `Cover Letter` |
| `date_format = "%b %d %Y"` | `Nov 11 2026` |
| `resume_label = "CV"` | `Jordan_Rivera_Acme_Corp_CV.pdf` |

Characters that aren't allowed in file names, such as `/ : * ? " < > |`, are removed automatically. If you apply to the same role twice on the same day, the second folder gets `_2` added to its name, so nothing is overwritten. To name a single run differently, use `--name` and `--folder-name`, or use `--out` to choose the exact folder.

## Tweak and re-render

To change something by hand, edit the `.md` file in the application's `.source/` folder, then run:
```bash
python tailor.py render "~/Job Applications/.../.source/Jordan_Rivera_Acme_Corp_Resume.md"
```
This rebuilds the .docx and .pdf without calling the API.

## Settings

Edit `config.toml` to change the default output folder, cover letter behavior, naming, model, font, file formats, and the standing instructions applied to every application.

## Cost

Each run makes one API call. With the default `claude-sonnet-5-5` model, a resume plus cover letter typically costs a few cents.

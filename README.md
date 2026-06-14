# enigma-cv-tools

Keep your academic CV publication list automatically up to date — powered by PubMed, Google Sheets, and a little AI.

Built by the [ENIGMA Brain Injury Working Group](https://enigma.ini.usc.edu/ongoing-studies/enigma-tbi/) at the University of Utah. Designed for researchers with large, growing publication lists who are tired of updating their CV manually.

---

## What this does

- **Bootstrap** a fully-populated Publications Master Google Sheet from PubMed + an existing CV (one-time setup per person)
- **Track status** of papers through submission → accepted → in press → published
- **Auto-update** via PubMed: finds newly published papers and fills in missing DOIs/PMIDs
- **Email intake**: forward acceptance/submission emails and have them parsed and logged automatically using Gemini AI
- **Export** a formatted CV section (publications + talks, grouped by type) to a Google Doc

---

## Quick start

### 1. Install Python dependencies

```bash
pip install pyyaml python-docx pdfminer.six
```

Python 3.8+ required. No NCBI API key needed (the script respects NCBI's rate limits automatically).

### 2. Create a config file

Copy `config_template.yaml` and fill in your details:

```yaml
name: Your Name
orcid: 0000-0002-XXXX-XXXX        # find yours at orcid.org
pubmed_query: LastnameAB[Author]   # how your name appears in PubMed
affiliation_hints:
  - Your University
  - Previous Institution
cv_path: ./my_cv.docx              # optional — skip if you don't have one
```

### 3. Run the bootstrap script

```bash
python bootstrap_colleague.py --config my_config.yaml
```

This generates two files:
- `LastnameAB_publications.csv` — your full publication list, ready to import
- `LastnameAB_reconciliation.txt` — a log of what matched, what didn't, and any CV-only entries to review

### 4. Set up your Google Sheet

1. Create a new Google Sheet
2. Rename the first tab to **Publications Master**
3. File → Import → upload the CSV → **Replace current sheet**
4. Add a second tab called **Talks** with columns: `id, year, month, type, title, venue, location, virtual, cancelled, notes, raw_text`

### 5. Install the Apps Script

1. In your Sheet: **Extensions → Apps Script**
2. Delete any existing code and paste the entire contents of `cv_automation.gs`
3. Update the `AUTHOR_QUERY` constant at the top to match your PubMed query
4. Save (Cmd/Ctrl+S) and reload the Sheet
5. A **CV Tools** menu will appear — click it and authorize when prompted

### 6. Add your Gemini API key (for email intake)

The email parsing feature uses Google's Gemini AI. To enable it:

1. Get a free key at [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. In Apps Script: **Project Settings (gear icon) → Script Properties → Add property**
   - Name: `GEMINI_API_KEY`
   - Value: your key
3. Test it: **CV Tools → Test Gemini connection**

Email intake is optional — everything else works without it.

---

## Setting up email intake

Once your API key is in place:

1. In Gmail, create a label called `CV/process`
2. Create a Gmail filter: emails **To** `youraddress+cv@gmail.com` → apply label `CV/process`
3. Forward any submission or acceptance emails to `youraddress+cv@gmail.com`
4. In your Sheet: **CV Tools → Process new emails**

The script will auto-apply high-confidence matches and put uncertain ones in an **Email Review Queue** tab for you to check.

---

## Day-to-day workflow

Open your Sheet, then open the **CV Dashboard** (CV Tools → Open CV Dashboard). Four buttons:

| Button | What it does |
|--------|-------------|
| 🔍 Check PubMed for new publications | Finds accepted/in-press papers that are now published |
| 🔗 Fill in missing DOIs & PMIDs | Backfills identifiers for any paper missing them |
| ✅ Apply approved matches | Writes approved PubMed Log entries into your sheet |
| 📬 Process new emails | Reads forwarded submission/acceptance emails |

After running **Check PubMed** or **Fill in missing DOIs**, open the **PubMed Log** tab, type `y` in column A for correct matches, then click **Apply approved matches**. The script never overwrites data you already have.

To export your CV: **CV Tools → Export CV to Google Doc**. Creates a timestamped Google Doc in your Drive with publications grouped by type and talks.

---

## Onboarding a colleague

The same bootstrap script works for anyone. Create a config file for them (copy `config_template.yaml`, fill in their details), run the script, and import the CSV into their own Google Sheet. Then paste `cv_automation.gs` into their Apps Script editor and update `AUTHOR_QUERY` at the top.

See [docs/user_guide.md](docs/user_guide.md) for a step-by-step guide written for users with no coding experience.

---

## Schema reference

### Publications Master columns

| Column | Values | Notes |
|--------|--------|-------|
| `id` | PUB0001, PUB0002… | Sequential, assigned by bootstrap script |
| `category` | `journal` / `review` / `book_chapter` / `conference` / `editorial` | |
| `status` | `published` / `accepted` / `in_press` / `under_review` / `submitted` | |
| `year` | YYYY | |
| `authors` | Full author list | As returned by PubMed |
| `title` | | |
| `venue` | Journal / book / conference name | |
| `volume` / `issue` / `pages` | | |
| `doi` | | |
| `pmid` | | PubMed ID |
| `submitted_date` / `accepted_date` / `published_date` | YYYY-MM-DD | |
| `notes` | | Free text |
| `raw_text` | | Original import text if parsed from CV |

### Talks columns

| Column | Values |
|--------|--------|
| `id` | TALK0001, TALK0002… |
| `year` / `month` | YYYY / MM |
| `type` | `invited` / `keynote` / `lecture` / `grand_rounds` / `workshop` / `symposium` / `conference_paper` / `panel` |
| `title` | Talk title (optional) |
| `venue` | Conference or institution |
| `location` | City, Region/Country |
| `virtual` | `yes` / `hybrid` / (blank) |
| `cancelled` | `yes` / (blank) |
| `notes` | |

---

## Troubleshooting

**PubMed returns no results**
Check your `pubmed_query` by pasting it directly into [pubmed.ncbi.nlm.nih.gov](https://pubmed.ncbi.nlm.nih.gov). Common issues: unusual name formatting, name collision without affiliation hints.

**CV parsing finds very few titles**
The heuristic extractor works best on standard CV formats. For highly unusual layouts, you can skip CV parsing (`cv_path: ""`) and rely on PubMed alone — with an ORCID this is usually comprehensive.

**CV Tools menu doesn't appear**
Reload the Sheet. If it still doesn't appear, check that the script was saved in Apps Script (no asterisk in the tab name).

**Gemini API errors**
Verify your `GEMINI_API_KEY` in Script Properties. Run **CV Tools → Test Gemini connection** to confirm it works before processing real emails.

**"Publications Master tab not found"**
The sheet tab must be named exactly `Publications Master` (capital P, capital M).

---

## Requirements

- Python 3.8+
- `pyyaml`, `python-docx`, `pdfminer.six` (install with pip)
- A Google account
- A free Gemini API key (for email intake only)

No NCBI API key required. No paid services required.

---

## Citation / credit

If you use this in your research group, a mention in your lab tools or methods would be appreciated:

> Dennis EL. enigma-cv-tools: PubMed-powered CV automation for academic researchers. GitHub. https://github.com/emilydennis/enigma-cv-tools

---

## License

MIT License — free to use, modify, and distribute with attribution.

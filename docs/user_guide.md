# CV Automation System — User Guide

This guide is for researchers who have had the system set up for them and want to understand how to use it day-to-day. No coding experience required.

---

## What this system does

It keeps your Publications Master Google Sheet automatically up to date. Once set up, your regular workflow is just a few button clicks:

- Check PubMed for papers that have moved from "accepted" to fully published
- Fill in missing DOIs and PubMed IDs
- Log forwarded submission and acceptance emails automatically
- Export a formatted publication list and talks section to a Google Doc

You never need to touch the code. Everything runs through a dashboard panel in your sheet.

---

## First-time setup: Authorize the script

The first time you use the dashboard, Google will ask you to confirm that you trust the script. This is normal — it's just Google checking that you want your own script to run.

1. Open your Publications Master Google Sheet
2. Reload the page (Cmd+R on Mac, F5 on Windows)
3. A **CV Tools** menu will appear in the top menu bar
4. Click **CV Tools → Open CV Dashboard**
5. Click any button. Google will ask you to authorize.
6. Click **Review permissions** and choose your account
7. You may see a warning: "Google hasn't verified this app." Click **Advanced**, then **Go to [project name] (unsafe)**. This is safe — it is your own script.
8. Click **Allow**. You only need to do this once.

---

## First-time setup: Email intake (optional)

If you want to forward submission/acceptance emails and have them automatically logged:

### Create a Gmail label

1. Open Gmail
2. On the left sidebar, scroll down and click **More**, then **Create new label**
3. Type `CV/process` and click Create
   - The slash creates a nested label (it will appear as CV > process). That's correct.

### Create a Gmail filter

You'll forward manuscript emails to a special alias address — your Gmail address with `+cv` added before the `@`.

For example: if your Gmail is `ewilde@gmail.com`, your intake address is `ewilde+cv@gmail.com`. Gmail delivers `+alias` mail to your regular inbox automatically.

1. In Gmail, click the Settings gear (top right) → **See all settings**
2. Click the **Filters and Blocked Addresses** tab
3. Click **Create a new filter**
4. In the **To** field, enter your `+cv` alias address
5. Click **Create filter**
6. Check **Apply the label** and choose `CV/process`
7. Click **Create filter** to save

### Test it

Forward any manuscript-related email to your `+cv` alias. Check that it arrives with the `CV/process` label. Then open your Sheet, open the Dashboard, and click **Process new emails**.

---

## Day-to-day use

### Opening the dashboard

Click **CV Tools** in the top menu bar, then **Open CV Dashboard**. A panel opens on the right side of your screen.

---

### 🔍 Check PubMed for new publications

Use this when you have papers at "accepted" or "in press" status and want to see if they've been officially published.

1. Click the button and wait 30–60 seconds
2. A message tells you how many matches were found
3. Open the **PubMed Log** tab in your Sheet
4. For each row, compare column E (your title) with column F (PubMed's title)
5. Type `y` in column A for rows you approve
6. Come back to the Dashboard and click **Apply approved matches**

> The script never overwrites data you already have — it only fills in empty fields.

---

### 🔗 Fill in missing DOIs & PMIDs

Use this periodically to find DOIs and PubMed IDs for papers that are missing them. The workflow is the same: review the PubMed Log, type `y` to approve, then apply.

---

### ✅ Apply approved matches

Run this after reviewing the PubMed Log and typing `y` next to the rows you approve. It writes the confirmed information into your Publications Master sheet and marks papers as published where appropriate.

---

### 📬 Process new emails

Run this after forwarding submission or acceptance emails to your `+cv` alias. The script reads them, extracts the publication details using AI, and either:

- **Auto-applies** them if it's confident (high-confidence acceptances and new submissions)
- **Queues** them in the **Email Review Queue** tab if it's uncertain

Check the Email Review Queue tab after running this to see if anything needs your attention.

---

### Exporting your CV

Click **CV Tools → Export CV to Google Doc**. The script builds a formatted document from your entire publication list and talks, and saves it to your Google Drive. A popup shows the link.

- Sections are ordered: Journal Articles, Reviews, Book Chapters, Conference Papers, Editorials, Talks
- Within each section, papers are sorted newest first
- Non-published papers get a status badge (e.g. `[in press]`, `[accepted]`)
- Each export creates a new timestamped file — it does not overwrite previous exports

---

## Troubleshooting

**CV Tools menu doesn't appear**
Reload the page. If it still doesn't appear, contact whoever set up your system.

**"Publications Master tab not found"**
Your main sheet tab must be named exactly `Publications Master` (capital P, capital M).

**"No Gmail label CV/process found"**
Go back to the email setup section above and create the label in Gmail.

**PubMed Log shows no matches**
Either your accepted/in-press papers aren't indexed in PubMed yet, or the titles in your sheet differ significantly from how they appear in PubMed. You can look up individual papers manually — contact whoever set up your system.

**Email processed but nothing appeared in the sheet**
Check the Email Review Queue tab — the email likely landed there for manual review. Low-confidence extractions and revision requests/rejections always go to the queue rather than being applied automatically.

---

## Questions

Contact the person who set up your system, or reach out to the ENIGMA Brain Injury Working Group.

/**
 * enigma-cv-tools — CV Automation Script
 * ========================================
 * Paste the entire contents of this file into the Apps Script editor
 * of your Publications Master Google Sheet (Extensions > Apps Script),
 * then update the CONFIG section below with your own details.
 *
 * Full setup instructions: see README.md or docs/user_guide.md
 *
 * Required Google Sheet tabs:
 *   - "Publications Master"   (import your bootstrap CSV here)
 *   - "Talks"                 (optional, for talks tracking)
 *
 * These tabs are created automatically when needed:
 *   - "PubMed Log"
 *   - "Email Review Queue"
 *
 * One-time setup:
 *   1. Update CONFIG below
 *   2. Save the script (Cmd/Ctrl+S)
 *   3. Reload your Google Sheet — a "CV Tools" menu will appear
 *   4. For email intake: add your Gemini API key in
 *      Apps Script > Project Settings > Script Properties
 *      Property name: GEMINI_API_KEY
 */

// =====================================================================
// CONFIG — update these for your own use
// =====================================================================

const AUTHOR_QUERY     = 'Dennis EL[Author]';   // your PubMed author query
const MASTER_SHEET     = 'Publications Master';
const TALKS_SHEET      = 'Talks';
const LOG_SHEET        = 'PubMed Log';
const FORM_RESPONSES_SHEET = 'Form Responses (Talks)';

const MATCH_THRESHOLD     = 0.70;   // title similarity required to log a candidate
const MAX_PUBMED_RESULTS  = 200;    // increase if you have >200 publications
const NCBI_BASE           = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils';

const LOG_HEADERS = [
  'Approve? (y/n)', 'CV id', 'Master row', 'Score',
  'CV title', 'PubMed title',
  'CV year', 'PubMed year',
  'CV venue', 'PubMed journal',
  'Volume', 'Issue', 'Pages', 'DOI', 'PMID',
];

const STOPWORDS = new Set([
  'a','an','and','the','of','in','on','for','to','with','by','from',
  'is','are','was','were','be','been','at','as','that','this','it','or'
]);

// Email intake config
const INTAKE_LABEL        = 'CV/process';
const PROCESSED_LABEL     = 'CV/processed';
const FAILED_LABEL        = 'CV/failed';
const REVIEW_SHEET        = 'Email Review Queue';
const GEMINI_MODEL        = 'gemini-2.5-flash';
const AUTO_APPLY_CONFIDENCE = 0.85;
const MAX_RETRY_COUNT     = 5;

// =====================================================================
// MENU
// =====================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CV Tools')
    .addItem('Open CV Dashboard', 'showSidebar')
    .addItem('Export CV to Google Doc', 'exportCvToDoc')
    .addSeparator()
    .addItem('Find candidates (accepted → published)', 'findCandidates')
    .addItem('Backfill missing DOIs/PMIDs', 'findDoiCandidates')
    .addItem('Apply approved matches', 'applyApprovedMatches')
    .addSeparator()
    .addItem('Lookup selected row by PMID or DOI…', 'lookupSelectedRow')
    .addSeparator()
    .addItem('Process emails manually', 'processIntakeEmails')
    .addItem('Test Gemini connection', 'testGeminiConnection')
    .addSeparator()
    .addItem('Clear PubMed log', 'clearLog')
    .addItem('Clear email review queue', 'clearReviewQueue')
    .addToUi();
}

// =====================================================================
// SIDEBAR
// =====================================================================

function showSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('CV Dashboard')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function sidebar_checkPubMed()   { findCandidates(); }
function sidebar_backfillDois()  { findDoiCandidates(); }
function sidebar_applyMatches()  { applyApprovedMatches(); }
function sidebar_processEmails() { processIntakeEmails(); }

function sidebar_getStats() {
  try {
    const master = SpreadsheetApp.getActive().getSheetByName(MASTER_SHEET);
    if (!master) return { error: `"${MASTER_SHEET}" tab not found.` };
    const data    = master.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf('status');
    if (statusCol < 0) return { error: 'No "status" column found.' };
    const counts = { published: 0, accepted: 0, under_review: 0, in_press: 0, other: 0 };
    for (let i = 1; i < data.length; i++) {
      const s = String(data[i][statusCol]).toLowerCase().trim();
      if (counts[s] !== undefined) counts[s]++;
      else counts.other++;
    }
    return {
      total:        data.length - 1,
      published:    counts.published,
      accepted:     counts.accepted,
      under_review: counts.under_review,
      in_press:     counts.in_press,
      other:        counts.other,
    };
  } catch (e) {
    return { error: e.message };
  }
}

function getSidebarHtml() {
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    font-size: 13px; color: #1a1a1a; background: #f8f8f6;
    padding: 16px; line-height: 1.5;
  }
  h1 { font-size: 15px; font-weight: 600; margin-bottom: 4px; }
  .subtitle { font-size: 12px; color: #666; margin-bottom: 16px; }
  .stats-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 20px; }
  .stat-card { background: #fff; border: 0.5px solid #e0e0d8; border-radius: 8px; padding: 10px 12px; }
  .stat-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 2px; }
  .stat-value { font-size: 20px; font-weight: 600; }
  .stat-card.highlight .stat-value { color: #1a6b3c; }
  .section-label { font-size: 11px; font-weight: 600; color: #888; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 8px; }
  .action-list { display: flex; flex-direction: column; gap: 8px; margin-bottom: 20px; }
  .action-btn {
    background: #fff; border: 0.5px solid #d0d0c8; border-radius: 8px;
    padding: 10px 12px; cursor: pointer; text-align: left; font-family: inherit;
    font-size: 13px; color: #1a1a1a; transition: background 0.1s;
    display: flex; align-items: flex-start; gap: 10px;
  }
  .action-btn:hover { background: #f0f0ea; }
  .action-btn:disabled { opacity: 0.5; cursor: not-allowed; }
  .btn-icon { font-size: 18px; flex-shrink: 0; margin-top: 1px; }
  .btn-text strong { display: block; font-weight: 600; margin-bottom: 1px; }
  .btn-text span { font-size: 11px; color: #666; }
  .status-bar {
    background: #fff; border: 0.5px solid #e0e0d8; border-radius: 8px;
    padding: 10px 12px; font-size: 12px; color: #444;
    min-height: 40px; display: flex; align-items: center; gap: 8px;
  }
  .status-bar.running { color: #7a5a00; background: #fffbeb; border-color: #e8d88a; }
  .status-bar.success { color: #1a6b3c; background: #f0faf4; border-color: #a3d9b8; }
  .status-bar.error   { color: #a33030; background: #fff0f0; border-color: #f5b8b8; }
  .spinner {
    width: 14px; height: 14px; border: 2px solid #d4a800;
    border-top-color: transparent; border-radius: 50%;
    animation: spin 0.7s linear infinite; flex-shrink: 0;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  .divider { border: none; border-top: 0.5px solid #e0e0d8; margin: 16px 0; }
</style>
</head>
<body>
<h1>CV Dashboard</h1>
<p class="subtitle" id="subtitle">Publications Master</p>
<div class="stats-grid" id="stats-grid">
  <div class="stat-card" style="grid-column:span 2">
    <p class="stat-label">Loading&hellip;</p><p class="stat-value">&mdash;</p>
  </div>
</div>
<hr class="divider">
<p class="section-label">Actions</p>
<div class="action-list">
  <button class="action-btn" onclick="run('sidebar_checkPubMed',
    'Searching PubMed&hellip; This may take 30&ndash;60 seconds.',
    'Done. Check the PubMed Log tab, approve matches with y, then click Apply.')">
    <span class="btn-icon">🔍</span>
    <span class="btn-text"><strong>Check PubMed for new publications</strong>
    <span>Finds papers that moved from accepted to published</span></span>
  </button>
  <button class="action-btn" onclick="run('sidebar_backfillDois',
    'Looking up missing DOIs and PMIDs&hellip;',
    'Done. Check PubMed Log and approve matches.')">
    <span class="btn-icon">🔗</span>
    <span class="btn-text"><strong>Fill in missing DOIs &amp; PMIDs</strong>
    <span>Finds IDs for papers that are missing them</span></span>
  </button>
  <button class="action-btn" onclick="run('sidebar_applyMatches',
    'Applying approved matches&hellip;',
    'Done. Approved rows updated in Publications Master.')">
    <span class="btn-icon">✅</span>
    <span class="btn-text"><strong>Apply approved matches</strong>
    <span>Run after reviewing the PubMed Log tab</span></span>
  </button>
  <button class="action-btn" onclick="run('sidebar_processEmails',
    'Processing emails in CV/process label&hellip;',
    'Done. Check Email Review Queue tab for items needing review.')">
    <span class="btn-icon">📬</span>
    <span class="btn-text"><strong>Process new emails</strong>
    <span>Reads forwarded submission &amp; acceptance emails</span></span>
  </button>
</div>
<p class="section-label">Status</p>
<div class="status-bar" id="status-bar">Ready. Select an action above.</div>
<script>
  function run(fnName, runningMsg, successMsg) {
    setStatus('running', runningMsg);
    document.querySelectorAll('.action-btn').forEach(b => b.disabled = true);
    google.script.run
      .withSuccessHandler(function() {
        setStatus('success', successMsg);
        document.querySelectorAll('.action-btn').forEach(b => b.disabled = false);
        loadStats();
      })
      .withFailureHandler(function(err) {
        setStatus('error', 'Error: ' + (err.message || String(err)));
        document.querySelectorAll('.action-btn').forEach(b => b.disabled = false);
      })
      [fnName]();
  }
  function setStatus(type, msg) {
    const bar = document.getElementById('status-bar');
    bar.className = 'status-bar ' + type;
    bar.innerHTML = (type === 'running' ? '<div class="spinner"></div>' : '') + '<span>' + msg + '</span>';
  }
  function loadStats() {
    google.script.run
      .withSuccessHandler(function(s) {
        const grid = document.getElementById('stats-grid');
        if (s.error) {
          grid.innerHTML = '<div class="stat-card" style="grid-column:span 2"><p class="stat-label">Error</p><p style="font-size:12px;color:#a33">' + s.error + '</p></div>';
          return;
        }
        grid.innerHTML =
          card('Total', s.total, false) +
          card('Published', s.published, true) +
          card('Accepted / in press', s.accepted + s.in_press, false) +
          card('Under review', s.under_review, false);
      })
      .withFailureHandler(function() {
        document.getElementById('stats-grid').innerHTML =
          '<div class="stat-card" style="grid-column:span 2"><p class="stat-label">Stats unavailable</p></div>';
      })
      .sidebar_getStats();
  }
  function card(label, value, highlight) {
    return '<div class="stat-card' + (highlight ? ' highlight' : '') + '">' +
      '<p class="stat-label">' + label + '</p>' +
      '<p class="stat-value">' + value + '</p></div>';
  }
  loadStats();
<\/script>
</body>
</html>`;
}

// =====================================================================
// PUBMED: BULK BACKFILL (accepted/in_press → published)
// =====================================================================

function findCandidates() {
  const ss     = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  if (!master) throw new Error(`Sheet "${MASTER_SHEET}" not found.`);

  const data    = master.getDataRange().getValues();
  const headers = data[0];
  const col     = name => headers.indexOf(name);

  const eligible = [];
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][col('status')] || '').toLowerCase().trim();
    if (status === 'accepted' || status === 'in_press') {
      eligible.push({
        rowIdx: i + 1,
        id:     data[i][col('id')],
        title:  String(data[i][col('title')] || ''),
        venue:  String(data[i][col('venue')]  || ''),
        year:   String(data[i][col('year')]   || ''),
      });
    }
  }
  if (!eligible.length) {
    SpreadsheetApp.getUi().alert('No rows with status "accepted" or "in_press" found.');
    return;
  }

  const pubmedRecords = fetchPubMedForAuthor(AUTHOR_QUERY, MAX_PUBMED_RESULTS);
  if (!pubmedRecords.length) {
    SpreadsheetApp.getUi().alert('PubMed returned no records. Check AUTHOR_QUERY in the script config.');
    return;
  }

  const logRows = [];
  eligible.forEach(row => {
    const best = findBestMatch(row, pubmedRecords);
    if (best && best.score >= MATCH_THRESHOLD) {
      logRows.push([
        '', row.id, row.rowIdx, best.score.toFixed(2),
        row.title, best.record.title,
        row.year,  best.record.year,
        row.venue, best.record.journal,
        best.record.volume, best.record.issue, best.record.pages,
        best.record.doi, best.record.pmid,
      ]);
    }
  });

  writeLog(logRows);
  SpreadsheetApp.getUi().alert(
    `Found ${logRows.length} candidate match${logRows.length === 1 ? '' : 'es'} ` +
    `from ${eligible.length} accepted/in-press rows.\n\n` +
    `Open the "PubMed Log" tab, type "y" in column A to approve, ` +
    `then click CV Tools → Apply approved matches.`
  );
}

// =====================================================================
// PUBMED: DOI/PMID BACKFILL
// =====================================================================

function findDoiCandidates() {
  const ss     = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  if (!master) throw new Error(`Sheet "${MASTER_SHEET}" not found.`);

  const data    = master.getDataRange().getValues();
  const headers = data[0];
  const col     = name => headers.indexOf(name);

  const eligible = [];
  for (let i = 1; i < data.length; i++) {
    const doi   = String(data[i][col('doi')]   || '').trim();
    const pmid  = String(data[i][col('pmid')]  || '').trim();
    const title = String(data[i][col('title')] || '').trim();
    if (title && (!doi || !pmid)) {
      eligible.push({
        rowIdx:  i + 1,
        id:      data[i][col('id')],
        title,
        venue:   String(data[i][col('venue')]  || ''),
        year:    String(data[i][col('year')]   || ''),
        hasDoi:  !!doi,
        hasPmid: !!pmid,
      });
    }
  }
  if (!eligible.length) {
    SpreadsheetApp.getUi().alert('All rows already have DOI and PMID. Nothing to do!');
    return;
  }

  const pubmedRecords = fetchPubMedForAuthor(AUTHOR_QUERY, MAX_PUBMED_RESULTS);
  if (!pubmedRecords.length) {
    SpreadsheetApp.getUi().alert('PubMed returned no records.');
    return;
  }

  const logRows = [];
  eligible.forEach(row => {
    const best = findBestMatch(row, pubmedRecords);
    if (best && best.score >= MATCH_THRESHOLD) {
      if (!(!row.hasDoi && best.record.doi) && !(!row.hasPmid && best.record.pmid)) return;
      logRows.push([
        '', row.id, row.rowIdx, best.score.toFixed(2),
        row.title, best.record.title,
        row.year,  best.record.year,
        row.venue, best.record.journal,
        best.record.volume, best.record.issue, best.record.pages,
        best.record.doi, best.record.pmid,
      ]);
    }
  });

  writeLog(logRows);
  SpreadsheetApp.getUi().alert(
    `Found ${logRows.length} row${logRows.length === 1 ? '' : 's'} with new DOI/PMID info.\n\n` +
    `Review the PubMed Log, approve with "y", then Apply approved matches.\n` +
    `Note: only empty fields will be filled — existing data is never overwritten.`
  );
}

// =====================================================================
// PUBMED: APPLY APPROVED MATCHES
// =====================================================================

function applyApprovedMatches() {
  const ss     = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  const log    = ss.getSheetByName(LOG_SHEET);
  if (!log) throw new Error('No PubMed Log tab — run a Find command first.');

  const logData = log.getDataRange().getValues();
  if (logData.length < 2) { SpreadsheetApp.getUi().alert('PubMed Log is empty.'); return; }

  const headers = master.getDataRange().getValues()[0];
  const col     = name => headers.indexOf(name) + 1;

  let applied = 0;
  for (let i = 1; i < logData.length; i++) {
    const approve = String(logData[i][0]).trim().toLowerCase();
    if (approve !== 'y' && approve !== 'yes') continue;
    const masterRow = Number(logData[i][2]);
    if (!masterRow) continue;

    const [, , , , , , , pmYear, , pmJournal, vol, issue, pages, doi, pmid] = logData[i];
    const fillIfEmpty = (columnName, newValue) => {
      if (!newValue) return;
      const cell = master.getRange(masterRow, col(columnName));
      if (!String(cell.getValue()).trim()) cell.setValue(newValue);
    };

    fillIfEmpty('year',   pmYear);
    fillIfEmpty('venue',  pmJournal);
    fillIfEmpty('volume', vol);
    fillIfEmpty('issue',  issue);
    fillIfEmpty('pages',  pages);
    fillIfEmpty('doi',    doi);
    fillIfEmpty('pmid',   pmid);

    const currentStatus = String(master.getRange(masterRow, col('status')).getValue()).toLowerCase();
    if (currentStatus === 'accepted' || currentStatus === 'in_press') {
      master.getRange(masterRow, col('status')).setValue('published');
    }

    log.getRange(i + 1, 1).setValue('applied');
    applied++;
  }
  SpreadsheetApp.getUi().alert(`Applied ${applied} update${applied === 1 ? '' : 's'} to ${MASTER_SHEET}.`);
}

// =====================================================================
// PUBMED: MANUAL SINGLE-ROW LOOKUP
// =====================================================================

function lookupSelectedRow() {
  const ss     = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  if (!master) throw new Error(`Sheet "${MASTER_SHEET}" not found.`);

  const active = ss.getActiveSheet();
  if (active.getName() !== MASTER_SHEET) {
    SpreadsheetApp.getUi().alert(`Switch to the "${MASTER_SHEET}" tab first, then click on the row you want to update.`);
    return;
  }
  const row = active.getActiveRange().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert('Click on a data row (not the header).'); return; }

  const headers      = master.getDataRange().getValues()[0];
  const col          = name => headers.indexOf(name) + 1;
  const currentTitle = master.getRange(row, col('title')).getValue();
  const currentId    = master.getRange(row, col('id')).getValue();

  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    `Lookup for ${currentId}`,
    `Title: "${currentTitle}"\n\nPaste a PMID (digits) or DOI (starts with 10.):`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const input = resp.getResponseText().trim();
  if (!input) return;

  let record;
  try {
    if (/^\d+$/.test(input)) {
      record = fetchPubMedByPmid(input);
    } else if (/^10\./.test(input)) {
      record = fetchPubMedByDoi(input);
      if (!record) record = fetchCrossRefByDoi(input);
    } else {
      ui.alert('Input must be a PMID (digits only) or DOI (starts with "10.").');
      return;
    }
  } catch (e) { ui.alert(`Lookup failed: ${e.message}`); return; }

  if (!record) { ui.alert('No record found.'); return; }

  const confirm = ui.alert(
    'Confirm match',
    `"${record.title}"\n${record.journal} ${record.year}` +
    (record.volume ? `, ${record.volume}` : '') +
    (record.issue  ? `(${record.issue})`  : '') +
    (record.pages  ? `:${record.pages}`   : '') +
    `\n\nApply to row ${row} (${currentId})?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  if (record.year)    master.getRange(row, col('year')).setValue(record.year);
  if (record.journal) master.getRange(row, col('venue')).setValue(record.journal);
  if (record.volume)  master.getRange(row, col('volume')).setValue(record.volume);
  if (record.issue)   master.getRange(row, col('issue')).setValue(record.issue);
  if (record.pages)   master.getRange(row, col('pages')).setValue(record.pages);
  if (record.doi)     master.getRange(row, col('doi')).setValue(record.doi);
  if (record.pmid)    master.getRange(row, col('pmid')).setValue(record.pmid);
  master.getRange(row, col('status')).setValue('published');
  ui.alert(`Updated row ${row}.`);
}

// =====================================================================
// PUBMED / CROSSREF API
// =====================================================================

function fetchPubMedForAuthor(query, retmax) {
  const searchUrl  = `${NCBI_BASE}/esearch.fcgi?db=pubmed&retmode=json&retmax=${retmax}&term=${encodeURIComponent(query)}`;
  const searchResp = JSON.parse(UrlFetchApp.fetch(searchUrl).getContentText());
  const pmids      = (searchResp.esearchresult && searchResp.esearchresult.idlist) || [];
  if (!pmids.length) return [];

  const summaryUrl  = `${NCBI_BASE}/esummary.fcgi?db=pubmed&retmode=json&id=${pmids.join(',')}`;
  const summaryResp = JSON.parse(UrlFetchApp.fetch(summaryUrl).getContentText());
  const result      = summaryResp.result || {};

  return pmids.reduce((acc, pmid) => {
    const r = result[pmid];
    if (!r) return acc;
    let doi = '';
    (r.articleids || []).forEach(a => { if (a.idtype === 'doi') doi = a.value; });
    acc.push({
      pmid,
      title:   String(r.title || '').replace(/\.$/, ''),
      journal: r.fulljournalname || r.source || '',
      year:    (r.pubdate || '').slice(0, 4),
      volume:  r.volume || '',
      issue:   r.issue  || '',
      pages:   r.pages  || '',
      doi,
    });
    return acc;
  }, []);
}

function fetchPubMedByPmid(pmid) {
  const url  = `${NCBI_BASE}/esummary.fcgi?db=pubmed&retmode=json&id=${pmid}`;
  const resp = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  const r    = resp.result && resp.result[pmid];
  if (!r || r.error) return null;
  let doi = '';
  (r.articleids || []).forEach(a => { if (a.idtype === 'doi') doi = a.value; });
  return {
    pmid,
    title:   String(r.title || '').replace(/\.$/, ''),
    journal: r.fulljournalname || r.source || '',
    year:    (r.pubdate || '').slice(0, 4),
    volume:  r.volume || '',
    issue:   r.issue  || '',
    pages:   r.pages  || '',
    doi,
  };
}

function fetchPubMedByDoi(doi) {
  const url  = `${NCBI_BASE}/esearch.fcgi?db=pubmed&retmode=json&term=${encodeURIComponent(doi + '[AID]')}`;
  const resp = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  const pmids = (resp.esearchresult && resp.esearchresult.idlist) || [];
  return pmids.length ? fetchPubMedByPmid(pmids[0]) : null;
}

function fetchCrossRefByDoi(doi) {
  try {
    const url  = `https://api.crossref.org/works/${encodeURIComponent(doi)}`;
    const resp = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    const m    = resp.message;
    if (!m) return null;
    const year = m.issued && m.issued['date-parts'] && m.issued['date-parts'][0] && m.issued['date-parts'][0][0];
    return {
      pmid: '', doi: m.DOI || doi,
      title:   (m.title && m.title[0]) || '',
      journal: (m['container-title'] && m['container-title'][0]) || '',
      year:    String(year || ''),
      volume:  m.volume || '',
      issue:   m.issue  || '',
      pages:   m.page   || '',
    };
  } catch (e) { return null; }
}

// =====================================================================
// TITLE MATCHING
// =====================================================================

function findBestMatch(row, records) {
  let best = null;
  records.forEach(rec => {
    const score = titleSimilarity(row.title, rec.title);
    if (!best || score > best.score) best = { record: rec, score };
  });
  return best;
}

function titleSimilarity(a, b) {
  const ta = tokenize(a), tb = tokenize(b);
  if (!ta.length || !tb.length) return 0;
  const sa = new Set(ta), sb = new Set(tb);
  let inter = 0;
  sa.forEach(t => { if (sb.has(t)) inter++; });
  const jaccard = inter / (sa.size + sb.size - inter);
  const n = Math.min(8, ta.length, tb.length);
  let ordered = 0;
  for (let i = 0; i < n; i++) if (ta[i] === tb[i]) ordered++;
  return 0.75 * jaccard + 0.25 * (n ? ordered / n : 0);
}

function tokenize(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/)
    .filter(t => t && !STOPWORDS.has(t));
}

// =====================================================================
// LOG SHEET
// =====================================================================

function writeLog(rows) {
  const ss = SpreadsheetApp.getActive();
  let log  = ss.getSheetByName(LOG_SHEET);
  if (!log) log = ss.insertSheet(LOG_SHEET);
  log.clear();
  log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
  log.setFrozenRows(1);
  if (rows.length) log.getRange(2, 1, rows.length, LOG_HEADERS.length).setValues(rows);
  if (rows.length > 1) log.getRange(2, 1, rows.length, LOG_HEADERS.length).sort({ column: 4, ascending: false });
  log.autoResizeColumns(1, LOG_HEADERS.length);
  log.setColumnWidth(1, 110);
}

function clearLog() {
  const log = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  if (log) log.clear();
}

// =====================================================================
// TALKS: GOOGLE FORM
// =====================================================================

function createTalkForm() {
  const ss    = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  const ui    = SpreadsheetApp.getUi();

  const existing = props.getProperty('talkFormId');
  if (existing) {
    const resp = ui.alert('A talk form already exists', `ID: ${existing}\n\nCreate a new one anyway?`, ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  const form = FormApp.create('Add a Talk — CV');
  form.setCollectEmail(false).setRequireLogin(false);
  form.addTextItem().setTitle('Talk title').setRequired(false);
  form.addTextItem().setTitle('Venue / conference / institution').setRequired(true);
  form.addTextItem().setTitle('Location (City, Region/Country)').setRequired(true);
  form.addListItem().setTitle('Month').setChoiceValues(
    ['01 Jan','02 Feb','03 Mar','04 Apr','05 May','06 Jun',
     '07 Jul','08 Aug','09 Sep','10 Oct','11 Nov','12 Dec']
  ).setRequired(true);
  const y = new Date().getFullYear();
  const yrs = [];
  for (let i = y + 2; i >= y - 1; i--) yrs.push(String(i));
  form.addListItem().setTitle('Year').setChoiceValues(yrs).setRequired(true);
  form.addListItem().setTitle('Type').setChoiceValues(
    ['invited','keynote','lecture','grand_rounds','workshop','symposium','conference_paper','panel']
  ).setRequired(true);
  form.addMultipleChoiceItem().setTitle('Virtual?').setChoiceValues(['In person','Virtual','Hybrid']).setRequired(true);
  form.addParagraphTextItem().setTitle('Notes (optional)').setRequired(false);

  props.setProperty('talkFormId',  form.getId());
  props.setProperty('talkFormUrl', form.getPublishedUrl());

  let linkedOk = false;
  try { form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId()); linkedOk = true; }
  catch (e) { Logger.log('Form link failed: ' + e.message); }

  if (linkedOk) {
    try {
      Utilities.sleep(1500);
      const rs = ss.getSheets().filter(s => /Form Responses/i.test(s.getName()));
      if (rs.length && !ss.getSheetByName(FORM_RESPONSES_SHEET)) rs[rs.length - 1].setName(FORM_RESPONSES_SHEET);
    } catch (e) { Logger.log('Rename failed: ' + e.message); }
  }

  try {
    if (!ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'handleTalkFormSubmit').length) {
      ScriptApp.newTrigger('handleTalkFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
    }
  } catch (e) { Logger.log('Trigger creation failed: ' + e.message); }

  ui.alert('Form created!', `URL:\n\n${form.getPublishedUrl()}\n\nBookmark this on your phone.` +
    (linkedOk ? '' : '\n\nNote: form created but could not auto-link to the sheet.'), ui.ButtonSet.OK);
}

function showTalkFormUrl() {
  const url = PropertiesService.getDocumentProperties().getProperty('talkFormUrl');
  if (!url) { SpreadsheetApp.getUi().alert('No form found. Run CV Tools → Create talk form first.'); return; }
  SpreadsheetApp.getUi().alert('Talk form URL', url, SpreadsheetApp.getUi().ButtonSet.OK);
}

function handleTalkFormSubmit(e) {
  if (!e || !e.namedValues || !('Talk title' in e.namedValues)) return;
  const nv    = e.namedValues;
  const ss    = SpreadsheetApp.getActive();
  const talks = ss.getSheetByName(TALKS_SHEET);
  if (!talks) return;

  const headers  = talks.getDataRange().getValues()[0];
  const colIndex = name => headers.indexOf(name);
  const ids = talks.getRange(2, colIndex('id') + 1, talks.getLastRow() - 1, 1).getValues();
  let maxNum = 0;
  ids.forEach(r => { const m = String(r[0]).match(/TALK(\d+)/); if (m) maxNum = Math.max(maxNum, Number(m[1])); });
  const newId = 'TALK' + String(maxNum + 1).padStart(4, '0');

  const f = k => (nv[k] && nv[k][0]) || '';
  const vr = f('Virtual?');
  const newRow = headers.map(h => {
    switch (h) {
      case 'id':       return newId;
      case 'year':     return f('Year');
      case 'month':    return f('Month').slice(0, 2);
      case 'type':     return f('Type');
      case 'title':    return f('Talk title');
      case 'venue':    return f('Venue / conference / institution');
      case 'location': return f('Location (City, Region/Country)');
      case 'virtual':  return vr === 'Virtual' ? 'yes' : vr === 'Hybrid' ? 'hybrid' : '';
      case 'notes':    return f('Notes (optional)');
      default:         return '';
    }
  });
  talks.appendRow(newRow);
  const lr = talks.getLastRow();
  if (lr > 2) {
    talks.getRange(2, 1, lr - 1, talks.getLastColumn()).sort([
      { column: colIndex('year')  + 1, ascending: false },
      { column: colIndex('month') + 1, ascending: false },
    ]);
  }
}

// =====================================================================
// EMAIL INTAKE (Gemini-powered)
// =====================================================================

function processIntakeEmails() {
  const label = GmailApp.getUserLabelByName(INTAKE_LABEL);
  if (!label) {
    SpreadsheetApp.getUi().alert(`No Gmail label "${INTAKE_LABEL}" found.\n\nCreate it in Gmail, then forward emails to your +cv alias.`);
    return;
  }

  let processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL) || GmailApp.createLabel(PROCESSED_LABEL);
  let failedLabel    = GmailApp.getUserLabelByName(FAILED_LABEL)    || GmailApp.createLabel(FAILED_LABEL);

  const threads = label.getThreads(0, 50);
  if (!threads.length) { SpreadsheetApp.getUi().alert('No new emails to process.'); return; }

  const ss      = SpreadsheetApp.getActive();
  const master  = ss.getSheetByName(MASTER_SHEET);
  const headers = master.getDataRange().getValues()[0];
  const col     = name => headers.indexOf(name) + 1;
  const props   = PropertiesService.getScriptProperties();

  let autoApplied = 0, queued = 0, errors = 0, permFailed = 0;

  threads.forEach(thread => {
    const threadId   = thread.getId();
    const retryKey   = 'retry_' + threadId;
    const retryCount = Number(props.getProperty(retryKey) || 0);
    const msg        = thread.getMessages()[0];
    const subject    = msg.getSubject();
    const body       = msg.getPlainBody().slice(0, 8000);

    try {
      const extracted = extractWithGemini(subject, body);
      if (!extracted) throw new Error('Extraction returned null');
      const result = applyOrQueue(master, headers, col, extracted, thread, subject, body);
      if (result === 'applied') autoApplied++;
      else if (result === 'queued') queued++;
      thread.removeLabel(label);
      thread.addLabel(processedLabel);
      props.deleteProperty(retryKey);
    } catch (e) {
      errors++;
      const n = retryCount + 1;
      if (n >= MAX_RETRY_COUNT) {
        addToReviewQueue(thread, subject, body, null, 0, `Failed after ${MAX_RETRY_COUNT} attempts: ${e.message}`);
        thread.removeLabel(label);
        thread.addLabel(failedLabel);
        props.deleteProperty(retryKey);
        permFailed++;
      } else {
        props.setProperty(retryKey, String(n));
      }
    }
  });

  let msg = `Processed ${threads.length} email${threads.length === 1 ? '' : 's'}:\n\n` +
    `  Auto-applied: ${autoApplied}\n  Queued for review: ${queued}\n` +
    `  Transient errors (will retry): ${errors - permFailed}\n  Permanently failed: ${permFailed}`;
  if (queued || permFailed) msg += `\n\nCheck "${REVIEW_SHEET}" tab.`;
  SpreadsheetApp.getUi().alert(msg);
}

function extractWithGemini(subject, body) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('No GEMINI_API_KEY found in Script Properties. See setup instructions.');

  const prompt = `Extract manuscript metadata from this forwarded email.
Return ONLY a JSON object with these fields:
- event: "submission" | "acceptance" | "revision_request" | "rejection" | "published" | "unknown"
- title: manuscript title or ""
- journal: journal name or ""
- manuscript_id: tracking ID or ""
- doi: DOI or ""
- confidence: 0.0–1.0
No markdown, no explanation. If not a manuscript lifecycle email, set event="unknown" and confidence=0.

SUBJECT: ${subject}
BODY: ${body}`;

  const url  = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json' },
    }),
  });

  if (resp.getResponseCode() !== 200) throw new Error(`Gemini ${resp.getResponseCode()}: ${resp.getContentText().slice(0, 200)}`);

  const text = JSON.parse(resp.getContentText())
    ?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) return null;
  try { return JSON.parse(text); }
  catch (e) { return JSON.parse(text.replace(/```json\s*|```\s*$/g, '').trim()); }
}

function applyOrQueue(master, headers, col, extracted, thread, subject, body) {
  const { event, title, journal, doi, confidence } = extracted;
  if (event === 'unknown' || !event) {
    addToReviewQueue(thread, subject, body, extracted, confidence || 0, 'Not a manuscript lifecycle email');
    return 'queued';
  }

  const data     = master.getDataRange().getValues();
  const titleCol = col('title') - 1;
  const doiCol   = col('doi')   - 1;
  let matchRow   = -1, matchScore = 0;

  if (title) {
    for (let i = 1; i < data.length; i++) {
      const s = titleSimilarity(title, data[i][titleCol]);
      if (s > matchScore) { matchScore = s; matchRow = i; }
    }
  }
  if (doi) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][doiCol]).trim().toLowerCase() === doi.toLowerCase()) {
        matchRow = i; matchScore = 1.0; break;
      }
    }
  }

  const found  = matchRow >= 0 && matchScore >= 0.70;
  const hiConf = (confidence || 0) >= AUTO_APPLY_CONFIDENCE;

  if (event === 'submission') {
    if (found) {
      addToReviewQueue(thread, subject, body, extracted, confidence, `Submission but matched existing row ${matchRow + 1} (score ${matchScore.toFixed(2)})`);
      return 'queued';
    }
    if (!hiConf || !title || !journal) {
      addToReviewQueue(thread, subject, body, extracted, confidence, 'Submission: low confidence or missing fields');
      return 'queued';
    }
    addNewSubmissionRow(master, headers, col, extracted);
    return 'applied';
  }

  if (event === 'acceptance') {
    if (!found) {
      addToReviewQueue(thread, subject, body, extracted, confidence, 'Acceptance: no matching row found');
      return 'queued';
    }
    if (!hiConf) {
      addToReviewQueue(thread, subject, body, extracted, confidence, `Acceptance: low confidence (score ${matchScore.toFixed(2)})`);
      return 'queued';
    }
    master.getRange(matchRow + 1, col('status')).setValue('accepted');
    if (doi && !data[matchRow][doiCol]) master.getRange(matchRow + 1, col('doi')).setValue(doi);
    return 'applied';
  }

  addToReviewQueue(thread, subject, body, extracted, confidence, `Event: ${event} (not auto-applied)`);
  return 'queued';
}

function addNewSubmissionRow(master, headers, col, extracted) {
  const allIds = master.getRange(2, col('id'), master.getLastRow() - 1, 1).getValues();
  let maxNum = 0;
  allIds.forEach(r => { const m = String(r[0]).match(/PUB(\d+)/); if (m) maxNum = Math.max(maxNum, Number(m[1])); });
  const newId = 'PUB' + String(maxNum + 1).padStart(4, '0');
  master.appendRow(headers.map(h => {
    switch (h) {
      case 'id':             return newId;
      case 'category':       return 'journal';
      case 'status':         return 'under_review';
      case 'year':           return String(new Date().getFullYear());
      case 'title':          return extracted.title  || '';
      case 'venue':          return extracted.journal || '';
      case 'doi':            return extracted.doi    || '';
      case 'submitted_date': return new Date().toISOString().slice(0, 10);
      case 'notes':          return extracted.manuscript_id ? `Manuscript ID: ${extracted.manuscript_id}` : '';
      default:               return '';
    }
  }));
}

const REVIEW_HEADERS = [
  'Timestamp', 'Apply?', 'Thread link', 'Reason',
  'Event', 'Confidence', 'Title', 'Journal', 'DOI', 'Manuscript ID',
  'Email subject', 'Email body (first 500 chars)',
];

function addToReviewQueue(thread, subject, body, extracted, confidence, reason) {
  const ss = SpreadsheetApp.getActive();
  let queue = ss.getSheetByName(REVIEW_SHEET);
  if (!queue) {
    queue = ss.insertSheet(REVIEW_SHEET);
    queue.getRange(1, 1, 1, REVIEW_HEADERS.length).setValues([REVIEW_HEADERS]).setFontWeight('bold');
    queue.setFrozenRows(1);
  }
  queue.appendRow([
    new Date(), '', thread ? `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}` : '',
    reason,
    extracted ? extracted.event          : '',
    confidence,
    extracted ? extracted.title          : '',
    extracted ? extracted.journal        : '',
    extracted ? extracted.doi            : '',
    extracted ? extracted.manuscript_id  : '',
    subject, body.slice(0, 500),
  ]);
}

function clearReviewQueue() {
  const q = SpreadsheetApp.getActive().getSheetByName(REVIEW_SHEET);
  if (q) q.clear();
}

function testGeminiConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = extractWithGemini(
      'Manuscript accepted: "A Test Paper on Nothing"',
      'Dear Author, We are pleased to accept "A Test Paper on Nothing" (TST-2026-001) for publication in Journal of Testing. DOI: 10.1234/test.2026. Regards, The Editors.'
    );
    ui.alert('Gemini connection OK', JSON.stringify(r, null, 2), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Gemini connection failed', e.message, ui.ButtonSet.OK);
  }
}

// =====================================================================
// CV EXPORT — Google Doc
// =====================================================================

function exportCvToDoc() {
  const ss     = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  const talks  = ss.getSheetByName(TALKS_SHEET);
  if (!master) throw new Error(`Sheet "${MASTER_SHEET}" not found.`);

  const pubData    = master.getDataRange().getValues();
  const pubHeaders = pubData[0];
  const pcol       = name => pubHeaders.indexOf(name);

  const pubs = [];
  for (let i = 1; i < pubData.length; i++) {
    const r = pubData[i];
    pubs.push({
      category: String(r[pcol('category')] || '').toLowerCase().trim(),
      status:   String(r[pcol('status')]   || '').toLowerCase().trim(),
      year:     String(r[pcol('year')]     || '').trim(),
      authors:  String(r[pcol('authors')]  || '').trim(),
      title:    String(r[pcol('title')]    || '').trim(),
      venue:    String(r[pcol('venue')]    || '').trim(),
      volume:   String(r[pcol('volume')]   || '').trim(),
      issue:    String(r[pcol('issue')]    || '').trim(),
      pages:    String(r[pcol('pages')]    || '').trim(),
      doi:      String(r[pcol('doi')]      || '').trim(),
    });
  }

  const talkRows = [];
  if (talks) {
    const td = talks.getDataRange().getValues();
    const th = td[0];
    const tc = name => th.indexOf(name);
    for (let i = 1; i < td.length; i++) {
      const r = td[i];
      talkRows.push({
        year: String(r[tc('year')] || '').trim(), month: String(r[tc('month')] || '').trim(),
        type: String(r[tc('type')] || '').trim(), title: String(r[tc('title')] || '').trim(),
        venue: String(r[tc('venue')] || '').trim(), location: String(r[tc('location')] || '').trim(),
        virtual: String(r[tc('virtual')] || '').trim(), notes: String(r[tc('notes')] || '').trim(),
      });
    }
  }

  const SECTIONS = [
    { label: 'Journal Articles',  category: 'journal' },
    { label: 'Reviews',           category: 'review' },
    { label: 'Book Chapters',     category: 'book_chapter' },
    { label: 'Conference Papers', category: 'conference' },
    { label: 'Editorials',        category: 'editorial' },
  ];

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const doc  = DocumentApp.create(`CV Export — ${timestamp}`);
  const body = doc.getBody();
  body.setMarginTop(72).setMarginBottom(72).setMarginLeft(72).setMarginRight(72);

  const addHeading = text => {
    const p = body.appendParagraph(text);
    p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const s = {};
    s[DocumentApp.Attribute.FONT_SIZE]       = 13;
    s[DocumentApp.Attribute.BOLD]             = true;
    s[DocumentApp.Attribute.SPACE_BEFORE]     = 16;
    s[DocumentApp.Attribute.SPACE_AFTER]      = 4;
    p.setAttributes(s);
  };

  const addCitation = (n, pub) => {
    const parts = [];
    if (pub.authors) parts.push(pub.authors + '.');
    if (pub.title)   parts.push(pub.title + '.');
    let v = pub.venue || '';
    if (pub.year)   v += (v ? '. ' : '') + pub.year;
    if (pub.volume) { v += ';' + pub.volume; if (pub.issue) v += '(' + pub.issue + ')'; }
    if (pub.pages)  v += ':' + pub.pages;
    if (v)          parts.push(v + '.');
    if (pub.doi)    parts.push('doi:' + pub.doi);
    const badge = pub.status && pub.status !== 'published' ? ' [' + pub.status.replace(/_/g, ' ') + ']' : '';
    const p = body.appendParagraph('');
    p.setIndentFirstLine(0).setIndentStart(36).setSpacingAfter(4);
    p.appendText(n + '. ').setBold(false).setFontSize(10);
    p.appendText(parts.join(' ') + badge).setBold(false).setFontSize(10);
  };

  const sortDesc = arr => arr.slice().sort((a, b) => (Number(b.year) || 0) - (Number(a.year) || 0));

  // Title block
  const tp = body.appendParagraph('CV Export');
  tp.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const ts = {};
  ts[DocumentApp.Attribute.FONT_SIZE] = 16; ts[DocumentApp.Attribute.BOLD] = true; ts[DocumentApp.Attribute.SPACE_AFTER] = 4;
  tp.setAttributes(ts);
  const dp = body.appendParagraph('Generated: ' + timestamp);
  const ds = {}; ds[DocumentApp.Attribute.FONT_SIZE] = 9; ds[DocumentApp.Attribute.SPACE_AFTER] = 12;
  dp.setAttributes(ds);

  let n = 1;
  SECTIONS.forEach(sec => {
    const sp = sortDesc(pubs.filter(p => p.category === sec.category));
    if (!sp.length) return;
    addHeading(sec.label + ' (' + sp.length + ')');
    sp.forEach(pub => addCitation(n++, pub));
  });

  if (talkRows.length) {
    const st = talkRows.slice().sort((a, b) => {
      const yd = (Number(b.year) || 0) - (Number(a.year) || 0);
      return yd !== 0 ? yd : (Number(b.month) || 0) - (Number(a.month) || 0);
    });
    addHeading('Talks & Presentations (' + st.length + ')');
    const MN = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    st.forEach((t, i) => {
      const parts = [];
      if (t.title)    parts.push(t.title + '.');
      if (t.venue)    parts.push(t.venue + '.');
      let loc = t.location || '';
      if (t.virtual === 'yes')    loc = loc ? loc + ' (virtual)' : 'Virtual';
      if (t.virtual === 'hybrid') loc = loc ? loc + ' (hybrid)'  : 'Hybrid';
      if (loc) parts.push(loc + '.');
      const m = parseInt(t.month, 10);
      if (t.month && t.year) parts.push((MN[m] || t.month) + ' ' + t.year + '.');
      else if (t.year) parts.push(t.year + '.');
      if (t.type)  parts.push('[' + t.type + ']');
      if (t.notes) parts.push(t.notes);
      const p = body.appendParagraph('');
      p.setIndentFirstLine(0).setIndentStart(36).setSpacingAfter(4);
      p.appendText((i + 1) + '. ').setBold(false).setFontSize(10);
      p.appendText(parts.join(' ')).setBold(false).setFontSize(10);
    });
  }

  doc.saveAndClose();
  SpreadsheetApp.getUi().alert(
    'CV exported!',
    `Saved to Google Drive as:\n"CV Export — ${timestamp}"\n\n${doc.getUrl()}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

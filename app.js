// ─────────────────────────────────────────────────────────────
// PDF Assignment Renamer — app.js
// ─────────────────────────────────────────────────────────────

// pdfjsLib is initialised lazily in initPdfJs() so a failed import
// does NOT prevent the folder/Excel pickers from working.
let pdfjsLib = null;

async function initPdfJs() {
  if (pdfjsLib) return;
  pdfjsLib = await import('./lib/pdf.min.js');
  pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL('lib/pdf.worker.min.mjs');
}

// ─────────────────────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────────────────────
const state = {
  dirHandle: null,        // FileSystemDirectoryHandle for the PDF folder
  pdfFiles: [],           // [{ name, handle }]
  students: [],           // [{ name, atsId }] — optional; empty = no reference file loaded
  results: [],            // one entry per file after OCR
  tesseractWorker: null,
  template: {             // renaming template configured by user in Step 1
    pattern: '',          // e.g. "ENE61_GC5.1_5.3_P_{id}.{ext}"
  },
  customTokens: {},       // tokens discovered by "Scan a file" e.g. { title: 'Milestone 1' }
  finalReport: [],
};

// ─────────────────────────────────────────────────────────────
// STEP NAVIGATION
// ─────────────────────────────────────────────────────────────
function showStep(n) {
  [1, 2, 3, 4].forEach(i => {
    document.getElementById(`step-${i}`).classList.toggle('hidden', i !== n);
    const tab = document.getElementById(`tab-${i}`);
    tab.classList.remove('active', 'done');
    if (i === n) tab.classList.add('active');
    if (i < n) tab.classList.add('done');
  });
}

// ─────────────────────────────────────────────────────────────
// STEP 1 — FILE LOADING
// ─────────────────────────────────────────────────────────────
async function pickFolder() {
  try {
    state.dirHandle = await window.showDirectoryPicker();
    state.pdfFiles = [];
    for await (const [name, handle] of state.dirHandle.entries()) {
      // Accept all file types — skip hidden files and the output subfolder
      if (handle.kind === 'file' && !name.startsWith('.')) {
        state.pdfFiles.push({ name, handle });
      }
    }
    state.pdfFiles.sort((a, b) => a.name.localeCompare(b.name));

    const zone = document.getElementById('zone-pdf');
    zone.classList.add('loaded');
    zone.removeEventListener('click', pickFolder);
    zone.querySelector('h3').textContent = `${state.pdfFiles.length} files loaded`;
    zone.querySelector('p').textContent = state.dirHandle.name;
    const tag = zone.querySelector('.status-tag');
    tag.textContent = `${state.pdfFiles.length} files`;
    tag.classList.remove('hidden');
    updateStartBtn();
    // Reveal the "Scan a file" button now that we have files
    document.getElementById('scan-section')?.classList.remove('hidden');
  } catch (e) {
    if (e.name !== 'AbortError') alert('Could not open folder: ' + e.message);
  }
};

async function pickExcel() {
  try {
    const [fh] = await window.showOpenFilePicker({
      types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'] } }],
    });
    const file = await fh.getFile();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find the header row (contains "Student Name" or "Name")
    let headerRow = -1;
    let nameCol = 1, idCol = 2;
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const rowStr = rows[r].map(c => String(c).toLowerCase()).join('|');
      if (rowStr.includes('student name') || rowStr.includes('student id')) {
        headerRow = r;
        // Find column indexes
        rows[r].forEach((cell, ci) => {
          const s = String(cell).toLowerCase();
          if (s.includes('name')) nameCol = ci;
          if (s.includes('id')) idCol = ci;
        });
        break;
      }
    }

    state.students = [];
    for (let r = headerRow + 1; r < rows.length; r++) {
      const row = rows[r];
      const name = String(row[nameCol] || '').trim();
      const atsId = String(row[idCol] || '').trim();
      if (name.length > 3 && !/^\d+$/.test(name)) {
        state.students.push({ name, atsId });
      }
    }

    const zone = document.getElementById('zone-excel');
    zone.classList.add('loaded');
    zone.removeEventListener('click', pickExcel);
    zone.querySelector('h3').textContent = `${state.students.length} students loaded`;
    zone.querySelector('p').textContent = file.name;
    const tag = zone.querySelector('.status-tag');
    tag.textContent = `${state.students.length} students`;
    tag.classList.remove('hidden');
    updateStartBtn();
  } catch (e) {
    if (e.name !== 'AbortError') alert('Could not read Excel file: ' + e.message);
  }
};

function validateTemplate() {
  return (state.template.pattern || '').trim().length > 0;
}

// True when pattern contains {name} or {id} — roster is required in that case
function needsRoster() {
  return /\{(name|id)\}/i.test(state.template.pattern || '');
}

// Update the excel zone badge (Required/Optional) based on current slots
function updateExcelZoneBadge() {
  const badge = document.getElementById('excel-badge');
  const desc  = document.getElementById('excel-zone-desc');
  if (!badge) return;
  const required = needsRoster();
  badge.textContent = required ? 'Required' : 'Optional';
  badge.style.background = required ? '#fee2e2' : '#fef9c3';
  badge.style.color      = required ? '#991b1b' : '#854d0e';
  badge.style.border     = required ? '1px solid #fecaca' : '1px solid #fde68a';
  if (desc) {
    desc.textContent = required
      ? 'Upload your .xlsx roster — needed to match names and IDs found in the files'
      : 'Optional — not needed when all slots use Fixed Template text only';
  }
}

function updateStartBtn() {
  const btn   = document.getElementById('start-btn');
  const rosterOk = !needsRoster() || state.students.length > 0;
  const ready    = state.pdfFiles.length > 0 && rosterOk && validateTemplate();
  btn.disabled = !ready;
  if (ready) {
    const rosterInfo = state.students.length > 0
      ? ` · ${state.students.length} roster entries`
      : needsRoster() ? '' : ' · no roster needed';
    btn.innerHTML = `<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
    Start Renaming (${state.pdfFiles.length} files${rosterInfo})`;
  }
}



// ─────────────────────────────────────────────────────────────
// STEP 2 — FILENAME → ROSTER MATCHING (no OCR)
// ─────────────────────────────────────────────────────────────
// OCR was unreliable (failed ~80% of the time on grey-cell labels).
// The filename already contains the student's name, which is a reliable
// source. Matching filename against the roster is fast and accurate.
async function startProcessing() {
  showStep(2);

  const hasRoster = state.students.length > 0;
  addLog(hasRoster
    ? `Matching ${state.pdfFiles.length} filenames against ${state.students.length} roster entries…`
    : `Processing ${state.pdfFiles.length} files (no roster — template uses Fixed Template only)…`,
    'ok');

  try {
    state.results = [];

    for (let i = 0; i < state.pdfFiles.length; i++) {
      const pdf = state.pdfFiles[i];
      updateProgress(
        Math.round((i / state.pdfFiles.length) * 100),
        `Matching ${i + 1} of ${state.pdfFiles.length}: ${pdf.name}`
      );

      const filenameHint = extractNameFromFilename(pdf.name);
      const parsed = { learnerName: '', rawLearnerName: '', atsId: '', courseCode: '', criteriaType: '', criteriaStart: '', criteriaEnd: '' };

      // Extract file metadata automatically — fast, no OCR
      const fileMeta = {};
      try {
        const f = await pdf.handle.getFile();
        if (f.lastModified) {
          fileMeta.modified = new Date(f.lastModified).toISOString().slice(0, 10);
        }
        const pdfM = await extractPdfMetadata(pdf.handle);
        Object.assign(fileMeta, pdfM); // created, title, author, subject
      } catch (_) {}

      let match;
      if (hasRoster && filenameHint) {
        // Compare filename name against roster — no OCR needed
        match = matchStudent('', filenameHint, '', state.students);
      } else if (!filenameHint) {
        match = { status: 'unreadable', student: null, confidence: 0, candidates: state.students.slice(0, 5) };
      } else {
        // No roster: use filename directly (for Fixed Template-only setups)
        match = { status: 'extracted', student: null, confidence: 1, candidates: [] };
      }

      const newName = buildFilenameFromTemplate(parsed, match.student, filenameHint, pdf.name, i + 1, fileMeta);
      state.results.push({ pdf, parsed, filenameHint, fileMeta, match, newName, selected: !!match.student || match.status === 'extracted' });

      const conf = Math.round((match.confidence || 0) * 100);
      const statusText = match.status === 'matched'        ? `✓ matched (${conf}%)`
        : match.status === 'extracted'                     ? '→ no roster'
        : match.status === 'low-confidence'                ? `⚠ low confidence (${conf}%)`
        : '✗ needs review';
      addLog(`[${i + 1}/${state.pdfFiles.length}] ${pdf.name} → ${statusText}`,
        match.status === 'matched' ? 'ok' : match.status === 'low-confidence' ? 'warn' : 'err');
    }

    updateProgress(100, 'Matching complete!');
    addLog(`Done — ${state.results.filter(r => r.match.status === 'matched').length} matched, ` +
      `${state.results.filter(r => r.match.status === 'low-confidence').length} low-confidence, ` +
      `${state.results.filter(r => !['matched','low-confidence','extracted'].includes(r.match.status)).length} need review.`, 'ok');

    setTimeout(() => { buildResultsTable(); showStep(3); }, 300);

  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    addLog('Error: ' + msg, 'err');
    console.error('startProcessing error:', e);
  }
};

function updateProgress(pct, label) {
  if (pct !== null) {
    document.getElementById('prog-fill').style.width = pct + '%';
    document.getElementById('prog-pct').textContent = pct + '%';
  }
  if (label) {
    document.getElementById('prog-label').textContent = label;
    document.getElementById('prog-current').textContent = label;
  }
}

function addLog(msg, cls = '') {
  const log = document.getElementById('ocr-log');
  const line = document.createElement('div');
  line.className = 'ocr-log-line' + (cls ? ' ' + cls : '');
  line.textContent = msg;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}

// ─────────────────────────────────────────────────────────────
// PDF → CANVAS → OCR
// ─────────────────────────────────────────────────────────────
// File types that can be OCR'd as images
const IMAGE_EXTS = new Set(['jpg','jpeg','png','gif','bmp','webp','tiff','tif']);
const PDF_EXTS   = new Set(['pdf']);

// Greyscale + contrast-stretch preprocessing — helps OCR on grey-shaded cells
function preprocessCanvas(canvas) {
  const ctx = canvas.getContext('2d');
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const d   = img.data;
  for (let i = 0; i < d.length; i += 4) {
    const grey     = Math.round(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
    const enhanced = grey >= 220 ? 255 : Math.round(grey / 220 * 255);
    d[i] = d[i + 1] = d[i + 2] = enhanced;
  }
  ctx.putImageData(img, 0, 0);
}

// OCR a PDF — renders page 1 to canvas then recognises
async function ocrPdfPage(file) {
  await initPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const pdf  = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
  const page = await pdf.getPage(1);
  const viewport = page.getViewport({ scale: 4.0 });
  const canvas   = document.createElement('canvas');
  canvas.width   = viewport.width;
  canvas.height  = viewport.height;
  await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
  await pdf.destroy();
  preprocessCanvas(canvas);
  const { data: { text } } = await state.tesseractWorker.recognize(canvas);
  return text;
}

// OCR an image file — loads as <img>, draws on canvas, recognises
async function ocrImageFile(file) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = async () => {
      try {
        const scale  = 3.0;
        const canvas = document.createElement('canvas');
        canvas.width  = img.naturalWidth  * scale;
        canvas.height = img.naturalHeight * scale;
        canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);
        preprocessCanvas(canvas);
        const { data: { text } } = await state.tesseractWorker.recognize(canvas);
        resolve(text);
      } catch (e) { URL.revokeObjectURL(url); reject(e); }
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Image failed to load')); };
    img.src = url;
  });
}

// Router: choose OCR path based on file extension
async function ocrFirstPage(fileHandle) {
  const file = await fileHandle.getFile();
  const ext  = file.name.split('.').pop().toLowerCase();
  if (PDF_EXTS.has(ext))   return await ocrPdfPage(file);
  if (IMAGE_EXTS.has(ext)) return await ocrImageFile(file);
  return ''; // Word, Excel, PPT etc — no content OCR; rely on filename + roster
}

// ─────────────────────────────────────────────────────────────
// OCR TEXT PARSING
// ─────────────────────────────────────────────────────────────

// Strings that look like names but are actually form field labels — reject if candidate starts with these
const FORM_LABEL_RE = /^(course|ats|subject|unit|program|nqc|submission|assessment|milestone|outcome|performance|task|module|criteria|reference|date|signature|teacher|instructor|evaluator|mark|score|grade|result|total|pc\b|gc\b|pc\s*#|gc\s*#|description|ladder|circuit|relay|normally|closed|pushbutton|safety|identify|determine|label|describe|explain|interpret|document|using|apply|with|for|the|and|or|by|of|in|to)/i;
// Reject if the candidate is a sentence fragment (contains verb-like words)
const BAD_NAME_PREFIX = /^(course|assessment|subject|unit|program|task|module|outcome|criteria|milestone|pc|gc|describe|identify|determine|label|explain|interpret|document)/i;

/**
 * Scan OCR text for a label keyword and return the text that immediately
 * follows it — on the same line or the next line.
 * labelPattern: regex string for the label (e.g. "Learner\\s+Name")
 * Returns { value, raw } or null.
 */
function extractAfterLabel(text, labelPattern) {
  // Same-line: LABEL [: | - spaces]* VALUE [end of column or line]
  const re1 = new RegExp(
    `\\b${labelPattern}\\s*[:\\-|]*\\s*([A-Za-z][^\\n|\\t]{1,60}?)(?:\\s{2,}|\\s*[|]|\\s*\\n|$)`,
    'i'
  );
  // Next-line: LABEL\n VALUE
  const re2 = new RegExp(
    `\\b${labelPattern}\\s*[:\\-|]*\\s*\\n+\\s*([A-Za-z][^\\n|\\t]{1,60})`,
    'i'
  );
  for (const re of [re1, re2]) {
    const m = text.match(re);
    if (m) {
      // Strip trailing label bleed (e.g. "Aysha   Course Code" → "Aysha")
      const raw = m[1].trim()
        .replace(/\s+(Course|Subject|ATS|Unit|Program|NQC|Task|Assessment|Submission|Outcome).*$/i, '')
        .trim();
      if (raw.length >= 2) return { value: raw, raw: m[1].trim() };
    }
  }
  return null;
}

/** Extract a digit sequence that follows a label (for ID fields). */
function extractIdAfterLabel(text, labelPattern) {
  const re1 = new RegExp(
    `\\b${labelPattern}\\s*[:\\-|#.]*\\s*([A-Za-z0-9][^\\n|\\t]{1,40}?)(?:\\s{2,}|\\s*[|]|\\s*\\n|$)`,
    'i'
  );
  const re2 = new RegExp(
    `\\b${labelPattern}\\s*[:\\-|#.]*\\s*\\n+\\s*([A-Za-z0-9][^\\n]{1,40})`,
    'i'
  );
  for (const re of [re1, re2]) {
    const m = text.match(re);
    if (m) {
      const digits = m[1].replace(/\D/g, '');
      if (digits.length >= 4) return digits;
    }
  }
  return '';
}

function parseOcrText(text) {
  const result = {
    learnerName: '', rawLearnerName: '',
    atsId: '', courseCode: '',
    criteriaType: '', criteriaStart: '', criteriaEnd: '',
  };

  const lines = text.split('\n');

  // ── Helpers ──────────────────────────────────────────────────
  // Quality check for a name candidate
  function isValidName(s) {
    if (!s || s.length < 4) return false;
    const clean = s.trim().replace(/\s{2,}/g, ' ');
    const words = clean.split(/\s+/);
    return (
      /^[A-Za-z][A-Za-z\s.\-']{3,}$/.test(clean) &&  // only letters/spaces/dots
      words.length >= 2 && words.length <= 6 &&         // 2–6 words
      words.every(w => w.length >= 3) &&                // all words ≥ 3 chars
      words.some(w => w.length >= 5) &&                 // at least one ≥ 5 chars
      !FORM_LABEL_RE.test(clean)                        // not a form label
    );
  }

  // Find the line that contains a label keyword.
  // Uses three tiers — stops at the first tier that finds a match:
  //   Tier 1 — exact word match (e.g. "Learner Name")
  //   Tier 2 — partial/fuzzy match (tolerates OCR character swaps)
  //   Tier 3 — fragment match (just a distinctive substring)
  function findLabelLine(tiers) {
    for (const tier of tiers) {
      for (const pat of tier) {
        // With word boundaries first
        const reWb  = new RegExp(`\\b${pat}\\b`, 'i');
        const idx1  = lines.findIndex(l => reWb.test(l));
        if (idx1 >= 0) return { idx: idx1, re: reWb };
        // Without word boundaries (catches run-together OCR like "LeamerName")
        const reNwb = new RegExp(pat, 'i');
        const idx2  = lines.findIndex(l => reNwb.test(l));
        if (idx2 >= 0) return { idx: idx2, re: reNwb };
      }
    }
    return null;
  }

  // Extract a value adjacent to the label:
  //   1. Same line — text after the label, right of a pipe or separator
  //   2. Next 1–3 lines — look for value between pipes or before a trailing label
  function extractValueNearLabel(labelResult) {
    if (!labelResult) return null;
    const { idx, re } = labelResult;

    // ── Same-line extraction ──
    const labelLine = lines[idx];
    const matchPos  = labelLine.search(re);
    if (matchPos >= 0) {
      const afterLabel = labelLine.slice(matchPos + labelLine.match(re)[0].length);
      // Pattern: "| VALUE |" or "| VALUE Course Code" or ": VALUE"
      for (const pattern of [
        /\|\s*([A-Za-z][^|\n]{3,60}?)\s*(?:\||\bCourse|\bSubject|$)/i,
        /[:\-]\s*([A-Za-z][^|\n]{3,60}?)\s*(?:\||\bCourse|\bSubject|$)/i,
        /\s{2,}([A-Za-z][^|\n]{3,50}?)\s*(?:\||\bCourse|\bSubject|$)/i,
      ]) {
        const m = afterLabel.match(pattern);
        if (m) {
          const val = m[1].trim();
          if (isValidName(val)) return val;
        }
      }
    }

    // ── Next-line extraction ──
    // The form table puts the value row BELOW the label row.
    // OCR reads: Label row → "Learner Name  Course Code | ENE61"
    //            Value row → "[garbled] | VALUE  Course Code |"
    for (let j = idx + 1; j <= Math.min(idx + 3, lines.length - 1); j++) {
      const line = lines[j].trim();
      if (!line) continue;

      // Value between | and a trailing label keyword
      const m1 = line.match(/\|\s*([A-Za-z][^|\n]{3,60}?)\s+(?:Course\s*Code|Subject|ATS|Submission)/i);
      if (m1) { const v = m1[1].trim(); if (isValidName(v)) return v; }

      // Value between two pipes
      const m2 = line.match(/\|\s*([A-Za-z][^|\n]{3,50}?)\s*\|/i);
      if (m2) { const v = m2[1].trim(); if (isValidName(v)) return v; }

      // Any pipe-separated segment that looks like a name
      const segs = line.split('|').map(s => s.trim());
      for (const seg of segs) {
        if (isValidName(seg)) return seg;
      }
    }
    return null;
  }

  // ── NAME: label-first with 3-tier fuzzy keyword search ──────
  // Tier 1 — exact standard labels (preferred)
  // Tier 2 — OCR-error tolerant: common character substitutions in these words
  //   OCR often swaps: n↔m, r↔n, e↔a, i↔l, N↔H, d↔cl, etc.
  //   "Learner" → "Leamer", "Learner", "Leaner", "Lea rner"
  //   "Name"    → "Hame", "Nam", "Hame", "Mame", "Nam e"
  // Tier 3 — fragment fallback: just distinctive substrings
  const NAME_LABEL_TIERS = [
    // Tier 1: exact standard labels
    [
      'Learner\\s+Name', 'Student\\s+Name', 'Staff\\s+Name',
      'Employee\\s+Name', 'Candidate\\s+Name', 'Trainee\\s+Name', 'Full\\s+Name',
    ],
    // Tier 2: tolerant of 1-2 character OCR errors in the label words
    [
      'Le[ae]?r?n[aeo]r?\\s+[NH]?[Nn]am[eo]?',  // Leaner/Learner + Name/Hame/Nam
      'L[ea]{1,2}r?n\\w{0,3}\\s+[Nn]am',          // L(ea)rn... + Nam
      'earn\\w*\\s+[Nn]am',                         // ...earn... + Nam
      'learn\\w*\\s+[Nn]am',                        // learn... + Nam
      'Student\\s+[Nn]am', 'Staff\\s+[Nn]am',       // other labels with fuzzy Nam
    ],
    // Tier 3: just the fragment — "earn" is unique enough on a form
    [
      'earn\\w{0,3}\\s+[Nn]a',  // "earner Na" from "Learner Name"
      'arner\\s',                 // "arner " (end of Learner)
      'rner\\s+N',               // "rner N" (Learner N...)
    ],
  ];

  const nameLabelResult = findLabelLine(NAME_LABEL_TIERS);
  if (nameLabelResult) {
    const val = extractValueNearLabel(nameLabelResult);
    if (val) {
      result.rawLearnerName = val;
      result.learnerName    = val;
    }
  }

  // Fallback: structural pattern — value precedes "Course Code" after a pipe
  // Used when the "Learner Name" label cell was too grey for OCR to read.
  if (!result.learnerName) {
    for (const m of text.matchAll(/\|\s*([A-Za-z][^|\n]{3,60}?)\s+(?=Course\s*Code)/gi)) {
      const candidate = m[1].trim().replace(/\s{2,}/g, ' ');
      if (isValidName(candidate)) {
        result.rawLearnerName = candidate;
        result.learnerName    = candidate;
        break;
      }
    }
  }

  if (!result.rawLearnerName) {
    result.rawLearnerName = nameLabelResult
      ? `(label found on line ${nameLabelResult.idx + 1} but value was unreadable)`
      : '("Learner Name" label not found in OCR — label cell may have been unreadable)';
  }

  // ── ID: label-first ──────────────────────────────────────────
  // "ATS ID", "Student ID" etc. are the markers for the ID field.
  // Find the label, then look right (same line) or below (next row).
  // Tier 1: exact ATS/Student/Staff ID labels
  // Tier 2: OCR tolerant — "ATS ID" often read as "ASI", "ATS lD", "ATslD",
  //   "A1S", "ATG ID", "AT5 ID" (digit/letter swaps)
  // Tier 3: fragments — just the "ATS" prefix or standalone "ID" near digits
  const ID_LABEL_TIERS = [
    // Tier 1: exact labels
    [
      'ATS\\s+ID', 'Student\\s+ID', 'Staff\\s+ID',
      'Employee\\s+ID', 'Learner\\s+ID', 'Trainee\\s+ID', 'Emirates\\s+ID',
    ],
    // Tier 2: 1-2 char OCR errors
    [
      'A[TtGg1][SsI5]\\s*[IlL1][Dd]',  // ATS ID → ATSlD, A1S ID, ATG ID
      'A[TtGg][SsI]\\s+[IilL1Dd]',      // ATS I, ATG I...
      'ASI', 'AST\\s*ID', 'ATS\\s*lD', 'ATS\\s*1D',  // common single-word misreads
      'A\\.T\\.S', 'A\\.T\\.S\\.',       // "A.T.S." style
      'Student\\s+[IilL1]', 'Staff\\s+[IilL1]',  // fuzzy ID
    ],
    // Tier 3: just the prefix
    [
      '\\bATS\\b', '\\bASI\\b', '\\bATSID\\b',  // bare ATS
      'Student\\s+No', 'Staff\\s+No', 'ID\\s+No', 'ID\\s+Number',
    ],
  ];

  function extractIdNearLabel(labelIdx) {
    if (labelIdx < 0) return '';
    // Check same line after the label
    const labelLine = lines[labelIdx];
    const afterPipe = labelLine.match(/\|\s*([A-Za-z0-9][^|\n]{1,30})/);
    if (afterPipe) {
      const digits = afterPipe[1].replace(/\D/g, '');
      if (digits.length >= 4) return digits;
    }
    // Check next 1–3 lines
    for (let j = labelIdx + 1; j <= Math.min(labelIdx + 3, lines.length - 1); j++) {
      const line = lines[j].trim();
      if (!line) continue;
      // Value between pipe and next pipe/end
      const m = line.match(/\|\s*([A-Za-z0-9][^|\n]{1,30}?)(?:\s*\||$)/i);
      if (m) {
        const digits = m[1].replace(/\D/g, '');
        if (digits.length >= 4) return digits;
      }
      // Any long digit sequence on this line
      const digits = line.replace(/\D/g, '');
      if (digits.length >= 5) return digits;
    }
    return '';
  }

  const idLabelResult = findLabelLine(ID_LABEL_TIERS);
  if (idLabelResult) {
    result.atsId = extractIdNearLabel(idLabelResult.idx);
  }

  // Fallback: any long digit sequence near the name section
  if (!result.atsId) {
    const nameLineIdx = nameLabelResult?.idx ?? 0;
    const scanEnd     = Math.min(nameLineIdx + 8, lines.length);
    for (let j = nameLineIdx; j < scanEnd; j++) {
      const digits = lines[j].replace(/\D/g, '');
      if (digits.length >= 6) { result.atsId = digits; break; }
    }
  }

  // ── Course Code ── e.g. ENE61, EEE61, ETE61
  result.courseCode = (() => {
    // 1. "Course Code" label — tolerates any separator and optional space between letters and digits
    let m = text.match(/[Cc]ourse\s*[Cc]ode[\s|:\-]*([A-Za-z]{2,4})\s*(\d{1,4})/i);
    if (m) return (m[1] + m[2]).toUpperCase();

    // 2. Just "Code" label (handles OCR garbling "Course")
    m = text.match(/\bcode[\s|:\-]+([A-Za-z]{2,4})\s*(\d{1,4})/i);
    if (m) return (m[1] + m[2]).toUpperCase();

    // 3. Scan all letter+digit codes, exclude GC/PC/PS/IT/IS/AS/AD/ID patterns
    const excluded = /^(GC|PC|PS|IT|IS|AS|AD|ID|NO|PO)/i;
    const allCodes = [...text.matchAll(/\b([A-Za-z]{2,4})\s*(\d{2,3})\b/g)]
      .map(x => (x[1] + x[2]).toUpperCase())
      .filter(c => !excluded.test(c));
    if (allCodes.length > 0) return allCodes[0];

    return '';
  })();

  // ── Criteria type and range ── e.g. "GC 5.1-5.3" or "PC 2.1-2.4"
  const rangeMatch = text.match(/\b(GC|PC)\s*(\d+\.\d+)\s*[-–]\s*(\d+\.\d+)/i);
  if (rangeMatch) {
    result.criteriaType  = rangeMatch[1].toUpperCase();
    result.criteriaStart = rangeMatch[2];
    result.criteriaEnd   = rangeMatch[3];
  } else {
    // Single criterion or range written without dash
    const singleMatch = text.match(/\b(GC|PC)\s*(\d+\.\d+)/i);
    if (singleMatch) {
      result.criteriaType  = singleMatch[1].toUpperCase();
      result.criteriaStart = singleMatch[2];
      result.criteriaEnd   = singleMatch[2];
      // Try to find the last criterion number in the criteria table
      const allMatches = [...text.matchAll(/\b(?:GC|PC)\s*(\d+\.\d+)/gi)];
      if (allMatches.length > 1) {
        result.criteriaEnd = allMatches[allMatches.length - 1][1];
      }
    }
  }

  return result;
}

// ─────────────────────────────────────────────────────────────
// FILENAME HELPERS
// ─────────────────────────────────────────────────────────────
function extractNameFromFilename(filename) {
  // Strip any extension, then extract name after " - " separator
  const m = filename.replace(/\.[^.]+$/, '').match(/[-–]\s+(.+)$/);
  return m ? m[1].trim().replace(/\s+/g, ' ') : '';
}

function fileExtension(filename) {
  const m = (filename || '').match(/(\.[^.]+)$/);
  return m ? m[1].toLowerCase() : '';
}

// ─────────────────────────────────────────────────────────────
// SMART SCAN — discover content fields from a sample file
// ─────────────────────────────────────────────────────────────

// Extract PDF metadata from file headers (fast, no OCR needed)
async function extractPdfMetadata(fileHandle) {
  try {
    await initPdfJs();
    const file   = await fileHandle.getFile();
    const buf    = await file.arrayBuffer();
    const pdf    = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
    const meta   = await pdf.getMetadata();
    await pdf.destroy();
    const info   = meta?.info || {};
    const result = {};
    if (info.Title)       result.title    = info.Title;
    if (info.Author)      result.author   = info.Author;
    if (info.Subject)     result.subject  = info.Subject;
    if (info.Keywords)    result.keywords = info.Keywords;
    if (info.CreationDate) {
      // PDF dates are like "D:20260315120000+04'00'"
      const m = String(info.CreationDate).match(/D:(\d{4})(\d{2})(\d{2})/);
      if (m) result.created = `${m[1]}-${m[2]}-${m[3]}`;
    }
    return result;
  } catch (_) { return {}; }
}

// Orchestrate: PDF metadata + OCR first page → discovered fields
async function scanSampleFile(fileHandle) {
  const pdfMeta = await extractPdfMetadata(fileHandle);
  let   parsed  = {};
  try {
    // Reuse existing OCR pipeline (already in codebase)
    if (!state.tesseractWorker) {
      // Init Tesseract if not already running
      const workerPath = chrome.runtime.getURL('lib/worker.min.js');
      const corePath   = chrome.runtime.getURL('lib/');
      const langPath   = chrome.runtime.getURL('lib');
      const _Orig = window.Worker;
      window.Worker = (url, opts) =>
        new _Orig((typeof url === 'string' && url.startsWith('blob:')) ? workerPath : url, opts);
      window.Worker.prototype = _Orig.prototype;
      try {
        state.tesseractWorker = await Tesseract.createWorker('eng', Tesseract.OEM?.LSTM_ONLY ?? 1,
          { workerPath, corePath, langPath });
      } finally { window.Worker = _Orig; }
    }
    const ocrText = await ocrFirstPage(fileHandle);
    parsed = parseOcrText(ocrText);
  } catch (_) {}

  return { pdfMeta, parsed };
}

// Build the discovered-fields panel rows and show them
function showScanResults(fileName, results) {
  const { pdfMeta, parsed } = results;
  const container = document.getElementById('scan-fields');
  const nameEl    = document.getElementById('scan-filename');
  if (!container) return;
  nameEl.textContent = fileName;
  container.innerHTML = '';

  const rowStyle = `display:flex;align-items:center;gap:10px;padding:8px 14px;
    border-bottom:1px solid var(--gray-100);font-size:12px;`;
  const labelStyle = `color:var(--gray-400);min-width:90px;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;`;
  const valStyle   = `flex:1;color:var(--gray-800);font-weight:500;`;
  const btnStyle   = (col) => `padding:3px 9px;border:none;border-radius:5px;font-size:11px;
    font-weight:600;cursor:pointer;background:${col};color:white;white-space:nowrap;`;

  // Helper — create one field row
  function addRow(section, label, value, tokenKey, mapTo) {
    if (!value || value.length < 2) return;
    const row = document.createElement('div');
    row.style.cssText = rowStyle;
    if (section) {
      // Section header
      const hdr = document.createElement('div');
      hdr.style.cssText = `padding:6px 14px 3px;font-size:10px;font-weight:700;
        text-transform:uppercase;letter-spacing:0.07em;color:var(--gray-400);
        background:var(--gray-50);border-bottom:1px solid var(--gray-100);`;
      hdr.textContent = section;
      container.appendChild(hdr);
    }
    row.innerHTML = `
      <span style="${labelStyle}">${escHtml(label)}</span>
      <span style="${valStyle}" title="${escHtml(value)}">${escHtml(value.slice(0, 60))}${value.length > 60 ? '…' : ''}</span>
    `;
    // Insert token button
    const insertBtn = document.createElement('button');
    insertBtn.textContent = `Insert {${tokenKey}}`;
    insertBtn.style.cssText = btnStyle('#3b82f6');
    insertBtn.onclick = () => insertDiscoveredToken(tokenKey, value, value.slice(0,50));
    row.appendChild(insertBtn);
    // Map to {name} or {id} button
    if (mapTo) {
      const mapBtn = document.createElement('button');
      mapBtn.textContent = `→ {${mapTo}}`;
      mapBtn.style.cssText = btnStyle('#22c55e') + 'margin-left:4px;';
      mapBtn.onclick = () => {
        state.customTokens[mapTo] = value;
        mapBtn.textContent = `✓ {${mapTo}} set`;
        mapBtn.disabled = true;
      };
      row.appendChild(mapBtn);
    }
    container.appendChild(row);
  }

  // PDF metadata is now extracted automatically per-file — show it here as info only
  const metaNote = document.createElement('div');
  metaNote.style.cssText = 'padding:8px 14px;font-size:11px;color:var(--gray-400);background:var(--gray-50);border-bottom:1px solid var(--gray-100);';
  const metaItems = Object.entries(pdfMeta).filter(([,v]) => v);
  metaNote.textContent = metaItems.length
    ? `ℹ️ PDF metadata auto-extracted: ${metaItems.map(([k,v])=>`{${k}}=${v}`).join(' · ')} — use {created}, {title} etc. in your pattern without scanning.`
    : 'ℹ️ No PDF metadata found in this file. Metadata tokens ({created}, {title}) will be empty for this file.';
  container.appendChild(metaNote);

  // OCR-extracted text fields
  const ocrFields = [
    ['Learner Name', parsed.learnerName, 'learner-name', 'name'],
    ['ATS ID',       parsed.atsId,       'ats-id',       'id'],
    ['Course Code',  parsed.courseCode,   'course-code',  null],
  ];
  let hasOcr = false;
  let ocrFirst = true;
  for (const [label, val, tkey, mapTo] of ocrFields) {
    if (!val) continue;
    addRow(ocrFirst ? '📝 Text found on first page' : null, label, val, tkey, mapTo);
    ocrFirst = false;
    hasOcr = true;
  }

  if (!hasOcr) {
    const msg = document.createElement('div');
    msg.style.cssText = 'padding:12px 14px;color:var(--gray-400);font-size:12px;';
    msg.textContent = 'No readable text fields found on the first page.';
    container.appendChild(msg);
  }

  document.getElementById('scan-results').classList.remove('hidden');
}

// Insert a discovered token into the pattern and add it to the dropdown
function insertDiscoveredToken(tokenKey, value, label) {
  state.customTokens[tokenKey] = value;

  // Insert token in pattern input
  const input = document.getElementById('template-pattern');
  if (input) {
    const pos = input.selectionStart ?? input.value.length;
    const tok = `{${tokenKey}}`;
    input.value = input.value.slice(0, pos) + tok + input.value.slice(pos);
    state.template.pattern = input.value;
    input.focus();
    updatePreview();
    updateExcelZoneBadge();
    updateStartBtn();
    saveTemplate();
  }

  // Add to dropdown if not already there
  const sel = document.getElementById('token-insert-select');
  if (sel && !sel.querySelector(`option[value="{${tokenKey}}"]`)) {
    // Find or create the "Discovered" optgroup
    let grp = sel.querySelector('optgroup[label="Discovered from your files"]');
    if (!grp) {
      grp = document.createElement('optgroup');
      grp.label = 'Discovered from your files';
      sel.appendChild(grp);
    }
    const opt = document.createElement('option');
    opt.value = `{${tokenKey}}`;
    opt.textContent = `{${tokenKey}} — ${label}`;
    grp.appendChild(opt);
  }
}

function buildFilenameFromTemplate(parsed, student, filenameHint = '', originalName = '', count = 0, fileMeta = {}) {
  const hasRoster  = state.students.length > 0;
  const extWithDot = fileExtension(originalName) || '.pdf';
  const ext        = extWithDot.slice(1);
  const today      = new Date().toISOString().slice(0, 10);

  const values = {
    name:     hasRoster ? (student?.name  || 'UNKNOWN') : (parsed?.learnerName || filenameHint || 'UNKNOWN'),
    id:       hasRoster ? (student?.atsId || 'UNKNOWN') : (parsed?.atsId || 'UNKNOWN'),
    ext,
    filetype: ext,
    date:     today,
    year:     today.slice(0, 4),
    month:    today.slice(5, 7),
    day:      today.slice(8, 10),
    filename: (originalName || '').replace(/\.[^.]+$/, ''),
    count:    String(count).padStart(3, '0'),
    index:    String(count).padStart(3, '0'),
    // Per-file metadata — populated automatically, no scan needed
    created:  fileMeta.created  || '',
    modified: fileMeta.modified || '',
    title:    fileMeta.title    || '',
    author:   fileMeta.author   || '',
    subject:  fileMeta.subject  || '',
  };

  // Merge with OCR-discovered custom tokens (from Scan button)
  const allValues = { ...values, ...state.customTokens };

  let result = (state.template.pattern || '').trim();
  for (const [token, value] of Object.entries(allValues)) {
    result = result.replace(new RegExp(`\\{${token}\\}`, 'gi'), String(value));
  }
  return result || 'UNKNOWN' + extWithDot;
}

// ─────────────────────────────────────────────────────────────
// MATCHING ALGORITHM
// ─────────────────────────────────────────────────────────────
function normalise(str) {
  return (str || '')
    .toLowerCase()
    .replace(/\b(al|el)\s+(?=[a-z])/g, '$1') // "Al Ktebi" → "Alktebi", "Al Rashdi" → "Alrashdi"
    .replace(/[^a-z\s]/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function tokenize(str) {
  return normalise(str).split(' ').filter(t => t.length > 1);
}

function jaccardTokens(a, b) {
  if (!a || !b) return 0;
  const sa = new Set(tokenize(a));
  const sb = new Set(tokenize(b));
  if (sa.size === 0 || sb.size === 0) return 0;
  const inter = [...sa].filter(t => sb.has(t)).length;
  const union  = new Set([...sa, ...sb]).size;
  return inter / union;
}

// Character trigram Jaccard — tolerates OCR misreadings
function trigramJaccard(a, b) {
  if (!a || !b) return 0;
  const norm = s => s.toLowerCase().replace(/[^a-z]/g, '');
  const tg   = s => {
    const n = norm(s);
    const g = new Set();
    for (let i = 0; i <= n.length - 3; i++) g.add(n.slice(i, i + 3));
    return g;
  };
  const ga = tg(a);
  const gb = tg(b);
  if (ga.size === 0 || gb.size === 0) return 0;
  const inter = [...ga].filter(g => gb.has(g)).length;
  const union  = new Set([...ga, ...gb]).size;
  return inter / union;
}

// Extract first name token
function getFirstName(nameStr) {
  const tokens = tokenize(nameStr);
  return tokens.length > 0 ? tokens[0] : '';
}

// Extract second name token (middle name)
function getSecondName(nameStr) {
  const tokens = tokenize(nameStr);
  return tokens.length > 1 ? tokens[1] : '';
}

// Extract last name — merges Arabic "Al/El" prefix with the following word
// e.g. "Rauda Hamad Al Ktebi" → "alktebi"  (not just "ktebi")
function getLastName(nameStr) {
  const tokens = tokenize(nameStr); // all tokens, no length filter
  if (tokens.length === 0) return '';
  const last = tokens[tokens.length - 1];
  if (tokens.length >= 2) {
    const prev = tokens[tokens.length - 2];
    if (/^(al|el)$/i.test(prev)) {
      return prev + last; // "al" + "ktebi" → "alktebi"
    }
  }
  // Skip very short tokens when looking for the family name
  const long = tokens.filter(t => t.length >= 4);
  return long.length > 0 ? long[long.length - 1] : last;
}

// Consonant skeleton — strips vowels so Arabic transliteration variants match
// "Alketbi" → "lktb",  "Alktebi" → "lktb"  ← identical ✓
// "Rouda"   → "rd",    "Rauda"   → "rd"     ← identical ✓
function consonantSkeleton(str) {
  return (str || '').toLowerCase().replace(/[aeiou\s]/g, '');
}

// Last-name similarity: consonant skeleton first, trigram as fallback
function lastNameSimilarity(a, b) {
  if (!a || !b) return 0;
  const ca = consonantSkeleton(a);
  const cb = consonantSkeleton(b);
  if (ca && cb && ca === cb) return 1.0;           // exact consonant match
  const skelScore = trigramJaccard(ca, cb);        // trigram on skeletons
  const rawScore  = trigramJaccard(a, b);          // trigram on full strings
  return Math.max(skelScore, rawScore);
}

// Digit sequence overlap — for ATS ID tiebreaker
function digitOverlap(ocrDigits, excelId) {
  if (!ocrDigits || ocrDigits.length < 2) return 0;
  const ed = (excelId || '').replace(/\D/g, '');
  if (!ed) return 0;
  let hits = 0, total = 0;
  for (let len = 2; len <= Math.min(ocrDigits.length, 5); len++) {
    for (let i = 0; i <= ocrDigits.length - len; i++) {
      const sub = ocrDigits.slice(i, i + len);
      if (ed.includes(sub)) hits++;
      total++;
    }
  }
  return total ? hits / total : 0;
}

function scoreAgainstExcel(nameStr, students) {
  const fnTokens    = tokenize(nameStr);         // filename tokens e.g. ["wadima","matar"]
  const fnTokenSet  = new Set(fnTokens);
  const ocrLastName = getLastName(nameStr);

  return students.map(s => {
    const excelTokens   = new Set(tokenize(s.name));
    const jaccard       = jaccardTokens(nameStr, s.name);
    const trigram       = trigramJaccard(nameStr, s.name);
    const excelLastName = getLastName(s.name);

    // ── Subset recall bonus ──────────────────────────────────────
    // "Wadima Matar" vs "Wadima Matar Salem Mubarak Alkhateri":
    // All filename tokens exist in the roster name → 100% recall.
    // Jaccard alone gives only 40% (2/5) which unfairly punishes short filenames.
    // When all filename words appear in the roster name, the recall is 1.0
    // and we use that as a strong signal (score ≥ 0.85).
    const matchedTokens = fnTokens.filter(t => excelTokens.has(t)).length;
    const recall        = fnTokens.length > 0 ? matchedTokens / fnTokens.length : 0;

    // Last-name similarity (consonant skeleton)
    const lastNameScore = ocrLastName.length >= 3
      ? lastNameSimilarity(ocrLastName, excelLastName)
      : 0;

    // Base score: max of (jaccard+trigram weighted) and (recall-based)
    // Recall-based score is capped at 0.88 to leave room for a better exact match
    const weightedScore = jaccard * 0.5 + trigram * 0.3 + lastNameScore * 0.2;
    const recallScore   = recall >= 1.0 ? 0.88        // all words matched → very high
                        : recall >= 0.5 ? recall * 0.7  // partial → moderate
                        : 0;
    const score = Math.max(weightedScore, recallScore);

    return { student: s, score, jaccard, trigram, lastNameScore, recall };
  }).sort((a, b) => b.score - a.score);
}

function matchStudent(ocrName, filenameHint, ocrId, students) {

  // ── Stage 0: First name + Last name direct check ──────────────────────────
  // Use BOTH the OCR name and the filename hint as name sources.
  // If first name AND last name both match a student clearly:
  //   → Check ATS ID: if match → high confidence (matched)
  //   → If ID doesn't match → check second/middle name to choose between candidates
  const nameSources = [ocrName, filenameHint].filter(n => n && n.length > 2);

  for (const src of nameSources) {
    const srcFirst = getFirstName(src);
    const srcLast  = getLastName(src);
    if (!srcFirst || !srcLast || srcFirst === srcLast) continue; // need both

    const firstLastMatches = students.filter(s => {
      const exFirst = getFirstName(s.name);
      const exLast  = getLastName(s.name);
      // Use consonant skeletons for both first and last names
      const firstOk = trigramJaccard(srcFirst, exFirst) >= 0.45
                   || consonantSkeleton(srcFirst) === consonantSkeleton(exFirst);
      const lastOk  = lastNameSimilarity(srcLast, exLast) >= 0.5;
      return firstOk && lastOk;
    });

    if (firstLastMatches.length === 1) {
      // Unique first+last match — check ATS ID for confidence level
      const student = firstLastMatches[0];
      const idScore = digitOverlap(ocrId, student.atsId);
      const conf    = idScore >= 0.3 ? 0.92 : 0.72;
      return { status: 'matched', student, confidence: conf, candidates: [] };
    }

    if (firstLastMatches.length > 1) {
      // Multiple students share first+last name pattern — use ATS ID to pick
      const withId = firstLastMatches.map(s => ({
        student: s,
        score: 0.7,
        idScore: digitOverlap(ocrId, s.atsId),
      })).sort((a, b) => b.idScore - a.idScore);

      const top  = withId[0];
      const next = withId[1];

      if (top.idScore > 0.25 && top.idScore - next.idScore > 0.15) {
        return { status: 'matched', student: top.student, confidence: 0.85, candidates: withId };
      }

      // ID didn't resolve it — check second (middle) name
      const srcSecond = getSecondName(src);
      if (srcSecond) {
        const withSecond = firstLastMatches.map(s => ({
          student: s,
          score: 0.7,
          secondScore: trigramJaccard(srcSecond, getSecondName(s.name)),
        })).sort((a, b) => b.secondScore - a.secondScore);

        const s0 = withSecond[0], s1 = withSecond[1];
        if (s0.secondScore >= 0.5 && s0.secondScore - s1.secondScore > 0.2) {
          return { status: 'matched', student: s0.student, confidence: 0.80, candidates: withSecond };
        }
      }

      // Still ambiguous — flag for manual
      return {
        status: 'manual',
        student: null,
        confidence: 0.65,
        candidates: firstLastMatches.map(s => ({ student: s, score: 0.65 })),
        bestGuess: firstLastMatches[0],
      };
    }
  }

  // ── Score using OCR name ──
  const ocrScored = ocrName.length > 2 ? scoreAgainstExcel(ocrName, students) : [];

  // ── Score using filename hint ──
  const fnScored  = filenameHint.length > 2 ? scoreAgainstExcel(filenameHint, students) : [];

  // ── Merge: best score for each student from either source ──
  const merged = students.map(s => {
    const o = ocrScored.find(r => r.student === s);
    const f = fnScored.find(r => r.student === s);
    const ocrScore = o?.score || 0;
    const fnScore  = f?.score || 0;
    const score    = Math.max(ocrScore, fnScore);
    return { student: s, score, ocrScore, fnScore };
  }).sort((a, b) => b.score - a.score);

  const top    = merged[0];
  const second = merged[1];
  const gap    = second ? top.score - second.score : 1;

  // ── Filename-vs-roster matching at 50% threshold ────────────
  // OCR is no longer used. Matching is filename name vs roster only.
  // User-confirmed threshold: confidence ≥ 50% = matched.

  // Clear winner: ≥50% with meaningful gap over next candidate
  if (top.score >= 0.5 && gap >= 0.15) {
    return { status: 'matched', student: top.student, confidence: top.score, candidates: [] };
  }

  // Good match but gap is narrow (two similar names in roster)
  if (top.score >= 0.5) {
    return { status: 'low-confidence', student: top.student, confidence: top.score, candidates: merged.slice(0, 3) };
  }

  // Below threshold but not zero — flag for manual review
  if (top.score >= 0.2) {
    return { status: 'low-confidence', student: top.student, confidence: top.score, candidates: merged.slice(0, 5) };
  }

  // ── Too low to work with ──
  if (top.score < 0.2) {
    return { status: 'unreadable', student: null, confidence: 0, candidates: merged.slice(0, 3) };
  }

  // ── Stage 2: Narrowing via more tokens (already in merged score) — check top candidates ──
  const candidates = merged.filter(r => r.score >= 0.2);

  if (candidates.length === 1) {
    const conf = candidates[0].score;
    const status = conf >= 0.5 ? 'matched' : 'low-confidence';
    return { status, student: candidates[0].student, confidence: conf, candidates };
  }

  // ── Stage 2b: Last name tiebreaker ──
  // When two candidates are close in score, compare the last name from
  // both the OCR text and the filename hint against each candidate's family name.
  if (candidates.length > 1) {
    const ocrLastName = getLastName(ocrName);
    const fnLastName  = getLastName(filenameHint);

    const withLastName = candidates.map(c => {
      const excelLast    = getLastName(c.student.name);
      const ocrLnScore   = ocrLastName.length >= 3 ? lastNameSimilarity(ocrLastName, excelLast) : 0;
      const fnLnScore    = fnLastName.length  >= 3 ? lastNameSimilarity(fnLastName,  excelLast) : 0;
      const lnScore      = Math.max(ocrLnScore, fnLnScore);
      return { ...c, lnScore, combinedScore: c.score * 0.7 + lnScore * 0.3 };
    }).sort((a, b) => b.combinedScore - a.combinedScore);

    const lnTop  = withLastName[0];
    const lnNext = withLastName[1];
    const lnGap  = lnNext ? lnTop.combinedScore - lnNext.combinedScore : 1;

    if (lnTop.lnScore >= 0.4 && lnGap >= 0.15) {
      const conf = lnTop.combinedScore;
      return {
        status: conf >= 0.45 ? 'matched' : 'low-confidence',
        student: lnTop.student,
        confidence: conf,
        candidates: withLastName,
      };
    }
  }

  // ── Stage 3: ATS ID digit tiebreaker ──
  if (ocrId && ocrId.length >= 3) {
    const withId = candidates.map(c => ({
      ...c,
      idScore: digitOverlap(ocrId, c.student.atsId),
    })).sort((a, b) => {
      const diff = b.idScore - a.idScore;
      return diff !== 0 ? diff : b.score - a.score;
    });

    const idTop  = withId[0];
    const idNext = withId[1];
    const idGap  = idNext ? idTop.idScore - idNext.idScore : 1;

    if (idTop.idScore > 0.25 && idGap >= 0.15) {
      const conf = (idTop.score + idTop.idScore) / 2;
      return { status: conf >= 0.4 ? 'matched' : 'low-confidence', student: idTop.student, confidence: conf, candidates: withId };
    }
  }

  // ── Stage 4: Manual assignment needed ──
  const bestConf = top.score;
  return {
    status: 'manual',
    student: null,
    confidence: bestConf,
    candidates: candidates.slice(0, 5),
    bestGuess: top.student,
  };
}

// ─────────────────────────────────────────────────────────────
// STEP 3 — RESULTS TABLE
// ─────────────────────────────────────────────────────────────
function buildResultsTable() {
  const tbody = document.getElementById('results-tbody');
  tbody.innerHTML = '';

  const noRef = state.students.length === 0;

  // ── Populate template banner ──
  const templateDisplay = document.getElementById('review-template-display');
  if (templateDisplay) {
    templateDisplay.textContent = state.template.pattern || '— no template set —';
  }

  // Show UNKNOWN hint if any filenames contain UNKNOWN and no reference file
  const hasUnknown = state.results.some(r => r.newName && r.newName.includes('UNKNOWN'));
  const unknownHint = document.getElementById('review-unknown-hint');
  if (unknownHint) {
    unknownHint.classList.toggle('hidden', !hasUnknown);
  }

  // Update table column headers for no-reference mode
  const thead = document.querySelector('#results-table thead tr');
  if (thead) {
    const ths = thead.querySelectorAll('th');
    if (ths[3]) ths[3].textContent = noRef ? 'Extracted Name' : 'Matched Student';
    if (ths[4]) ths[4].textContent = noRef ? 'Extracted ID' : 'ATS ID';
  }

  let nMatched = 0, nExtracted = 0, nLow = 0, nManual = 0, nUnread = 0;
  state.results.forEach(r => {
    const s = r.match.status;
    if (s === 'matched')         nMatched++;
    else if (s === 'extracted')  nExtracted++;
    else if (s === 'low-confidence') nLow++;
    else nManual++;
    if (s === 'unreadable' || s === 'error') nUnread++;
  });

  const pills = document.getElementById('summary-pills');
  pills.innerHTML = `
    ${nMatched > 0    ? `<span class="pill pill-green">✅ ${nMatched} matched</span>` : ''}
    ${nExtracted > 0  ? `<span class="pill pill-blue">→ ${nExtracted} extracted</span>` : ''}
    ${nLow > 0        ? `<span class="pill pill-yellow">⚠️ ${nLow} low-confidence</span>` : ''}
    ${(nManual - nUnread) > 0 ? `<span class="pill pill-red">⚠️ ${nManual - nUnread} need manual</span>` : ''}
    ${nUnread > 0     ? `<span class="pill pill-red">❌ ${nUnread} unreadable</span>` : ''}
  `;

  const nSelected = state.results.filter(r => r.selected).length;
  document.getElementById('apply-label').textContent = `Apply Renaming (${nSelected} files)`;

  state.results.forEach((r, idx) => {
    const tr = document.createElement('tr');
    const st = r.match.status;
    tr.className = (st === 'matched' || st === 'extracted') ? 'row-matched'
      : st === 'low-confidence' ? 'row-low'
      : 'row-manual';

    const conf    = r.match.confidence || 0;
    const confPct = Math.round(conf * 100);
    const confColor = conf >= 0.6 ? '#22c55e' : conf >= 0.4 ? '#f59e0b' : '#ef4444';

    const badgeHtml = st === 'matched'
      ? '<span class="badge badge-green">✅ Matched</span>'
      : st === 'extracted'
      ? '<span class="badge badge-blue">→ Extracted</span>'
      : st === 'low-confidence'
      ? '<span class="badge badge-yellow">⚠ Low confidence</span>'
      : st === 'unreadable' || st === 'error'
      ? '<span class="badge badge-red">❌ Unreadable</span>'
      : '<span class="badge badge-red">⚠ Manual needed</span>';

    // Student cell — dropdown for non-matched, text for matched
    let studentCell = '';
    const currentStudent = r.match.student || r.match.bestGuess;
    if (st === 'matched' && currentStudent) {
      studentCell = `<div class="student-cell truncate" title="${escHtml(currentStudent.name)}">${escHtml(currentStudent.name)}</div>`;
    } else {
      const opts = (r.match.candidates || state.students.slice(0, 5))
        .map(c => {
          const s = c.student || c;
          return `<option value="${escHtml(s.atsId)}" data-name="${escHtml(s.name)}">${escHtml(s.name)}</option>`;
        }).join('');
      const selectedVal = currentStudent ? `value="${escHtml(currentStudent.atsId)}"` : '';
      studentCell = `<select class="assign-select" data-idx="${idx}">
        <option value="">— Select student —</option>
        ${opts}
      </select>`;
    }

    // New filename cell
    const nameCell = r.newName
      ? `<span class="new-name-cell">${escHtml(r.newName)}</span>`
      : '<span style="color:var(--gray-400);font-size:12px;">— pending assignment —</span>';

    // Name extracted from the filename (the matching signal)
    const fnName    = r.filenameHint || '—';
    const fnDisplay = fnName !== '—'
      ? escHtml(fnName)
      : `<span style="color:var(--gray-400);font-size:11px;">no name in filename</span>`;

    tr.innerHTML = `
      <td><input type="checkbox" class="cb row-cb" data-idx="${idx}" ${r.selected ? 'checked' : ''}></td>
      <td><div class="filename-cell truncate" title="${escHtml(r.pdf.name)}">${escHtml(r.pdf.name)}</div></td>
      <td><div class="ocr-name-cell">${fnDisplay}</div></td>
      <td>${studentCell}</td>
      <td style="font-size:12px;color:var(--gray-600);font-family:monospace;">${escHtml((r.match.student || currentStudent)?.atsId || '—')}</td>
      <td>
        <div class="conf-bar">
          <span class="conf-track"><span class="conf-fill" style="width:${confPct}%;background:${confColor}"></span></span>
          <span style="font-size:11px;color:var(--gray-500)">${confPct}%</span>
        </div>
      </td>
      <td>${nameCell}</td>
      <td>${badgeHtml}</td>
    `;
    tbody.appendChild(tr);
  });

  // Pre-fill dropdowns for low-confidence best guesses
  state.results.forEach((r, idx) => {
    if ((r.match.status === 'low-confidence' || r.match.status === 'manual') && r.match.bestGuess) {
      const sel = document.querySelector(`select[data-idx="${idx}"]`);
      if (sel) {
        sel.value = r.match.bestGuess.atsId;
        // Trigger the assignment
        const ev = new Event('change');
        sel.dispatchEvent(ev);
      }
    }
  });
}

function toggleRow(idx, checked) {
  state.results[idx].selected = checked;
  const n = state.results.filter(r => r.selected).length;
  document.getElementById('apply-label').textContent = `Apply Renaming (${n} files)`;
  // Sync select-all checkbox
  document.getElementById('select-all').checked = state.results.every(r => r.selected || !r.newName);
};

function toggleAll(checked) {
  state.results.forEach((r, idx) => {
    if (r.newName) {
      r.selected = checked;
      const cb = document.querySelector(`.row-cb[data-idx="${idx}"]`);
      if (cb) cb.checked = checked;
    }
  });
  const n = state.results.filter(r => r.selected).length;
  document.getElementById('apply-label').textContent = `Apply Renaming (${n} files)`;
};

function handleManualAssign(select) {
  const idx  = parseInt(select.dataset.idx, 10);
  const atsId = select.value;
  const r    = state.results[idx];

  if (!atsId) {
    r.match.student = null;
    r.newName = null;
    r.selected = false;
    // Update cells
    const row = select.closest('tr');
    row.querySelector('td:nth-child(5)').textContent = '—';
    row.querySelector('td:nth-child(7)').innerHTML = '<span style="color:var(--gray-400);font-size:12px;">— pending assignment —</span>';
    row.querySelector('.row-cb').checked = false;
    toggleRow(idx, false);
    return;
  }

  const student = state.students.find(s => s.atsId === atsId);
  if (!student) return;

  r.match.student = student;
  r.match.status  = 'low-confidence';
  const p = r.parsed;
  r.newName = buildFilenameFromTemplate(p, student, r.filenameHint, r.pdf.name, idx + 1, r.fileMeta || {});
  r.selected = true;

  // Update the row cells without full re-render
  const row = select.closest('tr');
  row.querySelector('td:nth-child(5)').textContent = student.atsId;
  row.querySelector('td:nth-child(7)').innerHTML = `<span class="new-name-cell">${escHtml(r.newName)}</span>`;
  const cb = row.querySelector('.row-cb');
  cb.checked = true;
  toggleRow(idx, true);
};

// ─────────────────────────────────────────────────────────────
// STEP 4 — APPLY RENAMING
// ─────────────────────────────────────────────────────────────
async function applyRenaming() {
  const toProcess = state.results.filter(r => r.selected && r.newName);
  if (toProcess.length === 0) { alert('No files selected for renaming.'); return; }

  const applyBtn = document.getElementById('apply-btn');
  applyBtn.disabled = true;
  applyBtn.innerHTML = '⏳ Copying files…';

  try {
    // Request write permission at apply time (not at folder selection)
    const perm = await state.dirHandle.requestPermission({ mode: 'readwrite' });
    if (perm !== 'granted') {
      alert('Write permission is required to create the output subfolder. Please allow it.');
      applyBtn.disabled = false;
      applyBtn.innerHTML = '⚠ Retry — write permission needed';
      return;
    }

    const folderName = `${state.dirHandle.name}_renamed`;
    const outDir = await state.dirHandle.getDirectoryHandle(folderName, { create: true });
    state.finalReport = [];
    let done = 0, errors = 0;

    for (const r of toProcess) {
      try {
        const srcHandle  = await state.dirHandle.getFileHandle(r.pdf.name);
        const file       = await srcHandle.getFile();
        const destHandle = await outDir.getFileHandle(r.newName, { create: true });
        const writable   = await destHandle.createWritable();
        await writable.write(file);
        await writable.close();
        state.finalReport.push({ ...r, applyStatus: 'ok' });
        done++;
      } catch (e) {
        state.finalReport.push({ ...r, applyStatus: 'error', applyError: e.message });
        errors++;
      }
    }

    // Also add unselected/unmatched to report
    state.results.filter(r => !r.selected || !r.newName).forEach(r => {
      state.finalReport.push({ ...r, applyStatus: 'skipped' });
    });

    // Show done step

    document.getElementById('done-title').textContent = `${done} file${done !== 1 ? 's' : ''} renamed!`;
    document.getElementById('done-sub').textContent = `Renamed copies saved to subfolder "${folderName}" inside your selected folder.`;
    document.getElementById('done-stats').innerHTML = `
      <span class="pill pill-green">✅ ${done} renamed</span>
      ${errors > 0 ? `<span class="pill pill-red">❌ ${errors} errors</span>` : ''}
      <span class="pill pill-blue">⏭ ${state.results.length - toProcess.length} skipped</span>
    `;
    showStep(4);

  } catch (e) {
    applyBtn.disabled = false;
    applyBtn.innerHTML = '⚠ Error — retry';
    alert('Error during renaming: ' + e.message);
    console.error(e);
  }
};

// ─────────────────────────────────────────────────────────────
// CSV REPORT
// ─────────────────────────────────────────────────────────────
function downloadCSV() {
  const report = state.finalReport.length ? state.finalReport : state.results;
  const rows = [
    ['Original Filename', 'Name from Filename', 'Matched Student', 'ATS ID', 'New Filename', 'Status'],
    ...report.map(r => [
      r.pdf.name,
      r.filenameHint || '',
      r.match?.student?.name || '',
      r.match?.student?.atsId || '',
      r.newName || '',
      r.applyStatus || r.match?.status || '',
    ]),
  ];

  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `renaming_report_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
};

// Re-apply the current template to existing OCR results — no re-scan needed
function reapplyTemplate() {
  if (state.results.length === 0) return;
  state.results.forEach((r, i) => {
    r.newName = buildFilenameFromTemplate(r.parsed, r.match?.student, r.filenameHint, r.pdf.name, i + 1, r.fileMeta || {});
    r.selected = true;
  });
  buildResultsTable();
  showStep(3);
  // Restore the normal start button for future use
  const btn = document.getElementById('start-btn');
  btn.textContent = 'Start Renaming';
  btn.onclick = null; // DOMContentLoaded listener still handles it via addEventListener
}

// ─────────────────────────────────────────────────────────────
// RESTART
// ─────────────────────────────────────────────────────────────
function restart() {
  state.dirHandle  = null;
  state.pdfFiles   = [];
  state.students   = [];
  state.results    = [];
  state.template      = { pattern: '' };
  state.customTokens  = {};
  state.finalReport = [];

  ['zone-pdf', 'zone-excel'].forEach(id => {
    const z = document.getElementById(id);
    z.classList.remove('loaded');
  });
  document.getElementById('zone-pdf').addEventListener('click', pickFolder);
  document.getElementById('zone-excel').addEventListener('click', pickExcel);
  document.getElementById('pdf-tag').classList.add('hidden');
  document.getElementById('excel-tag').classList.add('hidden');
  document.getElementById('zone-pdf').querySelector('h3').textContent = 'Folder with Files';
  document.getElementById('zone-pdf').querySelector('p').textContent = 'Click to select the folder containing all files to be renamed';
  document.getElementById('zone-excel').querySelector('h3').textContent = 'Name & ID File';
  document.getElementById('zone-excel').querySelector('p').textContent = 'Upload a .xlsx/.csv with names and IDs to cross-check against file content';
  document.getElementById('ocr-log').innerHTML = '';
  document.getElementById('start-btn').disabled = true;
  document.getElementById('start-btn').innerHTML = `<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg> Start Renaming`;

  // Reset template UI
  const patternEl = document.getElementById('template-pattern');
  if (patternEl) patternEl.value = '';
  document.getElementById('scan-section')?.classList.add('hidden');
  document.getElementById('scan-results')?.classList.add('hidden');
  // Remove discovered optgroup from dropdown
  document.querySelector('#token-insert-select optgroup[label="Discovered from your files"]')?.remove();
  updatePreview();
  updateExcelZoneBadge();  // reset badge to Optional (slots are all null after restart)
  updateStartBtn();         // ensure start button is disabled until folder is re-selected

  showStep(1);
};

// ─────────────────────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────────────────────
function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Surface any unhandled module errors to the page console
window.addEventListener('error', e => console.error('app.js error:', e.message, e.filename, e.lineno));
window.addEventListener('unhandledrejection', e => console.error('app.js unhandled promise:', e.reason));

// ─────────────────────────────────────────────────────────────
// TEMPLATE CONFIGURATOR UI
// ─────────────────────────────────────────────────────────────
const STORAGE_KEY = 'filerenamer_template';

// Sample values used in the live preview
const PREVIEW_SAMPLES = {
  name: 'Aysha Abdulla', id: '571883', ext: 'pdf', filetype: 'pdf',
  date: '2026-06-26', year: '2026', month: '06', day: '26',
  filename: 'Milestone 1 - Aysha', count: '001', index: '001',
  created: '2026-03-15', modified: '2026-06-22',
  title: 'Assignment Milestone 1', author: '', subject: '',
};

function saveTemplate() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ pattern: state.template.pattern }));
  } catch (_) {}
}

function loadTemplate() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const t = JSON.parse(raw);
    if (t.pattern) {
      state.template.pattern = t.pattern;
      const el = document.getElementById('template-pattern');
      if (el) el.value = t.pattern;
    }
  } catch (_) {}
}

function updatePreview() {
  const pattern   = (state.template.pattern || '').trim();
  const previewFn = document.getElementById('preview-filename');
  const previewEx = document.getElementById('preview-example');
  if (!previewFn) return;

  if (!pattern) {
    previewFn.textContent = '—';
    previewEx.textContent = 'Type a pattern above to see a preview';
    previewEx.style.color = '#ef4444';
    return;
  }

  previewEx.style.color = '';
  previewFn.textContent = pattern;   // show the pattern itself as the "template"

  // Substitute sample values to show an example filename
  let example = pattern;
  for (const [k, v] of Object.entries(PREVIEW_SAMPLES)) {
    example = example.replace(new RegExp(`\\{${k}\\}`, 'gi'), v);
  }
  previewEx.textContent = 'e.g.  ' + example;
}

function initTemplateUI() {
  const patternInput = document.getElementById('template-pattern');
  const tokenSelect  = document.getElementById('token-insert-select');

  // Track cursor position — clicking the dropdown steals focus from the input,
  // resetting selectionStart. We capture it on blur so insertion stays at cursor.
  let savedCursor = 0;

  if (patternInput) {
    patternInput.addEventListener('input', () => {
      state.template.pattern = patternInput.value;
      updatePreview();
      updateExcelZoneBadge();
      updateStartBtn();
      saveTemplate();
    });
    patternInput.addEventListener('focus', () => {
      patternInput.style.borderColor = 'var(--blue)';
    });
    patternInput.addEventListener('blur', () => {
      patternInput.style.borderColor = 'var(--gray-300)';
      savedCursor = patternInput.selectionStart ?? patternInput.value.length;
    });
    patternInput.addEventListener('keyup', () => {
      savedCursor = patternInput.selectionStart ?? patternInput.value.length;
    });
    patternInput.addEventListener('click', () => {
      savedCursor = patternInput.selectionStart ?? patternInput.value.length;
    });
  }

  // Token insert dropdown — uses savedCursor so position is correct even after focus leaves
  if (tokenSelect) {
    tokenSelect.addEventListener('change', () => {
      const token = tokenSelect.value;
      if (!token || !patternInput) { tokenSelect.value = ''; return; }
      const pos = savedCursor;
      const val = patternInput.value;
      patternInput.value = val.slice(0, pos) + token + val.slice(pos);
      savedCursor = pos + token.length;
      patternInput.selectionStart = patternInput.selectionEnd = savedCursor;
      patternInput.focus();
      state.template.pattern = patternInput.value;
      updatePreview();
      updateExcelZoneBadge();
      updateStartBtn();
      saveTemplate();
      tokenSelect.value = ''; // reset to placeholder
    });
  }

  updatePreview();
}

// ─────────────────────────────────────────────────────────────
// WIRE UP ALL EVENT LISTENERS (replaces all removed inline handlers)
// ─────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Template configurator — init first, then restore saved state
  initTemplateUI();
  loadTemplate();              // restore last-used template from localStorage
  updateExcelZoneBadge();      // set badge based on restored slots

  // Step 1 — file loaders
  document.getElementById('zone-pdf').addEventListener('click', pickFolder);
  document.getElementById('zone-excel').addEventListener('click', pickExcel);
  document.getElementById('start-btn').addEventListener('click', startProcessing);

  // Step 3 — back to setup (preserves loaded files and OCR results)
  document.getElementById('back-to-setup-btn').addEventListener('click', () => {
    showStep(1);
    // If we already have OCR results, add a quick "Re-apply template" button
    if (state.results.length > 0) {
      const btn = document.getElementById('start-btn');
      btn.disabled = false;
      btn.textContent = '↺ Re-apply Template (no re-scan)';
      btn.onclick = reapplyTemplate;
    }
  });

  // Step 3 — select-all checkbox
  document.getElementById('select-all').addEventListener('change', e => toggleAll(e.target.checked));

  // Step 3 — apply + CSV buttons
  document.getElementById('apply-btn').addEventListener('click', applyRenaming);
  document.getElementById('csv-btn').addEventListener('click', downloadCSV);

  // Step 4 — done-page buttons
  document.getElementById('done-csv-btn').addEventListener('click', downloadCSV);
  document.getElementById('restart-btn').addEventListener('click', restart);

  // Smart Scan button
  document.getElementById('scan-btn').addEventListener('click', async () => {
    if (!state.pdfFiles.length) return;
    const btn   = document.getElementById('scan-btn');
    const label = document.getElementById('scan-btn-label');
    label.textContent = 'Scanning with OCR…';
    btn.disabled = true;
    try {
      const sample  = state.pdfFiles[0]; // use first file as sample
      const results = await scanSampleFile(sample.handle);
      showScanResults(sample.name, results);
    } catch (e) {
      label.textContent = 'Scan failed — ' + e.message;
    } finally {
      label.textContent = 'Scan text content in a file (OCR)';
      btn.disabled = false;
    }
  });

  // Step 3 — event delegation for dynamic table rows (checkboxes + dropdowns)
  document.getElementById('results-tbody').addEventListener('change', e => {
    if (e.target.classList.contains('row-cb')) {
      const idx = parseInt(e.target.dataset.idx, 10);
      toggleRow(idx, e.target.checked);
    }
    if (e.target.classList.contains('assign-select')) {
      handleManualAssign(e.target);
    }
  });
});

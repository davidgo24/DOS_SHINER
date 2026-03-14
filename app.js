/**
 * DOS Report Formatter
 * Parses raw Excel report and transforms to formatted layout with configurable options
 */

// Default column mapping - maps output display name to possible raw column names
const DEFAULT_COLUMN_MAP = {
  paddle: ['Paddle', 'paddle'],
  block: ['Block', 'block'],
  plannedStart: ['Planned Shift Start Time', 'Planned Shift Start', 'planned_start'],
  plannedEnd: ['Planned Shift End Time', 'Planned Shift End', 'planned_end'],
  plannedHrs: ['Hrs (Planned Duration)', 'Planned Duration', 'planned_hrs'],
  vehicle: ['Vehicle', 'vehicle'],
  actualStart: ['Actual Start Time', 'actual_start'],
  actualEnd: ['Actual End Time', 'actual_end'],
  trim: ['Trim (Actual Duration, Hrs.)', 'Trim', 'trim', 'actual_duration'],
  primaryDriver: ['Primary Driver Name', 'Primary Driver', 'primary_driver'],
  primaryId: ['Primary Driver ID', 'Primary ID', 'primary_id'],
  altDriver: ['Alternative Driver Name', 'Alternative Driver', 'alt_driver'],
  altId: ['Alternative Driver ID', 'Alternative ID', 'alt_id'],
  labels: ['Labels', 'labels'],
  driverNotes: ['Driver Notes', 'driver_notes'],
  internalNotes: ['Internal Notes', 'internal_notes'],
  cancelled: ['Was Cancelled', 'cancelled'],
};

const DISPLAY_COLUMNS = [
  'paddle', 'block', 'plannedStart', 'plannedEnd', 'plannedHrs', 'vehicle',
  'actualStart', 'actualEnd', 'trim', 'primaryDriver', 'primaryId', 'altDriver', 'altId',
  'labels', 'driverNotes', 'internalNotes', 'cancelled'
];

// Bucket-based configuration - fixed order, each bucket has paddles and color
const BUCKET_DEFAULTS = [
  { id: 'paddle', label: 'Paddle', color: 'FFFFFF', paddles: ['*paddle'], locked: true },
  { id: 'extraBoard', label: 'Extra-Board', color: 'CCFBF1', paddles: ['AM Extra-Board', 'Mid Extra-Board', 'PM Extra-Board'], locked: false },
  { id: 'supervisors', label: 'Supervisors', color: 'E0E7FF', paddles: ['FIELD SUPERVISOR', 'Field 1', 'Field 2', 'Field 3', 'Field 4', 'Mid-Field', 'OPS', 'MID/OPS', 'Open', 'Closing'], locked: false },
  { id: 'trainees', label: 'Trainees', color: 'DDD6FE', paddles: ['BTW/TRN', 'Classroom / BTW', 'Classroom BTW / BTW', '(REV/TRN)', 'REV/TRN'], locked: false },
  { id: 'leave', label: 'Leave', color: 'FCE7F3', paddles: ['Sick', 'TTD', 'FMLA', 'P/L', 'VAC', 'Admin Leave', 'C/B'], locked: false },
  { id: 'other', label: 'Other', color: 'FFFFFF', paddles: [], locked: true },
];

const STORAGE_KEYS = { buckets: 'dos_buckets', reportTitle: 'dos_report_title', reportDate: 'dos_report_date', exportLabel: 'dos_export_label', columnPadding: 'dos_column_padding', theme: 'dos_theme' };

function findColumn(row, possibleNames) {
  for (const name of possibleNames) {
    const idx = row.findIndex(cell => {
      const val = cell != null ? String(cell).trim() : '';
      return val.toLowerCase() === name.toLowerCase();
    });
    if (idx >= 0) return idx;
  }
  return -1;
}

function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', raw: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  const data = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[addr];
      row.push(cell ? cell.v : null);
    }
    data.push(row);
  }
  return data;
}

function getColumnOverride() {
  try {
    const raw = document.getElementById('columnOverride').value.trim();
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (_) {
    return null;
  }
}

function buildColumnIndex(headerRow) {
  const override = getColumnOverride();
  const index = {};
  for (const [key, defaultNames] of Object.entries(DEFAULT_COLUMN_MAP)) {
    const names = override && override[key] ? [override[key]] : defaultNames;
    const idx = findColumn(headerRow, names);
    if (idx >= 0) index[key] = idx;
  }
  return index;
}

/** Find the first row that looks like the header (has "Paddle" etc.) */
function findHeaderRowIndex(raw) {
  for (let r = 0; r < raw.length; r++) {
    const row = raw[r];
    if (!row || !row.length) continue;
    const idx = findColumn(row, DEFAULT_COLUMN_MAP.paddle);
    if (idx >= 0) return r;
  }
  return 0;
}

const DAY_NAMES = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];

function parseDateFromString(s) {
  const parts = s.replace(/\./g, '-').split(/[\/\-]/);
  if (parts.length !== 3) return null;
  let year, month, day;
  if (parts[0].length === 4) {
    year = parseInt(parts[0], 10);
    month = parseInt(parts[1], 10) - 1;
    day = parseInt(parts[2], 10);
  } else {
    month = parseInt(parts[0], 10) - 1;
    day = parseInt(parts[1], 10);
    year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
  }
  const d = new Date(year, month, day);
  if (isNaN(d.getTime())) return null;
  return d;
}

function formatDateWithDay(date) {
  const dayName = DAY_NAMES[date.getDay()];
  const m = date.getMonth() + 1;
  const d = date.getDate();
  const y = date.getFullYear();
  return `${dayName} ${m}/${d}/${y}`;
}

/** Extract date from filename and format as "WEDNESDAY 3/14/2026" */
function extractDateFromFilename(filename) {
  const base = filename.replace(/\.(xlsx|xls)$/i, '');
  const match = base.match(/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})/);
  if (match) {
    const d = parseDateFromString(match[1]);
    if (d) return formatDateWithDay(d);
  }
  return '';
}

/** Extract report title from preamble rows (before header) - only when field is empty */
function extractTitleFromPreamble(raw, headerRowIndex) {
  if (headerRowIndex === 0) return '';
  const preamble = raw.slice(0, headerRowIndex);
  for (const row of preamble) {
    if (!row) continue;
    for (const cell of row) {
      const s = cell != null ? String(cell).trim() : '';
      if (s && s.length < 100 && !/\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}/.test(s)) {
        return s;
      }
    }
  }
  return preamble[0]?.[0] != null ? String(preamble[0][0]).trim() : '';
}

function getValue(row, colIndex) {
  if (colIndex == null || colIndex < 0) return '';
  const val = row[colIndex];
  if (val == null) return '';
  if (typeof val === 'number') return val === Math.floor(val) ? String(val) : String(val);
  return String(val).trim();
}

function getBucketConfig() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.buckets);
    if (!raw) return BUCKET_DEFAULTS.map(b => ({ ...b, paddles: [...b.paddles] }));
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return BUCKET_DEFAULTS.map(b => ({ ...b, paddles: [...b.paddles] }));
    const defaultsById = Object.fromEntries(BUCKET_DEFAULTS.map(b => [b.id, b]));
    return parsed.map(b => {
      const def = defaultsById[b.id];
      return {
        id: b.id,
        label: b.label ?? def?.label ?? b.id,
        color: b.color ?? def?.color ?? 'FFFFFF',
        paddles: Array.isArray(b.paddles) ? b.paddles : (def?.paddles ? [...def.paddles] : []),
        locked: def ? def.locked : (b.locked ?? false),
      };
    });
  } catch (_) { return BUCKET_DEFAULTS.map(b => ({ ...b, paddles: [...b.paddles] })); }
}

function saveBucketConfig(buckets) {
  localStorage.setItem(STORAGE_KEYS.buckets, JSON.stringify(buckets.map(b => ({ id: b.id, label: b.label, color: b.color, paddles: b.paddles, locked: b.locked }))));
}

function classifyRowToBucket(paddleVal, buckets) {
  const s = String(paddleVal || '').trim();
  if (!s) return 'other';

  // Numeric paddle blocks first
  if (/^\d{5}\s*$/.test(s) || /^\d{5}\s*\(\d+\/\d+\)/.test(s)) {
    return 'paddle';
  }

  const valNorm = s.toUpperCase();
  for (const bucket of buckets) {
    if (bucket.id === 'paddle' || bucket.id === 'other') continue;
    for (const p of bucket.paddles) {
      if (p === '*paddle') continue;
      const pNorm = p.toUpperCase();
      if (valNorm === pNorm || valNorm.includes(pNorm) || pNorm.includes(valNorm)) {
        return bucket.id;
      }
    }
  }
  return 'other';
}

function getBucketById(buckets, id) {
  return buckets.find(b => b.id === id) || null;
}

function getBucketColor(bucketId, buckets) {
  const b = getBucketById(buckets, bucketId);
  return b ? b.color : 'FFFFFF';
}

function loadFromStorage(key) {
  try {
    return localStorage.getItem(key);
  } catch (_) { return null; }
}

function saveToStorage(key, value) {
  try { localStorage.setItem(key, value); } catch (_) {}
}

function initBuckets() {
  const container = document.getElementById('bucketsList');
  renderBuckets();
  document.getElementById('addBucketBtn').addEventListener('click', addCustomBucket);
}

function renderBuckets() {
  const buckets = getBucketConfig();
  const container = document.getElementById('bucketsList');
  container.innerHTML = '';
  buckets.forEach((bucket, idx) => {
    const card = document.createElement('div');
    card.className = 'bucket-card';
    card.dataset.bucketId = bucket.id;
    const paddlesStr = bucket.paddles.join(', ');
    card.innerHTML = `
      <div class="bucket-header">
        <div class="bucket-swatch" style="background-color:#${bucket.color}"></div>
        ${bucket.locked
          ? `<span class="bucket-label">${escapeHtml(bucket.label)}</span>`
          : `<input type="text" class="bucket-label-input" placeholder="Bucket name">`
        }
        ${!bucket.locked && bucket.id !== 'other' ? `<button type="button" class="btn-icon bucket-remove" title="Remove bucket">×</button>` : ''}
      </div>
      <div class="bucket-paddles">
        <input type="text" data-bucket-id="${bucket.id}" placeholder="Paddles (comma-separated)…">
      </div>
    `;
    const labelEl = card.querySelector('.bucket-label-input');
    if (labelEl) labelEl.value = bucket.label;
    const paddlesEl = card.querySelector('.bucket-paddles input');
    paddlesEl.value = paddlesStr;
    if (bucket.id === 'other') {
      paddlesEl.disabled = true;
      paddlesEl.placeholder = 'Unassigned rows appear here automatically';
    } else if (bucket.id === 'paddle') {
      paddlesEl.disabled = true;
      paddlesEl.placeholder = 'Numeric blocks (10001–10077)';
    }
    const labelInput = card.querySelector('.bucket-label-input');
    const paddlesInput = card.querySelector('.bucket-paddles input');
    const removeBtn = card.querySelector('.bucket-remove');

    if (labelInput) {
      labelInput.addEventListener('change', () => {
        const cfg = getBucketConfig();
        const b = cfg.find(x => x.id === bucket.id);
        if (b) { b.label = labelInput.value.trim() || b.label; saveBucketConfig(cfg); }
      });
    }
    paddlesInput.addEventListener('change', () => {
      const cfg = getBucketConfig();
      const b = cfg.find(x => x.id === bucket.id);
      if (b) {
        b.paddles = paddlesInput.value.split(/[,\n]+/).map(s => s.trim()).filter(Boolean);
        saveBucketConfig(cfg);
        reapplyTransform();
      }
    });
    if (removeBtn) {
      removeBtn.addEventListener('click', () => {
        const cfg = getBucketConfig();
        const filtered = cfg.filter(b => b.id !== bucket.id);
        saveBucketConfig(filtered);
        renderBuckets();
        reapplyTransform();
      });
    }
    container.appendChild(card);
  });
}

function addCustomBucket() {
  const cfg = getBucketConfig();
  const otherIdx = cfg.findIndex(b => b.id === 'other');
  const newBucket = { id: 'custom_' + Date.now(), label: 'New bucket', color: 'E8E8E8', paddles: [], locked: false };
  cfg.splice(otherIdx >= 0 ? otherIdx : cfg.length, 0, newBucket);
  saveBucketConfig(cfg);
  renderBuckets();
}

function transformRows(rows, colIndex) {
  const buckets = getBucketConfig();
  const byBucket = new Map();
  for (const b of buckets) {
    byBucket.set(b.id, []);
  }

  for (const row of rows) {
    const paddleVal = getValue(row, colIndex.paddle);
    const bucketId = classifyRowToBucket(paddleVal, buckets);

    const primaryDriver = getValue(row, colIndex.primaryDriver);
    const altDriver = getValue(row, colIndex.altDriver);
    const primaryId = getValue(row, colIndex.primaryId);
    const altId = getValue(row, colIndex.altId);

    const displayPrimary = primaryDriver || altDriver;
    const displayPrimaryId = primaryDriver ? primaryId : altId;
    const displayAlt = primaryDriver ? altDriver : '';
    const displayAltId = primaryDriver ? altId : '';

    const driverNotes = getValue(row, colIndex.driverNotes);
    const internalNotes = getValue(row, colIndex.internalNotes);
    const notes = [driverNotes, internalNotes].filter(Boolean).join(' ');

    const rec = {
      paddle: paddleVal,
      block: getValue(row, colIndex.block),
      plannedStart: getValue(row, colIndex.plannedStart),
      plannedEnd: getValue(row, colIndex.plannedEnd),
      plannedHrs: getValue(row, colIndex.plannedHrs),
      vehicle: getValue(row, colIndex.vehicle),
      actualStart: getValue(row, colIndex.actualStart),
      actualEnd: getValue(row, colIndex.actualEnd),
      trim: getValue(row, colIndex.trim),
      primaryDriver: displayPrimary,
      primaryId: displayPrimaryId,
      altDriver: displayAlt,
      altId: displayAltId,
      labels: getValue(row, colIndex.labels),
      driverNotes: notes,
      internalNotes: '',
      cancelled: getValue(row, colIndex.cancelled),
      _section: bucketId,
    };

    const target = byBucket.get(bucketId) || byBucket.get('other');
    if (target) target.push(rec);
  }

  // Sort paddle bucket rows by block then paddle number
  const paddleRows = byBucket.get('paddle') || [];
  paddleRows.sort((a, b) => {
    const blockA = parseInt(a.block, 10) || 99999;
    const blockB = parseInt(b.block, 10) || 99999;
    if (blockA !== blockB) return blockA - blockB;
    const numA = parseInt(String(a.paddle).replace(/\D/g, ''), 10) || 99999;
    const numB = parseInt(String(b.paddle).replace(/\D/g, ''), 10) || 99999;
    return numA - numB;
  });

  const result = [];
  for (const b of buckets) {
    const rowsForBucket = b.id === 'paddle' ? paddleRows : (byBucket.get(b.id) || []);
    result.push(...rowsForBucket);
  }

  return result;
}

function renderTable(records, isHeader = false) {
  const headers = [
    'Paddle', 'Block', 'Planned Shift Start', 'Planned Shift End', 'Hrs (Planned)',
    'Vehicle', 'Actual Start', 'Actual End', 'Trim', 'Primary Driver', 'ID',
    'Alternative Driver', 'Alt ID', 'Labels', 'Driver Notes', 'Internal Notes', 'Was Cancelled'
  ];
  const keys = [
    'paddle', 'block', 'plannedStart', 'plannedEnd', 'plannedHrs', 'vehicle',
    'actualStart', 'actualEnd', 'trim', 'primaryDriver', 'primaryId', 'altDriver', 'altId',
    'labels', 'driverNotes', 'internalNotes', 'cancelled'
  ];

  let html = '<thead><tr>';
  for (const h of headers) {
    html += `<th>${h}</th>`;
  }
  html += '</tr></thead><tbody>';

  const buckets = getBucketConfig();
  for (const r of records) {
    const bucketId = r._section || '';
    const hasAltDriver = bucketId === 'paddle' && !!(r.altDriver && String(r.altDriver).trim());
    let fill = bucketId === 'footer' ? 'FFFFFF' : getBucketColor(bucketId, buckets);
    if (bucketId === 'paddle') fill = hasAltDriver ? 'FFEB3B' : 'FFFFFF';
    html += `<tr data-section="${escapeHtml(bucketId)}"${hasAltDriver ? ' data-has-alt-driver="true"' : ''} style="background-color:#${fill}">`;
    for (const k of keys) {
      html += `<td>${escapeHtml(r[k] || '')}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody>';
  return html;
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function processFile(buffer, filename = '') {
  const raw = parseExcel(buffer);
  if (!raw.length) {
    document.getElementById('noData').textContent = 'No data found in the file.';
    document.getElementById('outputTable').innerHTML = '';
    document.getElementById('outputTable').style.display = 'none';
    return null;
  }

  const headerRowIndex = findHeaderRowIndex(raw);
  const headerRow = raw[headerRowIndex];
  const colIndex = buildColumnIndex(headerRow);

  const dataRows = raw.slice(headerRowIndex + 1);
  const transformed = transformRows(dataRows, colIndex);

  // Auto-populate report date from filename (e.g. "Report_3-11-2026_to_3-17-2026.xlsx")
  const extractedDate = extractDateFromFilename(filename);
  if (extractedDate) {
    const reportDateEl = document.getElementById('reportDate');
    reportDateEl.value = extractedDate;
    saveToStorage(STORAGE_KEYS.reportDate, extractedDate);
  }
  // Auto-populate report title from preamble only when field is empty
  if (headerRowIndex > 0 && !document.getElementById('reportTitle').value.trim()) {
    const extractedTitle = extractTitleFromPreamble(raw, headerRowIndex);
    if (extractedTitle) {
      document.getElementById('reportTitle').value = extractedTitle;
    }
  }

  const reportTitle = document.getElementById('reportTitle').value ?? '';
  const reportDate = document.getElementById('reportDate').value || '';

  const footerRows = [
    { paddle: reportTitle, block: '', plannedStart: '', plannedEnd: '', plannedHrs: '', vehicle: '', actualStart: '', actualEnd: '', trim: '', primaryDriver: reportDate, primaryId: '', altDriver: '', altId: '', labels: '', driverNotes: '', internalNotes: '', cancelled: '', _section: 'footer' },
    { paddle: '-- 1 of 1 --', block: '', plannedStart: '', plannedEnd: '', plannedHrs: '', vehicle: '', actualStart: '', actualEnd: '', trim: '', primaryDriver: '', primaryId: '', altDriver: '', altId: '', labels: '', driverNotes: '', internalNotes: '', cancelled: '', _section: 'footer' },
  ];

  const allRows = [...transformed, ...footerRows];

  const table = document.getElementById('outputTable');
  table.innerHTML = renderTable(allRows);
  table.style.display = 'table';
  document.getElementById('noData').style.display = 'none';

  const otherRows = transformed.filter(r => r._section === 'other');
  const newPaddles = getNewPaddlesFromOther(otherRows);
  if (newPaddles.length > 0) {
    showNewPaddlesBanner(newPaddles);
  } else {
    hideNewPaddlesBanner();
  }

  window.__lastTransformedData = { headers: Object.keys(transformed[0] || {}), rows: allRows, raw };
  return allRows;
}

function initDropZone() {
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      handleFile(file);
    } else {
      alert('Please drop an Excel file (.xlsx or .xls)');
    }
  });

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
  });
}

async function handleFile(file) {
  showLoading(true);
  try {
    const buffer = await file.arrayBuffer();
    window.__lastBuffer = buffer;
    window.__lastFilename = file.name;
    processFile(buffer, file.name);
  } catch (err) {
    alert('Could not process the file. Make sure it is a valid Excel file with the expected columns.');
    console.error(err);
  } finally {
    showLoading(false);
  }
}

function showLoading(show) {
  let overlay = document.getElementById('loadingOverlay');
  if (show) {
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'loadingOverlay';
      overlay.className = 'loading-overlay';
      overlay.innerHTML = '<div class="spinner"></div><span class="text">Processing…</span>';
      document.body.appendChild(overlay);
    }
    overlay.style.display = 'flex';
  } else if (overlay) {
    overlay.style.display = 'none';
  }
}

function resetApp() {
  window.__lastBuffer = null;
  window.__lastTransformedData = null;
  window.__lastFilename = '';
  document.getElementById('fileInput').value = '';
  document.getElementById('noData').textContent = 'Upload an Excel file to see the formatted preview';
  document.getElementById('noData').style.display = 'block';
  document.getElementById('outputTable').innerHTML = '';
  document.getElementById('outputTable').style.display = 'none';
  hideNewPaddlesBanner();
}

const DEFAULT_FILL = 'FFFFFF';
const HEADER_FILL = 'F3F4F6';

function exportExcel() {
  const data = window.__lastTransformedData;
  if (!data || !data.rows || !data.rows.length) {
    alert('No data to export. Please upload a file first.');
    return;
  }

  const headers = ['Paddle', 'Block', 'Planned Shift Start', 'Planned Shift End', 'Hrs (Planned)', 'Vehicle', 'Actual Start', 'Actual End', 'Trim', 'Primary Driver', 'ID', 'Alternative Driver', 'Alt ID', 'Labels', 'Driver Notes', 'Internal Notes', 'Was Cancelled'];
  const keys = ['paddle', 'block', 'plannedStart', 'plannedEnd', 'plannedHrs', 'vehicle', 'actualStart', 'actualEnd', 'trim', 'primaryDriver', 'primaryId', 'altDriver', 'altId', 'labels', 'driverNotes', 'internalNotes', 'cancelled'];

  const wsData = [headers, ...data.rows.map(r => keys.map(k => r[k] ?? ''))];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  const thinBlack = { style: 'thin', color: { rgb: '000000' } };
  const border = { top: thinBlack, bottom: thinBlack, left: thinBlack, right: thinBlack };

  const centerAlign = { horizontal: 'center', vertical: 'center' };
  const centerWrap = { horizontal: 'center', vertical: 'center', wrapText: true };

  // Style header row - larger font, bold, centered; columns C-D wrap so text fits in narrow columns
  for (let c = 0; c < headers.length; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[addr]) {
      const isWrapHeader = c === 2 || c === 3; // C: Planned Shift Start, D: Planned Shift End
      ws[addr].s = {
        font: { bold: true, sz: 12 },
        fill: { fgColor: { rgb: HEADER_FILL } },
        border,
        alignment: isWrapHeader ? centerWrap : centerAlign,
      };
    }
  }

  const lastRowIdx = data.rows.length; // 0-based sheet row index of last row
  const footerRow1 = lastRowIdx - 1; // title/date row
  const footerRow2 = lastRowIdx;     // "-- 1 of 1 --" row

  // Style data rows - bold all, section fill
  data.rows.forEach((r, rowIdx) => {
    const bucketId = r._section || '';
    const buckets = getBucketConfig();
    let fill = getBucketColor(bucketId, buckets);
    if (bucketId === 'paddle') {
      fill = (r.altDriver && String(r.altDriver).trim()) ? 'FFEB3B' : 'FFFFFF'; // Highlighter yellow
    }
    const isFooter = section === 'footer';
    const isTitleDateRow = isFooter && rowIdx === footerRow1 - 1; // -1 because rowIdx is 0-based in data.rows
    const isPageBreakRow = isFooter && rowIdx === footerRow2 - 1;

    for (let c = 0; c < keys.length; c++) {
      const addr = XLSX.utils.encode_cell({ r: rowIdx + 1, c });
      if (ws[addr]) {
        const fontSize = isTitleDateRow ? 16 : (isPageBreakRow ? 14 : 11);
        ws[addr].s = {
          font: { bold: true, sz: fontSize },
          fill: { fgColor: { rgb: fill } },
          border,
          alignment: (isTitleDateRow || isPageBreakRow) ? { horizontal: 'center', vertical: 'center', wrapText: true } : centerAlign,
        };
      }
    }
  });

  // Merge and style footer rows - title/date and "-- 1 of 1 --"
  ws['!merges'] = [
    { s: { r: footerRow1, c: 0 }, e: { r: footerRow1, c: headers.length - 1 } },
    { s: { r: footerRow2, c: 0 }, e: { r: footerRow2, c: headers.length - 1 } },
  ];

  // Set combined title+date in merged cell (top-left holds value)
  const titleDateRow = data.rows[footerRow1 - 1];
  const title = titleDateRow?.paddle ?? '';
  const date = titleDateRow?.primaryDriver ?? '';
  const combinedText = [title, date].filter(Boolean).join('\n');
  const titleDateAddr = XLSX.utils.encode_cell({ r: footerRow1, c: 0 });
  if (ws[titleDateAddr]) {
    ws[titleDateAddr].v = combinedText;
  }

  // Auto-size column widths - proportional padding so long text doesn't get cut off
  // Columns C-I (indices 2-8): time/vehicle data - cap at 12 chars
  const TIME_COLS = new Set([2, 3, 4, 5, 6, 7, 8]); // Planned Start/End, Hrs, Vehicle, Actual Start/End, Trim
  const colPadding = parseInt(document.getElementById('columnPadding')?.value, 10);
  const userPad = Number.isNaN(colPadding) || colPadding < 0 ? 0 : Math.min(colPadding, 5);
  ws['!cols'] = headers.map((_, c) => {
    let maxLen = headers[c]?.length ?? 6;
    for (let r = 0; r < wsData.length; r++) {
      const val = wsData[r]?.[c];
      const len = val != null ? String(val).length : 0;
      maxLen = Math.max(maxLen, len);
    }
    if (TIME_COLS.has(c)) {
      return { wch: Math.min(12, Math.max(6, maxLen + 1 + userPad)) };
    }
    // Long columns get extra padding to prevent cut-off; short columns stay tight
    const proportionalPad = maxLen > 25 ? Math.ceil(maxLen * 0.15) : (maxLen > 15 ? 3 : 2);
    const pad = userPad + proportionalPad;
    return { wch: Math.min(55, Math.max(5, maxLen + pad)) };
  });

  // Row heights - taller for header and footer rows
  ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 42 };           // Header row - taller for wrapped C/D headers
  ws['!rows'][footerRow1] = { hpt: 36 }; // Title/date row
  ws['!rows'][footerRow2] = { hpt: 24 }; // "-- 1 of 1 --" row

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Formatted Report');
  const reportDate = document.getElementById('reportDate')?.value?.trim();
  const dateSlug = reportDate ? reportDate.replace(/\s+/g, '_').replace(/\//g, '-').slice(0, 30) : '';
  const label = document.getElementById('exportLabel')?.value?.trim() || '';
  const base = dateSlug ? `DOS_Report_${dateSlug}` : 'DOS_Report_Formatted';
  const filename = label ? `${base}_${label}.xlsx` : `${base}.xlsx`;
  XLSX.writeFile(wb, filename, { cellStyles: true });
}

function printPdf() {
  window.print();
}

function reapplyTransform() {
  if (window.__lastBuffer) processFile(window.__lastBuffer, window.__lastFilename || '');
}

function getNewPaddlesFromOther(rows) {
  const seen = new Set();
  const result = [];
  for (const r of rows) {
    if (r._section !== 'other') continue;
    const p = String(r.paddle || '').trim();
    if (p && !seen.has(p)) {
      seen.add(p);
      result.push(p);
    }
  }
  return result;
}

function showNewPaddlesBanner(newPaddles) {
  const banner = document.getElementById('newPaddlesBanner');
  const list = document.getElementById('newPaddlesList');
  list.innerHTML = '';
  const buckets = getBucketConfig().filter(b => b.id !== 'paddle' && b.id !== 'other');
  window.__pendingNewPaddleAssignments = {};
  newPaddles.forEach(p => {
    window.__pendingNewPaddleAssignments[p] = '';
    const row = document.createElement('div');
    row.className = 'new-paddle-row';
    const select = document.createElement('select');
    select.dataset.paddle = p;
    select.innerHTML = `<option value="">— Skip (stay in Other) —</option>${buckets.map(b => `<option value="${b.id}">${escapeHtml(b.label)}</option>`).join('')}`;
    select.addEventListener('change', () => { window.__pendingNewPaddleAssignments[p] = select.value; });
    row.innerHTML = `<label>${escapeHtml(p)}</label>`;
    row.appendChild(select);
    list.appendChild(row);
  });
  banner.style.display = 'block';
}

function hideNewPaddlesBanner() {
  document.getElementById('newPaddlesBanner').style.display = 'none';
}

function applyNewPaddleAssignments() {
  const assignments = window.__pendingNewPaddleAssignments || {};
  const cfg = getBucketConfig();
  for (const [paddle, bucketId] of Object.entries(assignments)) {
    if (!bucketId) continue;
    const b = cfg.find(x => x.id === bucketId);
    if (b && !b.paddles.includes(paddle)) {
      b.paddles.push(paddle);
    }
  }
  saveBucketConfig(cfg);
  hideNewPaddlesBanner();
  reapplyTransform();
}

function initTheme() {
  const saved = loadFromStorage(STORAGE_KEYS.theme);
  const theme = saved === 'light' ? 'light' : 'dark';
  document.documentElement.dataset.theme = theme;
  const btn = document.getElementById('themeToggle');
  btn.textContent = theme === 'dark' ? '☀️' : '🌙';
  btn.title = theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode';
  btn.addEventListener('click', () => {
    const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    btn.textContent = next === 'dark' ? '☀️' : '🌙';
    btn.title = next === 'dark' ? 'Switch to light mode' : 'Switch to dark mode';
    saveToStorage(STORAGE_KEYS.theme, next);
  });
}

document.addEventListener('DOMContentLoaded', () => {
  const reportTitleEl = document.getElementById('reportTitle');
  const reportDateEl = document.getElementById('reportDate');
  const columnPaddingEl = document.getElementById('columnPadding');

  if (loadFromStorage(STORAGE_KEYS.reportTitle)) reportTitleEl.value = loadFromStorage(STORAGE_KEYS.reportTitle);
  if (loadFromStorage(STORAGE_KEYS.reportDate)) reportDateEl.value = loadFromStorage(STORAGE_KEYS.reportDate);
  const savedLabel = loadFromStorage(STORAGE_KEYS.exportLabel);
  if (savedLabel) document.getElementById('exportLabel').value = savedLabel;
  const pad = loadFromStorage(STORAGE_KEYS.columnPadding);
  if (pad != null && pad !== '') columnPaddingEl.value = pad;

  initDropZone();
  initBuckets();
  initTheme();

  document.getElementById('newPaddlesSkip').addEventListener('click', () => { hideNewPaddlesBanner(); });
  document.getElementById('newPaddlesApply').addEventListener('click', applyNewPaddleAssignments);

  document.getElementById('exportExcel').addEventListener('click', exportExcel);
  document.getElementById('printPdf').addEventListener('click', printPdf);
  document.getElementById('resetBtn').addEventListener('click', resetApp);

  const persist = (key, el) => { el.addEventListener('input', () => saveToStorage(key, el.value)); el.addEventListener('change', () => saveToStorage(key, el.value)); };
  persist(STORAGE_KEYS.reportTitle, reportTitleEl);
  persist(STORAGE_KEYS.reportDate, reportDateEl);
  persist(STORAGE_KEYS.exportLabel, document.getElementById('exportLabel'));
  persist(STORAGE_KEYS.columnPadding, columnPaddingEl);

  reportTitleEl.addEventListener('input', reapplyTransform);
  reportDateEl.addEventListener('input', reapplyTransform);
  document.getElementById('columnOverride').addEventListener('input', reapplyTransform);

  });

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

// Starter template sections - used for the dropdown and default order
const SECTION_TEMPLATES = [
  '*paddle',
  'AM Extra-Board',
  'Mid Extra-Board',
  'PM Extra-Board',
  'Field 1',
  'Field 2',
  'Field 3',
  'Field 4',
  'Mid-Field',
  'OPS',
  'MID/OPS',
  'Open',
  'Closing',
  'BTW/TRN',
  'Classroom / BTW',
  '(REV/TRN)',
  'C/B',
  'Sick',
  'FMLA',
  'P/L',
  'VAC',
  'Admin Leave',
  'TTD',
];

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

function classifyRow(paddleVal, sectionOrder) {
  const s = String(paddleVal || '').trim();
  if (!s) return '_other';

  // First try to match explicit section names (non-*paddle)
  for (const sec of sectionOrder) {
    if (sec === '*paddle') continue;
    const secNorm = sec.toUpperCase();
    const valNorm = s.toUpperCase();
    if (valNorm === secNorm || valNorm.includes(secNorm) || secNorm.includes(valNorm)) {
      return sec;
    }
  }

  // Numeric paddle blocks: 10001, 10002, 10011 (1/2), 10011 (2/2)
  // NOT (REV/TRN) 10051 - that goes to (REV/TRN) section
  if (/^\d{5}\s*$/.test(s) || /^\d{5}\s*\(\d+\/\d+\)/.test(s)) {
    return '*paddle';
  }
  return '_other';
}

function getSectionOrder() {
  const items = document.querySelectorAll('.section-item[data-section]');
  return Array.from(items).map(el => el.dataset.section);
}

function createSectionItem(sectionName) {
  const div = document.createElement('div');
  div.className = 'section-item';
  div.dataset.section = sectionName;
  div.draggable = true;
  div.innerHTML = `
    <span class="section-drag-handle" title="Drag to reorder">⋮⋮</span>
    <span class="section-label">${escapeHtml(sectionName)}</span>
    <div class="section-actions">
      <button type="button" class="btn-icon section-remove" title="Remove">×</button>
    </div>
  `;

  div.querySelector('.section-remove').addEventListener('click', () => div.remove());

  div.addEventListener('dragstart', (e) => {
    e.dataTransfer.setData('text/plain', sectionName);
    e.dataTransfer.effectAllowed = 'move';
    div.classList.add('section-dragging');
  });
  div.addEventListener('dragend', () => div.classList.remove('section-dragging'));

  div.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (div.classList.contains('section-dragging')) return;
    const list = document.getElementById('sectionList');
    const dragged = list.querySelector('.section-dragging');
    if (!dragged) return;
    const items = Array.from(list.querySelectorAll('.section-item'));
    const rect = div.getBoundingClientRect();
    const midY = rect.top + rect.height / 2;
    const insertBefore = e.clientY < midY;
    if (insertBefore && div.previousElementSibling !== dragged) {
      list.insertBefore(dragged, div);
    } else if (!insertBefore && div.nextElementSibling !== dragged) {
      list.insertBefore(dragged, div.nextElementSibling);
    }
  });

  return div;
}

function addSection(name) {
  const trimmed = String(name || '').trim();
  if (!trimmed) return;
  const list = document.getElementById('sectionList');
  const existing = Array.from(list.querySelectorAll('.section-item')).map(el => el.dataset.section);
  if (existing.includes(trimmed)) return;
  list.appendChild(createSectionItem(trimmed));
}

function initSectionOrder() {
  const list = document.getElementById('sectionList');
  const templateSelect = document.getElementById('sectionTemplate');
  const customInput = document.getElementById('sectionCustomInput');

  // Populate with default order
  for (const name of SECTION_TEMPLATES) {
    list.appendChild(createSectionItem(name));
  }

  templateSelect.addEventListener('change', () => {
    const val = templateSelect.value;
    if (!val) return;
    if (val === '__custom__') {
      customInput.style.display = 'inline-block';
      customInput.focus();
    } else {
      addSection(val);
    }
    templateSelect.value = '';
  });

  customInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      addSection(customInput.value);
      customInput.value = '';
      customInput.style.display = 'none';
    } else if (e.key === 'Escape') {
      customInput.value = '';
      customInput.style.display = 'none';
    }
  });
  customInput.addEventListener('blur', () => {
    if (customInput.value.trim()) {
      addSection(customInput.value);
      customInput.value = '';
    }
    customInput.style.display = 'none';
  });
}

function transformRows(rows, colIndex) {
  const sectionOrder = getSectionOrder();
  const bySection = new Map();
  for (const sec of sectionOrder) {
    bySection.set(sec, []);
  }
  bySection.set('_other', []);

  for (const row of rows) {
    const paddleVal = getValue(row, colIndex.paddle);
    const section = classifyRow(paddleVal, sectionOrder);

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
      internalNotes: '', // merged into driverNotes for display
      cancelled: getValue(row, colIndex.cancelled),
      _section: section,
    };

    if (bySection.has(section)) {
      bySection.get(section).push(rec);
    } else {
      // Try partial match for section names (e.g. "AM Extra-Board" in section order)
      let placed = false;
      for (const sec of sectionOrder) {
        if (sec === '*paddle') continue;
        if (paddleVal.toUpperCase().includes(sec.toUpperCase()) || sec.toUpperCase().includes(paddleVal.toUpperCase())) {
          bySection.get(sec).push(rec);
          placed = true;
          break;
        }
      }
      if (!placed) {
        bySection.get('_other').push(rec);
      }
    }
  }

  // Sort *paddle rows by block then paddle number
  const paddleRows = bySection.get('*paddle') || [];
  paddleRows.sort((a, b) => {
    const blockA = parseInt(a.block, 10) || 99999;
    const blockB = parseInt(b.block, 10) || 99999;
    if (blockA !== blockB) return blockA - blockB;
    const numA = parseInt(String(a.paddle).replace(/\D/g, ''), 10) || 99999;
    const numB = parseInt(String(b.paddle).replace(/\D/g, ''), 10) || 99999;
    return numA - numB;
  });

  const result = [];
  for (const sec of sectionOrder) {
    const rows = sec === '*paddle' ? paddleRows : (bySection.get(sec) || []);
    result.push(...rows);
  }
  result.push(...(bySection.get('_other') || []));

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

  for (const r of records) {
    const section = r._section || '';
    const hasAltDriver = section === '*paddle' && !!(r.altDriver && String(r.altDriver).trim());
    html += `<tr data-section="${escapeHtml(section)}"${hasAltDriver ? ' data-has-alt-driver="true"' : ''}>`;
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
    document.getElementById('reportDate').value = extractedDate;
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
  const buffer = await file.arrayBuffer();
  window.__lastBuffer = buffer;
  processFile(buffer, file.name);
}

// Section-based fill colors (Julio's look) - hex without #
const SECTION_FILL = {
  '*paddle': 'FFFFFF', // Overridden per-row: highlighter yellow only when altDriver present
  'AM Extra-Board': 'CCFBF1',
  'PM Extra-Board': 'CCFBF1',
  'Mid Extra-Board': 'CCFBF1',
  'Mid-Field': 'E0E7FF',
  'OPS': 'E0E7FF',
  'MID/OPS': 'E0E7FF',
  'Open': 'E0E7FF',
  'Closing': 'E0E7FF',
  'Field 1': 'E0E7FF',
  'Field 2': 'E0E7FF',
  'Field 3': 'E0E7FF',
  'Field 4': 'E0E7FF',
  'BTW/TRN': 'DDD6FE',
  'Classroom / BTW': 'DDD6FE',
  '(REV/TRN)': 'DDD6FE',
  'Sick': 'FCE7F3',
  'TTD': 'FCE7F3',
  'FMLA': 'FCE7F3',
  'P/L': 'FCE7F3',
  'VAC': 'FCE7F3',
  'Admin Leave': 'FCE7F3',
  'C/B': 'FCE7F3',
};
const DEFAULT_FILL = 'FFFFFF';
const HEADER_FILL = 'F3F4F6';

function getFillForSection(section) {
  if (!section) return DEFAULT_FILL;
  if (SECTION_FILL[section]) return SECTION_FILL[section];
  if (section.startsWith('Field')) return SECTION_FILL[section] || 'E0E7FF';
  return DEFAULT_FILL;
}

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
    const section = r._section || '';
    let fill = getFillForSection(section);
    // Numbered paddle rows only: yellow if alternate driver present, else white
    if (section === '*paddle') {
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
  XLSX.writeFile(wb, 'DOS_Report_Formatted.xlsx', { cellStyles: true });
}

function printPdf() {
  window.print();
}

function reapplyTransform() {
  if (window.__lastBuffer) {
    processFile(window.__lastBuffer);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  initDropZone();
  initSectionOrder();

  document.getElementById('exportExcel').addEventListener('click', exportExcel);
  document.getElementById('printPdf').addEventListener('click', printPdf);

  document.getElementById('reportTitle').addEventListener('input', reapplyTransform);
  document.getElementById('reportDate').addEventListener('input', reapplyTransform);
  document.getElementById('columnOverride').addEventListener('input', reapplyTransform);

  // Re-run transform when section list changes (add/remove/move)
  const sectionList = document.getElementById('sectionList');
  const observer = new MutationObserver(reapplyTransform);
  observer.observe(sectionList, { childList: true });
});

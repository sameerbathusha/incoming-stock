// =============================================================================
// mapping.js
// Turns a parsed Excel workbook into clean "canonical" rows the database
// understands, using a mapping template (the ones seeded in column_mappings).
// No browser/DOM code here on purpose, so the exact same logic can be tested
// outside the browser. Pass in the SheetJS library as `XLSX`.
// =============================================================================

export function normalizeHeader(s) {
  return String(s ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function hasDigit(s) {
  return /\d/.test(String(s ?? ''));
}

// Excel date cells come back as JS Date objects (cellDates:true). Use the
// local date parts so a Dubai user sees the same day that's in the sheet.
function cellToString(val) {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(val).trim();
  return s === '' ? null : s;
}

// Some sheets store dates as plain text ("30/12/2025", day-first in the Gulf)
// instead of real Excel dates. Normalize anything date-like to YYYY-MM-DD, or
// return null if it can't be understood (better a blank date than a failed
// upload).
function coerceDate(val) {
  const s = cellToString(val);
  if (s === null) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;                 // already ISO
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let day = parseInt(m[1], 10);
    let mon = parseInt(m[2], 10);
    let yr = m[3];
    if (day <= 12 && mon > 12) [day, mon] = [mon, day];        // tolerate month-first
    if (yr.length === 2) yr = '20' + yr;
    if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) {
      return `${yr}-${String(mon).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }
  const d = new Date(s);
  return isNaN(d) ? null : cellToString(d);
}

const DATE_FIELDS = new Set(['order_date', 'po_deadline', 'eta']);

// Recognize a "new item" column (named New / Tag / New Item) without needing it
// in the template. Truthy cells -> is_new true; explicit no/0/false -> false.
function findNewColumn(headers) {
  for (const cand of ['new', 'tag', 'new item', 'new launch']) {
    let i = headers.findIndex((h) => h === cand);
    if (i < 0) i = headers.findIndex((h) => h.startsWith(cand + ' '));
    if (i >= 0) return i;
  }
  return -1;
}
function parseNewFlag(val) {
  const s = String(val ?? '').trim().toLowerCase();
  if (s === '') return undefined;
  if (['new', 'yes', 'y', '1', 'true', 'x', '✓', 'newlaunch', 'launch'].includes(s)) return true;
  if (['no', 'n', '0', 'false', '-'].includes(s)) return false;
  return true; // any other non-empty note in a New column means "new"
}

// Find the column index whose normalized header matches a source header,
// trying exact, then prefix, then contains (handles multi-line headers like
// "Delivery Status\n( China warehouse )").
function findColumn(headers, sourceHeader) {
  const nk = normalizeHeader(sourceHeader);
  let idx = headers.findIndex((h) => h === nk);
  if (idx < 0) idx = headers.findIndex((h) => h.startsWith(nk));
  if (idx < 0) idx = headers.findIndex((h) => h.includes(nk));
  return idx;
}

/**
 * @param {object} workbook  SheetJS workbook (from XLSX.read)
 * @param {object} mapping   a column_mappings.mapping config object
 * @param {object} XLSX      the SheetJS library (browser global or node import)
 * @returns {{rows: object[], report: object}}
 */
export function mapWorkbook(workbook, mapping, XLSX) {
  const fields = mapping.fields || {};
  const sheetsCfg = mapping.sheets || {};
  const skipFirst = (mapping.skip_row_if_first_cell_matches || []).map((s) =>
    String(s).toLowerCase()
  );
  const headerToken = normalizeHeader(mapping.header_row_contains || 'voucher');
  const defaults = mapping.defaults || {};
  const statusMap = mapping.status_normalization || {};

  const rows = [];
  const report = { sheetsSeen: workbook.SheetNames, sheetsUsed: [], skipped: 0, perSheet: {} };

  for (const sheetName of workbook.SheetNames) {
    const cfg = sheetsCfg[sheetName];
    if (!cfg) continue; // only import sheets named in the template
    report.sheetsUsed.push(sheetName);

    const ws = workbook.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

    // locate the header row (first row containing the header token)
    let hr = 0;
    for (let i = 0; i < Math.min(8, grid.length); i++) {
      if ((grid[i] || []).some((c) => normalizeHeader(c) === headerToken)) {
        hr = i;
        break;
      }
    }
    const headers = (grid[hr] || []).map((h) => normalizeHeader(h));
    const newCol = findNewColumn(headers);

    // canonical field -> column index (first match wins)
    const fieldCol = {};
    for (const [src, canon] of Object.entries(fields)) {
      if (canon in fieldCol) continue;
      const idx = findColumn(headers, src);
      if (idx >= 0) fieldCol[canon] = idx;
    }

    let count = 0;
    for (let r = hr + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const get = (canon) => (canon in fieldCol ? row[fieldCol[canon]] : null);

      const voucher = cellToString(get('po_number'));
      const code = cellToString(get('sku'));
      if (!voucher || !code) continue;          // need both identity fields
      if (!hasDigit(voucher)) continue;          // skip title/total rows
      if (skipFirst.includes(String(row[0] ?? '').toLowerCase())) continue;

      const obj = {
        brand: mapping.brand || null,
        region: cfg.region || null,
        voucher_type: cfg.voucher_type || null,
      };
      for (const canon of Object.keys(fieldCol)) {
        const val = DATE_FIELDS.has(canon) ? coerceDate(get(canon)) : cellToString(get(canon));
        if (val !== null) obj[canon] = val;
      }

      // normalize free-text status into factory_status
      const raw = obj.raw_status_text || obj.remarks;
      if (raw) {
        const t = normalizeHeader(raw);
        for (const [k, v] of Object.entries(statusMap)) {
          if (t.includes(normalizeHeader(k))) {
            obj.factory_status = v;
            break;
          }
        }
      }

      // apply template defaults (e.g. Tangem -> available) without overriding
      for (const [k, v] of Object.entries(defaults)) {
        if (obj[k] === undefined || obj[k] === null) obj[k] = v;
      }

      // "New" item flag from a New/Tag column, if present
      if (newCol >= 0) {
        const nf = parseNewFlag(row[newCol]);
        if (nf !== undefined) obj.is_new = nf;
      }

      rows.push(obj);
      count++;
    }
    report.perSheet[sheetName] = count;
  }

  return { rows, report };
}

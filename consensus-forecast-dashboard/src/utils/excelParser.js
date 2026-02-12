import * as XLSX from 'xlsx';

/**
 * Normalize a model name: trim whitespace, uppercase
 */
function normalizeModel(name) {
  if (!name) return '';
  return String(name).trim().toUpperCase();
}

/**
 * Parse a date value from Excel into a YYYY-MM string.
 * Handles: Excel serial dates, "Nov-24", "2024-11", "11/1/2024", Date objects, etc.
 */
function parseMonth(val) {
  if (val == null || val === '') return null;

  // If it's an Excel serial number
  if (typeof val === 'number') {
    const date = XLSX.SSF.parse_date_code(val);
    if (date) {
      const y = date.y;
      const m = String(date.m).padStart(2, '0');
      return `${y}-${m}`;
    }
  }

  const str = String(val).trim();

  // YYYY-MM
  const isoMatch = str.match(/^(\d{4})-(\d{1,2})$/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2].padStart(2, '0')}`;

  // MM/DD/YYYY or M/D/YYYY
  const mdyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdyMatch) return `${mdyMatch[3]}-${mdyMatch[1].padStart(2, '0')}`;

  // Mon-YY or Mon-YYYY (e.g., "Nov-24", "Nov-2024")
  const monthNames = {
    jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
    jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
  };
  const monYrMatch = str.match(/^([A-Za-z]{3})-(\d{2,4})$/);
  if (monYrMatch) {
    const mon = monthNames[monYrMatch[1].toLowerCase()];
    let yr = monYrMatch[2];
    if (yr.length === 2) yr = (parseInt(yr) > 50 ? '19' : '20') + yr;
    if (mon) return `${yr}-${mon}`;
  }

  // Try native Date parse as last resort
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }

  return null;
}

/**
 * Auto-detect header row: find the first row that contains multiple expected column keywords.
 */
function findHeaderRow(data, expectedColumns) {
  const lcExpected = expectedColumns.map(c => c.toLowerCase());
  for (let i = 0; i < Math.min(data.length, 10); i++) {
    const row = data[i];
    if (!row) continue;
    const cells = row.map(c => String(c || '').toLowerCase().trim());
    const matches = lcExpected.filter(exp => cells.some(cell => cell.includes(exp)));
    if (matches.length >= 2) return i;
  }
  return 0; // default to first row
}

/**
 * Map header names to canonical column names.
 */
function mapHeaders(headers, columnMap) {
  const mapped = {};
  headers.forEach((h, idx) => {
    const lc = String(h || '').toLowerCase().trim().replace(/[\s_]+/g, '_');
    for (const [canonical, aliases] of Object.entries(columnMap)) {
      if (aliases.some(a => lc.includes(a))) {
        mapped[canonical] = idx;
        break;
      }
    }
  });
  return mapped;
}

/**
 * Parse File 1: Field Forecast & Sell-Through Data
 */
export function parseFieldForecast(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const columnMap = {
    model: ['model'],
    account: ['account'],
    month: ['month', 'date', 'period'],
    field_forecast: ['field_forecast', 'field forecast', 'forecast'],
    actual_sell_through: ['actual_sell_through', 'actual sell through', 'sell_through', 'actual'],
    units_sold_mtd: ['units_sold_mtd', 'units sold mtd', 'mtd'],
    promo_weeks: ['promo_weeks', 'promo weeks'],
    promo_discount_pct: ['promo_discount', 'discount_pct', 'discount pct', 'promo_discount_pct']
  };

  const headerIdx = findHeaderRow(data, ['model', 'account', 'month']);
  const headers = data[headerIdx];
  const mapping = mapHeaders(headers, columnMap);
  const rows = [];
  const warnings = [];

  const requiredCols = ['model', 'account', 'month'];
  const missingCols = requiredCols.filter(c => mapping[c] === undefined);
  if (missingCols.length > 0) {
    warnings.push(`File 1 missing columns: ${missingCols.join(', ')}`);
  }

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === '' || c == null)) continue;

    const model = normalizeModel(row[mapping.model]);
    if (!model) continue;

    const month = parseMonth(row[mapping.month]);
    rows.push({
      model,
      account: String(row[mapping.account] || '').trim(),
      month,
      field_forecast: Number(row[mapping.field_forecast]) || 0,
      actual_sell_through: Number(row[mapping.actual_sell_through]) || 0,
      units_sold_mtd: Number(row[mapping.units_sold_mtd]) || 0,
      promo_weeks: Number(row[mapping.promo_weeks]) || 0,
      promo_discount_pct: Number(row[mapping.promo_discount_pct]) || 0,
    });
  }

  const models = [...new Set(rows.map(r => r.model))];
  const months = [...new Set(rows.map(r => r.month).filter(Boolean))].sort();

  return {
    data: rows,
    models,
    dateRange: months.length ? { start: months[0], end: months[months.length - 1] } : null,
    warnings,
    rowCount: rows.length
  };
}

/**
 * Parse File 2: Planning & Inventory Data
 */
export function parsePlanningData(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const columnMap = {
    model: ['model'],
    month: ['month', 'date', 'period'],
    business_plan_target: ['business_plan', 'plan_target', 'business plan'],
    fc3_target: ['fc3'],
    previous_consensus_forecast: ['previous_consensus', 'prev_consensus', 'consensus_forecast', 'previous consensus'],
    inventory_units: ['inventory', 'inventory_units']
  };

  const headerIdx = findHeaderRow(data, ['model', 'month']);
  const headers = data[headerIdx];
  const mapping = mapHeaders(headers, columnMap);
  const rows = [];
  const warnings = [];

  const requiredCols = ['model', 'month'];
  const missingCols = requiredCols.filter(c => mapping[c] === undefined);
  if (missingCols.length > 0) {
    warnings.push(`File 2 missing columns: ${missingCols.join(', ')}`);
  }

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === '' || c == null)) continue;

    const model = normalizeModel(row[mapping.model]);
    if (!model) continue;

    rows.push({
      model,
      month: parseMonth(row[mapping.month]),
      business_plan_target: Number(row[mapping.business_plan_target]) || 0,
      fc3_target: Number(row[mapping.fc3_target]) || 0,
      previous_consensus_forecast: Number(row[mapping.previous_consensus_forecast]) || 0,
      inventory_units: Number(row[mapping.inventory_units]) || 0,
    });
  }

  const models = [...new Set(rows.map(r => r.model))];
  const months = [...new Set(rows.map(r => r.month).filter(Boolean))].sort();

  return {
    data: rows,
    models,
    dateRange: months.length ? { start: months[0], end: months[months.length - 1] } : null,
    warnings,
    rowCount: rows.length
  };
}

/**
 * Parse File 3: Historical Sell-Through (3-Year)
 */
export function parseHistorical(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const columnMap = {
    model: ['model'],
    month: ['month', 'date', 'period'],
    sell_through: ['sell_through', 'sell through', 'units', 'actual']
  };

  const headerIdx = findHeaderRow(data, ['model', 'month']);
  const headers = data[headerIdx];
  const mapping = mapHeaders(headers, columnMap);
  const rows = [];
  const warnings = [];

  const requiredCols = ['model', 'month', 'sell_through'];
  const missingCols = requiredCols.filter(c => mapping[c] === undefined);
  if (missingCols.length > 0) {
    warnings.push(`File 3 missing columns: ${missingCols.join(', ')}`);
  }

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === '' || c == null)) continue;

    const model = normalizeModel(row[mapping.model]);
    if (!model) continue;

    rows.push({
      model,
      month: parseMonth(row[mapping.month]),
      sell_through: Number(row[mapping.sell_through]) || 0,
    });
  }

  const models = [...new Set(rows.map(r => r.model))];
  const months = [...new Set(rows.map(r => r.month).filter(Boolean))].sort();

  return {
    data: rows,
    models,
    dateRange: months.length ? { start: months[0], end: months[months.length - 1] } : null,
    warnings,
    rowCount: rows.length
  };
}

/**
 * Read a File object and return an XLSX workbook
 */
export function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        resolve(workbook);
      } catch (err) {
        reject(new Error(`Failed to parse Excel file: ${err.message}`));
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}

import * as XLSX from 'xlsx';

const MODELS = [
  { id: 'ES-580W', name: 'ES-580W', category: 'Document Scanner', baseVolume: 450 },
  { id: 'ES-C380W', name: 'ES-C380W', category: 'Compact Scanner', baseVolume: 300 },
  { id: 'RR-600W', name: 'RR-600W', category: 'Receipt Scanner', baseVolume: 200 },
  { id: 'RR-400W', name: 'RR-400W', category: 'Receipt Scanner', baseVolume: 180 },
  { id: 'FF-680W', name: 'FF-680W', category: 'Photo Scanner', baseVolume: 350 },
  { id: 'DS-790WN', name: 'DS-790WN', category: 'Network Scanner', baseVolume: 250 },
];

const ACCOUNTS = [
  { name: 'Amazon', share: 0.30, accuracyBias: 0.05 },
  { name: 'Best Buy', share: 0.25, accuracyBias: -0.03 },
  { name: 'Staples', share: 0.15, accuracyBias: 0.08 },
  { name: 'B&H Photo', share: 0.10, accuracyBias: -0.02 },
  { name: 'CDW', share: 0.12, accuracyBias: 0.06 },
  { name: 'Direct/Epson.com', share: 0.08, accuracyBias: -0.04 },
];

/**
 * Generate seasonal multiplier for a given month (1-12)
 */
function seasonality(month) {
  const factors = [0.7, 0.65, 0.8, 0.85, 0.9, 0.85, 0.8, 0.9, 1.0, 1.1, 1.3, 1.5];
  return factors[month - 1];
}

/**
 * Add some randomness
 */
function jitter(value, pct = 0.1) {
  return Math.round(value * (1 + (Math.random() - 0.5) * 2 * pct));
}

/**
 * Generate 36 months of dates from Jan 2023 to Dec 2025
 */
function generateMonths() {
  const months = [];
  for (let y = 2023; y <= 2025; y++) {
    for (let m = 1; m <= 12; m++) {
      months.push(`${y}-${String(m).padStart(2, '0')}`);
    }
  }
  return months;
}

/**
 * Generate File 1: Field Forecast & Sell-Through Data
 */
function generateFile1() {
  const months = generateMonths();
  const rows = [['Model', 'Account', 'Month', 'Field_Forecast', 'Actual_Sell_Through', 'Units_Sold_MTD', 'Promo_Weeks', 'Promo_Discount_Pct']];

  MODELS.forEach(model => {
    ACCOUNTS.forEach(account => {
      months.forEach(monthStr => {
        const [y, m] = monthStr.split('-').map(Number);
        const monthNum = m;
        const yearIdx = y - 2023;

        // Base volume for this account
        const base = model.baseVolume * account.share;
        // Apply seasonality
        const seasonal = base * seasonality(monthNum);
        // Year-over-year growth (~5% per year)
        const grown = seasonal * (1 + yearIdx * 0.05);
        // Actual sell-through with jitter
        const actual = jitter(grown, 0.12);
        // Field forecast with account-specific bias
        const forecast = jitter(actual * (1 + account.accuracyBias), 0.08);

        // Promo weeks: mainly in Q4 and some in summer
        let promoWeeks = 0;
        let promoDiscount = 0;
        if (monthNum === 11 || monthNum === 12) {
          promoWeeks = Math.random() > 0.3 ? Math.floor(Math.random() * 3) + 1 : 0;
          promoDiscount = promoWeeks > 0 ? Math.floor(Math.random() * 15) + 10 : 0;
        } else if (monthNum === 7 || monthNum === 8) {
          promoWeeks = Math.random() > 0.5 ? Math.floor(Math.random() * 2) + 1 : 0;
          promoDiscount = promoWeeks > 0 ? Math.floor(Math.random() * 10) + 5 : 0;
        }

        // Units sold MTD - only meaningful for the last month
        const isCurrent = (y === 2025 && m === 12);
        const mtd = isCurrent ? jitter(actual * 0.6, 0.15) : actual;

        rows.push([model.id, account.name, monthStr, forecast, actual, mtd, promoWeeks, promoDiscount]);
      });
    });
  });

  return rows;
}

/**
 * Generate File 2: Planning & Inventory Data
 */
function generateFile2() {
  const months = generateMonths();
  const rows = [['Model', 'Month', 'Business_Plan_Target', 'FC3_Target', 'Previous_Consensus_Forecast', 'Inventory_Units']];

  MODELS.forEach(model => {
    months.forEach(monthStr => {
      const [y, m] = monthStr.split('-').map(Number);
      const monthNum = m;
      const yearIdx = y - 2023;

      const base = model.baseVolume;
      const seasonal = base * seasonality(monthNum);
      const grown = seasonal * (1 + yearIdx * 0.05);

      const businessPlan = jitter(grown * 1.05, 0.05);
      const fc3 = jitter(grown * 1.02, 0.04);
      const prevConsensus = jitter(grown, 0.06);
      const inventory = jitter(model.baseVolume * 2.5, 0.2);

      rows.push([model.id, monthStr, businessPlan, fc3, prevConsensus, inventory]);
    });
  });

  return rows;
}

/**
 * Generate File 3: Historical Sell-Through
 */
function generateFile3() {
  const months = generateMonths();
  const rows = [['Model', 'Month', 'Sell_Through']];

  MODELS.forEach(model => {
    months.forEach(monthStr => {
      const [y, m] = monthStr.split('-').map(Number);
      const monthNum = m;
      const yearIdx = y - 2023;

      const base = model.baseVolume;
      const seasonal = base * seasonality(monthNum);
      const grown = seasonal * (1 + yearIdx * 0.05);
      const value = jitter(grown, 0.1);

      rows.push([model.id, monthStr, value]);
    });
  });

  return rows;
}

/**
 * Create an XLSX workbook from an array of arrays
 */
function createWorkbook(data, sheetName = 'Sheet1') {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return wb;
}

/**
 * Download a workbook as a file
 */
function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename);
}

/**
 * Generate and download all 3 sample files
 */
export function downloadSampleFiles() {
  const file1Data = generateFile1();
  const file2Data = generateFile2();
  const file3Data = generateFile3();

  downloadWorkbook(createWorkbook(file1Data, 'Field_Forecast'), 'Sample_Field_Forecast.xlsx');
  setTimeout(() => {
    downloadWorkbook(createWorkbook(file2Data, 'Planning_Data'), 'Sample_Planning_Data.xlsx');
  }, 500);
  setTimeout(() => {
    downloadWorkbook(createWorkbook(file3Data, 'Historical'), 'Sample_Historical_Data.xlsx');
  }, 1000);
}

/**
 * Generate sample data in-memory (returns parsed data objects, skipping file I/O)
 */
export function generateSampleDataInMemory() {
  const file1Data = generateFile1();
  const file2Data = generateFile2();
  const file3Data = generateFile3();

  // Convert to the same format as the parser output
  const parseRows = (rows) => {
    const headers = rows[0];
    return rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
  };

  const fieldRows = parseRows(file1Data).map(r => ({
    model: r.Model,
    account: r.Account,
    month: r.Month,
    field_forecast: r.Field_Forecast,
    actual_sell_through: r.Actual_Sell_Through,
    units_sold_mtd: r.Units_Sold_MTD,
    promo_weeks: r.Promo_Weeks,
    promo_discount_pct: r.Promo_Discount_Pct,
  }));

  const planningRows = parseRows(file2Data).map(r => ({
    model: r.Model,
    month: r.Month,
    business_plan_target: r.Business_Plan_Target,
    fc3_target: r.FC3_Target,
    previous_consensus_forecast: r.Previous_Consensus_Forecast,
    inventory_units: r.Inventory_Units,
  }));

  const historicalRows = parseRows(file3Data).map(r => ({
    model: r.Model,
    month: r.Month,
    sell_through: r.Sell_Through,
  }));

  const models = MODELS.map(m => m.id);

  return {
    fieldData: {
      data: fieldRows,
      models,
      dateRange: { start: '2023-01', end: '2025-12' },
      warnings: [],
      rowCount: fieldRows.length,
    },
    planningData: {
      data: planningRows,
      models,
      dateRange: { start: '2023-01', end: '2025-12' },
      warnings: [],
      rowCount: planningRows.length,
    },
    historicalData: {
      data: historicalRows,
      models,
      dateRange: { start: '2023-01', end: '2025-12' },
      warnings: [],
      rowCount: historicalRows.length,
    },
  };
}

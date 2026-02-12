/**
 * Utility: shift a YYYY-MM string by N months
 */
export function shiftMonth(yyyymm, delta) {
  const [y, m] = yyyymm.split('-').map(Number);
  const date = new Date(y, m - 1 + delta, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

/**
 * Format a YYYY-MM string to "Mon-YY" display format
 */
export function formatMonth(yyyymm) {
  if (!yyyymm) return '';
  const [y, m] = yyyymm.split('-').map(Number);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[m - 1]}-${String(y).slice(2)}`;
}

/**
 * Generate an array of N month strings starting from startMonth
 */
export function generateMonthRange(startMonth, count) {
  const months = [];
  for (let i = 0; i < count; i++) {
    months.push(shiftMonth(startMonth, i));
  }
  return months;
}

/**
 * Get the last N months ending at endMonth (inclusive)
 */
export function getLastNMonths(endMonth, n) {
  const months = [];
  for (let i = n - 1; i >= 0; i--) {
    months.push(shiftMonth(endMonth, -i));
  }
  return months;
}

/**
 * Get field forecast total for a model and month (summed across accounts)
 */
export function getFieldForecastTotal(fieldData, model, month) {
  return fieldData
    .filter(r => r.model === model && r.month === month)
    .reduce((sum, r) => sum + r.field_forecast, 0);
}

/**
 * Get field forecast breakdown by account for a model and month
 */
export function getFieldForecastByAccount(fieldData, model, month) {
  return fieldData
    .filter(r => r.model === model && r.month === month)
    .map(r => ({ account: r.account, forecast: r.field_forecast }));
}

/**
 * Get planning data for a model and month
 */
export function getPlanningRow(planningData, model, month) {
  return planningData.find(r => r.model === model && r.month === month) || null;
}

/**
 * Get historical sell-through for a specific model and month
 */
export function getHistoricalSellThrough(historicalData, model, month) {
  const row = historicalData.find(r => r.model === model && r.month === month);
  return row ? row.sell_through : null;
}

/**
 * Get total units sold MTD for a model (sum across accounts)
 */
export function getUnitsSoldMTD(fieldData, model, month) {
  return fieldData
    .filter(r => r.model === model && r.month === month)
    .reduce((sum, r) => sum + r.units_sold_mtd, 0);
}

/**
 * Calculate average monthly sell-through over the last N months
 */
export function getAvgMonthlySellThrough(historicalData, model, endMonth, nMonths) {
  const months = getLastNMonths(endMonth, nMonths);
  const values = months
    .map(m => getHistoricalSellThrough(historicalData, model, m))
    .filter(v => v !== null && v > 0);
  if (values.length === 0) return 0;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

/**
 * Calculate forecast accuracy for an account over a rolling window.
 * Accuracy = 1 - ABS(SUM(Forecast) - SUM(Actual)) / SUM(Actual)
 */
export function calculateAccuracy(fieldData, model, account, endMonth, windowMonths) {
  const months = getLastNMonths(endMonth, windowMonths);
  const rows = fieldData.filter(
    r => r.model === model && r.account === account && months.includes(r.month)
  );
  const totalForecast = rows.reduce((s, r) => s + r.field_forecast, 0);
  const totalActual = rows.reduce((s, r) => s + r.actual_sell_through, 0);
  if (totalActual === 0) return null;
  return 1 - Math.abs(totalForecast - totalActual) / totalActual;
}

/**
 * Calculate accuracy for all accounts for a model
 */
export function calculateAccuracyByAccount(fieldData, model, endMonth) {
  const accounts = [...new Set(
    fieldData.filter(r => r.model === model).map(r => r.account)
  )];

  return accounts.map(account => ({
    account,
    accuracy3: calculateAccuracy(fieldData, model, account, endMonth, 3),
    accuracy6: calculateAccuracy(fieldData, model, account, endMonth, 6),
    accuracy12: calculateAccuracy(fieldData, model, account, endMonth, 12),
  }));
}

/**
 * Calculate weighted average accuracy across accounts
 */
export function calculateWeightedAccuracy(fieldData, model, endMonth, windowMonths) {
  const accounts = [...new Set(
    fieldData.filter(r => r.model === model).map(r => r.account)
  )];

  const months = getLastNMonths(endMonth, windowMonths);
  let totalActual = 0;
  let weightedSum = 0;

  accounts.forEach(account => {
    const rows = fieldData.filter(
      r => r.model === model && r.account === account && months.includes(r.month)
    );
    const actual = rows.reduce((s, r) => s + r.actual_sell_through, 0);
    const acc = calculateAccuracy(fieldData, model, account, endMonth, windowMonths);
    if (acc !== null && actual > 0) {
      weightedSum += acc * actual;
      totalActual += actual;
    }
  });

  return totalActual > 0 ? weightedSum / totalActual : null;
}

/**
 * Calculate YoY variance: (current - lastYear) / lastYear * 100
 */
export function calculateYoYVariance(historicalData, model, currentMonth) {
  const lastYearMonth = shiftMonth(currentMonth, -12);
  const current = getHistoricalSellThrough(historicalData, model, currentMonth);
  const lastYear = getHistoricalSellThrough(historicalData, model, lastYearMonth);

  if (lastYear === null || lastYear === 0 || current === null) return null;
  return ((current - lastYear) / lastYear) * 100;
}

/**
 * Calculate sell-through variance from forecast over a rolling window
 * AVG((Actual - Forecast) / Forecast * 100)
 */
export function calculateSellThroughVariance(fieldData, model, endMonth, windowMonths) {
  const months = getLastNMonths(endMonth, windowMonths);
  const variances = [];

  months.forEach(month => {
    const rows = fieldData.filter(r => r.model === model && r.month === month);
    const totalForecast = rows.reduce((s, r) => s + r.field_forecast, 0);
    const totalActual = rows.reduce((s, r) => s + r.actual_sell_through, 0);
    if (totalForecast > 0) {
      variances.push(((totalActual - totalForecast) / totalForecast) * 100);
    }
  });

  if (variances.length === 0) return null;
  return variances.reduce((a, b) => a + b, 0) / variances.length;
}

/**
 * Get promo data for a model for a specific month (aggregated across accounts)
 */
export function getPromoData(fieldData, model, month) {
  const rows = fieldData.filter(r => r.model === model && r.month === month);
  if (rows.length === 0) return null;

  const totalWeeks = rows.reduce((s, r) => s + r.promo_weeks, 0);
  const avgDiscount = rows.filter(r => r.promo_discount_pct > 0).length > 0
    ? rows.reduce((s, r) => s + r.promo_discount_pct, 0) / rows.filter(r => r.promo_discount_pct > 0).length
    : 0;

  return {
    weeks: Math.round(totalWeeks / rows.length * 10) / 10,
    discountPct: Math.round(avgDiscount * 10) / 10,
    intensity: (totalWeeks / rows.length) * avgDiscount
  };
}

/**
 * Get forecast vs actual trend data for the last N months
 */
export function getForecastVsActualTrend(fieldData, model, endMonth, nMonths) {
  const months = getLastNMonths(endMonth, nMonths);
  return months.map(month => {
    const rows = fieldData.filter(r => r.model === model && r.month === month);
    return {
      month,
      forecast: rows.reduce((s, r) => s + r.field_forecast, 0),
      actual: rows.reduce((s, r) => s + r.actual_sell_through, 0)
    };
  });
}

/**
 * Suggest forecast values for 9 months, based on last year + trend
 */
export function suggestForecast(historicalData, fieldData, planningData, model, startMonth) {
  const months = generateMonthRange(startMonth, 9);
  return months.map(month => {
    const lastYearMonth = shiftMonth(month, -12);
    const lastYearValue = getHistoricalSellThrough(historicalData, model, lastYearMonth);
    const fieldForecast = getFieldForecastTotal(fieldData, model, month);
    const planRow = getPlanningRow(planningData, model, month);
    const prevConsensus = planRow ? planRow.previous_consensus_forecast : 0;

    // Average of available non-zero values
    const values = [lastYearValue, fieldForecast, prevConsensus].filter(v => v && v > 0);
    const suggested = values.length > 0
      ? Math.round(values.reduce((a, b) => a + b, 0) / values.length)
      : 0;

    return { month, suggested };
  });
}

/**
 * Get all unique accounts for a model from field data
 */
export function getAccountsForModel(fieldData, model) {
  return [...new Set(fieldData.filter(r => r.model === model).map(r => r.account))];
}

/**
 * Get forecast vs actual by account for sparkline data
 */
export function getAccountTrend(fieldData, model, account, endMonth, nMonths) {
  const months = getLastNMonths(endMonth, nMonths);
  return months.map(month => {
    const row = fieldData.find(r => r.model === model && r.account === account && r.month === month);
    return {
      month,
      forecast: row ? row.field_forecast : 0,
      actual: row ? row.actual_sell_through : 0
    };
  });
}

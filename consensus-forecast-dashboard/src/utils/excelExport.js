import * as XLSX from 'xlsx';
import { formatMonth } from './calculations';

/**
 * Generate the SAP-ready Consensus Forecast Excel export
 */
export function exportConsensusForecast(forecasts, fieldData, planningData, historicalData, forecastMonth, notes) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Consensus Forecast (9-month rolling)
  const models = Object.keys(forecasts).filter(model => {
    const vals = forecasts[model];
    return vals && Object.values(vals).some(v => v > 0);
  });

  if (models.length === 0) {
    alert('No forecast data to export. Please enter forecasts for at least one model.');
    return;
  }

  // Determine the 9 months
  const monthKeys = models.length > 0 ? Object.keys(forecasts[models[0]]).sort() : [];

  // Sheet 1 data
  const sheet1Header = ['Model', ...monthKeys.map(m => formatMonth(m))];
  const sheet1Data = [sheet1Header];
  models.forEach(model => {
    const row = [model, ...monthKeys.map(m => forecasts[model][m] || 0)];
    sheet1Data.push(row);
  });

  const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
  // Set column widths
  ws1['!cols'] = sheet1Header.map((_, i) => ({ wch: i === 0 ? 14 : 10 }));
  XLSX.utils.book_append_sheet(wb, ws1, 'Consensus Forecast');

  // Sheet 2: Forecast Summary
  const sheet2Header = [
    'Model', 'Current_Month_Forecast', 'Field_Forecast', 'Previous_Consensus',
    'YoY_Change_Pct', 'Inventory', 'Business_Plan_Target', 'FC3_Target', 'Notes'
  ];
  const sheet2Data = [sheet2Header];

  models.forEach(model => {
    const currentForecast = forecasts[model][forecastMonth] || 0;

    // Field forecast total
    const fieldTotal = fieldData
      .filter(r => r.model === model && r.month === forecastMonth)
      .reduce((s, r) => s + r.field_forecast, 0);

    // Planning data
    const planRow = planningData.find(r => r.model === model && r.month === forecastMonth);
    const prevConsensus = planRow ? planRow.previous_consensus_forecast : 0;
    const inventory = planRow ? planRow.inventory_units : 0;
    const bpTarget = planRow ? planRow.business_plan_target : 0;
    const fc3 = planRow ? planRow.fc3_target : 0;

    // YoY
    const [y, m] = forecastMonth.split('-').map(Number);
    const lastYearMonth = `${y - 1}-${String(m).padStart(2, '0')}`;
    const currentST = historicalData.find(r => r.model === model && r.month === forecastMonth);
    const lastYearST = historicalData.find(r => r.model === model && r.month === lastYearMonth);
    let yoy = '';
    if (currentST && lastYearST && lastYearST.sell_through > 0) {
      yoy = Math.round(((currentST.sell_through - lastYearST.sell_through) / lastYearST.sell_through) * 100);
    }

    const modelNotes = (notes && notes[model]) || '';

    sheet2Data.push([model, currentForecast, fieldTotal, prevConsensus, yoy, inventory, bpTarget, fc3, modelNotes]);
  });

  const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);
  ws2['!cols'] = sheet2Header.map((h, i) => ({ wch: i === 0 ? 14 : i === 8 ? 30 : 18 }));
  XLSX.utils.book_append_sheet(wb, ws2, 'Forecast Summary');

  // Sheet 3: Accuracy Reference
  const allModels = [...new Set(fieldData.map(r => r.model))];
  const allAccounts = [...new Set(fieldData.map(r => r.account))];

  const sheet3Header = ['Model', 'Account', 'Last_3Mo_Accuracy', 'Last_6Mo_Accuracy', 'Last_12Mo_Accuracy'];
  const sheet3Data = [sheet3Header];

  allModels.forEach(model => {
    allAccounts.forEach(account => {
      const calcAcc = (windowMonths) => {
        const endDate = new Date();
        const months = [];
        for (let i = windowMonths - 1; i >= 0; i--) {
          const d = new Date(endDate.getFullYear(), endDate.getMonth() - i, 1);
          months.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
        }
        const rows = fieldData.filter(
          r => r.model === model && r.account === account && months.includes(r.month)
        );
        const totalForecast = rows.reduce((s, r) => s + r.field_forecast, 0);
        const totalActual = rows.reduce((s, r) => s + r.actual_sell_through, 0);
        if (totalActual === 0) return '';
        return Math.round((1 - Math.abs(totalForecast - totalActual) / totalActual) * 100) + '%';
      };

      sheet3Data.push([model, account, calcAcc(3), calcAcc(6), calcAcc(12)]);
    });
  });

  const ws3 = XLSX.utils.aoa_to_sheet(sheet3Data);
  ws3['!cols'] = sheet3Header.map((_, i) => ({ wch: i < 2 ? 18 : 16 }));
  XLSX.utils.book_append_sheet(wb, ws3, 'Accuracy Reference');

  // Generate filename and download
  const timestamp = new Date().toISOString().slice(0, 16).replace(/[:-]/g, '');
  const filename = `Consensus_Forecast_${forecastMonth}_${timestamp}.xlsx`;
  XLSX.writeFile(wb, filename);
}

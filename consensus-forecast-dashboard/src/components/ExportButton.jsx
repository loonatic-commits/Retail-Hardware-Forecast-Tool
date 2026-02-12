import React from 'react';
import { exportConsensusForecast } from '../utils/excelExport';

export default function ExportButton({ forecasts, fieldData, planningData, historicalData, forecastMonth, notes }) {
  const modelsWithForecasts = Object.keys(forecasts || {}).filter(model => {
    const vals = forecasts[model];
    return vals && Object.values(vals).some(v => v > 0);
  });

  const handleExport = () => {
    exportConsensusForecast(
      forecasts,
      fieldData,
      planningData,
      historicalData,
      forecastMonth,
      notes
    );
  };

  return (
    <button
      onClick={handleExport}
      disabled={modelsWithForecasts.length === 0}
      className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed shadow-sm"
      title={modelsWithForecasts.length === 0 ? 'Enter forecasts for at least one model' : `Export ${modelsWithForecasts.length} models`}
    >
      <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
      </svg>
      Export All Models ({modelsWithForecasts.length})
    </button>
  );
}

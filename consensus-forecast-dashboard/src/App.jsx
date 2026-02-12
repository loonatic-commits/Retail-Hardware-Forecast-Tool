import React, { useState, useEffect, useCallback, useMemo } from 'react';
import FileUpload from './components/FileUpload';
import ModelSelector from './components/ModelSelector';
import MetricsCards from './components/MetricsCards';
import AccuracyTable from './components/AccuracyTable';
import VarianceAnalysis from './components/VarianceAnalysis';
import TrendChart from './components/TrendChart';
import ForecastInput from './components/ForecastInput';
import ExportButton from './components/ExportButton';
import { generateSampleDataInMemory } from './utils/sampleData';

const STORAGE_KEYS = {
  FIELD_DATA: 'cfd_fieldData',
  PLANNING_DATA: 'cfd_planningData',
  HISTORICAL_DATA: 'cfd_historicalData',
  FORECASTS: 'cfd_forecasts',
  NOTES: 'cfd_notes',
  SELECTED_MODEL: 'cfd_selectedModel',
  FORECAST_MONTH: 'cfd_forecastMonth',
};

function loadFromStorage(key, fallback) {
  try {
    const val = localStorage.getItem(key);
    return val ? JSON.parse(val) : fallback;
  } catch {
    return fallback;
  }
}

function saveToStorage(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // storage full or unavailable
  }
}

function getCurrentMonth() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

export default function App() {
  const [fieldData, setFieldData] = useState(() => loadFromStorage(STORAGE_KEYS.FIELD_DATA, null));
  const [planningData, setPlanningData] = useState(() => loadFromStorage(STORAGE_KEYS.PLANNING_DATA, null));
  const [historicalData, setHistoricalData] = useState(() => loadFromStorage(STORAGE_KEYS.HISTORICAL_DATA, null));
  const [forecasts, setForecasts] = useState(() => loadFromStorage(STORAGE_KEYS.FORECASTS, {}));
  const [notes, setNotes] = useState(() => loadFromStorage(STORAGE_KEYS.NOTES, {}));
  const [selectedModel, setSelectedModel] = useState(() => loadFromStorage(STORAGE_KEYS.SELECTED_MODEL, ''));
  const [forecastMonth, setForecastMonth] = useState(() => loadFromStorage(STORAGE_KEYS.FORECAST_MONTH, getCurrentMonth()));
  const [uploadOpen, setUploadOpen] = useState(false);

  const hasData = fieldData && planningData && historicalData;

  // All models from uploaded data
  const allModels = useMemo(() => {
    if (!hasData) return [];
    const modelSets = [
      fieldData.models || [],
      planningData.models || [],
      historicalData.models || [],
    ];
    const union = [...new Set(modelSets.flat())];
    return union.sort();
  }, [fieldData, planningData, historicalData, hasData]);

  // Auto-select first model if none selected
  useEffect(() => {
    if (allModels.length > 0 && (!selectedModel || !allModels.includes(selectedModel))) {
      setSelectedModel(allModels[0]);
    }
  }, [allModels, selectedModel]);

  // Persist state changes
  useEffect(() => { saveToStorage(STORAGE_KEYS.FIELD_DATA, fieldData); }, [fieldData]);
  useEffect(() => { saveToStorage(STORAGE_KEYS.PLANNING_DATA, planningData); }, [planningData]);
  useEffect(() => { saveToStorage(STORAGE_KEYS.HISTORICAL_DATA, historicalData); }, [historicalData]);
  useEffect(() => { saveToStorage(STORAGE_KEYS.FORECASTS, forecasts); }, [forecasts]);
  useEffect(() => { saveToStorage(STORAGE_KEYS.NOTES, notes); }, [notes]);
  useEffect(() => { saveToStorage(STORAGE_KEYS.SELECTED_MODEL, selectedModel); }, [selectedModel]);
  useEffect(() => { saveToStorage(STORAGE_KEYS.FORECAST_MONTH, forecastMonth); }, [forecastMonth]);

  const handleDataLoaded = useCallback((data) => {
    if (data.fieldData) setFieldData(data.fieldData);
    if (data.planningData) setPlanningData(data.planningData);
    if (data.historicalData) setHistoricalData(data.historicalData);
  }, []);

  const handleForecastChange = useCallback((model, values) => {
    setForecasts(prev => ({ ...prev, [model]: values }));
  }, []);

  const handleNotesChange = useCallback((model, text) => {
    setNotes(prev => ({ ...prev, [model]: text }));
  }, []);

  const handleLoadSampleData = () => {
    const sample = generateSampleDataInMemory();
    handleDataLoaded(sample);
  };

  // Models that have forecast data entered
  const completedModels = useMemo(() => {
    return Object.keys(forecasts).filter(m => {
      const vals = forecasts[m];
      return vals && Object.values(vals).some(v => v > 0);
    });
  }, [forecasts]);

  const handleClearData = () => {
    Object.values(STORAGE_KEYS).forEach(key => localStorage.removeItem(key));
    setFieldData(null);
    setPlanningData(null);
    setHistoricalData(null);
    setForecasts({});
    setNotes({});
    setSelectedModel('');
    setForecastMonth(getCurrentMonth());
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Top Bar */}
      <header className="bg-white border-b border-gray-200 shadow-sm sticky top-0 z-40">
        <div className="max-w-screen-2xl mx-auto px-4 py-3">
          <div className="flex items-center justify-between mb-2">
            <div className="flex items-center gap-3">
              <h1 className="text-lg font-bold text-gray-900">Consensus Forecast Dashboard</h1>
              {hasData && (
                <span className="text-xs text-gray-400 bg-gray-100 px-2 py-0.5 rounded">
                  {allModels.length} models loaded
                </span>
              )}
            </div>
            <div className="flex items-center gap-2">
              {hasData && (
                <ExportButton
                  forecasts={forecasts}
                  fieldData={fieldData?.data || []}
                  planningData={planningData?.data || []}
                  historicalData={historicalData?.data || []}
                  forecastMonth={forecastMonth}
                  notes={notes}
                />
              )}
              <button
                onClick={() => setUploadOpen(true)}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 shadow-sm"
              >
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                </svg>
                Upload Files
              </button>
              {hasData && (
                <button
                  onClick={handleClearData}
                  className="px-3 py-2 text-sm font-medium text-red-600 bg-white border border-red-200 rounded-lg hover:bg-red-50"
                >
                  Clear All
                </button>
              )}
            </div>
          </div>

          {hasData && (
            <ModelSelector
              models={allModels}
              selectedModel={selectedModel}
              onSelectModel={setSelectedModel}
              forecastMonth={forecastMonth}
              onForecastMonthChange={setForecastMonth}
              completedModels={completedModels}
            />
          )}
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-screen-2xl mx-auto px-4 py-6">
        {!hasData ? (
          /* Empty State */
          <div className="flex flex-col items-center justify-center py-20 text-center">
            <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mb-4">
              <svg className="w-8 h-8 text-blue-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
            </div>
            <h2 className="text-xl font-bold text-gray-900 mb-2">Consensus Forecast Dashboard</h2>
            <p className="text-gray-500 mb-6 max-w-md">
              Upload your Excel data files to begin the consensus forecasting process, or load sample data to explore the dashboard.
            </p>

            <div className="flex gap-3 mb-8">
              <button
                onClick={() => setUploadOpen(true)}
                className="px-5 py-2.5 text-sm font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 shadow"
              >
                Upload Excel Files
              </button>
              <button
                onClick={handleLoadSampleData}
                className="px-5 py-2.5 text-sm font-medium text-blue-600 bg-blue-50 rounded-lg hover:bg-blue-100 border border-blue-200"
              >
                Load Sample Data
              </button>
            </div>

            <div className="bg-white rounded-lg border border-gray-200 p-6 max-w-xl text-left">
              <h3 className="text-sm font-semibold text-gray-900 mb-3">Expected File Formats</h3>
              <div className="space-y-3 text-xs">
                <div>
                  <span className="font-semibold text-blue-700">File 1: Field Forecast & Sell-Through</span>
                  <p className="text-gray-500 font-mono mt-0.5">Model, Account, Month, Field_Forecast, Actual_Sell_Through, Units_Sold_MTD, Promo_Weeks, Promo_Discount_Pct</p>
                </div>
                <div>
                  <span className="font-semibold text-blue-700">File 2: Planning & Inventory</span>
                  <p className="text-gray-500 font-mono mt-0.5">Model, Month, Business_Plan_Target, FC3_Target, Previous_Consensus_Forecast, Inventory_Units</p>
                </div>
                <div>
                  <span className="font-semibold text-blue-700">File 3: Historical Sell-Through (3-Year)</span>
                  <p className="text-gray-500 font-mono mt-0.5">Model, Month, Sell_Through</p>
                </div>
              </div>
            </div>
          </div>
        ) : (
          /* Dashboard Content */
          <div className="space-y-6">
            {/* Section 1: Key Metrics */}
            <MetricsCards
              model={selectedModel}
              forecastMonth={forecastMonth}
              fieldData={fieldData?.data || []}
              planningData={planningData?.data || []}
              historicalData={historicalData?.data || []}
            />

            {/* Section 2 & 3: Accuracy + Variance side by side on large screens */}
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
              <AccuracyTable
                model={selectedModel}
                forecastMonth={forecastMonth}
                fieldData={fieldData?.data || []}
              />
              <VarianceAnalysis
                model={selectedModel}
                forecastMonth={forecastMonth}
                fieldData={fieldData?.data || []}
                historicalData={historicalData?.data || []}
              />
            </div>

            {/* Section 4: Historical Trend */}
            <TrendChart
              model={selectedModel}
              forecastMonth={forecastMonth}
              fieldData={fieldData?.data || []}
              planningData={planningData?.data || []}
              historicalData={historicalData?.data || []}
            />

            {/* Section 5: Forecast Input */}
            <ForecastInput
              model={selectedModel}
              forecastMonth={forecastMonth}
              fieldData={fieldData?.data || []}
              planningData={planningData?.data || []}
              historicalData={historicalData?.data || []}
              forecasts={forecasts}
              onForecastChange={handleForecastChange}
              notes={notes}
              onNotesChange={handleNotesChange}
            />
          </div>
        )}
      </main>

      {/* File Upload Modal */}
      <FileUpload
        isOpen={uploadOpen}
        onClose={() => setUploadOpen(false)}
        onDataLoaded={handleDataLoaded}
        currentData={{ fieldData, planningData, historicalData }}
      />
    </div>
  );
}

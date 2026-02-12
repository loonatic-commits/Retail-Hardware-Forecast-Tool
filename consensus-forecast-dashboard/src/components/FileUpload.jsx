import React, { useState, useCallback } from 'react';
import { readExcelFile, parseFieldForecast, parsePlanningData, parseHistorical } from '../utils/excelParser';
import { downloadSampleFiles } from '../utils/sampleData';

const FILE_TYPES = [
  {
    key: 'field',
    label: 'Field Forecast & Sell-Through',
    description: 'Model, Account, Month, Field_Forecast, Actual_Sell_Through, Units_Sold_MTD, Promo_Weeks, Promo_Discount_Pct',
    parser: parseFieldForecast,
  },
  {
    key: 'planning',
    label: 'Planning & Inventory',
    description: 'Model, Month, Business_Plan_Target, FC3_Target, Previous_Consensus_Forecast, Inventory_Units',
    parser: parsePlanningData,
  },
  {
    key: 'historical',
    label: 'Historical Sell-Through (3-Year)',
    description: 'Model, Month, Sell_Through',
    parser: parseHistorical,
  },
];

export default function FileUpload({ onDataLoaded, isOpen, onClose, currentData }) {
  const [fileStates, setFileStates] = useState({
    field: null,
    planning: null,
    historical: null,
  });
  const [parsing, setParsing] = useState({});
  const [errors, setErrors] = useState({});

  const handleFile = useCallback(async (fileType, file) => {
    setParsing(prev => ({ ...prev, [fileType.key]: true }));
    setErrors(prev => ({ ...prev, [fileType.key]: null }));

    try {
      const workbook = await readExcelFile(file);
      const result = fileType.parser(workbook);
      setFileStates(prev => ({ ...prev, [fileType.key]: { ...result, fileName: file.name } }));

      // Check if all files are loaded to trigger data update
      const newStates = { ...fileStates, [fileType.key]: result };
      if (newStates.field && newStates.planning && newStates.historical) {
        onDataLoaded({
          fieldData: newStates.field,
          planningData: newStates.planning,
          historicalData: newStates.historical,
        });
      }
    } catch (err) {
      setErrors(prev => ({ ...prev, [fileType.key]: err.message }));
    } finally {
      setParsing(prev => ({ ...prev, [fileType.key]: false }));
    }
  }, [fileStates, onDataLoaded]);

  const handleApply = () => {
    const data = {};
    if (fileStates.field) data.fieldData = fileStates.field;
    if (fileStates.planning) data.planningData = fileStates.planning;
    if (fileStates.historical) data.historicalData = fileStates.historical;
    if (Object.keys(data).length > 0) {
      onDataLoaded(data);
    }
    onClose();
  };

  const allLoaded = fileStates.field && fileStates.planning && fileStates.historical;

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl shadow-2xl max-w-3xl w-full max-h-[90vh] overflow-y-auto">
        <div className="p-6 border-b border-gray-200 flex justify-between items-center">
          <div>
            <h2 className="text-xl font-bold text-gray-900">Upload Data Files</h2>
            <p className="text-sm text-gray-500 mt-1">Upload 3 Excel files to populate the dashboard</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 p-1">
            <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        <div className="p-6 space-y-6">
          {FILE_TYPES.map((fileType, idx) => {
            const state = fileStates[fileType.key];
            const existing = currentData && currentData[fileType.key === 'field' ? 'fieldData' : fileType.key === 'planning' ? 'planningData' : 'historicalData'];
            const isParsing = parsing[fileType.key];
            const error = errors[fileType.key];

            return (
              <div key={fileType.key} className="border border-gray-200 rounded-lg p-4">
                <div className="flex items-start gap-3">
                  <div className="flex-shrink-0 w-8 h-8 bg-blue-100 text-blue-700 rounded-full flex items-center justify-center font-bold text-sm">
                    {idx + 1}
                  </div>
                  <div className="flex-1">
                    <h3 className="font-semibold text-gray-900">{fileType.label}</h3>
                    <p className="text-xs text-gray-500 mt-0.5 font-mono">{fileType.description}</p>

                    <div className="mt-3">
                      {isParsing ? (
                        <div className="flex items-center gap-2 text-blue-600 text-sm">
                          <svg className="animate-spin h-4 w-4" viewBox="0 0 24 24">
                            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                          </svg>
                          Parsing...
                        </div>
                      ) : (
                        <label className="flex items-center gap-2 cursor-pointer">
                          <span className="inline-flex items-center px-3 py-1.5 border border-gray-300 rounded-md text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 transition">
                            {state ? 'Replace File' : 'Choose File'}
                          </span>
                          <input
                            type="file"
                            accept=".xlsx,.xls,.csv"
                            className="hidden"
                            onChange={(e) => {
                              if (e.target.files[0]) handleFile(fileType, e.target.files[0]);
                            }}
                          />
                          {state && <span className="text-sm text-gray-500">{state.fileName}</span>}
                        </label>
                      )}
                    </div>

                    {error && (
                      <div className="mt-2 text-sm text-red-600 bg-red-50 rounded px-3 py-1.5">
                        {error}
                      </div>
                    )}

                    {state && !error && (
                      <div className="mt-2 text-sm space-y-1">
                        <div className="flex gap-4 text-gray-600">
                          <span>{state.rowCount} rows</span>
                          <span>{state.models.length} models</span>
                          {state.dateRange && (
                            <span>{state.dateRange.start} to {state.dateRange.end}</span>
                          )}
                        </div>
                        {state.warnings.length > 0 && (
                          <div className="text-amber-600">
                            {state.warnings.map((w, i) => <div key={i}>{w}</div>)}
                          </div>
                        )}
                      </div>
                    )}

                    {existing && !state && (
                      <div className="mt-2 text-xs text-green-600">Currently loaded: {existing.rowCount} rows, {existing.models.length} models</div>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        <div className="p-6 border-t border-gray-200 flex justify-between items-center">
          <button
            onClick={() => downloadSampleFiles()}
            className="text-sm text-blue-600 hover:text-blue-800 underline"
          >
            Download sample data files
          </button>
          <div className="flex gap-3">
            <button
              onClick={onClose}
              className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50"
            >
              Cancel
            </button>
            <button
              onClick={handleApply}
              disabled={!fileStates.field && !fileStates.planning && !fileStates.historical}
              className="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {allLoaded ? 'Apply & Close' : 'Apply Changes'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

import React, { useState, useEffect, useMemo } from 'react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import {
  generateMonthRange,
  formatMonth,
  suggestForecast,
  getFieldForecastTotal,
  getHistoricalSellThrough,
  getPlanningRow,
  shiftMonth,
} from '../utils/calculations';

export default function ForecastInput({
  model,
  forecastMonth,
  fieldData,
  planningData,
  historicalData,
  forecasts,
  onForecastChange,
  notes,
  onNotesChange,
}) {
  const months9 = useMemo(
    () => generateMonthRange(forecastMonth, 9),
    [forecastMonth]
  );

  // Get suggestions
  const suggestions = useMemo(
    () => suggestForecast(historicalData, fieldData, planningData, model, forecastMonth),
    [historicalData, fieldData, planningData, model, forecastMonth]
  );

  // Local values from forecasts state
  const currentValues = useMemo(() => {
    const vals = {};
    months9.forEach((m, i) => {
      vals[m] = (forecasts && forecasts[model] && forecasts[model][m] !== undefined)
        ? forecasts[model][m]
        : suggestions[i]?.suggested || 0;
    });
    return vals;
  }, [months9, forecasts, model, suggestions]);

  const [localValues, setLocalValues] = useState(currentValues);

  useEffect(() => {
    setLocalValues(currentValues);
  }, [currentValues]);

  const handleValueChange = (month, value) => {
    const numVal = parseInt(value) || 0;
    const newVals = { ...localValues, [month]: numVal };
    setLocalValues(newVals);
    onForecastChange(model, newVals);
  };

  const applyAll = (values) => {
    setLocalValues(values);
    onForecastChange(model, values);
  };

  // Quick-fill handlers
  const copyFromFieldForecast = () => {
    const vals = {};
    months9.forEach(m => {
      vals[m] = getFieldForecastTotal(fieldData, model, m) || localValues[m];
    });
    applyAll(vals);
  };

  const copyFromLastYear = (pctAdjust) => {
    const vals = {};
    months9.forEach(m => {
      const ly = getHistoricalSellThrough(historicalData, model, shiftMonth(m, -12));
      vals[m] = ly ? Math.round(ly * (1 + pctAdjust / 100)) : localValues[m];
    });
    applyAll(vals);
  };

  const copyFromPrevConsensus = () => {
    const vals = {};
    months9.forEach(m => {
      const plan = getPlanningRow(planningData, model, m);
      vals[m] = plan ? plan.previous_consensus_forecast : localValues[m];
    });
    applyAll(vals);
  };

  const [flatValue, setFlatValue] = useState('');
  const applyFlat = () => {
    const val = parseInt(flatValue) || 0;
    if (val <= 0) return;
    const vals = {};
    months9.forEach(m => { vals[m] = val; });
    applyAll(vals);
  };

  const [lyPct, setLyPct] = useState(0);

  // Chart data: overlay new forecast vs historical & field forecast
  const chartData = months9.map(m => ({
    month: formatMonth(m),
    forecast: localValues[m] || 0,
    lastYear: getHistoricalSellThrough(historicalData, model, shiftMonth(m, -12)) || 0,
    fieldForecast: getFieldForecastTotal(fieldData, model, m) || 0,
  }));

  return (
    <div className="bg-white rounded-lg border border-gray-200 shadow-sm p-4 space-y-4">
      <div className="flex items-center justify-between">
        <h3 className="text-sm font-semibold text-gray-900">
          9-Month Rolling Forecast — {model}
        </h3>
      </div>

      {/* Quick-fill buttons */}
      <div className="flex flex-wrap gap-2">
        <button
          onClick={copyFromFieldForecast}
          className="px-3 py-1.5 text-xs font-medium bg-blue-50 text-blue-700 rounded-md hover:bg-blue-100 border border-blue-200"
        >
          Copy from Field Forecast
        </button>

        <div className="flex items-center gap-1">
          <button
            onClick={() => copyFromLastYear(lyPct)}
            className="px-3 py-1.5 text-xs font-medium bg-amber-50 text-amber-700 rounded-md hover:bg-amber-100 border border-amber-200"
          >
            Last Year
          </button>
          <div className="flex items-center gap-1">
            <input
              type="number"
              value={lyPct}
              onChange={(e) => setLyPct(Number(e.target.value))}
              className="w-14 px-1.5 py-1 text-xs border border-gray-300 rounded"
              placeholder="%"
            />
            <span className="text-xs text-gray-500">% adj</span>
          </div>
        </div>

        <button
          onClick={copyFromPrevConsensus}
          className="px-3 py-1.5 text-xs font-medium bg-green-50 text-green-700 rounded-md hover:bg-green-100 border border-green-200"
        >
          Copy from Prev Consensus
        </button>

        <div className="flex items-center gap-1">
          <input
            type="number"
            value={flatValue}
            onChange={(e) => setFlatValue(e.target.value)}
            className="w-20 px-1.5 py-1 text-xs border border-gray-300 rounded"
            placeholder="Units"
          />
          <button
            onClick={applyFlat}
            className="px-3 py-1.5 text-xs font-medium bg-gray-50 text-gray-700 rounded-md hover:bg-gray-100 border border-gray-200"
          >
            Flat Fill
          </button>
        </div>
      </div>

      {/* 9-month input grid */}
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr>
              {months9.map(m => (
                <th key={m} className="text-center px-2 py-1 text-xs font-medium text-gray-500 uppercase">
                  {formatMonth(m)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {/* Reference rows */}
            <tr>
              {months9.map(m => {
                const ly = getHistoricalSellThrough(historicalData, model, shiftMonth(m, -12));
                return (
                  <td key={m} className="text-center px-2 py-0.5 text-xs text-gray-400">
                    LY: {ly ? ly.toLocaleString() : '--'}
                  </td>
                );
              })}
            </tr>
            <tr>
              {months9.map(m => {
                const ff = getFieldForecastTotal(fieldData, model, m);
                return (
                  <td key={m} className="text-center px-2 py-0.5 text-xs text-gray-400">
                    FF: {ff ? ff.toLocaleString() : '--'}
                  </td>
                );
              })}
            </tr>
            {/* Input row */}
            <tr>
              {months9.map(m => (
                <td key={m} className="px-1 py-1">
                  <input
                    type="number"
                    value={localValues[m] || ''}
                    onChange={(e) => handleValueChange(m, e.target.value)}
                    className="w-full text-center px-1 py-1.5 text-sm font-semibold border-2 border-blue-200 rounded-md focus:border-blue-500 focus:ring-1 focus:ring-blue-500 bg-blue-50"
                  />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>

      {/* Mini chart */}
      <div className="h-40">
        <ResponsiveContainer width="100%" height="100%">
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis dataKey="month" tick={{ fontSize: 10 }} />
            <YAxis tick={{ fontSize: 10 }} />
            <Tooltip formatter={(val) => val.toLocaleString()} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            <Line
              type="monotone"
              dataKey="forecast"
              stroke="#2563eb"
              name="Your Forecast"
              strokeWidth={2.5}
              dot={{ r: 3 }}
            />
            <Line
              type="monotone"
              dataKey="lastYear"
              stroke="#9ca3af"
              name="Last Year"
              strokeWidth={1.5}
              strokeDasharray="4 4"
              dot={{ r: 2 }}
            />
            <Line
              type="monotone"
              dataKey="fieldForecast"
              stroke="#f59e0b"
              name="Field Forecast"
              strokeWidth={1.5}
              strokeDasharray="4 4"
              dot={{ r: 2 }}
            />
          </LineChart>
        </ResponsiveContainer>
      </div>

      {/* Notes */}
      <div>
        <label className="text-xs font-medium text-gray-500 uppercase">Notes for {model}</label>
        <textarea
          value={(notes && notes[model]) || ''}
          onChange={(e) => onNotesChange(model, e.target.value)}
          placeholder="Optional notes for this model's forecast decision..."
          className="mt-1 w-full px-3 py-2 text-sm border border-gray-300 rounded-md focus:border-blue-500 focus:ring-1 focus:ring-blue-500 resize-none"
          rows={2}
        />
      </div>
    </div>
  );
}

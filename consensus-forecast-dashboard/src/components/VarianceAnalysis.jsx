import React, { useState } from 'react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import {
  calculateYoYVariance,
  calculateSellThroughVariance,
  getPromoData,
  getForecastVsActualTrend,
  shiftMonth,
  formatMonth,
} from '../utils/calculations';

function VarianceBadge({ value, suffix = '%' }) {
  if (value === null || value === undefined) {
    return <span className="text-gray-400">--</span>;
  }
  const formatted = value > 0 ? `+${value.toFixed(1)}${suffix}` : `${value.toFixed(1)}${suffix}`;
  const colorClass = value > 0 ? 'text-green-600' : value < 0 ? 'text-red-600' : 'text-gray-600';
  const arrow = value > 0 ? '\u2191' : value < 0 ? '\u2193' : '';

  return (
    <span className={`font-semibold ${colorClass}`}>
      {arrow} {formatted}
    </span>
  );
}

export default function VarianceAnalysis({ model, forecastMonth, fieldData, historicalData }) {
  const [isExpanded, setIsExpanded] = useState(true);

  if (!model || !forecastMonth) return null;

  // YoY Variance
  const yoyVariance = calculateYoYVariance(historicalData, model, forecastMonth);

  // Sell-through variance from forecast
  const stVariance3 = calculateSellThroughVariance(fieldData, model, forecastMonth, 3);
  const stVariance6 = calculateSellThroughVariance(fieldData, model, forecastMonth, 6);

  // Promo comparison
  const currentPromo = getPromoData(fieldData, model, forecastMonth);
  const lastYearMonth = shiftMonth(forecastMonth, -12);
  const lastYearPromo = getPromoData(fieldData, model, lastYearMonth);

  // Trend data for chart
  const trendData = getForecastVsActualTrend(fieldData, model, forecastMonth, 12).map(d => ({
    ...d,
    monthLabel: formatMonth(d.month),
  }));

  const promoIntensityCurrent = currentPromo ? currentPromo.intensity : 0;
  const promoIntensityLY = lastYearPromo ? lastYearPromo.intensity : 0;
  const promoIntensityDelta = promoIntensityCurrent - promoIntensityLY;

  return (
    <div className="bg-white rounded-lg border border-gray-200 shadow-sm">
      <button
        onClick={() => setIsExpanded(!isExpanded)}
        className="w-full px-4 py-3 flex items-center justify-between text-left hover:bg-gray-50"
      >
        <h3 className="text-sm font-semibold text-gray-900">Variance Analysis</h3>
        <svg
          className={`w-4 h-4 text-gray-400 transition-transform ${isExpanded ? 'rotate-180' : ''}`}
          fill="none" viewBox="0 0 24 24" stroke="currentColor"
        >
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
        </svg>
      </button>

      {isExpanded && (
        <div className="px-4 pb-4 space-y-6">
          {/* Variance cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {/* YoY */}
            <div className="bg-gray-50 rounded-lg p-3">
              <div className="text-xs font-medium text-gray-500 uppercase">Year-over-Year</div>
              <div className="mt-1 text-lg">
                <VarianceBadge value={yoyVariance} />
              </div>
              <div className="text-xs text-gray-500 mt-0.5">
                vs {formatMonth(shiftMonth(forecastMonth, -12))}
              </div>
            </div>

            {/* Sell-through variance */}
            <div className="bg-gray-50 rounded-lg p-3">
              <div className="text-xs font-medium text-gray-500 uppercase">Sell-Through vs Forecast</div>
              <div className="mt-1 flex gap-4">
                <div>
                  <div className="text-xs text-gray-400">3 Mo</div>
                  <VarianceBadge value={stVariance3} />
                </div>
                <div>
                  <div className="text-xs text-gray-400">6 Mo</div>
                  <VarianceBadge value={stVariance6} />
                </div>
              </div>
            </div>

            {/* Promo comparison */}
            <div className="bg-gray-50 rounded-lg p-3">
              <div className="text-xs font-medium text-gray-500 uppercase">Promo Comparison</div>
              <div className="mt-1 text-xs space-y-1">
                <div className="text-gray-700">
                  <span className="font-medium">This {formatMonth(forecastMonth)}:</span>{' '}
                  {currentPromo
                    ? `${currentPromo.weeks} wks at ${currentPromo.discountPct}% off`
                    : 'No promo data'}
                </div>
                <div className="text-gray-500">
                  <span className="font-medium">Last {formatMonth(lastYearMonth)}:</span>{' '}
                  {lastYearPromo
                    ? `${lastYearPromo.weeks} wks at ${lastYearPromo.discountPct}% off`
                    : 'No promo data'}
                </div>
                <div className="text-gray-600 pt-1 border-t border-gray-200">
                  Intensity delta:{' '}
                  <span className={promoIntensityDelta > 0 ? 'text-green-600' : promoIntensityDelta < 0 ? 'text-red-600' : ''}>
                    {promoIntensityDelta > 0 ? '+' : ''}{promoIntensityDelta.toFixed(1)}
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Forecast vs Actual trend chart */}
          <div>
            <h4 className="text-xs font-medium text-gray-500 uppercase mb-2">
              Forecast vs Actual (Last 12 Months)
            </h4>
            <div className="h-48">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={trendData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                  <XAxis dataKey="monthLabel" tick={{ fontSize: 10 }} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip formatter={(val) => val.toLocaleString()} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Line
                    type="monotone"
                    dataKey="forecast"
                    stroke="#93c5fd"
                    name="Field Forecast"
                    strokeWidth={2}
                    dot={{ r: 2 }}
                  />
                  <Line
                    type="monotone"
                    dataKey="actual"
                    stroke="#2563eb"
                    name="Actual"
                    strokeWidth={2}
                    dot={{ r: 2 }}
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

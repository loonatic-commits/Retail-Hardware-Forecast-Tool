import React, { useState } from 'react';
import { BarChart, Bar, ResponsiveContainer } from 'recharts';
import {
  calculateAccuracyByAccount,
  calculateWeightedAccuracy,
  getAccountTrend,
} from '../utils/calculations';

function AccuracyBadge({ value }) {
  if (value === null || value === undefined) {
    return <span className="text-gray-400 text-sm">--</span>;
  }
  const pct = Math.round(value * 100);
  let colorClass = 'bg-red-100 text-red-700';
  if (pct >= 90) colorClass = 'bg-green-100 text-green-700';
  else if (pct >= 75) colorClass = 'bg-yellow-100 text-yellow-700';

  return (
    <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${colorClass}`}>
      {pct}%
    </span>
  );
}

function MiniSparkline({ data }) {
  if (!data || data.length === 0) return null;
  return (
    <div className="w-32 h-8">
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={data} barGap={0} barCategoryGap={1}>
          <Bar dataKey="forecast" fill="#93c5fd" radius={[1, 1, 0, 0]} />
          <Bar dataKey="actual" fill="#3b82f6" radius={[1, 1, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

export default function AccuracyTable({ model, forecastMonth, fieldData }) {
  const [isExpanded, setIsExpanded] = useState(true);

  if (!model || !forecastMonth || !fieldData || fieldData.length === 0) return null;

  const accuracyData = calculateAccuracyByAccount(fieldData, model, forecastMonth);

  const weighted3 = calculateWeightedAccuracy(fieldData, model, forecastMonth, 3);
  const weighted6 = calculateWeightedAccuracy(fieldData, model, forecastMonth, 6);
  const weighted12 = calculateWeightedAccuracy(fieldData, model, forecastMonth, 12);

  return (
    <div className="bg-white rounded-lg border border-gray-200 shadow-sm">
      <button
        onClick={() => setIsExpanded(!isExpanded)}
        className="w-full px-4 py-3 flex items-center justify-between text-left hover:bg-gray-50"
      >
        <h3 className="text-sm font-semibold text-gray-900">Field Forecast Accuracy by Account</h3>
        <svg
          className={`w-4 h-4 text-gray-400 transition-transform ${isExpanded ? 'rotate-180' : ''}`}
          fill="none" viewBox="0 0 24 24" stroke="currentColor"
        >
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
        </svg>
      </button>

      {isExpanded && (
        <div className="px-4 pb-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-gray-200">
                <th className="text-left py-2 text-xs font-medium text-gray-500 uppercase">Account</th>
                <th className="text-center py-2 text-xs font-medium text-gray-500 uppercase">Last 3 Mo</th>
                <th className="text-center py-2 text-xs font-medium text-gray-500 uppercase">Last 6 Mo</th>
                <th className="text-center py-2 text-xs font-medium text-gray-500 uppercase">Last 12 Mo</th>
                <th className="text-center py-2 text-xs font-medium text-gray-500 uppercase">Forecast vs Actual</th>
              </tr>
            </thead>
            <tbody>
              {accuracyData.map(({ account, accuracy3, accuracy6, accuracy12 }) => {
                const trend = getAccountTrend(fieldData, model, account, forecastMonth, 6);
                return (
                  <tr key={account} className="border-b border-gray-100">
                    <td className="py-2 text-gray-800 font-medium">{account}</td>
                    <td className="py-2 text-center"><AccuracyBadge value={accuracy3} /></td>
                    <td className="py-2 text-center"><AccuracyBadge value={accuracy6} /></td>
                    <td className="py-2 text-center"><AccuracyBadge value={accuracy12} /></td>
                    <td className="py-2 flex justify-center"><MiniSparkline data={trend} /></td>
                  </tr>
                );
              })}
              <tr className="border-t-2 border-gray-300 font-semibold">
                <td className="py-2 text-gray-900">Weighted Average</td>
                <td className="py-2 text-center"><AccuracyBadge value={weighted3} /></td>
                <td className="py-2 text-center"><AccuracyBadge value={weighted6} /></td>
                <td className="py-2 text-center"><AccuracyBadge value={weighted12} /></td>
                <td></td>
              </tr>
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

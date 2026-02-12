import React from 'react';
import {
  ComposedChart, Line, Area, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ReferenceLine
} from 'recharts';
import {
  getHistoricalSellThrough,
  getFieldForecastTotal,
  getPlanningRow,
  formatMonth,
  getLastNMonths,
} from '../utils/calculations';

export default function TrendChart({ model, forecastMonth, fieldData, planningData, historicalData }) {
  if (!model || !forecastMonth || !historicalData || historicalData.length === 0) return null;

  // Build 36-month history
  const months36 = getLastNMonths(forecastMonth, 36);

  const chartData = months36.map(month => {
    const sellThrough = getHistoricalSellThrough(historicalData, model, month);
    const fieldForecast = getFieldForecastTotal(fieldData, model, month);
    const planRow = getPlanningRow(planningData, model, month);
    const consensus = planRow ? planRow.previous_consensus_forecast : null;

    // Determine if this month had promos
    const promoRows = fieldData.filter(r => r.model === model && r.month === month && r.promo_weeks > 0);
    const hasPromo = promoRows.length > 0;

    // Only show forecast/consensus for last 12 months
    const last12 = getLastNMonths(forecastMonth, 12);
    const isRecent = last12.includes(month);

    return {
      month,
      monthLabel: formatMonth(month),
      sellThrough: sellThrough || 0,
      fieldForecast: isRecent && fieldForecast > 0 ? fieldForecast : null,
      consensus: isRecent && consensus > 0 ? consensus : null,
      promo: hasPromo ? sellThrough || 0 : null,
    };
  });

  // Business Plan and FC3 for current month
  const currentPlan = getPlanningRow(planningData, model, forecastMonth);
  const bpTarget = currentPlan ? currentPlan.business_plan_target : null;
  const fc3Target = currentPlan ? currentPlan.fc3_target : null;

  return (
    <div className="bg-white rounded-lg border border-gray-200 shadow-sm p-4">
      <h3 className="text-sm font-semibold text-gray-900 mb-3">
        Historical Sell-Through Trend — {model}
      </h3>
      <div className="h-72">
        <ResponsiveContainer width="100%" height="100%">
          <ComposedChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis
              dataKey="monthLabel"
              tick={{ fontSize: 9 }}
              interval={2}
            />
            <YAxis tick={{ fontSize: 10 }} />
            <Tooltip
              formatter={(value, name) => [value ? value.toLocaleString() : '--', name]}
              labelFormatter={(label) => label}
            />
            <Legend wrapperStyle={{ fontSize: 11 }} />

            {/* Promo periods as area fill */}
            <Area
              type="monotone"
              dataKey="promo"
              fill="#fef3c7"
              stroke="none"
              name="Promo Period"
              fillOpacity={0.6}
            />

            {/* Sell-through line */}
            <Line
              type="monotone"
              dataKey="sellThrough"
              stroke="#1e40af"
              name="Sell-Through"
              strokeWidth={2}
              dot={{ r: 1.5 }}
              activeDot={{ r: 4 }}
            />

            {/* Field Forecast overlay */}
            <Line
              type="monotone"
              dataKey="fieldForecast"
              stroke="#f59e0b"
              name="Field Forecast"
              strokeWidth={2}
              strokeDasharray="4 4"
              dot={{ r: 2 }}
              connectNulls={false}
            />

            {/* Consensus overlay */}
            <Line
              type="monotone"
              dataKey="consensus"
              stroke="#10b981"
              name="Prev Consensus"
              strokeWidth={2}
              strokeDasharray="6 3"
              dot={{ r: 2 }}
              connectNulls={false}
            />

            {/* Reference lines for BP and FC3 */}
            {bpTarget && (
              <ReferenceLine
                y={bpTarget}
                stroke="#ef4444"
                strokeDasharray="8 4"
                label={{ value: 'BP', position: 'right', fontSize: 10, fill: '#ef4444' }}
              />
            )}
            {fc3Target && (
              <ReferenceLine
                y={fc3Target}
                stroke="#8b5cf6"
                strokeDasharray="8 4"
                label={{ value: 'FC3', position: 'right', fontSize: 10, fill: '#8b5cf6' }}
              />
            )}
          </ComposedChart>
        </ResponsiveContainer>
      </div>
      <div className="mt-2 flex flex-wrap gap-x-4 gap-y-1 text-xs text-gray-500">
        <span className="flex items-center gap-1">
          <span className="w-3 h-0.5 bg-red-500 inline-block" style={{ borderTop: '2px dashed #ef4444' }}></span>
          BP = Business Plan Target
        </span>
        <span className="flex items-center gap-1">
          <span className="w-3 h-0.5 inline-block" style={{ borderTop: '2px dashed #8b5cf6' }}></span>
          FC3 = FC3 Target
        </span>
        <span className="flex items-center gap-1">
          <span className="w-3 h-3 bg-amber-100 inline-block rounded"></span>
          Shaded = Promo Period
        </span>
      </div>
    </div>
  );
}

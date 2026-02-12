import React, { useState } from 'react';
import {
  getFieldForecastTotal,
  getFieldForecastByAccount,
  getPlanningRow,
  getHistoricalSellThrough,
  getUnitsSoldMTD,
  getAvgMonthlySellThrough,
  shiftMonth,
  formatMonth,
} from '../utils/calculations';

function MetricCard({ title, value, subtitle, detail, expandable, children }) {
  const [expanded, setExpanded] = useState(false);

  return (
    <div
      className={`bg-white rounded-lg border border-gray-200 p-4 shadow-sm ${expandable ? 'cursor-pointer hover:border-blue-300' : ''}`}
      onClick={expandable ? () => setExpanded(!expanded) : undefined}
    >
      <div className="text-xs font-medium text-gray-500 uppercase tracking-wider">{title}</div>
      <div className="mt-1 text-2xl font-bold text-gray-900">
        {value !== null && value !== undefined ? value.toLocaleString() : '--'}
      </div>
      {subtitle && <div className="mt-0.5 text-xs text-gray-500">{subtitle}</div>}
      {detail && <div className="mt-0.5 text-xs text-blue-600">{detail}</div>}
      {expanded && children && (
        <div className="mt-3 pt-3 border-t border-gray-100">
          {children}
        </div>
      )}
      {expandable && (
        <div className="mt-1 text-xs text-gray-400">
          {expanded ? 'Click to collapse' : 'Click to expand'}
        </div>
      )}
    </div>
  );
}

export default function MetricsCards({ model, forecastMonth, fieldData, planningData, historicalData }) {
  if (!model || !forecastMonth) return null;

  // Field Forecast Total
  const fieldTotal = getFieldForecastTotal(fieldData, model, forecastMonth);
  const fieldBreakdown = getFieldForecastByAccount(fieldData, model, forecastMonth);

  // Planning data
  const planRow = getPlanningRow(planningData, model, forecastMonth);
  const prevConsensus = planRow ? planRow.previous_consensus_forecast : null;
  const prevMonthPlan = getPlanningRow(planningData, model, shiftMonth(forecastMonth, -1));
  const prevMonthConsensus = prevMonthPlan ? prevMonthPlan.previous_consensus_forecast : null;
  const consensusChange = (prevConsensus && prevMonthConsensus && prevMonthConsensus > 0)
    ? ((prevConsensus - prevMonthConsensus) / prevMonthConsensus * 100).toFixed(1)
    : null;

  // Historical same month comparisons
  const sameMonth1YA = getHistoricalSellThrough(historicalData, model, shiftMonth(forecastMonth, -12));
  const sameMonth2YA = getHistoricalSellThrough(historicalData, model, shiftMonth(forecastMonth, -24));
  const sameMonth3YA = getHistoricalSellThrough(historicalData, model, shiftMonth(forecastMonth, -36));

  // Inventory
  const inventory = planRow ? planRow.inventory_units : null;
  const avgMonthly = getAvgMonthlySellThrough(historicalData, model, forecastMonth, 6);
  const monthsOfSupply = (inventory && avgMonthly > 0) ? (inventory / avgMonthly).toFixed(1) : null;

  // Units sold MTD
  const mtd = getUnitsSoldMTD(fieldData, model, forecastMonth);
  const mtdPctOfForecast = fieldTotal > 0 ? ((mtd / fieldTotal) * 100).toFixed(0) : null;

  // Business Plan & FC3
  const bpTarget = planRow ? planRow.business_plan_target : null;
  const fc3Target = planRow ? planRow.fc3_target : null;

  return (
    <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3">
      <MetricCard
        title="Field Forecast"
        value={fieldTotal}
        detail={`${fieldBreakdown.length} accounts`}
        expandable={fieldBreakdown.length > 0}
      >
        <div className="space-y-1">
          {fieldBreakdown.map(({ account, forecast }) => (
            <div key={account} className="flex justify-between text-xs">
              <span className="text-gray-600">{account}</span>
              <span className="font-medium text-gray-900">{forecast.toLocaleString()}</span>
            </div>
          ))}
        </div>
      </MetricCard>

      <MetricCard
        title="Prev Consensus"
        value={prevConsensus}
        subtitle={consensusChange !== null
          ? `${consensusChange > 0 ? '+' : ''}${consensusChange}% vs prior month`
          : undefined}
      />

      <MetricCard
        title={`Same Month LY`}
        value={sameMonth1YA}
        subtitle={formatMonth(shiftMonth(forecastMonth, -12))}
      />

      <MetricCard
        title="Same Month 2YA"
        value={sameMonth2YA}
        subtitle={formatMonth(shiftMonth(forecastMonth, -24))}
      />

      <MetricCard
        title="Same Month 3YA"
        value={sameMonth3YA}
        subtitle={formatMonth(shiftMonth(forecastMonth, -36))}
      />

      <MetricCard
        title="Inventory"
        value={inventory}
        subtitle={monthsOfSupply ? `${monthsOfSupply} months of supply` : undefined}
      />

      <MetricCard
        title="Units Sold MTD"
        value={mtd}
        detail={mtdPctOfForecast !== null ? `${mtdPctOfForecast}% of field forecast` : undefined}
      />

      <MetricCard
        title="Business Plan"
        value={bpTarget}
      />

      <MetricCard
        title="FC3 Target"
        value={fc3Target}
      />
    </div>
  );
}

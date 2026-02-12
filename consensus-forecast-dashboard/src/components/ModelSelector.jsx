import React from 'react';

export default function ModelSelector({
  models,
  selectedModel,
  onSelectModel,
  forecastMonth,
  onForecastMonthChange,
  completedModels,
}) {
  const currentIdx = models.indexOf(selectedModel);
  const completedCount = completedModels ? completedModels.length : 0;

  const handlePrev = () => {
    if (currentIdx > 0) onSelectModel(models[currentIdx - 1]);
  };

  const handleNext = () => {
    if (currentIdx < models.length - 1) onSelectModel(models[currentIdx + 1]);
  };

  return (
    <div className="flex items-center gap-4 flex-wrap">
      {/* Model selector */}
      <div className="flex items-center gap-2">
        <button
          onClick={handlePrev}
          disabled={currentIdx <= 0}
          className="p-1.5 rounded-md border border-gray-300 text-gray-500 hover:bg-gray-100 disabled:opacity-30 disabled:cursor-not-allowed"
          title="Previous Model"
        >
          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
          </svg>
        </button>

        <select
          value={selectedModel}
          onChange={(e) => onSelectModel(e.target.value)}
          className="block w-44 rounded-lg border-gray-300 bg-white py-2 px-3 text-sm font-semibold text-gray-900 shadow-sm focus:border-blue-500 focus:ring-blue-500"
        >
          {models.map(m => (
            <option key={m} value={m}>
              {m} {completedModels && completedModels.includes(m) ? '\u2713' : ''}
            </option>
          ))}
        </select>

        <button
          onClick={handleNext}
          disabled={currentIdx >= models.length - 1}
          className="p-1.5 rounded-md border border-gray-300 text-gray-500 hover:bg-gray-100 disabled:opacity-30 disabled:cursor-not-allowed"
          title="Next Model"
        >
          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
          </svg>
        </button>
      </div>

      {/* Forecast month */}
      <div className="flex items-center gap-2">
        <label className="text-sm font-medium text-gray-500">Forecast Month:</label>
        <input
          type="month"
          value={forecastMonth}
          onChange={(e) => onForecastMonthChange(e.target.value)}
          className="rounded-lg border-gray-300 bg-white py-1.5 px-3 text-sm text-gray-900 shadow-sm focus:border-blue-500 focus:ring-blue-500"
        />
      </div>

      {/* Progress */}
      {models.length > 0 && (
        <div className="flex items-center gap-2 ml-auto">
          <div className="text-sm text-gray-500">
            <span className="font-semibold text-gray-900">{completedCount}</span> of{' '}
            <span className="font-semibold text-gray-900">{models.length}</span> models complete
          </div>
          <div className="w-24 h-2 bg-gray-200 rounded-full overflow-hidden">
            <div
              className="h-full bg-blue-600 rounded-full transition-all duration-300"
              style={{ width: `${models.length > 0 ? (completedCount / models.length) * 100 : 0}%` }}
            />
          </div>
        </div>
      )}
    </div>
  );
}

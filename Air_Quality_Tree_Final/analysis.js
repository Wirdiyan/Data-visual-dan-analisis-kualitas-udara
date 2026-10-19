/* =========================================================
   AIR QUALITY ANALYSIS ENGINE
   Standalone browser module exposed as window.AirQualityAnalysis
   Missing code -200 is excluded from statistical analysis.
========================================================= */
(function (global) {
  "use strict";

  const VARIABLES = [
    "CO(GT)", "NMHC(GT)", "C6H6(GT)", "NOx(GT)", "NO2(GT)",
    "T", "RH", "AH"
  ];

  function cleanValue(value) {
    if (value === null || value === undefined || value === "" || value === -200 || String(value).trim() === "-200") return null;
    const n = Number(value);
    return Number.isFinite(n) ? n : null;
  }

  function rowValue(row, variable) {
    if (!row) return null;
    if (Object.prototype.hasOwnProperty.call(row, variable)) return row[variable];
    const key = String(variable).trim().replace(/\s+/g, " ").toLowerCase();
    return row[key];
  }

  function values(dataset, variable) {
    return dataset.map(r => cleanValue(rowValue(r, variable))).filter(v => v !== null);
  }

  function sum(arr) { return arr.reduce((a, b) => a + b, 0); }

  function mean(arr) { return arr.length ? sum(arr) / arr.length : null; }

  function median(arr) {
    if (!arr.length) return null;
    const s = [...arr].sort((a, b) => a - b);
    const m = Math.floor(s.length / 2);
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }

  function percentile(arr, p) {
    if (!arr.length) return null;
    const s = [...arr].sort((a, b) => a - b);
    const pos = (s.length - 1) * p;
    const base = Math.floor(pos);
    const rest = pos - base;
    return s[base + 1] !== undefined ? s[base] + rest * (s[base + 1] - s[base]) : s[base];
  }

  function standardDeviation(arr) {
    if (arr.length < 2) return arr.length ? 0 : null;
    const m = mean(arr);
    return Math.sqrt(sum(arr.map(v => (v - m) ** 2)) / (arr.length - 1));
  }

  function descriptiveStats(dataset, variable) {
    const v = values(dataset, variable);
    if (!v.length) return null;
    const q1 = percentile(v, 0.25);
    const q3 = percentile(v, 0.75);
    const iqr = q3 - q1;
    const lower = q1 - 1.5 * iqr;
    const upper = q3 + 1.5 * iqr;
    const outliers = v.filter(x => x < lower || x > upper);
    return {
      count: v.length,
      missing: dataset.length - v.length,
      mean: mean(v),
      median: median(v),
      min: Math.min(...v),
      max: Math.max(...v),
      std: standardDeviation(v),
      q1, q3, iqr, lower, upper,
      outliers,
      outlierCount: outliers.length
    };
  }

  function pearson(dataset, xVar, yVar) {
    const pairs = [];
    dataset.forEach(row => {
      const x = cleanValue(rowValue(row, xVar));
      const y = cleanValue(rowValue(row, yVar));
      if (x !== null && y !== null) pairs.push([x, y]);
    });
    if (pairs.length < 2) return { r: null, pairs };
    const xs = pairs.map(p => p[0]);
    const ys = pairs.map(p => p[1]);
    const mx = mean(xs), my = mean(ys);
    const numerator = sum(pairs.map(([x, y]) => (x - mx) * (y - my)));
    const dx = Math.sqrt(sum(xs.map(x => (x - mx) ** 2)));
    const dy = Math.sqrt(sum(ys.map(y => (y - my) ** 2)));
    const r = dx && dy ? numerator / (dx * dy) : null;
    return { r, pairs };
  }

  function correlationMatrix(dataset, variables = VARIABLES) {
    const matrix = {};
    variables.forEach(x => {
      matrix[x] = {};
      variables.forEach(y => {
        matrix[x][y] = x === y ? 1 : pearson(dataset, x, y).r;
      });
    });
    return matrix;
  }

  function strongestCorrelation(dataset, variables = VARIABLES) {
    let best = null;
    for (let i = 0; i < variables.length; i++) {
      for (let j = i + 1; j < variables.length; j++) {
        const { r } = pearson(dataset, variables[i], variables[j]);
        if (r === null) continue;
        if (!best || Math.abs(r) > Math.abs(best.r)) best = { x: variables[i], y: variables[j], r };
      }
    }
    return best;
  }

  function timeSeries(dataset, variable, dateKey) {
    return dataset
      .filter(r => (!dateKey || r.dateKey === dateKey) && cleanValue(rowValue(r, variable)) !== null)
      .sort((a, b) => a.__timestamp - b.__timestamp)
      .map(r => ({ time: r.displayTime, value: cleanValue(rowValue(r, variable)), row: r }));
  }

  function countMissing(dataset, variables = VARIABLES) {
    let total = 0;
    dataset.forEach(row => variables.forEach(v => { if (cleanValue(rowValue(row, v)) === null) total++; }));
    return total;
  }

  function pollutionScore(row, dataset) {
    const pollutantVars = ["CO(GT)", "NOx(GT)", "NO2(GT)"];
    const scores = pollutantVars.map(v => {
      const stats = descriptiveStats(dataset, v);
      const value = cleanValue(rowValue(row, v));
      if (!stats || value === null) return 0;
      const lo = percentile(values(dataset, v), 0.05);
      const hi = percentile(values(dataset, v), 0.95);
      return hi > lo ? Math.max(0, Math.min(1, (value - lo) / (hi - lo))) : 0;
    });
    return scores[0] * 0.40 + scores[1] * 0.30 + scores[2] * 0.30;
  }

  function pollutionScoresForDataset(dataset) {
    return dataset.map(row => ({ row, score: pollutionScore(row, dataset) }));
  }

  global.AirQualityAnalysis = {
    VARIABLES,
    cleanValue,
    values,
    mean,
    median,
    percentile,
    standardDeviation,
    descriptiveStats,
    pearson,
    correlationMatrix,
    strongestCorrelation,
    timeSeries,
    countMissing,
    pollutionScore,
    pollutionScoresForDataset
  };
})(window);

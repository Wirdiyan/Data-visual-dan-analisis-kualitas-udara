/* =========================================================
   AIR QUALITY TREE — FINAL APPLICATION
   Excel reader + temporal control + 3D tree + statistical analysis
========================================================= */
"use strict";

let airQualityData = [];
let currentDayData = [];
let selectedRow = null;
let selectedDateKey = "";
let currentIndex = 0;
let analysisState = null;

let scene = null;
let camera = null;
let renderer = null;
let controls = null;
let treeGroup = null;
let trunk = null;
let foliageGroup = null;
let ground = null;
let fogLight = null;

const A = () => window.AirQualityAnalysis;
const $ = id => document.getElementById(id);

const LABELS = {
  "CO(GT)": "CO",
  "NMHC(GT)": "NMHC",
  "C6H6(GT)": "C6H6",
  "NOx(GT)": "NOx",
  "NO2(GT)": "NO₂",
  "T": "Temperature",
  "RH": "Humidity",
  "AH": "Absolute Humidity"
};

const DISPLAY_VARIABLES = ["CO(GT)", "NMHC(GT)", "C6H6(GT)", "NOx(GT)", "NO2(GT)", "T", "RH", "AH"];

/* =========================================================
   INIT
========================================================= */
document.addEventListener("DOMContentLoaded", () => {
  const fileInput = $("excelFile");
  if (!fileInput) return console.error("#excelFile tidak ditemukan.");
  if (typeof XLSX === "undefined") return setDatasetStatus("Library Excel belum tersedia.");
  if (!A()) return setDatasetStatus("Analysis engine belum tersedia.");

  fileInput.addEventListener("change", handleFileUpload);
  $("dateSelect")?.addEventListener("change", handleDateChange);
  $("timeSlider")?.addEventListener("input", handleTimeChange);
  $("corrX")?.addEventListener("change", updateCorrelationAnalysis);
  $("corrY")?.addEventListener("change", updateCorrelationAnalysis);
  $("distVariable")?.addEventListener("change", updateDistributionAnalysis);
  $("timeSeriesVariable")?.addEventListener("change", updateTimeSeriesAnalysis);

  initThreeJS();
  window.addEventListener("resize", resizeThreeJS);
});

/* =========================================================
   EXCEL
========================================================= */
function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  setText("fileName", file.name);
  setDatasetStatus("Membaca dataset...");
  resetDataInterface();

  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    setDatasetStatus("File harus berformat .xlsx atau .xls.");
    return;
  }

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const workbook = XLSX.read(e.target.result, { type: "array", cellDates: false, raw: true });
      const datasetSheet = findDatasetSheet(workbook);
      if (!datasetSheet) throw new Error("Tidak ditemukan sheet yang memiliki kolom Date dan Time.");

      const parsed = parseDatasetRows(datasetSheet.rows, datasetSheet.headers);
      if (!parsed.length) throw new Error("Tidak ada baris Date/Time yang valid.");

      airQualityData = parsed;
      populateDateSelect();
      const dates = getUniqueDateKeys();
      selectedDateKey = dates[0] || "";
      const dateSelect = $("dateSelect");
      dateSelect.disabled = !selectedDateKey;
      dateSelect.value = selectedDateKey;

      setupAnalysisUI();
      runDatasetAnalysis();
      updateSelectedDate();

      setDatasetStatus(`Dataset berhasil dibaca • ${parsed.length.toLocaleString("id-ID")} baris valid • ${dates.length} tanggal`);
      setControlStatus("DATASET LOADED", "ready");
    } catch (error) {
      console.error(error);
      airQualityData = [];
      currentDayData = [];
      selectedRow = null;
      resetDataInterface();
      setDatasetStatus("Gagal membaca dataset.");
      setControlStatus("DATASET ERROR", "error");
      alert(`Gagal membaca dataset:\n\n${error.message || "Kesalahan tidak diketahui."}`);
    }
  };
  reader.onerror = () => {
    setDatasetStatus("Browser gagal membaca file.");
    setControlStatus("DATASET ERROR", "error");
  };
  reader.readAsArrayBuffer(file);
}

function findDatasetSheet(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    for (let i = 0; i < Math.min(rows.length, 100); i++) {
      const headers = (rows[i] || []).map(normalizeHeader);
      if (findHeaderIndex(headers, ["date", "tanggal"]) !== -1 && findHeaderIndex(headers, ["time", "waktu"]) !== -1) {
        return { sheetName, headerRowIndex: i, headers, rows: rows.slice(i + 1) };
      }
    }
  }
  return null;
}

function normalizeHeader(value) {
  return value == null ? "" : String(value).replace(/^\uFEFF/, "").trim().replace(/\s+/g, " ").toLowerCase();
}

function findHeaderIndex(headers, candidates) {
  for (const c of candidates) {
    const i = headers.indexOf(normalizeHeader(c));
    if (i !== -1) return i;
  }
  return -1;
}

function parseDatasetRows(rows, headers) {
  const dateIndex = findHeaderIndex(headers, ["date", "tanggal"]);
  const timeIndex = findHeaderIndex(headers, ["time", "waktu"]);
  const result = [];

  rows.forEach((row, rowIndex) => {
    if (!Array.isArray(row)) return;
    const d = parseExcelDate(row[dateIndex]);
    const t = parseExcelTime(row[timeIndex]);
    if (!d || !t) return;

    const timestamp = Date.UTC(d.year, d.month - 1, d.day, t.hours, t.minutes, t.seconds);
    const parsed = {
      __rowNumber: rowIndex + 2,
      __timestamp: timestamp,
      dateKey: makeDateKey(d.year, d.month, d.day),
      year: d.year, month: d.month, day: d.day,
      hours: t.hours, minutes: t.minutes, seconds: t.seconds,
      displayDate: formatDisplayDateParts(d.year, d.month, d.day),
      displayTime: formatTime(t.hours, t.minutes, t.seconds)
    };

    headers.forEach((header, col) => {
      if (header) parsed[header] = parseNumericValue(row[col]);
    });
    result.push(parsed);
  });

  return result.sort((a, b) => a.__timestamp - b.__timestamp);
}

function parseExcelDate(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  if (typeof value === "number" && Number.isFinite(value)) {
    if (XLSX.SSF?.parse_date_code) {
      const p = XLSX.SSF.parse_date_code(value);
      if (p?.y && p?.m && p?.d && isValidDate(p.y, p.m, p.d)) return { year: p.y, month: p.m, day: p.d };
    }
    const date = new Date(Date.UTC(1899, 11, 30) + Math.round(value * 86400000));
    if (!Number.isNaN(date.getTime())) return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
  }
  let text = String(value).trim().split(/\s+/)[0];
  let m = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) {
    const day = +m[1], month = +m[2], year = +m[3];
    return isValidDate(year, month, day) ? { year, month, day } : null;
  }
  m = text.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (m) {
    const year = +m[1], month = +m[2], day = +m[3];
    return isValidDate(year, month, day) ? { year, month, day } : null;
  }
  const months = { januari:1, jan:1, februari:2, feb:2, maret:3, mar:3, april:4, apr:4, mei:5, may:5, juni:6, jun:6, juli:7, jul:7, agustus:8, agu:8, september:9, sep:9, oktober:10, okt:10, november:11, nov:11, desember:12, des:12 };
  m = text.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (m && months[m[2].toLowerCase()]) {
    const day = +m[1], month = months[m[2].toLowerCase()], year = +m[3];
    return isValidDate(year, month, day) ? { year, month, day } : null;
  }
  return null;
}

function parseExcelTime(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return { hours:value.getHours(), minutes:value.getMinutes(), seconds:value.getSeconds() };
  if (typeof value === "number" && Number.isFinite(value)) {
    const total = Math.round(((value % 1 + 1) % 1) * 86400) % 86400;
    return { hours:Math.floor(total/3600), minutes:Math.floor((total%3600)/60), seconds:total%60 };
  }
  const text = String(value).trim().replace(/\./g, ":").replace(/\s+/g, "");
  const m = text.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (!m) return null;
  const hours = +m[1], minutes = +m[2], seconds = m[3] ? +m[3] : 0;
  return hours <= 23 && minutes <= 59 && seconds <= 59 ? { hours, minutes, seconds } : null;
}

function parseNumericValue(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const text = String(value).trim();
  const direct = Number(text);
  if (Number.isFinite(direct)) return direct;
  const comma = Number(text.replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(comma) ? comma : value;
}

function isValidDate(y, m, d) {
  const date = new Date(Date.UTC(y, m - 1, d));
  return date.getUTCFullYear() === y && date.getUTCMonth() === m - 1 && date.getUTCDate() === d;
}
function makeDateKey(y, m, d) { return `${String(y).padStart(4,"0")}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`; }
function formatDisplayDateParts(y, m, d) { return new Intl.DateTimeFormat("id-ID", { day:"2-digit", month:"long", year:"numeric", timeZone:"UTC" }).format(new Date(Date.UTC(y,m-1,d))); }
function formatTime(h,m,s) { return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:${String(s).padStart(2,"0")}`; }
function getUniqueDateKeys() { return [...new Set(airQualityData.map(r => r.dateKey))].sort(); }

/* =========================================================
   TEMPORAL CONTROL
========================================================= */
function populateDateSelect() {
  const select = $("dateSelect");
  if (!select) return;
  select.innerHTML = "";
  getUniqueDateKeys().forEach(key => {
    const row = airQualityData.find(r => r.dateKey === key);
    const option = document.createElement("option");
    option.value = key;
    option.textContent = row?.displayDate || key;
    select.appendChild(option);
  });
}

function handleDateChange(event) {
  selectedDateKey = event.target.value;
  currentIndex = 0;
  updateSelectedDate();
  updateTimeSeriesAnalysis();
}

function updateSelectedDate() {
  currentDayData = airQualityData.filter(r => r.dateKey === selectedDateKey).sort((a,b) => a.__timestamp - b.__timestamp);
  const slider = $("timeSlider");
  if (!currentDayData.length) {
    slider.disabled = true;
    return;
  }
  slider.disabled = false;
  slider.min = 0;
  slider.max = currentDayData.length - 1;
  slider.value = 0;
  currentIndex = 0;
  updateTimeLabels();
  updateCurrentRow();
  setText("currentDateTime", `${currentDayData[0].displayDate} • ${currentDayData[0].displayTime}`);
}

function handleTimeChange(event) {
  currentIndex = Number(event.target.value);
  updateCurrentRow();
}

function updateTimeLabels() {
  if (!currentDayData.length) return;
  const mid = currentDayData[Math.floor((currentDayData.length - 1) / 2)];
  setText("startTime", currentDayData[0].displayTime);
  setText("middleTime", mid.displayTime);
  setText("endTime", currentDayData[currentDayData.length - 1].displayTime);
}

function updateCurrentRow() {
  selectedRow = currentDayData[currentIndex];
  if (!selectedRow) return;
  setText("currentDateTime", `${selectedRow.displayDate} • ${selectedRow.displayTime}`);
  setText("timeValue", selectedRow.displayTime);
  updateConditionPanel(selectedRow);
  updateAnalysisCards(selectedRow);
  updateTree3D(selectedRow);
  updateTimeSeriesAnalysis();
}

/* =========================================================
   CURRENT CONDITION
========================================================= */
function getDataValue(row, variable) {
  if (!row) return null;
  const value = row[normalizeHeader(variable)];
  return A().cleanValue(value);
}

function formatValue(value) { return value == null ? "--" : Number(value).toFixed(2); }

function getRelativePercentage(value, variable) {
  if (value == null) return 0;
  const max = { "CO(GT)":10, "PT08.S1(CO)":2000, "NMHC(GT)":500, "C6H6(GT)":60, "PT08.S2(NMHC)":2000, "NOx(GT)":2000, "PT08.S3(NOx)":2000, "NO2(GT)":1000, "PT08.S4(NO2)":2500, "PT08.S5(O3)":2500, T:50, RH:100, AH:0.03 }[variable] || 100;
  return Math.max(0, Math.min(100, (Number(value) / max) * 100));
}

function updateConditionPanel(row) {
  const container = $("conditionContent");
  if (!container) return;
  const vars = [
    ["CO(GT)","CO(GT)"],["PT08.S1(CO)","PT08.S1(CO)"],["NMHC(GT)","NMHC(GT)"],["C6H6(GT)","C6H6(GT)"],["PT08.S2(NMHC)","PT08.S2(NMHC)"],["NOx(GT)","NOx(GT)"],["PT08.S3(NOx)","PT08.S3(NOx)"],["NO2(GT)","NO2(GT)"],["PT08.S4(NO2)","PT08.S4(NO2)"],["PT08.S5(O3)","PT08.S5(O3)"],["T","Temperature"],["RH","Humidity"],["AH","Absolute Humidity"]
  ];
  container.innerHTML = "";
  vars.forEach(([key,label]) => {
    const value = getDataValue(row,key);
    const item = document.createElement("div"); item.className = "variable";
    item.innerHTML = `<div class="variable-header"><span>${label}</span><strong>${formatValue(value)}</strong></div><div class="variable-bar"><span style="width:${getRelativePercentage(value,key)}%"></span></div>`;
    container.appendChild(item);
  });
}

function updateAnalysisCards(row) {
  const score = A().pollutionScore(row, airQualityData);
  setText("pollutionIndex", `${Math.round(score * 100)}%`);
  setText("dominantPollutant", getDominantPollutant(row));
  const t = getDataValue(row,"T"), rh = getDataValue(row,"RH");
  setText("environmentStatus", [t == null ? null : `${t.toFixed(1)}°C`, rh == null ? null : `${rh.toFixed(1)}% RH`].filter(Boolean).join(" • ") || "--");
  updateQualityBadge(score);
  updateControlStatus(score);
}

function getDominantPollutant(row) {
  const configs = [["CO","CO(GT)"],["NOx","NOx(GT)"],["NO₂","NO2(GT)"]];
  const scores = configs.map(([name,v]) => {
    const stats = A().descriptiveStats(airQualityData,v); const value = getDataValue(row,v);
    if (!stats || value == null) return {name, score:0};
    const lo = A().percentile(A().values(airQualityData,v),0.05), hi = A().percentile(A().values(airQualityData,v),0.95);
    return {name, score:hi > lo ? Math.max(0,Math.min(1,(value-lo)/(hi-lo))) : 0};
  }).sort((a,b)=>b.score-a.score);
  return scores[0]?.score > 0 ? scores[0].name : "--";
}

function updateQualityBadge(score) {
  const el = $("qualityBadge"); if (!el) return;
  el.className = "quality-badge";
  if (score < .25) { el.classList.add("good"); el.textContent = "GOOD AIR QUALITY"; }
  else if (score < .50) { el.classList.add("medium"); el.textContent = "MODERATE AIR QUALITY"; }
  else if (score < .75) { el.classList.add("bad"); el.textContent = "POOR AIR QUALITY"; }
  else { el.classList.add("bad"); el.textContent = "VERY POOR AIR QUALITY"; }
}
function updateControlStatus(score) {
  const el = $("controlStatus"); if (!el) return;
  el.className = "control-status";
  if (score < .25) { el.classList.add("good"); el.textContent = "GOOD AIR QUALITY"; }
  else if (score < .50) { el.classList.add("medium"); el.textContent = "MODERATE AIR QUALITY"; }
  else if (score < .75) { el.classList.add("bad"); el.textContent = "POOR AIR QUALITY"; }
  else { el.classList.add("bad"); el.textContent = "VERY POOR AIR QUALITY"; }
}

/* =========================================================
   ANALYSIS UI
========================================================= */
function setupAnalysisUI() {
  const selects = [$("corrX"), $("corrY"), $("distVariable"), $("timeSeriesVariable")];
  selects.forEach(s => { if (s) s.innerHTML = ""; });
  DISPLAY_VARIABLES.forEach(v => {
    selects.forEach(s => {
      if (!s) return;
      const o = document.createElement("option"); o.value = v; o.textContent = LABELS[v] || v; s.appendChild(o);
    });
  });
  $("corrX").value = "CO(GT)";
  $("corrY").value = "NOx(GT)";
  $("distVariable").value = "CO(GT)";
  $("timeSeriesVariable").value = "CO(GT)";
}

function runDatasetAnalysis() {
  const vars = A().VARIABLES;
  analysisState = {
    missing: A().countMissing(airQualityData, vars),
    strongest: A().strongestCorrelation(airQualityData, vars),
    stats: Object.fromEntries(vars.map(v => [v, A().descriptiveStats(airQualityData,v)]))
  };
  const outliers = vars.reduce((n,v) => n + (analysisState.stats[v]?.outlierCount || 0), 0);
  setText("statRows", airQualityData.length.toLocaleString("id-ID"));
  setText("statMissing", analysisState.missing.toLocaleString("id-ID"));
  setText("statOutliers", outliers.toLocaleString("id-ID"));
  setText("statStrongest", analysisState.strongest ? `${LABELS[analysisState.strongest.x]} ↔ ${LABELS[analysisState.strongest.y]} (${analysisState.strongest.r.toFixed(2)})` : "--");
  setText("analysisStatus", `${vars.length} variabel dianalisis`);
  updateCorrelationAnalysis();
  updateDistributionAnalysis();
  updateTimeSeriesAnalysis();
  updateInsight();
}

function updateCorrelationAnalysis() {
  if (!airQualityData.length) return;
  const x = $("corrX").value, y = $("corrY").value;
  const result = A().pearson(airQualityData,x,y);
  setText("corrValue", result.r == null ? "r = --" : `r = ${result.r.toFixed(3)}`);
  setText("corrInterpretation", interpretCorrelation(result.r));
  renderScatterPlot(result.pairs, LABELS[x], LABELS[y]);
}

function interpretCorrelation(r) {
  if (r == null) return "Data tidak cukup.";
  const abs = Math.abs(r);
  const strength = abs >= .8 ? "sangat kuat" : abs >= .6 ? "kuat" : abs >= .4 ? "sedang" : abs >= .2 ? "lemah" : "sangat lemah";
  const direction = r > .05 ? "positif" : r < -.05 ? "negatif" : "mendekati nol";
  return `Hubungan ${direction} ${strength}. Korelasi tidak otomatis berarti sebab-akibat.`;
}

function renderScatterPlot(pairs, xLabel, yLabel) {
  const box = $("scatterPlot"); if (!box) return;
  if (!pairs.length) return box.innerHTML = emptyChart("Tidak ada pasangan data valid.");
  const width = 760, height = 300, pad = {l:54,r:18,t:18,b:42};
  const xs = pairs.map(p=>p[0]), ys = pairs.map(p=>p[1]);
  const xmin=Math.min(...xs), xmax=Math.max(...xs), ymin=Math.min(...ys), ymax=Math.max(...ys);
  const sx=v=>pad.l+(v-xmin)/(xmax-xmin||1)*(width-pad.l-pad.r);
  const sy=v=>height-pad.b-(v-ymin)/(ymax-ymin||1)*(height-pad.t-pad.b);
  const points = pairs.map(([x,y])=>`<circle cx="${sx(x).toFixed(1)}" cy="${sy(y).toFixed(1)}" r="2.8" class="scatter-point"><title>${xLabel}: ${formatCompact(x)} • ${yLabel}: ${formatCompact(y)}</title></circle>`).join("");
  const svg = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Scatter plot"><line class="axis" x1="${pad.l}" y1="${height-pad.b}" x2="${width-pad.r}" y2="${height-pad.b}"/><line class="axis" x1="${pad.l}" y1="${pad.t}" x2="${pad.l}" y2="${height-pad.b}"/>${gridLines(width,height,pad,4)}${points}<text class="axis-label" x="${width/2}" y="${height-8}" text-anchor="middle">${escapeHtml(xLabel)}</text><text class="axis-label" transform="translate(14 ${height/2}) rotate(-90)" text-anchor="middle">${escapeHtml(yLabel)}</text></svg>`;
  box.innerHTML = svg;
}

function updateDistributionAnalysis() {
  if (!airQualityData.length) return;
  const variable = $("distVariable").value;
  const stats = A().descriptiveStats(airQualityData,variable);
  if (!stats) { $("distributionPlot").innerHTML = emptyChart("Tidak ada data valid."); $("distributionStats").innerHTML=""; return; }
  renderDistribution(stats, A().values(airQualityData,variable), LABELS[variable]);
  $("distributionStats").innerHTML = `<div><span>Mean</span><strong>${formatCompact(stats.mean)}</strong></div><div><span>Median</span><strong>${formatCompact(stats.median)}</strong></div><div><span>Std. Dev.</span><strong>${formatCompact(stats.std)}</strong></div><div><span>Min</span><strong>${formatCompact(stats.min)}</strong></div><div><span>Max</span><strong>${formatCompact(stats.max)}</strong></div><div><span>Q1 / Q3</span><strong>${formatCompact(stats.q1)} / ${formatCompact(stats.q3)}</strong></div><div><span>IQR</span><strong>${formatCompact(stats.iqr)}</strong></div><div><span>Outliers</span><strong>${stats.outlierCount}</strong></div>`;
}

function renderDistribution(stats, values, label) {
  const box=$("distributionPlot"); if(!box)return;
  const width=760,height=300,pad={l:54,r:20,t:28,b:42};
  const min=stats.min,max=stats.max;
  const sx=v=>pad.l+(v-min)/(max-min||1)*(width-pad.l-pad.r);
  const bins=Math.min(14,Math.max(6,Math.ceil(Math.sqrt(values.length))));
  const counts=new Array(bins).fill(0);
  values.forEach(v=>{const i=Math.min(bins-1,Math.floor((v-min)/(max-min||1)*bins));counts[i]++;});
  const maxCount=Math.max(...counts,1), bw=(width-pad.l-pad.r)/bins;
  const bars=counts.map((c,i)=>{const h=(c/maxCount)*(height-pad.t-pad.b); return `<rect class="hist-bar" x="${pad.l+i*bw+2}" y="${height-pad.b-h}" width="${Math.max(1,bw-4)}" height="${h}" rx="3"><title>${c} observasi</title></rect>`;}).join("");
  const q1=sx(stats.q1), med=sx(stats.median), q3=sx(stats.q3), low=sx(Math.max(min,stats.lower)), high=sx(Math.min(max,stats.upper));
  const boxY=height-82, boxH=30;
  const svg=`<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Distribution and outlier chart">${gridLines(width,height,pad,4)}${bars}<line class="marker" x1="${q1}" y1="${boxY}" x2="${q1}" y2="${boxY+boxH}"/><rect class="boxplot" x="${q1}" y="${boxY}" width="${Math.max(2,q3-q1)}" height="${boxH}" rx="5"/><line class="median" x1="${med}" y1="${boxY}" x2="${med}" y2="${boxY+boxH}"/><line class="whisker" x1="${low}" y1="${boxY+boxH/2}" x2="${q1}" y2="${boxY+boxH/2}"/><line class="whisker" x1="${q3}" y1="${boxY+boxH/2}" x2="${high}" y2="${boxY+boxH/2}"/>${values.filter(v=>v<stats.lower||v>stats.upper).slice(0,80).map(v=>`<circle class="outlier" cx="${sx(v)}" cy="${boxY+boxH/2}" r="3"><title>Outlier: ${formatCompact(v)}</title></circle>`).join("")}<text class="axis-label" x="${width/2}" y="${height-8}" text-anchor="middle">${escapeHtml(label)} • histogram + box plot</text></svg>`;
  box.innerHTML=svg;
}

function updateTimeSeriesAnalysis() {
  if (!airQualityData.length) return;
  const variable=$("timeSeriesVariable").value;
  const series=A().timeSeries(airQualityData,variable,selectedDateKey);
  setText("timeSeriesDescription", `${LABELS[variable] || variable} dari waktu ke waktu pada ${selectedDateKey ? formatDateKey(selectedDateKey) : "tanggal aktif"}.`);
  renderTimeSeries(series,LABELS[variable]||variable);
}

function renderTimeSeries(series,label) {
  const box=$("timeSeriesPlot"); if(!box)return;
  if(!series.length)return box.innerHTML=emptyChart("Tidak ada data valid pada tanggal aktif.");
  const width=1100,height=330,pad={l:58,r:24,t:24,b:48};
  const ys=series.map(p=>p.value), min=Math.min(...ys), max=Math.max(...ys);
  const sx=i=>pad.l+i/(Math.max(series.length-1,1))*(width-pad.l-pad.r);
  const sy=v=>height-pad.b-(v-min)/(max-min||1)*(height-pad.t-pad.b);
  const path=series.map((p,i)=>`${i?"L":"M"}${sx(i).toFixed(1)},${sy(p.value).toFixed(1)}`).join(" ");
  const dots=series.map((p,i)=>`<circle class="line-point" cx="${sx(i).toFixed(1)}" cy="${sy(p.value).toFixed(1)}" r="2.5"><title>${p.time} • ${formatCompact(p.value)}</title></circle>`).join("");
  const labels=series.filter((_,i)=>i===0||i===Math.floor(series.length/2)||i===series.length-1).map(p=>p.time.slice(0,5));
  const svg=`<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Time series">${gridLines(width,height,pad,5)}<path class="series-line" d="${path}"/>${dots}<text class="axis-label" x="${width/2}" y="${height-10}" text-anchor="middle">Waktu</text><text class="axis-label" transform="translate(14 ${height/2}) rotate(-90)" text-anchor="middle">${escapeHtml(label)}</text><text class="axis-tick" x="${pad.l}" y="${height-27}">${labels[0]||""}</text><text class="axis-tick" x="${width/2}" y="${height-27}" text-anchor="middle">${labels[1]||""}</text><text class="axis-tick" x="${width-pad.r}" y="${height-27}" text-anchor="end">${labels[2]||""}</text></svg>`;
  box.innerHTML=svg;
}

function updateInsight() {
  if (!analysisState) return;
  const s=analysisState.strongest;
  const d=analysisState.stats["CO(GT)"];
  const out=d?.outlierCount||0;
  const corr=s ? `${LABELS[s.x]} dan ${LABELS[s.y]} memiliki korelasi ${s.r>=0?"positif":"negatif"} paling kuat (${s.r.toFixed(3)}).` : "Belum ada pasangan variabel yang cukup untuk korelasi.";
  $("analysisInsight").innerHTML=`<strong>💡 Insight:</strong> ${corr} Pada CO terdapat <strong>${out}</strong> outlier berdasarkan aturan IQR. Gunakan hasil ini sebagai dasar interpretasi, bukan sebagai bukti hubungan sebab-akibat.`;
}

function gridLines(width,height,pad,n) {
  let s="";
  for(let i=1;i<n;i++){const y=pad.t+i/n*(height-pad.t-pad.b);s+=`<line class="grid-line" x1="${pad.l}" y1="${y}" x2="${width-pad.r}" y2="${y}"/>`;}
  return s;
}
function emptyChart(text){return `<div class="empty-chart">${escapeHtml(text)}</div>`;}
function escapeHtml(s){return String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]));}
function formatCompact(v){ if(v==null||!Number.isFinite(Number(v)))return "--"; const n=Number(v); return Math.abs(n)>=1000?n.toLocaleString("id-ID",{maximumFractionDigits:1}):n.toLocaleString("id-ID",{maximumFractionDigits:3}); }
function formatDateKey(key){const [y,m,d]=key.split("-").map(Number);return formatDisplayDateParts(y,m,d);}

/* =========================================================
   THREE.JS TREE
========================================================= */
function initThreeJS(){
  const container=$("treeContainer");
  if(!container||typeof THREE==="undefined")return;
  scene=new THREE.Scene();
  scene.background=new THREE.Color(0x06120d);
  scene.fog=new THREE.Fog(0x06120d,12,34);
  camera=new THREE.PerspectiveCamera(45,container.clientWidth/Math.max(container.clientHeight,1),0.1,1000);
  camera.position.set(0,5.4,14);
  renderer=new THREE.WebGLRenderer({antialias:true,alpha:true});
  renderer.setPixelRatio(Math.min(window.devicePixelRatio||1,2));
  renderer.setSize(container.clientWidth,container.clientHeight);
  renderer.shadowMap.enabled=true;
  container.appendChild(renderer.domElement);

  if(typeof THREE.OrbitControls !== "undefined"){
    controls=new THREE.OrbitControls(camera,renderer.domElement);
    controls.enableDamping=true; controls.enablePan=false; controls.minDistance=8; controls.maxDistance=22;
    controls.target.set(0,4.2,0);
  }
  scene.add(new THREE.AmbientLight(0xffffff,1.6));
  const light=new THREE.DirectionalLight(0xffffff,2.1); light.position.set(5,11,8); light.castShadow=true; scene.add(light);
  fogLight=new THREE.PointLight(0x8bd8a4,0.8,20); fogLight.position.set(-4,7,2); scene.add(fogLight);
  createTree();
  const geo=new THREE.CylinderGeometry(4,4,.25,64), mat=new THREE.MeshStandardMaterial({color:0x132c20,roughness:.9});
  ground=new THREE.Mesh(geo,mat); ground.position.y=-.1; ground.receiveShadow=true; scene.add(ground);
  $("treePlaceholder")?.style.setProperty("display","none");
  animateTree();
}

function createTree(){
  treeGroup=new THREE.Group(); scene.add(treeGroup);
  foliageGroup=new THREE.Group(); treeGroup.add(foliageGroup);
}

function updateTree3D(row){
  if(!treeGroup||!foliageGroup)return;
  const score=A().pollutionScore(row,airQualityData);
  const co=getDataValue(row,"CO(GT)"), nox=getDataValue(row,"NOx(GT)"), no2=getDataValue(row,"NO2(GT)"), t=getDataValue(row,"T"), rh=getDataValue(row,"RH");
  const normalized=(variable,value)=>{if(value==null)return .5;const st=A().descriptiveStats(airQualityData,variable);if(!st)return .5;const lo=A().percentile(A().values(airQualityData,variable),.05),hi=A().percentile(A().values(airQualityData,variable),.95);return hi>lo?Math.max(0,Math.min(1,(value-lo)/(hi-lo))):.5;};
  const coN=normalized("CO(GT)",co), noxN=normalized("NOx(GT)",nox), no2N=normalized("NO2(GT)",no2), tempN=normalized("T",t), rhN=normalized("RH",rh);
  foliageGroup.traverse(obj=>{if(obj.isMesh){obj.geometry.dispose();if(Array.isArray(obj.material))obj.material.forEach(m=>m.dispose());else obj.material.dispose();}});
  foliageGroup.clear();
  const leafColor=new THREE.Color().setHSL(.37-no2N*.34,.52+no2N*.18,.46-no2N*.06);
  const woodMat=new THREE.MeshStandardMaterial({color:0x62432d,roughness:.92});
  const leafMat=new THREE.MeshStandardMaterial({color:leafColor,roughness:.82});
  const leafGeometry=new THREE.IcosahedronGeometry(.12+(1-coN)*.12,1);
  const branchSides=noxN>.68?3:2;
  const branchAngle=.34+noxN*.30;
  const treeHeight=7.1+tempN*2.1;
  const leafDensity=.35+(1-coN)*.65;
  const up=new THREE.Vector3(0,1,0);
  const addBranch=(start,direction,length,radius,depth,phase)=>{
    const end=start.clone().addScaledVector(direction,length);
    const geo=new THREE.CylinderGeometry(Math.max(.018,radius*.68),radius,length,7);
    const segment=new THREE.Mesh(geo,woodMat);
    segment.position.copy(start).add(end).multiplyScalar(.5);
    segment.quaternion.setFromUnitVectors(up,direction);
    segment.castShadow=true; foliageGroup.add(segment);
    if(depth===0){
      const leafCount=1+Math.round(leafDensity*2);
      for(let i=0;i<leafCount;i++){
        if(((i+phase)%3)/3>leafDensity)continue;
        const leaf=new THREE.Mesh(leafGeometry,leafMat);
        const spread=.12+i*.08;
        leaf.position.copy(end).add(new THREE.Vector3(Math.cos(phase+i*2.4)*spread,.04+i%2*.08,Math.sin(phase+i*2.4)*spread));
        leaf.scale.set(1,.78,1); leaf.castShadow=true; foliageGroup.add(leaf);
      }
      return;
    }
    const tilt=branchAngle+(depth===3?-.05:0);
    const baseAngle=phase+depth*1.7;
    for(let i=0;i<branchSides;i++){
      const a=baseAngle+i*Math.PI*2/branchSides;
      const childDir=new THREE.Vector3(Math.sin(tilt),Math.cos(tilt),0).applyAxisAngle(up,a).normalize();
      addBranch(end,childDir,length*(.69+noxN*.04),radius*.68,depth-1,a);
    }
    if(depth>1){
      const topDir=direction.clone().add(new THREE.Vector3(Math.sin(baseAngle)*.18,.18,Math.cos(baseAngle)*.18)).normalize();
      addBranch(end,topDir,length*.72,radius*.7,depth-1,baseAngle+1.1);
    }
  };
  // A tapered trunk and recursively split limbs form the data-driven fractal crown.
  addBranch(new THREE.Vector3(0,0,0),up,treeHeight*.32,.34,3,0);
  const crownBase=treeHeight*.32;
  for(let i=0;i<4;i++){
    const angle=i*Math.PI/2+.35;
    const dir=new THREE.Vector3(Math.cos(angle)*.45,.89,Math.sin(angle)*.45).normalize();
    addBranch(new THREE.Vector3(0,crownBase,0),dir,treeHeight*.25,.22,3,angle);
  }
  const top=new THREE.Vector3(0,crownBase+treeHeight*.25,0);
  addBranch(top,up,treeHeight*.2,.18,3,1.2);
  treeGroup.scale.set(1,.86+tempN*.28,1);
  foliageGroup.rotation.z=(rhN-.5)*.08;
  scene.fog.near=13-rhN*5; scene.fog.far=35-rhN*11;
  fogLight.intensity=.35+rhN*.8;
}

function animateTree(){requestAnimationFrame(animateTree);if(treeGroup)treeGroup.rotation.y+=.0012;if(controls)controls.update();if(renderer&&scene&&camera)renderer.render(scene,camera);}
function resizeThreeJS(){const c=$("treeContainer");if(!c||!camera||!renderer)return;const w=Math.max(c.clientWidth,1),h=Math.max(c.clientHeight,1);camera.aspect=w/h;camera.updateProjectionMatrix();renderer.setSize(w,h);}

/* =========================================================
   RESET / DOM
========================================================= */
function resetDataInterface(){
  const ds=$("dateSelect"); if(ds){ds.disabled=true;ds.innerHTML='<option value="">Pilih tanggal</option>';}
  const slider=$("timeSlider"); if(slider){slider.disabled=true;slider.min=0;slider.max=0;slider.value=0;}
  ["currentDateTime","timeValue","pollutionIndex","dominantPollutant","environmentStatus","statRows","statMissing","statStrongest","statOutliers"].forEach(id=>setText(id,"--"));
  setText("analysisStatus","Menunggu dataset");
  $("conditionContent").innerHTML='<div class="empty-message">Pilih dataset terlebih dahulu.</div>';
  $("qualityBadge").className="quality-badge good";$("qualityBadge").textContent="WAITING FOR DATA";
  setControlStatus("WAITING FOR DATA","waiting");
  ["scatterPlot","distributionPlot","timeSeriesPlot"].forEach(id=>$(id)&&( $(id).innerHTML=emptyChart("Upload dataset untuk melihat analisis.") ));
  $("distributionStats").innerHTML="";
  $("analysisInsight").textContent="Upload dataset untuk menghasilkan insight analisis.";
}
function setDatasetStatus(text){setText("datasetStatus",text);}
function setControlStatus(text,state){const el=$("controlStatus");if(!el)return;el.className="control-status"+(state?` ${state}`:"");el.textContent=text;}
function setText(id,text){const el=$(id);if(el)el.textContent=text;}

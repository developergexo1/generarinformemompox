// server.js
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { ChartJSNodeCanvas } = require("chartjs-node-canvas");
const ChartJSLib = require("chart.js");
const ImageModule = require("docxtemplater-image-module-free");

const app = express();
const upload = multer({ dest: "uploads/" });
const PORT = process.env.PORT || 3001;

// === GRÁFICAS (único bloque, eje X = tiempo real) ===
const OUT_IMG_DIR = path.join(__dirname, "out", "imgs");
const IMG_TAGS = new Set([
  "img_ph",
  "img_oxigeno",
  "img_conductividad",
  "img_temperatura",
  "img_profundidad",
  "img_turbidez",
]);

function ensureDirSync(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

// Busca hoja y métrica según WATER_MAP
function findSheetAndMetric(data, sheetRegex, metricRegex) {
  let sheetName = null;
  for (const n of Object.keys(data)) {
    if (sheetRegex.test(normKey(n))) { sheetName = n; break; }
  }
  if (!sheetName) return null;

  const rowsArriba = data[sheetName]?.["Aguas arriba"] || [];
  const rowsAbajo  = data[sheetName]?.["Aguas abajo"]  || [];
  const sample = rowsArriba[0] || rowsAbajo[0] || null;
  if (!sample) return null;

  let metricKey = null;
  for (const k of Object.keys(sample)) {
    if (metricRegex.test(normKey(k))) { metricKey = k; break; }
  }
  if (!metricKey) {
    for (const k of Object.keys(sample)) { if (isValueKey(k)) { metricKey = k; break; } }
  }
  if (!metricKey) return null;

  return { sheetName, metricKey };
}

// Convierte filas -> puntos XY [{x:tsMs, y:num}] ordenados
function rowsToPoints(rows, metricKey) {
  const pts = [];
  for (const r of rows) {
    const v = toNumber(r[metricKey]);
    if (v == null) continue;
    const dtRaw = r["Created At -5 (copia)"] ?? r["Created At"];
    const dt = parseDateLoose(dtRaw);
    if (!dt) continue;
    pts.push({ x: dt.getTime(), y: Number(v) });
  }
  pts.sort((a, b) => a.x - b.x);
  return pts;
}

// Formato tick X (ms -> 'YYYY-MM-DD HH:mm')
function fmtTickTs(ms) {
  const d = new Date(ms);
  const pad = (n) => String(n).padStart(2, "0");
  const Y = d.getFullYear();
  const M = pad(d.getMonth() + 1);
  const D = pad(d.getDate());
  const h = pad(d.getHours());
  const m = pad(d.getMinutes());
  return `${Y}-${M}-${D} ${h}:${m}`;
}

// Render XY (sin labels categóricos)
async function renderChartPNG_XY({ title, pointsUp, pointsDown, filename }) {
  ensureDirSync(OUT_IMG_DIR);
  const width = 800, height = 450;

  const chartJSNodeCanvas = new ChartJSNodeCanvas({
    width, height,
    chartCallback: (Chart) => { Chart.register(...ChartJSLib.registerables); },
  });

  const minX = Math.min(
    pointsUp.length ? pointsUp[0].x : Infinity,
    pointsDown.length ? pointsDown[0].x : Infinity
  );
  const maxX = Math.max(
    pointsUp.length ? pointsUp[pointsUp.length - 1].x : -Infinity,
    pointsDown.length ? pointsDown[pointsDown.length - 1].x : -Infinity
  );

  const config = {
    type: "line",
    data: {
      datasets: [
        { label: "Aguas arriba", data: pointsUp,   parsing: false, spanGaps: true, borderWidth: 2, pointRadius: 0 },
        { label: "Aguas abajo",  data: pointsDown, parsing: false, spanGaps: true, borderWidth: 2, pointRadius: 0 },
      ],
    },
    options: {
      responsive: false,
      animation: false,
      plugins: { title: { display: true, text: title }, legend: { display: true } },
      scales: {
        x: {
          type: "linear",
          min: Number.isFinite(minX) ? minX : undefined,
          max: Number.isFinite(maxX) ? maxX : undefined,
          ticks: { callback: (value) => fmtTickTs(value), maxRotation: 0, autoSkip: true },
          grid: { display: true },
        },
        y: { beginAtZero: false, grid: { display: true } },
      },
    },
  };

  const buffer = await chartJSNodeCanvas.renderToBuffer(config);
  const outPath = path.join(OUT_IMG_DIR, filename);
  fs.writeFileSync(outPath, buffer);
  return outPath;
}

// Genera TODAS las gráficas y devuelve { img_ph, img_oxigeno, ... }
async function generateWaterCharts(data) {
  const TITLES = {
    conductividad: "Conductividad (µS/cm)",
    profundidad:   "Profundidad (m)",
    turbidez:      "Turbidez (NTU)",
    temperatura:   "Temperatura (°C)",
    ph:            "pH (Unidades de pH)",
    oxigenodisuelto: "Oxígeno Disuelto (mg/L)",
  };

  const images = {};
  for (const item of WATER_MAP) {
    const fm = findSheetAndMetric(data, item.sheet, item.metric);
    if (!fm) continue;
    const { sheetName, metricKey } = fm;

    const upRows   = data[sheetName]?.["Aguas arriba"] || [];
    const downRows = data[sheetName]?.["Aguas abajo"]  || [];

    const ptsUp   = rowsToPoints(upRows, metricKey);
    const ptsDown = rowsToPoints(downRows, metricKey);

    const outPath = await renderChartPNG_XY({
      title: TITLES[item.key] || item.key,
      pointsUp: ptsUp,
      pointsDown: ptsDown,
      filename: `chart_${item.key}.png`,
    });

    if (item.key === "ph")                   images["img_ph"] = outPath;
    else if (item.key === "oxigenodisuelto") images["img_oxigeno"] = outPath;
    else if (item.key === "conductividad")   images["img_conductividad"] = outPath;
    else if (item.key === "temperatura")     images["img_temperatura"]   = outPath;
    else if (item.key === "profundidad")     images["img_profundidad"]   = outPath;
    else if (item.key === "turbidez")        images["img_turbidez"]      = outPath;
  }
  return images;
}


// ==== Helpers específicos para la tabla de agua (Aguas arriba) ====
function removeDiacritics(s) {
  return String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
function normKey(s) {
  return removeDiacritics(s).toLowerCase().replace(/\s+/g, " ").trim();
}
function formatBogota(iso) {
  if (!iso) return "-";
  try {
    const fmt = new Intl.DateTimeFormat("es-CO", {
      dateStyle: "medium",
      timeStyle: "short",
      timeZone: "America/Bogota",
    });
    return fmt.format(new Date(iso));
  } catch {
    return String(iso).replace("T", " ").replace("Z", "");
  }
}
function fmtNum(n, digits = 2) {
  if (n == null || Number.isNaN(Number(n))) return "-";
  return Number(n).toFixed(digits);
}

// Calcula media/min/max/fechas SOLO para “Aguas arriba” dentro de una hoja
function statsAguasArriba(data, sheetRegex, metricRegex) {
  // hoja
  let sheetName = null;
  for (const n of Object.keys(data)) {
    if (sheetRegex.test(normKey(n))) { sheetName = n; break; }
  }
  if (!sheetName) return null;

  const rows = data[sheetName]?.["Aguas arriba"] || [];
  if (!rows.length) return null;

  // métrica (columna de valor)
  let metricKey = null;
  const keys = Object.keys(rows[0] || {});
  for (const k of keys) {
    if (metricRegex.test(normKey(k))) { metricKey = k; break; }
  }
  if (!metricKey) {
    for (const k of keys) { if (isValueKey(k)) { metricKey = k; break; } }
  }
  if (!metricKey) return null;

  // stats con media
  let min = Infinity, max = -Infinity, sum = 0, cnt = 0;
  let dtMin = null, dtMax = null;

  for (const r of rows) {
    const v = toNumber(r[metricKey]);
    if (v == null) continue;
    const dtRaw = r["Created At -5 (copia)"] ?? r["Created At"];
    const dt = parseDateLoose(dtRaw);
    if (!dt) continue;

    if (v < min) { min = v; dtMin = dt; }
    if (v > max) { max = v; dtMax = dt; }
    sum += v; cnt += 1;
  }
  if (!Number.isFinite(min) || !Number.isFinite(max) || cnt === 0) return null;

  const media = sum / cnt;
  return {
    media, max, min,
    fechaMax: dtMax ? dtMax.toISOString() : null,
    fechaMin: dtMin ? dtMin.toISOString() : null,
  };
}

// Mapeo hoja/métrica → placeholders del nuevo template
const WATER_MAP = [
  // key = prefijo usado en placeholders del DOCX
  { key: "conductividad", sheet: /conductividad|µs.?\/?cm|u?s.?\/?cm/, metric: /conduct/ },
  { key: "profundidad",   sheet: /profundidad|depth/,                metric: /(recdist|profund|depth)/ },
  { key: "turbidez",      sheet: /turbidez|ntu/,                      metric: /(turb|ntu)/ },
  { key: "temperatura",   sheet: /temperatura/,                       metric: /temper/ },
  { key: "ph",            sheet: /^ph$|^p\s*h$/,                      metric: /^ph$|^p\s*h$/ },
  { key: "oxigenodisuelto", sheet: /ox[ií]geno.*disuelto|dissolved.*oxygen/, metric: /(ox[ií]geno.*disuelt|dissolved.*oxygen)/, 
    // alias para las 2 variantes que trae tu DOCX
    aliases: ["oxigenodisuelt"] 
  },
];

// Construye el contexto con SOLO “Aguas arriba”
function buildWaterContextAguasArriba(data) {
  const ctx = {};
  for (const item of WATER_MAP) {
    const s = statsAguasArriba(data, item.sheet, item.metric);
    const put = (prefix, stat) => {
      ctx[`${prefix}media`]    = stat ? fmtNum(stat.media) : "-";
      ctx[`${prefix}max`]      = stat ? fmtNum(stat.max)   : "-";
      ctx[`${prefix}min`]      = stat ? fmtNum(stat.min)   : "-";
      ctx[`${prefix}fechamax`] = stat ? formatBogota(stat.fechaMax) : "-";
      ctx[`${prefix}fechamin`] = stat ? formatBogota(stat.fechaMin) : "-";
    };
    put(item.key, s);
    // coloca también alias (oxigenodisueltomedia/min/max + oxigenodisueltofechas…)
    if (item.aliases) {
      for (const a of item.aliases) put(a, s);
    }
  }

  // Si tu plantilla trae el bug en la celda de Temperatura (usa {turbidezfechamin}),
  // igual rellenamos la correcta:
  if (ctx["temperaturafechamin"]) {
    // nada: ya está
  } else {
    // por si algún día agregas {temperaturafechamin}, deja el valor preparado
    const sTemp = statsAguasArriba(data, /temperatura/, /temper/);
    if (sTemp) ctx["temperaturafechamin"] = formatBogota(sTemp.fechaMin);
  }
  return ctx;
}

/* -------------------- Helpers -------------------- */
const sanitizeSheet = (s) => String(s ?? "").normalize("NFKC").trim();

const normalizeName = (s) =>
  String(s ?? "undefined").normalize("NFKC").replace(/\s+/g, " ").trim();

function classifySite(nameRaw) {
  const n = normalizeName(nameRaw).toLowerCase();
  if (n.includes("arriba")) return "Aguas arriba";
  if (n.includes("abajo")) return "Aguas abajo";
  return "Otro";
}

function normalizeCell(v) {
  if (v instanceof Date) return v.toISOString();
  if (v && typeof v === "object") {
    if (v.text) return v.text;
    if (v.result) return v.result;
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (/^-?\d+,\d+$/.test(s)) {
      const n = Number(s.replace(",", "."));
      if (!Number.isNaN(n)) return n;
    }
    const maybe = Number(s);
    if (!Number.isNaN(maybe) && s !== "") return maybe;
    return s;
  }
  return v ?? "";
}

function readSheetAsRows(worksheet) {
  const headerRow = worksheet.getRow(1);
  const headers = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    headers[colNumber] = String(cell.value ?? "").trim();
  });

  const rows = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const obj = {};
    Object.keys(headers).forEach((k) => {
      const col = Number(k);
      if (!headers[col]) return;
      obj[headers[col]] = normalizeCell(row.getCell(col).value);
    });

    const hasAny = Object.values(obj).some(
      (v) => v !== "" && v !== null && v !== undefined
    );
    if (hasAny) rows.push(obj);
  });
  return rows;
}

function groupRowsBySite(rows) {
  const groups = { "Aguas arriba": [], "Aguas abajo": [], Otro: [] };
  for (const r of rows) {
    const site = classifySite(r.Name);
    groups[site].push({
      ...r,
      _NameClean: normalizeName(r.Name),
      _Site: site,
    });
  }
  return groups;
}

/* ---- Parse fecha robusto (ISO o dd/mm/yyyy [hh:mm[:ss]]) ---- */
function parseDateLoose(v) {
  if (v instanceof Date) return v;
  if (typeof v !== "string") {
    if (v && typeof v === "object" && v.text) return new Date(v.text);
    return null;
  }
  const s = v.trim().replace(/\u00A0/g, " ");

  // ISO con Z => UTC real
  if (/^\d{4}-\d{2}-\d{2}T.+Z$/.test(s)) {
    const t = Date.parse(s);
    return Number.isNaN(t) ? null : new Date(t);
  }

  // dd/mm/yyyy [hh:mm[:ss]] => **local**
  const m = s.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (m) {
    let d = parseInt(m[1], 10);
    let mo = parseInt(m[2], 10);
    const y = parseInt(m[3], 10);
    const hh = parseInt(m[4] || "0", 10);
    const mm = parseInt(m[5] || "0", 10);
    const ss = parseInt(m[6] || "0", 10);
    // si parece mm/dd, corrige
    if (mo > 12 && d <= 12) [d, mo] = [mo, d];
    // **LOCAL**
    return new Date(y, mo - 1, d, hh, mm, ss);
  }

  // ISO sin Z (o con espacio) => **local**
  const isoLocal = s.match(
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/
  );
  if (isoLocal) {
    const [, Y, M, D, H, Mi, S = "0"] = isoLocal;
    return new Date(+Y, +M - 1, +D, +H, +Mi, +S);
  }

  // último intento
  const t = Date.parse(s);
  return Number.isNaN(t) ? null : new Date(t);
}


/* ---- detectar columnas de valor (métricas) ---- */
const VALUE_COL_PATTERNS = [
  /^SUMA?\s*\(.+\)$/i,
  /^RECDIST\(.+\)$/i,
  /^AVG\(.+\)$/i,
  /^AVERAGE\(.+\)$/i,
];
function isValueKey(key) {
  return VALUE_COL_PATTERNS.some((rx) => rx.test(key));
}
function toNumber(val) {
  if (typeof val === "number") return val;
  if (typeof val === "string") {
    const s = val.trim();
    if (/^-?\d+,\d+$/.test(s)) return Number(s.replace(",", "."));
    const n = Number(s);
    return Number.isNaN(n) ? null : n;
  }
  return null;
}

/* ---- stats: min, max, fechaMin, fechaMax por métrica ---- */
function computeStatsForRows(rows, metricKey) {
  let min = Infinity;
  let max = -Infinity;
  let dateMin = null;
  let dateMax = null;

  for (const r of rows) {
    const v = toNumber(r[metricKey]);
    if (v == null) continue;

    const dtRaw = r["Created At -5 (copia)"] ?? r["Created At"];
    const dt = parseDateLoose(dtRaw);
    if (!dt) continue;

    if (v < min) { min = v; dateMin = dt; }
    if (v > max) { max = v; dateMax = dt; }
  }

  if (!Number.isFinite(min) || !Number.isFinite(max)) return null;
  return {
    min,
    max,
    fechaMin: dateMin ? dateMin.toISOString() : null,
    fechaMax: dateMax ? dateMax.toISOString() : null,
  };
}

/* ---- periodRange (min/max global) ---- */
function computePeriodRange(data) {
  let minTs = Infinity;
  let maxTs = -Infinity;

  const pick = (val) => {
    const dt = parseDateLoose(val);
    if (!dt) return;
    const ts = dt.getTime();
    if (!Number.isFinite(ts)) return;
    if (ts < minTs) minTs = ts;
    if (ts > maxTs) maxTs = ts;
  };

  for (const sheet of Object.values(data)) {
    for (const siteArr of Object.values(sheet)) {
      for (const row of siteArr) {
        if ("Created At -5 (copia)" in row) pick(row["Created At -5 (copia)"]);
        else if ("Created At" in row) pick(row["Created At"]);
      }
    }
  }

  const start = Number.isFinite(minTs) ? new Date(minTs) : null;
  const end = Number.isFinite(maxTs) ? new Date(maxTs) : null;

  return {
    isoStart: start ? start.toISOString() : null,
    isoEnd: end ? end.toISOString() : null,
  };
}

// formateo humano para el DOCX (zona Bogotá)
function formatBogota(iso) {
  if (!iso) return "";
  try {
    const fmt = new Intl.DateTimeFormat("es-CO", {
      dateStyle: "medium",
      timeStyle: "short",
      timeZone: "America/Bogota",
    });
    return fmt.format(new Date(iso));
  } catch {
    return String(iso);
  }
}

/* -------------------- DOCX: solo {periodRange} -------------------- */
app.post("/api/docx/period", upload.single("file"), async (req, res) => {
  let tmp = null;
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo 'file'" });
    tmp = req.file.path;

    // 1) Leer Excel y construir "data"
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp);
    const data = {};
    wb.worksheets.forEach((ws) => {
      const sheetName = sanitizeSheet(ws.name);
      const rows = readSheetAsRows(ws);
      const grouped = groupRowsBySite(rows);
      data[sheetName] = grouped;
    });

    // 2) Calcular periodRange
    const periodRange = computePeriodRange(data);
    const startLabel = formatBogota(periodRange.isoStart);
    const endLabel   = formatBogota(periodRange.isoEnd);
    const periodRangeLabel = `${startLabel} — ${endLabel} (GMT-5)`;
    console.log("periodRange:", { startLabel, endLabel, periodRangeLabel });

    // 3) Cargar y validar template
    const templatePath = path.join(__dirname, "template.docx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "No se encontró template.docx" });
    }
    const content = fs.readFileSync(templatePath); // Buffer
    let zip;
    try {
      zip = new PizZip(content);
      if (!zip.files || !zip.files["[Content_Types].xml"]) {
        throw new Error("El archivo no parece un DOCX válido");
      }
    } catch (e) {
      console.error("PizZip error:", e);
      return res.status(500).json({ error: "template.docx inválido o corrupto" });
    }

    // 4) Render con Docxtemplater
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    try {
      doc.render({
        periodRange: periodRangeLabel,     // por compatibilidad
        rangoFechaInicio: startLabel,      // NUEVO
        rangoFechaFin: endLabel            // NUEVO
      });
    } catch (e) {
      const info = {
        message: e.message,
        explanation: e.properties && e.properties.explanation,
        id: e.properties && e.properties.id,
        errors: e.properties && e.properties.errors,
      };
      console.error("Docxtemplater render error:", info);
      return res.status(400).json({ error: "Error al renderizar DOCX", detail: info });
    }

    // 5) Generar, guardar y descargar
    const buf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

    const outDir = path.join(__dirname, "out");
    await fsp.mkdir(outDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const outFile = path.join(outDir, `informe_periodo_${stamp}.docx`);
    fs.writeFileSync(outFile, buf);

    return res.download(outFile, "informe_periodo.docx");
  } catch (err) {
    console.error("docx/period error:", err);
    res.status(500).json({ error: err.message });
  } finally {
    if (tmp) { try { await fsp.unlink(tmp); } catch {} }
  }
});


/* -------------------- Endpoint de datos + stats -------------------- */
app.post("/api/xlsx-by-site", upload.single("file"), async (req, res) => {
  let tmp = null;
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo 'file'" });
    tmp = req.file.path;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp);

    const data = {};
    const stats = {};

    wb.worksheets.forEach((ws) => {
      const sheetName = sanitizeSheet(ws.name);
      const rows = readSheetAsRows(ws);
      const grouped = groupRowsBySite(rows);

      const allKeys = new Set();
      rows.forEach(r => Object.keys(r).forEach(k => allKeys.add(k)));
      const metricKeys = [...allKeys].filter(isValueKey);

      const sheetStats = {};
      Object.entries(grouped).forEach(([site, arr]) => {
        const siteStats = {};
        metricKeys.forEach((mk) => {
          const s = computeStatsForRows(arr, mk);
          if (s) siteStats[mk] = s;
        });
        sheetStats[site] = siteStats;
      });

      data[sheetName] = grouped;
      stats[sheetName] = sheetStats;

      const uniqNames = Array.from(
        new Set(rows.filter(r => r.Name != null).map(r => normalizeName(r.Name)))
      );
      console.log(`Hoja "${sheetName}" → Names únicos:`, uniqNames);
      console.log(
        `Hoja "${sheetName}" → Conteo por sitio:`,
        Object.fromEntries(Object.entries(grouped).map(([k,v]) => [k, v.length]))
      );
    });

    const periodRange = computePeriodRange(data);
    console.log("periodRange:", periodRange);

    const payload = { ok: true, data, periodRange, stats };
    const logsDir = path.join(__dirname, "logs");
    await fsp.mkdir(logsDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const outFile = path.join(logsDir, `by-site_${stamp}.json`);
    await fsp.writeFile(outFile, JSON.stringify(payload, null, 2), "utf8");
    console.log(`Respuesta guardada en: ${outFile}`);

    res.json({ ...payload, savedTo: outFile });
  } catch (err) {
    console.error("xlsx-by-site error:", err);
    res.status(500).json({ ok: false, error: err.message });
  } finally {
    if (tmp) { try { await fsp.unlink(tmp); } catch {} }
  }
});

/* ========= ENDPOINT: DOCX con tabla (solo Aguas arriba) ========= */
app.post("/api/docx/tabla-aguas-arriba", upload.single("file"), async (req, res) => {
  let tmp = null;
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo 'file'" });
    tmp = req.file.path;

    // 1) Excel → data
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp);
    const data = {};
    wb.worksheets.forEach((ws) => {
      const sheetName = sanitizeSheet(ws.name);
      const rows = readSheetAsRows(ws);
      const grouped = groupRowsBySite(rows);
      data[sheetName] = grouped;
    });

    // 2) periodRange + contexto tabla (Aguas arriba)
    const periodRange = computePeriodRange(data);
    const ctxTabla = buildWaterContextAguasArriba(data);
    const context = {
      periodRange: `${formatBogota(periodRange.isoStart)} — ${formatBogota(periodRange.isoEnd)}`,
      rangoFechaInicio: formatBogota(periodRange.isoStart),
      rangoFechaFin:    formatBogota(periodRange.isoEnd),
      ...ctxTabla, // <<< mantiene TUS placeholders de estadísticas (Aguas arriba)
    };

    // === NUEVO: generar imágenes de las 6 variables (Arriba vs Abajo) ===
    const imgs = await generateWaterCharts(data);

    // === Cargar y validar template ===
    const templatePath = path.join(__dirname, "template.docx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "No se encontró template.docx" });
    }
    const content = fs.readFileSync(templatePath);
    let zip;
    try {
      zip = new PizZip(content);
      if (!zip.files || !zip.files["[Content_Types].xml"]) {
        throw new Error("El archivo no parece un DOCX válido");
      }
    } catch (e) {
      console.error("PizZip error:", e);
      return res.status(500).json({ error: "template.docx inválido o corrupto" });
    }

    // === ImageModule: marcar explícitamente qué tags son imagen ===
    const imModule = new ImageModule({
      fileType: "docx",
      centered: true,
      getTagType: (tag) => (IMG_TAGS.has(tag) ? "image" : "text"),
      getImage: (tagValue/*, tagName*/) => {
        try {
          if (!tagValue) return Buffer.alloc(0);
          if (fs.existsSync(tagValue)) return fs.readFileSync(tagValue);
          return Buffer.alloc(0);
        } catch { return Buffer.alloc(0); }
      },
      getSize: () => [600, 338], // ~15 cm ancho en Word
    });

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      modules: [imModule],
    });

    // === Render final: FECHAS + STATS + IMÁGENES ===
    try {
      doc.render({ ...context, ...imgs });
    } catch (e) {
      const info = {
        message: e.message,
        explanation: e.properties && e.properties.explanation,
        id: e.properties && e.properties.id,
        errors: e.properties && e.properties.errors,
      };
      console.error("Docxtemplater render error:", info);
      return res.status(400).json({ error: "Error al renderizar DOCX", detail: info });
    }

    // === Guardar y devolver ===
    const buf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
    const outDir = path.join(__dirname, "out");
    await fsp.mkdir(outDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const outFile = path.join(outDir, `informe_aguas_arriba_${stamp}.docx`);
    fs.writeFileSync(outFile, buf);
    return res.download(outFile, "informe_aguas_arriba.docx");
  } catch (err) {
    console.error("docx/tabla-aguas-arriba error:", err);
    res.status(500).json({ error: err.message });
  } finally {
    if (tmp) { try { await fsp.unlink(tmp); } catch {} }
  }
});

/* ========= ENDPOINT: DOCX con GRÁFICAS (Arriba vs Abajo) ========= */
app.post("/api/docx/graficas", upload.single("file"), async (req, res) => {
  let tmp = null;
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo 'file' (.xlsx)" });
    tmp = req.file.path;

    // 1) Excel → data (reutiliza tus helpers)
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp);
    const data = {};
    wb.worksheets.forEach((ws) => {
      const sheetName = sanitizeSheet(ws.name);
      const rows = readSheetAsRows(ws);
      const grouped = groupRowsBySite(rows);
      data[sheetName] = grouped;
    });

    // 2) Calcular rango de fechas para el contexto
    const periodRange = computePeriodRange(data);
    const ctxBase = {
      periodRange: `${formatBogota(periodRange.isoStart)} — ${formatBogota(periodRange.isoEnd)}`,
      rangoFechaInicio: formatBogota(periodRange.isoStart),
      rangoFechaFin:    formatBogota(periodRange.isoEnd),
    };

    // 3) Generar PNGs de todas las variables
    const imgs = await generateWaterCharts(data);

    // 4) Cargar template.docx y configurar módulo de imágenes
    const templatePath = path.join(__dirname, "template.docx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "No se encontró template.docx" });
    }
    const content = fs.readFileSync(templatePath);
    let zip;
    try {
      zip = new PizZip(content);
      if (!zip.files || !zip.files["[Content_Types].xml"]) {
        throw new Error("El archivo no parece un DOCX válido");
      }
    } catch (e) {
      console.error("PizZip error:", e);
      return res.status(500).json({ error: "template.docx inválido o corrupto" });
    }

    const imModule = new ImageModule({
      centered: true,
      getImage: (tagValue) => {
        if (!tagValue || !fs.existsSync(tagValue)) return Buffer.alloc(0);
        return fs.readFileSync(tagValue);
      },
      getSize: () => {
        // ~15 cm de ancho en Word: 15cm ≈ 567 pt ≈ 1134 px a 96dpi
        // Como generamos 1600x900, lo acercamos a 1200x675
        return [1200, 675];
      },
    });

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      modules: [imModule],
    });

    // 5) Renderizar con contexto (fechas + rutas de imágenes)
    const context = { ...ctxBase, ...imgs };
    try {
      doc.render(context);
    } catch (e) {
      const info = {
        message: e.message,
        explanation: e.properties && e.properties.explanation,
        id: e.properties && e.properties.id,
        errors: e.properties && e.properties.errors,
      };
      console.error("Docxtemplater render error:", info);
      return res.status(400).json({ error: "Error al renderizar DOCX", detail: info });
    }

    // 6) Guardar y descargar
    const buf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
    const outDir = path.join(__dirname, "out");
    await fsp.mkdir(outDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const outFile = path.join(outDir, `informe_graficas_${stamp}.docx`);
    fs.writeFileSync(outFile, buf);

    return res.download(outFile, "informe_graficas.docx");
  } catch (err) {
    console.error("docx/graficas error:", err);
    res.status(500).json({ error: err.message });
  } finally {
    if (tmp) { try { await fsp.unlink(tmp); } catch {} }
  }
});


app.get("/", (_req, res) => res.send("API XLSX agrupado por sitio activa"));
app.listen(PORT, () => console.log(`Server on :${PORT}`));

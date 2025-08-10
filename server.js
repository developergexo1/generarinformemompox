// server.js
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const fsp = fs.promises;
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });
const PORT = process.env.PORT || 3001;

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
    // "5,33" -> 5.33 (coma decimal sin miles)
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
  // lee encabezados (fila 1)
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

/* ---- Parse de fecha robusto (ISO ó dd/mm/yyyy [hh:mm[:ss]]) ---- */
function parseDateLoose(v) {
  if (v instanceof Date) return v;
  if (typeof v !== "string") {
    if (v && typeof v === "object" && v.text) return new Date(v.text);
    return null;
  }
  const s = v.trim().replace(/\u00A0/g, " "); // NBSP -> espacio
  if (/\d{4}-\d{2}-\d{2}T/.test(s)) {
    const t = Date.parse(s);
    return Number.isNaN(t) ? null : new Date(t);
  }
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
    if (mo > 12 && d <= 12) [d, mo] = [mo, d]; // por si viene mm/dd
    return new Date(Date.UTC(y, mo - 1, d, hh, mm, ss));
  }
  const t = Date.parse(s);
  return Number.isNaN(t) ? null : new Date(t);
}

/* ---- util: detectar columnas de valor (métricas) ---- */
const VALUE_COL_PATTERNS = [
  /^SUMA?\s*\(.+\)$/i,     // SUM(...), SUMA(...)
  /^RECDIST\(.+\)$/i,      // RECDIST(Depth)
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

/* -------------------- Endpoint -------------------- */
app.post("/api/xlsx-by-site", upload.single("file"), async (req, res) => {
  let tmp = null;
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo 'file'" });
    tmp = req.file.path;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp); // ExcelJS lee el XLSX (workbook/worksheets/cells) :contentReference[oaicite:0]{index=0}

    // data: { "Temperatura": { "Aguas arriba":[...], "Aguas abajo":[...], "Otro":[...] }, ... }
    const data = {};
    const stats = {}; // NUEVO: stats[sheet][site][metric] = { min, max, fechaMin, fechaMax }

    wb.worksheets.forEach((ws) => {
      const sheetName = sanitizeSheet(ws.name);
      const rows = readSheetAsRows(ws);
      const grouped = groupRowsBySite(rows);

      // Detectar métricas (keys tipo SUM(...)/SUMA(...)/RECDIST(...))
      const allKeys = new Set();
      rows.forEach(r => Object.keys(r).forEach(k => allKeys.add(k)));
      const metricKeys = [...allKeys].filter(isValueKey);

      // Calcular stats por sitio y por métrica
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

      // Debug
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

    // Guardar archivo con la respuesta completa
    const payload = { ok: true, data, periodRange, stats };
    const logsDir = path.join(__dirname, "logs");
    await fsp.mkdir(logsDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const outFile = path.join(logsDir, `by-site_${stamp}.json`);
    await fsp.writeFile(outFile, JSON.stringify(payload, null, 2), "utf8");
    console.log(`Respuesta guardada en: ${outFile}`);

    // Responder al cliente
    res.json({ ...payload, savedTo: outFile });
  } catch (err) {
    console.error("xlsx-by-site error:", err);
    res.status(500).json({ ok: false, error: err.message });
  } finally {
    if (tmp) { try { await fsp.unlink(tmp); } catch {} }
  }
});

app.get("/", (_req, res) => res.send("API XLSX agrupado por sitio activa"));
app.listen(PORT, () => console.log(`Server on :${PORT}`));

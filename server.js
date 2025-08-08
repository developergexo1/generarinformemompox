const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });
const PORT = process.env.PORT || 3001;

// Mapea nombres de hoja si quieres "amigables"
const SHEET_LABEL_MAP = {
  "Oxígeno Disuelto": "Oxigeno Disuelto",
  Temperatura: "Temperatura",
  "Turbidez (NTU)": "Turbidez (NTU)",
  "Conductividad (µScm) ": "Conductividad (µScm)",
  "Profundidad (m)": "Profundidad (m)",
  pH: "pH",
};

function formatCell(cell) {
  if (cell instanceof Date) return cell.toISOString();
  if (typeof cell === "object" && cell !== null) {
    if (cell.text) return cell.text;
    if (cell.richText) return cell.richText.map((t) => t.text).join("");
    if (cell.result) return cell.result;
  }
  return String(cell);
}

// calcula stats de una serie de tiempos y valores
function computeStats(times, values) {
  const paired = times
    .map((t, i) => ({ time: t, value: values[i] }))
    .filter((p) => typeof p.value === "number" && !isNaN(p.value));
  if (!paired.length) return null;

  const vals = paired.map((p) => p.value);
  const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
  const max = Math.max(...vals);
  const min = Math.min(...vals);
  const dateMax = paired.find((p) => p.value === max)?.time || null;
  const dateMin = paired.find((p) => p.value === min)?.time || null;

  return {
    mean,
    max,
    min,
    date_max: dateMax,
    date_min: dateMin,
  };
}

async function extractWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const result = [];

  for (const worksheet of workbook.worksheets) {
    const originalName = worksheet.name;
    const friendlyName = SHEET_LABEL_MAP[originalName] || originalName;

    const headerRow = worksheet.getRow(1);
    const headers = {};
    headerRow.eachCell((cell, colNumber) => {
      const txt = String(cell.value || "").trim();
      headers[colNumber] = txt;
    });

    // detectar columnas clave
    const createdAtCols = Object.entries(headers)
      .filter(([, name]) => /Created At -5/i.test(name))
      .map(([col]) => parseInt(col));
    const nameCols = Object.entries(headers)
      .filter(([, name]) => /^Name$/i.test(name))
      .map(([col]) => parseInt(col));
    const sumaCols = Object.entries(headers)
      .filter(([, name]) => /^SUMA\(.+\)/i.test(name))
      .map(([col]) => parseInt(col));

    if (createdAtCols.length === 0 || sumaCols.length === 0) {
      // no tiene la forma esperada; igual puede saltarse o incluir vacío
      continue;
    }

    const group = {};

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // header
      const createdAtRaw = row.getCell(createdAtCols[0]).value;
      const createdAt = formatCell(createdAtRaw);
      const name = nameCols.length
        ? formatCell(row.getCell(nameCols[0]).value)
        : "undefined";

      const entry = {
        "Created At -5 (copia)": createdAt,
      };

      sumaCols.forEach((col) => {
        const h = headers[col];
        entry[h] = formatCell(row.getCell(col).value);
      });

      if (!group[name]) group[name] = [];
      group[name].push(entry);
    });

    result.push({
      [friendlyName]: group,
    });
  }

  return result;
}

async function extractWithStats(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const structured = [];

  for (const worksheet of workbook.worksheets) {
    const originalName = worksheet.name;
    const friendlyName = SHEET_LABEL_MAP[originalName] || originalName;

    const headerRow = worksheet.getRow(1);
    const headers = {};
    headerRow.eachCell((cell, colNumber) => {
      headers[colNumber] = String(cell.value || "").trim();
    });

    const createdAtCols = Object.entries(headers)
      .filter(([, name]) => /Created At -5/i.test(name))
      .map(([col]) => parseInt(col));
    const nameCols = Object.entries(headers)
      .filter(([, name]) => /^Name$/i.test(name))
      .map(([col]) => parseInt(col));
    const sumaCols = Object.entries(headers)
      .filter(([, name]) => /^SUMA\(.+\)/i.test(name))
      .map(([col]) => parseInt(col));

    if (createdAtCols.length === 0 || sumaCols.length === 0) continue;

    const group = {};

    // primero agrupar por Name y dentro guardar series para stats
    const temp = {}; // temp[name][sumaHeader] = { times:[], values:[] }
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const createdAtRaw = row.getCell(createdAtCols[0]).value;
      const createdAt = createdAtRaw instanceof Date ? createdAtRaw : new Date(createdAtRaw);
      const name = nameCols.length
        ? formatCell(row.getCell(nameCols[0]).value)
        : "undefined";

      if (!temp[name]) temp[name] = {};
      sumaCols.forEach((col) => {
        const headerName = headers[col];
        const valRaw = row.getCell(col).value;
        const valNum = Number(String(valRaw).replace(",", ".")); // comas decimales
        if (isNaN(valNum)) return;
        if (!temp[name][headerName]) temp[name][headerName] = { times: [], values: [] };
        temp[name][headerName].times.push(createdAt);
        temp[name][headerName].values.push(valNum);
      });
    });

    // construir salida con stats por cada suma header
    for (const [name, sums] of Object.entries(temp)) {
      group[name] = {};
      for (const [sumaHeader, series] of Object.entries(sums)) {
        const stats = computeStats(series.times, series.values);
        group[name][sumaHeader] = {
          raw: series.times.map((t, i) => ({
            "Created At -5 (copia)": series.times[i].toISOString(),
            [sumaHeader]: series.values[i],
          })),
          stats,
        };
      }
    }

    structured.push({ [friendlyName]: group });
  }

  return structured;
}

// endpoint básico: devuelve estructura sin stats
app.post("/api/extract", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo" });
    const data = await extractWorkbook(req.file.path);
    console.log("Extract result:", JSON.stringify(data, null, 2));
    res.json({ extracted: data });
  } catch (err) {
    console.error("Error extract:", err);
    res.status(500).json({ error: err.message });
  }
});

// endpoint con stats
app.post("/api/extract-with-stats", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo" });
    const data = await extractWithStats(req.file.path);
    console.log("Extract with stats:", JSON.stringify(data, null, 2));
    res.json({ extracted: data });
  } catch (err) {
    console.error("Error extract-with-stats:", err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/", (_req, res) => {
  res.send("API de extracción activa");
});

app.listen(PORT, () => {
  console.log(`Servidor corriendo en puerto ${PORT}`);
});

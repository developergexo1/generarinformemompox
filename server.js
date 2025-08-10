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
    // agrega metadatos útiles normalizados (no molestan en tu consumo)
    const withMeta = {
      ...r,
      _NameClean: normalizeName(r.Name),
      _Site: site,
    };
    groups[site].push(withMeta);
  }
  return groups;
}

/* -------------------- Endpoint -------------------- */
app.post("/api/xlsx-by-site", upload.single("file"), async (req, res) => {
  let tmp = null;
  try {
    if (!req.file) return res.status(400).json({ error: "Falta archivo 'file'" });
    tmp = req.file.path;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tmp); // ExcelJS lee el .xlsx :contentReference[oaicite:0]{index=0}

    // Estructura final: { "Temperatura": { "Aguas arriba":[...], "Aguas abajo":[...] }, ... }
    const data = {};
    wb.worksheets.forEach((ws) => {
      const sheetName = sanitizeSheet(ws.name);
      const rows = readSheetAsRows(ws);
      const grouped = groupRowsBySite(rows);

      // Debug útil
      const uniqNames = Array.from(
        new Set(rows.filter(r => r.Name != null).map(r => normalizeName(r.Name)))
      );
      console.log(`Hoja "${sheetName}" → Names únicos:`, uniqNames);
      console.log(
        `Hoja "${sheetName}" → Conteo por sitio:`,
        Object.fromEntries(Object.entries(grouped).map(([k,v]) => [k, v.length]))
      );

      data[sheetName] = grouped;
    });

    // Guarda archivo con la respuesta completa
    const logsDir = path.join(__dirname, "logs");
    await fsp.mkdir(logsDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const outFile = path.join(logsDir, `by-site_${stamp}.json`);
    await fsp.writeFile(outFile, JSON.stringify({ ok: true, data }, null, 2), "utf8");
    console.log(`Respuesta guardada en: ${outFile}`);

    // Responde al cliente
    res.json({ ok: true, data, savedTo: outFile });
  } catch (err) {
    console.error("xlsx-by-site error:", err);
    res.status(500).json({ ok: false, error: err.message });
  } finally {
    if (tmp) { try { await fsp.unlink(tmp); } catch {} }
  }
});

app.get("/", (_req, res) => res.send("API XLSX agrupado por sitio activa"));

app.listen(PORT, () => console.log(`Server on :${PORT}`));

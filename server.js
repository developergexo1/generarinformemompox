const ExcelJS = require("exceljs");
const path = require("path");

if (process.argv.length < 3) {
  console.error("Uso: node extract.js <ruta-al-excel>");
  process.exit(1);
}

const filePath = process.argv[2];

const SHEET_LABEL_MAP = {
  "Oxígeno Disuelto": "Oxigeno Disuelto",
  Temperatura: "Temperatura",
  "Turbidez (NTU)": "Turbidez (NTU)",
  "Conductividad (µScm) ": "Conductividad (µScm)",
  "Profundidad (m)": "Profundidad (m)",
  pH: "pH",
  // agrega más si hay otras hojas
};

async function extract(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  const result = [];

  for (const worksheet of workbook.worksheets) {
    const originalName = worksheet.name;
    const friendlyName = SHEET_LABEL_MAP[originalName] || originalName;

    // leer encabezados para saber qué columnas existen
    const headerRow = worksheet.getRow(1);
    const headers = {};
    headerRow.eachCell((cell, colNumber) => {
      const txt = String(cell.value || "").trim();
      headers[colNumber] = txt;
    });

    // buscamos columnas clave: "Created At -5 (copia)" y la que empieza con "SUMA("
    // también tomamos "Name" si existe
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
      // no tiene la estructura esperada: saltar
      continue;
    }

    const group = {};
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // header

      const createdAt = row.getCell(createdAtCols[0]).value;
      const name = nameCols.length ? row.getCell(nameCols[0]).value : "undefined";
      // tomar todas las columnas SUMA(...) y ponerlas como campos separados
      const sumaValues = {};
      sumaCols.forEach((col) => {
        const headerName = headers[col];
        sumaValues[headerName] = row.getCell(col).value;
      });

      const entry = {
        "Created At -5 (copia)": formatCell(createdAt),
        ...sumaValues,
      };

      const groupKey = String(name);
      if (!group[groupKey]) group[groupKey] = [];
      group[groupKey].push(entry);
    });

    const obj = {
      [friendlyName]: group,
    };
    result.push(obj);
  }

  console.log(JSON.stringify(result, null, 2));
}

function formatCell(cell) {
  if (cell instanceof Date) return cell.toISOString();
  if (typeof cell === "object" && cell !== null) {
    if (cell.text) return cell.text;
    if (cell.richText) return cell.richText.map((t) => t.text).join("");
  }
  return String(cell);
}

extract(path.resolve(filePath)).catch((e) => {
  console.error("Error al procesar:", e);
});

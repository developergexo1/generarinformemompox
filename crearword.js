// proyecto.js
const officegen = require('officegen');
const fs = require('fs');

// 1) Crear documento Word
const docx = officegen('docx');

// 2) Párrafo justificado con interlineado 1.0 y estilo definido
const p1 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p1.addText('Proyecto: ', {
  bold: true,
  font_face: 'Microsoft YaHei UI',
  font_size: 14,
  color: '002060'
});
p1.addText(
  'Fortalecimiento de capacidades para la apropiación tecnológica, conectividad digital y construcción de un territorio digitalmente transformado en el municipio de Santa Cruz de Mompox del departamento de Bolívar',
  {
    font_face: 'Microsoft YaHei UI',
    font_size: 14,
    color: '002060'
  }
);
const prueba = "PERIODO DE FECHAS SELECCIONADO";
// 3) Párrafo centrado en negrita “PERIODO DE FECHAS SELECCIONADO”
const p2 = docx.createP({
  align: 'center',
  spacing: { before: 200, after: 200 }
});
p2.addText(`${prueba}`, {
  bold: true,
  font_face: 'Microsoft YaHei UI',
  font_size: 14,
  color: '002060'
});

// 4) Salto de página (como Ctrl+Enter)
docx.putPageBreak();

// 5) Encabezado "RESUMEN" centrado y en negrita
const p3 = docx.createP({
  align: 'center',
  spacing: { before: 200, after: 100 }
});
p3.addText('RESUMEN', {
  bold: true,
  font_face: 'Microsoft YaHei UI',
  font_size: 14,
  color: '002060'
});

// 6) Texto de descripción con tamaño 12, justificado
const p4 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p4.addText(
  'Los datos que se muestran en este informe corresponden a mediciones realizadas entre las fechas 2025-04-01 00:00:00 y 2025-04-30 23:59:59. Todos los valores se reportan cada 10 minutos en ese período de tiempo.',
  {
    font_face: 'Microsoft YaHei UI',
    font_size: 12,
    color: '002060'
  }
);

// 7) Sección “Dispositivo” en un solo párrafo con saltos suaves (Shift+Enter)
const p5 = docx.createP({
  align: 'left',
  spacing: { line: 240, before: 200 } // interlineado 1.0
});

p5.addText('Dispositivo', {
  bold: true,
  font_face: 'Microsoft YaHei UI',
  font_size: 13,
  color: '002060'
});

p5.addLineBreak(); // Shift+Enter
p5.addText('Nombre: Kunak', {
  font_face: 'Microsoft YaHei UI',
  font_size: 12,
  color: '002060'
});

p5.addLineBreak(); // Shift+Enter
p5.addText('Referencia: Kunak AIR Pro', {
  font_face: 'Microsoft YaHei UI',
  font_size: 12,
  color: '002060'
});

// 8) Sección “Sensores” en un solo párrafo con saltos suaves (Shift+Enter)
const p6 = docx.createP({
  align: 'left',
  spacing: { line: 240, before: 200 } // interlineado 1.0
});
p6.addText('Sensores', {
  bold: true,
  font_face: 'Microsoft YaHei UI',
  font_size: 13,
  color: '002060'
});

p6.addLineBreak();
p6.addText('PM₁₀ (µg/m³)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('PM₂.₅ (µg/m³)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('SO₂ (PPB)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('NO₂ (PPB)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('O₃ (PPB)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('CO (PPB)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('Velocidad del viento (m/s)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('Dirección del viento (º)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('Precipitación (mm)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('Presión atmosférica (mmHg)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('Temperatura (°C)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

p6.addLineBreak();
p6.addText('Humedad (%)', { font_face: 'Microsoft YaHei UI', font_size: 12, color: '002060' });

// 9) Sección “Estadísticas” con saltos suaves (Shift+Enter)
const p7 = docx.createP({
  align: 'left',
  spacing: { line: 240, before: 200 }
});

p7.addText('Estadísticas', {
  bold: true,
  font_face: 'Microsoft YaHei UI',
  font_size: 13,
  color: '002060'
});

p7.addLineBreak();
p7.addText('Media: Media de los valores en el análisis', {
  font_face: 'Microsoft YaHei UI',
  font_size: 12,
  color: '002060'
});

p7.addLineBreak();
p7.addText('Max: Valor máximo horario', {
  font_face: 'Microsoft YaHei UI',
  font_size: 12,
  color: '002060'
});

p7.addLineBreak();
p7.addText('Min: Valor mínimo horario', {
  font_face: 'Microsoft YaHei UI',
  font_size: 12,
  color: '002060'
});

p7.addLineBreak();
p7.addText('Fecha Max: Fecha cuando se reportó el valor máximo horario', {
  font_face: 'Microsoft YaHei UI',
  font_size: 12,
  color: '002060'
});

p7.addLineBreak();
p7.addText('Fecha Min: Fecha cuando se reportó el valor mínimo horario', {
  font_face: 'Microsoft YaHei UI',
  font_size: 12,
  color: '002060'
});

// 10) Página nueva y tabla "Estadísticas por sensor"
docx.putPageBreak();

// Cabeceras y estilos
const headerFill = 'FFFFFF';   // azul
const headerText = '002060';   // blanco
const col1 = 7000;             // ancho 1ra col (Sensor)
const colN = 1800;             // ancho columnas numéricas

const headerRow = [
  { val: 'Sensor',    opts: { b:true, color: headerText, shd:{ fill: headerFill }, align:'center', vAlign:'center', fontFamily:'Microsoft YaHei UI', sz:'18', cellColWidth: col1 } },
  { val: 'Media',     opts: { b:true, color: headerText, shd:{ fill: headerFill }, align:'center', vAlign:'center', fontFamily:'Microsoft YaHei UI', sz:'18', cellColWidth: colN } },
  { val: 'Max',       opts: { b:true, color: headerText, shd:{ fill: headerFill }, align:'center', vAlign:'center', fontFamily:'Microsoft YaHei UI', sz:'18', cellColWidth: colN } },
  { val: 'Min',       opts: { b:true, color: headerText, shd:{ fill: headerFill }, align:'center', vAlign:'center', fontFamily:'Microsoft YaHei UI', sz:'18', cellColWidth: colN } },
  { val: 'Fecha Max', opts: { b:true, color: headerText, shd:{ fill: headerFill }, align:'center', vAlign:'center', fontFamily:'Microsoft YaHei UI', sz:'18', cellColWidth: colN } },
  { val: 'Fecha Min', opts: { b:true, color: headerText, shd:{ fill: headerFill }, align:'center', vAlign:'center', fontFamily:'Microsoft YaHei UI', sz:'18', cellColWidth: colN } },
];

const sensores = [
  'PM₁₀ (µg/m³)',
  'PM₂.₅ (µg/m³)',
  'SO₂ (PPB)',
  'NO₂ (PPB)',
  'O₃ (PPB)',
  'CO (PPB)',
  'Velocidad del viento (m/s)',
  'Dirección del viento (º)',
  'Precipitación (mm)',
  'Presión atmosférica (mmHg)',
  'Temperatura (°C)',
  'Humedad (%)'
];

// Construir filas (valores vacíos por ahora)
const rows = sensores.map(nombre => ([
  { val: nombre, opts: { b:true, fontFamily:'Microsoft YaHei UI', sz:'18', color:'002060', cellColWidth: col1 } },
  { val: '-',     opts: { b:true, align:'center', fontFamily:'Microsoft YaHei UI', sz:'18', color:'002060', cellColWidth: colN } },
  { val: '-',     opts: { b:true, align:'center', fontFamily:'Microsoft YaHei UI', sz:'18', color:'002060', cellColWidth: colN } },
  { val: '-',     opts: { b:true, align:'center', fontFamily:'Microsoft YaHei UI', sz:'18', color:'002060', cellColWidth: colN } },
  { val: '-',     opts: { b:true, align:'center', fontFamily:'Microsoft YaHei UI', sz:'18', color:'002060', cellColWidth: colN } },
  { val: '-',     opts: { b:true, align:'center', fontFamily:'Microsoft YaHei UI', sz:'18', color:'002060', cellColWidth: colN } },
]));

const tableData = [headerRow, ...rows];

const tableStyle = {
  tableColWidth: 1200,                 // default (usamos cellColWidth por celda)
  tableSize: 18,
  tableColor: '002060',                // color de bordes
  tableAlign: 'center',
  tableFontFamily: 'Microsoft YaHei UI',
  borders: true
};

// Crear la tabla
docx.createTable(tableData, tableStyle);

// 4) Salto de página (como Ctrl+Enter)
docx.putPageBreak();

// 5) Encabezado "ANÁLISIS" centrado y en negrita
const p8 = docx.createP({
  align: 'center',
  spacing: { before: 200, after: 100 }
});
p8.addText('ANÁLISIS', {
  bold: true,
  font_face: 'Microsoft YaHei UI',
  font_size: 15,
  color: '002060'
});

// 6) Texto de descripción con tamaño 12, justificado
const p9 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p9.addText(
  'PM10 (µg/m³)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p10 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p10.addText(
  'PM2.5 (µg/m³)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p11 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p11.addText(
  'SO₂ (PPB)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p12 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p12.addText(
  'NO₂ (PPB)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p13 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p13.addText(
  'O3 (PPB)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p14 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p14.addText(
  'CO (PPB)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p15 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p15.addText('Velocidad del viento (m/s)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p16 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p16.addText('Dirección del viento (º)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p17 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p17.addText('Precipitación (mm)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p18 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p18.addText('Presión atmosférica (mmHg)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p19 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p19.addText('Temperatura (°C)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

const p20 = docx.createP({
  align: 'justify',
  spacing: { line: 240 }
});
p20.addText('Humedad (%)',
  {
    bold: true,
    font_face: 'Microsoft YaHei UI',
    font_size: 13,
    color: '002060'
  }
);

// 8) Generar archivo
const out = fs.createWriteStream('proyecto.docx');
out.on('error', err => console.error(err));
docx.generate(out, {
  finalize: () => console.log('¡Documento creado: proyecto.docx!')
});

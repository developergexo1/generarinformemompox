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


// 8) Generar archivo
const out = fs.createWriteStream('proyecto.docx');
out.on('error', err => console.error(err));
docx.generate(out, {
  finalize: () => console.log('¡Documento creado: proyecto.docx!')
});

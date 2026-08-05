/* Genera plantilla-invitados.xlsx: llénala y luego corre generar-invitaciones.js */
const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('Invitados');

  sheet.columns = [
    { header: 'Familia', key: 'familia', width: 28 },
    { header: 'Pases', key: 'pases', width: 10 },
    { header: 'WhatsApp', key: 'whatsapp', width: 20 },
    { header: 'Trae', key: 'trae', width: 14 },
  ];
  sheet.getRow(1).font = { bold: true };

  sheet.addRows([
    { familia: 'Familia Pérez', pases: 4, whatsapp: '573001234567', trae: 'Pañales' },
    { familia: 'Familia Gómez', pases: 2, whatsapp: '', trae: 'Pañitos' },
    { familia: 'Ana y Luis', pases: 2, whatsapp: '573007654321', trae: '' },
  ]);

  const notes = wb.addWorksheet('Instrucciones');
  notes.columns = [{ width: 90 }];
  [
    'Instrucciones',
    '',
    '1) Columna "Familia": el nombre que aparecerá en el pase de la invitación (ej: "Familia Pérez", "Ana y Luis").',
    '2) Columna "Pases": número de personas invitadas de esa familia/grupo.',
    '3) Columna "WhatsApp" (opcional): número con indicativo de país sin "+" ni espacios (ej: 573001234567). Si lo llenas, el generador te crea un link listo para enviar por WhatsApp.',
    '4) Columna "Trae" (opcional): escribe "Pañales" o "Pañitos" para que la invitación de esa familia muestre ese distintivo en la portada. Si la dejas vacía, el generador alterna automáticamente entre las dos para repartir parejo.',
    '5) No cambies los nombres de las columnas ni el nombre de la hoja "Invitados".',
    '6) Agrega todas las filas que necesites, una por familia/grupo.',
    '7) Guarda el archivo y corre: node generar-invitaciones.js',
  ].forEach((line, i) => { notes.getCell(`A${i + 1}`).value = line; });
  notes.getCell('A1').font = { bold: true, size: 14 };

  await wb.xlsx.writeFile('plantilla-invitados.xlsx');
  console.log('Creado: plantilla-invitados.xlsx');
}

main();

/* Lee plantilla-invitados.xlsx (llena) y genera invitaciones-generadas.xlsx con un link por familia. */
const ExcelJS = require('exceljs');

/* Cambia esto por la URL real una vez publiques el sitio (ver README).
   Debe apuntar al archivo HTML publicado, ej:
   https://tuusuario.github.io/invitaciones/invitacion-baby-shower-samuel_3.html */
const BASE_URL = 'https://TU-DOMINIO-AQUI/invitacion-baby-shower-samuel_3.html';

async function main() {
  const inPath = 'plantilla-invitados.xlsx';
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.readFile(inPath);
  } catch (e) {
    console.error(`No se pudo leer ${inPath}. Corre primero: node crear-plantilla.js`);
    process.exit(1);
  }

  const sheet = wb.getWorksheet('Invitados');
  if (!sheet) {
    console.error('La hoja "Invitados" no existe en el archivo.');
    process.exit(1);
  }

  const out = new ExcelJS.Workbook();
  const outSheet = out.addWorksheet('Invitaciones');
  outSheet.columns = [
    { header: 'Familia', key: 'familia', width: 28 },
    { header: 'Pases', key: 'pases', width: 8 },
    { header: 'Link Invitación', key: 'link', width: 70 },
    { header: 'Link para enviar por WhatsApp', key: 'wa', width: 70 },
  ];
  outSheet.getRow(1).font = { bold: true };

  let count = 0;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // encabezado
    const familia = String(row.getCell(1).value || '').trim();
    const pases = String(row.getCell(2).value || '').trim();
    const whatsapp = String(row.getCell(3).value || '').trim();
    if (!familia) return;

    const link = `${BASE_URL}?familia=${encodeURIComponent(familia)}&pases=${encodeURIComponent(pases || '2')}`;

    let wa = '';
    if (whatsapp) {
      const msg = encodeURIComponent(`¡Hola ${familia}! Con mucha alegría los invitamos al Baby Shower de Samuel 💙. Aquí su invitación personalizada: ${link}`);
      wa = `https://wa.me/${whatsapp.replace(/\D/g, '')}?text=${msg}`;
    }

    outSheet.addRow({ familia, pases: pases || '2', link, wa });
    count++;
  });

  await out.xlsx.writeFile('invitaciones-generadas.xlsx');
  console.log(`Listo: invitaciones-generadas.xlsx (${count} invitaciones generadas)`);
  if (BASE_URL.includes('TU-DOMINIO-AQUI')) {
    console.log('\n⚠ Aún no configuraste BASE_URL en generar-invitaciones.js — edítalo con la URL real donde publiques el sitio y vuelve a correr este script.');
  }
}

main();

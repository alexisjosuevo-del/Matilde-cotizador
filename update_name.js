const xlsx = require('xlsx');
const wb = xlsx.readFile('Matilde360.xlsx');
const sheet = wb.Sheets['Catálogo'] || wb.Sheets['Servicios'];

const range = xlsx.utils.decode_range(sheet['!ref']);
let rowIndex = -1;
let colServicio = -1;

for (let C = range.s.c; C <= range.e.c; ++C) {
    const cellAddress = {c: C, r: range.s.r};
    const cellRef = xlsx.utils.encode_cell(cellAddress);
    const cell = sheet[cellRef];
    if (cell && cell.v) {
        if (cell.v === 'Servicio') colServicio = C;
    }
}

for (let R = range.s.r + 1; R <= range.e.r; ++R) {
    const cellAddress = {c: colServicio, r: R};
    const cellRef = xlsx.utils.encode_cell(cellAddress);
    const cell = sheet[cellRef];
    if (cell && cell.v && cell.v.includes('Portal de educación médica continua')) {
        rowIndex = R;
        break;
    }
}

if (rowIndex !== -1) {
    const servRef = xlsx.utils.encode_cell({c: colServicio, r: rowIndex});
    if (!sheet[servRef]) sheet[servRef] = {};
    sheet[servRef].t = 's';
    sheet[servRef].v = "Proyecto de educación médica continua. (Desarrollo, implementación y seguimiento. Duración del proyecto 12 meses)";
    
    xlsx.writeFile(wb, 'Matilde360.xlsx');
    console.log("Updated name for row " + rowIndex);
} else {
    console.log("Not found.");
}

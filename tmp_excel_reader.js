import XLSX from 'xlsx';
import fs from 'fs';

const wb = XLSX.readFile('student_phone_register_v2 (1).xlsx');
let out = [];

out.push('All Sheets: ' + JSON.stringify(wb.SheetNames));

wb.SheetNames.forEach(name => {
  const ws = wb.Sheets[name];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  out.push('');
  out.push(`=== Sheet: "${name}" | Total Rows: ${data.length} ===`);
  data.slice(0, 20).forEach((row, i) => {
    out.push(`Row ${i}: ${JSON.stringify(row)}`);
  });
});

fs.writeFileSync('tmp_excel_structure.txt', out.join('\n'), 'utf8');
console.log('Done');

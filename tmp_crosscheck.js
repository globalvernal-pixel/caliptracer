import pg from 'pg';
import XLSX from 'xlsx';
import fs from 'fs';
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 15000
});

// --- Normalize helpers ---
const normName = n => (n || '').trim().toLowerCase().replace(/[.\-\s]+/g, ' ').replace(/\s+/g, ' ').trim();
const normReg = r => String(r || '').trim().replace(/[^0-9]/g, '');

async function run() {
  let report = [];
  report.push('=== PHONE REGISTER EXCEL vs DATABASE MISMATCH REPORT ===');
  report.push(`Generated: ${new Date().toLocaleString()}`);
  report.push('');

  // 1. Read Excel
  const wb = XLSX.readFile('student_phone_register_v2 (1).xlsx');
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

  // Find header row (row with "Register No")
  let headerIdx = rows.findIndex(r => r.some && r.some(c => String(c).toLowerCase().includes('register')));
  if (headerIdx < 0) headerIdx = 2;

  const excelStudents = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const lockerNo = row[0];
    const regNo = normReg(row[1]);
    const name = (row[2] || '').toString().trim();
    const status = (row[3] || '').toString().trim();
    if (!name && !regNo) continue;
    excelStudents.push({ lockerNo, regNo, name, normNameVal: normName(name), status, rowIdx: i + 1 });
  }
  report.push(`--- EXCEL DATA SUMMARY ---`);
  report.push(`Total entries in Excel: ${excelStudents.length}`);
  report.push(`Pending: ${excelStudents.filter(s => s.status === 'Pending').length}`);
  report.push(`Verified: ${excelStudents.filter(s => s.status === 'Verified').length}`);
  report.push('');

  // 2. Fetch DB students
  const res = await pool.query('SELECT id, name, class, register_number, phone_model, phone_type FROM students ORDER BY name ASC');
  const dbStudents = res.rows;

  report.push(`--- DATABASE SUMMARY ---`);
  report.push(`Total students in Database: ${dbStudents.length}`);
  report.push('');

  // Build lookup maps
  const dbByReg = {};
  const dbByNormName = {};
  dbStudents.forEach(s => {
    const reg = normReg(s.register_number);
    if (reg) {
      if (!dbByReg[reg]) dbByReg[reg] = [];
      dbByReg[reg].push(s);
    }
    const nn = normName(s.name);
    if (!dbByNormName[nn]) dbByNormName[nn] = [];
    dbByNormName[nn].push(s);
  });

  const excelByReg = {};
  excelStudents.forEach(s => {
    if (s.regNo) {
      if (!excelByReg[s.regNo]) excelByReg[s.regNo] = [];
      excelByReg[s.regNo].push(s);
    }
  });

  // 3. CROSS CHECK: Excel entries NOT FOUND in DB
  report.push('=== [A] EXCEL ENTRIES NOT FOUND IN DATABASE ===');
  report.push('(Students in Excel register but missing from DB - either by reg number or name)');
  report.push('');
  let notInDB = [];
  excelStudents.forEach(es => {
    const foundByReg = es.regNo && dbByReg[es.regNo];
    const foundByName = dbByNormName[es.normNameVal];
    if (!foundByReg && !foundByName) {
      notInDB.push(es);
    }
  });
  if (notInDB.length === 0) {
    report.push('  No students found in Excel that are missing from DB.');
  } else {
    report.push(`  Total: ${notInDB.length} students`);
    report.push(`  ${'#'.padEnd(4)} | ${'Excel Row'.padEnd(10)} | ${'Reg No'.padEnd(10)} | ${'Excel Name'.padEnd(40)} | Status`);
    report.push(`  ${'-'.repeat(85)}`);
    notInDB.forEach((s, i) => {
      report.push(`  ${String(i+1).padEnd(4)} | Row ${String(s.rowIdx).padEnd(6)} | ${s.regNo.padEnd(10)} | ${s.name.padEnd(40)} | ${s.status}`);
    });
  }
  report.push('');

  // 4. CROSS CHECK: DB students NOT IN Excel
  report.push('=== [B] DATABASE STUDENTS NOT FOUND IN EXCEL REGISTER ===');
  report.push('(Students in DB but not in the physical phone register)');
  report.push('');
  let notInExcel = [];
  dbStudents.forEach(ds => {
    const dReg = normReg(ds.register_number);
    const dNorm = normName(ds.name);
    const foundByReg = dReg && excelByReg[dReg];
    const foundByName = excelStudents.some(es => es.normNameVal === dNorm);
    if (!foundByReg && !foundByName) {
      notInExcel.push(ds);
    }
  });
  if (notInExcel.length === 0) {
    report.push('  No DB students missing from Excel.');
  } else {
    report.push(`  Total: ${notInExcel.length} students`);
    report.push(`  ${'#'.padEnd(4)} | ${'DB ID'.padEnd(35)} | ${'Reg No'.padEnd(10)} | ${'Class'.padEnd(8)} | ${'DB Name'.padEnd(40)}`);
    report.push(`  ${'-'.repeat(100)}`);
    notInExcel.forEach((s, i) => {
      const reg = normReg(s.register_number) || 'N/A';
      report.push(`  ${String(i+1).padEnd(4)} | ${s.id.padEnd(35)} | ${reg.padEnd(10)} | ${(s.class||'').padEnd(8)} | ${s.name}`);
    });
  }
  report.push('');

  // 5. NAME MISMATCH: Same Reg No but different name
  report.push('=== [C] NAME MISMATCH - SAME REG NUMBER, DIFFERENT NAME ===');
  report.push('(Same register number exists in both places but names are different)');
  report.push('');
  let nameMismatches = [];
  excelStudents.forEach(es => {
    if (!es.regNo) return;
    const dbMatches = dbByReg[es.regNo];
    if (!dbMatches) return;
    dbMatches.forEach(ds => {
      if (normName(ds.name) !== es.normNameVal) {
        nameMismatches.push({ excelName: es.name, dbName: ds.name, regNo: es.regNo, dbClass: ds.class, dbId: ds.id, excelStatus: es.status });
      }
    });
  });
  if (nameMismatches.length === 0) {
    report.push('  No name mismatches found for matching register numbers.');
  } else {
    report.push(`  Total: ${nameMismatches.length} mismatches`);
    report.push('');
    nameMismatches.forEach((m, i) => {
      report.push(`  ${i+1}. Reg No: ${m.regNo} | Class: ${m.dbClass}`);
      report.push(`     Excel Name : "${m.excelName}" [${m.excelStatus}]`);
      report.push(`     DB Name    : "${m.dbName}" (ID: ${m.dbId})`);
      report.push('');
    });
  }

  // 6. DUPLICATE REG in Excel
  report.push('=== [D] DUPLICATE REGISTER NUMBERS IN EXCEL ===');
  const excelDupRegs = {};
  excelStudents.forEach(s => {
    if (!s.regNo) return;
    if (!excelDupRegs[s.regNo]) excelDupRegs[s.regNo] = [];
    excelDupRegs[s.regNo].push(s);
  });
  const dupRegEntries = Object.entries(excelDupRegs).filter(([_, v]) => v.length > 1);
  if (dupRegEntries.length === 0) {
    report.push('  No duplicate register numbers found in Excel.');
  } else {
    report.push(`  Total: ${dupRegEntries.length} duplicate reg groups`);
    dupRegEntries.forEach(([reg, list]) => {
      report.push(`  Reg No "${reg}" appears ${list.length} times:`);
      list.forEach(s => report.push(`    - Row ${s.rowIdx}: "${s.name}" [${s.status}]`));
    });
  }
  report.push('');

  // Write report
  fs.writeFileSync('excel_vs_db_mismatch_report.txt', report.join('\n'), 'utf8');
  console.log('Done! Report saved to excel_vs_db_mismatch_report.txt');
  await pool.end();
}

run().catch(async err => {
  console.error('Error:', err.message);
  await pool.end();
});

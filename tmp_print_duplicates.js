import pg from 'pg';
import fs from 'fs';
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 10000
});

async function analyzeDuplicates() {
  try {
    const res = await pool.query('SELECT * FROM students ORDER BY name ASC');
    const students = res.rows;

    const nameMap = {};
    const regMap = {};

    students.forEach(s => {
      const normName = (s.name || '').trim().toLowerCase();
      if (normName) {
        if (!nameMap[normName]) nameMap[normName] = [];
        nameMap[normName].push(s);
      }

      const reg = (s.register_number || '').trim();
      if (reg && reg !== '0' && reg !== 'N/A' && reg !== '') {
        if (!regMap[reg]) regMap[reg] = [];
        regMap[reg].push(s);
      }
    });

    const duplicateNames = [];
    for (const [nameKey, list] of Object.entries(nameMap)) {
      if (list.length > 1) {
        duplicateNames.push({
          name: list[0].name,
          count: list.length,
          records: list.map(s => ({
            id: s.id,
            name: s.name,
            class: s.class,
            registerNumber: s.register_number || '',
            star: s.star,
            tally: s.tally,
            fine: s.fine
          }))
        });
      }
    }

    const duplicateRegs = [];
    for (const [regKey, list] of Object.entries(regMap)) {
      if (list.length > 1) {
        duplicateRegs.push({
          registerNumber: regKey,
          count: list.length,
          records: list.map(s => ({
            id: s.id,
            name: s.name,
            class: s.class,
            star: s.star,
            tally: s.tally,
            fine: s.fine
          }))
        });
      }
    }

    const output = {
      totalStudents: students.length,
      duplicateNameGroupsCount: duplicateNames.length,
      duplicateRegGroupsCount: duplicateRegs.length,
      duplicateNames,
      duplicateRegs
    };

    fs.writeFileSync('duplicate_report.json', JSON.stringify(output, null, 2), 'utf-8');
    console.log(`Report generated successfully! Total students: ${students.length}, Duplicate Name Groups: ${duplicateNames.length}, Duplicate Reg Groups: ${duplicateRegs.length}`);

  } catch (err) {
    console.error('Error generating report:', err);
  } finally {
    await pool.end();
  }
}

analyzeDuplicates();

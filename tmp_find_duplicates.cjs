const pg = require('pg');
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 10000
});

async function findDuplicates() {
  try {
    const res = await pool.query('SELECT * FROM students ORDER BY name ASC');
    const students = res.rows;
    console.log(`TOTAL STUDENTS IN DB: ${students.length}`);

    // Map by lowercase trimmed name
    const nameMap = {};
    // Map by register number / student number
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

    console.log('\n========================================');
    console.log('DUPLICATE STUDENT NAMES IN DATABASE:');
    console.log('========================================');
    let duplicateNameCount = 0;
    for (const [nameKey, list] of Object.entries(nameMap)) {
      if (list.length > 1) {
        duplicateNameCount++;
        console.log(`\nDuplicate Group ${duplicateNameCount}: "${list[0].name}" (${list.length} records)`);
        list.forEach((s, idx) => {
          console.log(`  [Record ${idx + 1}] ID: ${s.id} | Class: ${s.class} | RegNo: "${s.register_number || ''}" | Stars: ${s.star} | Tallies: ${s.tally} | Fine: ₹${s.fine}`);
        });
      }
    }
    if (duplicateNameCount === 0) {
      console.log('No duplicate student names found in DB.');
    }

    console.log('\n========================================');
    console.log('DUPLICATE REGISTER NUMBERS IN DATABASE:');
    console.log('========================================');
    let duplicateRegCount = 0;
    for (const [reg, list] of Object.entries(regMap)) {
      if (list.length > 1) {
        duplicateRegCount++;
        console.log(`\nDuplicate RegNo Group ${duplicateRegCount}: "${reg}" (${list.length} records)`);
        list.forEach((s, idx) => {
          console.log(`  [Record ${idx + 1}] ID: ${s.id} | Name: "${s.name}" | Class: ${s.class} | Fine: ₹${s.fine}`);
        });
      }
    }
    if (duplicateRegCount === 0) {
      console.log('No duplicate register numbers found in DB.');
    }

    // Also check ROOM_STUDENT_MAPPING or hardcoded lists in App.jsx or update_user_regs.js if any!
  } catch (err) {
    console.error('Database connection / query error:', err.message);
  } finally {
    await pool.end();
  }
}

findDuplicates();

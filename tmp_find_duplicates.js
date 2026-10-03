import pg from 'pg';
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
    console.log(`TOTAL_STUDENTS_COUNT: ${students.length}`);

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

    console.log('--- DUPLICATE NAMES START ---');
    let duplicateNameCount = 0;
    for (const [nameKey, list] of Object.entries(nameMap)) {
      if (list.length > 1) {
        duplicateNameCount++;
        console.log(`[NAME_DUP] "${list[0].name}" count:${list.length}`);
        list.forEach((s) => {
          console.log(`  -> ID:${s.id} | Class:${s.class} | RegNo:${s.register_number || 'N/A'} | Fine:${s.fine} | Tally:${s.tally} | Star:${s.star}`);
        });
      }
    }
    console.log(`TOTAL_DUPLICATE_NAMES_GROUPS: ${duplicateNameCount}`);
    console.log('--- DUPLICATE NAMES END ---');

    console.log('--- DUPLICATE REG NUMBERS START ---');
    let duplicateRegCount = 0;
    for (const [reg, list] of Object.entries(regMap)) {
      if (list.length > 1) {
        duplicateRegCount++;
        console.log(`[REG_DUP] "${reg}" count:${list.length}`);
        list.forEach((s) => {
          console.log(`  -> ID:${s.id} | Name:"${s.name}" | Class:${s.class} | Fine:${s.fine}`);
        });
      }
    }
    console.log(`TOTAL_DUPLICATE_REG_GROUPS: ${duplicateRegCount}`);
    console.log('--- DUPLICATE REG NUMBERS END ---');

  } catch (err) {
    console.error('Database query error:', err);
  } finally {
    await pool.end();
  }
}

findDuplicates();

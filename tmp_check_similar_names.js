import pg from 'pg';
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 10000
});

async function checkNames() {
  const res = await pool.query('SELECT * FROM students');
  const students = res.rows;
  console.log(`Checking ${students.length} students...`);

  // Check simplified normalized names (without special chars / multiple spaces)
  const map = {};
  students.forEach(s => {
    const simplified = s.name.toLowerCase().replace(/[^a-z]/g, '');
    if (!map[simplified]) map[simplified] = [];
    map[simplified].push(s);
  });

  console.log('\n--- SIMILAR / NEAR-DUPLICATE NAMES ---');
  for (const [key, list] of Object.entries(map)) {
    if (list.length > 1) {
      console.log(`\nGroup "${key}":`);
      list.forEach(s => {
        console.log(`  ID: ${s.id} | Name: "${s.name}" | Class: ${s.class} | RegNo: "${s.register_number || ''}"`);
      });
    }
  }

  await pool.end();
}

checkNames();

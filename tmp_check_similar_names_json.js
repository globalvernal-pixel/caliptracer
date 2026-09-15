import pg from 'pg';
import fs from 'fs';
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 10000
});

async function checkNames() {
  const res = await pool.query('SELECT * FROM students');
  const students = res.rows;

  const map = {};
  students.forEach(s => {
    const simplified = s.name.toLowerCase().replace(/[^a-z]/g, '');
    if (!map[simplified]) map[simplified] = [];
    map[simplified].push(s);
  });

  const similarNameGroups = [];
  for (const [key, list] of Object.entries(map)) {
    if (list.length > 1) {
      similarNameGroups.push({
        key,
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

  fs.writeFileSync('similar_name_report.json', JSON.stringify(similarNameGroups, null, 2));
  console.log(`Found ${similarNameGroups.length} near-duplicate name groups.`);
  await pool.end();
}

checkNames();

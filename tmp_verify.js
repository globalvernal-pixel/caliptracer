import pg from 'pg';
const { Pool } = pg;
const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 15000
});
try {
  const r1 = await pool.query('SELECT COUNT(*) FROM phone_passes');
  console.log('Total passes: ' + r1.rows[0].count);
  const r2 = await pool.query('SELECT status, COUNT(*) FROM phone_passes GROUP BY status');
  r2.rows.forEach(r => console.log('  Status: ' + r.status + ' -> ' + r.count));
  const r3 = await pool.query("SELECT COUNT(*) FROM phone_register_students WHERE phone_model IS NOT NULL AND phone_model <> ''");
  console.log('With phone model: ' + r3.rows[0].count);
  const r4 = await pool.query("SELECT COUNT(*) FROM phone_register_students WHERE phone_model IS NULL OR phone_model = ''");
  console.log('Without phone model: ' + r4.rows[0].count);
} finally { await pool.end(); }

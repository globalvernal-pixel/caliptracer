import pg from 'pg';
const { Pool } = pg;
const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 15000
});
async function fix() {
  const res = await pool.query("UPDATE phone_passes SET status = 'IN' WHERE status = 'RETURNED'");
  console.log(`Updated ${res.rowCount} passes from RETURNED to IN status`);
  const counts = await pool.query('SELECT status, COUNT(*) FROM phone_passes GROUP BY status');
  counts.rows.forEach(r => console.log(`  Status: ${r.status} -> ${r.count}`));
  await pool.end();
}
fix();

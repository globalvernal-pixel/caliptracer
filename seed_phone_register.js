// Script: Import Excel phone register data into phone_register_students DB table
// Run once: node seed_phone_register.js
import XLSX from 'xlsx';
import pg from 'pg';
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 15000
});

async function seed() {
  console.log('Reading Excel file...');
  const wb = XLSX.readFile('student_phone_register_v2 (1).xlsx');
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

  // Find header row
  let headerIdx = rows.findIndex(r => r && r.some(c => String(c || '').toLowerCase().includes('register')));
  if (headerIdx < 0) headerIdx = 2;
  console.log(`Header found at row ${headerIdx}`);

  const students = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const lockerNo = String(row[0] || '').trim();
    const regNo = String(row[1] || '').trim().replace(/[^0-9]/g, '');
    const name = String(row[2] || '').trim();
    const status = String(row[3] || 'Pending').trim(); // 'Pending' or 'Verified'

    if (!name || name === 'Name') continue;
    if (!regNo) continue;

    students.push({
      lockerNo,
      registerNumber: regNo,
      name: name.toUpperCase(),
      // Pending = wants phone pass OUT (phone still in locker, needs collection)
      // Verified = phone status confirmed, shows as IN (already taken/returned)
      lockerStatus: status  // 'Pending' or 'Verified'
    });
  }

  console.log(`Total students parsed from Excel: ${students.length}`);

  // Ensure table exists
  await pool.query(`
    CREATE TABLE IF NOT EXISTS phone_register_students (
      id SERIAL PRIMARY KEY,
      locker_no VARCHAR(20) DEFAULT '',
      register_number VARCHAR(20) NOT NULL,
      name VARCHAR(255) NOT NULL,
      phone_type VARCHAR(10) DEFAULT 'school',
      locker_status VARCHAR(20) DEFAULT 'Pending',
      phone_model VARCHAR(100) DEFAULT '',
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
      updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(register_number, name)
    );
  `);

  const client = await pool.connect();
  let inserted = 0;
  let updated = 0;
  try {
    await client.query('BEGIN');
    for (const s of students) {
      const res = await client.query(`
        INSERT INTO phone_register_students (locker_no, register_number, name, locker_status)
        VALUES ($1, $2, $3, $4)
        ON CONFLICT (register_number, name)
        DO UPDATE SET
          locker_no = EXCLUDED.locker_no,
          locker_status = EXCLUDED.locker_status,
          updated_at = CURRENT_TIMESTAMP
        RETURNING (xmax = 0) AS is_insert
      `, [s.lockerNo, s.registerNumber, s.name, s.lockerStatus]);

      if (res.rows[0]?.is_insert) inserted++;
      else updated++;
    }
    await client.query('COMMIT');
    console.log(`\n✅ Done! Inserted: ${inserted} | Updated: ${updated}`);
    console.log(`Total records now in phone_register_students: ${inserted + updated}`);
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('Error during import:', err);
  } finally {
    client.release();
  }

  const countRes = await pool.query('SELECT COUNT(*) FROM phone_register_students');
  console.log(`DB total rows: ${countRes.rows[0].count}`);
  await pool.end();
}

seed().catch(async err => {
  console.error('Fatal error:', err.message);
  await pool.end();
});

import pg from 'pg';
import fs from 'fs';

const pool = new pg.Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false }
});

// Add your mapping array here:
// [{ name: 'AALIYA MARIYAM', class: 'c1a', registerNumber: '12345' }, ...]
const studentRegisterMap = [
  // PASTE OR ADD STUDENT DATA HERE
];

async function updateRegisterNumbers() {
  try {
    const allStudentsRes = await pool.query('SELECT id, name, class, register_number FROM students');
    const dbStudents = allStudentsRes.rows;
    let updatedCount = 0;

    for (const item of studentRegisterMap) {
      const searchName = item.name.trim().toLowerCase();
      const searchClass = item.class.trim().toLowerCase();
      const regNo = String(item.registerNumber).trim();

      const match = dbStudents.find(s => s.class.toLowerCase() === searchClass && s.name.trim().toLowerCase() === searchName);
      if (match) {
        await pool.query('UPDATE students SET register_number = $1 WHERE id = $2', [regNo, match.id]);
        updatedCount++;
        console.log(`Updated ${match.name} (${match.class}): #${regNo}`);
      } else {
        console.log(`Student not found: ${item.name} (${item.class})`);
      }
    }

    console.log(`\nSuccessfully updated ${updatedCount} student register numbers!`);
  } catch (err) {
    console.error('Error updating register numbers:', err);
  } finally {
    await pool.end();
  }
}

updateRegisterNumbers();

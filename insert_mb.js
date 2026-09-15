import pg from 'pg';
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false }
});

const reports = [
  { class: 's1a', name: 'FIDHA FAKRUDHEEN', topic: 'First Impression', duration: '05:47', score: 7.5, ev: 'SB' },
  { class: 'c1a', name: 'NIYA FATHIMA MP', topic: 'SUCCESS THROUGH HARDWORK', duration: '05:18', score: 4.5, ev: 'BS' },
  { class: 'c1b', name: 'AHAMMAD SHADIL', topic: 'The free miracle', duration: '04:04', score: 7.5, ev: 'MQ' },
  { class: 'c1c', name: 'AMIN RIHAN', topic: 'Drugs', duration: '04:08', score: 4, ev: 'SN' },
  { class: 'c2a', name: 'RUBIYA  S', topic: 'Cinemas Earn More From Popcorn than Movie Tickets', duration: '04:00', score: 2.5, ev: 'MM' },
  { class: 's2a', name: 'ZAREEN SHIBAS', topic: 'Anchoring effect', duration: '04:42', score: 7, ev: 'SK' },
  { class: 's1b', name: 'MUHAMMED NIHAL PB', topic: 'Butterfly effect', duration: '04:52', score: 4.5, ev: 'SK' },
  { class: 'c2c', name: 'ABDULLA KASSIM KURUNGOT', topic: 'The politics of real madrid', duration: '05:12', score: 7.5, ev: 'Mq' },
  { class: 's2b', name: 'HATHIM ZAMAN NABEEL', topic: 'Presence of being in a relationship', duration: '05:55', score: 8, ev: 'DS' }
];

async function run() {
  try {
    // Ensure table exists
    await pool.query(`
      CREATE TABLE IF NOT EXISTS morning_bliss_summary (
        id SERIAL PRIMARY KEY,
        date DATE NOT NULL UNIQUE DEFAULT CURRENT_DATE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
      ALTER TABLE students ADD COLUMN IF NOT EXISTS summary_id INT REFERENCES morning_bliss_summary(id) ON DELETE SET NULL;
    `);

    const todayStr = '2026-08-04';
    let sumRes = await pool.query('SELECT id FROM morning_bliss_summary WHERE date = $1', [todayStr]);
    let summaryId;
    if (sumRes.rows.length === 0) {
      const insRes = await pool.query('INSERT INTO morning_bliss_summary (date) VALUES ($1) RETURNING id', [todayStr]);
      summaryId = insRes.rows[0].id;
    } else {
      summaryId = sumRes.rows[0].id;
    }
    console.log('Today Summary ID:', summaryId);

    const allStudentsRes = await pool.query('SELECT * FROM students');
    const dbStudents = allStudentsRes.rows;

    for (const rep of reports) {
      const searchName = rep.name.trim().toLowerCase();
      let match = dbStudents.find(s => s.class.toLowerCase() === rep.class && s.name.trim().toLowerCase() === searchName);
      if (!match) {
        match = dbStudents.find(s => s.class.toLowerCase() === rep.class && (s.name.trim().toLowerCase().includes(searchName) || searchName.includes(s.name.trim().toLowerCase())));
      }
      
      if (match) {
        console.log(`Updating student "${match.name}" (${match.id})`);
        await pool.query(
          'UPDATE students SET morning_bliss_mark=$1, morning_bliss_topic=$2, morning_bliss_ev=$3, morning_bliss_duration=$4, summary_id=$5 WHERE id=$6',
          [rep.score, rep.topic, rep.ev, rep.duration, summaryId, match.id]
        );
      } else {
        const studentId = `${Date.now()}-${Math.floor(Math.random()*1000)}`;
        console.log(`Inserting new student "${rep.name}" (${studentId})`);
        await pool.query(
          'INSERT INTO students (id, name, class, morning_bliss_mark, morning_bliss_topic, morning_bliss_ev, morning_bliss_duration, summary_id) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)',
          [studentId, rep.name.trim(), rep.class.toLowerCase(), rep.score, rep.topic, rep.ev, rep.duration, summaryId]
        );
      }
    }

    console.log('SUCCESSFULLY inserted/updated all 9 Morning Bliss reports!');
  } catch (err) {
    console.error('Error:', err);
  } finally {
    await pool.end();
  }
}

run();

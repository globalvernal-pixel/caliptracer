import express from 'express';
import cors from 'cors';
import pg from 'pg';
const { Pool } = pg;
import path from 'path';
import compression from 'compression';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';

import dns from 'dns';
try {
  dns.setDefaultResultOrder('ipv4first');
} catch (e) {}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

dotenv.config();

const JWT_SECRET = process.env.JWT_SECRET || 'caliph_tracker_secret_2024_xk9mpl';

const app = express();
const port = process.env.PORT || 3000;

app.use(compression());
app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));
app.use(express.static(path.join(__dirname, 'dist')));

let dbInitPromise = null;
let isDbInitialized = false;

async function ensureDb(req, res, next) {
  if (!isDbInitialized) {
    if (!dbInitPromise) {
      dbInitPromise = initDb().then(() => {
        isDbInitialized = true;
      }).catch(err => {
        console.error('Database initialization error in middleware:', err);
        dbInitPromise = null;
      });
    }
    try {
      const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error('DB init timeout')), 2500));
      await Promise.race([dbInitPromise, timeoutPromise]);
    } catch (err) {
      console.warn('DB init skipped or timed out, proceeding:', err.message);
    }
  }
  next();
}

app.use(ensureDb);

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: {
    rejectUnauthorized: false
  },
  connectionTimeoutMillis: 5000,
  idleTimeoutMillis: 10000,
  max: 5
});

const INITIAL_STUDENTS = [];

// Helper to map DB row to frontend student structure
function mapRowToStudent(row) {
  const histFineCount = Number(row.history_fine_count) || 0;
  const rawFineCount = Number(row.fine_count) || 0;
  const rawFine = Number(row.fine) || 0;
  const finalFineCount = Math.max(rawFineCount, histFineCount, rawFine > 0 ? 1 : 0);
  const finalFine = Math.max(rawFine, histFineCount);

  return {
    id: row.id,
    name: row.name,
    class: row.class,
    star: row.star,
    tally: row.tally,
    starReason: row.star_reason || '',
    tallyReason: row.tally_reason || '',
    diaryStar: row.diary_star || 0,
    diaryTally: row.diary_tally || 0,
    neatAndOrderTally: row.neat_and_order_tally || 0,
    neatAndOrderReason: row.neat_and_order_reason || '',
    neatAndOrderIncidents: row.neat_and_order_incidents || 0,
    fine: finalFine,
    fineCount: finalFineCount,
    fineReason: row.fine_reason || '',
    ineligible: row.ineligible || false,
    ineligibleReason: row.ineligible_reason || '',
    morningBlissMark: row.morning_bliss_mark || null,
    morningBlissTopic: row.morning_bliss_topic || '',
    morningBlissEv: row.morning_bliss_ev || '',
    morningBlissScript: row.morning_bliss_script || null,
    morningBlissStar: row.morning_bliss_star || 0,
    morningBlissDuration: row.morning_bliss_duration || '',
    customTotal: (row.custom_total !== undefined && row.custom_total !== null) ? Number(row.custom_total) : null,
    customGrade: row.custom_grade || '',
    summaryId: row.summary_id || null,
    summaryDate: row.summary_date ? (row.summary_date instanceof Date ? row.summary_date.toISOString().split('T')[0] : String(row.summary_date).split('T')[0]) : null,
    phoneType: row.phone_type || 'school',
    phoneModel: row.phone_model || '',
    registerNumber: row.register_number || ''
  };
}

// Initialize Database Table and Auto-Seed
async function initDb() {
  try {
    // Create Students Table
    await pool.query(`
      CREATE TABLE IF NOT EXISTS students (
        id VARCHAR(255) PRIMARY KEY,
        name VARCHAR(255) NOT NULL,
        class VARCHAR(50) NOT NULL,
        star INT DEFAULT 0,
        tally INT DEFAULT 0,
        star_reason TEXT DEFAULT '',
        tally_reason TEXT DEFAULT ''
      );
    `);

    await pool.query(`ALTER TABLE students ADD COLUMN IF NOT EXISTS phone_type VARCHAR(20) DEFAULT 'school';`);
    await pool.query(`ALTER TABLE students ADD COLUMN IF NOT EXISTS phone_model VARCHAR(100) DEFAULT '';`);

    // Create History Table
    await pool.query(`
      CREATE TABLE IF NOT EXISTS student_history (
        id SERIAL PRIMARY KEY,
        student_id VARCHAR(255) NOT NULL,
        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        event_type VARCHAR(100) NOT NULL,
        amount INT DEFAULT 0,
        reason TEXT DEFAULT ''
      );
    `);

    // Create Users Table
    await pool.query(`
      CREATE TABLE IF NOT EXISTS app_users (
        id SERIAL PRIMARY KEY,
        username VARCHAR(100) UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role VARCHAR(20) DEFAULT 'user',
        permissions JSONB DEFAULT '[]',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Create App Config Table for Hostel & Room Sync across all users
    await pool.query(`
      CREATE TABLE IF NOT EXISTS app_config (
        key VARCHAR(255) PRIMARY KEY,
        data JSONB NOT NULL,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Create Phone Passes Table
    await pool.query(`
      CREATE TABLE IF NOT EXISTS phone_passes (
        id VARCHAR(255) PRIMARY KEY,
        student_id VARCHAR(255) NOT NULL,
        student_name VARCHAR(255) NOT NULL,
        student_class VARCHAR(50) NOT NULL,
        phone_model VARCHAR(100) DEFAULT '',
        student_type VARCHAR(20) DEFAULT 'school',
        start_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        allowed_until TIMESTAMP NOT NULL,
        return_time TIMESTAMP,
        reason TEXT NOT NULL,
        status VARCHAR(20) DEFAULT 'OUT',
        is_late BOOLEAN DEFAULT FALSE,
        issued_by VARCHAR(100) DEFAULT 'Admin',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Seed or upgrade default admin user to super_admin and ensure password is 1234
    const hash = await bcrypt.hash('1234', 12);
    const adminCheck = await pool.query("SELECT id, role FROM app_users WHERE username = 'admin'");
    if (adminCheck.rows.length === 0) {
      await pool.query(
        "INSERT INTO app_users (username, password_hash, role, permissions) VALUES ($1, $2, $3, $4)",
        ['admin', hash, 'super_admin', JSON.stringify(['all'])]
      );
      console.log('Default super admin user created: admin / 1234');
    } else {
      await pool.query(
        "UPDATE app_users SET password_hash = $1, role = 'super_admin', permissions = $2 WHERE username = 'admin'",
        [hash, JSON.stringify(['all'])]
      );
      console.log('Admin user password reset to 1234 with super_admin role');
    }

    // Create Morning Bliss Summary Table
    await pool.query(`
      CREATE TABLE IF NOT EXISTS morning_bliss_summary (
        id SERIAL PRIMARY KEY,
        date DATE NOT NULL UNIQUE DEFAULT CURRENT_DATE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Add new columns safely if they do not exist
    await pool.query(`
      ALTER TABLE students 
      ADD COLUMN IF NOT EXISTS diary_star INT DEFAULT 0,
      ADD COLUMN IF NOT EXISTS diary_tally INT DEFAULT 0,
      ADD COLUMN IF NOT EXISTS neat_and_order_tally INT DEFAULT 0,
      ADD COLUMN IF NOT EXISTS neat_and_order_reason TEXT DEFAULT '',
      ADD COLUMN IF NOT EXISTS neat_and_order_incidents INT DEFAULT 0,
      ADD COLUMN IF NOT EXISTS fine INT DEFAULT 0,
      ADD COLUMN IF NOT EXISTS fine_count INT DEFAULT 0,
      ADD COLUMN IF NOT EXISTS fine_reason TEXT DEFAULT '',
      ADD COLUMN IF NOT EXISTS ineligible BOOLEAN DEFAULT false,
      ADD COLUMN IF NOT EXISTS ineligible_reason TEXT DEFAULT '',
      ADD COLUMN IF NOT EXISTS morning_bliss_mark NUMERIC,
      ADD COLUMN IF NOT EXISTS morning_bliss_topic TEXT DEFAULT '',
      ADD COLUMN IF NOT EXISTS morning_bliss_ev TEXT DEFAULT '',
      ADD COLUMN IF NOT EXISTS morning_bliss_script NUMERIC,
      ADD COLUMN IF NOT EXISTS morning_bliss_star INT DEFAULT 0,
      ADD COLUMN IF NOT EXISTS phone_type VARCHAR(20) DEFAULT 'school',
      ADD COLUMN IF NOT EXISTS phone_model VARCHAR(100) DEFAULT '',
      ADD COLUMN IF NOT EXISTS register_number VARCHAR(10) DEFAULT '',
      ADD COLUMN IF NOT EXISTS summary_id INT REFERENCES morning_bliss_summary(id) ON DELETE SET NULL;
    `);

    // Ensure today's summary exists and map any existing unlinked Morning Bliss records to today's summary
    await pool.query(`
      INSERT INTO morning_bliss_summary (date) VALUES (CURRENT_DATE) ON CONFLICT (date) DO NOTHING;
      UPDATE students 
      SET summary_id = (SELECT id FROM morning_bliss_summary WHERE date = CURRENT_DATE LIMIT 1)
      WHERE morning_bliss_mark IS NOT NULL AND summary_id IS NULL;
    `);
    
    // Check if table is empty
    const checkRes = await pool.query('SELECT COUNT(*) FROM students');
    const count = parseInt(checkRes.rows[0].count);
    
    if (count === 0) {
      console.log('Database table "students" is empty. Auto-seeding initial students...');
      const client = await pool.connect();
      try {
        await client.query('BEGIN');
        for (const student of INITIAL_STUDENTS) {
          await client.query(`
            INSERT INTO students (id, name, class, star, tally, star_reason, tally_reason)
            VALUES ($1, $2, $3, $4, $5, $6, $7)
          `, [
            student.id,
            student.name,
            student.class,
            student.star,
            student.tally,
            student.starReason,
            student.tallyReason
          ]);
        }
        await client.query('COMMIT');
        console.log(`Auto-seeding complete! Inserted ${INITIAL_STUDENTS.length} students.`);
      } catch (err) {
        await client.query('ROLLBACK');
        console.error('Error during auto-seeding:', err);
      } finally {
        client.release();
      }
    } else {
      console.log(`Database is already populated with ${count} student records.`);
    }


    // Enable Row Level Security (RLS) on all public tables for Supabase security compliance
    await pool.query(`
      ALTER TABLE IF EXISTS app_users ENABLE ROW LEVEL SECURITY;
      ALTER TABLE IF EXISTS morning_bliss_summary ENABLE ROW LEVEL SECURITY;
      ALTER TABLE IF EXISTS app_config ENABLE ROW LEVEL SECURITY;
      ALTER TABLE IF EXISTS phone_passes ENABLE ROW LEVEL SECURITY;
      ALTER TABLE IF EXISTS students ENABLE ROW LEVEL SECURITY;
      ALTER TABLE IF EXISTS student_history ENABLE ROW LEVEL SECURITY;
    `);

  } catch (err) {
    console.error('Database initialization error:', err);
  }
}

// API Endpoints

// 1. Get all students
app.get('/api/students', async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT s.*, mbs.date AS summary_date,
        COALESCE((
          SELECT COUNT(*) FROM student_history h 
          WHERE h.student_id = s.id 
          AND LOWER(h.event_type) IN ('spot fine', 'room fine', 'spot_fine', 'room_fine', 'fine', 'spot', 'room')
        ), 0) AS history_fine_count
      FROM students s 
      LEFT JOIN morning_bliss_summary mbs ON s.summary_id = mbs.id
    `);
    const students = result.rows.map(mapRowToStudent);
    
    // Sort by id timestamp and roll number so it reflects the actual roll number order
    students.sort((a, b) => {
      const parseId = (id) => {
        if (!id) return { timestamp: 0, roll: 0 };
        const parts = String(id).split('-');
        let timestamp = 0;
        let roll = 0;
        for (const p of parts) {
          if (p.length === 13 && !isNaN(p)) timestamp = parseInt(p, 10);
          else if (!isNaN(p)) roll = parseInt(p, 10);
        }
        return { timestamp, roll };
      };
      const parsedA = parseId(a.id);
      const parsedB = parseId(b.id);
      if (parsedA.timestamp !== parsedB.timestamp) return parsedA.timestamp - parsedB.timestamp;
      return parsedA.roll - parsedB.roll;
    });

    res.json(students);
  } catch (err) {
    console.error(err);
    res.status(500).json({ 
      error: 'Failed to fetch students', 
      details: err.message,
      code: err.code 
    });
  }
});

// 2. Create a student
app.post('/api/students', async (req, res) => {
  const { id, name, class: className, star, tally, starReason, tallyReason } = req.body;
  try {
    const result = await pool.query(`
      INSERT INTO students (id, name, class, star, tally, star_reason, tally_reason)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING *
    `, [id, name, className, star || 0, tally || 0, starReason || '', tallyReason || '']);
    
    res.status(201).json(mapRowToStudent(result.rows[0]));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to create student' });
  }
});

// 3. Update an individual student (full object PUT)
app.put('/api/students/:id', async (req, res) => {
  const { id } = req.params;
  const { name, class: className, star, tally, starReason, tallyReason, diaryStar, diaryTally, neatAndOrderTally, neatAndOrderReason, neatAndOrderIncidents, fine, fineCount, fineReason, ineligible, ineligibleReason, customTotal, customGrade } = req.body;
  try {
    const cTotal = (customTotal !== undefined && customTotal !== null && customTotal !== '') ? Number(customTotal) : null;
    const cGrade = customGrade || '';
    const result = await pool.query(`
      UPDATE students 
      SET name = $1, class = $2, star = $3, tally = $4, star_reason = $5, tally_reason = $6, diary_star = $8, diary_tally = $9, neat_and_order_tally = $10, neat_and_order_reason = $11, neat_and_order_incidents = $17, fine = $12, fine_reason = $13, ineligible = $14, ineligible_reason = $15, fine_count = $16, custom_total = $18, custom_grade = $19
      WHERE id = $7
      RETURNING *
    `, [name, className, star, tally, starReason || '', tallyReason || '', id, diaryStar || 0, diaryTally || 0, neatAndOrderTally || 0, neatAndOrderReason || '', fine || 0, fineReason || '', ineligible || false, ineligibleReason || '', fineCount || 0, neatAndOrderIncidents || 0, cTotal, cGrade]);
    
    if (result.rows.length === 0) {
      return res.status(404).json({ error: 'Student not found' });
    }
    
    res.json(mapRowToStudent(result.rows[0]));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to update student' });
  }
});

// History Endpoints
app.post('/api/history', async (req, res) => {
  const { student_id, event_type, amount, reason } = req.body;
  if (!student_id || !event_type) return res.status(400).json({ error: 'Missing required fields' });
  try {
    const result = await pool.query(`
      INSERT INTO student_history (student_id, event_type, amount, reason)
      VALUES ($1, $2, $3, $4)
      RETURNING *
    `, [student_id, event_type, amount || 0, reason || '']);
    res.status(201).json(result.rows[0]);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to add history' });
  }
});

app.get('/api/history', async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM student_history ORDER BY date DESC');
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch history' });
  }
});

app.get('/api/history/:student_id', async (req, res) => {
  const { student_id } = req.params;
  try {
    const result = await pool.query('SELECT * FROM student_history WHERE student_id = $1 ORDER BY date DESC', [student_id]);
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch history' });
  }
});

app.put('/api/history/:id', async (req, res) => {
  const { id } = req.params;
  const { event_type, amount, reason, date, student_id } = req.body;
  try {
    const result = await pool.query(`
      UPDATE student_history
      SET event_type = COALESCE($1, event_type),
          amount = COALESCE($2, amount),
          reason = COALESCE($3, reason),
          date = COALESCE($4::timestamp, date),
          student_id = COALESCE($5, student_id)
      WHERE id = $6
      RETURNING *
    `, [event_type, amount, reason, date, student_id, id]);
    if (result.rows.length === 0) return res.status(404).json({ error: 'History record not found' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to update history' });
  }
});

app.delete('/api/history/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const result = await pool.query('DELETE FROM student_history WHERE id = $1 RETURNING *', [id]);
    if (result.rows.length === 0) return res.status(404).json({ error: 'History record not found' });
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to delete history' });
  }
});


// 3.5 Morning Bliss Summary endpoints
app.post('/api/morning-bliss/summary', async (req, res) => {
  const dateStr = req.body.date || new Date().toISOString().split('T')[0];
  try {
    const result = await pool.query(`
      INSERT INTO morning_bliss_summary (date)
      VALUES ($1)
      ON CONFLICT (date) DO UPDATE SET date = EXCLUDED.date
      RETURNING *
    `, [dateStr]);
    const row = result.rows[0];
    const formattedDate = row.date ? (row.date instanceof Date ? row.date.toISOString().split('T')[0] : String(row.date).split('T')[0]) : dateStr;
    res.json({ id: row.id, date: formattedDate });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch or create summary' });
  }
});

// 4. Bulk upsert (multiple students update/insert)
app.post('/api/students/bulk-upsert', async (req, res) => {
  const { students } = req.body;
  if (!Array.isArray(students)) {
    return res.status(400).json({ error: 'Invalid students list' });
  }
  
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    if (students.length > 0) {
      const ids = students.map(s => s.id);
      const names = students.map(s => s.name);
      const classes = students.map(s => s.class);
      const stars = students.map(s => s.star || 0);
      const tallies = students.map(s => s.tally || 0);
      const starReasons = students.map(s => s.starReason || '');
      const tallyReasons = students.map(s => s.tallyReason || '');
      const diaryStars = students.map(s => s.diaryStar || 0);
      const diaryTallies = students.map(s => s.diaryTally || 0);
      const noTallies = students.map(s => s.neatAndOrderTally || 0);
      const noReasons = students.map(s => s.neatAndOrderReason || '');
      const noIncidents = students.map(s => s.neatAndOrderIncidents || 0);
      const fines = students.map(s => s.fine || 0);
      const fineCounts = students.map(s => s.fineCount || 0);
      const fineReasons = students.map(s => s.fineReason || '');
      const ineligibles = students.map(s => s.ineligible || false);
      const ineligibleReasons = students.map(s => s.ineligibleReason || '');
      const mbMarks = students.map(s => s.morningBlissMark === undefined ? null : s.morningBlissMark);
      const mbTopics = students.map(s => s.morningBlissTopic || '');
      const mbEvs = students.map(s => s.morningBlissEv || '');
      const mbScripts = students.map(s => s.morningBlissScript === undefined ? null : s.morningBlissScript);
      const mbStars = students.map(s => s.morningBlissStar || 0);
      const mbDurations = students.map(s => s.morningBlissDuration || '');
      const customTotals = students.map(s => (s.customTotal !== undefined && s.customTotal !== null && s.customTotal !== '') ? Number(s.customTotal) : null);
      const customGrades = students.map(s => s.customGrade || '');
      const summaryIds = students.map(s => (s.summaryId !== undefined && s.summaryId !== null && s.summaryId !== '') ? Number(s.summaryId) : null);

      await client.query(`
        INSERT INTO students (id, name, class, star, tally, star_reason, tally_reason, diary_star, diary_tally, neat_and_order_tally, neat_and_order_reason, neat_and_order_incidents, fine, fine_count, fine_reason, ineligible, ineligible_reason, morning_bliss_mark, morning_bliss_topic, morning_bliss_ev, morning_bliss_script, morning_bliss_star, morning_bliss_duration, custom_total, custom_grade, summary_id)
        SELECT * FROM UNNEST($1::varchar[], $2::varchar[], $3::varchar[], $4::int[], $5::int[], $6::text[], $7::text[], $8::int[], $9::int[], $10::int[], $11::text[], $23::int[], $12::int[], $13::int[], $14::text[], $15::boolean[], $16::text[], $17::numeric[], $18::text[], $19::text[], $20::numeric[], $21::int[], $22::text[], $24::int[], $25::text[], $26::int[])
        ON CONFLICT (id) DO UPDATE SET
          name = EXCLUDED.name,
          class = EXCLUDED.class,
          star = EXCLUDED.star,
          tally = EXCLUDED.tally,
          star_reason = EXCLUDED.star_reason,
          tally_reason = EXCLUDED.tally_reason,
          diary_star = EXCLUDED.diary_star,
          diary_tally = EXCLUDED.diary_tally,
          neat_and_order_tally = EXCLUDED.neat_and_order_tally,
          neat_and_order_reason = EXCLUDED.neat_and_order_reason,
          neat_and_order_incidents = EXCLUDED.neat_and_order_incidents,
          fine = EXCLUDED.fine,
          fine_count = EXCLUDED.fine_count,
          fine_reason = EXCLUDED.fine_reason,
          ineligible = EXCLUDED.ineligible,
          ineligible_reason = EXCLUDED.ineligible_reason,
          morning_bliss_mark = EXCLUDED.morning_bliss_mark,
          morning_bliss_topic = EXCLUDED.morning_bliss_topic,
          morning_bliss_ev = EXCLUDED.morning_bliss_ev,
          morning_bliss_script = EXCLUDED.morning_bliss_script,
          morning_bliss_star = EXCLUDED.morning_bliss_star,
          morning_bliss_duration = EXCLUDED.morning_bliss_duration,
          custom_total = EXCLUDED.custom_total,
          custom_grade = EXCLUDED.custom_grade,
          summary_id = EXCLUDED.summary_id
      `, [ids, names, classes, stars, tallies, starReasons, tallyReasons, diaryStars, diaryTallies, noTallies, noReasons, fines, fineCounts, fineReasons, ineligibles, ineligibleReasons, mbMarks, mbTopics, mbEvs, mbScripts, mbStars, mbDurations, noIncidents, customTotals, customGrades, summaryIds]);
    }
    await client.query('COMMIT');
    res.json({ success: true, count: students.length });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: 'Failed to bulk upsert students' });
  } finally {
    client.release();
  }
});

// 5. Bulk score log submit (incremental addition)
app.post('/api/students/bulk-score', async (req, res) => {
  const { ids, type, amount, reason } = req.body;
  if (!Array.isArray(ids) || !['star', 'tally'].includes(type) || typeof amount !== 'number') {
    return res.status(400).json({ error: 'Invalid scoring parameters' });
  }
  
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const updateColumn = type === 'star' ? 'star' : 'tally';
    const reasonColumn = type === 'star' ? 'star_reason' : 'tally_reason';
    
    if (ids.length > 0) {
      if (reason && reason.trim()) {
        await client.query(`
          UPDATE students 
          SET ${updateColumn} = ${updateColumn} + $1, ${reasonColumn} = $2
          WHERE id = ANY($3::varchar[])
        `, [amount, reason.trim(), ids]);
      } else {
        await client.query(`
          UPDATE students 
          SET ${updateColumn} = ${updateColumn} + $1
          WHERE id = ANY($2::varchar[])
        `, [amount, ids]);
      }
    }
    await client.query('COMMIT');
    res.json({ success: true, count: ids.length });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: 'Failed to log scoring session' });
  } finally {
    client.release();
  }
});

// 6. Delete student
app.delete('/api/students/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const result = await pool.query('DELETE FROM students WHERE id = $1 RETURNING *', [id]);
    if (result.rows.length === 0) {
      return res.status(404).json({ error: 'Student not found' });
    }
    res.json({ success: true, student: mapRowToStudent(result.rows[0]) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to delete student' });
  }
});

// 6.5 Delete all students in a class
app.delete('/api/classes/:classId', async (req, res) => {
  const { classId } = req.params;
  try {
    await pool.query('DELETE FROM students WHERE class = $1', [classId]);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to delete class students' });
  }
});

// 7. Clear class rosters and import Excel roster
app.post('/api/students/import', async (req, res) => {
  const { classesToClear, newStudents } = req.body;
  if (!Array.isArray(classesToClear) || !Array.isArray(newStudents)) {
    return res.status(400).json({ error: 'Invalid import format' });
  }
  
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    
    // Clear existing rosters for classes in import set
    if (classesToClear.length > 0) {
      await client.query('DELETE FROM students WHERE class = ANY($1)', [classesToClear]);
    }
    
    // Insert new students
    if (newStudents.length > 0) {
      const ids = newStudents.map(s => s.id);
      const names = newStudents.map(s => s.name);
      const classes = newStudents.map(s => s.class);
      const stars = newStudents.map(s => s.star || 0);
      const tallies = newStudents.map(s => s.tally || 0);
      const starReasons = newStudents.map(s => s.starReason || '');
      const tallyReasons = newStudents.map(s => s.tallyReason || '');

      await client.query(`
        INSERT INTO students (id, name, class, star, tally, star_reason, tally_reason)
        SELECT * FROM UNNEST($1::varchar[], $2::varchar[], $3::varchar[], $4::int[], $5::int[], $6::text[], $7::text[])
        ON CONFLICT (id) DO NOTHING
      `, [ids, names, classes, stars, tallies, starReasons, tallyReasons]);
    }
    
    await client.query('COMMIT');
    res.json({ success: true, count: newStudents.length });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: 'Failed to import student records' });
  } finally {
    client.release();
  }
});

// 8. Clear class rosters
app.post('/api/students/clear', async (req, res) => {
  const { classesToClear } = req.body;
  if (!Array.isArray(classesToClear)) {
    return res.status(400).json({ error: 'Invalid request format' });
  }
  
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    if (classesToClear.length > 0) {
      await client.query(`
        UPDATE students 
        SET star = 0, tally = 0, star_reason = '', tally_reason = '', diary_star = 0, diary_tally = 0, neat_and_order_tally = 0, neat_and_order_reason = '', fine = 0, fine_count = 0, fine_reason = ''
        WHERE class = ANY($1)
      `, [classesToClear]);
    }
    await client.query('COMMIT');
    res.json({ success: true, count: classesToClear.length });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: 'Failed to clear class records' });
  } finally {
    client.release();
  }
});



// ─── AUTH MIDDLEWARE ────────────────────────────────────────────────────────
function requireAuth(req, res, next) {
  const auth = req.headers.authorization;
  if (!auth || !auth.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  try {
    const decoded = jwt.verify(auth.slice(7), JWT_SECRET);
    req.user = decoded;
    next();
  } catch {
    return res.status(401).json({ error: 'Invalid or expired token' });
  }
}

function requireAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.user.role !== 'admin' && req.user.role !== 'super_admin') {
      return res.status(403).json({ error: 'Admin access required' });
    }
    next();
  });
}

function requireSuperAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.user.role !== 'super_admin') {
      return res.status(403).json({ error: 'Super Admin access required for user management' });
    }
    next();
  });
}

// ─── AUTH ROUTES ────────────────────────────────────────────────────────────

// Login
app.post('/api/auth/login', async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Username and password required' });
  }
  const cleanUsername = username.trim().toLowerCase();

  try {
    const queryPromise = pool.query('SELECT * FROM app_users WHERE username = $1', [cleanUsername]);
    const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error('Query timeout')), 3000));
    const result = await Promise.race([queryPromise, timeoutPromise]);

    if (result && result.rows && result.rows.length > 0) {
      const user = result.rows[0];
      const valid = await bcrypt.compare(password, user.password_hash);
      if (valid) {
        if (user.username === 'admin' && user.role !== 'super_admin') {
          user.role = 'super_admin';
          pool.query("UPDATE app_users SET role = 'super_admin' WHERE username = 'admin'").catch(console.error);
        }
        const token = jwt.sign(
          { id: user.id, username: user.username, role: user.role, permissions: user.permissions || ['all'] },
          JWT_SECRET,
          { expiresIn: '30d' }
        );
        return res.json({
          token,
          user: { id: user.id, username: user.username, role: user.role, permissions: user.permissions || ['all'] }
        });
      }
    }
  } catch (err) {
    console.error('DB query error during login:', err.message);
  }

  // Fallback for default admin credentials if DB connection fails or times out
  if (cleanUsername === 'admin' && password === '1234') {
    const token = jwt.sign(
      { id: 1, username: 'admin', role: 'super_admin', permissions: ['all'] },
      JWT_SECRET,
      { expiresIn: '30d' }
    );
    return res.json({
      token,
      user: { id: 1, username: 'admin', role: 'super_admin', permissions: ['all'] }
    });
  }

  return res.status(401).json({ error: 'Invalid username or password' });
});

// Verify token (used on app startup)
app.get('/api/auth/verify', requireAuth, (req, res) => {
  res.json({ user: req.user });
});

// ─── USER MANAGEMENT ROUTES (super admin only) ───────────────────────────────────

// Get all users
app.get('/api/users', requireSuperAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT id, username, role, permissions, created_at FROM app_users ORDER BY created_at ASC');
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch users' });
  }
});

// Create user
app.post('/api/users', requireSuperAdmin, async (req, res) => {
  const { username, password, role, permissions } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Username and password required' });
  }
  try {
    const hash = await bcrypt.hash(password, 12);
    const result = await pool.query(
      'INSERT INTO app_users (username, password_hash, role, permissions) VALUES ($1, $2, $3, $4) RETURNING id, username, role, permissions, created_at',
      [username.trim().toLowerCase(), hash, role || 'user', JSON.stringify(permissions || [])]
    );
    res.status(201).json(result.rows[0]);
  } catch (err) {
    if (err.code === '23505') {
      return res.status(409).json({ error: 'Username already exists' });
    }
    console.error(err);
    res.status(500).json({ error: 'Failed to create user' });
  }
});

// Update user
app.put('/api/users/:id', requireSuperAdmin, async (req, res) => {
  const { id } = req.params;
  const { username, password, role, permissions } = req.body;
  try {
    let query, params;
    if (password && password.trim()) {
      const hash = await bcrypt.hash(password, 12);
      query = 'UPDATE app_users SET username=$1, password_hash=$2, role=$3, permissions=$4 WHERE id=$5 RETURNING id, username, role, permissions, created_at';
      params = [username.trim().toLowerCase(), hash, role || 'user', JSON.stringify(permissions || []), id];
    } else {
      query = 'UPDATE app_users SET username=$1, role=$2, permissions=$3 WHERE id=$4 RETURNING id, username, role, permissions, created_at';
      params = [username.trim().toLowerCase(), role || 'user', JSON.stringify(permissions || []), id];
    }
    const result = await pool.query(query, params);
    if (result.rows.length === 0) return res.status(404).json({ error: 'User not found' });
    res.json(result.rows[0]);
  } catch (err) {
    if (err.code === '23505') {
      return res.status(409).json({ error: 'Username already exists' });
    }
    console.error(err);
    res.status(500).json({ error: 'Failed to update user' });
  }
});

// Delete user
app.delete('/api/users/:id', requireSuperAdmin, async (req, res) => {
  const { id } = req.params;
  try {
    const result = await pool.query("DELETE FROM app_users WHERE id=$1 AND username != 'admin' RETURNING id", [id]);
    if (result.rows.length === 0) return res.status(404).json({ error: 'User not found or cannot delete master admin' });
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to delete user' });
  }
});

// --- Hostel & Room Configuration Synchronization Endpoints ---
app.get('/api/hostel-config', async (req, res) => {
  try {
    const result = await pool.query("SELECT key, data FROM app_config WHERE key IN ('hostel_blocks', 'hostel_block_rooms', 'room_student_mapping')");
    const config = {};
    result.rows.forEach(row => {
      config[row.key] = row.data;
    });
    res.json(config);
  } catch (err) {
    console.error('Error fetching hostel config:', err);
    res.status(500).json({ error: 'Failed to fetch hostel config' });
  }
});

app.post('/api/hostel-config', async (req, res) => {
  try {
    const { hostel_blocks, hostel_block_rooms, room_student_mapping } = req.body;

    if (hostel_blocks !== undefined) {
      await pool.query(
        "INSERT INTO app_config (key, data, updated_at) VALUES ('hostel_blocks', $1, CURRENT_TIMESTAMP) ON CONFLICT (key) DO UPDATE SET data = $1, updated_at = CURRENT_TIMESTAMP",
        [JSON.stringify(hostel_blocks)]
      );
    }
    if (hostel_block_rooms !== undefined) {
      await pool.query(
        "INSERT INTO app_config (key, data, updated_at) VALUES ('hostel_block_rooms', $1, CURRENT_TIMESTAMP) ON CONFLICT (key) DO UPDATE SET data = $1, updated_at = CURRENT_TIMESTAMP",
        [JSON.stringify(hostel_block_rooms)]
      );
    }
    if (room_student_mapping !== undefined) {
      await pool.query(
        "INSERT INTO app_config (key, data, updated_at) VALUES ('room_student_mapping', $1, CURRENT_TIMESTAMP) ON CONFLICT (key) DO UPDATE SET data = $1, updated_at = CURRENT_TIMESTAMP",
        [JSON.stringify(room_student_mapping)]
      );
    }

    res.json({ success: true });
  } catch (err) {
    console.error('Error saving hostel config:', err);
    res.status(500).json({ error: 'Failed to save hostel config' });
  }
});

// --- Phone Pass Management API Endpoints ---
app.get('/api/phone-passes', async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM phone_passes ORDER BY created_at DESC');
    const passes = result.rows.map(row => ({
      id: row.id,
      studentId: row.student_id,
      studentName: row.student_name,
      studentClass: row.student_class,
      phoneModel: row.phone_model || '',
      studentType: row.student_type || 'school',
      startTime: row.start_time,
      allowedUntil: row.allowed_until,
      returnTime: row.return_time,
      reason: row.reason,
      status: row.status || 'ISSUED',
      isLate: row.is_late || false,
      issuedBy: row.issued_by || 'Admin',
      createdAt: row.created_at
    }));
    res.json(passes);
  } catch (err) {
    console.error('Error fetching phone passes:', err);
    res.status(500).json({ error: 'Failed to fetch phone passes' });
  }
});

app.delete('/api/phone-passes/all', async (req, res) => {
  try {
    await pool.query('DELETE FROM phone_passes');
    res.json({ success: true });
  } catch (err) {
    console.error('Error clearing all phone passes:', err);
    res.status(500).json({ error: 'Failed to clear all phone passes' });
  }
});

app.post('/api/phone-passes', async (req, res) => {
  try {
    const { id, studentId, studentName, studentClass, phoneModel, studentType, startTime, allowedUntil, reason, issuedBy } = req.body;

    // Check if student already has an active pass (status 'ISSUED' or 'OUT')
    const activeCheck = await pool.query(
      "SELECT id FROM phone_passes WHERE student_id = $1 AND status IN ('ISSUED', 'OUT')",
      [studentId]
    );

    if (activeCheck.rows.length > 0) {
      return res.status(400).json({
        error: `Student ${studentName} already has an active phone pass! Must return previous pass before issuing a new one.`
      });
    }

    const passId = id || ('pass-' + Date.now());
    const startTs = startTime ? new Date(startTime) : new Date();
    const allowedTs = allowedUntil ? new Date(allowedUntil) : new Date(Date.now() + 3600000);

    const query = `
      INSERT INTO phone_passes (id, student_id, student_name, student_class, phone_model, student_type, start_time, allowed_until, reason, status, issued_by)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, 'ISSUED', $10)
      RETURNING *
    `;
    const values = [passId, studentId, studentName, studentClass, phoneModel || '', studentType || 'school', startTs, allowedTs, reason || 'Calling Home', issuedBy || 'Admin'];

    const result = await pool.query(query, values);
    const row = result.rows[0];

    res.status(201).json({
      id: row.id,
      studentId: row.student_id,
      studentName: row.student_name,
      studentClass: row.student_class,
      phoneModel: row.phone_model,
      studentType: row.student_type,
      startTime: row.start_time,
      allowedUntil: row.allowed_until,
      returnTime: row.return_time,
      reason: row.reason,
      status: row.status,
      isLate: row.is_late,
      issuedBy: row.issued_by,
      createdAt: row.created_at
    });
  } catch (err) {
    console.error('Error issuing phone pass:', err);
    res.status(500).json({ error: 'Failed to issue phone pass' });
  }
});

app.post('/api/phone-passes/batch', async (req, res) => {
  try {
    const { passes, allowedUntil, reason, issuedBy, classFilter } = req.body;

    // Option A: If server-side generation requested (passes omitted or full list)
    if (!passes || Array.isArray(passes)) {
      const passItems = passes || [];
      if (passItems.length === 0) {
        return res.status(400).json({ error: 'No passes provided' });
      }

      // Filter out students who ALREADY have an active pass
      const activeRes = await pool.query("SELECT student_id FROM phone_passes WHERE status IN ('ISSUED', 'OUT')");
      const activeStudentIds = new Set(activeRes.rows.map(r => r.student_id));

      const eligibleItems = passItems.filter(p => !activeStudentIds.has(String(p.studentId)));

      if (eligibleItems.length === 0) {
        return res.json({ count: 0, passes: [] });
      }

      // Perform single multi-row INSERT for high performance & minimal DB calls
      const valueTuples = [];
      const queryValues = [];
      let paramIdx = 1;

      const now = new Date();
      eligibleItems.forEach(p => {
        const passId = p.id || ('pass-ml-' + p.studentId + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000));
        const startTs = p.startTime ? new Date(p.startTime) : now;
        const allowedTs = p.allowedUntil ? new Date(p.allowedUntil) : new Date(now.getTime() + 3600000);

        valueTuples.push(`($${paramIdx}, $${paramIdx+1}, $${paramIdx+2}, $${paramIdx+3}, $${paramIdx+4}, $${paramIdx+5}, $${paramIdx+6}, $${paramIdx+7}, $${paramIdx+8}, 'ISSUED', $${paramIdx+9})`);
        queryValues.push(
          passId,
          p.studentId,
          p.studentName,
          p.studentClass,
          p.phoneModel || '',
          p.studentType || 'school',
          startTs,
          allowedTs,
          p.reason || 'Monthly Leave',
          p.issuedBy || issuedBy || 'Admin'
        );
        paramIdx += 10;
      });

      const batchQuery = `
        INSERT INTO phone_passes (id, student_id, student_name, student_class, phone_model, student_type, start_time, allowed_until, reason, status, issued_by)
        VALUES ${valueTuples.join(', ')}
        RETURNING *
      `;

      const result = await pool.query(batchQuery, queryValues);
      const createdPasses = result.rows.map(row => ({
        id: row.id,
        studentId: row.student_id,
        studentName: row.student_name,
        studentClass: row.student_class,
        phoneModel: row.phone_model,
        studentType: row.student_type,
        startTime: row.start_time,
        allowedUntil: row.allowed_until,
        returnTime: row.return_time,
        reason: row.reason,
        status: row.status,
        isLate: row.is_late,
        issuedBy: row.issued_by,
        createdAt: row.created_at
      }));

      return res.status(201).json({ count: createdPasses.length, passes: createdPasses });
    }
  } catch (err) {
    console.error('Error batch issuing phone passes:', err);
    res.status(500).json({ error: 'Failed to batch issue phone passes: ' + err.message });
  }
});

app.put('/api/phone-passes/:id/pickup', async (req, res) => {
  try {
    const { id } = req.params;
    const now = new Date();

    const passCheck = await pool.query('SELECT start_time, allowed_until FROM phone_passes WHERE id = $1', [id]);
    if (passCheck.rows.length === 0) {
      return res.status(404).json({ error: 'Phone pass not found' });
    }

    const oldStart = new Date(passCheck.rows[0].start_time);
    const oldAllowed = new Date(passCheck.rows[0].allowed_until);
    const durationMs = oldAllowed.getTime() - oldStart.getTime();
    const newAllowed = new Date(now.getTime() + (durationMs > 0 ? durationMs : 3600000));

    const result = await pool.query(
      "UPDATE phone_passes SET status = 'OUT', start_time = $1, allowed_until = $2 WHERE id = $3 RETURNING *",
      [now, newAllowed, id]
    );

    const row = result.rows[0];
    res.json({
      id: row.id,
      studentId: row.student_id,
      studentName: row.student_name,
      studentClass: row.student_class,
      phoneModel: row.phone_model,
      studentType: row.student_type,
      startTime: row.start_time,
      allowedUntil: row.allowed_until,
      returnTime: row.return_time,
      reason: row.reason,
      status: row.status,
      isLate: row.is_late,
      issuedBy: row.issued_by,
      createdAt: row.created_at
    });
  } catch (err) {
    console.error('Error marking phone pass pickup (OUT):', err);
    res.status(500).json({ error: 'Failed to pickup phone pass' });
  }
});

app.put('/api/phone-passes/:id/return', async (req, res) => {
  try {
    const { id } = req.params;
    const now = new Date();

    const passCheck = await pool.query('SELECT allowed_until FROM phone_passes WHERE id = $1', [id]);
    if (passCheck.rows.length === 0) {
      return res.status(404).json({ error: 'Phone pass not found' });
    }

    const allowedUntil = new Date(passCheck.rows[0].allowed_until);
    const isLate = now > allowedUntil;

    const result = await pool.query(
      "UPDATE phone_passes SET status = 'IN', return_time = $1, is_late = $2 WHERE id = $3 RETURNING *",
      [now, isLate, id]
    );

    const row = result.rows[0];
    res.json({
      id: row.id,
      studentId: row.student_id,
      studentName: row.student_name,
      studentClass: row.student_class,
      phoneModel: row.phone_model,
      studentType: row.student_type,
      startTime: row.start_time,
      allowedUntil: row.allowed_until,
      returnTime: row.return_time,
      reason: row.reason,
      status: row.status,
      isLate: row.is_late,
      issuedBy: row.issued_by,
      createdAt: row.created_at
    });
  } catch (err) {
    console.error('Error marking phone pass returned (IN):', err);
    res.status(500).json({ error: 'Failed to return phone pass' });
  }
});

app.delete('/api/phone-passes/:id', async (req, res) => {
  try {
    const { id } = req.params;
    await pool.query('DELETE FROM phone_passes WHERE id = $1', [id]);
    res.json({ success: true });
  } catch (err) {
    console.error('Error deleting phone pass:', err);
    res.status(500).json({ error: 'Failed to delete phone pass' });
  }
});

app.put('/api/students/:id/phone-type', async (req, res) => {
  try {
    const { id } = req.params;
    const { phone_type, phone_model, register_number } = req.body;

    const result = await pool.query(
      'UPDATE students SET phone_type = $1, phone_model = $2, register_number = $3 WHERE id = $4 RETURNING *',
      [phone_type || 'school', phone_model || '', register_number || '', id]
    );

    if (result.rows.length === 0) return res.status(404).json({ error: 'Student not found' });
    res.json(mapRowToStudent(result.rows[0]));
  } catch (err) {
    console.error('Error updating student phone info:', err);
    res.status(500).json({ error: 'Failed to update student phone info' });
  }
});

// Initialize database and start the listening loop
if (process.argv[1] === __filename) {
  initDb().then(() => {
    app.listen(port, () => {
      console.log(`Backend server successfully listening on port ${port}`);
    });
  });
} else {
  // When imported as a module (e.g. on Vercel), just init the DB
  initDb().catch(console.error);
}

export default app;

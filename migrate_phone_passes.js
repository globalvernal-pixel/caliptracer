// Migration Script v2 — BATCH inserts for speed
import pg from 'pg';
const { Pool } = pg;

const pool = new Pool({
  connectionString: 'postgresql://postgres.xnkzwagnibshwcjenvxh:frrVFypkPqsnx7tl@aws-1-ap-southeast-2.pooler.supabase.com:6543/postgres',
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 20000,
  idleTimeoutMillis: 30000
});

const normName = n => (n || '').trim().toLowerCase().replace(/\s+/g, ' ');
const normReg  = r => String(r || '').trim().replace(/[^0-9]/g, '');

async function main() {
  console.log('\n═══ PHONE PASS MIGRATION v2 ═══\n');

  // ── STEP 1: Clear all phone passes ─────────────────────────────
  console.log('[1] Clearing phone_passes...');
  await pool.query('DELETE FROM phone_passes');
  console.log('    Done.\n');

  // ── STEP 2: Fetch phone register students ───────────────────────
  console.log('[2] Fetching phone_register_students...');
  const { rows: regStudents } = await pool.query('SELECT * FROM phone_register_students ORDER BY id ASC');
  console.log(`    → ${regStudents.length} students.\n`);

  // ── STEP 3: Fetch old students table phone models ───────────────
  console.log('[3] Fetching phone models from old students table...');
  const { rows: oldStudents } = await pool.query(
    "SELECT name, register_number, phone_model FROM students WHERE phone_model IS NOT NULL AND phone_model <> ''"
  );
  console.log(`    → ${oldStudents.length} students have phone models.\n`);

  // Build lookup maps
  const byName = {};
  const byReg  = {};
  oldStudents.forEach(s => {
    const nn  = normName(s.name);
    const reg = normReg(s.register_number);
    if (nn)  byName[nn]  = s.phone_model;
    if (reg) byReg[reg]  = s.phone_model;
  });

  // ── STEP 4: Build update SQL for phone models ───────────────────
  console.log('[4] Updating phone models in batch...');
  const modelUpdates = [];
  regStudents.forEach(rs => {
    const model = byName[normName(rs.name)] || byReg[normReg(rs.register_number)];
    if (model && model.trim() !== (rs.phone_model || '').trim()) {
      modelUpdates.push({ id: rs.id, model });
    }
  });

  if (modelUpdates.length > 0) {
    // Build a single multi-row UPDATE
    const caseWhen = modelUpdates.map((u, i) => `WHEN id = ${u.id} THEN $${i + 1}`).join('\n      ');
    const ids      = modelUpdates.map(u => u.id);
    const models   = modelUpdates.map(u => u.model);
    await pool.query(`
      UPDATE phone_register_students SET phone_model = CASE
        ${caseWhen}
        ELSE phone_model
      END,
      updated_at = CURRENT_TIMESTAMP
      WHERE id = ANY($${models.length + 1})
    `, [...models, ids]);
    console.log(`    → Updated phone models for ${modelUpdates.length} students.\n`);
  } else {
    console.log('    → No phone models to update.\n');
  }

  // Refetch with updated models
  const { rows: finalStudents } = await pool.query('SELECT * FROM phone_register_students ORDER BY id ASC');

  // ── STEP 5: Build phone passes in two bulk inserts ──────────────
  console.log('[5] Building phone passes...');
  const now      = new Date();
  const nowISO   = now.toISOString();
  const allowed  = new Date(now.getTime() + 24 * 60 * 60 * 1000).toISOString();

  const pendingStudents  = finalStudents.filter(s => s.locker_status === 'Pending');
  const verifiedStudents = finalStudents.filter(s => s.locker_status === 'Verified');

  console.log(`    Pending → OUT  : ${pendingStudents.length}`);
  console.log(`    Verified → IN  : ${verifiedStudents.length}`);

  // Helper: build bulk insert SQL
  function buildBulkInsert(studentList, status, hasReturnTime) {
    if (studentList.length === 0) return null;
    const colNames = hasReturnTime
      ? '(id, student_id, student_name, student_class, phone_model, student_type, start_time, allowed_until, return_time, reason, status, is_late, issued_by, created_at)'
      : '(id, student_id, student_name, student_class, phone_model, student_type, start_time, allowed_until, reason, status, is_late, issued_by, created_at)';

    const rows   = [];
    const values = [];
    let   p      = 1;

    studentList.forEach((rs, i) => {
      const passId    = `preg-${rs.id}-${Date.now() + i}`;
      const stuId     = `preg-${rs.id}`;
      const model     = rs.phone_model || '';
      const pType     = rs.phone_type || 'school';
      const reason    = status === 'OUT' ? 'Phone Pass (Pending)' : 'Phone Pass (Verified)';

      if (hasReturnTime) {
        rows.push(`($${p},$${p+1},$${p+2},$${p+3},$${p+4},$${p+5},$${p+6},$${p+7},$${p+8},$${p+9},$${p+10},$${p+11},$${p+12},$${p+13})`);
        values.push(passId, stuId, rs.name, '', model, pType, nowISO, allowed, nowISO, reason, status, false, 'System Migration', nowISO);
        p += 14;
      } else {
        rows.push(`($${p},$${p+1},$${p+2},$${p+3},$${p+4},$${p+5},$${p+6},$${p+7},$${p+8},$${p+9},$${p+10},$${p+11},$${p+12})`);
        values.push(passId, stuId, rs.name, '', model, pType, nowISO, allowed, reason, status, false, 'System Migration', nowISO);
        p += 13;
      }
    });

    const sql = `INSERT INTO phone_passes ${colNames} VALUES ${rows.join(',')} ON CONFLICT (id) DO NOTHING`;
    return { sql, values };
  }

  // ── Chunk helper (avoid too-large queries) ──
  async function insertChunked(studentList, status, hasReturn, chunkSize = 50) {
    let count = 0;
    for (let i = 0; i < studentList.length; i += chunkSize) {
      const chunk = studentList.slice(i, i + chunkSize);
      const built = buildBulkInsert(chunk, status, hasReturn);
      if (built) {
        await pool.query(built.sql, built.values);
        count += chunk.length;
        process.stdout.write(`\r    → Inserted ${count}/${studentList.length}...`);
      }
    }
    console.log('');
    return count;
  }

  console.log('\n  Inserting OUT passes (Pending):');
  const outInserted = await insertChunked(pendingStudents, 'OUT', false, 50);

  console.log('  Inserting IN passes (Verified):');
  const inInserted  = await insertChunked(verifiedStudents, 'RETURNED', true, 50);

  // ── Final summary ──────────────────────────────────────────────
  const { rows: [{ count }] } = await pool.query('SELECT COUNT(*) FROM phone_passes');
  const noModelCount = finalStudents.filter(s => !s.phone_model || !s.phone_model.trim()).length;

  console.log('\n═══════════════════════════════════════');
  console.log(' MIGRATION COMPLETE ✅');
  console.log('═══════════════════════════════════════');
  console.log(`  Phone passes total    : ${count}`);
  console.log(`    OUT (Pending)       : ${outInserted}`);
  console.log(`    IN  (Verified)      : ${inInserted}`);
  console.log(`  Phone models updated  : ${modelUpdates.length}`);
  console.log(`  Students w/o model    : ${noModelCount}`);
  console.log('═══════════════════════════════════════\n');

  await pool.end();
}

main().catch(async err => {
  console.error('\nFatal error:', err.message);
  await pool.end();
  process.exit(1);
});

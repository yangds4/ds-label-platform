const express = require('express');
const http = require('http');
const https = require('https');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const fsp = require('fs/promises');
const crypto = require('crypto');
const config = require('./config');
const db = require('./db');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 }
});

// multer 默认按 latin1 读取上传文件名，中文 UTF-8 文件名会出现双重编码（如「房屋瑕疵」→「æ¿å±ç¥çµ」）。
// 写入数据库前用 latin1→utf8 反向解码还原成正确文件名。
function fixOriginalName(name) {
  if (!name) return '';
  try {
    const buf = Buffer.from(name, 'latin1');
    const decoded = buf.toString('utf8');
    // 验证解码后是否含中文字符或常见后缀；否则说明不是双重编码，按原值返回
    if (/[\u4e00-\u9fa5]/.test(decoded) || /\.(xlsx|xls|csv)$/i.test(decoded)) return decoded;
  } catch {}
  return name;
}

// 表尾的非标注列（不作为标注项）
const TRAILING_COLS = new Set(['人工备注', '标注人', '标注日期', '备注']);
// 定位/辅助列（不作为标注项，仅用于任务定位与导出）
const META_COLS = new Set(['full_image_url', 'prospecting_id', 'cell_name']);
// 表头的定位列
const URL_COL = 'full_image_url';

function ok(res, data) { res.json({ ok: true, data }); }
function fail(res, msg, code) { res.status(code || 400).json({ ok: false, msg }); }
// 按本地时区取 YYYY-MM-DD（兼容 MySQL 返回的 Date 与 SQLite 返回的字符串）
function localDate(v) {
  if (!v) return '';
  const d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v).slice(0, 10);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/* ==================== 管理员：导入 Excel ==================== */
app.post('/api/admin/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return fail(res, '请上传 Excel 文件');
    const batchName = (req.body.batchName || '').trim() || ('批次_' + new Date().toISOString().slice(0, 16).replace(/[-T:]/g, ''));

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!rows.length) return fail(res, 'Excel 为空');

    const header = rows[0].map(h => (h == null ? '' : String(h).trim()));
    const urlIdx = header.indexOf(URL_COL);
    if (urlIdx === -1) return fail(res, `未找到 ${URL_COL} 列，请检查表头`);

    // 动态识别标注列：full_image_url 之后、且不属于表尾备注列的所有列
    const labelColumns = [];
    for (let i = urlIdx + 1; i < header.length; i++) {
      const h = header[i];
      if (!h || TRAILING_COLS.has(h) || META_COLS.has(h)) continue;
      labelColumns.push(h);
    }
    if (!labelColumns.length) return fail(res, '未识别到任何标注列');

    const pidIdx = header.indexOf('prospecting_id');
    const cellIdx = header.indexOf('cell_name');

    // 严格按 Excel 每一行生成一条任务，不做任何去重（同一图片可重复出现，逐行保留）
    const dataRows = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.every(v => v == null || v === '')) continue;
      const url = row[urlIdx] == null ? '' : String(row[urlIdx]).trim();
      if (!url) continue;
      const rowObj = {};
      for (let c = 0; c < header.length; c++) {
        if (header[c]) rowObj[header[c]] = row[c] == null ? '' : row[c];
      }
      const pid  = pidIdx  >= 0 && row[pidIdx]  != null ? String(row[pidIdx])  : '';
      const cell = cellIdx >= 0 && row[cellIdx] != null ? String(row[cellIdx]) : '';
      dataRows.push({ rowIndex: r, url, pid, cell, data: rowObj });
    }
    if (!dataRows.length) return fail(res, '没有有效数据行');

    const pool = db.getPool();
    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();
      const [br] = await conn.query(
        'INSERT INTO batches (name, source_file, status, label_columns, total_tasks) VALUES (?,?,?,?,?)',
        [batchName, fixOriginalName(req.file.originalname || ''), 'active', JSON.stringify(labelColumns), dataRows.length]
      );
      const batchId = br.insertId;

      // 1) 先插入原始行，每行拿到自增 id，用于任务与行一一对应
      const CHUNK = 1000;
      for (let i = 0; i < dataRows.length; i += CHUNK) {
        const vals = dataRows.slice(i, i + CHUNK).map(d => [batchId, d.rowIndex, d.url, JSON.stringify(d.data)]);
        await conn.query('INSERT INTO excel_rows (batch_id, row_index, full_image_url, data) VALUES ?', [vals]);
      }
      const [er] = await conn.query('SELECT id, row_index FROM excel_rows WHERE batch_id=? ORDER BY row_index', [batchId]);
      const idByRow = new Map();
      for (const x of er) idByRow.set(x.row_index, x.id);

      // 2) 每条 Excel 行 -> 一条任务（1:1），excel_row 指向对应的 excel_rows.id
      const taskValues = [];
      for (const d of dataRows) {
        taskValues.push([batchId, d.url, d.pid, d.cell, 1, idByRow.get(d.rowIndex) || 0]);
      }
      for (let i = 0; i < taskValues.length; i += CHUNK) {
        await conn.query(
          'INSERT INTO tasks (batch_id, full_image_url, prospecting_id, cell_name, row_count, excel_row) VALUES ?',
          [taskValues.slice(i, i + CHUNK)]
        );
      }
      await conn.commit();
      ok(res, { batchId, batchName, totalTasks: dataRows.length, totalRows: dataRows.length, labelColumns });
    } catch (e) {
      await conn.rollback();
      throw e;
    } finally {
      conn.release();
    }
  } catch (e) {
    console.error(e);
    fail(res, '导入失败: ' + e.message, 500);
  }
});

/* ==================== 管理员：批次列表与进度 ==================== */
app.get('/api/admin/batches', async (req, res) => {
  try {
    const pool = db.getPool();
    const [batches] = await pool.query('SELECT * FROM batches ORDER BY id DESC');
    const [stats] = await pool.query(
      `SELECT batch_id,
              SUM(status='pending') AS pending,
              SUM(status='claimed') AS claimed,
              SUM(status='submitted') AS submitted
       FROM tasks GROUP BY batch_id`
    );
    const map = {};
    for (const s of stats) map[s.batch_id] = s;
    ok(res, batches.map(b => ({
      ...b,
      label_columns: typeof b.label_columns === 'string' ? JSON.parse(b.label_columns) : b.label_columns,
      pending: Number(map[b.id]?.pending || 0),
      claimed: Number(map[b.id]?.claimed || 0),
      submitted: Number(map[b.id]?.submitted || 0)
    })));
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 管理员：删除批次（任务 + Excel行 + 批次本身） ==================== */
app.delete('/api/admin/batches/:id', async (req, res) => {
  let conn;
  try {
    const pool = db.getPool();
    conn = await pool.getConnection();
    await conn.beginTransaction();
    const id = Number(req.params.id);
    if (!id) { await conn.rollback(); return fail(res, 'id 不合法'); }

    const [brows] = await conn.query('SELECT id, name, total_tasks FROM batches WHERE id=?', [id]);
    if (!brows.length) { await conn.rollback(); return fail(res, '批次不存在', 404); }
    const batch = brows[0];

    const [tstats] = await conn.query(
      `SELECT SUM(status='pending') AS pending,
              SUM(status='claimed') AS claimed,
              SUM(status='submitted') AS submitted,
              COUNT(*) AS total
       FROM tasks WHERE batch_id=?`, [id]
    );
    const s = tstats[0] || {};
    const pending   = Number(s.pending || 0);
    const claimed   = Number(s.claimed || 0);
    const submitted = Number(s.submitted || 0);
    const total     = Number(s.total || 0);

    // 顺序删除；SQLite 没有 ON DELETE CASCADE，所以手动；MySQL 有 cascade 也兼容
    const [erows] = await conn.query('DELETE FROM excel_rows WHERE batch_id=?', [id]);
    const [trows] = await conn.query('DELETE FROM tasks WHERE batch_id=?', [id]);
    const [bd]    = await conn.query('DELETE FROM batches WHERE id=?', [id]);
    if (!bd.affectedRows) { await conn.rollback(); return fail(res, '批次已被他人删除', 404); }
    await conn.commit();

    ok(res, {
      id,
      name: batch.name,
      deletedTasks: trows.affectedRows,
      deletedExcelRows: erows.affectedRows,
      hadTotal: total,
      hadSubmitted: submitted,
      hadClaimed: claimed,
      hadPending: pending
    });
  } catch (e) {
    try { if (conn) await conn.rollback(); } catch {}
    fail(res, e.message, 500);
  } finally {
    try { if (conn) conn.release(); } catch {}
  }
});

/* ==================== 管理员：批次状态（暂停/恢复/停止） ==================== */
app.post('/api/admin/batches/:id/status', async (req, res) => {
  try {
    const { status } = req.body;
    if (!['active', 'paused', 'stopped'].includes(status)) return fail(res, '状态不合法');
    const pool = db.getPool();
    const [r] = await pool.query('UPDATE batches SET status=? WHERE id=?', [status, req.params.id]);
    if (!r.affectedRows) return fail(res, '批次不存在', 404);
    ok(res, { id: Number(req.params.id), status });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 管理员：标注员统计 ==================== */
app.get('/api/admin/annotators', async (req, res) => {
  try {
    const pool = db.getPool();
    const where = req.query.batchId ? 'WHERE batch_id=' + Number(req.query.batchId) : '';
    const [rows] = await pool.query(
      `SELECT assignee,
              SUM(status='claimed') AS claimed,
              SUM(status='submitted') AS submitted
       FROM tasks ${where}
       ${where ? 'AND' : 'WHERE'} assignee IS NOT NULL
       GROUP BY assignee ORDER BY submitted DESC`
    );
    ok(res, rows.map(r => ({ assignee: r.assignee, claimed: Number(r.claimed), submitted: Number(r.submitted) })));
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 管理员：释放组（未提交的退回待领取） ==================== */
app.post('/api/admin/release', async (req, res) => {
  try {
    const { batchId, assignee, taskIds } = req.body;
    const pool = db.getPool();
    const conds = ["status='claimed'"];
    const params = [];
    if (batchId) { conds.push('batch_id=?'); params.push(batchId); }
    if (assignee) { conds.push('assignee=?'); params.push(assignee); }
    if (Array.isArray(taskIds) && taskIds.length) { conds.push(`id IN (${taskIds.map(() => '?').join(',')})`); params.push(...taskIds); }
    if (params.length === 0) return fail(res, '必须指定批次、标注员或任务ID');
    const [r] = await pool.query(
      `UPDATE tasks SET status='pending', assignee=NULL, claimed_at=NULL, draft=NULL WHERE ${conds.join(' AND ')}`,
      params
    );
    ok(res, { released: r.affectedRows });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 管理员：转派组 ==================== */
app.post('/api/admin/reassign', async (req, res) => {
  try {
    const { from, to, batchId, taskIds } = req.body;
    if (!to || !String(to).trim()) return fail(res, '必须指定接收人');
    const pool = db.getPool();
    const conds = ["status='claimed'"];
    const params = [String(to).trim()];
    if (from) { conds.push('assignee=?'); params.push(from); }
    if (batchId) { conds.push('batch_id=?'); params.push(batchId); }
    if (Array.isArray(taskIds) && taskIds.length) { conds.push(`id IN (${taskIds.map(() => '?').join(',')})`); params.push(...taskIds); }
    if (params.length === 1) return fail(res, '必须指定来源标注员、批次或任务ID');
    const [r] = await pool.query(
      `UPDATE tasks SET assignee=?, claimed_at=NOW() WHERE ${conds.join(' AND ')}`,
      params
    );
    ok(res, { reassigned: r.affectedRows });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 管理员：标签统计（每个标注项的 是/否/无法判断 分布） ==================== */
app.get('/api/admin/label-stats/:batchId', async (req, res) => {
  try {
    const batchId = Number(req.params.batchId);
    const pool = db.getPool();
    const [[batch]] = await pool.query('SELECT label_columns FROM batches WHERE id=?', [batchId]);
    if (!batch) return fail(res, '批次不存在', 404);
    const labelColumns = typeof batch.label_columns === 'string' ? JSON.parse(batch.label_columns) : batch.label_columns;

    // 查询该批次所有已提交任务的 answers
    const [tasks] = await pool.query(
      "SELECT answers FROM tasks WHERE batch_id=? AND status='submitted'",
      [batchId]
    );

    // 逐标签聚合
    const stats = {};
    for (const col of labelColumns) {
      stats[col] = { '是': 0, '否': 0, '无法判断': 0 };
    }
    let totalSubmitted = 0;
    for (const t of tasks) {
      totalSubmitted++;
      const ans = typeof t.answers === 'string' ? JSON.parse(t.answers || '{}') : (t.answers || {});
      for (const col of labelColumns) {
        const v = ans[col];
        if (v && stats[col][v] !== undefined) stats[col][v]++;
      }
    }

    ok(res, { batchId, totalSubmitted, labelColumns, stats });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 管理员：查看已提交数据 ==================== */
app.get('/api/admin/submitted', async (req, res) => {
  try {
    const pool = db.getPool();
    const batchId = req.query.batchId ? Number(req.query.batchId) : null;
    const assignee = (req.query.assignee || '').trim();
    const kw = (req.query.kw || '').trim();
    const page = Math.max(1, Number(req.query.page) || 1);
    const pageSize = Math.min(Math.max(Number(req.query.pageSize) || 20, 5), 100);

    const conds = ["t.status='submitted'"];
    const params = [];
    if (batchId) { conds.push('t.batch_id=?'); params.push(batchId); }
    if (assignee) { conds.push('t.assignee=?'); params.push(assignee); }
    if (kw) {
      conds.push('(t.cell_name LIKE ? OR t.prospecting_id LIKE ? OR t.full_image_url LIKE ?)');
      const like = `%${kw}%`;
      params.push(like, like, like);
    }
    const where = conds.join(' AND ');

    const [[cnt]] = await pool.query(
      `SELECT COUNT(*) AS n FROM tasks t WHERE ${where}`, params
    );
    const total = Number(cnt.n || 0);

    const [rows] = await pool.query(
      `SELECT t.id, t.batch_id, t.prospecting_id, t.cell_name, t.full_image_url,
              t.answers, t.remark, t.assignee, t.submitted_at, b.name AS batch_name
       FROM tasks t JOIN batches b ON t.batch_id=b.id
       WHERE ${where}
       ORDER BY t.submitted_at DESC, t.id DESC
       LIMIT ? OFFSET ?`,
      [...params, pageSize, (page - 1) * pageSize]
    );

    const list = rows.map(r => {
      const answers = typeof r.answers === 'string' ? JSON.parse(r.answers || '{}') : (r.answers || {});
      let yes = 0, no = 0, na = 0;
      for (const v of Object.values(answers)) {
        if (v === '是') yes++; else if (v === '否') no++; else if (v === '无法判断') na++;
      }
      return {
        id: r.id, batch_id: r.batch_id, batch_name: r.batch_name,
        prospecting_id: r.prospecting_id, cell_name: r.cell_name,
        full_image_url: r.full_image_url, assignee: r.assignee,
        submitted_at: r.submitted_at ? String(r.submitted_at instanceof Date ? localDate(r.submitted_at) + ' ' + r.submitted_at.toTimeString().slice(0, 8) : r.submitted_at).replace('T', ' ').slice(0, 19) : '',
        remark: r.remark || '', answers, yes, no, na
      };
    });

    ok(res, { total, page, pageSize, rows: list });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 管理员：导出结果（可筛选 已提交/未提交/全部） ==================== */
app.get('/api/admin/export/:batchId', async (req, res) => {
  try {
    const batchId = Number(req.params.batchId);
    // filter: all(默认) | submitted(仅已提交) | unsubmitted(仅未提交)
    const filter = ['submitted', 'unsubmitted'].includes(req.query.filter) ? req.query.filter : 'all';
    const pool = db.getPool();
    const [[batch]] = await pool.query('SELECT * FROM batches WHERE id=?', [batchId]);
    if (!batch) return fail(res, '批次不存在', 404);
    const labelColumns = typeof batch.label_columns === 'string' ? JSON.parse(batch.label_columns) : batch.label_columns;

    const [tasks] = await pool.query(
      "SELECT full_image_url, excel_row, answers, remark, assignee, submitted_at, status FROM tasks WHERE batch_id=?",
      [batchId]
    );
    const ansMap = new Map();
    for (const t of tasks) {
      if (t.status !== 'submitted') continue;
      const answers = typeof t.answers === 'string' ? JSON.parse(t.answers || '{}') : (t.answers || {});
      const a = { answers, remark: t.remark || '', assignee: t.assignee || '', submitted_at: t.submitted_at };
      if (t.excel_row) ansMap.set('r' + t.excel_row, a);   // 新数据：按 excel_rows 行 id 精确匹配（支持同 URL 多行）
      ansMap.set('u' + t.full_image_url, a);               // 旧数据：按 url 兼容（旧批次已去重，不会冲突）
    }

    const [rows] = await pool.query(
      'SELECT id, row_index, full_image_url, data FROM excel_rows WHERE batch_id=? ORDER BY row_index',
      [batchId]
    );

    const header = ['prospecting_id', 'cell_name', 'full_image_url', ...labelColumns, '人工备注', '标注人', '标注日期'];
    const out = [header];
    for (const r of rows) {
      const a = ansMap.get('r' + r.id) || ansMap.get('u' + r.full_image_url);
      // 按筛选条件跳过不需要的行
      if (filter === 'submitted' && !a) continue;
      if (filter === 'unsubmitted' && a) continue;
      const data = typeof r.data === 'string' ? JSON.parse(r.data) : r.data;
      const line = [data['prospecting_id'] ?? '', data['cell_name'] ?? '', r.full_image_url];
      for (const col of labelColumns) line.push(a ? (a.answers[col] ?? '') : '');
      line.push(a ? a.remark : '');
      line.push(a ? a.assignee : '');
      line.push(a && a.submitted_at ? localDate(a.submitted_at) : '');
      out.push(line);
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(out);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const suffix = filter === 'submitted' ? '_仅已提交' : (filter === 'unsubmitted' ? '_仅未提交' : '');
    const fname = encodeURIComponent(`标注结果_${batch.name}${suffix}.xlsx`);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fname}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 标注员：领取组 ==================== */
app.post('/api/claim', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    if (!name) return fail(res, '请填写姓名');
    let count = Number(req.body.count) || config.claimBatchSize;
    count = Math.max(1, Math.min(count, 200));

    const pool = db.getPool();
    // 只从进行中的批次领取；用 UPDATE...LIMIT 保证并发安全
    const [r] = await pool.query(
      `UPDATE tasks t JOIN batches b ON t.batch_id=b.id
       SET t.status='claimed', t.assignee=?, t.claimed_at=NOW()
       WHERE t.status='pending' AND b.status='active'
       ORDER BY t.batch_id, t.id LIMIT ${count}`,
      [name]
    );
    ok(res, { claimed: r.affectedRows });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 标注员：查看可选任务（待领取池） ==================== */
app.get('/api/available-tasks', async (req, res) => {
  try {
    const pool = db.getPool();
    const limit = Math.min(Number(req.query.limit) || 200, 500);
    const batchId = req.query.batchId ? Number(req.query.batchId) : null;
    let sql = `SELECT t.id, t.batch_id, t.full_image_url, t.prospecting_id, t.cell_name, t.row_count AS img_count,
                      b.name AS batch_name
               FROM tasks t JOIN batches b ON t.batch_id=b.id
               WHERE t.status='pending' AND b.status='active'`;
    const params = [];
    if (batchId) { sql += ' AND t.batch_id=?'; params.push(batchId); }
    sql += ' ORDER BY t.batch_id, t.id LIMIT ?';
    params.push(limit);
    const [rows] = await pool.query(sql, params);
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 标注员：单个领取 ==================== */
app.post('/api/claim-single', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    if (!name) return fail(res, '请填写姓名');
    const taskId = Number(req.body.taskId);
    if (!taskId) return fail(res, '请指定任务ID');
    const pool = db.getPool();
    const [r] = await pool.query(
      `UPDATE tasks t JOIN batches b ON t.batch_id=b.id
       SET t.status='claimed', t.assignee=?, t.claimed_at=NOW()
       WHERE t.id=? AND t.status='pending' AND b.status='active'`,
      [name, taskId]
    );
    if (!r.affectedRows) return fail(res, '该任务已被他人领取或批次不可用', 409);
    ok(res, { claimed: 1 });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 标注员：我的任务列表 ==================== */
app.get('/api/my-tasks', async (req, res) => {
  try {
    const name = (req.query.name || '').trim();
    if (!name) return fail(res, '请填写姓名');
    const pool = db.getPool();
    const [rows] = await pool.query(
      `SELECT t.id, t.batch_id, t.full_image_url, t.prospecting_id, t.cell_name, t.status,
              t.draft IS NOT NULL AS has_draft, b.name AS batch_name, b.status AS batch_status
       FROM tasks t JOIN batches b ON t.batch_id=b.id
       WHERE t.assignee=? AND t.status IN ('claimed','submitted')
       ORDER BY t.status='submitted', t.id`,
      [name]
    );
    // 已提交数量
    const [[cnt]] = await pool.query(
      "SELECT SUM(status='claimed') AS claimed, SUM(status='submitted') AS submitted FROM tasks WHERE assignee=?",
      [name]
    );
    ok(res, { tasks: rows, claimed: Number(cnt.claimed || 0), submitted: Number(cnt.submitted || 0) });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 主管：身份检查 ==================== */
app.get('/api/whoami', (req, res) => {
  const name = (req.query.name || '').trim();
  ok(res, { name, isSupervisor: (config.supervisors || []).includes(name) });
});

/* ==================== 主管：查看全员已提交（按标注员分组） ==================== */
app.get('/api/supervisor/all-submitted', async (req, res) => {
  try {
    const name = (req.query.name || '').trim();
    if (!(config.supervisors || []).includes(name)) return fail(res, '无权限：你不在主管名单中', 403);
    const pool = db.getPool();
    const batchId = req.query.batchId ? Number(req.query.batchId) : null;
    const PER_GROUP = 500; // 每个标注员最多返回多少条（按提交时间倒序）
    const conds = ["t.status='submitted'"];
    const params = [];
    if (batchId) { conds.push('t.batch_id=?'); params.push(batchId); }
    const whereSql = conds.join(' AND ');
    // 0) 建索引加速（两种 DB 都兼容，重复执行无害）
    try {
      await pool.query('CREATE INDEX IF NOT EXISTS idx_tasks_sub_assignee ON tasks(status, assignee, submitted_at, id)');
    } catch (e) { /* 忽略索引错误 */ }
    // 1) 各标注员真实总数（不受 PER_GROUP 限制）
    const countSql = `SELECT assignee, COUNT(*) AS total
                      FROM tasks t WHERE ${whereSql}
                      GROUP BY assignee`;
    const [countsRaw] = await pool.query(countSql, params);
    const totalMap = {};
    let totalInDb = 0;
    for (const r of countsRaw) {
      totalMap[r.assignee || '(未知)'] = Number(r.total);
      totalInDb += Number(r.total);
    }
    // 2) 每个标注员取前 PER_GROUP 条
    //    主方案：窗口函数 ROW_NUMBER()（SQLite / MySQL 8.0 均支持，极快）
    //    不支持时（如 MySQL 5.7）回退到相关子查询写法
    const winSql = `SELECT x.id, x.batch_id, x.prospecting_id, x.cell_name, x.full_image_url,
                           x.assignee, x.submitted_at, b.name AS batch_name
                    FROM (
                      SELECT t.id, t.batch_id, t.prospecting_id, t.cell_name, t.full_image_url,
                             t.assignee, t.submitted_at,
                             ROW_NUMBER() OVER (PARTITION BY t.assignee ORDER BY t.submitted_at DESC, t.id DESC) AS rn
                      FROM tasks t
                      WHERE ${whereSql}
                    ) x
                    JOIN batches b ON x.batch_id = b.id
                    WHERE x.rn <= ?
                    ORDER BY x.assignee, x.submitted_at DESC, x.id DESC`;
    let rows;
    try {
      [rows] = await pool.query(winSql, [...params, PER_GROUP]);
    } catch (e) {
      // 回退：相关子查询（兼容老 MySQL）
      const fbSql = `SELECT t.id, t.batch_id, t.prospecting_id, t.cell_name, t.full_image_url,
                            t.assignee, t.submitted_at, b.name AS batch_name
                     FROM tasks t JOIN batches b ON t.batch_id=b.id
                     WHERE ${whereSql}
                       AND (
                         SELECT COUNT(*) FROM tasks t2
                         WHERE t2.status='submitted'
                           ${batchId ? 'AND t2.batch_id=t.batch_id' : ''}
                           AND t2.assignee = t.assignee
                           AND (t2.submitted_at > t.submitted_at
                                OR (t2.submitted_at = t.submitted_at AND t2.id > t.id))
                       ) < ?
                     ORDER BY t.assignee, t.submitted_at DESC, t.id DESC`;
      [rows] = await pool.query(fbSql, [...params, PER_GROUP]);
    }
    // 按标注员分组
    const groups = {};
    const loadedCount = {};
    for (const r of rows) {
      const key = r.assignee || '(未知)';
      if (!groups[key]) groups[key] = [];
      loadedCount[key] = (loadedCount[key] || 0) + 1;
      groups[key].push({
        id: r.id, batch_id: r.batch_id, batch_name: r.batch_name,
        prospecting_id: r.prospecting_id, cell_name: r.cell_name,
        full_image_url: r.full_image_url,
        submitted_at: r.submitted_at ? String(r.submitted_at instanceof Date ? localDate(r.submitted_at) + ' ' + r.submitted_at.toTimeString().slice(0, 8) : r.submitted_at).replace('T', ' ').slice(0, 19) : ''
      });
    }
    // 附加每个组的真实总数（> PER_GROUP 时前端显示「加载了 X/Y」）
    const groupInfo = {};
    for (const k of Object.keys(totalMap)) {
      groupInfo[k] = { total: totalMap[k], loaded: loadedCount[k] || 0 };
    }
    ok(res, { groups, groupInfo, totalInDb, perGroup: PER_GROUP });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 标注员：任务详情 ==================== */
app.get('/api/task/:id', async (req, res) => {
  try {
    const pool = db.getPool();
    const [[t]] = await pool.query(
      `SELECT t.*, b.name AS batch_name, b.status AS batch_status, b.label_columns
       FROM tasks t JOIN batches b ON t.batch_id=b.id WHERE t.id=?`,
      [req.params.id]
    );
    if (!t) return fail(res, '任务不存在', 404);
    t.label_columns = typeof t.label_columns === 'string' ? JSON.parse(t.label_columns) : t.label_columns;
    t.draft = typeof t.draft === 'string' ? JSON.parse(t.draft || 'null') : t.draft;
    t.answers = typeof t.answers === 'string' ? JSON.parse(t.answers || 'null') : t.answers;
    ok(res, t);
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 标注员：保存草稿（实时） ==================== */
app.post('/api/task/:id/draft', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    const draft = req.body.draft || {};
    const pool = db.getPool();
    const [r] = await pool.query(
      "UPDATE tasks SET draft=? WHERE id=? AND assignee=? AND status='claimed'",
      [JSON.stringify(draft), req.params.id, name]
    );
    if (!r.affectedRows) return fail(res, '保存失败：任务不属于你或已提交', 409);
    ok(res, { saved: true });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 标注员：提交整组 ==================== */
app.post('/api/task/:id/submit', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    const answers = req.body.answers || {};
    const remark = (req.body.remark || '').trim();
    const pool = db.getPool();

    const [[t]] = await pool.query(
      `SELECT t.id, t.assignee, t.status, b.status AS batch_status, b.label_columns
       FROM tasks t JOIN batches b ON t.batch_id=b.id WHERE t.id=?`,
      [req.params.id]
    );
    if (!t) return fail(res, '任务不存在', 404);
    if (t.assignee !== name) return fail(res, '任务不属于你', 409);
    // 允许两种情况：首次提交（claimed）、修改后重新提交（submitted 且是本人）
    if (t.status !== 'claimed' && t.status !== 'submitted') return fail(res, '任务状态已变更，无法提交', 409);
    if (t.batch_status === 'stopped') return fail(res, '该批次已停止，无法提交', 403);

    const labelColumns = typeof t.label_columns === 'string' ? JSON.parse(t.label_columns) : t.label_columns;
    const VALID = new Set(['是', '否', '无法判断']);
    const missing = labelColumns.filter(c => !VALID.has(answers[c]));
    if (missing.length) return fail(res, '以下标注项未完成：' + missing.join('、'));

    const clean = {};
    for (const c of labelColumns) clean[c] = answers[c];
    const isResubmit = t.status === 'submitted';
    await pool.query(
      "UPDATE tasks SET status='submitted', answers=?, remark=?, draft=NULL, submitted_at=NOW() WHERE id=?",
      [JSON.stringify(clean), remark, t.id]
    );
    ok(res, { submitted: true, resubmit: isResubmit });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 主管：覆盖他人已提交标注 ==================== */
app.post('/api/task/:id/supervisor-resubmit', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    if (!(config.supervisors || []).includes(name)) return fail(res, '无权限：仅主管可执行此操作', 403);
    const answers = req.body.answers || {};
    const remark = (req.body.remark || '').trim();
    const pool = db.getPool();

    const [[t]] = await pool.query(
      `SELECT t.id, t.status, b.status AS batch_status, b.label_columns
       FROM tasks t JOIN batches b ON t.batch_id=b.id WHERE t.id=?`,
      [req.params.id]
    );
    if (!t) return fail(res, '任务不存在', 404);
    if (t.status !== 'submitted') return fail(res, '只能覆盖已提交的标注', 409);
    if (t.batch_status === 'stopped') return fail(res, '该批次已停止', 403);

    const labelColumns = typeof t.label_columns === 'string' ? JSON.parse(t.label_columns) : t.label_columns;
    const VALID = new Set(['是', '否', '无法判断']);
    const missing = labelColumns.filter(c => !VALID.has(answers[c]));
    if (missing.length) return fail(res, '以下标注项未完成：' + missing.join('、'));

    const clean = {};
    for (const c of labelColumns) clean[c] = answers[c];
    await pool.query(
      "UPDATE tasks SET answers=?, remark=?, submitted_at=NOW() WHERE id=?",
      [JSON.stringify(clean), remark, t.id]
    );
    ok(res, { submitted: true });
  } catch (e) { fail(res, e.message, 500); }
});

/* ==================== 图片代理（绕过图床限流/防盗链） ==================== */
const IMG_CACHE = path.join(__dirname, 'cache', 'img');
fs.mkdirSync(IMG_CACHE, { recursive: true });
const IMG_EXTS = ['jpg','jpeg','png','gif','webp','bmp','svg'];

// 服务端抓取远程图片：用 http/https 模块（可忽略证书不匹配 + 跟随重定向），带超时
function fetchRemote(urlStr, timeoutMs) {
  return new Promise((resolve, reject) => {
    let url;
    try { url = new URL(urlStr); } catch { return reject(new Error('invalid url')); }
    const lib = url.protocol === 'https:' ? https : http;
    const req = lib.get(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36',
        'Referer': url.origin + '/',
        'Accept': 'image/avif,image/webp,image/png,image/*,*/*;q=0.8'
      },
      timeout: timeoutMs,
      // 部分图床证书与域名不匹配（如 img.ljcdn.com 实际证书为 *.bdydns.com），忽略校验避免抓取失败
      rejectUnauthorized: false
    }, (res) => {
      // 跟随 3xx 重定向（CDN 常见）
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        return resolve(fetchRemote(new URL(res.headers.location, url).toString(), timeoutMs));
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error('HTTP ' + res.statusCode));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
    });
    req.on('error', reject);
    req.on('timeout', () => req.destroy(new Error('timeout')));
  });
}

// 带重试与超时，模拟浏览器头避免防盗链
async function fetchImageWithRetry(url, retries) {
  let lastErr;
  for (let i = 0; i < retries; i++) {
    try {
      return await fetchRemote(url, 15000);
    } catch (e) {
      lastErr = e;
      await new Promise(r => setTimeout(r, 600 * (i + 1))); // 退避后重试
    }
  }
  throw lastErr;
}

app.get('/api/img', async (req, res) => {
  const u = req.query.u;
  if (!u || !/^https?:\/\//i.test(u)) return fail(res, 'invalid url');
  let ext = '';
  try { ext = path.extname(new URL(u).pathname).toLowerCase().replace(/^\./, ''); } catch {}
  if (!IMG_EXTS.includes(ext)) ext = 'jpg';
  const key = crypto.createHash('md5').update(u).digest('hex');
  const file = path.join(IMG_CACHE, key + '.' + ext);
  // 命中本地缓存：直接返回，不再请求远程图床
  if (fs.existsSync(file)) {
    try {
      const stat = await fsp.stat(file);
      res.set('Cache-Control', 'public, max-age=86400');
      res.set('Content-Type', ext === 'svg' ? 'image/svg+xml' : 'image/' + (ext === 'jpg' ? 'jpeg' : ext));
      res.set('X-Cache', 'HIT');
      return res.sendFile(file);
    } catch {}
  }
  // 未命中：服务端抓取并缓存（带重试，避免限流导致整页空白）
  try {
    const buf = await fetchImageWithRetry(u, 3);
    await fsp.writeFile(file, buf).catch(() => {});
    res.set('Content-Type', ext === 'svg' ? 'image/svg+xml' : 'image/' + (ext === 'jpg' ? 'jpeg' : ext));
    res.set('Cache-Control', 'public, max-age=86400');
    res.set('X-Cache', 'MISS');
    res.send(buf);
  } catch (e) {
    fail(res, '图片抓取失败: ' + (e && e.message ? e.message : e), 502);
  }
});

/* ==================== 启动 ==================== */
db.init().then(() => {
  const mode = db.getMode() === 'sqlite' ? '本地预览(SQLite)' : 'MySQL';
  app.listen(config.server.port, config.server.host, () => {
    console.log(`标注平台已启动 [${mode}]: http://localhost:${config.server.port}`);
    console.log(`标注员/管理员入口: http://<内网IP>:${config.server.port}/（顶部切换标签）`);
  });
}).catch(e => {
  console.error('启动失败（请检查 config.js 中 MySQL 配置）:', e.message);
  process.exit(1);
});

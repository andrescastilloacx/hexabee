// HEXABEE · Función serverless: valida PIN por socio y actualiza data/tasks.json
// en el repo de GitHub vía API (commit = trazabilidad completa en el historial git).
// Variables de entorno requeridas en Netlify:
//   GITHUB_TOKEN  → token fine-grained con permiso Contents:write SOLO en el repo hexabee
//   PIN_ANDRES, PIN_OLGA, PIN_YUFENG, PIN_ALVARO → PIN personal de cada socio
// Opcionales: GITHUB_REPO (default andrescastilloacx/hexabee), GITHUB_BRANCH (default main)

const crypto = require('crypto');

const ALLOWED_ORIGINS = [
  'https://andrescastilloacx.github.io',
  'http://localhost:8888',
  'http://localhost:3000'
];
const VALID_STATUS = ['pendiente', 'en_progreso', 'bloqueada', 'completada'];
const PINS = { andres: 'PIN_ANDRES', olga: 'PIN_OLGA', yufeng: 'PIN_YUFENG', alvaro: 'PIN_ALVARO' };

function corsHeaders(origin) {
  const o = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    'Access-Control-Allow-Origin': o,
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };
}

function safeEqual(a, b) {
  const ba = Buffer.from(String(a)), bb = Buffer.from(String(b));
  if (ba.length !== bb.length) return false;
  return crypto.timingSafeEqual(ba, bb);
}

function resp(code, obj, headers) {
  return { statusCode: code, headers: { ...headers, 'Content-Type': 'application/json' }, body: JSON.stringify(obj) };
}

exports.handler = async (event) => {
  const headers = corsHeaders(event.headers.origin || event.headers.Origin || '');
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return resp(405, { error: 'Método no permitido' }, headers);

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return resp(400, { error: 'JSON inválido' }, headers); }

  const { member, pin, taskId, status, blocker, note } = body;
  if (!PINS[member]) return resp(400, { error: 'Socio desconocido' }, headers);
  if (!VALID_STATUS.includes(status)) return resp(400, { error: 'Estado inválido' }, headers);
  if (typeof taskId !== 'string' || !taskId) return resp(400, { error: 'taskId requerido' }, headers);

  const expected = process.env[PINS[member]];
  if (!expected) return resp(500, { error: 'PIN no configurado para ' + member }, headers);
  if (!pin || !safeEqual(pin, expected)) return resp(401, { error: 'PIN incorrecto' }, headers);

  const repo = process.env.GITHUB_REPO || 'andrescastilloacx/hexabee';
  const branch = process.env.GITHUB_BRANCH || 'main';
  const token = process.env.GITHUB_TOKEN;
  if (!token) return resp(500, { error: 'GITHUB_TOKEN no configurado' }, headers);

  const api = `https://api.github.com/repos/${repo}/contents/data/tasks.json`;
  const gh = {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'User-Agent': 'hexabee-tracker'
  };

  // Reintento simple ante conflicto de concurrencia (dos socios reportando a la vez)
  for (let attempt = 0; attempt < 2; attempt++) {
    const cur = await fetch(`${api}?ref=${branch}`, { headers: gh });
    if (!cur.ok) return resp(502, { error: 'No se pudo leer tasks.json (' + cur.status + ')' }, headers);
    const file = await cur.json();
    const data = JSON.parse(Buffer.from(file.content, 'base64').toString('utf8'));

    const task = (data.tasks || []).find(t => t.id === taskId);
    if (!task) return resp(404, { error: 'Tarea no encontrada: ' + taskId }, headers);

    const now = new Date().toISOString();
    const entry = {
      ts: now, member, taskId,
      from: task.status, to: status,
      blocker: (blocker || '').slice(0, 300) || null,
      note: (note || '').slice(0, 500) || null
    };
    task.status = status;
    task.blocker = entry.blocker;
    task.note = (note || '').slice(0, 500);
    task.updated_by = member;
    task.updated_at = now;
    data.updated_at = now;
    data.log = [entry].concat(data.log || []).slice(0, 300);

    const put = await fetch(api, {
      method: 'PUT',
      headers: { ...gh, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: `report(${member}): ${taskId} → ${status}`,
        content: Buffer.from(JSON.stringify(data, null, 2)).toString('base64'),
        sha: file.sha,
        branch
      })
    });
    if (put.ok) return resp(200, { ok: true, task }, headers);
    if (put.status === 409 || put.status === 422) continue; // sha viejo → reintentar
    return resp(502, { error: 'No se pudo guardar (' + put.status + ')' }, headers);
  }
  return resp(409, { error: 'Conflicto de concurrencia, reintenta en unos segundos' }, headers);
};

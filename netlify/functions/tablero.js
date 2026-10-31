// HEXABEE · Función serverless del tablero de socios.
// Valida el PIN del socio y recién ahí lee o escribe en la planilla "Hexabee · Tablero de socios"
// a través de su Apps Script (ver tablero/apps-script.gs).
// Variables de entorno en Netlify:
//   PIN_ANDRES, PIN_OLGA, PIN_YUFENG, PIN_ALVARO  (ya existen, las usa update-task.js)
//   TABLERO_URL    → URL /exec de la aplicación web de Apps Script
//   TABLERO_TOKEN  → token generado por setup() en Apps Script

const crypto = require('crypto');

const ALLOWED_ORIGINS = [
  'https://andrescastilloacx.github.io',
  'https://lucky-gingersnap-f5db20.netlify.app',
  'http://localhost:8888',
  'http://localhost:3000'
];
const PINS = { andres: 'PIN_ANDRES', olga: 'PIN_OLGA', yufeng: 'PIN_YUFENG', alvaro: 'PIN_ALVARO' };
const ACTIONS = ['read', 'log', 'update', 'nueva'];

function corsHeaders(origin) {
  const o = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    'Access-Control-Allow-Origin': o,
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Vary': 'Origin'
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
  if (event.httpMethod !== 'POST') return resp(405, { ok: false, error: 'method' }, headers);

  const url = process.env.TABLERO_URL, token = process.env.TABLERO_TOKEN;
  if (!url || !token) return resp(503, { ok: false, error: 'no_conectado' }, headers);

  let body;
  try { body = JSON.parse(event.body || '{}'); } catch (e) { return resp(400, { ok: false, error: 'json' }, headers); }

  const member = String(body.member || '');
  const envName = PINS[member];
  const expected = envName && process.env[envName];
  if (!expected || !safeEqual(body.pin || '', expected)) return resp(401, { ok: false, error: 'pin' }, headers);

  const action = String(body.action || 'read');
  if (!ACTIONS.includes(action)) return resp(400, { ok: false, error: 'action' }, headers);

  // El socio siempre es el dueño del PIN: nadie puede registrar a nombre de otro.
  const payload = { ...body, action, socio: member, token };
  delete payload.pin; delete payload.member;

  try {
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      redirect: 'follow'
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (e) { return resp(502, { ok: false, error: 'upstream' }, headers); }
    if (data.error === 'unauthorized') return resp(502, { ok: false, error: 'token' }, headers);
    return resp(data.ok ? 200 : 400, { ...data, member }, headers);
  } catch (e) {
    return resp(502, { ok: false, error: 'upstream' }, headers);
  }
};

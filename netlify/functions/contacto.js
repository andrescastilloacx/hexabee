// HEXABEE · Función pública del formulario de contacto de la web.
// Guarda cada mensaje en la pestaña "Contactos" del Tablero de socios (vía Apps Script),
// para que los 4 socios lo vean en el portal. El correo por Formspree sigue funcionando aparte.
// Variables de entorno: TABLERO_URL, TABLERO_TOKEN (las mismas del tablero).

const ALLOWED_ORIGINS = [
  'https://andrescastilloacx.github.io',
  'https://lucky-gingersnap-f5db20.netlify.app',
  'http://localhost:8888',
  'http://localhost:3000'
];

function corsHeaders(origin) {
  const o = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    'Access-Control-Allow-Origin': o,
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Vary': 'Origin'
  };
}

function resp(code, obj, headers) {
  return { statusCode: code, headers: { ...headers, 'Content-Type': 'application/json' }, body: JSON.stringify(obj) };
}

const clip = (v, n) => String(v == null ? '' : v).trim().slice(0, n);

exports.handler = async (event) => {
  const origin = event.headers.origin || event.headers.Origin || '';
  const headers = corsHeaders(origin);
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return resp(405, { ok: false, error: 'method' }, headers);
  if (origin && !ALLOWED_ORIGINS.includes(origin)) return resp(403, { ok: false, error: 'origin' }, headers);

  const url = process.env.TABLERO_URL, token = process.env.TABLERO_TOKEN;
  if (!url || !token) return resp(503, { ok: false, error: 'no_conectado' }, headers);

  let b;
  try { b = JSON.parse(event.body || '{}'); } catch (e) { return resp(400, { ok: false, error: 'json' }, headers); }

  // Trampa anti-bots: campo oculto que una persona nunca llena.
  if (b._gotcha) return resp(200, { ok: true }, headers);

  const data = {
    nombre: clip(b.nombre, 80),
    email: clip(b.email, 120),
    telefono: clip(b.telefono, 30),
    interes: clip(b.interes, 80),
    mensaje: clip(b.mensaje, 800),
    origen: clip(origin.replace(/^https?:\/\//, ''), 60) || 'web'
  };
  if (!data.nombre || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(data.email)) return resp(400, { ok: false, error: 'datos' }, headers);

  try {
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...data, action: 'contacto', token }),
      redirect: 'follow'
    });
    const j = JSON.parse(await r.text());
    return resp(j.ok ? 200 : 400, { ok: !!j.ok }, headers);
  } catch (e) {
    return resp(502, { ok: false, error: 'upstream' }, headers);
  }
};

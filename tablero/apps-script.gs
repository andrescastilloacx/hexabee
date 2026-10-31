/**
 * HEXABEE · Tablero de socios — Apps Script (vinculado a la planilla "Hexabee · Tablero de socios").
 *
 * FUENTE DE VERDAD DE LAS PROPUESTAS: el "Hexabee · Evaluador de propuestas (OFICIAL)" (Google Sheets).
 *  - Cada pestaña del Evaluador que tenga "Nombre de la propuesta" (A5) y "VEREDICTO" (A96) es una propuesta.
 *  - Una propuesta aparece en la página SOLO si sus 8 checks (C87:C94) están en OK.
 *  - Si una propuesta baja de 8/8, deja de mostrarse (su avance queda guardado por si vuelve a 8/8).
 *  - Si se BORRA la pestaña en el Evaluador, la propuesta se elimina del tablero y de la página.
 *  - La sincronización corre sola cada vez que alguien abre el portal (y cada 15 min por el disparador).
 *
 * Instalar / actualizar (lo hace Andrés):
 *  1. Pegar este archivo completo en Extensiones → Apps Script de la planilla del Tablero → Guardar.
 *  2. Ejecutar `setup` (▶) y autorizar. En "Registro de ejecución" aparece el TOKEN en su propia línea.
 *     (Si el token ya está en Netlify y no quieres cambiarlo, ejecuta `instalarDisparador` en vez de `setup`.)
 *  3. Implementar → Administrar implementaciones → editar → Versión: nueva → Implementar
 *     (la URL /exec no cambia).
 *  4. Netlify → Environment variables: TABLERO_URL = URL /exec, TABLERO_TOKEN = token.
 *
 * Seguridad: la URL /exec es pública, pero sin el TOKEN no entrega ni escribe nada.
 * La página valida el PIN de cada socio en Netlify (PIN_ANDRES, PIN_OLGA, PIN_YUFENG, PIN_ALVARO).
 */

const EVALUADOR_ID = '10ZzuwHGOBVZwRfsV3EN0gbLLFdg38KKTAXMCmRmmgRo';
const SOCIOS = ['andres', 'olga', 'yufeng', 'alvaro'];
const ETAPAS = ['Lista para vender', 'Vendiendo', 'Pausada'];
const HEAD = ['ID', 'Propuesta', 'Responsable', 'Checks OK (de 8)', 'Veredicto Evaluador', 'Hoja', 'GID',
  'Visible', 'Etapa', 'Inicio', 'Fin estimado', 'Avance %', 'Próximo paso', 'Actualizado'];

function setup() {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('HEXABEE_TOKEN', token);
  instalarDisparador();
  sincronizar();
  Logger.log('Copia SOLO la línea siguiente en Netlify como TABLERO_TOKEN:');
  Logger.log(token);
}

function instalarDisparador() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sincronizar') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sincronizar').timeBased().everyMinutes(15).create();
  sincronizar();
  Logger.log('Disparador instalado: sincroniza cada 15 minutos.');
}

/** Lee el Evaluador y deja la pestaña Propuestas exactamente alineada con él. */
function sincronizar() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ev = SpreadsheetApp.openById(EVALUADOR_ID);
    const found = {};
    ev.getSheets().forEach(sh => {
      if (sh.getMaxRows() < 96) return;
      const lab = sh.getRange('A5').getDisplayValue().trim();
      const ver = sh.getRange('A96').getDisplayValue().trim().toUpperCase();
      if (lab !== 'Nombre de la propuesta' || ver !== 'VEREDICTO') return;
      if (/^plantilla$/i.test(sh.getName())) return;
      const nombre = sh.getRange('B5').getDisplayValue().trim();
      if (!nombre) return;
      const checks = sh.getRange('C87:C94').getDisplayValues().filter(r => String(r[0]).trim().toUpperCase() === 'OK').length;
      found[String(sh.getSheetId())] = {
        nombre: nombre,
        responsable: socio_(sh.getRange('B6').getDisplayValue()),
        checks: checks,
        veredicto: sh.getRange('B96').getDisplayValue().trim(),
        hoja: sh.getName()
      };
    });

    const sh = SpreadsheetApp.getActive().getSheetByName('Propuestas');
    ensureHead_(sh);
    const data = sh.getDataRange().getValues();
    const head = data[0].map(String), c = n => head.indexOf(n);
    const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
    const seen = {};
    let maxN = 0;
    data.slice(1).forEach(r => { const m = /^P-(\d+)$/.exec(String(r[c('ID')])); if (m) maxN = Math.max(maxN, +m[1]); });

    const asText = v => (v instanceof Date ? Utilities.formatDate(v, 'America/Santiago', 'yyyy-MM-dd') : String(v));
    const out = [];
    data.slice(1).forEach(row => {
      const gid = String(row[c('GID')]);
      const f = found[gid];
      if (!gid || !f || seen[gid]) return;              // borrada en el Evaluador (o duplicada / manual) → fuera
      seen[gid] = true;
      const ok = f.checks === 8;
      const r = HEAD.map(h => asText(row[c(h)] === undefined ? '' : row[c(h)]));
      r[HEAD.indexOf('Propuesta')] = f.nombre;
      r[HEAD.indexOf('Responsable')] = f.responsable;
      r[HEAD.indexOf('Checks OK (de 8)')] = f.checks;
      r[HEAD.indexOf('Veredicto Evaluador')] = f.veredicto;
      r[HEAD.indexOf('Hoja')] = f.hoja;
      r[HEAD.indexOf('Visible')] = ok ? 'SI' : 'NO';
      if (ok && !r[HEAD.indexOf('Inicio')]) r[HEAD.indexOf('Inicio')] = hoy;
      if (ok && !r[HEAD.indexOf('Etapa')]) r[HEAD.indexOf('Etapa')] = 'Lista para vender';
      out.push(r);
    });
    Object.keys(found).forEach(gid => {
      if (seen[gid]) return;
      const f = found[gid], ok = f.checks === 8;
      maxN++;
      out.push(['P-' + String(maxN).padStart(2, '0'), f.nombre, f.responsable, f.checks, f.veredicto, f.hoja, gid,
        ok ? 'SI' : 'NO', ok ? 'Lista para vender' : '', ok ? hoy : '', '', 0, ok ? 'Definir fecha de término y primer cliente' : '', hoy]);
    });
    // Reescribe el bloque completo (como texto plano en fechas) para que quede idéntico al Evaluador.
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), HEAD.length)).clearContent();
    if (out.length) {
      const rng = sh.getRange(2, 1, out.length, HEAD.length);
      rng.setNumberFormat('@');
      rng.setValues(out);
    }
    PropertiesService.getScriptProperties().setProperty('ULTIMA_SYNC', new Date().toISOString());
  } finally { lock.releaseLock(); }
}

function doGet() {
  return json_({ ok: true, service: 'hexabee-tablero' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'json' }); }

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('HEXABEE_TOKEN');
  if (!token || body.token !== token) return json_({ ok: false, error: 'unauthorized' });

  const ss = SpreadsheetApp.getActive();

  if (body.action === 'read') {
    let syncError = '';
    try { sincronizar(); } catch (err) { syncError = String(err).slice(0, 160); }
    const todas = rows_(ss.getSheetByName('Propuestas'));
    return json_({
      ok: true,
      propuestas: todas.filter(p => p['Visible'] === 'SI'),
      enEvaluacion: todas.filter(p => p['Visible'] !== 'SI').map(p => ({ 'ID': p['ID'], 'Propuesta': p['Propuesta'], 'Responsable': p['Responsable'], 'Checks OK (de 8)': p['Checks OK (de 8)'] })),
      registro: rows_(ss.getSheetByName('Registro')).slice(-300),
      resumen: rows_(ss.getSheetByName('Resumen')).filter(r => r['Socio']),
      contactos: ss.getSheetByName('Contactos') ? rows_(ss.getSheetByName('Contactos')).slice(-150) : [],
      ultimaSync: props.getProperty('ULTIMA_SYNC') || '',
      syncError: syncError
    });
  }

  const lock = LockService.getScriptLock();

  if (body.action === 'contacto') {
    // Mensaje del formulario público de la web (llega vía Netlify, que valida y limpia los campos).
    const clip = (v, n) => String(v || '').replace(/[\r\n]+/g, ' ').trim().slice(0, n);
    const nombre = clip(body.nombre, 80), email = clip(body.email, 120);
    if (!nombre || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, error: 'datos' });
    lock.waitLock(10000);
    try {
      let sh = ss.getSheetByName('Contactos');
      if (!sh) sh = ss.insertSheet('Contactos');
      if (sh.getLastRow() === 0) {
        sh.appendRow(['Fecha', 'Nombre', 'Correo', 'Teléfono', 'Interés', 'Mensaje', 'Origen', 'Estado', 'Responsable']);
        sh.setFrozenRows(1);
      }
      sh.appendRow(["'" + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm'), nombre, email,
        "'" + clip(body.telefono, 30), clip(body.interes, 80), String(body.mensaje || '').trim().slice(0, 800),
        clip(body.origen, 60), 'Nuevo', '']);
    } finally { lock.releaseLock(); }
    return json_({ ok: true });
  }

  if (body.action === 'log') {
    const socio = String(body.socio || '');
    const horas = Number(body.horas);
    const fecha = String(body.fecha || '');
    const texto = String(body.texto || '').slice(0, 300);
    const prop = String(body.propuesta || '').slice(0, 20);
    if (SOCIOS.indexOf(socio) < 0) return json_({ ok: false, error: 'socio' });
    if (!(horas > 0 && horas <= 16)) return json_({ ok: false, error: 'horas' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return json_({ ok: false, error: 'fecha' });
    if (!texto.trim()) return json_({ ok: false, error: 'texto' });
    lock.waitLock(10000);
    try {
      ss.getSheetByName('Registro').appendRow(["'" + fecha, socio, horas, prop, texto, new Date()]);
    } finally { lock.releaseLock(); }
    return json_({ ok: true });
  }

  if (body.action === 'update') {
    // Solo campos de seguimiento. Nombre, responsable y checks vienen del Evaluador y no se editan aquí.
    const socio = String(body.socio || '');
    const id = String(body.id || '');
    const sh = ss.getSheetByName('Propuestas');
    lock.waitLock(10000);
    try {
      const data = sh.getDataRange().getValues();
      const head = data[0].map(String), col = n => head.indexOf(n);
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][col('ID')]) !== id) continue;
        if (String(data[i][col('Visible')]) !== 'SI') return json_({ ok: false, error: 'no_visible' });
        const owner = String(data[i][col('Responsable')]);
        if (socio !== owner && socio !== 'andres') return json_({ ok: false, error: 'forbidden' });
        const r = i + 1;
        if (body.etapa && ETAPAS.indexOf(body.etapa) >= 0) sh.getRange(r, col('Etapa') + 1).setValue(body.etapa);
        if (body.avance !== undefined && body.avance !== '') {
          const a = Math.max(0, Math.min(100, Number(body.avance)));
          if (!isNaN(a)) sh.getRange(r, col('Avance %') + 1).setValue(a);
        }
        if (body.fin !== undefined && /^(\d{4}-\d{2}-\d{2})?$/.test(String(body.fin))) {
          sh.getRange(r, col('Fin estimado') + 1).setValue(body.fin ? "'" + body.fin : '');
        }
        if (body.proximo) sh.getRange(r, col('Próximo paso') + 1).setValue(String(body.proximo).slice(0, 200));
        sh.getRange(r, col('Actualizado') + 1).setValue("'" + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd'));
        return json_({ ok: true });
      }
    } finally { lock.releaseLock(); }
    return json_({ ok: false, error: 'not_found' });
  }

  return json_({ ok: false, error: 'action' });
}

function ensureHead_(sh) {
  const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  if (cur.join('|') === HEAD.join('|')) return;
  sh.clear();
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function socio_(nombre) {
  const s = String(nombre || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  for (let i = 0; i < SOCIOS.length; i++) if (s.indexOf(SOCIOS[i]) >= 0) return SOCIOS[i];
  if (s.indexOf('yu') >= 0) return 'yufeng';
  return s.split(/\s+/)[0] || '';
}

function rows_(sh) {
  // Fechas siempre como AAAA-MM-DD (la planilla puede convertirlas a fecha local).
  const v = sh.getDataRange().getValues();
  const head = v.shift().map(String);
  return v.filter(r => r.join('') !== '').map(r => {
    const o = {};
    head.forEach((h, i) => {
      if (!h) return;
      const c = r[i];
      o[h] = c instanceof Date ? Utilities.formatDate(c, 'America/Santiago', 'yyyy-MM-dd') : String(c);
    });
    return o;
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

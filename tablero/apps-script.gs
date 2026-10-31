/**
 * HEXABEE · Tablero de socios — Apps Script (vinculado a la planilla "Hexabee · Tablero de socios").
 *
 * Instalar (una sola vez, lo hace Andrés):
 *  1. Abrir la planilla → Extensiones → Apps Script → pegar este archivo completo → Guardar.
 *  2. Ejecutar la función `setup` (botón ▶). Autorizar. En "Registro de ejecución" aparece el TOKEN.
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web.
 *       Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario.
 *     Copiar la URL que termina en /exec.
 *  4. En Netlify → Site settings → Environment variables, agregar:
 *       TABLERO_URL   = la URL /exec
 *       TABLERO_TOKEN = el TOKEN del paso 2
 *
 * Seguridad: la URL /exec es pública, pero sin el TOKEN no entrega ni escribe nada.
 * El TOKEN solo vive aquí y en Netlify; la página nunca lo ve. La página valida el PIN
 * de cada socio en Netlify (variables PIN_ANDRES, PIN_OLGA, PIN_YUFENG, PIN_ALVARO).
 */

const SOCIOS = ['andres', 'olga', 'yufeng', 'alvaro'];
const ETAPAS = ['Idea', 'Evaluando', 'Validando demanda', 'Lista para vender', 'Vendiendo', 'Pausada'];

function setup() {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('HEXABEE_TOKEN', token);
  Logger.log('Copia SOLO la línea siguiente en Netlify como TABLERO_TOKEN:');
  Logger.log(token);
}

function doGet() {
  return json_({ ok: true, service: 'hexabee-tablero' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'json' }); }

  const token = PropertiesService.getScriptProperties().getProperty('HEXABEE_TOKEN');
  if (!token || body.token !== token) return json_({ ok: false, error: 'unauthorized' });

  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();

  if (body.action === 'read') {
    return json_({
      ok: true,
      propuestas: rows_(ss.getSheetByName('Propuestas')),
      registro: rows_(ss.getSheetByName('Registro')).slice(-300),
      resumen: rows_(ss.getSheetByName('Resumen')).filter(r => r['Socio']),
      generado: new Date().toISOString()
    });
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
      ss.getSheetByName('Registro').appendRow([fecha, socio, horas, prop, texto, new Date()]);
    } finally { lock.releaseLock(); }
    return json_({ ok: true });
  }

  if (body.action === 'update') {
    // Actualiza etapa / avance / próximo paso / fin estimado de una propuesta.
    // Solo el responsable de la propuesta, o andres (admin), puede hacerlo.
    const socio = String(body.socio || '');
    const id = String(body.id || '');
    const sh = ss.getSheetByName('Propuestas');
    const data = sh.getDataRange().getValues();
    const head = data[0];
    const col = name => head.indexOf(name);
    lock.waitLock(10000);
    try {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][col('ID')]) !== id) continue;
        const owner = String(data[i][col('Responsable')]);
        if (socio !== owner && socio !== 'andres') return json_({ ok: false, error: 'forbidden' });
        const r = i + 1;
        if (body.etapa && ETAPAS.indexOf(body.etapa) >= 0) sh.getRange(r, col('Etapa') + 1).setValue(body.etapa);
        if (body.avance !== undefined && body.avance !== '') {
          const a = Math.max(0, Math.min(100, Number(body.avance)));
          if (!isNaN(a)) sh.getRange(r, col('Avance %') + 1).setValue(a);
        }
        if (body.fin !== undefined && /^(\d{4}-\d{2}-\d{2})?$/.test(String(body.fin))) {
          sh.getRange(r, col('Fin estimado') + 1).setValue(String(body.fin));
        }
        if (body.proximo) sh.getRange(r, col('Próximo paso') + 1).setValue(String(body.proximo).slice(0, 200));
        sh.getRange(r, col('Actualizado') + 1).setValue(Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd'));
        return json_({ ok: true });
      }
    } finally { lock.releaseLock(); }
    return json_({ ok: false, error: 'not_found' });
  }

  if (body.action === 'nueva') {
    // Un socio agrega una propuesta nueva a su nombre.
    const socio = String(body.socio || '');
    const nombre = String(body.nombre || '').trim().slice(0, 140);
    const tipo = ['Venta', 'Producto', 'Proceso'].indexOf(body.tipo) >= 0 ? body.tipo : 'Venta';
    if (SOCIOS.indexOf(socio) < 0 || !nombre) return json_({ ok: false, error: 'datos' });
    lock.waitLock(10000);
    try {
      const sh = ss.getSheetByName('Propuestas');
      const n = sh.getLastRow(); // incluye encabezado
      const id = 'P-' + String(n).padStart(2, '0');
      const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
      sh.appendRow([id, nombre, socio, tipo, 'Idea', '', 'Sin evaluar', hoy, String(body.fin || ''), 0,
        String(body.proximo || 'Evaluarla en el Evaluador de propuestas').slice(0, 200), hoy]);
      return json_({ ok: true, id: id });
    } finally { lock.releaseLock(); }
  }

  return json_({ ok: false, error: 'action' });
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

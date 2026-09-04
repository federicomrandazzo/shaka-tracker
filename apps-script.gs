/**
 * SHAKA · TAREAS — intermediario entre la app y la planilla
 *
 * Qué hace: valida el login con clave, entrega todas las tareas del equipo
 * y guarda los cambios que manda cada teléfono.
 *
 * Cómo se instala (una sola vez):
 *  1. Crear una planilla nueva en Google Sheets.
 *  2. Extensiones → Apps Script. Borrar lo que haya y pegar TODO este archivo.
 *  3. Cambiar SECRETO por cualquier frase larga inventada (línea de abajo).
 *  4. Guardar. Ejecutar la función  preparar  una vez (Ejecutar ▸ preparar)
 *     y aceptar los permisos que pide Google.
 *  5. Implementar ▸ Nueva implementación ▸ Aplicación web
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier usuario
 *     Copiar la URL que termina en /exec y pasársela a Fede.
 *  6. En la hoja "usuarios" de la planilla, cargar la clave de cada uno.
 *
 * Si más adelante se toca este código, hay que volver a Implementar ▸
 * Administrar implementaciones ▸ editar ▸ Versión: nueva, o los cambios no salen.
 */

var SECRETO = 'cambiar-esto-por-una-frase-larga-inventada';

// ---- avisos (notificaciones) ----
// PUSH_URL: la direccion de la app + /api/push
// PUSH_SECRET: la misma frase que cargues en Vercel como PUSH_SECRET
// Si dejas PUSH_SECRET vacio, los avisos quedan apagados y todo lo demas sigue igual.
var PUSH_URL    = 'https://shaka-tracker.vercel.app/api/push';
var PUSH_SECRET = '';

var HOJA_U = 'usuarios';
var HOJA_T = 'tareas';
var HOJA_A = 'accesos';
var HOJA_P = 'avisos';

var COLS = ['id','titulo','desc','quien','dia','vence','hecha','hechaEl',
            'corrida','drop','orden','borrada','actualizada'];

/* ============ instalación ============ */
function preparar(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var u = ss.getSheetByName(HOJA_U) || ss.insertSheet(HOJA_U);
  if (u.getLastRow() < 1 || !u.getRange(1,1).getValue()) {
    u.getRange(1,1,1,5).setValues([['nombre','clave','activo','nolog','oculto']]);
    var gente = ['FEDE','AXEL','MATE','MILI','TOMI','MARTU'];
    var filas = gente.map(function(n){ return [n, '', 'si', '', '']; });
    u.getRange(2,1,filas.length,5).setValues(filas);
    u.getRange(1,1,1,5).setFontWeight('bold');
    u.setFrozenRows(1);
    u.setColumnWidth(1,140); u.setColumnWidth(2,160);
    u.setColumnWidth(3,80); u.setColumnWidth(4,80); u.setColumnWidth(5,80);
  }
  // "nolog" = si  -> no queda registro de tus accesos
  if (String(u.getRange(1,4).getValue()).trim().toLowerCase() !== 'nolog') {
    u.getRange(1,4).setValue('nolog').setFontWeight('bold');
    u.setColumnWidth(4,80);
  }
  // "oculto" = si -> podés entrar y ver todo, pero no figurás para el resto:
  // no aparecés en la pantalla de entrada, ni en el equipo, ni te pueden asignar tareas
  if (String(u.getRange(1,5).getValue()).trim().toLowerCase() !== 'oculto') {
    u.getRange(1,5).setValue('oculto').setFontWeight('bold');
    u.setColumnWidth(5,80);
  }

  var a = ss.getSheetByName(HOJA_A) || ss.insertSheet(HOJA_A);
  if (a.getLastRow() < 1 || !a.getRange(1,1).getValue()) {
    a.getRange(1,1,1,4).setValues([['fecha','persona','evento','dispositivo']]);
    a.getRange(1,1,1,4).setFontWeight('bold');
    a.setFrozenRows(1);
    a.setColumnWidth(1,150); a.setColumnWidth(2,110); a.setColumnWidth(3,220); a.setColumnWidth(4,120);
  }

  var s = ss.getSheetByName(HOJA_P) || ss.insertSheet(HOJA_P);
  if (s.getLastRow() < 1 || !s.getRange(1,1).getValue()) {
    s.getRange(1,1,1,6).setValues([['persona','endpoint','p256dh','auth','aparato','fecha']]);
    s.getRange(1,1,1,6).setFontWeight('bold');
    s.setFrozenRows(1);
    s.setColumnWidth(1,110); s.setColumnWidth(2,260); s.setColumnWidth(5,120); s.setColumnWidth(6,150);
  }

  var t = ss.getSheetByName(HOJA_T) || ss.insertSheet(HOJA_T);
  if (t.getLastRow() < 1 || !t.getRange(1,1).getValue()) {
    t.getRange(1,1,1,COLS.length).setValues([COLS]);
    t.getRange(1,1,1,COLS.length).setFontWeight('bold');
    t.setFrozenRows(1);
  }
  // dia, vence y hechaEl como texto: si no, Sheets las convierte en fechas
  // con formato local y la app las recibe ilegibles
  t.getRange(1, 5, t.getMaxRows(), 1).setNumberFormat('@');
  t.getRange(1, 6, t.getMaxRows(), 1).setNumberFormat('@');
  t.getRange(1, 8, t.getMaxRows(), 1).setNumberFormat('@');
  return 'Listo: hojas "usuarios", "tareas", "accesos" y "avisos" preparadas.';
}

/* ============ utilidades ============ */
function hash_(txt){
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(txt), Utilities.Charset.UTF_8);
  return b.map(function(x){ return ('0'+(x & 0xFF).toString(16)).slice(-2); }).join('');
}
function token_(nombre, clave){ return hash_(nombre + '|' + clave + '|' + SECRETO); }

/** Las fechas viajan como texto "2026-08-30". Si la planilla las guardo
    como fecha de verdad, aca vuelven a texto para que la app las entienda. */
function fecha_(v){
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v == null ? '' : v).trim();
}

function hojaU_(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_U); }
function hojaA_(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_A); }
function hojaP_(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_P); }
function hojaT_(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_T); }

function usuarios_(){
  var h = hojaU_(), n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2,1,n,5).getValues()
    .filter(function(f){ return String(f[0]).trim(); })
    .map(function(f){
      return { nombre: String(f[0]).trim().toUpperCase(),
               clave:  String(f[1]).trim(),
               activo: String(f[2]).trim().toLowerCase() !== 'no',
               nolog:  String(f[3]).trim().toLowerCase() === 'si',
               oculto: String(f[4]).trim().toLowerCase() === 'si' };
    });
}

/** La gente que ve el equipo: los ocultos no figuran. */
function visibles_(){
  return usuarios_().filter(function(u){ return u.activo && !u.oculto; })
                    .map(function(u){ return u.nombre; });
}

/** Devuelve el usuario si el token es válido, o null. */
function porToken_(tok){
  if (!tok) return null;
  var us = usuarios_();
  for (var i=0;i<us.length;i++){
    if (us[i].activo && us[i].clave && token_(us[i].nombre, us[i].clave) === tok) return us[i];
  }
  return null;
}

function ok_(obj){
  obj = obj || {}; obj.ok = true;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function error_(msg){
  return ContentService.createTextOutput(JSON.stringify({ok:false, error:msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============ lectura y escritura de tareas ============ */
function leerTareas_(){
  var h = hojaT_(), n = h.getLastRow() - 1;
  if (n < 1) return [];
  var vals = h.getRange(2,1,n,COLS.length).getValues();
  var out = [];
  for (var i=0;i<vals.length;i++){
    var f = vals[i];
    if (!String(f[0]).trim()) continue;
    out.push({
      id:        String(f[0]),
      title:     String(f[1]),
      desc:      String(f[2]),
      who:       String(f[3]) ? String(f[3]).split('|').filter(String) : [],
      do:        fecha_(f[4]),
      due:       fecha_(f[5]),
      done:      String(f[6]) === 'si',
      doneAt:    fecha_(f[7]) || null,
      pushed:    Number(f[8]) || 0,
      drop:      String(f[9]),
      ord:       Number(f[10]) || 0,
      borrada:   String(f[11]) === 'si',
      actualizada: Number(f[12]) || 0
    });
  }
  return out;
}

function filaDe_(t){
  return [ t.id, t.title||'', t.desc||'',
           (t.who||[]).join('|'), t.do||'', t.due||'',
           t.done?'si':'no', t.doneAt||'',
           Number(t.pushed)||0, t.drop||'',
           Number(t.ord)||0, t.borrada?'si':'no',
           Number(t.actualizada)||0 ];
}

/**
 * Guarda los cambios que manda un teléfono.
 * Gana el cambio más nuevo (comparando "actualizada"), así dos personas
 * editando a la vez no se pisan lo que hizo la otra.
 */
function guardar_(cambios, quien){
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  var avisos = [];   // [{personas:[...], titulo, cuerpo}]
  try {
    var h = hojaT_(), n = h.getLastRow() - 1;
    var vals = n > 0 ? h.getRange(2,1,n,COLS.length).getValues() : [];
    var fila = {};
    for (var i=0;i<vals.length;i++){
      var id = String(vals[i][0]).trim();
      if (id) fila[id] = i + 2;
    }
    var nuevas = [];
    for (var k=0;k<cambios.length;k++){
      var t = cambios[k];
      if (!t || !t.id) continue;
      var r = fila[t.id];
      var antes = r ? String(vals[r-2][3]) : '';
      if (r) {
        var actual = Number(vals[r-2][12]) || 0;
        if ((Number(t.actualizada)||0) >= actual) {
          h.getRange(r,1,1,COLS.length).setValues([filaDe_(t)]);
        } else {
          continue;   // gano el otro cambio, no avisamos por este
        }
      } else {
        nuevas.push(filaDe_(t));
      }
      // a quien le tocaron una tarea que antes no tenia (y que no sea uno mismo)
      if (!t.borrada) {
        var ahora = (t.who || []).map(function(x){ return String(x).trim().toUpperCase(); });
        var previos = antes ? antes.split('|').map(function(x){ return x.trim().toUpperCase(); }) : [];
        var nuevos = ahora.filter(function(x){
          return x && x !== quien && previos.indexOf(x) < 0;
        });
        if (nuevos.length) {
          avisos.push({ personas: nuevos,
                        titulo: quien + ' te asignó una tarea',
                        cuerpo: String(t.title || '').slice(0, 90) });
        }
      }
    }
    if (nuevas.length) {
      h.getRange(h.getLastRow()+1, 1, nuevas.length, COLS.length).setValues(nuevas);
    }
    var salida = leerTareas_();
    lock.releaseLock(); lock = null;
    avisos.forEach(function(a){ avisar_(a.personas, a.titulo, a.cuerpo); });
    return salida;
  } finally {
    if (lock) lock.releaseLock();
  }
}

/* ============ avisos ============ */

/** Guarda (o actualiza) la suscripcion de un telefono. */
function suscribir_(persona, sub){
  var h = hojaP_();
  if (!h || !sub || !sub.endpoint) return false;
  var n = h.getLastRow() - 1;
  var vals = n > 0 ? h.getRange(2,1,n,6).getValues() : [];
  for (var i=0;i<vals.length;i++){
    if (String(vals[i][1]) === sub.endpoint) {          // ya estaba: la actualizamos
      h.getRange(i+2,1,1,6).setValues([[persona, sub.endpoint, sub.p256dh, sub.auth,
                                        sub.aparato||'', new Date()]]);
      return true;
    }
  }
  h.appendRow([persona, sub.endpoint, sub.p256dh, sub.auth, sub.aparato||'', new Date()]);
  return true;
}

/** Las suscripciones de una persona (un mismo usuario puede tener varios telefonos). */
function subsDe_(persona){
  var h = hojaP_();
  if (!h) return [];
  var n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2,1,n,4).getValues()
    .filter(function(f){ return String(f[0]).trim().toUpperCase() === persona && String(f[1]).trim(); })
    .map(function(f){ return { endpoint:String(f[1]), p256dh:String(f[2]), auth:String(f[3]) }; });
}

/** Borra las suscripciones que el navegador ya dio de baja. */
function limpiarSubs_(endpoints){
  if (!endpoints || !endpoints.length) return;
  var h = hojaP_(); if (!h) return;
  var n = h.getLastRow() - 1; if (n < 1) return;
  var vals = h.getRange(2,1,n,1).getValues();
  var endpointsCol = h.getRange(2,2,n,1).getValues();
  for (var i = endpointsCol.length - 1; i >= 0; i--){
    if (endpoints.indexOf(String(endpointsCol[i][0])) >= 0) h.deleteRow(i+2);
  }
}

/** Le pide a Vercel que mande el aviso. Nunca corta el flujo si falla. */
function avisar_(personas, titulo, cuerpo){
  if (!PUSH_SECRET || !PUSH_URL) return;
  try {
    var subs = [], vistos = {};
    personas.forEach(function(p){
      subsDe_(String(p).trim().toUpperCase()).forEach(function(s){
        if (!vistos[s.endpoint]) { vistos[s.endpoint] = 1; subs.push(s); }
      });
    });
    if (!subs.length) return;
    var r = UrlFetchApp.fetch(PUSH_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ secreto: PUSH_SECRET, subs: subs,
                                title: titulo, body: cuerpo,
                                url: 'https://shaka-tracker.vercel.app/' }),
      muteHttpExceptions: true
    });
    var res = JSON.parse(r.getContentText() || '{}');
    if (res && res.vencidas && res.vencidas.length) limpiarSubs_(res.vencidas);
  } catch (err) { /* si el aviso falla, la tarea igual se guardo */ }
}

/* ============ puerta de entrada ============ */
function doGet(e){
  // sirve para probar desde el navegador que la implementación está viva
  return ok_({ vivo:true, usuarios: visibles_() });
}

function doPost(e){
  try {
    var p = JSON.parse(e.postData.contents || '{}');
    var accion = p.accion;

    if (accion === 'gente') {
      return ok_({ gente: visibles_() });
    }

    if (accion === 'login') {
      var nombre = String(p.nombre||'').trim().toUpperCase();
      var clave  = String(p.clave||'').trim();
      var us = usuarios_();
      for (var i=0;i<us.length;i++){
        if (us[i].nombre === nombre && us[i].activo) {
          if (!us[i].clave) return error_('Ese usuario todavía no tiene clave cargada. Avisale a Fede.');
          if (us[i].clave !== clave) return error_('Clave incorrecta.');
          return ok_({ nombre: nombre, token: token_(nombre, clave),
                       nolog: us[i].nolog, oculto: us[i].oculto,
                       gente: visibles_(), tareas: leerTareas_() });
        }
      }
      return error_('No encontré ese usuario.');
    }

    var quien = porToken_(p.token);
    if (!quien) return error_('Sesión vencida, volvé a entrar.');

    if (accion === 'log') {
      if (!quien.nolog) {
        var a = hojaA_();
        if (a) a.appendRow([ new Date(), quien.nombre,
                             String(p.evento||'').slice(0,120),
                             String(p.dispositivo||'').slice(0,60) ]);
      }
      return ok_({});
    }

    if (accion === 'pull') {
      return ok_({ tareas: leerTareas_(), gente: visibles_() });
    }

    if (accion === 'push') {
      var tareas = guardar_(p.cambios || [], quien.nombre);
      return ok_({ tareas: tareas });
    }

    if (accion === 'suscribir') {
      return ok_({ guardada: suscribir_(quien.nombre, p.sub || {}) });
    }

    if (accion === 'probar') {   // para probar los avisos desde la app
      avisar_([quien.nombre], 'Shaka · Tareas', 'Los avisos están andando.');
      return ok_({});
    }

    return error_('No entendí qué hacer: ' + accion);
  } catch (err) {
    return error_('Se rompió algo: ' + err);
  }
}

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

var HOJA_U = 'usuarios';
var HOJA_T = 'tareas';

var COLS = ['id','titulo','desc','quien','dia','vence','hecha','hechaEl',
            'corrida','drop','orden','borrada','actualizada'];

/* ============ instalación ============ */
function preparar(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var u = ss.getSheetByName(HOJA_U) || ss.insertSheet(HOJA_U);
  if (u.getLastRow() < 1 || !u.getRange(1,1).getValue()) {
    u.getRange(1,1,1,3).setValues([['nombre','clave','activo']]);
    var gente = ['FEDE','AXEL','MATE','MILI','TOMI','MARTU'];
    var filas = gente.map(function(n){ return [n, '', 'si']; });
    u.getRange(2,1,filas.length,3).setValues(filas);
    u.getRange(1,1,1,3).setFontWeight('bold');
    u.setFrozenRows(1);
    u.setColumnWidth(1,140); u.setColumnWidth(2,160); u.setColumnWidth(3,80);
  }

  var t = ss.getSheetByName(HOJA_T) || ss.insertSheet(HOJA_T);
  if (t.getLastRow() < 1 || !t.getRange(1,1).getValue()) {
    t.getRange(1,1,1,COLS.length).setValues([COLS]);
    t.getRange(1,1,1,COLS.length).setFontWeight('bold');
    t.setFrozenRows(1);
  }
  return 'Listo: hojas "usuarios" y "tareas" preparadas. Cargá las claves en la hoja usuarios.';
}

/* ============ utilidades ============ */
function hash_(txt){
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(txt), Utilities.Charset.UTF_8);
  return b.map(function(x){ return ('0'+(x & 0xFF).toString(16)).slice(-2); }).join('');
}
function token_(nombre, clave){ return hash_(nombre + '|' + clave + '|' + SECRETO); }

function hojaU_(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_U); }
function hojaT_(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_T); }

function usuarios_(){
  var h = hojaU_(), n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2,1,n,3).getValues()
    .filter(function(f){ return String(f[0]).trim(); })
    .map(function(f){
      return { nombre: String(f[0]).trim().toUpperCase(),
               clave:  String(f[1]).trim(),
               activo: String(f[2]).trim().toLowerCase() !== 'no' };
    });
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
      do:        String(f[4]),
      due:       String(f[5]),
      done:      String(f[6]) === 'si',
      doneAt:    String(f[7]) || null,
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
function guardar_(cambios){
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
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
      if (r) {
        var actual = Number(vals[r-2][12]) || 0;
        if ((Number(t.actualizada)||0) >= actual) {
          h.getRange(r,1,1,COLS.length).setValues([filaDe_(t)]);
        }
      } else {
        nuevas.push(filaDe_(t));
      }
    }
    if (nuevas.length) {
      h.getRange(h.getLastRow()+1, 1, nuevas.length, COLS.length).setValues(nuevas);
    }
    return leerTareas_();
  } finally {
    lock.releaseLock();
  }
}

/* ============ puerta de entrada ============ */
function doGet(e){
  // sirve para probar desde el navegador que la implementación está viva
  return ok_({ vivo:true, usuarios: usuarios_().filter(function(u){return u.activo;})
                                               .map(function(u){return u.nombre;}) });
}

function doPost(e){
  try {
    var p = JSON.parse(e.postData.contents || '{}');
    var accion = p.accion;

    if (accion === 'gente') {
      return ok_({ gente: usuarios_().filter(function(u){return u.activo;})
                                     .map(function(u){return u.nombre;}) });
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
                       gente: us.filter(function(u){return u.activo;}).map(function(u){return u.nombre;}),
                       tareas: leerTareas_() });
        }
      }
      return error_('No encontré ese usuario.');
    }

    var quien = porToken_(p.token);
    if (!quien) return error_('Sesión vencida, volvé a entrar.');

    if (accion === 'pull') {
      return ok_({ tareas: leerTareas_(),
                   gente: usuarios_().filter(function(u){return u.activo;}).map(function(u){return u.nombre;}) });
    }

    if (accion === 'push') {
      var tareas = guardar_(p.cambios || []);
      return ok_({ tareas: tareas });
    }

    return error_('No entendí qué hacer: ' + accion);
  } catch (err) {
    return error_('Se rompió algo: ' + err);
  }
}

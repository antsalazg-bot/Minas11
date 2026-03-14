// ════════════════════════════════════════════════════════════════
// MINAS 11 — SISTEMA DE PAQUETERÍA (CASETA)
// Google Apps Script — Web App independiente
// Spreadsheet: https://docs.google.com/spreadsheets/d/1GC2mxLZMtDO4-NmT2nruqNOIOOU2k6IdEkcXMyV_8oY
// ════════════════════════════════════════════════════════════════

var SS_ID = '1GC2mxLZMtDO4-NmT2nruqNOIOOU2k6IdEkcXMyV_8oY';

// ── Sesiones activas en memoria de script ───────────────────────
// (Se pierde al redesplegar, las sesiones duran ~6h por ejecución)
var _SESSIONS = {};

// ════════════════════════════════════════════════════════════════
// ENTRY POINT
// ════════════════════════════════════════════════════════════════
function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    var data = JSON.parse(e.postData.contents);
    var accion = data.action || data.accion || '';
    var ss = SpreadsheetApp.openById(SS_ID);
    var resultado;

    // ── Acción pública: login ──────────────────────────────────
    if (accion === 'login') {
      resultado = doLogin(data, ss);

    // ── Acciones protegidas ───────────────────────────────────
    } else if (accion === 'registrar-paquete')   { resultado = registrarPaquete(data, ss);
    } else if (accion === 'entregar-paquete')    { resultado = entregarPaquete(data, ss);
    } else if (accion === 'paquetes-pendientes') { resultado = getPaquetesPendientes(data, ss);
    } else if (accion === 'paquetes-admin')      { resultado = getPaquetesAdmin(data, ss);
    } else if (accion === 'get-contactos-wa')    { resultado = getContactosWA(data, ss);
    } else if (accion === 'save-contacto-wa')    { resultado = saveContactoWA(data, ss);

    } else {
      resultado = { error: 'Acción desconocida: ' + accion };
    }

    output.setContent(JSON.stringify(resultado));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.message }));
  }

  return output;
}

// También responder GET para verificar que el script está activo
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, servicio: 'Caseta Paquetería Minas 11' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
// AUTH
// ════════════════════════════════════════════════════════════════
function doLogin(data, ss) {
  var user = String(data.user || '').trim().toLowerCase();
  var pass = String(data.pass || '').trim();
  if (!user || !pass) return { error: 'Usuario y contraseña requeridos' };

  var sh    = getOAsegurarHojaUsuarios(ss);
  var filas = sh.getDataRange().getValues();

  for (var i = 1; i < filas.length; i++) {
    var u = String(filas[i][0]).trim().toLowerCase();
    var p = String(filas[i][1]).trim();
    var nombre = String(filas[i][2]).trim();
    var rol    = String(filas[i][3]).trim().toLowerCase();
    if (u === user && p === pass) {
      var token = generarToken();
      // Guardar en Properties para persistencia entre invocaciones
      var props = PropertiesService.getScriptProperties();
      props.setProperty('tok_' + token, JSON.stringify({
        user: user, nombre: nombre, rol: rol,
        exp: Date.now() + 8 * 3600 * 1000  // 8 horas
      }));
      return { ok: true, token: token, userName: nombre, role: rol };
    }
  }
  return { error: 'Usuario o contraseña incorrectos' };
}

function generarToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var t = '';
  for (var i = 0; i < 32; i++) t += chars[Math.floor(Math.random() * chars.length)];
  return t;
}

function getUserFromToken(token) {
  if (!token) return null;
  try {
    var props = PropertiesService.getScriptProperties();
    var raw   = props.getProperty('tok_' + token);
    if (!raw) return null;
    var obj = JSON.parse(raw);
    if (Date.now() > obj.exp) { props.deleteProperty('tok_' + token); return null; }
    return obj;
  } catch(e) { return null; }
}

function checkAuth(token) {
  var u = getUserFromToken(token);
  if (!u) return null;
  // Roles válidos para caseta
  if (['caseta','admin'].indexOf(u.rol) === -1) return null;
  return u;
}

// ════════════════════════════════════════════════════════════════
// HOJAS — INICIALIZACIÓN AUTOMÁTICA
// ════════════════════════════════════════════════════════════════
function getOAsegurarHojaUsuarios(ss) {
  var sh = ss.getSheetByName('Usuarios');
  if (!sh) {
    sh = ss.insertSheet('Usuarios');
    sh.appendRow(['usuario', 'password', 'nombre', 'rol']);
    sh.appendRow(['guardia01', 'Caseta2025', 'Guardia Principal', 'caseta']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,4).setFontWeight('bold')
      .setBackground('#8b3a0f').setFontColor('#ffffff');
  }
  return sh;
}

function getOAsegurarHojaPaqueteria(ss) {
  var sh = ss.getSheetByName('Paqueteria');
  if (!sh) {
    sh = ss.insertSheet('Paqueteria');
    sh.appendRow([
      'Folio','Departamento','Propietario','Tracking','Courier',
      'FechaEntrada','HoraEntrada','Guardia','Estado',
      'FechaSalida','HoraSalida','QuienRecibio'
    ]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,12).setFontWeight('bold')
      .setBackground('#c49a22').setFontColor('#ffffff');
    // Anchos de columna cómodos
    sh.setColumnWidth(1, 160);   // Folio
    sh.setColumnWidth(2, 110);   // Departamento
    sh.setColumnWidth(3, 160);   // Propietario
    sh.setColumnWidth(4, 200);   // Tracking
    sh.setColumnWidth(5, 120);   // Courier
    sh.setColumnWidth(9, 100);   // Estado
  }
  return sh;
}

function getOAsegurarHojaContactosWA(ss) {
  var sh = ss.getSheetByName('Contactos_WA');
  if (!sh) {
    sh = ss.insertSheet('Contactos_WA');
    sh.appendRow(['Departamento','Nombre','Telefono','ApiKey_Callmebot']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,4).setFontWeight('bold')
      .setBackground('#3a5c3c').setFontColor('#ffffff');
    sh.setColumnWidth(1, 120);
    sh.setColumnWidth(2, 180);
    sh.setColumnWidth(3, 160);
    sh.setColumnWidth(4, 160);
  }
  return sh;
}

// ════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════
function json(obj) { return obj; }

function generarFolio(sh) {
  var ahora = new Date();
  var yy = String(ahora.getFullYear()).slice(2);
  var mm = String(ahora.getMonth()+1).padStart(2,'0');
  var dd = String(ahora.getDate()).padStart(2,'0');
  var seq = String(Math.max(sh.getLastRow(), 1)).padStart(4,'0');
  return 'PKG-' + yy + mm + dd + '-' + seq;
}

function getContactoWA(dept, ss) {
  var sh    = getOAsegurarHojaContactosWA(ss);
  var datos = sh.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(dept).trim()) {
      return { departamento: datos[i][0], nombre: datos[i][1],
               telefono: datos[i][2], apikey: datos[i][3], fila: i+1 };
    }
  }
  return null;
}

function sendWhatsApp(telefono, apikey, mensaje) {
  if (!telefono || !apikey) return 'sin-contacto';
  try {
    var url = 'https://api.callmebot.com/whatsapp.php'
      + '?phone='  + encodeURIComponent(String(telefono))
      + '&text='   + encodeURIComponent(mensaje)
      + '&apikey=' + encodeURIComponent(String(apikey));
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return r.getResponseCode() === 200 ? 'ok' : ('error-' + r.getResponseCode());
  } catch(err) { return 'error-' + err.message; }
}

function fmtFecha(d) {
  return Utilities.formatDate(d, 'America/Mexico_City', 'yyyy-MM-dd');
}
function fmtHora(d) {
  return Utilities.formatDate(d, 'America/Mexico_City', 'HH:mm:ss');
}

// ════════════════════════════════════════════════════════════════
// REGISTRAR PAQUETE
// ════════════════════════════════════════════════════════════════
function registrarPaquete(data, ss) {
  var u = checkAuth(data.token);
  if (!u) return { error: 'No autorizado' };

  var dept     = String(data.departamento || '').trim();
  var tracking = String(data.tracking     || '').trim();
  var courier  = String(data.courier      || '').trim();

  if (!dept)    return { error: 'Departamento es obligatorio' };
  if (!courier) return { error: 'Courier es obligatorio' };

  var sh    = getOAsegurarHojaPaqueteria(ss);
  var folio = generarFolio(sh);
  var ahora = new Date();

  sh.appendRow([
    folio,            // Folio
    dept,             // Departamento
    '',               // Propietario (se puede llenar después)
    tracking,         // Tracking
    courier,          // Courier
    fmtFecha(ahora),  // FechaEntrada
    fmtHora(ahora),   // HoraEntrada
    u.nombre,         // Guardia
    'pendiente',      // Estado
    '',               // FechaSalida
    '',               // HoraSalida
    ''                // QuienRecibio
  ]);

  // WhatsApp al residente
  var contacto = getContactoWA(dept, ss);
  var waStatus = 'sin-contacto';
  if (contacto && contacto.telefono && contacto.apikey) {
    var msg = '📦 *Nuevo paquete en caseta*\n'
      + 'Depto: *' + dept + '*\n'
      + 'Paquetería: ' + courier + '\n'
      + (tracking ? 'Guía: ' + tracking + '\n' : '')
      + 'Folio: ' + folio + '\n'
      + 'Recibido: ' + fmtFecha(ahora) + ' a las ' + fmtHora(ahora).slice(0,5) + '\n'
      + '🏠 Pasa a recogerlo en caseta cuando gustes.';
    waStatus = sendWhatsApp(contacto.telefono, contacto.apikey, msg);
  }

  return { ok: true, folio: folio, whatsapp: waStatus };
}

// ════════════════════════════════════════════════════════════════
// ENTREGAR PAQUETE
// ════════════════════════════════════════════════════════════════
function entregarPaquete(data, ss) {
  var u = checkAuth(data.token);
  if (!u) return { error: 'No autorizado' };

  var folios = data.folios;
  if (!folios || !folios.length) return { error: 'Se requiere al menos un folio' };

  var sh      = getOAsegurarHojaPaqueteria(ss);
  var filas   = sh.getDataRange().getValues();
  var ahora   = new Date();
  var fecha   = fmtFecha(ahora);
  var hora    = fmtHora(ahora);
  var entDept = '';
  var marcados = 0;

  for (var i = 1; i < filas.length; i++) {
    var folio = String(filas[i][0]).trim();
    if (folios.indexOf(folio) >= 0 &&
        String(filas[i][8]).trim().toLowerCase() === 'pendiente') {
      sh.getRange(i+1,  9).setValue('entregado');
      sh.getRange(i+1, 10).setValue(fecha);
      sh.getRange(i+1, 11).setValue(hora);
      sh.getRange(i+1, 12).setValue(u.nombre);
      entDept = String(filas[i][1]).trim();
      marcados++;
    }
  }

  if (marcados === 0) return { error: 'No se encontraron paquetes pendientes con esos folios' };

  // WhatsApp confirmación de entrega
  var waStatus = 'sin-contacto';
  if (entDept) {
    var contacto = getContactoWA(entDept, ss);
    if (contacto && contacto.telefono && contacto.apikey) {
      var msg2 = '✅ *Paquete(s) entregados*\n'
        + 'Depto: *' + entDept + '*\n'
        + marcados + ' paquete(s) entregados el ' + fecha + ' a las ' + hora.slice(0,5) + '\n'
        + 'Entregó: ' + u.nombre;
      waStatus = sendWhatsApp(contacto.telefono, contacto.apikey, msg2);
    }
  }

  return { ok: true, marcados: marcados, whatsapp: waStatus };
}

// ════════════════════════════════════════════════════════════════
// PAQUETES PENDIENTES
// ════════════════════════════════════════════════════════════════
function getPaquetesPendientes(data, ss) {
  var u = checkAuth(data.token);
  if (!u) return { error: 'No autorizado' };

  var sh    = getOAsegurarHojaPaqueteria(ss);
  var filas = sh.getDataRange().getValues();
  var hdrs  = filas[0];
  var dept  = data.departamento ? String(data.departamento).trim() : null;
  var lista = [];

  for (var i = 1; i < filas.length; i++) {
    var obj = {};
    hdrs.forEach(function(h, j) { obj[h] = filas[i][j]; });
    if (String(obj.Estado).trim().toLowerCase() !== 'pendiente') continue;
    if (dept && String(obj.Departamento).trim() !== dept) continue;
    lista.push(obj);
  }

  return { ok: true, paquetes: lista, total: lista.length };
}

// ════════════════════════════════════════════════════════════════
// HISTORIAL COMPLETO (ADMIN)
// ════════════════════════════════════════════════════════════════
function getPaquetesAdmin(data, ss) {
  var u = checkAuth(data.token);
  if (!u) return { error: 'No autorizado' };

  var sh     = getOAsegurarHojaPaqueteria(ss);
  var filas  = sh.getDataRange().getValues();
  var hdrs   = filas[0];
  var limite = data.limite ? parseInt(data.limite) : 100;
  var lista  = [];

  // Más recientes primero
  for (var i = filas.length - 1; i >= 1; i--) {
    if (lista.length >= limite) break;
    var obj = {};
    hdrs.forEach(function(h, j) { obj[h] = filas[i][j]; });
    lista.push(obj);
  }

  return { ok: true, paquetes: lista, total: filas.length - 1 };
}

// ════════════════════════════════════════════════════════════════
// CONTACTOS WHATSAPP — LEER
// ════════════════════════════════════════════════════════════════
function getContactosWA(data, ss) {
  var u = checkAuth(data.token);
  if (!u) return { error: 'No autorizado' };

  var sh    = getOAsegurarHojaContactosWA(ss);
  var filas = sh.getDataRange().getValues();
  var hdrs  = filas[0];
  var lista = [];

  for (var i = 1; i < filas.length; i++) {
    var obj = {};
    hdrs.forEach(function(h, j) { obj[h] = filas[i][j]; });
    if (obj.Departamento) lista.push(obj);
  }

  return { ok: true, contactos: lista };
}

// ════════════════════════════════════════════════════════════════
// CONTACTOS WHATSAPP — GUARDAR / ACTUALIZAR
// ════════════════════════════════════════════════════════════════
function saveContactoWA(data, ss) {
  var u = checkAuth(data.token);
  if (!u) return { error: 'No autorizado' };

  var dept   = String(data.departamento || '').trim();
  var nombre = String(data.nombre       || '').trim();
  var tel    = String(data.telefono     || '').trim();
  var apikey = String(data.apikey       || '').trim();

  if (!dept) return { error: 'Departamento es obligatorio' };
  if (!tel)  return { error: 'Teléfono es obligatorio' };

  var sh    = getOAsegurarHojaContactosWA(ss);
  var filas = sh.getDataRange().getValues();

  // Actualizar si ya existe
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === dept) {
      sh.getRange(i+1, 1, 1, 4).setValues([[dept, nombre, tel, apikey]]);
      return { ok: true, accion: 'actualizado', departamento: dept };
    }
  }

  // Crear nuevo
  sh.appendRow([dept, nombre, tel, apikey]);
  return { ok: true, accion: 'creado', departamento: dept };
}

// ════════════════════════════════════════════════════════════════
// INICIALIZACIÓN MANUAL (ejecutar una vez desde el editor)
// ════════════════════════════════════════════════════════════════
function inicializarHojas() {
  var ss = SpreadsheetApp.openById(SS_ID);
  getOAsegurarHojaUsuarios(ss);
  getOAsegurarHojaPaqueteria(ss);
  getOAsegurarHojaContactosWA(ss);
  Logger.log('✅ Hojas inicializadas correctamente');
  Logger.log('👤 Usuario por defecto: guardia01 / Caseta2025');
}

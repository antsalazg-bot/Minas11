// ═══════════════════════════════════════════════
//  CONFIGURACIÓN — Solo editar aquí
// ═══════════════════════════════════════════════
var ADMIN_USER = 'admin';
var ADMIN_PASS = 'Minas2025';
var ONESIGNAL_APP_ID  = 'e6b778b4-c510-4ded-886e-1b3821b6a14a';
var ONESIGNAL_API_KEY = 'os_v2_app_423xrngfcbg63cdodm4cdnvbjinrrz5cdutuwnes7ccsmxl7zqb4tzj43qdo4d6uiohwmvoddrd4wosj576gxsaipzc4jqd2malobqy';
var RECIBOS_FOLDER_ID = '1-4PtEcnhDD6V0VNlZstguFBbARiBywUW';
var ENVIAR_CORREOS    = false; // cambiar a true cuando estés listo para producción
// ═══════════════════════════════════════════════
//  AUTH
// ═══════════════════════════════════════════════
function generateToken(user, pass) {
  var raw = user + ':' + pass + ':minas11-salt-2025';
  return Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8
  ));
}
function isValidToken(token) {
  if (!token) return false;
  return token === generateToken(ADMIN_USER, ADMIN_PASS);
}
// ─── Roles y permisos ───────────────────────────────────────────────────────
function generateUserToken(usuario, password) {
  var raw = usuario + ':' + password + ':minas11-salt-2025';
  return Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8));
}
function getUserFromToken(token, ss) {
  if (!token) return null;
  if (token === generateToken(ADMIN_USER, ADMIN_PASS))
    return {usuario: ADMIN_USER, rol: 'admin', nombre: 'Administrador'};
  var sh = ss.getSheetByName('Usuarios');
  if (!sh) return null;
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][3]) continue; // Activo = FALSE
    var tok = generateUserToken(String(rows[i][0]).trim(), String(rows[i][4]).trim());
    if (tok === token)
      return {usuario: String(rows[i][0]).trim(), nombre: String(rows[i][1]).trim(),
              rol: String(rows[i][2]).trim().toLowerCase()};
  }
  return null;
}
var _PERMISOS = {
  editor:      ['append','editar','eliminar','cancelar-pago','guardar-cuota-extra',
                'eliminar-cuota-extra','editar-deuda-hist','quitar-multa','pwd'],
  operaciones: ['leer','leer-cuotas-extras','leer-tarifas','saldos-admin','detalle-depto',
                'generar-recibos-mes','generar-recibo','cancelar-recibo',
                'verificar-ahora','crear-hoja-mes','notificar'],
  consulta:    ['leer','leer-cuotas-extras','leer-tarifas','saldos-admin','detalle-depto']
};
function hasPermiso(userInfo, accion) {
  if (!userInfo) return false;
  if (userInfo.rol === 'admin') return true;
  return (_PERMISOS[userInfo.rol] || []).indexOf(accion) !== -1;
}
// ═══════════════════════════════════════════════
//  HELPERS GLOBALES
// ═══════════════════════════════════════════════
var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio',
             'Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
/**
 * Google Sheets auto-convierte "Diciembre 2025" → Date en locale español.
 * Esta función normaliza el valor almacenado en la columna Periodo/Mes
 * a un string "Mes Año" para poder comparar con el parámetro mes.
 */
function periodoAMes(val) {
  if (val instanceof Date) {
    return MESES[val.getMonth()] + ' ' + val.getFullYear();
  }
  return String(val).trim();
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function normDept(d) {
  return String(d||'').trim().toUpperCase()
    .replace(/^DEPTO\s*/i,'').replace(/-/g,'').trim();
}
function parseIndiviso(raw) {
  if (typeof raw === 'number') return raw > 1 ? raw/100 : raw;
  if (typeof raw === 'string') return parseFloat(raw.replace('%','').trim())/100 || 0;
  return 0;
}
function parseFecha(v) {
  if (!v) return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'string') {
    var p = v.split('/');
    if (p.length === 3) return new Date(+p[2], +p[1]-1, +p[0]);
  }
  return null;
}
// ═══════════════════════════════════════════════
//  getMesesActivos
// ═══════════════════════════════════════════════
function getMesesActivos(ss, hasta) {
  var nombres = ss.getSheets().map(function(s){ return s.getName(); });
  var activos = [];
  for (var i = 0; i < nombres.length; i++) {
    var parts = nombres[i].trim().split(' ');
    if (parts.length === 2 && MESES.indexOf(parts[0]) !== -1 && /^\d{4}$/.test(parts[1])) {
      activos.push(nombres[i]);
    }
  }
  activos.sort(function(a, b) {
    var pa = a.split(' '), pb = b.split(' ');
    var ya = parseInt(pa[1]), yb = parseInt(pb[1]);
    if (ya !== yb) return ya - yb;
    return MESES.indexOf(pa[0]) - MESES.indexOf(pb[0]);
  });
  if (hasta) {
    var idx = activos.indexOf(hasta);
    if (idx !== -1) activos = activos.slice(0, idx + 1);
  }
  return activos;
}
// ═══════════════════════════════════════════════
//  normMesAplica
// ═══════════════════════════════════════════════
function normMesAplica(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    var tz  = Session.getScriptTimeZone();
    var idx = parseInt(Utilities.formatDate(v, tz, 'M'), 10) - 1;
    var ano = Utilities.formatDate(v, tz, 'yyyy');
    return MESES[idx] + ' ' + ano;
  }
  var s = String(v || '').trim();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1);
}
// ═══════════════════════════════════════════════
//  TARIFAS
// ═══════════════════════════════════════════════
function getTarifaVigente(tarifasRows, concepto, mesNombre, anio) {
  var mesIdx = MESES.indexOf(mesNombre);
  if (mesIdx === -1) return 0;
  var fechaMes = new Date(anio, mesIdx, 1);
  var mejor = null, mejorFI = null;
  for (var i = 1; i < tarifasRows.length; i++) {
    if (String(tarifasRows[i][0]||'').trim() !== concepto) continue;
    var fi = parseFecha(tarifasRows[i][2]);
    var ff = parseFecha(tarifasRows[i][3]);
    if (!fi || fi > fechaMes) continue;
    if (ff && ff < fechaMes) continue;
    if (!mejorFI || fi > mejorFI) { mejor = tarifasRows[i]; mejorFI = fi; }
  }
  return mejor ? Number(mejor[1]||0) : 0;
}
// ═══════════════════════════════════════════════
//  CUOTAS EXTRAS
// ═══════════════════════════════════════════════
function getExtrasDelMes(extrasRows, sheetName, indiviso) {
  var ind = (indiviso && !isNaN(indiviso)) ? indiviso : 1;
  var total = 0;
  for (var i = 1; i < extrasRows.length; i++) {
    if (normMesAplica(extrasRows[i][2]).toLowerCase() !== sheetName.toLowerCase()) continue;
    var montoBase = Number(extrasRows[i][1]||0);
    var tipo = String(extrasRows[i][4]||'FLAT').trim().toUpperCase();
    total += (tipo === 'INDIVISO') ? Math.round(montoBase * ind * 100) / 100 : montoBase;
  }
  return total;
}
function getExtrasDetalleConMonto(extrasRows, sheetName, indiviso) {
  var ind = (indiviso && !isNaN(indiviso)) ? indiviso : 1;
  var result = [];
  var total   = 0;
  for (var i = 1; i < extrasRows.length; i++) {
    if (normMesAplica(extrasRows[i][2]).toLowerCase() !== sheetName.toLowerCase()) continue;
    var montoBase = Number(extrasRows[i][1]||0);
    if (!montoBase) continue;
    var tipo = String(extrasRows[i][4]||'FLAT').trim().toUpperCase();
    var monto = (tipo === 'INDIVISO') ? Math.round(montoBase * ind * 100) / 100 : montoBase;
    result.push({
      concepto:   String(extrasRows[i][0]||'').trim(),
      monto:      monto,
      tipo:       tipo,
      idConcepto: String(extrasRows[i][3]||'').trim()
    });
    total += monto;
  }
  return { total: total, extras: result };
}
// ═══════════════════════════════════════════════
//  getSaldosCompletos
// ═══════════════════════════════════════════════
function getSaldosCompletos(ss) {
  var sh = ss.getSheetByName('Saldos');
  if (!sh) return [];
  var rows = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      depto:           String(rows[i][0]).trim(),
      propietario:     String(rows[i][1]||'').trim(),
      indiviso:        parseIndiviso(rows[i][2]),
      cuotaVigente:    Number(rows[i][3]||0),
      deudaHist:       Number(rows[i][4]||0),
      deudaAcum:       Number(rows[i][5]||0),
      multas:          Number(rows[i][6]||0),
      total:           Number(rows[i][7]||0),
      ultimoMesPagado: (function(v){ return v instanceof Date ? Utilities.formatDate(v,'America/Mexico_City','dd/MM/yyyy') : String(v||''); })(rows[i][8]),
      actualizado:     (function(v){ return v instanceof Date ? Utilities.formatDate(v,'America/Mexico_City','dd/MM/yyyy') : String(v||''); })(rows[i][9])
    });
  }
  return result;
}
// ═══════════════════════════════════════════════
//  correrVerificacion
// ═══════════════════════════════════════════════
function pad2(n){ return n < 10 ? '0'+n : String(n); }
function fechaRowStr(f) {
  if (!f) return null;
  if (f instanceof Date && !isNaN(f)) {
    return { str: pad2(f.getDate())+'/'+pad2(f.getMonth()+1)+'/'+f.getFullYear(), dia: f.getDate() };
  }
  if (typeof f === 'number' && f > 1) {
    var d = new Date(Math.round((f - 25569) * 86400000));
    if (!isNaN(d)) return { str: pad2(d.getUTCDate())+'/'+pad2(d.getUTCMonth()+1)+'/'+d.getUTCFullYear(), dia: d.getUTCDate() };
  }
  if (typeof f === 'string') {
    var p = f.trim().split('/');
    if (p.length >= 2) { var dia = parseInt(p[0]); if (dia > 0) return { str: f.trim(), dia: dia }; }
  }
  return null;
}
function deptDeRow(row) {
  var deptoI = String(row[8]||'').trim();
  // Solo usar col I si contiene dígitos (es un nro de depto, no un concepto como 'ELV')
  if (deptoI && /\d/.test(deptoI)) return normDept(deptoI);
  var idConc = String(row[6]||'').trim();
  if (idConc) {
    var info = getConceptoYDepto(idConc);
    if (info && info.dept) return normDept(info.dept);
  }
  var concepto = String(row[2]||'').trim();
  if (concepto) {
    var mConc = concepto.match(/\b(\d{3}|PH\d)\s*$/i);
    if (mConc) return normDept(mConc[1]);
  }
  var b = String(row[1]||'').trim();
  if (b) return normDept(b);
  return null;
}
function getMorososBalance(morososRows, deptNorm) {
  for (var i = 1; i < morososRows.length; i++) {
    if (!morososRows[i][0]) continue;
    if (normDept(String(morososRows[i][0])) === deptNorm) {
      var total = 0;
      for (var c = 1; c < morososRows[i].length; c++) {
        total += Number(morososRows[i][c]||0);
      }
      return Math.round(total * 100) / 100;
    }
  }
  return 0;
}
function correrVerificacion(ss, ultimoMes) {
  // ── FIX: permitir llamar directo desde el editor sin argumentos ──
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  var saldosSh = ss.getSheetByName('Saldos');
  if (!saldosSh) return {ok:false, error:'Hoja Saldos no encontrada'};
  var tarifasSh = ss.getSheetByName('Tarifas');
  if (!tarifasSh) return {ok:false, error:'Hoja Tarifas no encontrada'};
  var extrasSh   = ss.getSheetByName('Cuotas Extras');
  var morososSh  = ss.getSheetByName('Morosos');
  var tarifasRows  = tarifasSh.getDataRange().getValues();
  var extrasRows   = extrasSh  ? extrasSh.getDataRange().getValues()  : [];
  var morososRows  = morososSh ? morososSh.getDataRange().getValues() : [];
  var meses = getMesesActivos(ss, ultimoMes);
  if (!meses.length) return {ok:false, error:'No hay hojas de meses activos'};
  var mesData = {};
  for (var m = 0; m < meses.length; m++) {
    var msh = ss.getSheetByName(meses[m]);
    if (msh) mesData[meses[m]] = msh.getDataRange().getValues();
  }
  var saldosRows = saldosSh.getDataRange().getValues();
  var ahora = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy');
  var deptosProcesados = 0;
  var multaLog = [['Depto','Propietario','Mes','DiaPago','Monto Ingreso','IDConcepto','MultaMonto']];
  // ── Tabla de detalle por rubro ──
  var detalleRows = [['Depto','Propietario','CM Adeudo','CE Adeudo','Multas','Total','Actualizado']];
  for (var i = 1; i < saldosRows.length; i++) {
    if (!saldosRows[i][0]) continue;
    var depto       = String(saldosRows[i][0]).trim();
    var propietario = String(saldosRows[i][1]||'').trim();
    var indiviso    = parseIndiviso(saldosRows[i][2]);
    var deptNorm    = normDept(depto);
    var deudaHist   = getMorososBalance(morososRows, deptNorm);
    var saldoCorr     = 0;
    var cmSaldoCorr   = 0;
    var ceSaldoCorr   = 0;
    var multasTotal   = 0;
    var ultimaFecha   = '';
    var cuotaVigenteD = 0;
    for (var m = 0; m < meses.length; m++) {
      var sheetName = meses[m];
      var mesParts  = sheetName.split(' ');
      var mesNombre = mesParts[0];
      var mesAnio   = parseInt(mesParts[1]);
      var rows = mesData[sheetName];
      if (!rows) continue;
      var cuotaTotal = getTarifaVigente(tarifasRows, 'CuotaTotal', mesNombre, mesAnio);
      var multaMonto = getTarifaVigente(tarifasRows, 'Multa', mesNombre, mesAnio);
      var cuotaOrd   = Math.round(indiviso * cuotaTotal * 100) / 100;
      var cuotaExtra = getExtrasDelMes(extrasRows, sheetName, indiviso);
      if (m === meses.length - 1) cuotaVigenteD = cuotaOrd;
      var creditoAntes = Math.max(0, -saldoCorr);
      saldoCorr   += cuotaOrd + cuotaExtra;
      cmSaldoCorr += cuotaOrd;
      ceSaldoCorr += cuotaExtra;
      var cmPendiente = Math.max(0, cuotaOrd - creditoAntes);
      var pagosMes = [];
      for (var r = 1; r < rows.length; r++) {
        var rowDept = deptDeRow(rows[r]);
        if (!rowDept || rowDept !== deptNorm) continue;
        if (String(rows[r][9]||'').toUpperCase() === 'CANCELADO') continue;
        var monto = Number(rows[r][3]||0);
        if (monto <= 0) continue;
        var fInfo = fechaRowStr(rows[r][0]);
        if (!fInfo) continue;
        pagosMes.push({
          dia:    fInfo.dia,
          str:    fInfo.str,
          monto:  monto,
          idConc: String(rows[r][6]||'').trim().toUpperCase(),
          nombre: String(rows[r][1]||'').trim()
        });
      }
      pagosMes.sort(function(a, b) { return a.dia - b.dia; });
      var multaMes    = false;
      var multaLogRow = null;
      for (var p = 0; p < pagosMes.length; p++) {
        var pago  = pagosMes[p];
        var esCM  = pago.idConc.indexOf('CM-') === 0;
        var esCE  = /^(CE|CV)-/.test(pago.idConc);
        ultimaFecha = pago.str;
        if (esCM && pago.dia > 10 && cmPendiente > 0 && !multaMes) {
          multaMes    = true;
          multaLogRow = [depto, propietario, sheetName, pago.dia, pago.monto, pago.idConc, multaMonto];
        }
        saldoCorr -= pago.monto;
        if (esCM)      cmSaldoCorr -= pago.monto;
        else if (esCE) ceSaldoCorr -= pago.monto;
        if (esCM) cmPendiente = Math.max(0, cmPendiente - pago.monto);
        if (!propietario && pago.nombre) propietario = pago.nombre;
      }
      if (multaMes) {
        multasTotal += multaMonto;
        if (multaLogRow) {
          multaLogRow[1] = propietario || multaLogRow[1];
          multaLog.push(multaLogRow);
        }
      }
    }
    var deudaAcum   = Math.max(0, Math.round(saldoCorr * 100) / 100);
    var cmDeuda     = Math.max(0, Math.round(cmSaldoCorr * 100) / 100);
    var ceDeuda     = Math.max(0, Math.round(ceSaldoCorr * 100) / 100);
    var totalAdeudo = Math.round((deudaHist + deudaAcum + multasTotal) * 100) / 100;
    var row = i + 1;
    if (propietario) saldosSh.getRange(row, 2).setValue(propietario);
    saldosSh.getRange(row, 4).setValue(cuotaVigenteD);
    saldosSh.getRange(row, 5).setValue(deudaHist);
    saldosSh.getRange(row, 6).setValue(deudaAcum);
    saldosSh.getRange(row, 7).setValue(multasTotal);
    saldosSh.getRange(row, 8).setValue(totalAdeudo);
    saldosSh.getRange(row, 9).setValue(ultimaFecha);
    saldosSh.getRange(row, 10).setValue(ahora);
    deptosProcesados++;
    // Acumular fila para Detalle Saldos
    detalleRows.push([depto, propietario || '', cmDeuda, ceDeuda, multasTotal, totalAdeudo, ahora]);
  }
  // ── Log Multas ──
  var logSh = ss.getSheetByName('Log Multas');
  if (!logSh) logSh = ss.insertSheet('Log Multas');
  else logSh.clearContents();
  if (multaLog.length > 1) {
    logSh.getRange(1, 1, multaLog.length, multaLog[0].length).setValues(multaLog);
    logSh.getRange(1, 1, 1, multaLog[0].length).setFontWeight('bold').setBackground('#f3f3f3');
    logSh.autoResizeColumns(1, multaLog[0].length);
  } else {
    logSh.getRange(1,1).setValue('Sin multas generadas en este período');
  }
  // ── Escribir hoja "Detalle Saldos" ──
  var detalleSh = ss.getSheetByName('Detalle Saldos');
  if (!detalleSh) detalleSh = ss.insertSheet('Detalle Saldos');
  else detalleSh.clearContents();
  detalleSh.getRange(1, 1, detalleRows.length, detalleRows[0].length).setValues(detalleRows);
  detalleSh.getRange(1, 1, 1, detalleRows[0].length)
    .setFontWeight('bold').setBackground('#f3f3f3');
  detalleSh.autoResizeColumns(1, detalleRows[0].length);
  return {ok:true, meses:meses.length, deptos:deptosProcesados, multas: multaLog.length - 1};
}
// ═══════════════════════════════════════════════
//  getNombrePropietario
// ═══════════════════════════════════════════════
function getNombrePropietario(dept) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Saldos');
  if (!sh) return dept;
  var rows = sh.getDataRange().getValues();
  var d = normDept(dept);
  for (var i = 1; i < rows.length; i++) {
    if (normDept(String(rows[i][0]||'')) === d) return String(rows[i][1]||dept).trim();
  }
  return dept;
}
// col K (índice 10) de la hoja Saldos
function getEmailPropietario(dept) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Saldos');
  if (!sh) return '';
  var rows = sh.getDataRange().getValues();
  var d = normDept(dept);
  for (var i = 1; i < rows.length; i++) {
    if (normDept(String(rows[i][0]||'')) === d) return String(rows[i][10]||'').trim();
  }
  return '';
}
// ═══════════════════════════════════════════════
//  getConceptoYDepto
// ═══════════════════════════════════════════════
function getConceptoYDepto(idConc) {
  if (!idConc) return null;
  var str = String(idConc).trim().toUpperCase();
  var parts = str.split('-');
  if (parts.length < 2) return null;
  var dept = parts[parts.length - 1];
  if (!/\d/.test(dept)) return null;
  var tipo = parts.slice(0, parts.length - 1).join('-');
  var conceptos = {
    'CM':      'Cuota Mantenimiento',
    'CE-ELV':  'Cuota Extra Elevador',
    'CE-CIS':  'Cuota Extra Cisternas',
    'CV-VIG':  'Cuota Vigilancia',
    'OI-INT':  'Otros Ingresos'
  };
  return { dept: dept, concepto: conceptos[tipo] || tipo };
}
// ═══════════════════════════════════════════════
//  Token de verificación de recibos
// ═══════════════════════════════════════════════
function generateReciboToken(folio, depto, mes, monto) {
  var raw = folio + ':' + depto + ':' + mes + ':' + Number(monto).toFixed(2) + ':minas11-recibo-2025';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(function(b){ return (b < 0 ? b + 256 : b).toString(16).padStart(2,'0'); }).join('').substring(0, 16);
}
// ═══════════════════════════════════════════════
//  generarRecibo
// ═══════════════════════════════════════════════
function generarRecibo(dept, nombre, mes, fechaPago, monto, concepto) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rs = ss.getSheetByName('Recibos');
    if (!rs) {
      rs = ss.insertSheet('Recibos');
      rs.appendRow(['Folio','Depto','Nombre','Mes','Fecha','Monto','Link','Estado','Token']);
    }
    // Asegurar encabezado columna Token
    if (!rs.getRange(1, 9).getValue()) rs.getRange(1, 9).setValue('Token');

    // Usar el año del PERIODO del recibo, no el año actual de impresión
    var year = new Date().getFullYear();
    var mesParts = String(mes).trim().split(' ');
    if (mesParts.length >= 2 && /^\d{4}$/.test(mesParts[mesParts.length - 1])) {
      year = parseInt(mesParts[mesParts.length - 1]);
    }
    var count = rs.getLastRow();
    var folio = 'REC-' + year + '-' + String(count).padStart(4,'0');
    var montoFmt = '$' + Number(monto).toLocaleString('es-MX', {minimumFractionDigits:2, maximumFractionDigits:2});
    var fechaEmision = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy');

    // ── Token de verificación ─────────────────────────────────────────────
    var token = generateReciboToken(folio, dept, mes, monto);
    var verificarUrl = 'https://antsalazg-bot.github.io/Minas11/verificar.html?f='
      + encodeURIComponent(folio) + '&t=' + token;

    // ── Obtener imagen QR (opcional — si falla se omite la imagen) ────────
    var qrBlob = null;
    try {
      var qrApiUrl = 'https://quickchart.io/qr?text='
        + encodeURIComponent(verificarUrl) + '&size=200&format=png';
      var qrResp = UrlFetchApp.fetch(qrApiUrl, {muteHttpExceptions: true});
      if (qrResp.getResponseCode() === 200) qrBlob = qrResp.getBlob().setName('qr.png');
    } catch(qrErr) { /* QR no disponible, continúa sin imagen */ }

    // ── Crear Google Doc con diseño estilizado ────────────────────────────
    var doc = DocumentApp.create(folio);
    var body = doc.getBody();
    body.setMarginTop(0); body.setMarginBottom(0);
    body.setMarginLeft(0); body.setMarginRight(0);

    var folioNum = folio.split('-')[2] || folio;

    body.setPageWidth(612);
    body.setPageHeight(qrBlob ? 649 : 534);

    // ── 1. HEADER (fondo oscuro) ──────────────────────────────────────────
    var hTbl = body.appendTable([['', '']]);
    hTbl.setBorderWidth(0);
    hTbl.setColumnWidth(0, 420); hTbl.setColumnWidth(1, 192);

    var hL = hTbl.getCell(0, 0);
    hL.setBackgroundColor('#0d1b2a');
    hL.setPaddingTop(16); hL.setPaddingBottom(16);
    hL.setPaddingLeft(20); hL.setPaddingRight(8);
    hL.getChild(0).asParagraph().editAsText()
      .setText('Real de Minas 11').setForegroundColor('#d4a017').setFontSize(17).setBold(true);
    hL.appendParagraph('Río San Ángel')
      .editAsText().setForegroundColor('#aaaaaa').setFontSize(9);
    hL.appendParagraph('Camino Real de Minas 11, Col. Lomas de los Ángeles Tetelpan')
      .editAsText().setForegroundColor('#778899').setFontSize(8);
    hL.appendParagraph('01790, Álvaro Obregón · Ciudad de México')
      .editAsText().setForegroundColor('#778899').setFontSize(8);

    var hR = hTbl.getCell(0, 1);
    hR.setBackgroundColor('#0d1b2a');
    hR.setPaddingTop(16); hR.setPaddingBottom(16); hR.setPaddingRight(20);
    var hrp1 = hR.getChild(0).asParagraph();
    hrp1.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    hrp1.editAsText().setText('F O L I O').setForegroundColor('#888888').setFontSize(8);
    var hrp2 = hR.appendParagraph('#' + folioNum);
    hrp2.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    hrp2.editAsText().setForegroundColor('#d4a017').setFontSize(26).setBold(true);
    var hrp3 = hR.appendParagraph(folio);
    hrp3.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    hrp3.editAsText().setForegroundColor('#888888').setFontSize(8);

    // ── 2. BANNER VERDE ───────────────────────────────────────────────────
    var bTbl = body.appendTable([['✓  PAGO VERIFICADO Y REGISTRADO']]);
    bTbl.setBorderWidth(0); bTbl.setColumnWidth(0, 612);
    var bCell = bTbl.getCell(0, 0);
    bCell.setBackgroundColor('#2d6a4f');
    bCell.setPaddingTop(9); bCell.setPaddingBottom(9); bCell.setPaddingLeft(20);
    bCell.getChild(0).asParagraph().editAsText()
      .setForegroundColor('#ffffff').setFontSize(10).setBold(true);

    // ── 3. TÍTULO DEL COMPROBANTE ─────────────────────────────────────────
    var cTbl = body.appendTable([['']]);
    cTbl.setBorderWidth(0); cTbl.setColumnWidth(0, 612);
    var cCell = cTbl.getCell(0, 0);
    cCell.setBackgroundColor('#f7f7f7');
    cCell.setPaddingTop(18); cCell.setPaddingBottom(2);
    cCell.setPaddingLeft(28); cCell.setPaddingRight(28);
    cCell.getChild(0).asParagraph().editAsText()
      .setText('COMPROBANTE OFICIAL DE PAGO').setForegroundColor('#999999').setFontSize(8);
    var recTxt = 'Recibo de Mantenimiento';
    var rPar = cCell.appendParagraph(recTxt);
    rPar.editAsText().setFontSize(22).setBold(true)
      .setForegroundColor(0, 9, '#1a1a1a')
      .setForegroundColor(10, recTxt.length - 1, '#c8860a');

    // ── 4. LÍNEA DORADA ───────────────────────────────────────────────────
    var lTbl = body.appendTable([['']]);
    lTbl.setBorderWidth(0); lTbl.setColumnWidth(0, 612);
    var lCell = lTbl.getCell(0, 0);
    lCell.setBackgroundColor('#d4a017');
    lCell.setPaddingTop(0); lCell.setPaddingBottom(0);
    lCell.setPaddingLeft(0); lCell.setPaddingRight(0);
    lCell.getChild(0).asParagraph().editAsText().setText('').setFontSize(2);

    // ── 5. PROPIETARIO ────────────────────────────────────────────────────
    var pTbl = body.appendTable([['']]);
    pTbl.setBorderWidth(0); pTbl.setColumnWidth(0, 612);
    var pCell = pTbl.getCell(0, 0);
    pCell.setBackgroundColor('#f7f7f7');
    pCell.setPaddingTop(14); pCell.setPaddingBottom(2);
    pCell.setPaddingLeft(28); pCell.setPaddingRight(28);
    pCell.getChild(0).asParagraph().editAsText()
      .setText('PROPIETARIO').setForegroundColor('#999999').setFontSize(8);
    pCell.appendParagraph(nombre).editAsText()
      .setForegroundColor('#1a1a1a').setFontSize(15).setBold(true);

    // ── 6. DEPTO + PERIODO ────────────────────────────────────────────────
    var dpTbl = body.appendTable([['', '']]);
    dpTbl.setBorderWidth(0); dpTbl.setColumnWidth(0, 306); dpTbl.setColumnWidth(1, 306);
    var dpL = dpTbl.getCell(0, 0);
    dpL.setBackgroundColor('#f7f7f7');
    dpL.setPaddingTop(10); dpL.setPaddingBottom(2); dpL.setPaddingLeft(28);
    dpL.getChild(0).asParagraph().editAsText()
      .setText('DEPARTAMENTO').setForegroundColor('#999999').setFontSize(8);
    dpL.appendParagraph(dept).editAsText()
      .setForegroundColor('#1a1a1a').setFontSize(12);
    var dpR = dpTbl.getCell(0, 1);
    dpR.setBackgroundColor('#f7f7f7');
    dpR.setPaddingTop(10); dpR.setPaddingBottom(2);
    dpR.getChild(0).asParagraph().editAsText()
      .setText('PERIODO').setForegroundColor('#999999').setFontSize(8);
    dpR.appendParagraph(mes).editAsText()
      .setForegroundColor('#1a1a1a').setFontSize(12);

    // ── 7. FECHAS ─────────────────────────────────────────────────────────
    var fTbl = body.appendTable([['', '']]);
    fTbl.setBorderWidth(0); fTbl.setColumnWidth(0, 306); fTbl.setColumnWidth(1, 306);
    var fL = fTbl.getCell(0, 0);
    fL.setBackgroundColor('#f7f7f7');
    fL.setPaddingTop(8); fL.setPaddingBottom(14); fL.setPaddingLeft(28);
    fL.getChild(0).asParagraph().editAsText()
      .setText('FECHA DE PAGO').setForegroundColor('#999999').setFontSize(8);
    fL.appendParagraph(fechaPago).editAsText()
      .setForegroundColor('#1a1a1a').setFontSize(12);
    var fR = fTbl.getCell(0, 1);
    fR.setBackgroundColor('#f7f7f7');
    fR.setPaddingTop(8); fR.setPaddingBottom(14);
    fR.getChild(0).asParagraph().editAsText()
      .setText('FECHA DE EMISIÓN').setForegroundColor('#999999').setFontSize(8);
    fR.appendParagraph(fechaEmision).editAsText()
      .setForegroundColor('#1a1a1a').setFontSize(12);

    // ── 8. CAJA TOTAL (oscuro) ────────────────────────────────────────────
    var tTbl = body.appendTable([['']]);
    tTbl.setBorderWidth(0); tTbl.setColumnWidth(0, 612);
    var tCell = tTbl.getCell(0, 0);
    tCell.setBackgroundColor('#0d1b2a');
    tCell.setPaddingTop(16); tCell.setPaddingBottom(16);
    tCell.setPaddingLeft(28); tCell.setPaddingRight(28);
    tCell.getChild(0).asParagraph().editAsText()
      .setText('TOTAL PAGADO').setForegroundColor('#aaaaaa').setFontSize(9);
    tCell.appendParagraph(montoFmt).editAsText()
      .setForegroundColor('#d4a017').setFontSize(28).setBold(true);
    tCell.appendParagraph(concepto).editAsText()
      .setForegroundColor('#7a9bb5').setFontSize(9);
    tCell.appendParagraph('El importe de este recibo no extingue adeudos ni pagos vencidos.')
      .editAsText().setForegroundColor('#aec6d8').setFontSize(8).setItalic(true);

    // ── 9. CAJA VERIFICACIÓN (crema) ──────────────────────────────────────
    var vTbl = body.appendTable([['']]);
    vTbl.setBorderWidth(0); vTbl.setColumnWidth(0, 612);
    var vCell = vTbl.getCell(0, 0);
    vCell.setBackgroundColor('#fef9ef');
    vCell.setPaddingTop(14); vCell.setPaddingBottom(14);
    vCell.setPaddingLeft(28); vCell.setPaddingRight(28);
    vCell.getChild(0).asParagraph().editAsText()
      .setText('Verificación de autenticidad')
      .setForegroundColor('#1a1a1a').setFontSize(11).setBold(true);
    vCell.appendParagraph('Recibo generado automáticamente por el sistema de administración del Condominio Real de Minas 11. Este documento es válido como comprobante de pago.')
      .editAsText().setForegroundColor('#666666').setFontSize(8);
    if (qrBlob) {
      var qrPara = vCell.appendParagraph('');
      qrPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      var qrImg = qrPara.appendInlineImage(qrBlob);
      qrImg.setWidth(110).setHeight(110);
    }
    var badgePar = vCell.appendParagraph('  ' + folio + '  ');
    badgePar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    badgePar.editAsText()
      .setForegroundColor('#ffffff').setFontSize(10).setBold(true)
      .setBackgroundColor('#0d1b2a');
    vCell.appendParagraph(verificarUrl)
      .editAsText().setForegroundColor('#aaaaaa').setFontSize(7).setItalic(true);

    // ── 10. FOOTER (oscuro) ───────────────────────────────────────────────
    var ftTbl = body.appendTable([['', '']]);
    ftTbl.setBorderWidth(0); ftTbl.setColumnWidth(0, 420); ftTbl.setColumnWidth(1, 192);
    var ftL = ftTbl.getCell(0, 0);
    ftL.setBackgroundColor('#0d1b2a');
    ftL.setPaddingTop(10); ftL.setPaddingBottom(10); ftL.setPaddingLeft(20);
    ftL.getChild(0).asParagraph().editAsText()
      .setText('© ' + year + ' Real de Minas 11 · Todos los derechos reservados')
      .setForegroundColor('#777777').setFontSize(8);
    ftL.appendParagraph('Developed by Antonio Salazar')
      .editAsText().setForegroundColor('#555555').setFontSize(7);
    var ftR = ftTbl.getCell(0, 1);
    ftR.setBackgroundColor('#0d1b2a');
    ftR.setPaddingTop(10); ftR.setPaddingBottom(10); ftR.setPaddingRight(20);
    var ftRp = ftR.getChild(0).asParagraph();
    ftRp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    ftRp.editAsText().setText(folio)
      .setForegroundColor('#d4a017').setFontSize(9).setBold(true);

    // Minimizar párrafos separadores (SIN background color para no contaminar celdas)
    for (var si = 0; si < body.getNumChildren(); si++) {
      var el = body.getChild(si);
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var sp = el.asParagraph();
        sp.editAsText().setFontSize(1);
        sp.setSpacingBefore(0);
        sp.setSpacingAfter(0);
        sp.setLineSpacing(0.5);
      }
    }
    doc.saveAndClose();

    // Exportar a PDF
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(folio + '.pdf');

    // ── Carpeta del año dentro de RECIBOS_FOLDER_ID ──────────────────────
    var mainFolder = DriveApp.getFolderById(RECIBOS_FOLDER_ID);
    var yearStr = String(year);
    var yearFolders = mainFolder.getFoldersByName(yearStr);
    var yearFolder = yearFolders.hasNext() ? yearFolders.next() : mainFolder.createFolder(yearStr);

    var pdfFile = yearFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    docFile.setTrashed(true);

    var link = pdfFile.getUrl();
    rs.appendRow([folio, dept, nombre, mes, fechaPago, monto, link, 'activo', token]);

    // ── Envío de correo (solo si ENVIAR_CORREOS = true y hay email) ───────
    if (ENVIAR_CORREOS) {
      var emailDest = getEmailPropietario(dept);
      if (emailDest) {
        try {
          MailApp.sendEmail({
            to: emailDest,
            subject: 'Recibo de pago · ' + mes + ' · Real de Minas 11',
            body: 'Estimado(a) ' + nombre + ',\n\n' +
              'Adjunto encontrará su recibo de pago correspondiente al mes de ' + mes + '.\n\n' +
              'Folio: ' + folio + '\n' +
              'Monto: ' + montoFmt + '\n' +
              'Fecha de pago: ' + fechaPago + '\n\n' +
              'También puede verificar la autenticidad de su recibo escaneando el código QR incluido en el documento.\n\n' +
              'Administración · Real de Minas 11',
            attachments: [pdfFile.getAs('application/pdf')]
          });
        } catch(mailErr) {
          // El correo falló pero el recibo ya fue guardado — no interrumpir
          Logger.log('Error enviando correo a ' + emailDest + ': ' + mailErr.toString());
        }
      }
    }

    return {ok:true, folio:folio, link:link};
  } catch(e) {
    return {ok:false, error:e.toString()};
  }
}
// ═══════════════════════════════════════════════
//  doPost — router principal
// ═══════════════════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.accion === 'login') {
      var u = (data.user || '').trim().toLowerCase();
      var p = (data.pass || '').trim();
      // Buscar en hoja Usuarios primero (roles configurados)
      var usuSh = ss.getSheetByName('Usuarios');
      if (usuSh) {
        var usuRows = usuSh.getDataRange().getValues();
        for (var ui = 1; ui < usuRows.length; ui++) {
          if (!usuRows[ui][3]) continue; // Activo = FALSE
          if (String(usuRows[ui][0]).trim().toLowerCase() === u && String(usuRows[ui][4]).trim() === p)
            return json({ok:true, token: generateUserToken(String(usuRows[ui][0]).trim(), p),
                         rol: String(usuRows[ui][2]).trim().toLowerCase(),
                         nombre: String(usuRows[ui][1]).trim()});
        }
      }
      // Fallback: admin hardcodeado (backwards compat)
      if (u === ADMIN_USER.toLowerCase() && p === ADMIN_PASS)
        return json({ok: true, token: generateToken(ADMIN_USER, ADMIN_PASS), rol: 'admin', nombre: 'Administrador'});
      return json({ok: false, error: 'Credenciales incorrectas'});
    }
    if (data.accion === 'login-portal') {
      var dept = String(data.dept || '').trim().toUpperCase();
      var pass = String(data.pass || '').trim();
      var pwdSheet = ss.getSheetByName('Contraseñas');
      if (!pwdSheet) return json({ok: false, error: 'Hoja Contraseñas no encontrada'});
      var values = pwdSheet.getDataRange().getValues();
      for (var i = 0; i < values.length; i++) {
        var cellDept = String(values[i][0] || '').trim();
        var cellPass = String(values[i][1] || '').trim();
        var norm = cellDept.replace(/^Depto\s*/i, '').replace(/-/g, '').trim().toUpperCase();
        if ((norm === dept || cellDept.toUpperCase() === dept) && cellPass === pass) {
          // ── Registrar acceso en hoja Logs ────────────────────────────────
          try {
            var logSh = ss.getSheetByName('Logs');
            if (!logSh) {
              logSh = ss.insertSheet('Logs');
              logSh.appendRow(['Timestamp','Depto','Acción','Dispositivo','Resultado']);
              logSh.getRange(1,1,1,5).setFontWeight('bold');
              logSh.setFrozenRows(1);
            }
            var ua = String(data.userAgent || '').trim();
            var dispositivo = ua.indexOf('iPhone') !== -1 ? 'iPhone' :
                              ua.indexOf('iPad')   !== -1 ? 'iPad'   :
                              ua.indexOf('Android') !== -1 ? 'Android' :
                              ua.indexOf('Mac')    !== -1 ? 'Mac'    :
                              ua.indexOf('Windows')!== -1 ? 'Windows' : 'Otro';
            logSh.appendRow([new Date(), cellDept, 'login', dispositivo, 'ok']);
          } catch(le) { /* no bloquear el login por error de log */ }
          return json({ok: true, token: generateToken(dept, pass)});
        }
      }
      // Registrar intento fallido
      try {
        var logShF = ss.getSheetByName('Logs');
        if (logShF) logShF.appendRow([new Date(), dept, 'login', '', 'fallido']);
      } catch(le2) {}
      return json({ok: false, error: 'Credenciales incorrectas'});
    }
    if (data.accion === 'get-logs') {
      if (!getUserFromToken(data.token, ss)) return json({ok:false, error:'No autorizado'});
      var logSh2 = ss.getSheetByName('Logs');
      if (!logSh2) return json({ok:true, logs:[]});
      var logRows = logSh2.getDataRange().getValues();
      var logs = [];
      for (var li = 1; li < logRows.length; li++) {
        var ts = logRows[li][0];
        logs.push({
          ts:    ts instanceof Date ? Utilities.formatDate(ts,'America/Mexico_City','dd/MM/yyyy HH:mm') : String(ts),
          depto: String(logRows[li][1]||''),
          accion:String(logRows[li][2]||'login'),
          disp:  String(logRows[li][3]||''),
          res:   String(logRows[li][4]||'ok')
        });
      }
      logs.reverse(); // más recientes primero
      return json({ok:true, logs:logs});
    }
    // ── Resolver usuario de admin portal desde token ──────────────────────
    var currentUser = getUserFromToken(data.token, ss);
    if (data.accion === 'get-user-info') {
      if (!currentUser) return json({ok: false, error: 'Token inválido'});
      return json({ok: true, rol: currentUser.rol, nombre: currentUser.nombre});
    }
    if (data.accion === 'leer') {
      if (!data.token) return json({ok: false, error: 'No autorizado'});
      var sheet = ss.getSheetByName(data.hoja);
      if (!sheet) return json({ok: false, error: 'Hoja no encontrada: ' + data.hoja});
      var rows = sheet.getDataRange().getValues().map(function(row) {
        return { c: row.map(function(cell) {
          return (cell === '' || cell === null || cell === undefined) ? null : {v: cell};
        })};
      });
      return json({ok: true, table: {rows: rows}});
    }
    if (data.accion === 'verificar-recibo') {
      // Acción pública — no requiere token de sesión
      var vFolio = String(data.folio || '').trim();
      var vToken = String(data.token || '').trim();
      if (!vFolio || !vToken) return json({ok:false, error:'Datos incompletos'});
      var vRs = ss.getSheetByName('Recibos');
      if (!vRs) return json({ok:true, valido:false, mensaje:'Sin registros de recibos'});
      var vRows = vRs.getDataRange().getValues();
      for (var vi = 1; vi < vRows.length; vi++) {
        if (String(vRows[vi][0]).trim() !== vFolio) continue;
        var storedTok = String(vRows[vi][8] || '').trim();
        if (!storedTok) return json({ok:true, valido:false, mensaje:'Este recibo no tiene verificación QR'});
        if (storedTok !== vToken) return json({ok:true, valido:false, mensaje:'Token de verificación inválido'});
        var vEstado = String(vRows[vi][7] || 'activo').trim().toLowerCase();
        var vFecha = vRows[vi][4];
        if (vFecha instanceof Date) vFecha = Utilities.formatDate(vFecha, 'America/Mexico_City', 'dd/MM/yyyy');
        return json({ok:true, valido: vEstado === 'activo', estado: vEstado,
          folio: String(vRows[vi][0]), depto: String(vRows[vi][1]),
          nombre: String(vRows[vi][2]), mes: String(vRows[vi][3]),
          fecha: String(vFecha), monto: Number(vRows[vi][5])});
      }
      return json({ok:true, valido:false, mensaje:'Folio no encontrado'});
    }
    if (data.accion === 'mis-recibos') {
      if (!data.token) return json({ok: false, error: 'No autorizado'});
      var dept = String(data.dept || '').trim().toUpperCase();
      var rs = ss.getSheetByName('Recibos');
      if (!rs) return json({ok: true, recibos: []});
      var rows = rs.getDataRange().getValues();
      var recibos = [];
      for (var i = 1; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        if (String(rows[i][1]).trim().toUpperCase() === dept) {
          recibos.push({
            folio: rows[i][0], depto: rows[i][1], nombre: rows[i][2],
            mes: rows[i][3], fecha: rows[i][4], monto: rows[i][5],
            link: rows[i][6], estado: rows[i][7] || 'activo'
          });
        }
      }
      recibos.reverse();
      return json({ok: true, recibos: recibos});
    }
    if (data.accion === 'cancelar-recibo') {
      if (!hasPermiso(currentUser, 'cancelar-recibo')) return json({ok: false, error: 'No autorizado'});
      var rs = ss.getSheetByName('Recibos');
      if (!rs) return json({ok: false, error: 'Hoja Recibos no encontrada'});
      var rows = rs.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(data.folio).trim()) {
          rs.getRange(i + 1, 8).setValue('cancelado');
          return json({ok: true});
        }
      }
      return json({ok: false, error: 'Folio no encontrado'});
    }
    if (data.accion === 'saldo-depto') {
      if (!data.token) return json({ok: false, error: 'No autorizado'});
      var dept = String(data.dept || '').trim().toUpperCase();
      var todos = getSaldosCompletos(ss);
      for (var k = 0; k < todos.length; k++) {
        if (todos[k].depto === dept) return json({ok: true, saldo: todos[k]});
      }
      return json({ok: true, saldo: null});
    }
    // ── Validación global: se requiere usuario con rol válido ─────────────
    if (!currentUser) return json({ok: false, error: 'No autorizado'});
    if (data.accion === 'saldos-admin')
      return json({ok: true, saldos: getSaldosCompletos(ss)});
    if (data.accion === 'verificar-ahora') {
      if (!hasPermiso(currentUser, 'verificar-ahora')) return json({ok:false, error:'No autorizado'});
      var resultado = correrVerificacion(ss, data.ultimoMes || null);
      return json(resultado);
    }
    if (data.accion === 'editar-deuda-hist') {
      if (!hasPermiso(currentUser, 'editar-deuda-hist')) return json({ok:false, error:'No autorizado'});
      var saldosSheet = ss.getSheetByName('Saldos');
      if (!saldosSheet) return json({ok: false, error: 'Hoja Saldos no encontrada'});
      var rows = saldosSheet.getDataRange().getValues();
      var dept = String(data.dept || '').trim().toUpperCase();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim().toUpperCase() === dept) {
          var nuevaHist = Number(data.monto);
          var deudaAcum = Number(rows[i][5] || 0);
          var multas    = Number(rows[i][6] || 0);
          saldosSheet.getRange(i+1, 5).setValue(nuevaHist);
          saldosSheet.getRange(i+1, 8).setValue(nuevaHist + deudaAcum + multas);
          saldosSheet.getRange(i+1, 10).setValue(Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy'));
          return json({ok: true});
        }
      }
      return json({ok: false, error: 'Depto no encontrado'});
    }
    if (data.accion === 'quitar-multa') {
      if (!hasPermiso(currentUser, 'quitar-multa')) return json({ok:false, error:'No autorizado'});
      var saldosSheet = ss.getSheetByName('Saldos');
      if (!saldosSheet) return json({ok: false, error: 'Hoja Saldos no encontrada'});
      var rows = saldosSheet.getDataRange().getValues();
      var dept = String(data.dept || '').trim().toUpperCase();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim().toUpperCase() === dept) {
          var deudaHist = Number(rows[i][4] || 0);
          var deudaAcum = Number(rows[i][5] || 0);
          saldosSheet.getRange(i+1, 7).setValue(0);
          saldosSheet.getRange(i+1, 8).setValue(deudaHist + deudaAcum);
          saldosSheet.getRange(i+1, 10).setValue(Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy'));
          return json({ok: true});
        }
      }
      return json({ok: false, error: 'Depto no encontrado'});
    }
    if (data.accion === 'leer-tarifas') {
      var sh = ss.getSheetByName('Tarifas');
      if (!sh) return json({ok:false, error:'Hoja Tarifas no encontrada'});
      var rows = sh.getDataRange().getValues();
      var tarifas = [];
      for (var i = 1; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        var fi = rows[i][2], ff = rows[i][3];
        tarifas.push({
          concepto:    String(rows[i][0]||'').trim(),
          monto:       Number(rows[i][1]||0),
          fechaInicio: fi instanceof Date ? Utilities.formatDate(fi,'America/Mexico_City','dd/MM/yyyy') : String(fi||''),
          fechaFin:    ff instanceof Date ? Utilities.formatDate(ff,'America/Mexico_City','dd/MM/yyyy') : String(ff||'')
        });
      }
      return json({ok:true, tarifas:tarifas});
    }
    if (data.accion === 'guardar-tarifa') {
      if (!hasPermiso(currentUser, 'guardar-tarifa')) return json({ok:false, error:'No autorizado'});
      var sh = ss.getSheetByName('Tarifas');
      if (!sh) return json({ok:false, error:'Hoja Tarifas no encontrada'});
      sh.appendRow([
        String(data.concepto||'').trim(),
        Number(data.monto||0),
        String(data.fechaInicio||'').trim(),
        String(data.fechaFin||'').trim()
      ]);
      return json({ok:true});
    }
    if (data.accion === 'crear-hoja-mes') {
      if (!hasPermiso(currentUser, 'crear-hoja-mes')) return json({ok:false, error:'No autorizado'});
      var mes = String(data.mes||'').trim();
      if (!mes) return json({ok:false, error:'Mes requerido'});
      var filas = data.filas || [];
      if (!filas.length) return json({ok:false, error:'Sin datos'});
      var existente = ss.getSheetByName(mes);
      if (existente && !data.sobreescribir) return json({ok:false, error:'La hoja "'+mes+'" ya existe', existe:true});
      if (existente) ss.deleteSheet(existente);
      var newSh = ss.insertSheet(mes);
      newSh.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
      return json({ok:true, filas:filas.length});
    }
    if (data.accion === 'generar-recibos-mes') {
      if (!hasPermiso(currentUser, 'generar-recibos-mes')) return json({ok:false, error:'No autorizado'});
      var mes = String(data.mes || '').trim();
      if (!mes) return json({ok: false, error: 'Mes requerido'});
      var sheet = ss.getSheetByName(mes);
      if (!sheet) return json({ok: false, error: 'Hoja no encontrada: ' + mes});
      var rows = sheet.getDataRange().getValues();
      var generados = 0, omitidos = 0, errores = 0;
      var saltados = []; // filas con monto pero sin dept detectable — para diagnóstico
      // Leer hoja Recibos UNA sola vez fuera del loop
      var rs = ss.getSheetByName('Recibos');
      var recibosCache = rs ? rs.getDataRange().getValues() : [];
      for (var i = 1; i < rows.length; i++) {
        var monto = rows[i][3];
        var idConc = String(rows[i][6] || '').trim();
        var nombreFila = String(rows[i][2] || '').trim();
        if (!monto || isNaN(monto) || Number(monto) <= 0) continue;
        // Saltar filas sin ningún tipo de identificador
        if (!idConc && !rows[i][8] && !nombreFila) continue;
        // Saltar filas donde col G sea un marcador de anulación, no un ID de concepto válido
        // Ej: "cancelado", "devuelto", "anulado", "reverso" → no generan recibo
        if (idConc && !/^(CM|CE|OI|CV)-/i.test(idConc)) {
          var idLow = idConc.toLowerCase();
          if (idLow === 'cancelado' || idLow === 'devuelto' || idLow === 'anulado' ||
              idLow === 'reverso' || idLow === 'void' || idLow === 'n/a') continue;
        }

        // ── Detección de depto con misma lógica de 3 capas que shared.js ────
        var info = getConceptoYDepto(idConc);

        // Fallback 1: Col I (depto directo, formato "201", "PH3", etc.)
        if (!info) {
          var colI = String(rows[i][8] || '').trim().toUpperCase();
          if (colI && /\d/.test(colI)) {
            info = { dept: colI, concepto: idConc || 'Pago' };
          }
        }

        // Fallback 2: número de depto al final del nombre/concepto (col C)
        if (!info && nombreFila) {
          var mNom = nombreFila.match(/\b(\d{3}|PH\d)\s*$/i);
          if (mNom) {
            info = { dept: mNom[1].toUpperCase(), concepto: nombreFila };
          }
        }

        // Sin depto detectable — registrar para diagnóstico y saltar
        if (!info) {
          if (idConc) saltados.push('Fila ' + (i+1) + ': idConc="' + idConc + '" nombre="' + nombreFila + '"');
          continue;
        }

        var dept = info.dept;
        var concepto = (info.concepto || idConc || 'Pago') + ' · Depto ' + dept;

        // Determinar período del recibo a partir del concepto.
        // Prioridad: 1) "MES AÑO" explícito  2) solo "MES" → usa año del sheet  3) sheet name
        // Ej: "DEPTO 302 JUNIO 2025"  → "Junio 2025"
        // Ej: "ABRIL 302"             → "Abril 2025"  (año tomado del sheet)
        // Ej: "DEPARTAMENTO 302"      → "Marzo 2025"  (= nombre del sheet)
        var MESES_RE = 'Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre';
        var periodoRecibo = mes;
        var anoSheet = mes.split(' ')[1] || String(new Date().getFullYear());
        var mPeriodo = nombreFila.match(new RegExp('\\b(' + MESES_RE + ')\\s+(\\d{4})\\b', 'i'));
        if (mPeriodo) {
          // Mes + año explícito
          periodoRecibo = mPeriodo[1].charAt(0).toUpperCase() + mPeriodo[1].slice(1).toLowerCase() + ' ' + mPeriodo[2];
        } else {
          var mSoloMes = nombreFila.match(new RegExp('\\b(' + MESES_RE + ')\\b', 'i'));
          if (mSoloMes) {
            // Mes sin año → tomar año del sheet
            periodoRecibo = mSoloMes[1].charAt(0).toUpperCase() + mSoloMes[1].slice(1).toLowerCase() + ' ' + anoSheet;
          }
        }

        var isDup = false;
        for (var j = 1; j < recibosCache.length; j++) {
          var rDept_  = String(recibosCache[j][1]).toUpperCase();
          var rPer_   = periodoAMes(recibosCache[j][3]);
          var rMonto_ = Number(recibosCache[j][5]).toFixed(2);
          var rEst_   = String(recibosCache[j][7]).trim().toLowerCase();
          // Bloquear si mismo depto + periodo + monto — sin importar si está cancelado
          if (rDept_ === dept && rPer_ === periodoRecibo && rMonto_ === Number(monto).toFixed(2)) {
            isDup = true;
            Logger.log('SKIP ' + dept + ' ' + periodoRecibo + ' $' + monto + ' estado=' + rEst_);
            break;
          }
        }
        if (isDup) { omitidos++; continue; }
        var nombre = getNombrePropietario(dept);
        var fecha = rows[i][0];
        var fechaStr = '';
        try {
          if (fecha instanceof Date) {
            fechaStr = Utilities.formatDate(fecha, 'America/Mexico_City', 'dd/MM/yyyy');
          } else if (typeof fecha === 'number') {
            fechaStr = Utilities.formatDate(new Date((fecha-25569)*86400*1000), 'America/Mexico_City', 'dd/MM/yyyy');
          } else { fechaStr = String(fecha); }
        } catch(ex) { fechaStr = String(fecha); }
        var result = generarRecibo(dept, nombre, periodoRecibo, fechaStr, monto, concepto);
        if (result.ok) {
          generados++;
          // Refrescar cache para que duplicados dentro del mismo mes también se detecten
          recibosCache = rs ? rs.getDataRange().getValues() : [];
        } else {
          errores++;
          // Cuota de Google Docs agotada — no tiene caso seguir, retornar de inmediato
          if (result.error && result.error.indexOf('too many times') !== -1) {
            return json({ok: true, generados: generados, omitidos: omitidos, errores: errores,
              advertencia: 'Cuota diaria de Google Docs agotada. Reintentar mañana.'});
          }
        }
      }
      var advertencia = saltados.length > 0
        ? saltados.length + ' fila(s) sin depto detectable (no generaron recibo): ' + saltados.join('; ')
        : null;
      return json({ok: true, generados: generados, omitidos: omitidos, errores: errores, advertencia: advertencia});
    }
    // ── AUDITORÍA DE RECIBOS ─────────────────────────────────────────────
    if (data.accion === 'auditoria-recibos') {
      var mes = String(data.mes || '').trim();
      if (!mes) return json({ok: false, error: 'Mes requerido'});
      var sheet = ss.getSheetByName(mes);
      if (!sheet) return json({ok: false, error: 'Hoja no encontrada: ' + mes});

      var rows = sheet.getDataRange().getValues();
      var rs2  = ss.getSheetByName('Recibos');
      var rd   = rs2 ? rs2.getDataRange().getValues() : [];

      // Construir lookup de recibos: {dept_periodo_monto → {folio,estado,link,nombre}}
      var recibos = [];
      for (var k = 1; k < rd.length; k++) {
        recibos.push({
          folio:   String(rd[k][0]),
          dept:    String(rd[k][1]).toUpperCase().trim(),
          nombre:  String(rd[k][2] || ''),
          periodo: periodoAMes(rd[k][3]),
          monto:   Number(rd[k][5]),
          estado:  String(rd[k][7] || '').trim().toLowerCase(),
          link:    String(rd[k][6] || '')
        });
      }

      var MESES_RE2 = 'Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre';
      var anoSheet2 = mes.split(' ')[1] || String(new Date().getFullYear());
      var pagos = [];

      for (var i = 1; i < rows.length; i++) {
        var monto2     = rows[i][3];
        var idConc2    = String(rows[i][6] || '').trim();
        var nombre2    = String(rows[i][2] || '').trim();
        if (!monto2 || isNaN(monto2) || Number(monto2) <= 0) continue;
        if (!idConc2 && !rows[i][8] && !nombre2) continue;
        // Saltar marcadores de anulación
        if (idConc2 && !/^(CM|CE|OI|CV)-/i.test(idConc2)) {
          var idL = idConc2.toLowerCase();
          if (idL === 'cancelado' || idL === 'devuelto' || idL === 'anulado' || idL === 'reverso' || idL === 'void' || idL === 'n/a') continue;
        }
        // Detectar depto
        var info2 = getConceptoYDepto(idConc2);
        if (!info2) {
          var colI2 = String(rows[i][8] || '').trim().toUpperCase();
          if (colI2 && /\d/.test(colI2)) info2 = {dept: colI2, concepto: idConc2 || 'Pago'};
        }
        if (!info2 && nombre2) {
          var mN2 = nombre2.match(/\b(\d{3}|PH\d)\s*$/i);
          if (mN2) info2 = {dept: mN2[1].toUpperCase(), concepto: nombre2};
        }
        if (!info2) continue;
        // Determinar periodoRecibo
        var periodoRecibo2 = mes;
        var mP2 = nombre2.match(new RegExp('\\b(' + MESES_RE2 + ')\\s+(\\d{4})\\b', 'i'));
        if (mP2) {
          periodoRecibo2 = mP2[1].charAt(0).toUpperCase() + mP2[1].slice(1).toLowerCase() + ' ' + mP2[2];
        } else {
          var mSM2 = nombre2.match(new RegExp('\\b(' + MESES_RE2 + ')\\b', 'i'));
          if (mSM2) periodoRecibo2 = mSM2[1].charAt(0).toUpperCase() + mSM2[1].slice(1).toLowerCase() + ' ' + anoSheet2;
        }
        // Buscar recibo coincidente
        var recibo2 = null;
        for (var r = 0; r < recibos.length; r++) {
          if (recibos[r].dept === info2.dept &&
              recibos[r].periodo === periodoRecibo2 &&
              Math.abs(recibos[r].monto - Number(monto2)) < 0.05) {
            recibo2 = recibos[r]; break;
          }
        }
        // Formatear fecha
        var fRaw = rows[i][0], fechaStr2 = '';
        try {
          fechaStr2 = (fRaw instanceof Date)
            ? Utilities.formatDate(fRaw, 'America/Mexico_City', 'dd/MM/yyyy')
            : String(fRaw);
        } catch(ex) { fechaStr2 = String(fRaw); }

        pagos.push({
          fila:         i + 1,
          dept:         info2.dept,
          nombre:       getNombrePropietario(info2.dept),
          concepto:     nombre2,
          monto:        Number(monto2),
          fecha:        fechaStr2,
          periodoRecibo: periodoRecibo2,
          esAdelanto:   periodoRecibo2 !== mes,
          folio:        recibo2 ? recibo2.folio  : null,
          reciboEstado: recibo2 ? recibo2.estado : 'sin_recibo',
          link:         recibo2 ? recibo2.link   : null
        });
      }

      // Recibos activos para este periodo (incluyendo los adelantados de otros meses)
      var recibosDelPeriodo = recibos.filter(function(r) {
        return r.periodo === mes && r.estado !== 'cancelado';
      });

      return json({ok: true, mes: mes, pagos: pagos, recibosDelPeriodo: recibosDelPeriodo});
    }

    if (data.accion === 'generar-recibo')
      return json(generarRecibo(data.dept, data.nombre, data.mes, data.fechaPago, data.monto, data.concepto));
    if (data.accion === 'notificar') {
      if (!hasPermiso(currentUser, 'notificar')) return json({ok:false, error:'No autorizado'});
      UrlFetchApp.fetch('https://api.onesignal.com/notifications', {
        method: 'post', contentType: 'application/json',
        headers: { 'Authorization': 'Key ' + ONESIGNAL_API_KEY },
        payload: JSON.stringify({
          app_id: ONESIGNAL_APP_ID, included_segments: ['Total Subscriptions'],
          headings: {en: data.titulo}, contents: {en: data.mensaje},
          url: 'https://antsalazg-bot.github.io/Minas11/portal.html'
        }), muteHttpExceptions: true
      });
      return json({ok: true});
    }
    if (data.accion === 'eliminar') {
      if (!hasPermiso(currentUser, 'eliminar')) return json({ok:false, error:'No autorizado'});
      var sheet = ss.getSheetByName(data.hoja);
      if (!sheet) throw new Error('Hoja no encontrada: ' + data.hoja);
      sheet.deleteRow(data.rowIdx + 1);
      return json({ok: true});
    }
    if (data.accion === 'pwd') {
      if (!hasPermiso(currentUser, 'pwd')) return json({ok:false, error:'No autorizado'});
      var pwdSheet = ss.getSheetByName('Contraseñas');
      if (!pwdSheet) throw new Error('Hoja Contraseñas no encontrada');
      var values = pwdSheet.getDataRange().getValues();
      var found = false;
      // Normalizar: quitar "Depto", guiones, espacios, convertir a mayúsculas → solo alfanumérico
      var dvNorm = String(data.dept).toUpperCase().replace(/[^A-Z0-9]/g, '');
      for (var i = 0; i < values.length; i++) {
        var cvNorm = String(values[i][0]).toUpperCase()
          .replace(/^DEPTO/i, '').replace(/[^A-Z0-9]/g, '');
        if (cvNorm && cvNorm === dvNorm) {
          pwdSheet.getRange(i+1, 2).setValue(data.pwd);
          found = true; break;
        }
      }
      if (!found) pwdSheet.appendRow(['Depto ' + data.dept, data.pwd]);
      return json({ok: true, found: found});
    }
    if (data.accion === 'editar' && data.rowIdx !== undefined) {
      if (!hasPermiso(currentUser, 'editar')) return json({ok:false, error:'No autorizado'});
      var sheet = ss.getSheetByName(data.hoja);
      if (!sheet) throw new Error('Hoja no encontrada: ' + data.hoja);
      for (var col = 0; col < data.fila.length; col++) {
        if (data.fila[col] !== null && data.fila[col] !== undefined)
          sheet.getRange(data.rowIdx + 1, col + 1).setValue(data.fila[col]);
      }
      if (data.fila[3] !== null) { sheet.getRange(data.rowIdx+1,5).clearContent(); sheet.getRange(data.rowIdx+1,8).clearContent(); }
      if (data.fila[4] !== null) { sheet.getRange(data.rowIdx+1,4).clearContent(); sheet.getRange(data.rowIdx+1,7).clearContent(); }
      return json({ok: true});
    }
    if (data.accion === 'leer-cuotas-extras') {
      var extrasSh3 = ss.getSheetByName('Cuotas Extras');
      var extrasRows3 = extrasSh3 ? extrasSh3.getDataRange().getValues() : [];
      var extras3 = [];
      for (var i = 1; i < extrasRows3.length; i++) {
        var r3 = extrasRows3[i];
        if (!r3[0] && !r3[1]) continue;
        extras3.push({
          fila:       i + 1,
          concepto:   String(r3[0]||'').trim(),
          monto:      Number(r3[1]||0),
          mesAplica:  (r3[2] instanceof Date) ? Utilities.formatDate(r3[2], Session.getScriptTimeZone(), "MMMM yyyy") : String(r3[2]||'').trim(),
          idConcepto: String(r3[3]||'').trim(),
          tipo:       String(r3[4]||'FLAT').trim().toUpperCase() || 'FLAT'
        });
      }
      var saldosParaCE = getSaldosCompletos(ss).map(function(s) {
        return { depto: s.depto, propietario: s.propietario, indiviso: s.indiviso };
      });
      return json({ ok: true, extras: extras3, saldos: saldosParaCE });
    }
    if (data.accion === 'guardar-cuota-extra') {
      if (!hasPermiso(currentUser, 'guardar-cuota-extra')) return json({ok:false, error:'No autorizado'});
      var extrasSh4 = ss.getSheetByName('Cuotas Extras');
      if (!extrasSh4) {
        extrasSh4 = ss.insertSheet('Cuotas Extras');
        extrasSh4.appendRow(['Concepto','Monto','MesAplica','IDConcepto','Tipo']);
        extrasSh4.getRange(1,1,1,5).setFontWeight('bold');
      }
      var concepto4   = String(data.concepto   || '').trim();
      var monto4      = Number(data.monto       || 0);
      var mes4        = String(data.mesAplica   || '').trim();
      var idConc4     = String(data.idConcepto  || '').trim();
      var tipo4       = (String(data.tipo || 'FLAT').trim().toUpperCase() === 'INDIVISO') ? 'INDIVISO' : 'FLAT';
      if (!concepto4 || !monto4 || !mes4) return json({ ok: false, error: 'Faltan campos' });
      extrasSh4.appendRow([concepto4, monto4, mes4, idConc4, tipo4]);
      return json({ ok: true });
    }
    if (data.accion === 'eliminar-cuota-extra') {
      if (!hasPermiso(currentUser, 'eliminar-cuota-extra')) return json({ok:false, error:'No autorizado'});
      var extrasSh5 = ss.getSheetByName('Cuotas Extras');
      if (!extrasSh5) return json({ ok: false, error: 'Hoja "Cuotas Extras" no encontrada' });
      var fila5 = parseInt(data.fila || 0);
      if (!fila5 || fila5 < 2) return json({ ok: false, error: 'Fila inválida' });
      extrasSh5.deleteRow(fila5);
      return json({ ok: true });
    }
    if (data.accion === 'detalle-depto') {
      var dept = String(data.dept || '').trim();
      var deptNorm = normDept(dept);
      var saldosSh2 = ss.getSheetByName('Saldos');
      var tarifasSh2 = ss.getSheetByName('Tarifas');
      if (!saldosSh2 || !tarifasSh2) return json({ok:false, error:'Hojas faltantes'});
      var saldosRows2 = saldosSh2.getDataRange().getValues();
      var tarifasRows2 = tarifasSh2.getDataRange().getValues();
      var extrasSh2 = ss.getSheetByName('Cuotas Extras');
      var extrasRows2 = extrasSh2 ? extrasSh2.getDataRange().getValues() : [];
      var morososSh2 = ss.getSheetByName('Morosos');
      var morososRows2 = morososSh2 ? morososSh2.getDataRange().getValues() : [];
      var indiviso2 = 0, propietario2 = '';
      for (var si = 1; si < saldosRows2.length; si++) {
        if (normDept(String(saldosRows2[si][0]||'')) === deptNorm) {
          indiviso2 = parseIndiviso(saldosRows2[si][2]);
          propietario2 = String(saldosRows2[si][1]||'').trim();
          break;
        }
      }
      var meses2 = getMesesActivos(ss, null);
      var deudaHist2 = getMorososBalance(morososRows2, deptNorm);
      var saldoCorr2 = 0;
      var detalleMeses = [];
      for (var mi = 0; mi < meses2.length; mi++) {
        var sn2 = meses2[mi];
        var mp2 = sn2.split(' ');
        var ct2 = getTarifaVigente(tarifasRows2, 'CuotaTotal', mp2[0], parseInt(mp2[1]));
        var mm2 = getTarifaVigente(tarifasRows2, 'Multa', mp2[0], parseInt(mp2[1]));
        var co2 = Math.round(indiviso2 * ct2 * 100) / 100;
        var extrasInfo2 = getExtrasDetalleConMonto(extrasRows2, sn2, indiviso2);
        var ce2 = extrasInfo2.total;
        var creditoAntes2 = Math.max(0, -saldoCorr2);
        saldoCorr2 += co2 + ce2;
        var cmPend2 = Math.max(0, co2 - creditoAntes2);
        var msh2 = ss.getSheetByName(sn2);
        var pagados2 = [], multaMes2 = false;
        if (msh2) {
          var rows2 = msh2.getDataRange().getValues();
          var pm2 = [];
          for (var rr = 1; rr < rows2.length; rr++) {
            if (deptDeRow(rows2[rr]) !== deptNorm) continue;
            if (String(rows2[rr][9]||'').toUpperCase() === 'CANCELADO') continue;
            var mo2 = Number(rows2[rr][3]||0);
            if (mo2 <= 0) continue;
            var fi2 = fechaRowStr(rows2[rr][0]);
            if (!fi2) continue;
            pm2.push({dia:fi2.dia, str:fi2.str, monto:mo2, idConc:String(rows2[rr][6]||'').trim().toUpperCase()});
          }
          pm2.sort(function(a,b){return a.dia-b.dia;});
          for (var pp = 0; pp < pm2.length; pp++) {
            var pg = pm2[pp];
            var esCM2 = pg.idConc.indexOf('CM-') === 0;
            if (esCM2 && pg.dia > 10 && cmPend2 > 0 && !multaMes2) multaMes2 = true;
            saldoCorr2 -= pg.monto;
            if (esCM2) cmPend2 = Math.max(0, cmPend2 - pg.monto);
            pagados2.push({dia:pg.dia, monto:pg.monto, tipo:esCM2?'CM':'CE', idConc:pg.idConc});
          }
        }
        detalleMeses.push({
          mes: sn2,
          cuotaMant: co2,
          cuotaExtra: ce2,
          cuotaTotal: co2 + ce2,
          extras: extrasInfo2.extras,
          pagados: pagados2,
          totalPagado: pagados2.reduce(function(s,p){return s+p.monto;},0),
          multa: multaMes2 ? mm2 : 0,
          saldoCorriente: Math.round(saldoCorr2 * 100) / 100
        });
      }
      return json({ok:true, dept:dept, propietario:propietario2, deudaHist:deudaHist2, detalle:detalleMeses});
    }
    if (data.accion === 'debug-depto') {
      var dept = String(data.dept||'101');
      var dn = normDept(dept);
      var meses = getMesesActivos(ss, null);
      var log = [];
      for (var m = 0; m < meses.length; m++) {
        var sh = ss.getSheetByName(meses[m]);
        if (!sh) continue;
        var rows = sh.getDataRange().getValues();
        var pagos = [];
        for (var r = 1; r < rows.length; r++) {
          var rd = deptDeRow(rows[r]);
          if (rd === dn) {
            pagos.push({row:r+1, colI:String(rows[r][8]), colG:String(rows[r][6]), colD:rows[r][3], colA:String(rows[r][0])});
          }
        }
        log.push({mes:meses[m], pagos:pagos});
      }
      return json({ok:true, dept:dept, deptNorm:dn, meses:log});
    }
    if (data.accion === 'cancelar-pago') {
      if (!hasPermiso(currentUser, 'cancelar-pago')) return json({ok:false, error:'No autorizado'});
      var cpSheet = ss.getSheetByName(data.hoja);
      if (!cpSheet) return json({ok:false, error:'Hoja no encontrada: ' + data.hoja});
      var cpRowNum = data.rowIdx + 1;
      var cpLastCol = Math.max(cpSheet.getLastColumn(), 13);
      var cpRowVals = cpSheet.getRange(cpRowNum, 1, 1, cpLastCol).getValues()[0];
      var cpMonto = Number(cpRowVals[3]) || Number(cpRowVals[4]); // col D (ingreso) o col E (egreso)
      var cpIdConc = String(cpRowVals[6] || '').trim(); // col G: idConcepto
      // Marcar CANCELADO en col J + auditoría en K, L, M
      cpSheet.getRange(cpRowNum, 10).setValue('CANCELADO');
      cpSheet.getRange(cpRowNum, 11).setValue(data.motivo || '');
      cpSheet.getRange(cpRowNum, 12).setValue(new Date());
      cpSheet.getRange(cpRowNum, 13).setValue(currentUser ? currentUser.usuario : '');
      // Cancelar recibos asociados (mismo depto, mismo mes, mismo monto)
      var cpFolios = [];
      var cpRs = ss.getSheetByName('Recibos');
      if (cpRs && cpIdConc) {
        var cpInfo = getConceptoYDepto(cpIdConc);
        var cpDept = cpInfo ? cpInfo.dept : null;
        if (cpDept) {
          var cpRecRows = cpRs.getDataRange().getValues();
          for (var ci = 1; ci < cpRecRows.length; ci++) {
            if (String(cpRecRows[ci][1]).trim() === cpDept &&
                String(cpRecRows[ci][3]).trim() === data.hoja &&
                String(cpRecRows[ci][7]).trim() !== 'cancelado' &&
                Math.abs(Number(cpRecRows[ci][5]) - Number(cpMonto)) < 0.01) {
              cpRs.getRange(ci + 1, 8).setValue('cancelado');
              cpFolios.push(cpRecRows[ci][0]);
            }
          }
        }
      }
      return json({ok:true, foliosCancelados:cpFolios, montoOriginal:cpMonto});
    }
    // ── Default: guardarPago / guardarEgreso (append) ────────────────────
    if (!hasPermiso(currentUser, 'append')) return json({ok:false, error:'No autorizado'});
    var sheet = ss.getSheetByName(data.hoja);
    if (!sheet) throw new Error('Hoja no encontrada: ' + data.hoja);
    sheet.appendRow(data.fila);
    return json({ok: true});
  } catch(err) {
    return json({ok: false, error: err.toString()});
  }
}
// ═══════════════════════════════════════════════
//  testVerificacion
// ═══════════════════════════════════════════════
function testVerificacion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var saldosSh  = ss.getSheetByName('Saldos');
  var tarifasSh = ss.getSheetByName('Tarifas');
  if (!saldosSh)  { Logger.log('ERROR: Hoja Saldos no encontrada');  return; }
  if (!tarifasSh) { Logger.log('ERROR: Hoja Tarifas no encontrada'); return; }
  var tarifasRows = tarifasSh.getDataRange().getValues();
  var extrasSh    = ss.getSheetByName('Cuotas Extras');
  var extrasRows  = extrasSh ? extrasSh.getDataRange().getValues() : [];
  var meses       = getMesesActivos(ss, null);
  Logger.log('=== testVerificacion === Meses: ' + meses.join(', '));
  var mesData = {};
  for (var m = 0; m < meses.length; m++) {
    var sh = ss.getSheetByName(meses[m]);
    if (sh) mesData[meses[m]] = sh.getDataRange().getValues();
  }
  var saldosRows = saldosSh.getDataRange().getValues();
  Logger.log('\nDepto       | Indiviso | CuotaMens  | TotalCuotas | TotalPagos | DeudaAcum  | Multas  | Total');
  Logger.log('------------|----------|------------|-------------|------------|------------|---------|----------');
  for (var i = 1; i < saldosRows.length; i++) {
    if (!saldosRows[i][0]) continue;
    var depto     = String(saldosRows[i][0]).trim();
    var indiviso  = parseIndiviso(saldosRows[i][2]);
    var deudaHist = Number(saldosRows[i][4]||0);
    var deptNorm  = normDept(depto);
    var totalCuotas = 0, totalPagos = 0, multasTotal = 0, cuotaVigente = 0;
    for (var m = 0; m < meses.length; m++) {
      var sheetName = meses[m];
      var mesParts  = sheetName.split(' ');
      var cuotaTotal = getTarifaVigente(tarifasRows, 'CuotaTotal', mesParts[0], parseInt(mesParts[1]));
      var multaMonto = getTarifaVigente(tarifasRows, 'Multa',      mesParts[0], parseInt(mesParts[1]));
      var cuotaOrd   = Math.round(indiviso * cuotaTotal * 100) / 100;
      var cuotaExtra = getExtrasDelMes(extrasRows, sheetName, indiviso);
      totalCuotas += cuotaOrd + cuotaExtra;
      if (m === meses.length - 1) cuotaVigente = cuotaOrd;
      var rows = mesData[sheetName];
      if (!rows) continue;
      var pagadoMes = 0, multaMes = false;
      for (var r = 1; r < rows.length; r++) {
        if (deptDeRow(rows[r]) !== deptNorm) continue;
        var monto = Number(rows[r][3]||0);
        if (monto <= 0) continue;
        pagadoMes += monto;
        var fInfo = fechaRowStr(rows[r][0]);
        var idConcT = String(rows[r][6]||'').trim().toUpperCase();
        if (fInfo && fInfo.dia > 10 && multaMonto > 0 && idConcT.indexOf('CM-') === 0) multaMes = true;
      }
      totalPagos += pagadoMes;
      if (multaMes) multasTotal += multaMonto;
    }
    var deudaAcum = Math.max(0, Math.round((totalCuotas - totalPagos) * 100) / 100);
    var total     = Math.round((deudaHist + deudaAcum + multasTotal) * 100) / 100;
    Logger.log(
      pad2(depto).padEnd(11) + ' | ' +
      (indiviso*100).toFixed(4)+'%' + ' | ' +
      '$'+cuotaVigente.toFixed(0).padStart(9) + ' | ' +
      '$'+totalCuotas.toFixed(0).padStart(10) + ' | ' +
      '$'+totalPagos.toFixed(0).padStart(9) + ' | ' +
      '$'+deudaAcum.toFixed(0).padStart(9) + ' | ' +
      '$'+multasTotal.toFixed(0).padStart(6) + ' | ' +
      '$'+total.toFixed(0)
    );
  }
}
// ═══════════════════════════════════════════════
//  runVerificacionDirecta
// ═══════════════════════════════════════════════
function runVerificacionDirecta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var morososSh = ss.getSheetByName('Morosos');
  if (morososSh) {
    var morososRows = morososSh.getDataRange().getValues();
    Logger.log('=== Morosos FOUND — ' + (morososRows.length - 1) + ' deptos ===');
    for (var i = 1; i < Math.min(6, morososRows.length); i++) {
      if (!morososRows[i][0]) continue;
      var bal = getMorososBalance(morososRows, normDept(String(morososRows[i][0])));
      Logger.log('  Depto ' + morososRows[i][0] + ' → balance=' + bal);
    }
  } else {
    Logger.log('AVISO: Hoja Morosos no encontrada — deudaHist = 0 para todos');
  }
  var meses = getMesesActivos(ss, null);
  var ultimoMes = meses.length > 0 ? meses[meses.length - 1] : null;
  Logger.log('Último mes detectado: ' + ultimoMes);
  Logger.log('Total meses: ' + meses.join(', '));
  var resultado = correrVerificacion(ss, ultimoMes);
  Logger.log('=== Resultado: ' + JSON.stringify(resultado) + ' ===');
  if (resultado.ok) {
    Logger.log('✓ Verificación completada: ' + resultado.deptos + ' deptos, ' + resultado.meses + ' meses');
  } else {
    Logger.log('✗ Error: ' + resultado.error);
  }
}
// ═══════════════════════════════════════════════
//  testMultas
// ═══════════════════════════════════════════════
function testMultas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var meses = getMesesActivos(ss, null);
  var DEPTO_TEST = '101';
  var deptNorm = normDept(DEPTO_TEST);
  Logger.log('=== testMultas para depto ' + DEPTO_TEST + ' ===');
  var tarifasSh = ss.getSheetByName('Tarifas');
  var tarifasRows = tarifasSh ? tarifasSh.getDataRange().getValues() : [];
  for (var m = 0; m < meses.length; m++) {
    var sheetName = meses[m];
    var mesParts  = sheetName.split(' ');
    var multaMonto = getTarifaVigente(tarifasRows, 'Multa', mesParts[0], parseInt(mesParts[1]));
    var sh = ss.getSheetByName(sheetName);
    if (!sh) continue;
    var rows = sh.getDataRange().getValues();
    for (var r = 1; r < rows.length; r++) {
      var rowDept = deptDeRow(rows[r]);
      if (rowDept !== deptNorm) continue;
      var monto = Number(rows[r][3]||0);
      if (monto <= 0) continue;
      var colG   = String(rows[r][6]||'').trim();
      var colI   = String(rows[r][8]||'').trim();
      var fInfo  = fechaRowStr(rows[r][0]);
      var dia    = fInfo ? fInfo.dia : '?';
      var esCM   = colG.toUpperCase().indexOf('CM-') === 0;
      var multa  = fInfo && fInfo.dia > 10 && multaMonto > 0 && esCM;
      Logger.log(sheetName+' | f'+(r+1)+' | G="'+colG+'" | I="'+colI+'" | dia='+dia+' | $'+monto+' | CM='+esCM+' | multa='+multa);
    }
  }
  Logger.log('=== FIN testMultas ===');
}
// ═══════════════════════════════════════════════
//  testDeptMatch
// ═══════════════════════════════════════════════
function testDeptMatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var meses = getMesesActivos(ss, null);
  Logger.log('=== testDeptMatch === Meses: ' + meses.length + ' → ' + JSON.stringify(meses));
  var saldosSh = ss.getSheetByName('Saldos');
  if (saldosSh) {
    var sRows = saldosSh.getDataRange().getValues();
    Logger.log('\n--- Saldos (primeros 5) ---');
    for (var i = 1; i < Math.min(6, sRows.length); i++) {
      var raw = sRows[i][0];
      Logger.log('  f'+(i+1)+': "'+raw+'" → "'+normDept(String(raw||''))+'"');
    }
  }
  if (meses.length > 0) {
    var sheetName = meses[meses.length - 1];
    var sh = ss.getSheetByName(sheetName);
    if (!sh) { Logger.log('ERROR: hoja "'+sheetName+'" no encontrada'); return; }
    var rows = sh.getDataRange().getValues();
    Logger.log('\n--- Hoja: "' + sheetName + '" ('+rows.length+' filas) ---');
    for (var r = 1; r < Math.min(9, rows.length); r++) {
      Logger.log('  f'+(r+1)+' | A='+JSON.stringify(rows[r][0])+' | D='+rows[r][3]+' | G='+JSON.stringify(rows[r][6])+' | I='+JSON.stringify(rows[r][8])+' → "'+deptDeRow(rows[r])+'"');
    }
    var encontrados = 0;
    for (var r = 1; r < rows.length; r++) {
      if (deptDeRow(rows[r]) === '101') {
        encontrados++;
        Logger.log('  ✓ f'+(r+1)+': $'+rows[r][3]+' I='+JSON.stringify(rows[r][8])+' G='+JSON.stringify(rows[r][6]));
      }
    }
    if (encontrados === 0) Logger.log('  ✗ Ninguna fila matcheó "101"');
  }
}
// ═══════════════════════════════════════════════
//  debugRecibosExistentes — DIAGNÓSTICO
//  Muestra en los logs exactamente qué hay en la
//  hoja Recibos para entender por qué duplica.
// ═══════════════════════════════════════════════
function debugRecibosExistentes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rs = ss.getSheetByName('Recibos');
  if (!rs) { Logger.log('Hoja Recibos no existe'); return; }
  var rows = rs.getDataRange().getValues();
  Logger.log('Total filas en Recibos (incl. header): ' + rows.length);
  Logger.log('HEADER: ' + JSON.stringify(rows[0]));
  for (var i = 1; i < rows.length; i++) {
    Logger.log('Fila ' + (i+1) + ' | dept=[' + rows[i][1] + '] mes=[' + rows[i][3] +
      '] monto=[' + rows[i][5] + '] tipo=' + typeof rows[i][5] +
      ' estado=[' + rows[i][7] + '] folio=[' + rows[i][0] + ']');
  }
}
// ═══════════════════════════════════════════════
//  generarRecibosHistoricos — FUNCIÓN TEMPORAL
//  Genera recibos de Dic 2025 a la fecha actual.
//  Ejecutar manualmente desde el editor GAS.
//  NO envía correos (ENVIAR_CORREOS = false).
// ═══════════════════════════════════════════════
function generarRecibosHistoricos() {
  var MESES_HISTORICOS = [
    'Diciembre 2025',
    'Enero 2026',
    'Febrero 2026',
    'Marzo 2026'
  ];

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var resumen = [];

  for (var m = 0; m < MESES_HISTORICOS.length; m++) {
    var mes   = MESES_HISTORICOS[m];
    var sheet = ss.getSheetByName(mes);

    if (!sheet) {
      var msg = '⚠️  Hoja no encontrada: ' + mes + ' — se omite';
      Logger.log(msg);
      resumen.push(msg);
      continue;
    }

    var rows = sheet.getDataRange().getValues();
    var generados = 0, omitidos = 0, errores = 0;

    // Leer cache de Recibos UNA vez por mes (se recarga después de cada generación)
    var rs = ss.getSheetByName('Recibos');
    var recibosCache = rs ? rs.getDataRange().getValues() : [];

    for (var i = 1; i < rows.length; i++) {
      var monto  = rows[i][3];
      var idConc = String(rows[i][6] || '').trim();

      if (!monto || isNaN(monto) || Number(monto) <= 0) continue;
      if (!idConc) continue;

      var info = getConceptoYDepto(idConc);
      if (!info) continue;

      var dept     = info.dept;
      var concepto = info.concepto + ' · Depto ' + dept;

      // Duplicado: mismo depto + mes + monto activo
      // Usamos toFixed(2) para evitar problemas de punto flotante en monto
      var montoCheck = Number(monto).toFixed(2);
      var isDup = false;
      for (var j = 1; j < recibosCache.length; j++) {
        var rDept   = String(recibosCache[j][1]).trim().toUpperCase();
        var rMes    = periodoAMes(recibosCache[j][3]); // convierte Date → "Mes Año"
        var rMonto  = Number(recibosCache[j][5]).toFixed(2);
        var rEstado = String(recibosCache[j][7]).trim().toLowerCase();
        if (rDept === dept && rMes === mes && rMonto === montoCheck) {
          isDup = true;
          Logger.log('  SKIP dup (incl. cancelado): ' + dept + ' ' + mes + ' $' + montoCheck + ' → ' + recibosCache[j][0]);
          break;
        }
      }
      if (!isDup) {
        Logger.log('  CHECK: dept=[' + dept + '] mes=[' + mes + '] monto=[' + montoCheck +
          '] — cache tiene ' + (recibosCache.length - 1) + ' registros, NO encontró dup');
      }
      if (isDup) { omitidos++; continue; }

      var nombre = getNombrePropietario(dept);

      // Formatear fecha de pago
      var fechaStr = '';
      try {
        var fecha = rows[i][0];
        if (fecha instanceof Date) {
          fechaStr = Utilities.formatDate(fecha, 'America/Mexico_City', 'dd/MM/yyyy');
        } else if (typeof fecha === 'number') {
          fechaStr = Utilities.formatDate(
            new Date((fecha - 25569) * 86400 * 1000), 'America/Mexico_City', 'dd/MM/yyyy');
        } else {
          fechaStr = String(fecha);
        }
      } catch(ex) { fechaStr = String(rows[i][0]); }

      var result = generarRecibo(dept, nombre, mes, fechaStr, monto, concepto);
      if (result.ok) {
        generados++;
        Logger.log('✓  ' + mes + ' · Depto ' + dept + ' → ' + result.folio);
        // Recargar cache para que el siguiente recibo ya vea este registro
        rs = ss.getSheetByName('Recibos');
        recibosCache = rs ? rs.getDataRange().getValues() : [];
      } else {
        errores++;
        Logger.log('✗  ' + mes + ' · Depto ' + dept + ' · ERROR: ' + result.error);
        // Si es error de cuota de Google Docs, no tiene caso seguir intentando
        if (result.error && result.error.indexOf('too many times') !== -1) {
          Logger.log('🚫 CUOTA AGOTADA — Se detiene. La cuota de Google Docs se resetea a medianoche.');
          Logger.log('   Los recibos ya generados están guardados. Vuelve a correr mañana para continuar.');
          resumen.push(mes + ': ' + generados + ' generados, ' + omitidos + ' omitidos (DETENIDO por cuota)');
          Logger.log('═══════════════════════════════════════');
          Logger.log('RESUMEN PARCIAL:');
          for (var r = 0; r < resumen.length; r++) Logger.log(resumen[r]);
          Logger.log('═══════════════════════════════════════');
          return;
        }
      }

      // Pequeña pausa para no saturar quickchart.io (QR)
      Utilities.sleep(300);
    }

    var linea = mes + ': ' + generados + ' generados, ' + omitidos + ' omitidos, ' + errores + ' errores';
    Logger.log('📋 ' + linea);
    resumen.push(linea);
  }

  Logger.log('═══════════════════════════════════════');
  Logger.log('RESUMEN FINAL — Recibos Históricos');
  Logger.log('═══════════════════════════════════════');
  for (var r = 0; r < resumen.length; r++) Logger.log(resumen[r]);
  Logger.log('═══════════════════════════════════════');
}

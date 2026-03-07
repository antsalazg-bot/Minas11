// ═══════════════════════════════════════════════════════════════
//  shared.js — Utilidades compartidas entre admin.html y portal.html
//
//  ⚠️  REGLA: cualquier fix de lógica va AQUÍ primero.
//     Ambos archivos incluyen este script y usan las funciones.
//     Equivalente JS del backend Codigo_completo.gs
//
//  Incluir ANTES del bloque <script> principal de cada archivo:
//    <script src="shared.js"></script>
// ═══════════════════════════════════════════════════════════════

// ── FORMATEO FECHA ────────────────────────────────────────────
// Acepta: Date obj | número serial Excel | ISO string | DD/MM/YYYY
// Retorna siempre: "DD/MM/YYYY"  (o '' si no hay valor)
function fmtFecha(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date && !isNaN(v)) {
    const d = String(v.getDate()).padStart(2,'0');
    const m = String(v.getMonth()+1).padStart(2,'0');
    return `${d}/${m}/${v.getFullYear()}`;
  }
  if (typeof v === 'number' && v > 1) {
    // Serial Excel/Sheets (días desde 1899-12-30)
    const dt = new Date(Math.round((v - 25569) * 86400000));
    return `${String(dt.getUTCDate()).padStart(2,'0')}/${String(dt.getUTCMonth()+1).padStart(2,'0')}/${dt.getUTCFullYear()}`;
  }
  const s = String(v).trim();
  // ISO: "2025-02-01" o "2025-02-01T00:00:00..."
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const parts = s.split('T')[0].split('-');
    if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  // Ya es DD/MM/YYYY u otro formato — devolver tal cual
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return s;
  // Último intento: new Date()
  const d = new Date(s);
  if (!isNaN(d)) {
    return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
  }
  return s;
}

// ── DETECCIÓN DE DEPARTAMENTO ─────────────────────────────────
// Determina el departamento de una fila de hoja mensual.
// ⚠️  Lógica idéntica a deptDeRow() en Codigo_completo.gs —
//     cualquier cambio debe reflejarse en AMBOS lados.
//
// Prioridad:
//   1. Col I (índice 8)   — DEPTO directo:  "302", "PH4"
//   2. Col G (idConc)     — ID Concepto:    CM-302, CE-CIS-305
//                           Solo si el último segmento tiene dígito
//   3. Col C (nombre)     — Concepto libre: "CUOTA EXTRA CISTERNAS 305" → "305"
//
// Retorna: { dept, mCM, deptColI }
//   dept     — string del depto (uppercase) o null
//   mCM      — true si el pago es cuota de mantenimiento (CM-xxx)
//   deptColI — valor de col I (puede ser null)
function detectDeptFromRow(row, nombre, idConc) {
  // 1) Col I — directo
  const deptColI = row[8] ? String(row[8]).trim().toUpperCase() : null;

  // 2) Col G — CM-302, CE-CIS-305, CV-VIG-101
  //    El último segmento DEBE contener un dígito (evita "CIS" como depto)
  const mCM     = String(idConc || '').match(/^CM-(.+)$/);
  const mCEfull = String(idConc || '').match(/^(?:CE|CV)-[A-Z]+-(\w*\d\w*)$/i);
  const deptFromG = mCM
    ? mCM[1].toUpperCase()
    : (mCEfull ? mCEfull[1].toUpperCase() : null);

  // 3) Col C — número de depto al final del texto (3 dígitos o PH+dígito)
  const mConc = (!deptColI && !deptFromG)
    ? ((nombre || '').match(/\b(\d{3}|PH\d)\s*$/i) || null)
    : null;
  const deptFromConc = mConc ? mConc[1].toUpperCase() : null;

  const dept = deptColI || deptFromG || deptFromConc || null;
  return { dept, mCM: !!mCM, deptColI };
}

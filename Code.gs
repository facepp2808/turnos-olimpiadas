// Backend de "Turnos Olimpiadas" para Google Apps Script.
// Pegar en: Google Sheet > Extensiones > Apps Script. Ver LEEME.md.

const HOJA = 'Cupos';
const CLAVE_ADMIN = 'CPSM'; // cámbiala antes de implementar
const COLS = ['id', 'pestana', 'descripcion', 'fecha', 'horario', 'inicio', 'n', 'nombre', 'curso'];

// Ejecutar UNA vez desde el editor para crear la hoja con los 88 cupos.
function cargarDatos() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
  sh.clear();
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  sh.getRange(2, 1, DATOS.length, COLS.length).setNumberFormat('@').setValues(DATOS);
  sh.getRange(1, 1, 1, COLS.length).setNumberFormat('@');
  sh.setFrozenRows(1);
}

function hoja_() { return SpreadsheetApp.getActive().getSheetByName(HOJA); }

function leer_() {
  const v = hoja_().getDataRange().getValues();
  const filas = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    const o = { _fila: i + 1 };
    COLS.forEach((c, j) => { o[c] = String(v[i][j]); });
    o.n = Number(o.n) || 0;
    filas.push(o);
  }
  return filas;
}

function publico_(r) {
  return { id: r.id, pestana: r.pestana, descripcion: r.descripcion, fecha: r.fecha,
           horario: r.horario, inicio: r.inicio, n: r.n, ocupado: !!(r.nombre || r.curso) };
}

function limpiar_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, 80).replace(/^[=+\-@]+/, '');
}

function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return salida_({ ok: true, filas: leer_().map(publico_) });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return salida_({ ok: false, error: 'solicitud_invalida' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.action === 'guardar') return salida_(guardar_(b.filas));
    if (b.clave !== CLAVE_ADMIN) return salida_({ ok: false, error: 'clave_incorrecta' });
    if (b.action === 'admin_listar') return salida_({ ok: true, filas: leer_() });
    if (b.action === 'admin_actualizar') return salida_(actualizar_(b.fila));
    if (b.action === 'admin_agregar') return salida_(agregar_(b.fila));
    if (b.action === 'admin_eliminar') return salida_(eliminar_(b.id));
    return salida_({ ok: false, error: 'accion_desconocida' });
  } finally {
    lock.releaseLock();
  }
}

// Guarda nombre y curso solo en cupos que siguen vacíos.
function guardar_(filas) {
  if (!Array.isArray(filas) || !filas.length || filas.length > 60) return { ok: false, error: 'solicitud_invalida' };
  const todas = leer_();
  const sh = hoja_();
  const resultados = {};
  filas.forEach(f => {
    const r = todas.find(x => x.id === String(f.id));
    const nombre = limpiar_(f.nombre), curso = limpiar_(f.curso);
    if (!r || !nombre || !curso) { resultados[f.id] = 'invalido'; return; }
    if (r.nombre || r.curso) { resultados[f.id] = 'ocupado'; return; }
    sh.getRange(r._fila, COLS.indexOf('nombre') + 1, 1, 2).setNumberFormat('@').setValues([[nombre, curso]]);
    r.nombre = nombre; r.curso = curso;
    resultados[f.id] = 'ok';
  });
  return { ok: true, resultados: resultados };
}

function valores_(f) {
  return COLS.map(c => c === 'n' ? Number(f[c]) || 0 : String(f[c] == null ? '' : f[c]).slice(0, 120));
}

function actualizar_(f) {
  const r = leer_().find(x => x.id === String(f.id));
  if (!r) return { ok: false, error: 'no_existe' };
  hoja_().getRange(r._fila, 1, 1, COLS.length).setNumberFormat('@').setValues([valores_(f)]);
  return { ok: true };
}

function agregar_(f) {
  f.id = 'X' + Date.now();
  const sh = hoja_();
  sh.getRange(sh.getLastRow() + 1, 1, 1, COLS.length).setNumberFormat('@').setValues([valores_(f)]);
  return { ok: true, id: f.id };
}

function eliminar_(id) {
  const r = leer_().find(x => x.id === String(id));
  if (!r) return { ok: false, error: 'no_existe' };
  hoja_().deleteRow(r._fila);
  return { ok: true };
}

const DATOS = [
  ["A01", "apoderados", "Apoyo ajedrez", "2026-10-24", "8:30 a 11:00", "08:30", 1, "", ""],
  ["A02", "apoderados", "Apoyo ajedrez", "2026-10-24", "8:30 a 11:00", "08:30", 2, "", ""],
  ["A03", "apoderados", "Apoyo ajedrez", "2026-10-24", "11:00 a 13:30", "11:00", 3, "", ""],
  ["A04", "apoderados", "Apoyo ajedrez", "2026-10-24", "11:00 a 13:30", "11:00", 4, "", ""],
  ["A05", "apoderados", "Apoyo gimnasio futbolito", "2026-10-24", "8:30 a 11:00", "08:30", 5, "", ""],
  ["A06", "apoderados", "Apoyo gimnasio futbolito", "2026-10-24", "8:30 a 11:00", "08:30", 6, "", ""],
  ["A07", "apoderados", "Apoyo gimnasio futbolito", "2026-10-24", "11:00 a 13:30", "11:00", 7, "", ""],
  ["A08", "apoderados", "Apoyo gimnasio futbolito", "2026-10-24", "11:00 a 13:30", "11:00", 8, "", ""],
  ["A09", "apoderados", "Apoyo ping-pong (salón)", "2026-10-24", "8:30 a 11:00", "08:30", 9, "", ""],
  ["A10", "apoderados", "Apoyo ping-pong (salón)", "2026-10-24", "8:30 a 11:00", "08:30", 10, "", ""],
  ["A11", "apoderados", "Apoyo ping-pong (salón)", "2026-10-24", "11:00 a 13:30", "11:00", 11, "", ""],
  ["A12", "apoderados", "Apoyo ping-pong (salón)", "2026-10-24", "11:00 a 13:30", "11:00", 12, "", ""],
  ["A13", "apoderados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "8:30 a 11:00", "08:30", 13, "", ""],
  ["A14", "apoderados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "8:30 a 11:00", "08:30", 14, "", ""],
  ["A15", "apoderados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "11:00 a 13:30", "11:00", 15, "", ""],
  ["A16", "apoderados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "11:00 a 13:30", "11:00", 16, "", ""],
  ["A17", "apoderados", "Apoyo voleibol (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 17, "", ""],
  ["A18", "apoderados", "Apoyo voleibol (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 18, "", ""],
  ["A19", "apoderados", "Apoyo yincana (patio central)", "2026-10-24", "8:30 a 11:00", "08:30", 19, "", ""],
  ["A20", "apoderados", "Apoyo yincana (patio central)", "2026-10-24", "8:30 a 11:00", "08:30", 20, "", ""],
  ["A21", "apoderados", "Apoyo yincana (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 21, "", ""],
  ["A22", "apoderados", "Apoyo yincana (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 22, "", ""],
  ["A23", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 23, "", ""],
  ["A24", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 24, "", ""],
  ["A25", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 25, "", ""],
  ["A26", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 26, "", ""],
  ["A27", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 27, "", ""],
  ["A28", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 28, "", ""],
  ["A29", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 29, "", ""],
  ["A30", "apoderados", "Armado", "2026-10-23", "18:00", "18:00", 30, "", ""],
  ["A31", "apoderados", "Aseo", "2026-10-24", "8:30 a 11:00", "08:30", 31, "", ""],
  ["A32", "apoderados", "Aseo", "2026-10-24", "8:30 a 11:00", "08:30", 32, "", ""],
  ["A33", "apoderados", "Aseo", "2026-10-24", "8:30 a 11:00", "08:30", 33, "", ""],
  ["A34", "apoderados", "Aseo y desarme", "2026-10-24", "11:00 a 13:30", "11:00", 34, "", ""],
  ["A35", "apoderados", "Aseo y desarme", "2026-10-24", "11:00 a 13:30", "11:00", 35, "", ""],
  ["A36", "apoderados", "Aseo y desarme", "2026-10-24", "11:00 a 13:30", "11:00", 36, "", ""],
  ["A37", "apoderados", "Aseo y desarme", "2026-10-24", "11:00 a 13:30", "11:00", 37, "", ""],
  ["A38", "apoderados", "Aseo y desarme", "2026-10-24", "11:00 a 13:30", "11:00", 38, "", ""],
  ["A39", "apoderados", "Control acceso Don Bosco", "2026-10-24", "8:30 a 11:00", "08:30", 39, "", ""],
  ["A40", "apoderados", "Control acceso Don Bosco", "2026-10-24", "8:30 a 11:00", "08:30", 40, "", ""],
  ["A41", "apoderados", "Control acceso Don Bosco", "2026-10-24", "8:30 a 11:00", "08:30", 41, "", ""],
  ["A42", "apoderados", "Control acceso Don Bosco", "2026-10-24", "11:00 a 13:30", "11:00", 42, "", ""],
  ["A43", "apoderados", "Control acceso Don Bosco", "2026-10-24", "11:00 a 13:30", "11:00", 43, "", ""],
  ["A44", "apoderados", "Control acceso Don Bosco", "2026-10-24", "11:00 a 13:30", "11:00", 44, "", ""],
  ["A45", "apoderados", "Seguridad espacios comunes", "2026-10-24", "8:30 a 11:00", "08:30", 45, "", ""],
  ["A46", "apoderados", "Seguridad espacios comunes", "2026-10-24", "8:30 a 11:00", "08:30", 46, "", ""],
  ["A47", "apoderados", "Seguridad espacios comunes", "2026-10-24", "8:30 a 11:00", "08:30", 47, "", ""],
  ["A48", "apoderados", "Seguridad espacios comunes", "2026-10-24", "8:30 a 11:00", "08:30", 48, "", ""],
  ["A49", "apoderados", "Seguridad espacios comunes", "2026-10-24", "8:30 a 11:00", "08:30", 49, "", ""],
  ["A50", "apoderados", "Seguridad espacios comunes", "2026-10-24", "11:00 a 13:30", "11:00", 50, "", ""],
  ["A51", "apoderados", "Seguridad espacios comunes", "2026-10-24", "11:00 a 13:30", "11:00", 51, "", ""],
  ["A52", "apoderados", "Seguridad espacios comunes", "2026-10-24", "11:00 a 13:30", "11:00", 52, "", ""],
  ["A53", "apoderados", "Seguridad espacios comunes", "2026-10-24", "11:00 a 13:30", "11:00", 53, "", ""],
  ["A54", "apoderados", "Seguridad espacios comunes", "2026-10-24", "11:00 a 13:30", "11:00", 54, "", ""],
  ["A55", "apoderados", "Seguridad estacionamiento", "2026-10-24", "8:30 a 11:00", "08:30", 55, "", ""],
  ["A56", "apoderados", "Seguridad estacionamiento", "2026-10-24", "8:30 a 11:00", "08:30", 56, "", ""],
  ["A57", "apoderados", "Seguridad estacionamiento", "2026-10-24", "11:00 a 13:30", "11:00", 57, "", ""],
  ["A58", "apoderados", "Seguridad estacionamiento", "2026-10-24", "11:00 a 13:30", "11:00", 58, "", ""],
  ["A59", "apoderados", "Seguridad y apoyo escenario", "2026-10-24", "8:30 a 11:00", "08:30", 59, "", ""],
  ["A60", "apoderados", "Seguridad y apoyo escenario", "2026-10-24", "11:00 a 13:30", "11:00", 60, "", ""],
  ["D01", "delegados", "Apoyo ajedrez", "2026-10-24", "8:30 a 11:00", "08:30", 1, "", ""],
  ["D02", "delegados", "Apoyo ajedrez", "2026-10-24", "11:00 a 13:30", "11:00", 2, "", ""],
  ["D03", "delegados", "Apoyo anfitriones acceso", "2026-10-24", "8:30 a 11:00", "08:30", 3, "", ""],
  ["D04", "delegados", "Apoyo anfitriones acceso", "2026-10-24", "11:00 a 13:30", "11:00", 4, "", ""],
  ["D05", "delegados", "Apoyo gimnasio futbolito", "2026-10-24", "8:30 a 11:00", "08:30", 5, "", ""],
  ["D06", "delegados", "Apoyo gimnasio futbolito", "2026-10-24", "8:30 a 11:00", "08:30", 6, "", ""],
  ["D07", "delegados", "Apoyo gimnasio futbolito", "2026-10-24", "8:30 a 11:00", "08:30", 7, "", ""],
  ["D08", "delegados", "Apoyo gimnasio futbolito", "2026-10-24", "11:00 a 13:30", "11:00", 8, "", ""],
  ["D09", "delegados", "Apoyo gimnasio futbolito", "2026-10-24", "11:00 a 13:30", "11:00", 9, "", ""],
  ["D10", "delegados", "Apoyo gimnasio futbolito", "2026-10-24", "11:00 a 13:30", "11:00", 10, "", ""],
  ["D11", "delegados", "Apoyo invitados", "2026-10-24", "8:30 a 11:00", "08:30", 11, "", ""],
  ["D12", "delegados", "Apoyo invitados", "2026-10-24", "8:30 a 11:00", "08:30", 12, "", ""],
  ["D13", "delegados", "Apoyo invitados", "2026-10-24", "11:00 a 13:30", "11:00", 13, "", ""],
  ["D14", "delegados", "Apoyo invitados", "2026-10-24", "11:00 a 13:30", "11:00", 14, "", ""],
  ["D15", "delegados", "Apoyo ping-pong (salón)", "2026-10-24", "8:30 a 11:00", "08:30", 15, "", ""],
  ["D16", "delegados", "Apoyo ping-pong (salón)", "2026-10-24", "8:30 a 11:00", "08:30", 16, "", ""],
  ["D17", "delegados", "Apoyo ping-pong (salón)", "2026-10-24", "11:00 a 13:30", "11:00", 17, "", ""],
  ["D18", "delegados", "Apoyo ping-pong (salón)", "2026-10-24", "11:00 a 13:30", "11:00", 18, "", ""],
  ["D19", "delegados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "8:30 a 11:00", "08:30", 19, "", ""],
  ["D20", "delegados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "8:30 a 11:00", "08:30", 20, "", ""],
  ["D21", "delegados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "11:00 a 13:30", "11:00", 21, "", ""],
  ["D22", "delegados", "Apoyo taca-taca (patio de colores)", "2026-10-24", "11:00 a 13:30", "11:00", 22, "", ""],
  ["D23", "delegados", "Apoyo voleibol (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 23, "", ""],
  ["D24", "delegados", "Apoyo voleibol (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 24, "", ""],
  ["D25", "delegados", "Apoyo yincana (patio central)", "2026-10-24", "8:30 a 11:00", "08:30", 25, "", ""],
  ["D26", "delegados", "Apoyo yincana (patio central)", "2026-10-24", "8:30 a 11:00", "08:30", 26, "", ""],
  ["D27", "delegados", "Apoyo yincana (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 27, "", ""],
  ["D28", "delegados", "Apoyo yincana (patio central)", "2026-10-24", "11:00 a 13:30", "11:00", 28, "", ""]
];

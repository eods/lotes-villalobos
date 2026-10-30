// API de propietarios para Lotes Villalobos / Magallanes (Google Apps Script, vinculado a la planilla).
// La contraseña NO va acá: se carga en Configuración del proyecto > Propiedades del script > CLAVE.
const HOJA = 'Lotes';
const BARRIOS = { V: { nombre: 'Villalobos', max: 454 }, M: { nombre: 'Magallanes', max: 406 } };
const MAX_TEXTO = 30, MAX_FALLOS = 10;

// Columnas: A Lote | B Propietario | C Actualizado | D Barrio (vacío = Villalobos, filas anteriores)
function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA);
  if (!h) {
    h = ss.insertSheet(HOJA);
    h.appendRow(['Lote', 'Propietario', 'Actualizado', 'Barrio']);
    h.setFrozenRows(1);
  } else if (h.getRange(1, 4).getValue() === '') {
    h.getRange(1, 4).setValue('Barrio');
  }
  return h;
}

function codigoBarrio_(texto) {
  return String(texto || '').trim().toLowerCase().startsWith('m') ? 'M' : 'V';
}

function leer_() {
  const datos = hoja_().getDataRange().getValues();
  const r = { V: {}, M: {} };
  for (let i = 1; i < datos.length; i++) {
    const n = parseInt(datos[i][0], 10), p = String(datos[i][1] || '').trim();
    if (n && p) r[codigoBarrio_(datos[i][3])][n] = p;
  }
  return r;
}

// Lectura pública: { ok, v: 2, lotes: { V: { "12": "Pérez" }, M: { ... } } }
function doGet() {
  return json_({ ok: true, v: 2, lotes: leer_() });
}

// Escritura con contraseña: { accion: "verificar" | "guardar", clave, barrio: "V" | "M", lote, propietario }
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Pedido inválido' }); }

  const cache = CacheService.getScriptCache();
  const fallos = +(cache.get('fallos') || 0);
  if (fallos >= MAX_FALLOS) return json_({ ok: false, error: 'Demasiados intentos fallidos. Probá de nuevo en 10 minutos.' });
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!clave) return json_({ ok: false, error: 'Falta configurar la propiedad CLAVE en el script.' });
  if (req.clave !== clave) {
    cache.put('fallos', String(fallos + 1), 600);
    return json_({ ok: false, error: 'Contraseña incorrecta' });
  }

  if (req.accion === 'verificar') return json_({ ok: true, v: 2 });

  if (req.accion === 'guardar') {
    const b = req.barrio === 'M' ? 'M' : 'V';
    const n = parseInt(req.lote, 10);
    if (!(n >= 1 && n <= BARRIOS[b].max)) return json_({ ok: false, error: 'Lote inválido' });
    let p = String(req.propietario || '').replace(/\s+/g, ' ').trim().slice(0, MAX_TEXTO);
    if (/^[=+\-@]/.test(p)) p = "'" + p; // que la planilla no lo tome como fórmula

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const h = hoja_();
      const filas = h.getRange(1, 1, h.getLastRow(), 4).getValues();
      let fila = 0;
      for (let i = 1; i < filas.length; i++) {
        if (parseInt(filas[i][0], 10) === n && codigoBarrio_(filas[i][3]) === b) { fila = i + 1; break; }
      }
      if (fila && p) h.getRange(fila, 2, 1, 3).setValues([[p, new Date(), BARRIOS[b].nombre]]);
      else if (fila) h.deleteRow(fila);
      else if (p) h.appendRow([n, p, new Date(), BARRIOS[b].nombre]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, v: 2, lotes: leer_() });
  }

  return json_({ ok: false, error: 'Acción desconocida' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

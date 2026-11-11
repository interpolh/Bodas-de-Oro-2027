/** BODAS DE ORO 2027 — Reemplazar el contenido de Código.gs con este código.
 * Debe ser un proyecto creado desde la hoja de cálculo: Extensiones > Apps Script.
 * Fecha límite: 20 diciembre 2026, 23:59:59, America/Mexico_City.
 */
const HOJA_CONFIRMACIONES = 'Confirmaciones Web';
const ZONA_HORARIA = 'America/Mexico_City';
const ULTIMO_DIA = '2026-12-20';
const MAXIMO = 9;
const ENCABEZADOS = ['Fecha', 'Nombre', 'Adultos', 'Niños', 'Total', 'Asistencia', 'Invitado de'];

function doPost(e) {
  const bloqueo = LockService.getScriptLock();
  let adquirido = false;
  try {
    const datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const nombre = String(datos.nombre || '').trim().replace(/[\r\n\t]+/g, ' ');
    const asistencia = String(datos.asistencia || '').trim();
    const invitadoDe = String(datos.invitadoDe || '').trim();
    const adultos = Number(datos.adultos);
    const ninos = Number(datos.ninos);
    const total = adultos + ninos;
    if (nombre.length < 3 || nombre.length > 100) throw new Error('Escribe un nombre válido.');
    if (!['Sí','No'].includes(asistencia)) throw new Error('Asistencia inválida.');
    if (!['Bodas de Oro','Nuestra Boda','Ambas parejas'].includes(invitadoDe)) throw new Error('Selecciona de quién eres invitado.');
    if (!Number.isInteger(adultos) || !Number.isInteger(ninos) || adultos < 0 || ninos < 0 || total > MAXIMO) throw new Error('Cantidad inválida. Máximo 9 personas.');
    if (asistencia === 'Sí' && adultos < 1) throw new Error('Debe asistir al menos un adulto.');
    if (asistencia === 'No' && total !== 0) throw new Error('Si no asistirás, el total debe ser 0.');
    adquirido = bloqueo.tryLock(15000);
    if (!adquirido) throw new Error('El sistema está ocupado. Intenta más tarde.');
    // Comprobación EN EL SERVIDOR, bajo bloqueo; la interfaz no puede eludirla.
    const hoy = Utilities.formatDate(new Date(), ZONA_HORARIA, 'yyyy-MM-dd');
    if (hoy > ULTIMO_DIA) throw new Error('El plazo de confirmación terminó el 20 de diciembre de 2026.');
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    if (!libro) throw new Error('Este script debe estar vinculado a una hoja de cálculo.');
    let hoja = libro.getSheetByName(HOJA_CONFIRMACIONES);
    if (!hoja) hoja = libro.insertSheet(HOJA_CONFIRMACIONES);
    // Conservar las columnas A:E y los registros anteriores de V18/V19.
    ENCABEZADOS.forEach((encabezado, i) => hoja.getRange(1, i + 1).setValue(encabezado));
    const clave = normalizar(nombre);
    const ultimo = hoja.getLastRow();
    if (ultimo >= 2) {
      const nombres = hoja.getRange(2, 2, ultimo - 1, 1).getDisplayValues();
      if (nombres.some(fila => normalizar(fila[0]) === clave)) {
        return responder({ok:false,codigo:'DUPLICADO',mensaje:'Este nombre ya fue registrado. Contacta a los organizadores para corregirlo.'});
      }
    }
    // Evitar que un nombre empiece como fórmula en Google Sheets.
    const nombreSeguro = "'" + nombre.replace(/^[=+\-@]/, '');
    hoja.appendRow([new Date(), nombreSeguro, adultos, ninos, total, asistencia, invitadoDe]);
    SpreadsheetApp.flush();
    return responder({ok:true,mensaje:'¡Gracias por confirmar tu asistencia!'});
  } catch (error) {
    return responder({ok:false,codigo:'ERROR',mensaje:String(error.message || error)});
  } finally {
    if (adquirido) bloqueo.releaseLock();
  }
}
function normalizar(valor) {
  return String(valor || '').replace(/^'/, '').normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
}
function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Confirmaciones Bodas de Oro 2027 — proyecto vinculado a Google Sheets. */
const HOJA_CONFIRMACIONES = 'Confirmaciones Web';
const ZONA_HORARIA = 'America/Mexico_City';
const ULTIMO_DIA = '2026-12-20';
const MAXIMO = 9;
const ENCABEZADOS = ['Fecha','Nombre','Adultos','Niños','Total','Asistencia','Invitado de'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const nombre = String(data.nombre || '').trim().replace(/[\r\n\t]+/g,' ');
    const asistencia = String(data.asistencia || '').trim();
    const invitadoDe = String(data.invitadoDe || '').trim();
    const adultos = Number(data.adultos), ninos = Number(data.ninos);
    const total = adultos + ninos;
    if (nombre.length < 3 || nombre.length > 100) throw Error('Escribe un nombre válido.');
    if (!['Sí','No'].includes(asistencia)) throw Error('Selecciona si asistirás.');
    if (!['Bodas de Oro','Nuestra Boda','Ambas parejas'].includes(invitadoDe)) throw Error('Selecciona de quién eres invitado.');
    if (!Number.isInteger(adultos) || !Number.isInteger(ninos) || adultos < 0 || ninos < 0 || total > MAXIMO) throw Error('Máximo 9 personas.');
    if (asistencia === 'Sí' && adultos < 1) throw Error('Debe asistir al menos un adulto.');
    if (asistencia === 'No' && total !== 0) throw Error('Si no asistirás, adultos y niños deben ser 0.');
    locked = lock.tryLock(15000);
    if (!locked) throw Error('Sistema ocupado; intenta más tarde.');
    const hoy = Utilities.formatDate(new Date(), ZONA_HORARIA, 'yyyy-MM-dd');
    if (hoy > ULTIMO_DIA) throw Error('Las confirmaciones cerraron el 20 de diciembre de 2026.');
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    if (!libro) throw Error('Vincula este proyecto a la hoja de cálculo.');
    let hoja = libro.getSheetByName(HOJA_CONFIRMACIONES);
    if (!hoja) hoja = libro.insertSheet(HOJA_CONFIRMACIONES);
    hoja.getRange(1,1,1,ENCABEZADOS.length).setValues([ENCABEZADOS]);
    const n = hoja.getLastRow();
    if (n > 1) {
      const previos = hoja.getRange(2,2,n-1,1).getDisplayValues();
      if (previos.some(fila => normalizar(fila[0]) === normalizar(nombre)))
        return responder({ok:false,codigo:'DUPLICADO',mensaje:'Este nombre ya está registrado.'});
    }
    // Prevenir inyección de fórmulas en las celdas.
    const seguro = "'" + nombre.replace(/^[=+\-@]/,'');
    hoja.appendRow([new Date(),seguro,adultos,ninos,total,asistencia,invitadoDe]);
    SpreadsheetApp.flush();
    return responder({ok:true,mensaje:'Confirmación registrada.'});
  } catch(err) {
    return responder({ok:false,codigo:'ERROR',mensaje:String(err.message || err)});
  } finally {
    if (locked) lock.releaseLock();
  }
}
function normalizar(x) {
  return String(x || '').replace(/^'/,'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase();
}
function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

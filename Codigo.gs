/**
 * Respaldo automático de Tierra Quin en Google Drive
 * Copia toda la información de Firebase a un archivo en la carpeta "Respaldos Tierra Quin".
 * El archivo se puede usar en el sistema con Ajustes → Restaurar desde archivo.
 */
const PROYECTO = 'tierra-quin';                 // ID del proyecto de Firebase
const CARPETA  = 'Respaldos Tierra Quin';       // carpeta en tu Google Drive
const GUARDAR  = 90;                            // cuántas copias se conservan (las más viejas van a la papelera)
const HORA     = 22;                            // hora del respaldo diario (hora de Colombia)
const COLECCIONES = ['cli','prod','ordenes','combos','gastos','config','compras','combosDia'];

/** Ejecuta esta función UNA vez: programa el respaldo diario y hace el primero. */
function instalar() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction()==='respaldar').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('respaldar').timeBased().everyDays(1).atHour(HORA).inTimezone('America/Bogota').create();
  respaldar();
}

/** Hace una copia ahora mismo. */
function respaldar() {
  const token = ScriptApp.getOAuthToken();
  const data = {};
  let total = 0;
  COLECCIONES.forEach(c => {
    data[c] = {};
    let pageToken = '';
    do {
      let url = `https://firestore.googleapis.com/v1/projects/${PROYECTO}/databases/(default)/documents/${c}?pageSize=300`;
      if (pageToken) url += '&pageToken=' + encodeURIComponent(pageToken);
      const r = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token, 'x-goog-user-project': PROYECTO },
        muteHttpExceptions: true
      });
      if (r.getResponseCode() !== 200) throw new Error(`No se pudo leer "${c}" (${r.getResponseCode()}): ${r.getContentText().slice(0, 300)}`);
      const j = JSON.parse(r.getContentText() || '{}');
      (j.documents || []).forEach(doc => {
        const id = decodeURIComponent(doc.name.split('/').pop());
        data[c][id] = campos(doc.fields || {});
        total++;
      });
      pageToken = j.nextPageToken || '';
    } while (pageToken);
  });

  const copia = { app: 'tierra-quin', version: 2, creado: new Date().toISOString(), origen: 'respaldo-drive', data: data };
  const carpeta = obtenerCarpeta();
  const nombre = 'tierra-quin-respaldo-' + Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd_HH-mm') + '.json';
  carpeta.createFile(nombre, JSON.stringify(copia), 'application/json');
  limpiar(carpeta);
  console.log(`Respaldo guardado: ${nombre} (${total} registros)`);
}

function obtenerCarpeta() {
  const it = DriveApp.getFoldersByName(CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA);
}

function limpiar(carpeta) {
  const archivos = [];
  const it = carpeta.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf('tierra-quin-respaldo-') === 0) archivos.push(f); }
  archivos.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  archivos.slice(GUARDAR).forEach(f => f.setTrashed(true));
}

/* Convierte el formato de Firestore a datos normales */
function campos(f) { const o = {}; Object.keys(f).forEach(k => o[k] = valor(f[k])); return o; }
function valor(v) {
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return Number(v.doubleValue);
  if ('booleanValue' in v) return v.booleanValue;
  if ('nullValue' in v) return null;
  if ('timestampValue' in v) return v.timestampValue;
  if ('mapValue' in v) return campos(v.mapValue.fields || {});
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(valor);
  return null;
}

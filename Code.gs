/**
 * Registro Scarti - Berloni Bagno
 * Backend API (JSON) per Google Apps Script.
 * Il frontend (index.html) è ospitato su GitHub Pages e chiama questa web app via fetch().
 *
 * Foglio "Base dati":  A Data | B Stabilimento | C Articolo | D Variante |
 *                      E Quantità | F Causale | G Operatore | H Immagini | I Registrato su Essentia
 * Foglio "Causali":    A Classificazione | B Causale | C Provenienza | D Sigla
 */

var SPREADSHEET_ID = '1KfuV1-iQcWAutqGJthjycV2aTDv1GFiT3ppYbWAE1i4';
var SHEET_DATI = 'Base dati';
var SHEET_CAUSALI = 'Causali';
var STABILIMENTI = ['BB1', 'BB3', 'Ipiemme', 'Zenobi'];
var DRIVE_FOLDER_NAME = 'Registro Scarti - Immagini';

// Lo script usa lo scope OAuth "drive.file": vede SOLO i file che crea lui,
// non il resto del Drive del proprietario. Per questo non si può cercare una
// cartella per nome (non la troverebbe): si ricorda l'ID della cartella foto
// nelle Script Properties dopo averla creata la prima volta.
function getOrCreateDriveFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOTO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* cartella rimossa: ne creo una nuova */ }
  }
  var folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  props.setProperty('FOTO_FOLDER_ID', folder.getId());
  return folder;
}

function uploadFotos_(fotos) {
  if (!fotos || !fotos.length) return [];
  var folder = getOrCreateDriveFolder_();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var urls = [];
  for (var i = 0; i < fotos.length; i++) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(fotos[i]), 'image/jpeg', timestamp + '_' + (i + 1) + '.jpg');
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls.push(file.getUrl());
    } catch (e) {
      console.warn('Errore upload foto ' + (i + 1), e);
    }
  }
  return urls;
}

/** Cella con un link cliccabile per foto ("Foto 1", "Foto 2", …) su righe separate. */
function buildFotoRichText_(urls) {
  var labels = [];
  for (var i = 0; i < urls.length; i++) labels.push('Foto ' + (i + 1));
  var text = labels.join('\n');
  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  var pos = 0;
  for (var j = 0; j < labels.length; j++) {
    builder.setLinkUrl(pos, pos + labels[j].length, urls[j]);
    pos += labels[j].length + 1; // +1 per il "\n"
  }
  return builder.build();
}

/* ============================ ROUTING API ============================ */

/** Letture: ?action=init  oppure  ?action=kpi&start=<ms>&end=<ms> */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'init') return json_(getInitData());
    if (action === 'kpi') {
      return json_(getKpiPareto(Number(e.parameter.start), Number(e.parameter.end)));
    }
    return json_({ error: 'Azione non valida' });
  } catch (err) {
    return json_({ error: (err && err.message) || String(err) });
  }
}

/** Scritture: body JSON { action: 'registra', data: {...} } */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'registra') return json_(registraScarto(body.data || {}));
    return json_({ error: 'Azione non valida' });
  } catch (err) {
    return json_({ error: (err && err.message) || String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/* ============================ LETTURE ============================ */

/** Dati iniziali per popolare i menu a tendina. */
function getInitData() {
  return {
    stabilimenti: STABILIMENTI,
    causali: getCausali_(),
    operatori: getOperatori_(),
    articoli: getArticoli_()
  };
}

/** Valori della colonna B "Causale" del foglio "Causali" (ordine del foglio, senza duplicati). */
function getCausali_() {
  var sheet = ss_().getSheetByName(SHEET_CAUSALI);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 2, last - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = (values[i][0] || '').toString().trim();
    if (v && out.indexOf(v) === -1) out.push(v);
  }
  return out;
}

/** Valori distinti di una colonna di "Base dati", ordinati alfabeticamente (italiano). */
function getDistinctSorted_(column) {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, column, last - 1, 1).getValues();
  var seen = {};
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = (values[i][0] || '').toString().trim();
    var key = v.toLowerCase();
    if (v && !seen[key]) { seen[key] = true; out.push(v); }
  }
  out.sort(function (a, b) { return a.localeCompare(b, 'it', { sensitivity: 'base' }); });
  return out;
}

/** Operatori già presenti nel file (colonna G di "Base dati"), distinti e ordinati alfabeticamente. */
function getOperatori_() { return getDistinctSorted_(7); }

/** Articoli già presenti nel file (colonna C di "Base dati"), distinti e ordinati alfabeticamente. */
function getArticoli_() { return getDistinctSorted_(3); }

/** Mappa Causale -> Sigla (per le etichette compatte del grafico di Pareto). */
function getCausaliSigla_() {
  var sheet = ss_().getSheetByName(SHEET_CAUSALI);
  var last = sheet.getLastRow();
  var map = {};
  if (last < 2) return map;
  var values = sheet.getRange(2, 2, last - 1, 3).getValues(); // B (Causale), C (Provenienza), D (Sigla)
  for (var i = 0; i < values.length; i++) {
    var causale = (values[i][0] || '').toString().trim();
    var sigla = (values[i][2] || '').toString().trim();
    if (causale) map[causale] = sigla || causale;
  }
  return map;
}

/**
 * Aggrega le quantità per causale nell'intervallo richiesto (Pareto, ordinato decrescente).
 * @param {number} startMs epoch ms inizio intervallo (incluso)
 * @param {number} endMs   epoch ms fine intervallo (incluso)
 */
function getKpiPareto(startMs, endMs) {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  var last = sheet.getLastRow();
  if (last < 2 || !isFinite(startMs) || !isFinite(endMs)) return { rows: [], total: 0 };

  var values = sheet.getRange(2, 1, last - 1, 6).getValues(); // A..F
  var map = {};
  for (var i = 0; i < values.length; i++) {
    var d = values[i][0];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < startMs || t > endMs) continue;
    var causale = (values[i][5] || '').toString().trim();
    if (!causale) continue;
    var q = Number(values[i][4]) || 0;
    map[causale] = (map[causale] || 0) + q;
  }

  var siglaMap = getCausaliSigla_();
  var rows = [];
  var total = 0;
  for (var k in map) {
    rows.push({ causale: k, sigla: siglaMap[k] || k, quantita: map[k] });
    total += map[k];
  }
  rows.sort(function (a, b) { return b.quantita - a.quantita; });
  return { rows: rows, total: total };
}

/* ============================ SCRITTURE ============================ */

/**
 * Registra un nuovo scarto in fondo al foglio "Base dati".
 * @param {Object} data { articolo, variante, quantita, stabilimento, operatore, causale }
 */
function registraScarto(data) {
  data = data || {};
  var articolo = (data.articolo || '').toString().trim();
  var varianteRaw = (data.variante || '').toString().trim();
  var quantita = Number(data.quantita);
  var stabilimento = (data.stabilimento || '').toString().trim();
  var operatore = (data.operatore || '').toString().trim();
  var causale = (data.causale || '').toString().trim();
  var fotos = Array.isArray(data.foto) ? data.foto : [];

  if (articolo.length !== 13) throw new Error('Il codice articolo deve avere 13 caratteri.');
  if (!/^\d+$/.test(varianteRaw)) throw new Error('La variante deve essere solo numerica.');
  if (!quantita || quantita <= 0) throw new Error('La quantità deve essere maggiore di zero.');
  if (STABILIMENTI.indexOf(stabilimento) === -1) throw new Error('Stabilimento non valido.');
  if (!operatore) throw new Error('Operatore obbligatorio.');
  if (!causale) throw new Error('Causale obbligatoria.');

  // Variante salvata come testo con zeri iniziali fino a un minimo di 6 cifre.
  var variante = varianteRaw;
  while (variante.length < 6) variante = '0' + variante;

  var fotoUrls = uploadFotos_(fotos);

  var sheet = ss_().getSheetByName(SHEET_DATI);
  var r = sheet.getLastRow() + 1;
  // Articolo (C) e Variante (D) forzati a testo PRIMA della scrittura, per non perdere
  // zeri iniziali / le 13 cifre del codice. Scrittura in un'unica getRange().setValues()
  // così le colonne C e D vengono sempre alimentate (niente dipendenza da getLastRow()
  // dopo appendRow()).
  sheet.getRange(r, 3, 1, 2).setNumberFormat('@');
  sheet.getRange(r, 1, 1, 9).setValues([[new Date(), stabilimento, articolo, variante, quantita, causale, operatore, '', 'NO']]);
  sheet.getRange(r, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  // Colonna H "Immagini": link cliccabili con testo breve ("Foto 1", "Foto 2", …).
  if (fotoUrls.length) sheet.getRange(r, 8).setRichTextValue(buildFotoRichText_(fotoUrls));

  return { ok: true, articolo: articolo, variante: variante, quantita: quantita };
}

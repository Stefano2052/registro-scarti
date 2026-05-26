/**
 * Registro Scarti - Berloni Bagno
 * Backend API JSON per Google Apps Script.
 *
 * IMPORTANTE:
 * - Apps Script deve essere modificabile solo da te o collaboratori fidati.
 * - La web app può restare pubblica/anonima.
 * - Progetto STANDALONE (non container-bound): accede al foglio via
 *   SpreadsheetApp.openById, quindi servono gli scope completi.
 * - Scope richiesti:
 *   - spreadsheets (accesso ai Fogli, non a Drive)
 *   - drive.file (solo i file creati da quest'app, non il resto del Drive)
 * - Servizio avanzato richiesto: Drive API v3 (abilitare in "Servizi avanzati")
 */

var SPREADSHEET_ID = '1KfuV1-iQcWAutqGJthjycV2aTDv1GFiT3ppYbWAE1i4';
var SHEET_DATI = 'Base dati';
var SHEET_CAUSALI = 'Causali';

var STABILIMENTI = ['BB1', 'BB3', 'Ipiemme', 'Zenobi'];

// Cartella Drive per le foto, creata e gestita dall'app stessa.
// DriveApp richiede lo scope drive completo: si usa invece il servizio
// avanzato Drive (Drive.Files.*) che funziona con drive.file per i file
// creati dall'app. L'ID della cartella è memorizzato nelle proprietà dello
// script e riusato a ogni upload successivo.
var DRIVE_FOLDER_NAME = 'Registro Scarti - Immagini';

/* ============================ DRIVE FOTO ============================ */

function getDriveFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOTO_FOLDER_ID');

  if (id) {
    try {
      Drive.Files.get(id, {fields: 'id'});
      return id;
    } catch (e) {
      // Cartella rimossa o non più accessibile: ne creo una nuova.
    }
  }

  var folder = Drive.Files.create(
    {name: DRIVE_FOLDER_NAME, mimeType: 'application/vnd.google-apps.folder'},
    null,
    {fields: 'id'}
  );
  props.setProperty('FOTO_FOLDER_ID', folder.id);
  return folder.id;
}

function uploadFotos_(fotos) {
  if (!fotos || !fotos.length) return [];

  var folderId = getDriveFolder_();
  var timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  );

  var urls = [];

  for (var i = 0; i < fotos.length; i++) {
    try {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(fotos[i]),
        'image/jpeg',
        timestamp + '_' + (i + 1) + '.jpg'
      );

      var file = Drive.Files.create(
        {name: blob.getName(), parents: [folderId]},
        blob,
        {fields: 'id'}
      );

      // Rende il singolo file visibile a chiunque abbia il link, così le
      // foto restano apribili dai link salvati nel foglio.
      Drive.Permissions.create({role: 'reader', type: 'anyone'}, file.id);

      urls.push('https://drive.google.com/file/d/' + file.id + '/view');
    } catch (e) {
      console.warn('Errore upload foto ' + (i + 1), e);
    }
  }

  return urls;
}

function buildFotoRichText_(urls) {
  var labels = [];

  for (var i = 0; i < urls.length; i++) {
    labels.push('Foto ' + (i + 1));
  }

  var text = labels.join('\n');
  var builder = SpreadsheetApp.newRichTextValue().setText(text);

  var pos = 0;

  for (var j = 0; j < labels.length; j++) {
    builder.setLinkUrl(pos, pos + labels[j].length, urls[j]);
    pos += labels[j].length + 1;
  }

  return builder.build();
}

/* ============================ ROUTING API ============================ */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  try {
    if (action === 'init') return json_(getInitData());

    if (action === 'kpi') {
      return json_(
        getKpiPareto(
          Number(e.parameter.start),
          Number(e.parameter.end),
          e.parameter.stabilimento || ''
        )
      );
    }

    if (action === 'storico') {
      return json_(
        getStorico(
          e.parameter.stabilimento || '',
          e.parameter.operatore || '',
          Number(e.parameter.limit)
        )
      );
    }

    return json_({ error: 'Azione non valida' });
  } catch (err) {
    return json_({ error: getErrorMessage_(err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Body JSON mancante.');
    }

    var body = JSON.parse(e.postData.contents);

    if (body.action === 'registra') {
      return json_(registraScarto(body.data || {}));
    }

    return json_({ error: 'Azione non valida' });
  } catch (err) {
    return json_({ error: getErrorMessage_(err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getErrorMessage_(err) {
  return (err && err.message) ? err.message : String(err);
}

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/* ============================ LETTURE ============================ */

function getInitData() {
  return {
    stabilimenti: STABILIMENTI,
    causali: getCausali_(),
    operatori: getOperatori_(),
    articoli: getArticoli_()
  };
}

function getCausali_() {
  var sheet = ss_().getSheetByName(SHEET_CAUSALI);
  if (!sheet) throw new Error('Foglio "' + SHEET_CAUSALI + '" non trovato.');

  var last = sheet.getLastRow();
  if (last < 2) return [];

  var values = sheet.getRange(2, 2, last - 1, 1).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || '').trim();

    if (v && out.indexOf(v) === -1) {
      out.push(v);
    }
  }

  return out;
}

function getDistinctSorted_(column) {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  if (!sheet) throw new Error('Foglio "' + SHEET_DATI + '" non trovato.');

  var last = sheet.getLastRow();
  if (last < 2) return [];

  var values = sheet.getRange(2, column, last - 1, 1).getValues();
  var seen = {};
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || '').trim();
    var key = v.toLowerCase();

    if (v && !seen[key]) {
      seen[key] = true;
      out.push(v);
    }
  }

  out.sort(function (a, b) {
    return a.localeCompare(b, 'it', { sensitivity: 'base' });
  });

  return out;
}

function getOperatori_() {
  return getDistinctSorted_(8);
}

function getArticoli_() {
  return getDistinctSorted_(3);
}

function getCausaliSigla_() {
  var sheet = ss_().getSheetByName(SHEET_CAUSALI);
  if (!sheet) throw new Error('Foglio "' + SHEET_CAUSALI + '" non trovato.');

  var last = sheet.getLastRow();
  var map = {};

  if (last < 2) return map;

  var values = sheet.getRange(2, 2, last - 1, 3).getValues();

  for (var i = 0; i < values.length; i++) {
    var causale = String(values[i][0] || '').trim();
    var sigla = String(values[i][2] || '').trim();

    if (causale) {
      map[causale] = sigla || causale;
    }
  }

  return map;
}

function getKpiPareto(startMs, endMs, stabilimento) {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  if (!sheet) throw new Error('Foglio "' + SHEET_DATI + '" non trovato.');

  var last = sheet.getLastRow();

  if (last < 2 || !isFinite(startMs) || !isFinite(endMs)) {
    return { rows: [], total: 0 };
  }

  var filtroStab = String(stabilimento || '').trim();

  var values = sheet.getRange(2, 1, last - 1, 7).getValues();
  var map = {};

  for (var i = 0; i < values.length; i++) {
    var d = values[i][0];

    if (!(d instanceof Date)) continue;

    var t = d.getTime();

    if (t < startMs || t > endMs) continue;

    if (filtroStab && String(values[i][1] || '').trim() !== filtroStab) continue;

    var causale = String(values[i][6] || '').trim();
    if (!causale) continue;

    var q = Number(values[i][5]) || 0;
    map[causale] = (map[causale] || 0) + q;
  }

  var siglaMap = getCausaliSigla_();
  var rows = [];
  var total = 0;

  for (var k in map) {
    rows.push({
      causale: k,
      sigla: siglaMap[k] || k,
      quantita: map[k]
    });

    total += map[k];
  }

  rows.sort(function (a, b) {
    return b.quantita - a.quantita;
  });

  return {
    rows: rows,
    total: total
  };
}

// Storico delle registrazioni ordinate dalla più recente, con filtro
// opzionale per stabilimento e operatore. Il limite cappa quante righe
// tornano al client (il front-end ne mostra poi 5 per volta).
function getStorico(stabilimento, operatore, limit) {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  if (!sheet) throw new Error('Foglio "' + SHEET_DATI + '" non trovato.');

  var last = sheet.getLastRow();
  if (last < 2) return { rows: [] };

  var filtroStab = String(stabilimento || '').trim();
  var filtroOp = String(operatore || '').trim();

  var values = sheet.getRange(2, 1, last - 1, 8).getValues();
  var siglaMap = getCausaliSigla_();
  var rows = [];

  for (var i = 0; i < values.length; i++) {
    var d = values[i][0];
    if (!(d instanceof Date)) continue;

    var stab = String(values[i][1] || '').trim();
    if (filtroStab && stab !== filtroStab) continue;

    var op = String(values[i][7] || '').trim();
    if (filtroOp && op !== filtroOp) continue;

    var causale = String(values[i][6] || '').trim();

    rows.push({
      data: d.getTime(),
      stabilimento: stab,
      articolo: String(values[i][2] || '').trim(),
      variante: String(values[i][3] || '').trim(),
      descrizione: String(values[i][4] || '').trim(),
      quantita: Number(values[i][5]) || 0,
      causale: causale,
      sigla: siglaMap[causale] || causale,
      operatore: op
    });
  }

  rows.sort(function (a, b) {
    return b.data - a.data;
  });

  var lim = Number(limit);
  if (isFinite(lim) && lim > 0 && rows.length > lim) {
    rows = rows.slice(0, lim);
  }

  return { rows: rows };
}

/* ============================ DIAGNOSTICA ============================ */

// Intestazioni attese nel foglio "Base dati", nell'ordine delle colonne da cui
// dipende il codice. Si confrontano in modo "tollerante" (minuscole, senza
// accenti, per prefisso) così "Registrato su Essentia" combacia con "Registrato".
var HEADER_ATTESI = [
  'Data',
  'Stabilimento',
  'Articolo',
  'Variante',
  'Descrizione',
  'Quantità',
  'Causale',
  'Operatore',
  'Immagini',
  'Registrato'
];

function normHeader_(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

// Verifica che le colonne di "Base dati" siano allineate a quanto si aspetta il
// codice. Eseguibile dall'editor Apps Script (Esegui -> doctor): l'esito è nel
// log e nel valore restituito. Utile dopo aver spostato/aggiunto colonne.
function doctor() {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  if (!sheet) throw new Error('Foglio "' + SHEET_DATI + '" non trovato.');

  var nCol = Math.max(sheet.getLastColumn(), HEADER_ATTESI.length);
  var headers = sheet.getRange(1, 1, 1, nCol).getValues()[0];

  var colonne = [];
  var ok = true;

  for (var i = 0; i < HEADER_ATTESI.length; i++) {
    var atteso = HEADER_ATTESI[i];
    var trovato = String(headers[i] || '').trim();
    var allineato = normHeader_(trovato).indexOf(normHeader_(atteso)) === 0;
    if (!allineato) ok = false;
    colonne.push({
      colonna: i + 1,
      atteso: atteso,
      trovato: trovato,
      esito: allineato ? 'OK' : 'DISALLINEATO'
    });
  }

  var result = { ok: ok, colonne: colonne };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/* ============================ SCRITTURE ============================ */

function registraScarto(data) {
  data = data || {};

  var articolo = String(data.articolo || '').trim();
  var varianteRaw = String(data.variante || '').trim();
  var descrizione = String(data.descrizione || '').trim();
  var quantita = Number(data.quantita);
  var stabilimento = String(data.stabilimento || '').trim();
  var operatore = String(data.operatore || '').trim();
  var causale = String(data.causale || '').trim();
  var fotos = Array.isArray(data.foto) ? data.foto : [];

  // Quando l'operatore non conosce articolo e variante usa la descrizione:
  // in quel caso articolo e variante non sono obbligatori.
  if (descrizione) {
    articolo = '';
    varianteRaw = '';
  } else {
    if (articolo.length !== 13) {
      throw new Error('Il codice articolo deve avere 13 caratteri.');
    }

    if (!/^\d+$/.test(varianteRaw)) {
      throw new Error('La variante deve essere solo numerica.');
    }
  }

  if (!quantita || quantita <= 0) {
    throw new Error('La quantità deve essere maggiore di zero.');
  }

  if (STABILIMENTI.indexOf(stabilimento) === -1) {
    throw new Error('Stabilimento non valido.');
  }

  if (!operatore) {
    throw new Error('Operatore obbligatorio.');
  }

  if (!causale) {
    throw new Error('Causale obbligatoria.');
  }

  var variante = varianteRaw;

  while (variante && variante.length < 6) {
    variante = '0' + variante;
  }

  var fotoUrls = uploadFotos_(fotos);

  var sheet = ss_().getSheetByName(SHEET_DATI);
  if (!sheet) throw new Error('Foglio "' + SHEET_DATI + '" non trovato.');

  var r = sheet.getLastRow() + 1;

  sheet.getRange(r, 3, 1, 2).setNumberFormat('@');

  sheet.getRange(r, 1, 1, 10).setValues([[
    new Date(),
    stabilimento,
    articolo,
    variante,
    descrizione,
    quantita,
    causale,
    operatore,
    '',
    'NO'
  ]]);

  sheet.getRange(r, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  if (fotoUrls.length) {
    sheet.getRange(r, 9).setRichTextValue(buildFotoRichText_(fotoUrls));
  }

  return {
    ok: true,
    articolo: articolo,
    variante: variante,
    descrizione: descrizione,
    quantita: quantita,
    foto: fotoUrls
  };
}

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
 *   - spreadsheets
 *   - drive.file (accesso ai soli file creati dall'app)
 */

var SPREADSHEET_ID = '1KfuV1-iQcWAutqGJthjycV2aTDv1GFiT3ppYbWAE1i4';
var SHEET_DATI = 'Base dati';
var SHEET_CAUSALI = 'Causali';

var STABILIMENTI = ['BB1', 'BB3', 'Ipiemme', 'Zenobi'];

// Cartella Drive per le foto, creata e gestita dall'app stessa.
// Con lo scope drive.file l'app può accedere solo ai file che crea, quindi
// non si può usare una cartella preesistente: la creiamo al primo upload e
// ne memorizziamo l'ID nelle proprietà dello script.
var DRIVE_FOLDER_NAME = 'Registro Scarti - Foto';

/* ============================ DRIVE FOTO ============================ */

function getDriveFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOTO_FOLDER_ID');

  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      // Cartella rimossa o non più accessibile: la ricreo sotto.
    }
  }

  var folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  props.setProperty('FOTO_FOLDER_ID', folder.getId());
  return folder;
}

function uploadFotos_(fotos) {
  if (!fotos || !fotos.length) return [];

  var folder = getDriveFolder_();
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

      var file = folder.createFile(blob);

      // Rende il singolo file visibile a chiunque abbia il link, così le
      // foto restano apribili dai link salvati nel foglio.
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      urls.push(file.getUrl());
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
          Number(e.parameter.end)
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
  return getDistinctSorted_(7);
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

function getKpiPareto(startMs, endMs) {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  if (!sheet) throw new Error('Foglio "' + SHEET_DATI + '" non trovato.');

  var last = sheet.getLastRow();

  if (last < 2 || !isFinite(startMs) || !isFinite(endMs)) {
    return { rows: [], total: 0 };
  }

  var values = sheet.getRange(2, 1, last - 1, 6).getValues();
  var map = {};

  for (var i = 0; i < values.length; i++) {
    var d = values[i][0];

    if (!(d instanceof Date)) continue;

    var t = d.getTime();

    if (t < startMs || t > endMs) continue;

    var causale = String(values[i][5] || '').trim();
    if (!causale) continue;

    var q = Number(values[i][4]) || 0;
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

/* ============================ SCRITTURE ============================ */

function registraScarto(data) {
  data = data || {};

  var articolo = String(data.articolo || '').trim();
  var varianteRaw = String(data.variante || '').trim();
  var quantita = Number(data.quantita);
  var stabilimento = String(data.stabilimento || '').trim();
  var operatore = String(data.operatore || '').trim();
  var causale = String(data.causale || '').trim();
  var fotos = Array.isArray(data.foto) ? data.foto : [];

  if (articolo.length !== 13) {
    throw new Error('Il codice articolo deve avere 13 caratteri.');
  }

  if (!/^\d+$/.test(varianteRaw)) {
    throw new Error('La variante deve essere solo numerica.');
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

  while (variante.length < 6) {
    variante = '0' + variante;
  }

  var fotoUrls = uploadFotos_(fotos);

  var sheet = ss_().getSheetByName(SHEET_DATI);
  if (!sheet) throw new Error('Foglio "' + SHEET_DATI + '" non trovato.');

  var r = sheet.getLastRow() + 1;

  sheet.getRange(r, 3, 1, 2).setNumberFormat('@');

  sheet.getRange(r, 1, 1, 9).setValues([[
    new Date(),
    stabilimento,
    articolo,
    variante,
    quantita,
    causale,
    operatore,
    '',
    'NO'
  ]]);

  sheet.getRange(r, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  if (fotoUrls.length) {
    sheet.getRange(r, 8).setRichTextValue(buildFotoRichText_(fotoUrls));
  }

  return {
    ok: true,
    articolo: articolo,
    variante: variante,
    quantita: quantita,
    foto: fotoUrls
  };
}

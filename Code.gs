/**
 * Registro Scarti - Berloni Bagno
 * Web app Google Apps Script per la registrazione degli scarti di produzione.
 *
 * Foglio "Base dati":  A Data | B Stabilimento | C Articolo | D Variante |
 *                      E Quantità | F Causale | G Operatore | H Registrato su Essentia
 * Foglio "Causali":    A Classificazione | B Causale | C Provenienza | D Sigla
 */

var SPREADSHEET_ID = '1KfuV1-iQcWAutqGJthjycV2aTDv1GFiT3ppYbWAE1i4';
var SHEET_DATI = 'Base dati';
var SHEET_CAUSALI = 'Causali';
var STABILIMENTI = ['BB3', 'Ipiemme'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Registro Scarti')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .addMetaTag('theme-color', '#E79E2F')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'default')
    .addMetaTag('apple-mobile-web-app-title', 'Scarti')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/** Dati iniziali per popolare i menu a tendina. */
function getInitData() {
  return {
    stabilimenti: STABILIMENTI,
    causali: getCausali(),
    operatori: getOperatori()
  };
}

/** Valori della colonna B "Causale" del foglio "Causali" (ordine del foglio, senza duplicati). */
function getCausali() {
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

/** Operatori già presenti nel file (colonna G di "Base dati"), distinti e ordinati alfabeticamente. */
function getOperatori() {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 7, last - 1, 1).getValues();
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

  if (articolo.length !== 13) throw new Error('Il codice articolo deve avere 13 caratteri.');
  if (!/^\d+$/.test(varianteRaw)) throw new Error('La variante deve essere solo numerica.');
  if (!quantita || quantita <= 0) throw new Error('La quantità deve essere maggiore di zero.');
  if (STABILIMENTI.indexOf(stabilimento) === -1) throw new Error('Stabilimento non valido.');
  if (!operatore) throw new Error('Operatore obbligatorio.');
  if (!causale) throw new Error('Causale obbligatoria.');

  // Variante salvata come testo con zeri iniziali fino a un minimo di 6 cifre.
  var variante = varianteRaw;
  while (variante.length < 6) variante = '0' + variante;

  var sheet = ss_().getSheetByName(SHEET_DATI);
  sheet.appendRow([new Date(), stabilimento, '', '', quantita, causale, operatore, 'NO']);
  var r = sheet.getLastRow();
  sheet.getRange(r, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  // Articolo e Variante forzati a formato testo per non perdere zeri iniziali / notazione.
  sheet.getRange(r, 3, 1, 2).setNumberFormat('@');
  sheet.getRange(r, 3).setValue(articolo);
  sheet.getRange(r, 4).setValue(variante);

  return { ok: true, articolo: articolo, variante: variante, quantita: quantita };
}

/**
 * Aggrega le quantità per causale nell'intervallo richiesto (Pareto, ordinato decrescente).
 * @param {number} startMs epoch ms inizio intervallo (incluso)
 * @param {number} endMs   epoch ms fine intervallo (incluso)
 */
function getKpiPareto(startMs, endMs) {
  var sheet = ss_().getSheetByName(SHEET_DATI);
  var last = sheet.getLastRow();
  if (last < 2) return { rows: [], total: 0 };

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

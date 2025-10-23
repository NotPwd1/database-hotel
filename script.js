/**
 * Script Apps Script Definitivo - Supporta Prenotazioni, Registrazioni PA, Rinnovi e Disdette
 */

const SHEET_ID = '1ygcJnTd6p9yX7x8XP3bZvMdNvYbaxLRBP_QBcaRNc50';
const SHEET_OGGETTI = 'OggettiPAO';
const SHEET_PRENOTAZIONI = 'Prenotazioni';
const SHEET_DISDETTE = 'Disdette';
const DISDETTE_GID = '692494043';

function addCorsHeaders(output) {
  if (output && typeof output.setHeader === 'function') {
    output.setHeader('Access-Control-Allow-Origin', '*');
    output.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  return output;
}

function doOptions(e) {
  return addCorsHeaders(ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT_PLAIN));
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const app = (params.app || params.page || 'prenotazioni').toString().toLowerCase();
  
  if (app === 'getuseremail') {
    return getUserEmail(e);
  }
  
  if (app === 'prenotazioni' && params.action === 'findLast') {
    return findLastPrenotazione(params.cliente, params.camera);
  }
  
  if (app === 'disdette' || app === 'disdire') {
    return doGetDisdette(e);
  }
  
  if (app.indexOf('registraz') === 0 || app === 'registrazioni') {
    return doGetRegistrazioniPA(e);
  }
  
  return doGetPrenotazioni(e);
}

function getUserEmail(e) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      email: userEmail || ''
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message,
      email: ''
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Trova l'ultima prenotazione per un cliente e camera specifici
 */
function findLastPrenotazione(cliente, camera) {
  try {
    if (!cliente || !camera) {
      throw new Error('Cliente e Camera sono obbligatori');
    }
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PRENOTAZIONI);
    if (!sheet) throw new Error('Sheet Prenotazioni not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const clienteIdx = headers.indexOf('Cliente');
    const cameraIdx = headers.findIndex(h => 
      h === 'Numero Camera' || h === 'Camera' || h === 'Numero'
    );
    
    if (clienteIdx === -1 || cameraIdx === -1) {
      throw new Error('Colonne Cliente o Camera non trovate');
    }
    
    let lastMatch = null;
    let lastRowNumber = -1;
    
    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      const rowCliente = String(row[clienteIdx] || '').trim();
      const rowCamera = String(row[cameraIdx] || '').trim();
      
      if (rowCliente === cliente && rowCamera === camera) {
        lastMatch = {};
        headers.forEach((h, idx) => {
          const v = row[idx];
          if (v instanceof Date) {
            try {
              const tz = Session.getScriptTimeZone();
              lastMatch[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
            } catch (e) {
              lastMatch[h] = v.toISOString().split('T')[0];
            }
          } else {
            lastMatch[h] = v;
          }
        });
        lastMatch.__rowNumber = i + 1;
        lastRowNumber = i + 1;
        break;
      }
    }
    
    if (!lastMatch) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        found: false,
        message: 'Nessuna prenotazione trovata'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      found: true,
      data: lastMatch
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    Logger.log(err);
    return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON));
  }
}

function doPost(e) {
  let app = '';
  try {
    if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents || '{}');
      app = (body.app || body.page || '').toString().toLowerCase();
    }
  } catch (err) {}
  
  if (!app && e && e.parameter && e.parameter.app) {
    app = e.parameter.app.toString().toLowerCase();
  }
  
  if (app === 'disdette' || app === 'disdire') {
    return doPostDisdette(e);
  }
  
  if (app.indexOf('registraz') === 0 || app === 'registrazioni') {
    return doPostRegistrazioniPA(e);
  }
  
  return doPostPrenotazioni(e);
}

/**
 * GET per Disdette - Restituisce tutti i dati del foglio Disdette
 */
function doGetDisdette(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.row && params.field && params.value !== undefined) {
      const finalValue = updateDisdetteValue(params.row, params.field, params.value);
      const payload = JSON.stringify({
        success: true,
        row: parseInt(params.row, 10),
        field: params.field,
        value: finalValue
      });
      if (params.callback) {
        return addCorsHeaders(ContentService.createTextOutput(params.callback + '(' + payload + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT));
      }
      return addCorsHeaders(ContentService.createTextOutput(payload)
        .setMimeType(ContentService.MimeType.JSON));
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DISDETTE);
    if (!sheet) throw new Error('Sheet Disdette not found');
    const data = sheetDataToObjects(sheet);
    const out = JSON.stringify(data);
    if (params.callback) {
      return addCorsHeaders(ContentService.createTextOutput(params.callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT));
    }
    return addCorsHeaders(ContentService.createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON));
  } catch (err) {
    Logger.log(err);
    return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON));
  }
}

function updateDisdetteValue(row, field, value) {
  const rowNumber = parseInt(row, 10);
  if (!rowNumber || rowNumber < 2) throw new Error('Invalid row');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_DISDETTE);
  if (!sheet) throw new Error('Sheet Disdette not found');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = headers.indexOf(field) + 1;
  if (col === 0) throw new Error('Column ' + field + ' not found');
  let finalValue = value;
  if (field === 'Disdetta' || field === 'Rinnovato') {
    finalValue = (value === true || value === 'true' || value === '1' || String(value).toLowerCase() === 'true');
  }
  sheet.getRange(rowNumber, col).setValue(finalValue);
  SpreadsheetApp.flush();
  return finalValue;
}

/**
 * POST per Disdette - Aggiorna il campo Disdetta per una riga specifica
 */
function doPostDisdette(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    let body = {};
    
    if (e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        Logger.log('Failed to parse JSON, falling back to parameters: ' + jsonErr.message);
        if (e.parameter) {
          body = e.parameter;
        }
      }
    } else if (e.parameter) {
      body = e.parameter;
    }

    if (body.row && body.field && (body.value !== undefined)) {
      const finalValue = updateDisdetteValue(body.row, body.field, body.value);
      return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
        success: true,
        row: parseInt(body.row, 10),
        field: body.field,
        value: finalValue
      })).setMimeType(ContentService.MimeType.JSON));
    }

    throw new Error('No valid action (row, field, and value required)');
  } catch (err) {
    Logger.log(err);
    return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON));
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doGetPrenotazioni(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PRENOTAZIONI);
    if (!sheet) throw new Error('Sheet not found');
    
    const params = (e && e.parameter) || {};
    const data = sheetDataToObjects(sheet);
    const out = JSON.stringify(data);
    
    if (params.callback) {
      return addCorsHeaders(ContentService.createTextOutput(params.callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT));
    }
    
    return addCorsHeaders(ContentService.createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON));
  } catch (err) {
    Logger.log(err);
    return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON));
  }
}

function doPostPrenotazioni(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    let body = {};
    
    if (e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        Logger.log('Failed to parse JSON, falling back to parameters: ' + jsonErr.message);
        if (e.parameter) {
          body = e.parameter;
        }
      }
    } else if (e.parameter) {
      body = e.parameter;
    }

    // Aggiorna campo specifico
    if (body.row && body.field && (body.value !== undefined)) {
      const row = parseInt(body.row, 10);
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_PRENOTAZIONI);
      if (!sheet) throw new Error('Sheet not found');
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const col = headers.indexOf(body.field) + 1;
      if (col === 0) throw new Error('Column ' + body.field + ' not found');
      
      sheet.getRange(row, col).setValue(body.value);
      SpreadsheetApp.flush();
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Append nuova riga
    if (body.append === 'true' || body.append === true || Array.isArray(body.values)) {
      const userEmail = Session.getActiveUser().getEmail();
      if (userEmail && !body.elaboratoDa && !body.elaborato) {
        body.elaboratoDa = userEmail;
      }
      
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_PRENOTAZIONI);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      let values;
      if (Array.isArray(body.values)) {
        values = body.values;
      } else {
        values = headers.map(h => {
          const key = (h || '').toString().trim();
          const lk = key.toLowerCase();
          
          if (lk === 'numero camera' || lk === 'camera' || lk === 'numero') {
            return body.numeroCamera || body.numero || '';
          }
          if (lk === 'cliente' || lk === 'nome cliente' || lk === 'nome') {
            return body.cliente || body.nomeCliente || '';
          }
          if (lk === 'telegram') return body.telegram || '';
          if (lk === 'data di nascita' || lk === 'datanascita' || lk === 'data nascita') {
            return body.dataNascita || '';
          }
          if (lk === 'totale giorni' || lk === 'giorni' || lk === 'totalegiorni') {
            return body.totaleGiorni || '';
          }
          if (lk === 'checkin' || lk === 'check-in' || lk === 'check in') {
            return body.checkIn || '';
          }
          if (lk === 'checkout' || lk === 'check-out' || lk === 'check out') {
            return body.checkOut || '';
          }
          if (lk === 'polizza') return body.polizza || '';
          if (lk === 'dipendente' || lk === 'responsabile') {
            return body.dipendente || '';
          }
          if (lk === 'giorni poseidon' || lk === 'giorniposeidon') {
            return body.giorniPoseidon || body.giorni_poseidon || '';
          }
          if (lk === 'prezzocamera' || lk === 'prezzo camera' || lk === 'prezzo_camera' || lk === 'prezzo') {
            return body.prezzoCamera ? cleanCurrencyValue(body.prezzoCamera) : '';
          }
          if (lk === 'prezzopolizza' || lk === 'prezzo polizza' || lk === 'prezzo_polizza') {
            return body.prezzoPolizza ? cleanCurrencyValue(body.prezzoPolizza) : '';
          }
          if (lk === 'elaborato da' || lk === 'elaboratoda' || lk === 'elaborato') {
            return body.elaboratoDa || body.elaborato || '';
          }
          
          return body[key] || '';
        });
      }
      
      sheet.appendRow(values);
      SpreadsheetApp.flush();
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        appended: true
      })).setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error('No valid action (body.append missing or false)');
  } catch (err) {
    Logger.log(err);
    return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON));
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doGetRegistrazioniPA(e) {
  try {
    const params = (e && e.parameter) || {};
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OGGETTI);
    if (!sheet) throw new Error('Sheet not found');

    if ((params.row || params.__rowNumber) && (params.Riscattato !== undefined)) {
      const rowRaw = params.row || params.__rowNumber;
      const row = parseInt(rowRaw, 10);
      if (isNaN(row) || row < 2) throw new Error('Invalid row number');
      
      const rawVal = params.Riscattato;
      const riscNormalized = (rawVal === true || rawVal === 'true' || rawVal === '1' || String(rawVal).toLowerCase() === 'true');
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const col = headers.indexOf('Riscattato') + 1;
      if (col === 0) throw new Error('Riscattato column not found');
      
      sheet.getRange(row, col).setValue(riscNormalized);
      SpreadsheetApp.flush();
      
      const updatedRowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const resultObj = { success: true, row: row, Riscattato: riscNormalized };
      headers.forEach(function(h, i) { resultObj[h] = updatedRowValues[i]; });
      
      const out = JSON.stringify(resultObj);
      if (params.callback) {
        return ContentService.createTextOutput(params.callback + '(' + out + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      
      return ContentService.createTextOutput(out)
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheetDataToObjects(sheet);
    const out = JSON.stringify(data);
    
    if (params.callback) {
      return addCorsHeaders(ContentService.createTextOutput(params.callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT));
    }
    
    return addCorsHeaders(ContentService.createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON));
  } catch (err) {
    Logger.log(err);
    return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON));
  }
}

function doPostRegistrazioniPA(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    var body = {};
    if (e.postData && e.postData.type && e.postData.type.indexOf('application/json') !== -1) {
      body = JSON.parse(e.postData.contents || '{}');
    } else if (e.parameter) {
      body = e.parameter;
    }

    if (body.append === 'true' || body.append === true || (body.nomeCliente && !body.row && !body.__rowNumber)) {
      const userEmail = Session.getActiveUser().getEmail();
      if (userEmail && !body.elaboratoDa && !body.elaborato) {
        body.elaboratoDa = userEmail;
      }
      
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_OGGETTI);
      if (!sheet) throw new Error('Sheet not found');
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      const values = headers.map(h => {
        const key = (h || '').toString().trim();
        const lk = key.toLowerCase();
        
        if (lk === 'cliente' || lk === 'nome' || lk === 'nome cliente') {
          return body.nomeCliente || body.Cliente || '';
        }
        if (lk === 'camera' || lk === 'numero camera' || lk === 'numero') {
          return body.numeroCamera || body.Camera || '';
        }
        if (lk === 'termine polizza' || lk === 'data fine polizza' || lk === 'termine') {
          return body.dataFinePolizza || body.TerminePolizza || body.Termine || '';
        }
        if (lk === 'dipendente' || lk === 'responsabile') {
          return body.dipendente || '';
        }
        if (lk === 'elaborato da' || lk === 'elaboratoda' || lk === 'elaborato') {
          return body.elaboratoDa || body.elaborato || '';
        }
        if (lk === 'descrizione' || lk === 'note') {
          return body.descrizione || '';
        }
        if (lk === 'riscattato') {
          return (body.Riscattato === true || body.Riscattato === 'true' || body.Riscattato === '1') ? true : false;
        }
        
        return body[key] || '';
      });

      sheet.appendRow(values);
      SpreadsheetApp.flush();
      
      const newRow = sheet.getLastRow();
      const updatedRowValues = sheet.getRange(newRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const resp = { success: true, appended: true, row: newRow };
      headers.forEach(function(h, i) { resp[h] = updatedRowValues[i]; });
      
      return ContentService.createTextOutput(JSON.stringify(resp))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rowRaw = body.row || body.__rowNumber;
    if (!rowRaw) throw new Error('row or __rowNumber required');
    
    const row = parseInt(rowRaw, 10);
    if (isNaN(row) || row < 2) throw new Error('Invalid row number');

    if (body.Riscattato === undefined) throw new Error('Riscattato value required');
    
    const rawVal = body.Riscattato;
    const riscNormalized = (rawVal === true || rawVal === 'true' || rawVal === '1' || String(rawVal).toLowerCase() === 'true');

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OGGETTI);
    if (!sheet) throw new Error('Sheet not found');
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col = headers.indexOf('Riscattato') + 1;
    if (col === 0) throw new Error('Riscattato column not found');

    sheet.getRange(row, col).setValue(riscNormalized);
    SpreadsheetApp.flush();

    const updatedRowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const resultObj = { success: true, row: row, Riscattato: riscNormalized };
    headers.forEach(function(h, i) { resultObj[h] = updatedRowValues[i]; });

    return ContentService.createTextOutput(JSON.stringify(resultObj))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return addCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON));
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sheetDataToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) return [];
  
  const headers = values.shift();
  var tz = Session.getScriptTimeZone();
  
  try {
    if (sheet && typeof sheet.getParent === 'function') {
      var ss = sheet.getParent();
      if (ss && typeof ss.getSpreadsheetTimeZone === 'function') {
        tz = ss.getSpreadsheetTimeZone();
      }
    }
  } catch (e) {}

  return values.map((row, idx) => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      if (v instanceof Date) {
        try {
          obj[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        } catch (e) {
          obj[h] = v.toISOString().split('T')[0];
        }
      } else {
        obj[h] = v;
      }
    });
    obj.__rowNumber = idx + 2;
    return obj;
  });
}

function cleanCurrencyValue(value) {
  if (!value) return '';
  return String(value).replace(/EUR/g, '').replace(/€/g, '').replace(/\s/g, '').trim();
}
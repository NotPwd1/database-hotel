/**
 * SCRIPT COMPLETO APPS SCRIPT
 * Copia tutto questo nel tuo script.gs
 */

// Authentication removed: this Apps Script deployment no longer enforces
// token-based authentication. All endpoints below are reachable without
// a server-side token. Keep in mind this makes the web app public; if
// you need protection, reintroduce server-side auth or restrict access
// via the Apps Script deployment configuration.

// ===== CONFIGURAZIONE SHEET =====
const SHEET_ID = '1ygcJnTd6p9yX7x8XP3bZvMdNvYbaxLRBP_QBcaRNc50';
const SHEET_OGGETTI = 'OggettiPAO';
const SHEET_PRENOTAZIONI = 'Prenotazioni';
const SHEET_DISDETTE = 'Disdette';
const SHEET_DIPENDENTI = 'Dipendenti';
const DISDETTE_GID = '692494043';

// ===== PASSWORD E GESTIONE SESSIONI =====
// Mappa password -> author identificativo leggibile
const AUTHORIZED_PASSWORDS_MAP = {
  'pwd3772': 'Operatore_A',
  'sharky1344': 'Operatore_B',
  'test': 'Tester'
};

// Password specifiche per disdette
const DISDETTE_PASSWORDS = {
  'disdette': 'Operatore_Disdette'
};

// Gestione sessioni (24 ore)
function createSession(author) {
  const token = Utilities.getUuid();
  const session = {
    token: token,
    author: author,
    createdAt: new Date().getTime()
  };
  
  // Salva sessione in PropertiesService (permanente) e in cache (veloce)
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty('session_' + token, JSON.stringify(session));
  
  // Anche in cache per accesso veloce
  CacheService.getScriptCache().put(
    'session_' + token, 
    JSON.stringify(session), 
    21600  // 6 ore
  );
  
  return token;
}

function getSession(token) {
  if (!token) return null;
  
  try {
    const cache = CacheService.getScriptCache();
    let sessionData = cache.get('session_' + token);
    
    // Se non in cache, prova a recuperare da PropertiesService
    if (!sessionData) {
      const scriptProps = PropertiesService.getScriptProperties();
      sessionData = scriptProps.getProperty('session_' + token);
      
      // Se trovato in Properties, rimetti in cache
      if (sessionData) {
        cache.put('session_' + token, sessionData, 21600);
      }
    }
    
    if (!sessionData) return null;
    
    const session = JSON.parse(sessionData);
    const now = new Date().getTime();
    
    // Verifica scadenza (24 ore)
    if (now - session.createdAt > 24 * 3600 * 1000) {
      cache.remove('session_' + token);
      PropertiesService.getScriptProperties().deleteProperty('session_' + token);
      return null;
    }
    
    return session;
  } catch (e) {
    console.error('Errore lettura sessione:', e);
    return null;
  }
}

function removeSession(token) {
  if (token) {
    // Rimuovi da entrambi Cache e Properties
    CacheService.getScriptCache().remove('session_' + token);
    PropertiesService.getScriptProperties().deleteProperty('session_' + token);
  }
}

// ===== FUNZIONI CORS =====
function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT_PLAIN);
}

/**
 * Endpoint pubblico per verificare una password inviata dal client.
 * Uso: GET ?app=checkPassword&password=laPassword
 */
function doLogin(e) {
  try {
    const params = (e && e.parameter) || {};
    const pwd = (params.password || '').toString();
    const page = (params.page || params.app || '').toString().toLowerCase();

    if (!pwd) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        message: 'Password mancante' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let author = null;

    // Gestione password specifiche per disdette
    if (page === 'disdette') {
      author = DISDETTE_PASSWORDS[pwd] || null;
      if (!author) {
        return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          message: 'Password disdette non valida' 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    } else {
      // Password globali
      author = AUTHORIZED_PASSWORDS_MAP[pwd] || null;
      if (!author) {
        return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          message: 'Password non valida' 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Crea sessione se password valida
    const token = createSession(author);
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true, 
      token: token,
      author: author
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doLogin error: ' + err);
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      error: err.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doLogout(e) {
  try {
    const params = (e && e.parameter) || {};
    const token = params.token;
    
    if (token) {
      removeSession(token);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      error: err.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Middleware per verificare l'autenticazione
function checkAuth(e) {
  const params = (e && e.parameter) || {};
  const token = params.token;
  
  if (!token) {
    return { 
      authenticated: false, 
      message: 'Token mancante' 
    };
  }
  
  const session = getSession(token);
  if (!session) {
    return { 
      authenticated: false, 
      message: 'Sessione scaduta o non valida' 
    };
  }
  
  return { 
    authenticated: true, 
    author: session.author 
  };
}

// ===== MAIN ROUTER =====
function doGet(e) {
  const params = (e && e.parameter) || {};
  const app = (params.app || params.page || 'prenotazioni').toString().toLowerCase();
  
  // Endpoint per verificare password dal client (password memorizzate server-side)
  if (app === 'login') {
    return doLogin(e);
  }
  
  if (app === 'logout') {
    return doLogout(e);
  }
  // Public routing (authentication removed)
  if (params.action === 'findLast' && app === 'prenotazioni') {
    return findLastPrenotazione(params.cliente, params.camera);
  }
  // Endpoint per ottenere le ultime camere registrate / rinnovate
  if (app === 'lastrooms' || app === 'ultimecamere') {
    const auth = checkAuth(e);
    if (!auth.authenticated) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: auth.message
      })).setMimeType(ContentService.MimeType.JSON)
      .setResponseCode(401);
    }
    return doGetLastRooms(e);
  }

  if (app === 'disdette' || app === 'disdire') {
    const auth = checkAuth(e);
    if (!auth.authenticated) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: auth.message
      })).setMimeType(ContentService.MimeType.JSON)
      .setResponseCode(401);
    }
    return doGetDisdette(e);
  }

  if (app === 'dipendenti') {
    return doGetDipendenti();
  }

  if (app.indexOf('registraz') === 0 || app === 'registrazioni') {
    return doGetRegistrazioniPA(e);
  }

  return doGetPrenotazioni(e);
}

function doPost(e) {
  let app = '';
  let body = {};
  
  // Estrai il parametro app
  try {
    if (e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents || '{}');
        app = (body.app || body.page || '').toString().toLowerCase();
      } catch (err) {
        body = e.parameter || {};
        app = (body.app || '').toString().toLowerCase();
      }
    } else if (e.parameter) {
      body = e.parameter;
      app = (body.app || '').toString().toLowerCase();
    }
  } catch (err) {
    Logger.log('Parse error: ' + err);
  }
  
  // Public routing (authentication removed)
  if (app === 'disdette' || app === 'disdire') {
    return doPostDisdette(e);
  }

  if (app.indexOf('registraz') === 0 || app === 'registrazioni') {
    return doPostRegistrazioniPA(e);
  }

  return doPostPrenotazioni(e);
}
// ===== FUNZIONI ORIGINALI =====

function doGetDisdette(e) {
  try {
    const params = (e && e.parameter) || {};
    
    if (params.row && params.field && params.value !== undefined) {
      const row = parseInt(params.row, 10);
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_DISDETTE);
      if (!sheet) throw new Error('Sheet Disdette not found');
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const col = headers.indexOf(params.field) + 1;
      if (col === 0) throw new Error('Column ' + params.field + ' not found');
      
      let finalValue = params.value;
      if (params.field === 'Disdetta' || params.field === 'Rinnovato') {
        finalValue = (params.value === true || params.value === 'true' || params.value === '1' || String(params.value).toLowerCase() === 'true');
      }
      
      sheet.getRange(row, col).setValue(finalValue);
      // Se l'autore è passato (es. password usata in client), scrivilo nella colonna 'Autore' se esiste
      if (params.Autore) {
        const authorCol = headers.indexOf('Autore') + 1;
        if (authorCol > 0) {
          sheet.getRange(row, authorCol).setValue(params.Autore);
        }
      }
      SpreadsheetApp.flush();
      
      const result = {
        success: true,
        row: row,
        field: params.field,
        value: finalValue
      };
      
      const out = JSON.stringify(result);
      
      if (params.callback) {
        return ContentService.createTextOutput(params.callback + '(' + out + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      
      return ContentService.createTextOutput(out)
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DISDETTE);
    if (!sheet) throw new Error('Sheet Disdette not found');
    
    const data = sheetDataToObjects(sheet);
    const out = JSON.stringify(data);
    
    if (params.callback) {
      return ContentService.createTextOutput(params.callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    return ContentService.createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPostDisdette(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    let body = {};
    
    if (e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        if (e.parameter) body = e.parameter;
      }
    } else if (e.parameter) {
      body = e.parameter;
    }

    if (body.row && body.field && (body.value !== undefined)) {
      const row = parseInt(body.row, 10);
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_DISDETTE);
      if (!sheet) throw new Error('Sheet Disdette not found');
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const col = headers.indexOf(body.field) + 1;
      if (col === 0) throw new Error('Column ' + body.field + ' not found');
      
      let finalValue = body.value;
      if (body.field === 'Disdetta' || body.field === 'Rinnovato') {
        finalValue = (body.value === true || body.value === 'true' || body.value === '1' || String(body.value).toLowerCase() === 'true');
      }
      
      sheet.getRange(row, col).setValue(finalValue);
      // Se l'autore è passato via POST, aggiornalo
      if (body.Autore) {
        const authorCol = headers.indexOf('Autore') + 1;
        if (authorCol > 0) {
          sheet.getRange(row, authorCol).setValue(body.Autore);
        }
      }
      SpreadsheetApp.flush();
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        row: row,
        field: body.field,
        value: finalValue
      })).setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error('No valid action');
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
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
      return ContentService.createTextOutput(params.callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    return ContentService.createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
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
        if (e.parameter) body = e.parameter;
      }
    } else if (e.parameter) {
      body = e.parameter;
    }

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

    if (body.append === 'true' || body.append === true) {
      const userEmail = Session.getActiveUser().getEmail();
      // Allow client to pass an 'Autore' which will be used as elaboratoDa
      if (body.Autore) {
        body.elaboratoDa = body.Autore;
      } else if (userEmail && !body.elaboratoDa) {
        body.elaboratoDa = userEmail;
      }
      
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_PRENOTAZIONI);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      const values = headers.map(h => {
        const key = (h || '').toString().trim();
        const lk = key.toLowerCase();
        
        if (lk === 'numero camera' || lk === 'camera' || lk === 'numero') {
          return body.numeroCamera || '';
        }
        if (lk === 'cliente' || lk === 'nome cliente') {
          return body.cliente || body.nomeCliente || '';
        }
        if (lk === 'telegram') return body.telegram || '';
        if (lk === 'data di nascita' || lk === 'datanascita') {
          return body.dataNascita || '';
        }
        if (lk === 'totale giorni') return body.totaleGiorni || '';
        if (lk === 'checkin' || lk === 'check-in') return body.checkIn || '';
        if (lk === 'check-in originale' || lk === 'checkinoriginale') {
          return body.checkInOriginale || body.checkIn || '';
        }
        if (lk === 'checkout' || lk === 'check-out') return body.checkOut || '';
        if (lk === 'polizza') return body.polizza || '';
        if (lk === 'dipendente') return body.dipendente || '';
        if (lk === 'giorni poseidon') return body.giorniPoseidon || '';
        if (lk === 'prezzocamera' || lk === 'prezzo camera') {
          return body.prezzoCamera ? cleanCurrencyValue(body.prezzoCamera) : '';
        }
        if (lk === 'prezzopolizza' || lk === 'prezzo polizza') {
          return body.prezzoPolizza ? cleanCurrencyValue(body.prezzoPolizza) : '';
        }
        if (lk === 'elaborato da' || lk === 'elaboratoda') {
          return body.elaboratoDa || '';
        }
        
        return body[key] || '';
      });
      
      sheet.appendRow(values);
      SpreadsheetApp.flush();
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        appended: true
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Se nessuna azione valida, non è un errore - ritorna success vuoto
    return ContentService.createTextOutput(JSON.stringify({
      success: true
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doGetDipendenti() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DIPENDENTI);
    if (!sheet) throw new Error('Sheet Dipendenti not found');
    const data = sheetDataToObjects(sheet);
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
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
      
      const riscNormalized = (params.Riscattato === true || params.Riscattato === 'true' || params.Riscattato === '1' || String(params.Riscattato).toLowerCase() === 'true');
      
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
      return ContentService.createTextOutput(params.callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    return ContentService.createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPostRegistrazioniPA(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    let body = {};
    if (e.postData && e.postData.type && e.postData.type.indexOf('application/json') !== -1) {
      body = JSON.parse(e.postData.contents || '{}');
    } else if (e.parameter) {
      body = e.parameter;
    }

    if (body.append === 'true' || body.append === true || (body.nomeCliente && !body.row)) {
      const userEmail = Session.getActiveUser().getEmail();
      if (body.Autore) {
        body.elaboratoDa = body.Autore;
      } else if (userEmail && !body.elaboratoDa) {
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
          return body.nomeCliente || '';
        }
        if (lk === 'camera' || lk === 'numero camera') {
          return body.numeroCamera || '';
        }
        if (lk === 'termine polizza' || lk === 'data fine polizza') {
          return body.dataFinePolizza || '';
        }
        if (lk === 'dipendente') {
          return body.dipendente || '';
        }
        if (lk === 'elaborato da') {
          return body.elaboratoDa || '';
        }
        if (lk === 'descrizione') {
          return body.descrizione || '';
        }
        if (lk === 'riscattato') {
          return (body.Riscattato === true || body.Riscattato === 'true') ? true : false;
        }
        
        return body[key] || '';
      });

      sheet.appendRow(values);
      SpreadsheetApp.flush();
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        appended: true
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const rowRaw = body.row || body.__rowNumber;
    if (!rowRaw) throw new Error('row required');
    
    const row = parseInt(rowRaw, 10);
    if (isNaN(row) || row < 2) throw new Error('Invalid row number');

    if (body.Riscattato === undefined) throw new Error('Riscattato required');
    
    const riscNormalized = (body.Riscattato === true || body.Riscattato === 'true' || body.Riscattato === '1');

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
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

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
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Ritorna le ultime N prenotazioni (ultime righe non vuote) dal foglio Prenotazioni.
 * Uso: GET ?app=lastrooms&limit=5
 */
function doGetLastRooms(e) {
  try {
    const params = (e && e.parameter) || {};
    const limit = Math.max(1, parseInt(params.limit, 10) || 5);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PRENOTAZIONI);
    if (!sheet) throw new Error('Sheet Prenotazioni not found');

    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: [] })).setMimeType(ContentService.MimeType.JSON);
    }

    const headers = data[0];
    // Identifica gli indici utili
    const cameraIdx = headers.findIndex(h => h === 'Numero Camera' || h === 'Camera' || h === 'Numero');
    const clienteIdx = headers.findIndex(h => h === 'Cliente' || h === 'Nome Cliente' || h === 'Nome');
    const checkInIdx = headers.findIndex(h => h.toLowerCase().indexOf('check') !== -1 && h.toLowerCase().indexOf('in') !== -1);
    const checkOutIdx = headers.findIndex(h => h.toLowerCase().indexOf('check') !== -1 && h.toLowerCase().indexOf('out') !== -1);
    const totaleIdx = headers.findIndex(h => h.toLowerCase().indexOf('totale') !== -1 || h.toLowerCase().indexOf('totale prenotazione') !== -1);
    const elaboratoIdx = headers.findIndex(h => h.toLowerCase().indexOf('elaborato') !== -1 || h.toLowerCase().indexOf('elaborato da') !== -1 || h.toLowerCase().indexOf('elaboratoda') !== -1);
    const rinnovatoIdx = headers.findIndex(h => h === 'Rinnovato' || h.toLowerCase() === 'rinnovato');

    const results = [];
    // Scorri le righe dal basso verso l'alto e raccogli le ultime `limit` prenotazioni con numero camera
    for (let i = data.length - 1; i > 0 && results.length < limit; i--) {
      const row = data[i];
      const cameraVal = (cameraIdx >= 0) ? row[cameraIdx] : '';
      if (!cameraVal || String(cameraVal).toString().trim() === '') continue;

      const item = {};
      item.Camera = cameraVal;
      item.Cliente = (clienteIdx >= 0) ? row[clienteIdx] : '';

      if (checkInIdx >= 0) item['CheckIn'] = row[checkInIdx] instanceof Date ? Utilities.formatDate(row[checkInIdx], Session.getScriptTimeZone(), 'yyyy-MM-dd') : (row[checkInIdx] || '');
      if (checkOutIdx >= 0) item['CheckOut'] = row[checkOutIdx] instanceof Date ? Utilities.formatDate(row[checkOutIdx], Session.getScriptTimeZone(), 'yyyy-MM-dd') : (row[checkOutIdx] || '');
      if (totaleIdx >= 0) item['Totale'] = row[totaleIdx] || '';
      if (elaboratoIdx >= 0) item['Autore'] = row[elaboratoIdx] || '';
      if (rinnovatoIdx >= 0) item['Rinnovato'] = !!row[rinnovatoIdx];
      item.__rowNumber = i + 1;

      results.push(item);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: results })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function sheetDataToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) return [];
  
  const headers = values.shift();
  let tz = Session.getScriptTimeZone();
  
  try {
    if (sheet && typeof sheet.getParent === 'function') {
      const ss = sheet.getParent();
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

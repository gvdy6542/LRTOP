const MAIN_SS_ID   = '1x_f-IMzhYpUpuhV8jL-Ij6qyTIpOEqwWzJgSUrW9Ihk';   // цени
const CONFIG_SS_ID = MAIN_SS_ID;                                      // конфигурационен Spreadsheet
const EUR_RATE     = 1.95583;                                         // курс евро

// Sheet and caching constants for items index
const ITEMS_SHEET_NAME = '666';
const ITEMS_CACHE_KEY = 'itemsIndex';
const ITEMS_CACHE_TTL = 300; // seconds
const INDEX_FILE_NAME = 'itemsIndex.json';
const ITEMS_CACHE_PARTS_KEY = ITEMS_CACHE_KEY + '_parts';
const ITEMS_CACHE_LIMIT = 100 * 1024; // 100 KB per cache entry

var processedFilesList = [];

/**
 * Reads the items sheet and builds an index by code and barcode.
 * Each entry keeps the original row number for quick reference.
 * @return {{byCode:Object,byBarcode:Object}}
 */
function buildItemsIndex_() {
  const sheet = SpreadsheetApp.openById(MAIN_SS_ID).getSheetByName(ITEMS_SHEET_NAME);
  if (!sheet) return { byCode: {}, byBarcode: {}, byShortCode: {} };

  const rows = sheet.getRange('A:E').getValues();
  const byCode = {};
  const byBarcode = {};
  const byShortCode = {};

  rows.forEach((r, i) => {
    const code = String(r[0]).trim();
    const name = String(r[1]).trim();
    const barcode = String(r[2]).trim();
    const shortCode = String(r[3]).trim();
    if (!code && !barcode && !shortCode) return;

    const rawPrice = String(r[4])
      .replace(/[^0-9.,]/g, '')
      .replace(',', '.')
      .trim();
    const price = parseFloat(rawPrice);

    const item = {
      code: code,
      name: name,
      barcode: barcode,
      shortCode: shortCode,
      price: isNaN(price) ? null : price,
      row: i + 1
    };
    if (code) byCode[code] = item;
    if (barcode) byBarcode[barcode] = item;
    if (shortCode) byShortCode[shortCode] = item;
  });

  return { byCode: byCode, byBarcode: byBarcode, byShortCode: byShortCode };
}

/**
 * Persists the given index JSON in a Drive file.
 * @param {{byCode:Object,byBarcode:Object}} index
 */
function saveIndexToFile_(index) {
  const json = JSON.stringify(index);
  const files = DriveApp.getFilesByName(INDEX_FILE_NAME);
  if (files.hasNext()) {
    const f = files.next();
    f.setContent(json);
  } else {
    DriveApp.createFile(INDEX_FILE_NAME, json, MimeType.PLAIN_TEXT);
  }
}

/**
 * Loads items index JSON from Drive if present.
 * @return {{byCode:Object,byBarcode:Object}|null}
 */
function loadIndexFromFile_() {
  const files = DriveApp.getFilesByName(INDEX_FILE_NAME);
  if (!files.hasNext()) return null;
  const file = files.next();
  try {
    return JSON.parse(file.getBlob().getDataAsString());
  } catch (e) {
    return null;
  }
}

/**
 * Removes all cached items index fragments.
 * @param {Cache} cache
 */
function clearItemsIndexCache_(cache) {
  cache.remove(ITEMS_CACHE_KEY);
  const partsStr = cache.get(ITEMS_CACHE_PARTS_KEY);
  if (partsStr) {
    const parts = parseInt(partsStr, 10);
    for (let i = 0; i < parts; i++) {
      cache.remove(ITEMS_CACHE_KEY + '_' + i);
    }
    cache.remove(ITEMS_CACHE_PARTS_KEY);
  }
}

/**
 * Saves the items index into cache, splitting into chunks if necessary.
 * @param {Cache} cache
 * @param {{byCode:Object,byBarcode:Object}} index
 */
function saveItemsIndexToCache_(cache, index) {
  clearItemsIndexCache_(cache);
  const raw = JSON.stringify(index);
  if (raw.length <= ITEMS_CACHE_LIMIT) {
    cache.put(ITEMS_CACHE_KEY, raw, ITEMS_CACHE_TTL);
  } else {
    const parts = Math.ceil(raw.length / ITEMS_CACHE_LIMIT);
    cache.put(ITEMS_CACHE_PARTS_KEY, String(parts), ITEMS_CACHE_TTL);
    for (let i = 0; i < parts; i++) {
      cache.put(
        ITEMS_CACHE_KEY + '_' + i,
        raw.slice(i * ITEMS_CACHE_LIMIT, (i + 1) * ITEMS_CACHE_LIMIT),
        ITEMS_CACHE_TTL
      );
    }
  }
}

/**
 * Attempts to load items index from cache, reassembling chunks if needed.
 * @param {Cache} cache
 * @return {{byCode:Object,byBarcode:Object}|null}
 */
function loadItemsIndexFromCache_(cache) {
  let raw = cache.get(ITEMS_CACHE_KEY);
  if (!raw) {
    const partsStr = cache.get(ITEMS_CACHE_PARTS_KEY);
    if (partsStr) {
      const parts = parseInt(partsStr, 10);
      const chunks = [];
      for (let i = 0; i < parts; i++) {
        const chunk = cache.get(ITEMS_CACHE_KEY + '_' + i);
        if (chunk) chunks.push(chunk);
      }
      if (chunks.length === parts) {
        raw = chunks.join('');
      }
    }
  }
  if (raw) {
    try { return JSON.parse(raw); } catch (e) {}
  }
  return null;
}

/**
 * Retrieves items index from cache, Drive or rebuilds if missing.
 * @return {{byCode:Object,byBarcode:Object,byShortCode:Object}}
 */
function getItemsIndex_() {
  const cache = CacheService.getScriptCache();
  let index = loadItemsIndexFromCache_(cache);
  if (index) return index;

  index = loadIndexFromFile_();
  if (!index) {
    index = buildItemsIndex_();
    saveIndexToFile_(index);
  }
  saveItemsIndexToCache_(cache, index);
  return index;
}

/**
 * Finds an item by its code.
 * @param {string|number} code
 * @return {?Object}
 */
function findByCode(code) {
  const key = String(code).trim();
  if (!key) return null;
  return getItemsIndex_().byCode[key] || null;
}

/**
 * Finds an item by its barcode.
 * @param {string|number} barcode
 * @return {?Object}
 */
function findByBarcode(barcode) {
  const key = String(barcode).trim();
  if (!key) return null;
  return getItemsIndex_().byBarcode[key] || null;
}

/**
 * Finds an item by its short code.
 * @param {string|number} shortCode
 * @return {?Object}
 */
function findByShortCode(shortCode) {
  const key = String(shortCode).trim();
  if (!key) return null;
  return getItemsIndex_().byShortCode[key] || null;
}

/**
 * Forces rebuilding of the items cache and persisting it to Drive.
 * @return {{ok:boolean,count:number}}
 */
function refreshItemsCache() {
  const index = buildItemsIndex_();
  saveIndexToFile_(index);
  const cache = CacheService.getScriptCache();
  saveItemsIndexToCache_(cache, index);
  return { ok: true, count: Object.keys(index.byCode || {}).length };
}

/**
 * Reads all rows from sheet "666", removes duplicates by article code
 * and writes the unique rows into sheet "666".
 * @return {{count:number}}
 */
function initializeItemsCache() {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const source = ss.getSheetByName(ITEMS_SHEET_NAME);
  if (!source) {
    return { count: 0 };
  }

  const data = source.getDataRange().getValues();
  const header = data.length ? data[0] : [];
  const unique = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const code = String(row[0]).trim();
    if (!code || unique.has(code)) continue;
    unique.set(code, row);
  }

    let target = ss.getSheetByName(ITEMS_SHEET_NAME);
  if (!target) {
    target = ss.insertSheet('666');
  } else {
    target.clear();
  }

  const rows = Array.from(unique.values());
  const output = header.length ? [header].concat(rows) : rows;
  if (output.length) {
    target.getRange(1, 1, output.length, output[0].length).setValues(output);
  }

  Logger.log('initializeItemsCache inserted %s records', rows.length);
  return { count: rows.length };
}

/**
 * Lists known cache entries with their values.
 * @return {{key:string,value:string}[]}
 */
function listCacheEntries() {
  const cache = CacheService.getScriptCache();
  const result = [];

  const add = key => {
    const val = cache.get(key);
    if (val !== null) result.push({ key: key, value: val });
  };

  // Simple keys
  ['itemsIndex', 'itemsCache'].forEach(add);

  // Handle itemsIndex parts
  const indexParts = cache.get('itemsIndex_parts');
  if (indexParts) {
    result.push({ key: 'itemsIndex_parts', value: indexParts });
    const parts = parseInt(indexParts, 10);
    for (let i = 0; i < parts; i++) {
      add('itemsIndex_' + i);
    }
  }

  // Handle itemsCache parts
  const cacheParts = cache.get('itemsCache_parts');
  if (cacheParts) {
    result.push({ key: 'itemsCache_parts', value: cacheParts });
    const parts = parseInt(cacheParts, 10);
    for (let i = 0; i < parts; i++) {
      add('itemsCache_' + i);
    }
  }

  return result;
}

/**
 * Updates a cache entry with a new value.
 * @param {string} key
 * @param {string} value
 * @return {{ok:boolean}}
 */
function updateCacheEntry(key, value) {
  const cache = CacheService.getScriptCache();
  cache.put(String(key), String(value), ITEMS_CACHE_TTL);
  return { ok: true };
}

/**
 * Зарежда данните от „666“ в кеш за бърз достъп.
 * @return {{byCode:Object,byBarcode:Object,byShortCode:Object}}
 */
function loadItemsCache() {
  const sheet = SpreadsheetApp.openById(MAIN_SS_ID)
                               .getSheetByName(ITEMS_SHEET_NAME);
  if (!sheet) return { byCode: {}, byBarcode: {}, byShortCode: {} };

  const rows = sheet.getRange('A:E').getValues();
  const byCode = {};
  const byBarcode = {};
  const byShortCode = {};

  rows.forEach(r => {
    const code = String(r[0]).trim();
    const name = String(r[1]).trim();
    const barcode = String(r[2]).trim();
    const shortCode = String(r[3]).trim();
    if (!code && !barcode && !shortCode) return;

    const rawPrice = String(r[4])
      .replace(/[^0-9.,]/g, '')
      .replace(',', '.')
      .trim();
    const price = parseFloat(rawPrice);

    const item = {
      code: code,
      name: name,
      barcode: barcode,
      shortCode: shortCode,
      price: isNaN(price) ? null : price
    };
    if (code) {
      byCode[code] = item;
    }
    if (barcode) {
      byBarcode[barcode] = item;
    }
    if (shortCode) {
      byShortCode[shortCode] = item;
    }
  });

  const data = { byCode: byCode, byBarcode: byBarcode, byShortCode: byShortCode };

  const cache = CacheService.getScriptCache();
  const raw = JSON.stringify(data);
  const LIMIT = 100 * 1024; // 100 KB
  if (raw.length <= LIMIT) {
    cache.put('itemsCache', raw, 300); // ~5 минути
  } else {
    const parts = Math.ceil(raw.length / LIMIT);
    cache.put('itemsCache_parts', String(parts), 300);
    for (let i = 0; i < parts; i++) {
      cache.put('itemsCache_' + i, raw.slice(i * LIMIT, (i + 1) * LIMIT), 300);
    }
  }
  return data;
}

/**
 * Връща всички артикули от кеша като масив.
 * @return {Array<{code:string,name:string,barcode:string,shortCode:string,price:(number|null)}>} 
 */
function getAllCachedItems() {
  const cache = CacheService.getScriptCache();
  let raw = cache.get('itemsCache');
  if (!raw) {
    const partsStr = cache.get('itemsCache_parts');
    if (partsStr) {
      const parts = parseInt(partsStr, 10);
      const chunks = [];
      for (let i = 0; i < parts; i++) {
        const chunk = cache.get('itemsCache_' + i);
        if (chunk) chunks.push(chunk);
      }
      if (chunks.length === parts) {
        raw = chunks.join('');
      }
    }
  }
  const data = raw ? JSON.parse(raw) : loadItemsCache();
  return Object.values(data.byCode).map(item => ({
    code: item.code,
    name: item.name,
    barcode: item.barcode,
    shortCode: item.shortCode,
    price: item.price
  }));
}

/**
 * Взема артикул от кеша по код, баркод или кратък код.
 * @param {string|number} codeOrBarcode
 * @return {{code:string,name:string,barcode?:string,shortCode?:string,price:(number|null)}|null}
 */
function getItemFromCache(codeOrBarcode) {
  const cache = CacheService.getScriptCache();
  let raw = cache.get('itemsCache');
  if (!raw) {
    const partsStr = cache.get('itemsCache_parts');
    if (partsStr) {
      const parts = parseInt(partsStr, 10);
      let chunks = [];
      for (let i = 0; i < parts; i++) {
        const chunk = cache.get('itemsCache_' + i);
        if (chunk) chunks.push(chunk);
      }
      if (chunks.length === parts) {
        raw = chunks.join('');
      }
    }
  }
  const data = raw ? JSON.parse(raw) : loadItemsCache();
  const key = String(codeOrBarcode).trim();

  const item =
    data.byCode[key] ||
    data.byBarcode[key] ||
    data.byShortCode[key] ||
    findByShortCode(key);
  return item
    ? {
        code: item.code,
        name: item.name,
        barcode: item.barcode,
        shortCode: item.shortCode,
        price: item.price
      }
    : null;
}

// Конфигурация по подразбиране
const DEFAULT_CONFIG = {
  parentFolderId: '',
  revisionParentFolderId: '',
  pprFolderId: '',
  showInterfaceButton: false,
  showReferenceButton: false,
  showLabelsButton: false,
  showPprButtons: false,
  showViewRevisionsBtn: false,
  adminEmails: ''
};

function doGet() {
  // Инициализира конфигурацията при първо зареждане
  getConfig();

  // генерира уникален идентификатор и записва активността
  const clientId = Session.getTemporaryActiveUserKey();
  logClientActivity(clientId, 'load');

  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function loadReferencePage() {
  return HtmlService.createHtmlOutputFromFile('reference.html').getContent();
}

function loadinterfacePage() {
  return HtmlService.createHtmlOutputFromFile('interface.html').getContent();
}

function loadLabelsPage() {
  return HtmlService.createHtmlOutputFromFile('labels.html').getContent();
}

function getConfig() {
  var ss = SpreadsheetApp.openById(CONFIG_SS_ID);
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  }
  var cfg = {};
  var data = sheet.getDataRange().getValues();
  const normalize = v => typeof v === 'string' ? v.trim() : v;
  for (var i = 1; i < data.length; i++) {
    var key = normalize(data[i][0]);
    var value = normalize(data[i][1]);
    if (key === '') continue;
    cfg[key] = value;
  }
  var bool = function(val, def) {
    const v = String(val).toLowerCase();
    return v === 'true' ? true : v === 'false' ? false : def;
  };
  return {
    parentFolderId: cfg.parentFolderId || DEFAULT_CONFIG.parentFolderId,
    revisionParentFolderId: cfg.revisionParentFolderId || DEFAULT_CONFIG.revisionParentFolderId,
    pprFolderId: cfg.pprFolderId || DEFAULT_CONFIG.pprFolderId,
    showInterfaceButton: bool(cfg.showInterfaceButton, DEFAULT_CONFIG.showInterfaceButton),
    showReferenceButton: bool(cfg.showReferenceButton, DEFAULT_CONFIG.showReferenceButton),
    showLabelsButton: bool(cfg.showLabelsButton, DEFAULT_CONFIG.showLabelsButton),
    showPprButtons: bool(cfg.showPprButtons, DEFAULT_CONFIG.showPprButtons),
    showViewRevisionsBtn: bool(cfg.showViewRevisionsBtn, DEFAULT_CONFIG.showViewRevisionsBtn),
    adminEmails: cfg.adminEmails || DEFAULT_CONFIG.adminEmails
  };
}

function saveConfig(config) {
  var ss = SpreadsheetApp.openById(CONFIG_SS_ID);
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
  }
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  var rows = [];
  for (var key in config) {
    if (config.hasOwnProperty(key)) {
      rows.push([key, config[key]]);
    }
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function openAdminPanel() {
  return HtmlService.createHtmlOutputFromFile("admin-panel.html").getContent();
}

function isAdminUser() {
  const conf = getConfig();
  const email = Session.getActiveUser().getEmail();
  const list = (conf.adminEmails || "").split(/,\s*/).filter(String);
  return list.includes(email);
}

/**
 * Returns the stored admin panel button configuration. Each button object
 * contains a label and visibility flag.
 * @return {{label:string,visible:boolean}[]}
 */
function getAdminButtons() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('adminButtons');
  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (e) {}
  }
  return [
    { label: 'Бутон 1', visible: true },
    { label: 'Бутон 2', visible: true },
    { label: 'Бутон 3', visible: true }
  ];
}

/**
 * Persists admin panel button configuration.
 * @param {{label:string,visible:boolean}[]} buttons
 */
function saveAdminButtons(buttons) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('adminButtons', JSON.stringify(buttons || []));
}

// записва или обновява информация за клиент
function logClientActivity(clientId, activity) {
  if (!clientId) return;
  const ss = SpreadsheetApp.openById(CONFIG_SS_ID);
  let sheet = ss.getSheetByName('Clients');
  if (!sheet) {
    sheet = ss.insertSheet('Clients');
    sheet.getRange(1,1,1,4).setValues([["clientId","lastActive","count","activity"]]);
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) {
      sheet.getRange(i+1,2,1,3).setValues([[new Date(), Number(data[i][2] || 0)+1, activity]]);
      return;
    }
  }
  sheet.appendRow([clientId, new Date(), 1, activity]);
}

function getClientId() {
  return Session.getTemporaryActiveUserKey();
}

// връща статистика за клиентите
function getClientStats() {
  const ss = SpreadsheetApp.openById(CONFIG_SS_ID);
  const sheet = ss.getSheetByName('Clients');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const res = [];
  for (let i = 1; i < data.length; i++) {
    res.push({
      clientId: data[i][0],
      lastActive: Utilities.formatDate(new Date(data[i][1]), tz, 'yyyy-MM-dd HH:mm:ss'),
      count: data[i][2],
      activity: data[i][3]
    });
  }
  return res;
}

function broadcastRefresh() {
  const props = PropertiesService.getScriptProperties();
  const ts = Date.now().toString();
  props.setProperty('refreshTimestamp', ts);
  return ts;
}

function getRefreshTimestamp() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('refreshTimestamp') || '';
}

function parseBarcodeSmart(bc) {
  const raw = String(bc).trim();
  // Allow special command barcodes to pass through unparsed so the UI can
  // handle them separately.
  if (/^\*000[0-2]$/.test(raw)) {
    return { code: raw };
  }
  if (raw.length > 20) {
    return { error: 'Невалиден баркод.' };
  }
  if (/^\d{6}$/.test(raw)) {
    const item = getItemFromCache(raw);
    return item
      ? { code: item.code, name: item.name, barcode: item.barcode || raw, qty: null }
      : { error: 'Артикулът не е намерен.' };
  }
  if (/^28\d{10,}/.test(raw)) {
    const itemCode = raw.substring(2, 7);
    const grams = raw.substring(7, 12);
    const qty = grams ? parseInt(grams, 10) / 1000 : 0;
    const codeE = '3' + itemCode;
    let item = getItemFromCache(codeE);
    if (!item) {
      item = getItemFromCache('4' + codeE.substring(1));
    }
    return item
      ? { code: item.code, name: item.name, barcode: raw, qty: qty }
      : { error: 'Артикулът не е намерен.' };
  }
  if (/^\d{8,13}$/.test(raw)) {
    const item = getItemFromCache(raw);
    return item
      ? { code: item.code, name: item.name, barcode: item.barcode || raw, qty: null }
      : { error: 'Артикулът не е намерен.' };
  }
  return { error: 'Невалиден баркод.' };
}

function processBarcode(barcode) {
  const bc = String(barcode).trim();
  const candidates = [bc];

  if (bc.startsWith('28') && bc.length >= 7) {
    const itemCode = bc.substring(2, 7);
    candidates.push('3' + itemCode);
    candidates.push('4' + itemCode);
    candidates.push(itemCode);
  } else if (bc.startsWith('8') && bc.length >= 6) {
    const itemCode = bc.substring(1, 6);
    candidates.push('3' + itemCode);
    candidates.push('4' + itemCode);
    candidates.push(itemCode);
  } else if ((bc.startsWith('3') || bc.startsWith('4')) && bc.length >= 6) {
    candidates.push(bc.substring(1, 6));
  }

  for (const code of candidates) {
    const item = getItemFromCache(code);
    if (item) return item;
  }
  return { error: 'Артикулът с този баркод не е намерен.' };
}
/**
 * Принуждава Apps Script да поиска Drive OAuth права.
 */
function authorizeDrive() {
  // това ще изисква drive scope
  DriveApp.getRootFolder();
}

function startNewRevision(names, store) {
  const ss = SpreadsheetApp.openById(CONFIG_SS_ID);
  const sheetName = 'StartedRevisions';
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1,1,1,3).setValues([["Timestamp","Names","Store"]]);
  }
  sh.appendRow([new Date(), (names || []).join(', '), store]);
  return 'OK';
}







/**
 * Връща списък ревизии за даден магазин
 * @param {string} storeName
 * @return {{id:string,name:string,date:string}[]}
 */
function listRevisions(storeName) {
  storeName = storeName.toLowerCase().trim();
  const conf = getConfig();
  const parentFolderId = conf.parentFolderId || '1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG';
  const root = DriveApp.getFolderById(parentFolderId);
  const matches = [];
  collectFilesRecursively(root, storeName, matches);
  return matches.map(f=>({
    id:   f.getId(),
    name: f.getName(),
    date: Utilities.formatDate(
            f.getDateCreated(),
            Session.getScriptTimeZone(),
            'yyyy-MM-dd')
  }));
}

/** Рекурсивен helper */
function collectFilesRecursively(folder, storeName, matches) {
  if (folder.getName().toLowerCase().startsWith(storeName)) {
    const files = folder.getFiles();
    while(files.hasNext()) matches.push(files.next());
  }
  const subs = folder.getFolders();
  while(subs.hasNext()) collectFilesRecursively(subs.next(), storeName, matches);
}

/**
 * Връща данните от лист „Ревизия“ на Spreadsheet-а с дадено id
 * @param {string} fileId
 * @return {any[][]}
 */
function getRevisionData(fileId) {
  const ss = SpreadsheetApp.openById(fileId);
  const sh = ss.getSheetByName('Ревизия');
  return sh ? sh.getDataRange().getValues() : [];
}






// Функция за намиране на артикулни детайли по баркод чрез кеша
function findItemDetailsByBarcode(itemBarcode) {
  const item = getItemFromCache(itemBarcode);
  if (item) {
    return {
      itemCode: item.code,
      itemName: item.name,
      itemBarcode: itemBarcode
    };
  }
  return null;
}
function saveToGoogleDrive(storeName, tableData) {
  const conf = getConfig();
  const parentFolderId = conf.parentFolderId || "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";

  // Проверка дали всички редове имат еднакъв брой колони
  const columnCount = tableData[0].length;
  for (let i = 1; i < tableData.length; i++) {
    if (tableData[i].length !== columnCount) {
      throw new Error(`Грешка: Броят на колоните в ред ${i + 1} не съвпада с броя на колоните в първия ред.`);
    }
  }

  // Създаване на нов Google Sheet файл
  const spreadsheet = SpreadsheetApp.create(storeName);
  const sheet = spreadsheet.getActiveSheet();

  // Добавяне на данни в листа
  sheet.getRange(1, 1, tableData.length, columnCount).setValues(tableData);

  // Създаване на папка или използване на съществуваща в Google Drive
const folder = DriveApp.getFolderById(parentFolderId); // Родителската папка

// Проверка за съществуването на папка с дадено име
let newFolder;
const folders = folder.getFoldersByName(storeName); // Търсене на папка с дадено име
if (folders.hasNext()) {
  newFolder = folders.next(); // Ако папката съществува, вземаме съществуващата
  Logger.log(`Папка с име "${storeName}" вече съществува.`);
} else {
  newFolder = folder.createFolder(storeName); // Ако няма такава папка, създаваме нова
  Logger.log(`Създадена е нова папка с име "${storeName}".`);
}

// Сега newFolder съдържа или новата, или вече съществуващата папка
Logger.log(`Работим с папка: ${newFolder.getName()}`);


  // Записване на файла в създадената папка
  const file = DriveApp.getFileById(spreadsheet.getId());
  file.moveTo(newFolder);

  // Експортиране на Google Sheet в Excel формат с помощта на Drive API
  const fileId = file.getId();
  const url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  // Създаване на Excel файл в новата папка
  const excelBlob = response.getBlob();
  const excelFile = newFolder.createFile(excelBlob);

  return "Файлът беше успешно създаден в Google Drive.";
}

function copySheetAndCreateNew() {
  // 1. Получаваме съществуващата таблица по ID
  const spreadsheetId = '11SCSY6HUrpu82aQtN0RnmgD8CO1plht9fbf1xv0yKfw'; // ID на оригиналния файл
  const sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // 2. Копираме цялата таблица (всички листове)
  const copiedSpreadsheet = sourceSpreadsheet.copy('Копие на Таблицата');
  
  // 3. Добавяме нов лист с име "ЛР"
  const newSheet = copiedSpreadsheet.insertSheet('ЛР');
  
  // 4. Получаваме данните от готовата функция
  const tableData = getTableData(); // Това е вашата функция за генериране на данни

  // 5. Записваме данните в новия лист
  newSheet.getRange(1, 1, tableData.length, tableData[0].length).setValues(tableData);

  // 6. Създаване на папка в Google Drive, където да се съхранява файлът
  const conf = getConfig();
  const parentFolderId = conf.parentFolderId || "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
  const folder = DriveApp.getFolderById(parentFolderId); // ID на родителската папка
  const file = DriveApp.getFileById(copiedSpreadsheet.getId());

  // Преместваме новия файл в съществуващата папка
  file.moveTo(folder);

  // 7. Експортиране на копието на таблицата в Excel формат
  const fileId = file.getId();
  const url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  // 8. Създаване на Excel файл в същата папка
  const excelBlob = response.getBlob();
  folder.createFile(excelBlob);

  return "Файлът беше успешно създаден и записан в Google Drive.";
}

// Примерна функция за генериране на данни
function getTableData() {
  return [
    ["Име", "Възраст", "Град"],
    ["Иван", 25, "София"],
    ["Мария", 30, "Пловдив"],
    ["Димитър", 22, "Варна"]
  ];
}







//function generateExcel(data) {
  //const sheet = SpreadsheetApp.create("Ревизия").getActiveSpreadsheet();
//  const sheetData = sheet.getActiveSheet();
  
 // data.forEach(row => {
//    sheetData.appendRow(row);
//  });
  
///  const file = DriveApp.getFileById(sheet.getId());
 // const blob = file.getBlob();
 // file.setTrashed(true);  // Изтрива временния файл след генерирането на Blob
  //return blob.getBytes();  // Връща Blob като байтове за запис
//}






function populateSheet(sheet, {name, code, b1Value, color, nameFontSize = 40, b1FontSize = 25, b1Align = "center", activateB1 = false}) {
  sheet.getRange("A1")
       .setValue(name)
       .setFontSize(nameFontSize)
       .setFontWeight("bold")
       .setHorizontalAlignment("center")
       .setBackground(color);
  sheet.getRange("B3").setValue(code);
  sheet.getRange("B1")
       .setValue(b1Value)
       .setFontSize(b1FontSize)
       .setHorizontalAlignment(b1Align)
       .setBackground(color);
  if (activateB1) {
    sheet.getRange("B1").activate();
  }
}

function handleSpecialCodes(code, sheet) {
  switch (code) {
    case "*0000":
      clearFoundSheet();
      SpreadsheetApp.getUi().alert("Страницата 'Намеренo' беше изчистена успешно.");
      resetScanField(sheet);
      return true;
    case "*0001":
      clearRevisionColumnC();
      SpreadsheetApp.getUi().alert("Колона C на 'Ревизия' беше изчистена успешно.");
      resetScanField(sheet);
      return true;
    case "*0002":
      clearDescriptionSheet();
      SpreadsheetApp.getUi().alert("Страницата 'Опис' беше изчистена успешно.");
      resetScanField(sheet);
      return true;
    default:
      return false;
  }
}

function handleShortCode(code, sheet, cache) {
  if (code.length !== 6) return false;

  let itemName = cache.get(code);
  if (!itemName) {
    const item = getItemFromCache(code);
    if (item) {
      itemName = item.name;
      cache.put(code, itemName, 1500);
    }
  }

  if (itemName) {
    populateSheet(sheet, {
      name: itemName,
      code: code,
      b1Value: "Въведи количество тук:",
      color: "#fcd4a9",
      activateB1: true
    });
  } else {
    SpreadsheetApp.getUi().alert("Артикулът с този баркод (6 символа) не е намерен.");
    resetScanField(sheet);
  }
  return true;
}

function handleWeightBarcode(barcode, sheet) {
  if (!barcode.startsWith("28")) return false;

  const itemCode = barcode.substring(0, 7).replace(/^28/, "3");
  let itemCodeE = "4" + itemCode.substring(1);
  const grams = parseInt(barcode.substring(7, 12), 10);
  const quantity = grams / 1000;

  let itemName = findItemNameByCode(itemCode);
  let codeToUse = itemCode;
  if (!itemName) {
    itemName = findItemNameByCode(itemCodeE);
    if (itemName) {
      codeToUse = itemCodeE;
    }
  }

  if (itemName) {
    populateSheet(sheet, {
      name: itemName,
      code: codeToUse,
      b1Value: quantity,
      color: "#95fb77"
    });
    transferToFoundSheet(codeToUse, itemName, quantity);
    transferToDescriptionSheet(codeToUse, itemName, quantity);
    resetScanField(sheet);
    return codeToUse;
  } else {
    SpreadsheetApp.getUi().alert("Артикулът с тегловен баркод не е намерен.");
    resetScanField(sheet);
  }
  return false;
}


function handlePieceBarcode(barcode, sheet) {
  const item = getItemFromCache(barcode.substring(0, 13));
  if (item) {
    populateSheet(sheet, {
      name: item.name,
      code: item.code,
      b1Value: "Въведи количество тук:",
      color: "#aaa9fc",
      nameFontSize: 28,
      activateB1: true
    });
  } else {
    SpreadsheetApp.getUi().alert("Артикулът с този баркод не е намерен.");
    resetScanField(sheet);
  }
  return true;
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();  // Инициализация на sheet тук
  const range = e.range;
  const cache = CacheService.getScriptCache(); // Инициализация на кеша

  // Проверка дали сме на лист "Сканирай"
  if (sheet.getName() === "Сканирай") {
    // Обработка на сканиран баркод в клетка A2
    if (range.getA1Notation() === "A2") {
      const scannedBarcode = String(range.getValue()).trim();
      if (handleSpecialCodes(scannedBarcode, sheet)) return;
      if (!scannedBarcode) {
        resetScanField(sheet);
        return;
      }
      if (handleShortCode(scannedBarcode, sheet, cache)) return;
      const weightCode = handleWeightBarcode(scannedBarcode, sheet);
      if (weightCode) return;
      if (handlePieceBarcode(scannedBarcode, sheet)) return;
      SpreadsheetApp.getUi().alert("Артикулът с този баркод не е намерен.");
      resetScanField(sheet);
      return;
    }

    // Обработка за клетка B1 (въвеждане на количество)
    if (range.getA1Notation() === "B1") {
      const quantity = parseFloat(range.getValue());
      const itemCode = sheet.getRange("B3").getValue();
      const itemName = sheet.getRange("A1").getValue();

      if (isNaN(quantity) || quantity <= 0) {
        SpreadsheetApp.getUi().alert("Моля, въведете валидно количество.");
        return;
      }

      if (itemCode && itemName) {
        transferToFoundSheet(itemCode, itemName, quantity);
        transferToDescriptionSheet(itemCode, itemName, quantity);
        resetScanField(sheet);
        return;
      }
    }
  }
}

// Функция за изчистване на полета и преместване на фокус
function resetScanField(sheet) {
  sheet.getRange("A2").setValue("Сканирай тук").setFontSize(100).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#6ce7db");
  sheet.getRange("A1").setValue("").setBackground("#ffffff");
  sheet.getRange("B1").setValue("").setBackground("#ffffff");
  sheet.getRange("B3").setValue("");
  sheet.getRange("B2").setValue("");
  sheet.getRange("A2").activate();
}

// Функция за запис в "Намеренo"
function transferToFoundSheet(itemCode, itemName, quantity) {
  const foundSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Намеренo");
  const foundData = foundSheet.getDataRange().getValues();

  const existingRow = foundData.find(row => row[0].toString() == itemCode.toString());
  if (existingRow) {
    const rowIndex = foundData.indexOf(existingRow) + 1;
    const quantityCell = foundSheet.getRange(rowIndex, 3);
    const currentQuantity = parseFloat(quantityCell.getValue()) || 0;
    quantityCell.setValue(currentQuantity + quantity);
  } else {
    foundSheet.appendRow([itemCode, itemName, quantity]);
  }

  updateRevisionSheet(itemCode, itemName, quantity);
}

// Функция за актуализиране на лист "Ревизия"
function updateRevisionSheet(itemCode, itemName, quantity) {
  const revisionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ревизия");

  if (!revisionSheet) {
    SpreadsheetApp.getUi().alert("Лист 'Ревизия' не е намерен!");
    return;
  }

  const revisionData = revisionSheet.getDataRange().getValues();
  const revisionCodes = revisionData.map(row => row[0].toString());
  const indexInRevision = revisionCodes.indexOf(itemCode.toString());

  if (indexInRevision === -1) {
    revisionSheet.appendRow([itemCode, itemName, quantity]);
  } else {
    const currentQuantity = parseFloat(revisionData[indexInRevision][2]) || 0;
    const updatedQuantity = currentQuantity + quantity;
    revisionSheet.getRange(indexInRevision + 1, 3).setValue(updatedQuantity);
  }
}

// Намиране на име на артикул по код чрез кеша
function findItemNameByCode(itemCode) {
  const item = getItemFromCache(itemCode);
  if (item) return item.name;
  return findItemNameByCodeInSheet(itemCode);
}

function findItemNameByBarcode(barcode) {
  const item = getItemFromCache(barcode);
  if (item) return item.name;
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ITEMS_SHEET_NAME);
  if (!sh) return null;
  const rng = sh.getDataRange().getValues();
  for (let i = 0; i < rng.length; i++) {
    if (String(rng[i][2]) === String(barcode)) {
      return String(rng[i][1] || '').trim() || null;
    }
  }
  return null;
}

function findItemNumberByBarcode(barcode) {
  const item = getItemFromCache(barcode);
  if (item) return item.code;
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ITEMS_SHEET_NAME);
  if (!sh) return null;
  const rng = sh.getDataRange().getValues();
  for (let i = 0; i < rng.length; i++) {
    if (String(rng[i][2]) === String(barcode)) {
      return String(rng[i][0] || '').trim() || null;
    }
  }
  return null;
}

/**
 * Searches for an item name by code in column A of '666'.
 * @param {string|number} itemCode
 * @return {?string}
 */
function findItemNameByCodeInSheet(itemCode) {
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ITEMS_SHEET_NAME);
  if (!sh) return null;
  const rng = sh.getDataRange().getValues();
  for (let i = 0; i < rng.length; i++) {
    if (
      String(rng[i][0]) === String(itemCode) || // column A
      String(rng[i][2]) === String(itemCode)    // column C (alternative code)
    ) {
      return String(rng[i][1] || '').trim() || null; // column B has the name
    }
  }
  return null;
}

/**
 * Връща баркода от колона C за даден артикулен код (A)
 */
function getBarcodeByCode(itemCode) {
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ITEMS_SHEET_NAME);
  const rng = sh.getDataRange().getValues(); // columns A..C
  for (let i = 0; i < rng.length; i++) {
    if (String(rng[i][0]) === String(itemCode)) { // column A
      return String(rng[i][2] || '');            // column C
    }
  }
  return '';
}
// Функция за запис в "Опис"
function transferToDescriptionSheet(itemCode, itemName, quantity) {
  const descriptionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Опис");
  const currentDate = new Date();  // Взимаме текущата дата и час

  // Добавяме нов ред в "Опис"
  descriptionSheet.appendRow([itemCode, itemName, quantity, currentDate]);
}

function clearFoundSheet() {
  const foundSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Намеренo");
  if (foundSheet) {
    foundSheet.clear(); // Изчистване на цялата страница
  }
}
function clearRevisionColumnC() {
  const revisionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ревизия");
  if (revisionSheet) {
    const columnC = revisionSheet.getRange("C:C");
    columnC.clearContent(); // Изчистване само на съдържанието в колона C
  }
}
function clearDescriptionSheet() {
  const descriptionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Опис");
  if (descriptionSheet) {
    descriptionSheet.clear(); // Изчистване на цялата страница
  }
}
// Функция за извличане на всички артикули чрез кеша
function getItemCodes() {
  const index = getItemsIndex_();
  return Object.keys(index.byCode);
}
function updateItemDetails(itemCode, itemName, barcode, quantity) {
  document.getElementById("outputMessage").textContent = "Артикул: " + itemName;

  // Вземи допълнителната информация
  const additionalInfo = document.getElementById("additionalInfoField").value.trim();

  if (quantity) {
    document.getElementById("outputMessage").textContent += ", Количество: " + quantity.toFixed(3) + " кг";
    appendOrUpdateTable(itemCode, itemName, barcode, quantity, additionalInfo);
  } else {
    showQuantityInput(itemCode, itemName, barcode, additionalInfo);
  }
}
function findMissingValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Лист16"); // Лист16
  var referenceSheet = ss.getSheetByName("Справка"); // Справка

  var sourceRange = sheet.getRange("E:E"); // Цялата колона E в Лист16 (променена от G на E)
  var referenceRange = referenceSheet.getRange("A37:A1000"); // Диапазон в Справка

  var sourceValues = sourceRange.getValues(); // Взимане на всички стойности от Лист16!E
  var referenceValues = referenceRange.getValues(); // Взимане на всички стойности от Справка!A37:A1000

  var missingValues = [];

  // Проверка дали стойностите от Лист16!E съществуват в Справка!A37:A1000
  for (var i = 0; i < sourceValues.length; i++) {
    var found = false;
    // Прекратяваме търсенето, ако срещнем празна клетка
    if (sourceValues[i][0] == "") continue;
    
    for (var j = 0; j < referenceValues.length; j++) {
      if (sourceValues[i][0] == referenceValues[j][0]) {
        found = true;
        break;
      }
    }
    if (!found) {
      missingValues.push(sourceValues[i][0]);
    }
  }

  // Поставяне на липсващите стойности в клетка A1 на Лист "Справка"
  if (missingValues.length > 0) {
    referenceSheet.getRange("A1").setValue("Липсваща стойност е добавена: " + missingValues.join(", "));
  } else {
    referenceSheet.getRange("A1").setValue("Няма липсващи стойности");
  }
}
function processFilesWithProgress() {
  processedFilesList = [];
  const conf = getConfig();
  var parentFolderId = conf.parentFolderId || "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
  var templateFileId = "11khFtYxY39OA9UYSfDStMPmGbbv76fBx-2-ziea7h50";
  var revisionParentFolderId = conf.revisionParentFolderId || "1Yo8oVkgYYmSR5z_cUFR7zdiRFJZkH3n5";
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var revisionFolderName = "ревизия " + dateString;
  
  var revisionParentFolder = DriveApp.getFolderById(revisionParentFolderId);
  var revisionFolder;
  var folders = revisionParentFolder.getFoldersByName(revisionFolderName);
  
  if (folders.hasNext()) {
    revisionFolder = folders.next();
  } else {
    revisionFolder = revisionParentFolder.createFolder(revisionFolderName);
  }

  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var subFolders = parentFolder.getFolders();
  
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    var files = subFolder.getFiles();
    
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();

      var skipProcessing = false;
      var subFolderFiles = subFolder.getFiles();
      while (subFolderFiles.hasNext()) {
        var subFolderFile = subFolderFiles.next();
        if (subFolderFile.getName().startsWith("обработка_")) {
          skipProcessing = true;
          break;
        }
      }
      if (skipProcessing) continue;

      if (!fileName.toLowerCase().includes("export")) {
        var templateFile = DriveApp.getFileById(templateFileId);
        var newFileName = "обработка_" + fileName;
        var copiedFile = templateFile.makeCopy(newFileName, subFolder);
        
        var sourceSpreadsheet = SpreadsheetApp.openById(file.getId());
        var targetSpreadsheet = SpreadsheetApp.openById(copiedFile.getId());
        var sourceSheet = sourceSpreadsheet.getSheets()[0];
        var targetSheet = targetSpreadsheet.getSheetByName("ЛР1");

        if (targetSheet) {
          var data = sourceSheet.getDataRange().getValues();
          targetSheet.clear();
          targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        }

        var revisionSheet = targetSpreadsheet.getSheetByName("ревизия");
        
        if (revisionSheet) {
          var revisionData = revisionSheet.getDataRange().getValues();
          var revisionFileName = "ревизия_" + fileName;
          var revisionFile = SpreadsheetApp.create(revisionFileName);
          var newRevisionSheet = revisionFile.getActiveSheet();
          newRevisionSheet.setName("ревизия");
          newRevisionSheet.getRange(1, 1, revisionData.length, revisionData[0].length).setValues(revisionData);
          
          var revisionFileId = revisionFile.getId();
          var revisionDriveFile = DriveApp.getFileById(revisionFileId);
          revisionFolder.addFile(revisionDriveFile);
          DriveApp.getRootFolder().removeFile(revisionDriveFile);
        }

        processedFilesList.push(newFileName);
      }
    }
  }

  return processedFilesList;
}
function uploadXlsxDataToFile(fileContent, targetFileName) {
    const conf = getConfig();
    try {
        var parentFolderId = conf.parentFolderId || "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
        var parentFolder = DriveApp.getFolderById(parentFolderId);
        
        // Качваме файла в Drive
        var blob = Utilities.newBlob(fileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", targetFileName);
        var uploadedFile = parentFolder.createFile(blob);

        // Конвертираме в Google Spreadsheet
        var fileId = uploadedFile.getId();
        var spreadsheet = SpreadsheetApp.openById(fileId);
        var sheet = spreadsheet.getSheets()[0];

        // Намираме файла, в който трябва да запишем данните
        var subFolders = parentFolder.getFolders();
        while (subFolders.hasNext()) {
            var subFolder = subFolders.next();
            var files = subFolder.getFiles();
            
            while (files.hasNext()) {
                var file = files.next();
                if (file.getName() === targetFileName) {
                    var targetSpreadsheet = SpreadsheetApp.openById(file.getId());
                    var targetSheet = targetSpreadsheet.getSheetByName("ПОСТАВИ ОЦ.СКЛ.ЗАП") || targetSpreadsheet.insertSheet("ПОСТАВИ ОЦ.СКЛ.ЗАП");

                    // Копираме данните
                    var data = sheet.getDataRange().getValues();
                    targetSheet.clear();
                    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

                    // Изтриваме качения .xlsx файл след като сме извлекли данните
                    DriveApp.getFileById(fileId).setTrashed(true);
                    return true;
                }
            }
        }
        return false;
    } catch (e) {
        Logger.log(e.message);
        return false;
    }
}


function processFilesWithProgress() {
  processedFilesList = [];
  const conf = getConfig();
  try {
    var parentFolderId = conf.parentFolderId || "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
    var templateFileId = "1vWgz8j2wWHrP2CYTRCnsiv970kbOfBfMzR0cP5ogodQ";
    var revisionParentFolderId = conf.revisionParentFolderId || "1Yo8oVkgYYmSR5z_cUFR7zdiRFJZkH3n5";
    var today = new Date();
    var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var revisionFolderName = "ревизия " + dateString;

    // Работа с папките и файловете
    var revisionParentFolder = DriveApp.getFolderById(revisionParentFolderId);
    var revisionFolder;
    var folders = revisionParentFolder.getFoldersByName(revisionFolderName);
    
    if (folders.hasNext()) {
      revisionFolder = folders.next();
    } else {
      revisionFolder = revisionParentFolder.createFolder(revisionFolderName);
    }

    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var subFolders = parentFolder.getFolders();
    
    while (subFolders.hasNext()) {
      var subFolder = subFolders.next();
      var files = subFolder.getFiles();
      
      while (files.hasNext()) {
        var file = files.next();
        var fileName = file.getName();

        // Проверка дали да пропуснем обработката
        var skipProcessing = false;
        var subFolderFiles = subFolder.getFiles();
        while (subFolderFiles.hasNext()) {
          var subFolderFile = subFolderFiles.next();
          if (subFolderFile.getName().startsWith("обработка_")) {
            skipProcessing = true;
            break;
          }
        }
        if (skipProcessing) continue;

        // Пропускаме файлове с името "export"
        if (!fileName.toLowerCase().includes("export")) {
          var templateFile = DriveApp.getFileById(templateFileId);
          var newFileName = "обработка_" + fileName;
          var copiedFile = templateFile.makeCopy(newFileName, subFolder);
          
          var sourceSpreadsheet = SpreadsheetApp.openById(file.getId());
          var targetSpreadsheet = SpreadsheetApp.openById(copiedFile.getId());
          var sourceSheet = sourceSpreadsheet.getSheets()[0];
          var targetSheet = targetSpreadsheet.getSheetByName("ЛР1");

          if (targetSheet) {
            var data = sourceSheet.getDataRange().getValues();
            targetSheet.clear();
            targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
          }

          var revisionSheet = targetSpreadsheet.getSheetByName("ревизия");
          
          if (revisionSheet) {
            var revisionData = revisionSheet.getDataRange().getValues();
            var revisionFileName = "ревизия_" + fileName;
            var revisionFile = SpreadsheetApp.create(revisionFileName);
            var newRevisionSheet = revisionFile.getActiveSheet();
            newRevisionSheet.setName("ревизия");
            newRevisionSheet.getRange(1, 1, revisionData.length, revisionData[0].length).setValues(revisionData);
            
            var revisionFileId = revisionFile.getId();
            var revisionDriveFile = DriveApp.getFileById(revisionFileId);
            revisionFolder.addFile(revisionDriveFile);
            DriveApp.getRootFolder().removeFile(revisionDriveFile);
          }

          processedFilesList.push(newFileName);
        }
      }
    }
    
    return processedFilesList; // Връщаме списъка с обработените файлове
  } catch (e) {
    Logger.log(e.message); // Логваме грешката в конзолата
    return ["Грешка при обработката на файловете: " + e.message];
  }
}

function uploadXlsxDataToFile(fileContent, targetFileName) {
  const conf = getConfig();
  try {
    var parentFolderId = conf.parentFolderId || "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    
    // Качваме файла в Drive
    var blob = Utilities.newBlob(fileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", targetFileName);
    var uploadedFile = parentFolder.createFile(blob);

    // Конвертираме в Google Spreadsheet
    var fileId = uploadedFile.getId();
    var spreadsheet = SpreadsheetApp.openById(fileId);
    var sheet = spreadsheet.getSheets()[0];

    // Намираме файла, в който трябва да запишем данните
    var subFolders = parentFolder.getFolders();
    while (subFolders.hasNext()) {
      var subFolder = subFolders.next();
      var files = subFolder.getFiles();
      
      while (files.hasNext()) {
        var file = files.next();
        if (file.getName() === targetFileName) {
          var targetSpreadsheet = SpreadsheetApp.openById(file.getId());
          var targetSheet = targetSpreadsheet.getSheetByName("ПОСТАВИ ОЦ.СКЛ.ЗАП") || targetSpreadsheet.insertSheet("ПОСТАВИ ОЦ.СКЛ.ЗАП");

          // Копираме данните
          var data = sheet.getDataRange().getValues();
          targetSheet.clear();
          targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

          // Изтриваме качения .xlsx файл след като сме извлекли данните
          DriveApp.getFileById(fileId).setTrashed(true);
          return true;
        }
      }
    }
    return false;
  } catch (e) {
    Logger.log(e.message); // Логваме грешката
    return false;
  }
}
function savePPRData(storeName, dateString, tableData, pprNumber, note, reasonType) {
  if (!pprNumber) throw new Error("❗ Моля, въведете номер на ППР.");

  const conf = getConfig();
  const TEMPLATE_ID = '1KBeWbFlYDMXPoMxxz4H0YvfWLQh2ZdK4i_9iMGk9H4I';
  const DESTINATION_FOLDER_ID = conf.pprFolderId || '1avn7paZvq3eHMdIMcH_PBF3sWA2tNM8l';
  const TARGET_SHEETS = ['МЛЯКО ВЪНШНА СТОКА', 'МЛЯКО', 'АГНЕШКО', 'ГОВЕЖДО', 'МЛЕНИ', 'МЕСО', 'КОЛБАСИ'];
  const EMAIL = 'sklad.pld@dmc.farm,mesokombinat_dobrotica@abv.bg,kristiyan.stoynev@dmc.farm, order@dmc.farm';

  try {
    const parentFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const subfolders = parentFolder.getFoldersByName(storeName);
    const targetSubfolder = subfolders.hasNext() ? subfolders.next() : parentFolder.createFolder(storeName);

    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const fileName = `${pprNumber}_${dateString}_${storeName}`;
    const newFile = templateFile.makeCopy(fileName, targetSubfolder);
    const spreadsheet = SpreadsheetApp.openById(newFile.getId());

    // Записваме типа в sheet "МЕТА"
    let metaSheet = spreadsheet.getSheetByName("МЕТА");
    if (!metaSheet) metaSheet = spreadsheet.insertSheet("МЕТА");
    metaSheet.getRange("A1").setValue(reasonType || "");

    const notFound = [];

    tableData.forEach(row => {
      const [code, name, barcode, qtyStr] = row;
      const quantity = parseFloat(qtyStr);
      if (!code || isNaN(quantity)) return;

      let found = false;
      for (const sheetName of TARGET_SHEETS) {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) continue;

        const values = sheet.getRange("A:A").getValues();
        for (let i = 0; i < values.length; i++) {
          if (String(values[i][0]).trim() === String(code).trim()) {
            const cell = sheet.getRange(i + 1, 4);
            const current = parseFloat(cell.getValue()) || 0;
            cell.setValue(current + quantity);
            found = true;
            break;
          }
        }
        if (found) break;
      }

      if (!found) notFound.push([code, name, quantity]);
    });

    if (notFound.length > 0) {
      let nfSheet = spreadsheet.getSheetByName('НЕРАЗПОЗНАТИ АРТИКУЛИ');
      if (nfSheet) spreadsheet.deleteSheet(nfSheet);
      nfSheet = spreadsheet.insertSheet('НЕРАЗПОЗНАТИ АРТИКУЛИ');
      nfSheet.getRange(1, 1, 1, 3).setValues([["Артикулен номер", "Име", "Количество"]]);

      const grouped = {};
      notFound.forEach(([code, name, qty]) => {
        if (!grouped[code]) grouped[code] = { name, qty: 0 };
        grouped[code].qty += qty;
      });

      const rows = Object.entries(grouped).map(([code, obj]) => [code, obj.name, obj.qty]);
      nfSheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }

    SpreadsheetApp.flush();

    const url = newFile.getUrl();
    const exportUrl = `https://docs.google.com/spreadsheets/d/${newFile.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });
    const attachment = response.getBlob().setName(fileName + ".xlsx");

   // 1. Построй HTML таблица с 6 колони
let htmlTable = `
  <table border="1" cellpadding="6" cellspacing="0"
         style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;margin-top:10px;">
    <thead style="background:#f0f0f0;">
      <tr>
        <th>Код</th>
        <th>Име</th>
        <th>Баркод</th>
        <th>Количество</th>
        <th>Ед. цена</th>
        <th>Общо</th>
      </tr>
    </thead>
    <tbody>`;

let grandTotal = 0;  // ще акумулираме общата сума
tableData.forEach(row => {
  const [code, name, barcode, qty, unitPrice, total] = row;
  const q  = parseFloat(qty);
  const up = parseFloat(unitPrice);
  const to = parseFloat(total);
  grandTotal += to;

  htmlTable += `
    <tr>
      <td>${code}</td>
      <td>${name}</td>
      <td>${barcode}</td>
      <td style="text-align:right;">${q.toFixed(3)}</td>
      <td style="text-align:right;">${up.toFixed(2)}</td>
      <td style="text-align:right;">${to.toFixed(2)}</td>
    </tr>`;
});

htmlTable += `</tbody>
  <tfoot>
    <tr style="background:#e0e0e0;">
      <td colspan="5" style="text-align:right;"><strong>Общо:</strong></td>
      <td style="text-align:right;"><strong>${grandTotal.toFixed(2)}</strong></td>
    </tr>
  </tfoot>
</table>`;

// 2. Изпращаме имейла, вграждайки новата htmlTable
MailApp.sendEmail({
  to:       EMAIL,
  subject:  `ППР №${pprNumber} за ${storeName} (${dateString})`,
  htmlBody: `
    <div style="font-family:Arial,sans-serif;color:#333;font-size:16px;">
      <p>✅ <strong>Данните от ППР №${pprNumber}</strong></p>
      <p>
        🏪 <strong>Магазин:</strong> ${storeName}<br>
        📅 <strong>Дата:</strong> ${dateString}<br>
        📂 <strong>Файл в Google Таблици:</strong>
        <a href="${url}" target="_blank" style="color:#1a73e8;">Отвори файла</a>
      </p>
      ${note ? `<p><strong>Причина:</strong> ${note}</p>` : ''}
      ${reasonType ? `<p><strong>Тип ППР:</strong> ${reasonType}</p>` : ''}
      <p><strong>Въведени артикули:</strong></p>
      ${htmlTable}
      <hr style="border:none;border-top:1px solid #ddd;margin:20px 0;">
      <p style="color:#2e7d32;font-size:16px;">
        С най-добри пожелания,<br>
        
      </p>
      <p style="color:#999;font-size:13px;">
        ⚠️ Този имейл е автоматично генериран.
      </p>
    </div>`,
  attachments: [attachment]
});


    return `✅ Данните са записани. Имейл с прикачен Excel файл е изпратен на ${EMAIL}`;
  } catch (e) {
    Logger.log("ГРЕШКА В savePPRData: " + e.message);
    throw new Error("❌ Грешка при създаване на файла: " + e.message);
  }
}



//function testMailAppPermission() {
 // MailApp.sendEmail({
  //  to: "VELICHKO_LIKOV@abv.bg",
  //  subject: "⚙️ Тест на MailApp разрешение",
 //   body: "Този имейл е изпратен с цел да активира нужните разрешения за MailApp.sendEmail."
 // });
//}




function sendDailyPPRReport() {
  const CACHE_SHEET_ID   = '1HCESdWuLCwUv5b-nB9HCqpWH5HniCK3zpqzMgDgSAgo';
  const CACHE_SHEET_NAME = 'история на брака';
  const EMAILS = 'sklad.pld@dmc.farm,kristiyan.stoynev@dmc.farm, v.likov@dmc.farm, m_margitina@abv.bg';
  const STORE_NAMES = {
  '810000': 'С- София, ТЦ Боила',
  '810001': 'М- София, жк. Люлин 1, бл.3 вх. А',
  '810002': 'М- София, жк. Младост 3, бл. 304',
  '810003': 'М- София, бул. Хр. Ботев 59',
  '810004': 'М- София, ул. Дойран 10а',
  '810005': 'М- София, бул. К. Величков',
  '810006': 'М- София, бул. Ал. Дондуков 50',
  '810007': 'М- София, жк. Младост 1, Магазин 6',
  '810008': 'М- София, жк. Младост 1, Магазин 18',
  '840001': 'М- Пловдив, ул. Житен пазар 5',
  '840002': 'М- Пловдив, ул. Солунска 1а',
  '840003': 'М- Пловдив, бул. Шипка 7, ет. 1, обект 19',
  '840004': 'М- Пловдив, бул. Дунав 66',
  '840005': 'М- Пловдив, ул. Петко Петков 11-13',
  '840007': 'М- Пловдив, ул. Георги Кондолов 3',
  '840008': 'М- Пловдив, ул. Георги Измирлиев 65, ет. 1',
  '840009': 'М- Пловдив, бул. Васил Априлов 84',
  '840010': 'М- Пловдив, ул. Славееви гори 95',
  '840011': 'М- Пловдив, ул. Патриарх Евтимий 24',
  '842180': 'С- Цалапица, ул. Тодор Ламбов 2а',
  '842181': 'М- Цалапица, ул. Тодор Ламбов 2а',
  '842182': 'ЕМ- Цалапица, ул. Тодор Ламбов 2а',
  '843001': 'М- Карлово, ул. Свежен 6',
  '844001': 'М- Пазарджик, ул. Ал. Стамболийски 38',
  '860002': 'М- Ст. Загора, ул. Пазарска 13',
  '860003': 'М- Ст. Загора, бул. Цар Симеон Велики 55',
  '863001': 'М- Хасково, бул. България 136 б',
  '876001': 'М- Тутракан, ул. Гео Милев 33',
  '880001': 'М- Бургас, бул. Демокрация 100',
  '880002': 'М- Бургас, жк. Меден Рудник, бл. 258, вх. А',
  '888888': 'Изложения',
  '890001': 'М- Варна, ЦКП до сладкарница Атлант',
  '890002': 'М- Варна, ул. Пирин 2',
  '890003': 'М- Варна, ул. Георги Бенковски 42',
  '890004': 'М- Варна, ул. Иван Вазов 41-43',
  '890005': 'М- Варна, ул. Васил Друмев 21',
  '890006': 'М- Варна, кв. Виница, ул. Цар Борис ІІІ № 10',
  '890007': 'М- Варна, Пазар Чаталджа обект 40 (затворен)',
  '890008': 'М- Варна, Пазар Чайка',
  '890009': 'М- Варна, ул. Д. Икономов 34',
  '890010': 'М- Варна, Пазар Чаталджа 12-13',
  '892951': 'М- Ген. Колево, Чаира',
  '892952': 'Г- Ген. Колево, Чаира',
  '893001': 'М- Добрич, ул. Отец Паисий 23',
  '893002': 'М- Добрич, ЦКП Гъбката',
  '893003': 'М- Добрич, ул. Хр. Ботев 91',
  '893004': 'М- Добрич, ул. Хан Аспарух 8',
  '893005': 'М- Добрич, жк. Добротица бл. 46',
  '893006': 'М- Добрич, ул. 25-ти септември 58',
  '893951': 'М- Овчарово, Стопански двор',
  '893952': 'Г- Овчарово, Стопански двор',
  '893953': 'Р- Овчарово, Стопански двор',
  '895001': 'М- Ген. Тошево, ул. Дочо Михайлов 22',
  '896001': 'М- Балчик, ул. Хр. Ботев 39',
  '896401': 'М- Соколово, ул. Кирил и Методий',
  '896402': 'Г- Соколово, ул. Кирил и Методий',
  '896491': 'М- Кранево, път към Албена',
  '896492': 'Г- Кранево, път към Албена',
  '896502': 'М- Каварна, ул. Георги Кирков 7',
  '896503': 'Г- Каварна, ул. Георги Кирков 7',
  '896801': 'М- Шабла, ул. П. Българанов 6',
  '897001': 'М- Шумен, ул. Цар Освободител 83',
  'HO':      'Главен офис',
  'TESTDOBRI4': 'Тест Магазин Добрич',
  'TESTVARNA':  'Тест Магазин Варна'
};


  const today = new Date();
  const currentMonth = today.getMonth();
  const currentYear  = today.getFullYear();
  const todayStr     = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd.MM.yyyy');

  const sheet = SpreadsheetApp.openById(CACHE_SHEET_ID).getSheetByName(CACHE_SHEET_NAME);
  const data  = sheet.getDataRange().getValues();

  const storesMap = {};

  data.forEach(([date, store, code, name, qty, unit, total]) => {
    const d = new Date(date);
    if (d.getMonth() !== currentMonth || d.getFullYear() !== currentYear) return;

    if (!storesMap[store]) storesMap[store] = {};
    if (!storesMap[store][code]) {
      storesMap[store][code] = { name, qty: 0, unit: parseFloat(unit) || 0 };
    }

    storesMap[store][code].qty += parseFloat(qty) || 0;
  });

  let fullHtmlBody = `
    <div style="font-family:Arial,sans-serif;font-size:15px">
      <h1>📋 Месечен отчет по ППР (само БРАК) – към ${todayStr}</h1>`;

  for (let store in storesMap) {
    let rows = '';
    let totalSum = 0;

    Object.entries(storesMap[store]).forEach(([code, info]) => {
      const amount = info.qty * info.unit;
      totalSum += amount;
      rows += `
        <tr>
          <td>${code}</td>
          <td>${info.name}</td>
          <td style="text-align:right;">${info.qty.toFixed(3)}</td>
          <td style="text-align:right;">${info.unit.toFixed(2)}</td>
          <td style="text-align:right;">${amount.toFixed(2)}</td>
        </tr>`;
    });

    rows += `
      <tr style="font-weight:bold; background:#f9f9f9;">
        <td colspan="4" style="text-align:right;">Общо за ${store}:</td>
        <td style="text-align:right;">${totalSum.toFixed(2)}</td>
      </tr>`;

    fullHtmlBody += `
     <h2>🏪 Магазин: ${STORE_NAMES[store] || store}</h2>

      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; margin-top:10px;">
        <thead style="background:#f0f0f0;">
          <tr>
            <th>Код</th>
            <th>Име</th>
            <th>Количество</th>
            <th>Ед. цена</th>
            <th>Общо</th>
          </tr>
        </thead>
        <tbody>
          ${rows}
        </tbody>
      </table>
      <hr style="margin:40px 0;">`;
  }

  fullHtmlBody += `
      <p style="color:#777;font-size:13px; margin-top:30px;">
        ⚠️ Този имейл е автоматично генериран. Моля, не отговаряйте.
      </p>
    </div>`;

  MailApp.sendEmail({
    to: EMAILS,
    subject: `📋 Месечен отчет по ППР (само БРАК) – към ${todayStr}`,
    htmlBody: fullHtmlBody
  });
}









/* 2. Нова безопасна функция – използва кеша вместо директно четене */
function findItemDetailsByBarcode_MAIN(barcode) {
  const item = getItemFromCache(barcode);
  if (item) {
    return { itemCode: item.code, itemName: item.name, itemBarcode: barcode };
  }
  return null;
}
/* === цена по артикулен код от кеша === */
/**
 * Връща единична цена от кешираните данни
 * @param {string|number} itemCode
 * @return {number|null}
 */
function getPriceByCode(itemCode) {
  const item = getItemFromCache(itemCode);
  return item ? item.price : null;
}

function buildPPRCacheToSheet() {
  const conf = getConfig();
  const PPR_FOLDER_ID   = conf.pprFolderId || '1avn7paZvq3eHMdIMcH_PBF3sWA2tNM8l';
  const PRICES_SHEET_ID = '1x_f-IMzhYpUpuhV8jL-Ij6qyTIpOEqwWzJgSUrW9Ihk';
  const CACHE_SHEET_ID  = '1HCESdWuLCwUv5b-nB9HCqpWH5HniCK3zpqzMgDgSAgo';
  const CACHE_SHEET_NAME = 'история на брака';

  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const thisMonth = today.getMonth();
  const thisYear  = today.getFullYear();

  const cacheSS = SpreadsheetApp.openById(CACHE_SHEET_ID);
  const cacheSheet = cacheSS.getSheetByName(CACHE_SHEET_NAME);
  if (!cacheSheet) throw new Error('Липсва лист "история на брака".');

  const existing = cacheSheet.getDataRange().getValues();
  const existingKeys = new Set();

  existing.forEach(row => {
    const [date, store, code, , , , , filename] = row;
    if (!date || !store || !code || !filename) return;
    const normalizedDate = Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd');
    const key = `${normalizedDate}|${store}|${code}|${filename}`;
    existingKeys.add(key);
  });

  const folder = DriveApp.getFolderById(PPR_FOLDER_ID);
  const subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const storeName = subfolder.getName();
    const files = subfolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const parts = file.getName().split('_');
      if (parts.length < 3) continue;

      const dateStr = parts[1];
      const fileDate = new Date(dateStr);
      if (fileDate.getMonth() !== thisMonth || fileDate.getFullYear() !== thisYear) continue;

      const normalizedDate = Utilities.formatDate(fileDate, tz, 'yyyy-MM-dd');
      const ss = SpreadsheetApp.openById(file.getId());
      const meta = ss.getSheetByName('МЕТА');
      if (!meta) continue;
      const reasonType = meta.getRange('A1').getValue();

      const sheets = ss.getSheets();
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        if (sheetName === 'МЕТА') continue;

        const values = sheet.getRange(`A2:D${sheet.getLastRow()}`).getValues();

        values.forEach(([code, name, , qty]) => {
          if (!code || !qty) return;

          const key = `${normalizedDate}|${storeName}|${code}|${file.getName()}`;
          if (existingKeys.has(key)) return;

          const unitPrice = getPriceByCode(code) || 0;
          const total = parseFloat(qty) * unitPrice;

          cacheSheet.appendRow([
            normalizedDate,
            storeName,
            code,
            name,
            parseFloat(qty).toFixed(3),
            unitPrice.toFixed(2),
            total.toFixed(2),
            file.getName(),
            reasonType
          ]);

          existingKeys.add(key);
        });
      }
    }
  }

  SpreadsheetApp.flush();
}

function getPprData(storeNumber, pprNumber) {
  const conf = getConfig();
  const PPR_FOLDER_ID = conf.pprFolderId || '1avn7paZvq3eHMdIMcH_PBF3sWA2tNM8l';
  const TARGET_SHEETS = ['МЛЯКО ВЪНШНА СТОКА','МЛЯКО','АГНЕШКО','ГОВЕЖДО','МЛЕНИ','МЕСО','КОЛБАСИ'];

  const mainFolder = DriveApp.getFolderById(PPR_FOLDER_ID);
  const storeFolders = mainFolder.getFoldersByName(storeNumber);
  if (!storeFolders.hasNext()) throw new Error('Няма папка за магазина.');

  const files = storeFolders.next().getFiles();
  let targetFile = null;
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (name.startsWith(pprNumber + '_') && name.endsWith('_' + storeNumber)) {
      targetFile = f;
      break;
    }
  }
  if (!targetFile) throw new Error('ППР файлът не е намерен.');

  const ss = SpreadsheetApp.openById(targetFile.getId());
  const rows = [];
  TARGET_SHEETS.forEach(sh => {
    const sheet = ss.getSheetByName(sh);
    if (!sheet) return;
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
    data.forEach(([code,name,,qty]) => {
      if (code && qty) rows.push([code,name,qty]);
    });
  });
  return rows;
}

function updatePprData(storeNumber, pprNumber, tableData) {
  const conf = getConfig();
  const PPR_FOLDER_ID = conf.pprFolderId || '1avn7paZvq3eHMdIMcH_PBF3sWA2tNM8l';
  const TARGET_SHEETS = ['МЛЯКО ВЪНШНА СТОКА','МЛЯКО','АГНЕШКО','ГОВЕЖДО','МЛЕНИ','МЕСО','КОЛБАСИ'];

  const mainFolder = DriveApp.getFolderById(PPR_FOLDER_ID);
  const storeFolders = mainFolder.getFoldersByName(storeNumber);
  if (!storeFolders.hasNext()) throw new Error('Няма папка за магазина.');

  const files = storeFolders.next().getFiles();
  let targetFile = null;
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (name.startsWith(pprNumber + '_') && name.endsWith('_' + storeNumber)) {
      targetFile = f;
      break;
    }
  }
  if (!targetFile) throw new Error('ППР файлът не е намерен.');

  const ss = SpreadsheetApp.openById(targetFile.getId());

  tableData.forEach(row => {
    const code = row[0];
    const qty = parseFloat(row[2]);
    if (!code || isNaN(qty)) return;
    for (const sh of TARGET_SHEETS) {
      const sheet = ss.getSheetByName(sh);
      if (!sheet) continue;
      const vals = sheet.getRange('A:A').getValues();
      for (let i=0; i<vals.length; i++) {
        if (String(vals[i][0]).trim() === String(code)) {
          sheet.getRange(i+1,4).setValue(qty);
          return;
        }
      }
    }
  });

  SpreadsheetApp.flush();
  return 'ППР е актуализиран.';
}

/**
 * Връща справка за брака за даден магазин от началото на
 * текущия месец до днес. Данните се четат от лист "история на брака".
 * @param {string} storeNumber
 * @return {{rows: any[][], total: string}}
 */
function getWasteReport(storeNumber) {
  const CACHE_SHEET_ID   = '1HCESdWuLCwUv5b-nB9HCqpWH5HniCK3zpqzMgDgSAgo';
  const CACHE_SHEET_NAME = 'история на брака';

  const now   = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const tz    = Session.getScriptTimeZone();

  const sheet = SpreadsheetApp.openById(CACHE_SHEET_ID)
                                 .getSheetByName(CACHE_SHEET_NAME);
  if (!sheet) return { rows: [], total: '0.00' };

  const data  = sheet.getDataRange().getValues();
  const rows  = [];
  let total   = 0;

  data.forEach(r => {
    const [date, store, , name, qty, , sum, , type = 'Брак'] = r;
    if (!date || String(store).trim() !== storeNumber) return;
    if (type && type !== 'Брак') return;
    const d = new Date(date);
    if (d < start || d > now) return;

    const q = parseFloat(qty) || 0;
    const s = parseFloat(sum) || 0;
    rows.push([
      Utilities.formatDate(d, tz, 'dd.MM.yyyy'),
      name,
      q.toFixed(3),
      s.toFixed(2)
    ]);
    total += s;
  });

  return { rows, total: total.toFixed(2) };
}

/* ==================== Label Generator ==================== */

function roundEuro(value) {
  const multiplied = value * 1000;
  const thirdDigit = Math.floor(multiplied) % 10;
  const base = Math.floor(multiplied / 10);
  const final = thirdDigit >= 5 ? base + 1 : base;
  return (final / 100).toFixed(2);
}

function fetchProductByBarcode(barcode) {
  return getItemFromCache(barcode);
}

function fetchPreviewData(barcodes) {
  if (!Array.isArray(barcodes)) return [];
  return barcodes
    .map(bc => getItemFromCache(bc))
    .filter(item => item);
}

function generateLabelsSheet(items) {
  if (!Array.isArray(items) || !items.length) {
    throw new Error('Липсват данни за етикети.');
  }

  const ss = SpreadsheetApp.create('Етикети');
  const sh = ss.getActiveSheet();
  sh.getRange(1, 1, 1, 5).setValues([
    ['Код', 'Име', 'Баркод', 'Цена (лв)', 'Цена (€)']
  ]);

  const data = items.map(it => [
    it.code,
    it.name,
    it.barcode,
    it.price ? it.price.toFixed(2) : '',
    it.price ? roundEuro(it.price / 1.95583) : ''
  ]);

  sh.getRange(2, 1, data.length, 5).setValues(data);
  return ss.getUrl();
}

function runGenerateLabels(barcodes) {
  const items = fetchPreviewData(barcodes);
  if (!items.length) {
    throw new Error('Няма намерени продукти.');
  }
  return generateLabelsSheet(items);
}

// Показва прозореца за избор (Selection.html)
function showSelectionSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Selection')
    .setTitle('Избор режим');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Показва твоя вече съществуващ генератор (index.html)
function showLabelsSidebar() {
  var html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ЛР');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Показва менюто (MenuView.html)
function showMenuSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('MenuView')
    .setTitle('Меню');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Показва Sidebar
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ЛР');
  SpreadsheetApp.getUi().showSidebar(html);
}


// Добавя меню в Google Sheets UI при отваряне
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Етикети')
    .addItem('Избор режим', 'showSelectionSidebar')
    .addToUi();
}

// Взима данни от лист "Меню"
function fetchMenuData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Меню');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    var priceRaw = data[i][1];
    if (!name) continue;
    var price = parseFloat(String(priceRaw).replace(',', '.'));
    if (isNaN(price)) continue;
    result.push({
      name: name,
      price: price.toFixed(2)
    });
  }
  return result;
}

function getUsers() {
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('USERS');
  return stored ? JSON.parse(stored) : {};
}

function saveUsers(users) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('USERS', JSON.stringify(users));
}

function hashPassword(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password);
  return digest.map(function(b){
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length == 1 ? '0' + v : v;
  }).join('');
}

function setUser(username, password) {
  var users = getUsers();
  var salt = Utilities.getUuid();
  users[username] = {
    salt: salt,
    hash: hashPassword(password, salt)
  };
  saveUsers(users);
}

function getUser(username) {
  var users = getUsers();
  return users[username];
}

function login(username, password) {
  // Позволява вход с администраторските данни по подразбиране
  if (username === 'admin' && password === 'admin') {
    return { success: true };
  }

  var user = getUser(username);
  if (!user) {
    return { success: false, message: 'Invalid username or password' };
  }
  var hash = hashPassword(password, user.salt);
  if (hash !== user.hash) {
    return { success: false, message: 'Invalid username or password' };
  }
  return { success: true };
}

const MAIN_SS_ID = '1x_f-IMzhYpUpuhV8jL-Ij6qyTIpOEqwWzJgSUrW9Ihk';   // цени
const EUR_RATE   = 1.95583;                                           // курс евро

var processedFilesList = [];

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
  return HtmlService.createHtmlOutputFromFile('index');
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
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  }
  var cfg = {};
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var value = data[i][1];
    if (key === '') continue;
    cfg[key] = value;
  }
  var bool = function(val, def) {
    return val === true || val === 'true' || val === 'TRUE' ? true : val === false || val === 'false' || val === 'FALSE' ? false : def;
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
  var ss = SpreadsheetApp.getActive();
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

function processBarcode(barcode) {
  let itemDetails = findItemDetailsByBarcode(barcode);      // старото търсене
  if (!itemDetails) itemDetails = findItemDetailsByBarcode_MAIN(barcode); // ← fallback

  if (!itemDetails) return { error: 'Артикулът с този баркод не е намерен.' };
  return itemDetails;
}
/**
 * Принуждава Apps Script да поиска Drive OAuth права.
 */
function authorizeDrive() {
  // това ще изисква drive scope
  DriveApp.getRootFolder();
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






// Функция за намиране на артикулни детайли по баркод от колона C
function findItemDetailsByBarcode(itemBarcode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  const data = sheet.getDataRange().getValues(); // Вземаме всички данни от листа

  for (const row of data) {
    // Проверяваме дали баркодът съвпада
    if (row[2].toString() === itemBarcode.toString()) { // Търсене в колона C (Баркод No.)
      return {
        itemCode: row[0],  // Артикулен номер от колона A
        itemName: row[1],  // Описание от колона B
        itemBarcode: row[2] // Баркод от колона C
      };
    }
  }
  return null;  // Ако не намерим съвпадение


  

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






function onEdit(e) {
  const sheet = e.source.getActiveSheet();  // Инициализация на sheet тук
  const range = e.range;
  const cache = CacheService.getScriptCache(); // Инициализация на кеша

  // Проверка дали сме на лист "Сканирай"
  if (sheet.getName() === "Сканирай") {
    // Обработка на сканиран баркод в клетка A2
    if (range.getA1Notation() === "A2") {
      const scannedBarcode = String(range.getValue()).trim();

      // Вмъкване на допълнителна обработка в onEdit
      if (scannedBarcode === "*0000") {
        clearFoundSheet();
        SpreadsheetApp.getUi().alert("Страницата 'Намеренo' беше изчистена успешно.");
        resetScanField(sheet);
        return;
      }

      if (scannedBarcode === "*0001") {
        clearRevisionColumnC();
        SpreadsheetApp.getUi().alert("Колона C на 'Ревизия' беше изчистена успешно.");
        resetScanField(sheet);
        return;
      }

      if (scannedBarcode === "*0002") {
        clearDescriptionSheet();
        SpreadsheetApp.getUi().alert("Страницата 'Опис' беше изчистена успешно.");
        resetScanField(sheet);
        return;
      }

      if (!scannedBarcode) {
        resetScanField(sheet);
        return;
      }

      let itemCode, itemName, isWeightBased = false;

      // Обработка на баркодове с точно 6 символа
      if (scannedBarcode.length === 6) {
        itemCode = scannedBarcode;
        const cachedData = cache.get(itemCode); // Търсене в кеша

        if (cachedData) {
          // Използване на кешираните данни
          itemName = cachedData;
          sheet.getRange("A1").setValue(itemName).setFontSize(40).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#d9fa45");
          sheet.getRange("B3").setValue(itemCode); // Показваме кода
          sheet.getRange("B1").setValue("Въведи количество тук:").setFontSize(25).setBackground("#d9fa45");
          sheet.getRange("B1").activate();
          return;
        }

        // Търсене на артикула в "Лист1" по колона A
        const sheetList1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
        const columnA = sheetList1.getRange("A:A").getValues();
        let rowIndex = -1;

        for (let i = 0; i < columnA.length; i++) {
          if (columnA[i][0] === itemCode) {
            rowIndex = i + 1;
            break;
          }
        }

        if (rowIndex !== -1) {
          itemName = sheetList1.getRange(rowIndex, 2).getValue();
          sheet.getRange("A1").setValue(itemName).setFontSize(40).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#fcd4a9");
          sheet.getRange("B3").setValue(itemCode);
          sheet.getRange("B1").setValue("Въведи количество тук:").setFontSize(25).setBackground("#fcd4a9");
          sheet.getRange("B1").activate();

          // Кеширане на резултата за бъдеща употреба
          cache.put(itemCode, itemName, 1500); // Запазваме в кеша за 25 минути (1500 секунди)
        } else {
          SpreadsheetApp.getUi().alert("Артикулът с този баркод (6 символа) не е намерен.");
          resetScanField(sheet);
        }
        return;
      }

      // Обработка на баркодовете започващи с "28" (тегловни артикули)
      if (scannedBarcode.startsWith("28")) {
        itemCodeE = scannedBarcode.substring(0, 7).replace(/^28/, "3");
        const grams = parseInt(scannedBarcode.substring(7, 12), 10); // Извличаме тегло в грамове
        const quantity = grams / 1000; // Преобразуваме в килограми
        isWeightBased = true;

        itemName = findItemNameByCode(itemCodeE) || findItemNameByCode("4" + itemCodeE.substring(1));

        if (itemName) {
          sheet.getRange("A1").setValue(itemName).setFontSize(40).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#95fb77"); // Показваме името на артикула
          sheet.getRange("B3").setValue(itemCodeE); // Записваме артикула в B3
          sheet.getRange("B1").setValue(quantity).setFontSize(25).setHorizontalAlignment("center").setBackground("#95fb77"); // Записваме количеството в килограми в B1

          // Прехвърляне на информацията в "Намеренo" и "Опис"
          transferToFoundSheet(itemCodeE, itemName, quantity);
          transferToDescriptionSheet(itemCodeE, itemName, quantity);

          resetScanField(sheet);
          return;
        } else {
          SpreadsheetApp.getUi().alert("Артикулът с тегловен баркод не е намерен.");
          resetScanField(sheet);
        }
      }

      // Обработка на бройкови артикули
      const itemBarcode = scannedBarcode.substring(0, 13);
      const itemNumber = findItemNumberByBarcode(itemBarcode);
      itemName = findItemNameByBarcode(itemBarcode);

      if (itemNumber) {
        sheet.getRange("A1").setValue(itemName).setFontSize(28).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#aaa9fc");
        sheet.getRange("B3").setValue(itemNumber);
        sheet.getRange("B1").setValue("Въведи количество тук:").setFontSize(25).setBackground("#aaa9fc");
        sheet.getRange("B1").activate();
      } else {
        SpreadsheetApp.getUi().alert("Артикулът с този баркод не е намерен.");
        resetScanField(sheet);
      }
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

// Намиране на име на артикул по код
function findItemNameByCode(itemCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  const data = sheet.getDataRange().getValues();
  for (const row of data) {
    if (row[0] === itemCode) {
      return row[1];
    }
  }
  return null;
}
// Намиране на артикулен номер по баркод
function findItemNumberByBarcode(itemBarcode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  const data = sheet.getDataRange().getValues();
  for (const row of data) {
    if (row[2] === itemBarcode) {
      return row[0];
    }
  }
  return null;
}

// Намиране на име на артикул по баркод
function findItemNameByBarcode(itemBarcode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  const data = sheet.getDataRange().getValues();
  for (const row of data) {
    if (row[2] === itemBarcode) {
      return row[1];
    }
  }
  return null;
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
// Функция за извличане на данни от колона A на "Лист1"
function getItemCodes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  const data = sheet.getRange("A:A").getValues();
  
  // Връщаме само уникалните стойности от колоната
  return data.filter(row => row[0]).map(row => row[0]);
}

function findItemNameByCode(itemCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Лист1');
  const data = sheet.getDataRange().getValues(); // Извличаме всички данни от листа
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === itemCode) { // Предполага се, че кода е в колона A (индекс 0)
      return data[i][1]; // Името на артикула се намира в колона B (индекс 1)
    }
  }
  
  return null; // Ако не намери артикула, връща null
}
function findItemNameInColumnA(barcode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Лист1');
  const range = sheet.getRange('A:A').getValues();
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] == barcode) {
      return sheet.getRange(i + 1, 2).getValue(); // Връща стойността от съответния ред в колона B
    }
  }
  return null; // Не е намерен артикул
}

function findItemNameInColumnC(barcode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Лист1');
  const range = sheet.getRange('C:C').getValues();
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] == barcode) {
      return sheet.getRange(i + 1, 4).getValue(); // Връща стойността от съответния ред в колона D
    }
  }
  return null; // Не е намерен артикул
}

function findItemNameByCode(itemCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  var data = sheet.getRange("A:A").getValues(); // Прочитаме само колоната A

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == itemCode) {
      return sheet.getRange(i + 1, 2).getValue(); // Връщаме името на артикула от колоната B
    }
  }

  // Ако не е намерено в колоната A, търсим в колоната C
  var dataC = sheet.getRange("C:C").getValues(); // Прочитаме само колоната C
  for (var j = 0; j < dataC.length; j++) {
    if (dataC[j][0] == itemCode) {
      return sheet.getRange(j + 1, 2).getValue(); // Връщаме името на артикула от колоната B
    }
  }

  return null; // Ако няма съвпадение нито в колоната A, нито в колоната C
}
function findItemNameByCodeInColumnC(itemCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  var dataC = sheet.getRange("C:C").getValues(); // Прочитаме само колоната C

  for (var i = 0; i < dataC.length; i++) {
    if (dataC[i][0] == itemCode) {
      return sheet.getRange(i + 1, 2).getValue(); // Връщаме името на артикула от колоната B
    }
  }

  return null; // Ако няма съвпадение в колоната C
}
function getItemDataFromColumnF(itemCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  var range = sheet.getRange("A:F");  // Преглеждаме диапазона от Колона A до Колона F
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == itemCode) {  // Сравняваме артикула в Колона A
      return values[i][5];  // Връщаме стойността от Колона F
    }
  }
  return null;  // Ако не намерим съвпадение
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
function findItemNameByCodeInColumnF(itemCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист1");
  const data = sheet.getRange("A:F").getValues(); // Вземаме всички стойности от колони A до F
  let itemNameFromF = null;

  // Търсене в колона A и C, но вземаме стойността от колона F
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === itemCode || data[i][2] === itemCode) { // Ако кодът е намерен в A или C
      itemNameFromF = data[i][5]; // Вземаме стойността от колона F
      break;
    }
  }

  return itemNameFromF; // Връща стойността от колона F
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









/* 2. Нова безопасна функция – чете директно от Лист1 */
function findItemDetailsByBarcode_MAIN(barcode) {
  const sheet = SpreadsheetApp.openById(MAIN_SS_ID)
                              .getSheetByName('Лист1');
  if (!sheet) return null;            // ако случайно липсва лист

  const data = sheet.getRange('A:C').getValues();   // A-код | B-име | C-баркод
  for (const row of data) {
    if (String(row[2]) === String(barcode)) {
      return { itemCode: row[0], itemName: row[1], itemBarcode: row[2] };
    }
  }
  return null;                        // няма съвпадение
}
/* === цена по артикулен код от Лист1 === */
/**
 * Връща единична цена от „Лист1“
 * – кодът е в колона A (или C); цената – в колона F
 * – поддържа “92,00”, “92.00”, “92,50 лв.” и т.н.
 * @param {string|number} itemCode
 * @return {number|null}
 */
function getPriceByCode(itemCode) {
  const sheet = SpreadsheetApp.openById(MAIN_SS_ID)
                              .getSheetByName('Лист1');
  if (!sheet) return null;

  const rows = sheet.getRange('A:F').getValues();   // A-код | F-цена
  for (const r of rows) {
    // съвпадение по код в A или C
    const codeA = String(r[0]).trim();
    const codeC = String(r[2]).trim();
    if (codeA === String(itemCode) || codeC === String(itemCode)) {

      // чистим цената – махаме букви/интервали, заменяме запетаи с точки
      const raw = String(r[5])
                    .replace(/[^0-9.,]/g, '')   // само цифри . ,
                    .replace(',', '.')
                    .trim();

      const price = parseFloat(raw);
      return isNaN(price) ? null : price;       // valid number → връщаме
    }
  }
  return null;                                   // няма цена
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
  const sheet = SpreadsheetApp.openById(MAIN_SS_ID).getSheetByName('Лист1');
  if (!sheet) return null;

  const rows = sheet.getRange('A:F').getValues();
  for (const r of rows) {
    if (String(r[2]) === String(barcode)) {
      const rawPrice = String(r[5])
        .replace(/[^0-9.,]/g, '')
        .replace(',', '.')
        .trim();
      const price = parseFloat(rawPrice);
      return {
        code: r[0],
        name: r[1],
        barcode: r[2],
        price: isNaN(price) ? null : price
      };
    }
  }
  return null;
}

function fetchPreviewData(barcodes) {
  if (!Array.isArray(barcodes)) return [];
  return barcodes
    .map(bc => fetchProductByBarcode(bc))
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
  var html = HtmlService.createHtmlOutputFromFile('index')
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
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ЛР');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Извиква генерация на листа и връща данни за preview
function runGenerateLabels() {
  generateLabelsSheet();
  return fetchPreviewData();
}

// Чете Sheet1 и подготвя данни за preview
function fetchPreviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('цени');
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var code = data[i][0], name = data[i][1], raw = data[i][2];
    if (!code || !name) continue;
    var price = parseFloat(String(raw).replace(',', '.'));
    if (isNaN(price)) continue;
    var euro = roundEuro(price / 1.95583);

    var barcodeUrl = 'https://bwipjs-api.metafloor.com/?bcid=code128&text=' + encodeURIComponent(code) + '&includetext';
    result.push({
      name: name,
      price: '<div class="price-line">' + price.toFixed(2) + ' лв.</div>' +
             '<div class="price-line">' + euro + ' €</div>',
      barcodeUrl: barcodeUrl
    });
  }
  return result;
}

// Генерира лист "Етикети" с формули IMAGE за баркод
function generateLabelsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName('цени');
  var dst = ss.getSheetByName('Етикети') || ss.insertSheet('Етикети');
  dst.clear();
  var cmToPx = function(cm) { return Math.round(cm * 37.8); };
  var w = cmToPx(6.5), h = cmToPx(4.5);
  for (var c = 1; c <= 4; c++) dst.setColumnWidth(c, w);
  for (var r = 1; r <= 50; r++) dst.setRowHeight(r, h);
  var rows = src.getDataRange().getValues(), r = 1, c = 1;
  for (var i = 1; i < rows.length; i++) {
    var code = rows[i][0], name = rows[i][1], raw = rows[i][2];
    var price = parseFloat(String(raw).replace(',', '.'));
    if (!code || !name || isNaN(price)) continue;
    var euro = roundEuro(price / 1.95583);

    dst.getRange(r, c).setWrap(true).setFontSize(12)
       .setValue(name + String.fromCharCode(10) + price.toFixed(2) + ' лв.   ' + euro + ' €');
    var url = 'https://bwipjs-api.metafloor.com/?bcid=code128&text=' + encodeURIComponent(code) + '&includetext';
    dst.getRange(r+1, c).setFormula('=IMAGE("' + url + '",4,' + h + ',' + cmToPx(0.5) + ')');
    c++; if (c > 4) { c = 1; r += 2; }
  }
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

function fetchProductByBarcode(barcode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('база данни');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() == barcode) {
      var name = data[i][0];
      var rawPrice = data[i][2];
      var price = parseFloat(String(rawPrice).replace(',', '.'));
      if (isNaN(price)) return null;
      return {
        name: name,
        code: barcode,
        price: price
      };
    }
  }
  return null;
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

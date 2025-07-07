const MAIN_SS_ID = '1x_f-IMzhYpUpuhV8jL-Ij6qyTIpOEqwWzJgSUrW9Ihk';   // цени 
const PPR_HISTORY_SHEET_ID = '14CAbMpgzss2KYF6hAiOqKBgdkJQFA8f4wE79bEgYe2k';

var processedFilesList = [];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index.html');
}

function loadReferencePage() {
  return HtmlService.createHtmlOutputFromFile('reference.html').getContent();
}

function loadinterfacePage() {
  return HtmlService.createHtmlOutputFromFile('interface.html').getContent();
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
  const root = DriveApp.getFolderById('1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG');
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
const folder = DriveApp.getFolderById("1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG"); // Родителската папка

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
  const folder = DriveApp.getFolderById("1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG"); // ID на родителската папка
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

  var parentFolderId = "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
  var templateFileId = "11khFtYxY39OA9UYSfDStMPmGbbv76fBx-2-ziea7h50";
  var revisionParentFolderId = "1Yo8oVkgYYmSR5z_cUFR7zdiRFJZkH3n5";
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
    try {
        var parentFolderId = "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
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
  
  try {
    var parentFolderId = "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
    var templateFileId = "1vWgz8j2wWHrP2CYTRCnsiv970kbOfBfMzR0cP5ogodQ";
    var revisionParentFolderId = "1Yo8oVkgYYmSR5z_cUFR7zdiRFJZkH3n5";
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
  try {
    var parentFolderId = "1PqVTfIpJKQRuxUcxwHVubfNOsvRTWzXG";
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

function savePPRData(storeName,dateString,tableData,pprNumber,note,reasonType){
  if(!pprNumber) throw new Error('❗ Моля, въведете номер на ППР.');

  /* 1. Проверяваме в история-таблицата */
  const histSh = SpreadsheetApp.openById(PPR_HISTORY_SHEET_ID).getSheetByName('ППР ИСТОРИЯ');
  const used   = histSh.getRange(2,1,histSh.getLastRow()-1,1).getValues().flat();
  if(used.includes(String(pprNumber).trim()))
      throw new Error('❌ ТОЗИ НОМЕР НА ППР ВЕЧЕ Е ИЗПОЛЗВАН!');

  /* 2. Генерираме PPR файла от шаблон */
  const TEMPLATE_ID            = '1KBeWbFlYDMXPoMxxz4H0YvfWLQh2ZdK4i_9iMGk9H4I';
  const DESTINATION_FOLDER_ID  = '1avn7paZvq3eHMdIMcH_PBF3sWA2tNM8l';
  const TARGET_SHEETS          = ['МЛЯКО ВЪНШНА СТОКА','МЛЯКО','АГНЕШКО','ГОВЕЖДО','МЛЕНИ','МЕСО','КОЛБАСИ'];
  const EMAIL_TO               = 'v.likov@dmc.farm, sklad.pld@dmc.farm, account@ovcharovo.com, m_margitina@abv.bg, kristiyan.stoynev@dmc.farm';

  const parent = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const sub    = parent.getFoldersByName(storeName).hasNext()
                 ? parent.getFoldersByName(storeName).next()
                 : parent.createFolder(storeName);

  const fileName   = `${pprNumber}_${dateString}_${storeName}`;
  const newFile    = DriveApp.getFileById(TEMPLATE_ID).makeCopy(fileName,sub);
  const ss         = SpreadsheetApp.openById(newFile.getId());

  /* записваме типа */
  const meta = ss.getSheetByName('МЕТА') || ss.insertSheet('МЕТА');
  meta.getRange('A1').setValue(reasonType||'');

  /* 3. Прехвърляме данните */
  const notFound=[];
  tableData.forEach(row=>{
    const [code,name,barcode,qtyStr] = row;
    const qty = parseFloat(qtyStr);
    if(!code || isNaN(qty)) return;

    let done=false;
    for(const shtName of TARGET_SHEETS){
      const sht = ss.getSheetByName(shtName);
      if(!sht) continue;
      const colA = sht.getRange('A:A').getValues().flat();
      const idx  = colA.findIndex(c=>String(c).trim()===String(code).trim());
      if(idx>-1){
        const cell = sht.getRange(idx+1,4);
        cell.setValue((parseFloat(cell.getValue())||0)+qty);
        done=true; break;
      }
    }
    if(!done) notFound.push([code,name,qty]);
  });

  if(notFound.length){
    let nf = ss.getSheetByName('НЕРАЗПОЗНАТИ АРТИКУЛИ');
    if(nf) ss.deleteSheet(nf);
    nf = ss.insertSheet('НЕРАЗПОЗНАТИ АРТИКУЛИ');
    nf.getRange(1,1,1,3).setValues([['Артикулен номер','Име','Количество']]);
    const grouped={};
    notFound.forEach(([c,n,q])=>{
      if(!grouped[c]) grouped[c]={n, q:0};
      grouped[c].q+=q;
    });
    const rows = Object.entries(grouped).map(([c,o])=>[c,o.n,o.q]);
    nf.getRange(2,1,rows.length,3).setValues(rows);
  }

  SpreadsheetApp.flush();

  /* 4. Добавяме записа в историята */
  histSh.appendRow([pprNumber,storeName,dateString,Session.getActiveUser().getEmail()]);

  /* 5. Изкарваме Excel + изпращаме мейл */
  const exportUrl = `https://docs.google.com/spreadsheets/d/${newFile.getId()}/export?format=xlsx`;
  const blob      = UrlFetchApp.fetch(exportUrl,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}})
                        .getBlob().setName(fileName+'.xlsx');

  sendViaSendGrid(
     EMAIL_TO,
     `ППР №${pprNumber} – ${storeName} (${dateString})`,
     `<p>✅ ППР №${pprNumber} е въведен.</p>
        <p><strong>Магазин:</strong> ${storeName}<br>
           <strong>Дата:</strong> ${dateString}</p>
        ${note?`<p><strong>Причина:</strong> ${note}</p>`:''}
        ${reasonType?`<p><strong>Тип ППР:</strong> ${reasonType}</p>`:''}
        <p><a href="${newFile.getUrl()}" target="_blank">Отвори файла в Google Таблици</a></p>`,
     blob
  );

  return '✅ Записът е направен и имейлът е изпратен.';
}

/**
 * Изпраща имейл чрез SendGrid API.
 * @param {string} to        - Списък получатели, разделени със запетаи.
 * @param {string} subject   - Тема на имейла.
 * @param {string} htmlBody  - HTML съдържание на имейла.
 * @param {Blob}   blob      - Прикачен файл за изпращане.
 */
function sendViaSendGrid(to, subject, htmlBody, blob) {
  var payload = {
    personalizations: [{ to: to.split(/,\s*/) .map(e => ({ email: e.trim() })) }],
    from: { email: 'noreply@yourdomain.com' },
    subject: subject,
    content: [{ type: 'text/html', value: htmlBody }],
    attachments: [{
      content: Utilities.base64Encode(blob.getBytes()),
      filename: blob.getName()
    }]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + PropertiesService.getScriptProperties().getProperty('SENDGRID_KEY')
    },
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch('https://api.sendgrid.com/v3/mail/send', options);
}

/* -------- Цена по код -------- */
function getPriceByCode(code){
  const sh = SpreadsheetApp.openById(MAIN_SS_ID).getSheetByName('Лист1');
  if(!sh) return null;
  const rows = sh.getRange('A:F').getValues();
  for(const r of rows){
    if(String(r[0]).trim()===String(code) || String(r[2]).trim()===String(code)){
      const price = parseFloat(String(r[5]).replace(/[^0-9.,]/g,'').replace(',','.'));
      return isNaN(price)?null:price;
    }
  }
  return null;
}

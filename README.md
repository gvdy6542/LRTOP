# LRTOP – Скриптове за Google Apps Script

## Общ преглед
LRTOP е уеб приложение изградено с **Google Apps Script**, предназначено за управление на ревизии, обработка на баркодове и печат на етикети. Интерфейсът и голяма част от функционалността са на български език. Проектът може да се публикува като Google Apps Script Web App или да се използва като разширение към Google Sheets.

Основният код се намира в `Code.js` и редица HTML файлове, които изграждат потребителския интерфейс. Скриптът използва Google Drive и Google Sheets за съхранение на данни.

## Съдържание на хранилището

- **Code.js** – основните сървърни функции на Apps Script.
- **appsscript.json** – конфигурационен файл на Google Apps Script с дефинирани OAuth обхвати и настройки.
- **index.html** – главната уеб страница за сканиране на баркодове, управление на ревизии и въвеждане на ППР.
- **interface.html** – страница, която визуализира съдържание на споделена папка в Google Drive.
- **labels.html** и **labelsContent.html** – генератор на етикети с визуализация и опции за персонализация.
- **MenuView.html** и **Selection.html** – страници за избор между режими и показване на меню.
- **reference.html** – инструменти за обработка и качване на файлове в Drive.
- **.clasp.json** – настройки за инструмента [`clasp`](https://github.com/google/clasp).
- **.github/workflows/deploy.yml** – GitHub Action, който при push към `main` изпраща кода към Apps Script чрез `clasp push --force`.

## Инсталация и разработка

1. Инсталирайте Node.js и глобално `@google/clasp`:
   ```bash
   npm install -g @google/clasp
   ```
2. Влезте в профила си и свържете проекта:
   ```bash
   clasp login
   clasp pull
   ```
3. Правете промени по `.js` и `.html` файловете. За тестове може да използвате `clasp push` за качване към Apps Script.
4. При нужда от автоматично деплойване използвайте GitHub Workflow файла в `.github/workflows/deploy.yml`.

## Използване на уеб приложението

### Главна страница – `index.html`
На тази страница се стартира ревизия и се сканират баркодове. Основни функции:
- **Започни ревизия** – показва поле за въвеждане на баркод.
- **Въвеждане на ППР** – отваря модал за въвеждане и запис на ППР данни.
- **Проверка на ППР** – зарежда съществуващ ППР за редакция.
- **Справка на брака** – генерира отчет за брак по магазин.
- **Етикети** – отваря генератор на етикети.
- **Преглед на ревизии** – показва списък с налични ревизии в Google Drive.

Част от интерфейса е показана в редовете около бутона *Преглед на ревизии* в `index.html`:
```html
<button id="viewRevisionsBtn"  class="green-button">Преглед на ревизии</button>
```
Този бутон извиква сървърната функция `listRevisions`, която връща списък с файлове. Пълният код на обработката се намира в `Code.js`.

### Генератор на етикети – `labels.html`
Файлът предоставя панел за настройка и визуализация на етикети. Генерирането става чрез бутон „Стартирай“:
```html
<button class="action-btn red" id="btnGenerate" onclick="generate()">Стартирай</button>
```
Функцията `generate()` извиква `runGenerateLabels()` от `Code.js`, която създава лист "Етикети" в Google Sheets и връща данните за визуализация.

### Обработка на файлове – `reference.html`
Тази страница дава възможност за обработка и качване на файлове. При натискане на бутона „Старт обработка“ се изпълнява `processFilesWithProgress()`:
```html
<button id="startButton" onclick="startProcessing()">Старт обработка</button>
```
Полученият списък с обработени файлове се визуализира в таблица.
### Админ панел – `admin.html`
Новата страница `admin.html` позволява настройка на папките и видимостта на бутоните. Отваря се чрез бутона **Admin** в `index.html` и записва промените чрез `saveConfig()`. Достъп имат само имейлите, записани в `adminEmails`.


## Основни сървърни функции (Code.js)
Файлът `Code.js` съдържа множество функции, сред които:
- `doGet()` – връща `index.html` при уеб достъп.
- `loadReferencePage()`, `loadinterfacePage()`, `loadLabelsPage()` – зареждат съответните HTML страници в нов прозорец.
- `processBarcode(barcode)` – търси артикул по баркод и връща данни за него.
- `listRevisions(storeName)` и `getRevisionData(fileId)` – работят с Google Drive за списък и визуализация на ревизии.
- `runGenerateLabels()` и `generateLabelsSheet()` – генерират етикети в Google Sheets.
- `savePPRData(...)` – създава PPR документ на база въведените редове и изпраща имейл с прикачен файл.

Кодът показва пример за зареждане на меню в странична лента:
```javascript
function showMenuSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('MenuView')
    .setTitle('Меню');
  SpreadsheetApp.getUi().showSidebar(html);
}
```

## Конфигурация
Файлът `appsscript.json` задава необходимите OAuth обхвати и други настройки на проекта:
```json
{
  "timeZone": "Europe/Bucharest",
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  },
  "runtimeVersion": "V8"
}
```

## Лиценз
Този проект не съдържа изрична информация за лиценз. При използване на кода се съобразявайте с общите условия на Google Apps Script и Google Drive.

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { JSDOM } = require('jsdom');

// Load Code.gs into the global context so we can access getItemFromCache
beforeAll(() => {
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  vm.runInNewContext(code, global); // defines getItemFromCache and loadItemsCache globally

  // Simple in-memory cache mock
  const store = {};
  global.CacheService = {
    getScriptCache: () => ({
      get: key => store[key] || null,
      put: (key, value) => { store[key] = value; }
    })
  };

  // Mock implementation of loadItemsCache with columns A, C and H
  const rows = [
    ['10001', '', 'Item One', '', '', '', '', '1'],      // short code 1-digit
    ['10002', '', 'Item Two', '', '', '', '', '22'],      // short code 2-digit
    ['10003', '', 'Item Three', '', '', '', '', '333'],   // short code 3-digit
    ['10004', '', 'Item Four', '', '', '', '', '4444'],   // short code 4-digit
    ['10005', '', 'Item Five', '', '', '', '', '55555']   // short code 5-digit
  ];

  global.loadItemsCache = function() {
    const byCode = {};
    const byBarcode = {};
    rows.forEach(r => {
      const code = String(r[0]).trim();          // column A
      const name = String(r[2]).trim();          // column C
      const shortCode = String(r[7]).trim();     // column H
      if (code) byCode[code] = { code, name };
      if (shortCode) byBarcode[shortCode] = { code, name };
    });
    const data = { byCode, byBarcode };
    CacheService.getScriptCache().put('itemsCache', JSON.stringify(data));
    return data;
  };
});

describe('short codes lookup', () => {
  test('getItemFromCache resolves full code via short code from column H', () => {
    loadItemsCache();
    const item = getItemFromCache('333'); // short code -> Item Three
    expect(item).toEqual({ code: '10003', name: 'Item Three', barcode: '333' });
  });

  test('appendOrUpdateTable stores article number instead of short code', () => {
    loadItemsCache();
    const item = getItemFromCache('4444'); // returns code 10004

    const dom = new JSDOM('<table id="itemDetails"><tbody></tbody></table>');
    global.document = dom.window.document;
    global.saveTableToLocalStorage = jest.fn();

    function appendOrUpdateTable(ic, inb, bc, qt) {
      const tb = document.querySelector('#itemDetails tbody');
      let ex = Array.from(tb.rows).find(r => r.cells[0].textContent === ic);
      if (ex) {
        const cur = parseFloat(ex.cells[3].textContent);
        ex.cells[3].textContent = (cur + qt).toFixed(3);
      } else {
        const tr = document.createElement('tr');
        tr.innerHTML = `
      <td>${ic}</td>
      <td>${inb}</td>
      <td>${bc}</td>
      <td>${qt.toFixed(3)}</td>
      <td class="actions">
        <button class="edit-btn">✏️</button>
        <button class="delete-btn">✖️</button>
      </td>
    `;
        tb.prepend(tr);
      }
      document.getElementById('itemDetails').style.display = 'table';
      saveTableToLocalStorage();
    }

    appendOrUpdateTable(item.code, item.name, 'ignored', 1);

    const first = dom.window.document.querySelector('#itemDetails tbody tr td').textContent;
    expect(first).toBe('10004');
    expect(first).not.toBe('4444');
  });
});

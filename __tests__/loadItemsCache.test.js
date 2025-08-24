const fs = require('fs');
const path = require('path');
const vm = require('vm');

function loadSandbox(rows) {
  const codePath = path.join(__dirname, '..', 'Code.gs');
  const code = fs.readFileSync(codePath, 'utf8');

  const cacheStore = {};
  const cache = {
    store: cacheStore,
    get(key) { return this.store[key] || null; },
    put(key, value) { this.store[key] = value; }
  };

  const sheet = {
    getRange() { return { getValues: () => rows }; }
  };

  const SpreadsheetApp = {
    openById() {
      return {
        getSheetByName() { return sheet; }
      };
    }
  };

  const CacheService = {
    getScriptCache() { return cache; }
  };

  const sandbox = { SpreadsheetApp, CacheService };
  vm.runInNewContext(code, sandbox);
  sandbox._cacheStore = cacheStore;
  return sandbox;
}

describe('loadItemsCache', () => {
  test('indexes by code, barcode and short codes', () => {
    const rows = [
      ['100', '', 'Item A', '', '', '', '123456', 'SC1, SC2'],
      ['200', '', 'Item B', '', '', '', '654321', 'SC3 SC4']
    ];
    const sb = loadSandbox(rows);
    const data = sb.loadItemsCache();

    expect(data.byCode['100']).toEqual({ code: '100', name: 'Item A', barcode: '123456' });
    expect(data.byBarcode['654321']).toEqual({ code: '200', name: 'Item B', barcode: '654321' });
    expect(data.byShortCode['SC1']).toEqual({ code: '100', name: 'Item A', barcode: '123456' });
    expect(data.byShortCode['SC4']).toEqual({ code: '200', name: 'Item B', barcode: '654321' });

    const cached = JSON.parse(sb._cacheStore['itemsCache']);
    expect(cached.byShortCode['SC2'].code).toBe('100');

    const item = sb.getItemFromCache('SC3');
    expect(item).toEqual({ code: '200', name: 'Item B', barcode: '654321' });
  });
});


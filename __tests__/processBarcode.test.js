const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

async function loadDom(mockItem) {
  const htmlPath = path.join(__dirname, '..', 'index.html');
  let html = fs.readFileSync(htmlPath, 'utf8');
  html = html.replace("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js\"></script>", '<script></script>');
  html = html.replace("<script><?!= include('common'); ?></script>", '<script></script>');
  const vc = new VirtualConsole();
  vc.on('error', () => {});
  const dom = new JSDOM(html, {
    runScripts: 'dangerously',
    resources: 'usable',
    virtualConsole: vc,
    beforeParse(window) {
      window.lastCacheArg = null;
      window.google = {
        script: {
          run: {
            withSuccessHandler(handler) {
              return {
                getItemFromCache(arg) {
                  window.lastCacheArg = arg;
                  handler(mockItem);
                }
              };
            }
          }
        }
      };
      window.sendActivity = () => {};
    }
  });
  await new Promise(resolve => dom.window.document.addEventListener('DOMContentLoaded', resolve));
  return dom;
}

describe('processBarcode weight barcodes', () => {
  test('uses extracted code and passes weight', async () => {
    const mockItem = { code: '301234', name: 'Item' };
    const dom = await loadDom(mockItem);
    const w = dom.window;
    const spy = jest.fn();
    w.updateItemDetails = spy;

    const barcode = '2801234012345';
    w.processBarcode(barcode);

    expect(w.lastCacheArg).toBe('301234');
    expect(spy).toHaveBeenCalledWith('301234', 'Item', barcode, 1.234);
  });
});

const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

async function loadDom(config) {
  const htmlPath = path.join(__dirname, '..', 'admin-panel.html');
  let html = fs.readFileSync(htmlPath, 'utf8');
  html = html.replace("<script><?!= include('common'); ?></script>", '<script></script>');
  const dom = new JSDOM(html, {
    runScripts: 'dangerously',
    resources: 'usable',
    beforeParse(window) {
      window.google = {
        script: {
          run: {
            withSuccessHandler(handler) {
              return {
                getConfig() { handler(config); },
                getClientStats() { handler([]); }
              };
            }
          }
        }
      };
      window.updateFolderLinks = () => {};
      window.sendActivity = () => {};
      window.refreshAllClients = () => {};
      window.setInterval = () => {};
    }
  });
  await new Promise(resolve => dom.window.document.addEventListener('DOMContentLoaded', resolve));
  return dom;
}

describe('loadConfig', () => {
  test('checkboxes reflect config values', async () => {
    const config = {
      showInterfaceButton: 'TRUE',
      showReferenceButton: 'False',
      showLabelsButton: ' true ',
      showPprButtons: true,
      showViewRevisionsBtn: false
    };
    const dom = await loadDom(config);
    const doc = dom.window.document;
    expect(doc.getElementById('showInterfaceButton').checked).toBe(true);
    expect(doc.getElementById('showReferenceButton').checked).toBe(false);
    expect(doc.getElementById('showLabelsButton').checked).toBe(false);
    expect(doc.getElementById('showPprButtons').checked).toBe(true);
    expect(doc.getElementById('showViewRevisionsBtn').checked).toBe(false);
  });
});

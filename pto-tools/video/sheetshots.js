const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const A = process.argv[2] + '/assets/';
  const b = await chromium.launch(); const p = await b.newPage({ deviceScaleFactor: 2, viewport: { width: 2400, height: 1400 } });
  for (const f of fs.readdirSync(A).filter(f => f.startsWith('sheet_') && f.endsWith('.html'))) {
    await p.goto('file://' + A + f);
    await (await p.$('#sheet')).screenshot({ path: A + f.replace('.html', '.png') });
  }
  await b.close();
})();

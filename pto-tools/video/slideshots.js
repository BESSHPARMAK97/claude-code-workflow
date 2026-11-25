const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const S = process.argv[2];
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
  for (const f of fs.readdirSync(S).filter(f => f.endsWith('.html')).sort()) {
    await p.goto('file://' + S + f); await p.waitForTimeout(100);
    await p.screenshot({ path: S + f.replace('.html', '.png') });
  }
  await b.close();
})();

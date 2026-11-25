const { chromium } = require('playwright');
(async () => {
  const A = process.argv[2] + '/assets/';
  const b = await chromium.launch();
  const ctx = await b.newContext({ acceptDownloads: true, viewport: { width: 1100, height: 760 }, deviceScaleFactor: 2 });
  const p = await ctx.newPage();
  await p.goto('file://' + require('path').resolve(__dirname, '../dist/Генератор_АОСР.html'));
  const hl = async (sel, on) => p.evaluate(([s, on]) => { document.querySelectorAll(s).forEach(e => { e.style.outline = on ? '4px solid #e8590c' : ''; e.style.outlineOffset = on ? '6px' : ''; e.style.borderRadius = on ? '6px' : ''; }); }, [sel, on]);
  await p.screenshot({ path: A + 'gen_empty.png' });
  await hl('#sample', true); await p.screenshot({ path: A + 'gen_sample.png' }); await hl('#sample', false);
  await hl('#drop', true); await p.evaluate(() => document.getElementById('drop').classList.add('over'));
  await p.screenshot({ path: A + 'gen_drop.png' });
  await p.evaluate(() => document.getElementById('drop').classList.remove('over')); await hl('#drop', false);
  // успешная загрузка
  let dlP = p.waitForEvent('download');
  await p.setInputFiles('#file', A + 'реестр_работ.xlsx');
  let d = await dlP; await d.saveAs(A + 'Комплект_АОСР.docx');
  await p.waitForTimeout(400);
  await p.screenshot({ path: A + 'gen_ok.png', fullPage: true });
  await hl('#extra button', true); await p.screenshot({ path: A + 'gen_buttons.png', fullPage: true }); await hl('#extra button', false);
  // ошибки
  dlP = p.waitForEvent('download', { timeout: 5000 }).catch(() => null);
  await p.setInputFiles('#file', A + 'реестр_с_ошибками.xlsx'); await dlP; await p.waitForTimeout(400);
  await p.screenshot({ path: A + 'gen_err.png', fullPage: true });
  console.log(await p.evaluate(() => document.getElementById('report').innerText));
  await b.close();
})();

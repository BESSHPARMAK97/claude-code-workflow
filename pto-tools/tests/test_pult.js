// ПТО-пульт: распознавание файлов, вкладки, передача реестра в журнал без .json.
// Запуск из pto-tools: node tests/test_pult.js
const { chromium } = require('playwright');
const path = require('path');
const ROOT = path.resolve(__dirname, '..');
(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 860 } })).newPage();
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  let bad = 0; const check = (ok, msg) => { if (!ok) bad++; console.log(`${ok ? 'OK  ' : 'FAIL'} ${msg}`); };
  await p.goto('file://' + path.join(ROOT, 'dist', 'ПТО_пульт.html')); await p.waitForTimeout(1500);
  const tabs = await p.$$eval('nav button', bs => bs.map(x => x.textContent));
  check(tabs.join('|') === 'Акты и журналы|Журнал работ|Формы 2 и 3', 'вкладки: ' + tabs.join(', '));
  const status = () => p.textContent('#status'), active = () => p.$eval('nav button.active', x => x.textContent);
  const frame = async i => (await p.$$('iframe'))[i].contentFrame();

  let dl = p.waitForEvent('download', { timeout: 10000 }).catch(() => null);
  await p.setInputFiles('#file', path.join(ROOT, 'forms_2_3/план_факт.xlsx'));
  let d = await dl; await p.waitForTimeout(400);
  check(await active() === 'Формы 2 и 3' && /план-факт/.test(await status()) && d && /Формы_2_3_2026-04/.test(d.suggestedFilename()),
    'план-факт → вкладка «Формы 2 и 3», формы скачаны');

  dl = p.waitForEvent('download', { timeout: 10000 }).catch(() => null);
  await p.setInputFiles('#file', path.join(ROOT, 'tests/fixtures/aosr_с_датами.xlsx'));
  d = await dl; await p.waitForTimeout(400);
  check(await active() === 'Акты и журналы' && /реестр работ/.test(await status()) && d && d.suggestedFilename() === 'Комплект_АОСР.docx',
    'реестр → вкладка «Акты и журналы», комплект скачан');

  const aosr = await frame(0);
  check(await aosr.isVisible('#jr-push'), 'в пульте есть кнопка «Передать в журнал работ»');
  await aosr.click('#jr-push'); await p.waitForTimeout(1000);
  check(/Журнал создан из реестра: ИТР — 3, актов — 6, записей о работах — 6/.test(await aosr.textContent('#report')), 'журнал создан из реестра');
  check(/Журнал работ обновлён/.test(await status()) && await p.$eval('nav button:nth-child(2)', x => !!x.querySelector('.dot')), 'пульт отметил вкладку журнала');

  await p.click('nav button:nth-child(2)'); await p.waitForTimeout(800);
  const jst = await (await frame(1)).evaluate(() => JSON.parse(localStorage.getItem('zpr-journal-v2')));
  check(jst && jst.s3.length === 6 && jst.s4.length === 6, 'журнал во вкладке видит данные реестра');

  await p.click('nav button:nth-child(1)'); await aosr.click('#jr-push'); await p.waitForTimeout(800);
  check(/Журнал дополнен: ИТР \+0, акты \+0, записи о работах \+0/.test(await aosr.textContent('#report')), 'повторная передача не дублирует записи');

  await p.setInputFiles('#file', path.join(ROOT, 'm29/материалы.xlsx')); await p.waitForTimeout(3000);
  check(/не распознан/.test(await status()), 'чужой файл: понятное сообщение');
  if (errs.length) { bad++; console.log('Ошибки JS:', errs); }
  await b.close(); process.exit(bad ? 1 : 0);
})();

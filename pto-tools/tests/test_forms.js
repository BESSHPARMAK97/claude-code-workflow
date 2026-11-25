// Прогон генератора форм 2 и 3 в браузере.  Запуск: node tests/test_forms.js
const { chromium } = require('playwright');
const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const CASES = [
  { file: 'forms_2_3/план_факт.xlsx', has: ['за месяц без НДС — 36 662 500,00 тг', 'с НДС — 42 528 500,00 тг', '50 734 500,00 тг из 179 745 000,00 тг'], download: true },
  { file: 'forms_2_3/пример_с_ошибками.xlsx', has: ['позиции № 12 нет в смете', 'перевыполнение 50,000'], download: false },
];
(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ acceptDownloads: true })).newPage();
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.goto('file://' + path.join(ROOT, 'dist', 'Формы_2_и_3.html'));
  let bad = 0;
  for (const c of CASES) {
    const dl = p.waitForEvent('download', { timeout: 8000 }).catch(() => null);
    await p.setInputFiles('#file', path.join(ROOT, c.file));
    const d = await dl; await p.waitForTimeout(300);
    const rep = await p.evaluate(() => document.getElementById('report').innerText);
    const miss = c.has.filter(s => !rep.includes(s));
    const ok = miss.length === 0 && !!d === c.download;
    if (!ok) bad++;
    console.log(`${ok ? 'OK  ' : 'FAIL'} ${c.file}${miss.length ? ': нет в отчёте: ' + miss.join(' | ') : ''}${!!d !== c.download ? ' (скачивание не как ожидалось)' : ''}`);
  }
  if (errs.length) { bad++; console.log('Ошибки JS:', errs); }
  await b.close();
  process.exit(bad ? 1 : 0);
})();

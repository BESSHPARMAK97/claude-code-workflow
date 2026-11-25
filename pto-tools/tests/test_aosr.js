// Прогон генератора АОСР в браузере на эталонных реестрах.  Запуск: node tests/test_aosr.js
const { chromium } = require('playwright');
const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const CASES = [
  { file: 'aosr/реестр_работ.xlsx', acts: 5, failed: 0 },
  { file: 'tests/fixtures/aosr_без_представителя.xlsx', acts: 4, failed: 1, has: ['не указан «Представитель СМО»', 'представитель СМО не из'] },
  { file: 'tests/fixtures/aosr_ошибки_дат_и_материалов.xlsx', acts: 2, failed: 3, has: ['истёк до начала работ', '«M99» не найден', 'позже окончания'] },
  { file: 'tests/fixtures/aosr_старый_формат.xlsx', acts: 10, failed: 1 },
];
(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ acceptDownloads: true })).newPage();
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.goto('file://' + path.join(ROOT, 'dist', 'Генератор_АОСР.html'));
  let bad = 0;
  for (const c of CASES) {
    const dl = p.waitForEvent('download', { timeout: 15000 }).catch(() => null);
    await p.setInputFiles('#file', path.join(ROOT, c.file));
    const d = await dl; await p.waitForTimeout(200);
    const rep = await p.evaluate(() => document.getElementById('report').innerText);
    const acts = +(rep.match(/Сформировано актов: (\d+)/) || [])[1];
    const failed = +(rep.match(/Пропущено из-за ошибок: (\d+)/) || [])[1];
    const miss = (c.has || []).filter(s => !rep.includes(s));
    const ok = acts === c.acts && failed === c.failed && miss.length === 0 && (c.acts === 0 || d);
    if (!ok) bad++;
    console.log(`${ok ? 'OK  ' : 'FAIL'} ${c.file}: актов ${acts} (ждали ${c.acts}), ошибок ${failed} (ждали ${c.failed})${miss.length ? ', нет в отчёте: ' + miss.join(' | ') : ''}`);
  }
  if (errs.length) { bad++; console.log('Ошибки JS:', errs); }
  await b.close();
  process.exit(bad ? 1 : 0);
})();

// ПТО-пульт: реестр — справочник объекта; журнал — быстрая запись и список работ; акты — по работам журнала.
// Запуск из pto-tools: node tests/test_pult.js
const { chromium } = require('playwright');
const path = require('path'); const fs = require('fs'); const os = require('os');
const PizZip = require('pizzip');
const ROOT = path.resolve(__dirname, '..');
const PULT = 'file://' + path.join(ROOT, 'dist', 'ПТО_пульт.html');
const REG = path.join(ROOT, 'tests/fixtures/aosr_с_датами.xlsx');
const iso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const ago = n => { const d = new Date(); d.setDate(d.getDate() - n); return iso(d); };
(async () => {
  const b = await chromium.launch();
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pult-'));
  let bad = 0; const check = (ok, msg) => { if (!ok) bad++; console.log(`${ok ? 'OK  ' : 'FAIL'} ${msg}`); };
  const errs = [];
  const open = async () => {
    const ctx = await b.newContext({ acceptDownloads: true, viewport: { width: 1400, height: 900 } });
    const p = await ctx.newPage(); p.on('pageerror', e => errs.push(String(e)));
    await p.goto(PULT); await p.waitForTimeout(1500);
    return { ctx, p };
  };
  const frames = async p => Promise.all((await p.$$('iframe')).map(f => f.contentFrame()));
  const jstate = f => f.evaluate(() => JSON.parse(localStorage.getItem('zpr-journal-v2')));

  let { ctx, p } = await open();
  const tabs = await p.$$eval('nav button', bs => bs.map(x => x.textContent));
  check(tabs.join('|') === 'Акты и журналы|Журнал работ|Формы 2 и 3', 'вкладки: ' + tabs.join(', '));
  const status = () => p.textContent('#status'), active = () => p.$eval('nav button.active', x => x.textContent);

  let dl = p.waitForEvent('download', { timeout: 10000 }).catch(() => null);
  await p.setInputFiles('#file', path.join(ROOT, 'forms_2_3/план_факт.xlsx'));
  let d = await dl; await p.waitForTimeout(400);
  check(await active() === 'Формы 2 и 3' && d && /Формы_2_3_2026-04/.test(d.suggestedFilename()), 'план-факт → вкладка «Формы 2 и 3», формы скачаны');

  // ---------- реестр: акты по его работам + справочник объекта; журнал сам не меняется
  dl = p.waitForEvent('download', { timeout: 15000 }).catch(() => null);
  await p.setInputFiles('#file', REG); d = await dl; await p.waitForTimeout(800);
  let [a, j] = await frames(p);
  check(await active() === 'Акты и журналы' && d && d.suggestedFilename() === 'Комплект_АОСР.docx', 'реестр → «Акты и журналы», комплект скачан');
  check(/запомнен как справочник объекта: людей — 6, партий материалов — 20/.test(await a.textContent('#report')), 'реестр запомнен как справочник (люди, материалы)');
  check(await jstate(a) === null, 'журнал сам не заполняется из реестра');

  // ---------- журнал: быстрая запись
  await p.click('nav button:nth-child(2)'); await p.waitForTimeout(900);
  [a, j] = await frames(p);
  check(await j.$eval('nav.tabs [aria-selected="true"]', x => x.dataset.tab) === 's4' && await j.isVisible('#quick'), 'журнал открывается на разделе 4 со строкой быстрой записи');
  const add = async (date, work, floor, vol, unit, enter) => {
    await j.fill('#q-date', date); await j.fill('#q-work', work); await j.fill('#q-floor', floor); await j.fill('#q-vol', vol);
    if (unit) await j.selectOption('#q-unit', unit);
    if (enter) await j.press('#q-vol', 'Enter'); else await j.click('#qAdd');
    await p.waitForTimeout(250);
  };
  await add('2026-03-10', 'Армирование плиты перекрытия', '2', '3,2', 'т', true);
  await add('2026-03-11', 'армирование  плиты перекрытия', '2', '2,8', '');
  await add('2026-03-12', 'Бетонирование плиты перекрытия', '2', '45', 'м³');
  await add('2026-03-13', 'Армирование плиты перекрытия', '3', '3,0', '');
  await p.waitForTimeout(500);
  let st = await jstate(j);
  check(st.s4.length === 4 && st.s4.every(r => !r._ex) && st.works.length === 2 && !st.s3.length, 'быстрая запись: 4 записи, 2 работы в списке, примеры убраны');
  check(st.s4.filter(r => /^Армирование/.test(r.work)).every(r => r.unit === 'т' && r.wid === st.s4[0].wid), 'та же работа (без учёта регистра) — одна работа, единица запомнилась');
  check(/Отставание \d+ дн/.test(await j.textContent('.lag')) && await j.$eval('.lag', x => x.classList.contains('bad')), 'отставание больше 5 дней — красная плашка');
  await add(ago(0), 'Армирование плиты перекрытия', '4', '1', '');
  await add(ago(3), 'Бетонирование плиты перекрытия', '3', '40', '');
  const lagText = await j.textContent('.lag'), ru = s => s.split('-').reverse().join('.');
  check(!(await j.$eval('.lag', x => x.classList.contains('bad'))) && lagText.includes(ru(ago(1))) && !lagText.includes(ru(ago(3))),
    'плашка: дни без записей за последнюю неделю (без дня, где запись есть)');
  await j.click(`[data-none="${ago(1)}"]`); await p.waitForTimeout(500);
  st = await jstate(j);
  check(st.s4.some(r => r.date === ago(1) && r.work === 'Работы не производились' && !r.wid), '«работ не было» — запись «Работы не производились»');
  check(/Дописать оси, отметки: 6 зап\./.test(await j.textContent('#checks')), 'сводка «что дописать» вместо замечаний к каждой строке');
  // бетон: графы лаборатории отдельно, напоминание о сроках испытаний
  const lab = await j.textContent('.lag.lab').catch(() => '');
  check(/Бетонирование плиты перекрытия, 2 этаж — 7 сут \(с 19\.03\.2026\)/.test(lab) && /28 сут \(с 09\.04\.2026\)/.test(lab) && !/3 этаж/.test(lab),
    'бетон: напоминание о прочности 7 и 28 суток (только где срок наступил)');
  const conc = st.s4.find(r => r.date === '2026-03-12');
  await j.click(`details[data-sid="${conc._id}"] summary`);
  const det = await j.textContent(`details[data-sid="${conc._id}"]`);
  check(/Журнал бетонных работ — заполняет мастер/.test(det) && /Заполняет лаборатория/.test(det) && det.indexOf('Подвижность') < det.indexOf('Заполняет лаборатория') && det.indexOf('Прочность 28') > det.indexOf('Заполняет лаборатория'),
    'бетон: «Подробнее» — блоки «мастер» (подвижность, t смеси) и «лаборатория» (образцы, прочность)');
  await j.fill(`[data-j="bj"][data-id="${conc._id}"][data-k="r7"]`, '21,5'); await j.fill(`[data-j="bj"][data-id="${conc._id}"][data-k="r28"]`, '31'); await p.waitForTimeout(400);
  await j.click('nav.tabs [data-tab="s4"]'); await p.waitForTimeout(300);
  check(!(await j.$('.lag.lab')), 'прочность внесена — напоминание исчезло');

  // печать за период: этаж — в графе «Место»
  await j.evaluate(() => { window.print = () => {}; });
  await j.fill('#pFrom', '2026-03-10'); await j.fill('#pTo', '2026-03-12'); await j.click('#bPrintPeriod');
  const pr = await j.$eval('#printArea', x => ({ rows: x.querySelectorAll('tbody tr').length, text: x.textContent }));
  check(pr.rows === 3 && /2 этаж/.test(pr.text) && /с 10\.03\.2026 по 12\.03\.2026/.test(pr.text), 'печать раздела 4 за период: 3 записи, этаж в графе «Место»');

  // ---------- работы и акты
  await j.click('nav.tabs [data-tab="wk"]'); await p.waitForTimeout(300);
  const sels = await j.$$eval('select[data-k="actsel"]', xs => xs.map(x => x.dataset.w));
  await j.selectOption(`select[data-w="${sels[0]}"][data-k="actsel"]`, 'aosr'); await p.waitForTimeout(300);
  await j.selectOption(`select[data-w="${sels[1]}"][data-k="actsel"]`, 'ppk'); await p.waitForTimeout(300);
  const ppl = await j.$$eval('#peopleList option', xs => xs.map(x => x.value));
  check(ppl.length === 6, 'в журнале — люди из реестра (для представителей в акте)');
  const ids = sels;
  const about = await j.textContent('details.about');
  check(/Когда оформлять/.test(about) && /до того, как их нагружают/.test(about) && /исполнительную геодезическую схему/.test(about), 'описание: когда оформлять акт скрытых работ и акт промежуточной приёмки');
  await j.click(`details[data-sid="w-${ids[1]}"] summary`);
  check(/до её нагружения и последующих работ/.test(await j.textContent(`details[data-sid="w-${ids[1]}"]`)), 'памятка в данных акта промежуточной приёмки');
  await j.click(`details[data-sid="w-${ids[1]}"] summary`);
  check(!(await j.$(`[data-w="${ids[1]}"][data-k="materials"]`)) && !!(await j.$(`[data-w="${ids[0]}"][data-k="materials"]`)), 'промежуточная приёмка: без материалов и доп. доказательств (их нет в форме)');
  for (const id of ids) {
    await j.click(`details[data-sid="w-${id}"] summary`);
    await j.fill(`[data-w="${id}"][data-k="drawings"]`, 'КЖ-1, листы 3–5');
    if (id === ids[0]) await j.fill(`[data-w="${id}"][data-k="materials"]`, 'M01');
    await j.fill(`[data-w="${id}"][data-k="cRep"]`, ppl[0]);
    await j.fill(`[data-w="${id}"][data-k="tRep"]`, ppl.find(x => /надзор/i.test(x)) || ppl[1]);
    if (id === ids[0]) await j.fill(`[data-w="${id}"][data-k="evidence"]`, 'исполнительная схема');
  }
  await j.fill(`[data-w="${ids[0]}"][data-k="no"]`, 'АОСР-101'); await j.press(`[data-w="${ids[0]}"][data-k="no"]`, 'Tab'); await p.waitForTimeout(300);
  await j.fill(`[data-w="${ids[1]}"][data-k="no"]`, 'АОСР-102'); await j.press(`[data-w="${ids[1]}"][data-k="no"]`, 'Tab'); await p.waitForTimeout(300);
  if (!(await j.isVisible(`[data-w="${ids[1]}"][data-k="next"]`))) await j.click(`details[data-sid="w-${ids[1]}"] summary`);
  await j.fill(`[data-w="${ids[1]}"][data-k="next"]`, 'устройство стяжки'); await p.waitForTimeout(500);
  // акт промежуточной приёмки: входящие акты скрытых работ подставляются сами (по этажам и датам)
  if (!(await j.isVisible(`[data-w="${ids[1]}"][data-k="incl"]`))) await j.click(`details[data-sid="w-${ids[1]}"] summary`);
  check(await j.getAttribute(`[data-w="${ids[1]}"][data-k="incl"]`, 'placeholder') === 'АОСР-101' && !(await j.$(`[data-w="${ids[0]}"][data-k="incl"]`)),
    'промежуточная приёмка: подсказка входящих актов скрытых работ (АОСР-101), у АОСР такого поля нет');
  await j.fill(`[data-w="${ids[1]}"][data-k="scheme"]`, 'ИС-12'); await j.fill(`[data-w="${ids[1]}"][data-k="tests"]`, 'протокол № 45 от 09.04.2026'); await p.waitForTimeout(400);
  st = await jstate(j);
  const arm = st.works.find(w => /^Армирование/.test(w.name));
  check(st.s3.some(x => x.no === 'АОСР-101' && /этажи 2, 3, 4/.test(x.name)) && st.s4.filter(r => r.wid === arm._id).every(r => r.aosr === 'АОСР-101'),
    'отметка «Акт»: акт в разделе 3 (с этажами), № акта — в записях');
  check(st.s3.some(x => x.no === 'АОСР-102' && /промежуточная приёмка/.test(x.note)), 'раздел 3: акт промежуточной приёмки помечен в примечании');
  dl = p.waitForEvent('download', { timeout: 15000 }).catch(() => null);
  await j.click('#bActs'); d = await dl; await p.waitForTimeout(800);
  [a, j] = await frames(p);
  const rep = await a.textContent('#report');
  check(await active() === 'Акты и журналы' && d && d.suggestedFilename() === 'Комплект_АОСР.docx' && /Сформировано актов: 2/.test(rep) && /Пропущено из-за ошибок: 0/.test(rep),
    '«Сформировать акты» в журнале → 2 акта по работам журнала');
  const kp = path.join(tmp, 'kit.docx'); if (d) await d.saveAs(kp);
  const xml = d ? new PizZip(fs.readFileSync(kp)).file('word/document.xml').asText().replace(/<[^>]+>/g, '') : '';
  check(/промежуточной приёмки ответственных конструкций/.test(xml) && /Конструкции принимаются/.test(xml) && /Приложение Е/.test(xml) && /промежуточной приёмки ответственных конструкций — 1/.test(xml),
    'комплект: акт промежуточной приёмки по форме (приложение Е) и счёт на титуле');
  check(/Приложения:/.test(xml) && /Исполнительная геодезическая схема № ИС-12/.test(xml) && /Акты освидетельствования скрытых работ № АОСР-101/.test(xml) && /Протоколы испытаний: протокол № 45/.test(xml)
    && (xml.match(/Приложения:/g) || []).length === 1, 'акт промежуточной приёмки: приложения — схема, входящие АОСР, протоколы (в АОСР блока нет)');
  check(/Из них актов промежуточной приёмки ответственных конструкций: 1/.test(rep), 'отчёт: сколько актов промежуточной приёмки');
  check(xml.includes('АОСР-101') && /«10» марта 2026/.test(xml) && xml.includes(`«${ago(0).slice(8)}»`), 'в акте даты работ — из записей журнала (первая и последняя)');

  // ссылки акт ⇄ запись
  await p.click('nav button:nth-child(2)'); await p.waitForTimeout(800);
  [a, j] = await frames(p);
  await j.click('nav.tabs [data-tab="s3"]'); await p.waitForTimeout(300);
  const lnks = await j.$$('tbody .lnk[data-goto="s4"]');
  if (lnks.length) { await lnks[0].click(); await p.waitForTimeout(300); }
  check(lnks.length === 6 && await j.$eval('nav.tabs [aria-selected="true"]', x => x.dataset.tab) === 's4' && !!(await j.$('tbody tr.hl')), 'ссылки «акт ⇄ запись» работают');

  await p.setInputFiles('#file', path.join(ROOT, 'm29/материалы.xlsx')); await p.waitForTimeout(3000);
  check(/не распознан/.test(await status()), 'чужой файл: понятное сообщение');
  await ctx.close();

  // ---------- акты из журнала без реестра — понятная подсказка
  ({ ctx, p } = await open());
  await p.click('nav button:nth-child(2)'); await p.waitForTimeout(800);
  [a, j] = await frames(p);
  await j.fill('#q-work', 'Монтаж закладных'); await j.click('#qAdd'); await p.waitForTimeout(300);
  await j.click('nav.tabs [data-tab="wk"]'); await p.waitForTimeout(300);
  await j.selectOption('select[data-k="actsel"]', 'aosr'); await p.waitForTimeout(300);
  await j.click('#bActs'); await p.waitForTimeout(800);
  [a] = await frames(p);
  check(/загрузите один раз реестр Excel объекта/.test(await a.textContent('#report')), 'без реестра: подсказка загрузить реестр один раз');
  await ctx.close();

  // ---------- журнал на телефоне: карточки записей, без горизонтальной прокрутки
  {
    const { devices } = require('playwright');
    const c2 = await b.newContext({ ...devices['iPhone 13'] });
    const ph = await c2.newPage(); ph.on('pageerror', e => errs.push(String(e)));
    await ph.goto('file://' + path.resolve(ROOT, '../templates/zhurnal-proizvodstva-rabot/zhurnal-proizvodstva-rabot.html')); await ph.waitForTimeout(400);
    await ph.fill('#q-work', 'Армирование плиты перекрытия'); await ph.fill('#q-floor', '5'); await ph.fill('#q-vol', '3'); await ph.click('#qAdd'); await ph.waitForTimeout(300);
    const m = await ph.evaluate(() => ({ sw: document.documentElement.scrollWidth, w: innerWidth,
      disp: getComputedStyle(document.querySelector('tr.rec')).display, head: getComputedStyle(document.querySelector('.s4t thead')).display,
      fs: getComputedStyle(document.getElementById('q-work')).fontSize }));
    check(m.sw <= m.w && m.disp === 'grid' && m.head === 'none' && m.fs === '16px', `телефон: запись — карточка, без прокрутки вбок, поля 16px (${m.sw}/${m.w}, ${m.disp}, ${m.fs})`);
    await c2.close();
  }

  if (errs.length) { bad++; console.log('Ошибки JS:', errs); }
  await b.close(); process.exit(bad ? 1 : 0);
})();

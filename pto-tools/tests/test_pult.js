// ПТО-пульт: окно объекта (ввод данных один раз) → акты и журнал работ; распознавание файлов; защита журнала, который вели вручную.
// Запуск из pto-tools: node tests/test_pult.js
const { chromium } = require('playwright');
const path = require('path'); const fs = require('fs'); const os = require('os');
const ROOT = path.resolve(__dirname, '..');
const PULT = 'file://' + path.join(ROOT, 'dist', 'ПТО_пульт.html');
const REG = path.join(ROOT, 'tests/fixtures/aosr_с_датами.xlsx');
(async () => {
  const b = await chromium.launch();
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pult-'));
  let bad = 0; const check = (ok, msg) => { if (!ok) bad++; console.log(`${ok ? 'OK  ' : 'FAIL'} ${msg}`); };
  const errs = [];
  const open = async () => {
    const ctx = await b.newContext({ acceptDownloads: true, viewport: { width: 1400, height: 900 } });
    const p = await ctx.newPage();
    p.on('pageerror', e => errs.push(String(e))); p.on('dialog', d => d.accept());
    return { ctx, p };
  };
  const frames = async p => Promise.all((await p.$$('iframe')).map(f => f.contentFrame()));
  const journal = f => f.evaluate(() => JSON.parse(localStorage.getItem('zpr-journal-v2')));
  const tabsText = f => f.textContent('#ok-tabs');

  // ---------- 1. новый браузер: окно, загрузка реестра, акты, журнал из окна
  let { ctx, p } = await open();
  await p.goto(PULT); await p.waitForTimeout(1500);
  const tabs = await p.$$eval('nav button', bs => bs.map(x => x.textContent));
  check(tabs.join('|') === 'Окно объекта|Журнал работ|Формы 2 и 3', 'вкладки: ' + tabs.join(', '));
  let [ok, jf] = await frames(p);
  check(!!(await ok.$('#okno')) && /С чего начать/.test(await ok.textContent('#ok-check')), 'пустое окно объекта: подсказка, с чего начать');
  const status = () => p.textContent('#status'), active = () => p.$eval('nav button.active', x => x.textContent);

  let dl = p.waitForEvent('download', { timeout: 10000 }).catch(() => null);
  await p.setInputFiles('#file', path.join(ROOT, 'forms_2_3/план_факт.xlsx'));
  let d = await dl; await p.waitForTimeout(400);
  check(await active() === 'Формы 2 и 3' && d && /Формы_2_3_2026-04/.test(d.suggestedFilename()), 'план-факт → вкладка «Формы 2 и 3», формы скачаны');

  await p.setInputFiles('#file', REG); await p.waitForTimeout(2000);
  check(await active() === 'Окно объекта' && /реестр работ/.test(await status()), 'реестр → вкладка «Окно объекта»');
  check(/Работы и акты 6/.test(await tabsText(ok)) && /Поступления 20/.test(await tabsText(ok)) && /Записи журнала 6/.test(await tabsText(ok)),
    'реестр загружен в окно: 6 работ, 20 партий, 6 записей журнала (по датам начала)');
  check(/ошибок нет/.test(await ok.textContent('#ok-check')), 'проверка окна: ошибок нет');

  let st = await journal(ok);
  check(st && st.meta.source === 'okno' && st.s3.length === 6 && st.s4.length === 6, 'журнал собран из окна: 6 актов, 6 записей');
  check(st && st.s3.map(a => a.no).sort().join() === st.s4.map(r => r.aosr).sort().join(), 'у каждой записи раздела 4 — номер её акта');

  dl = p.waitForEvent('download', { timeout: 15000 }).catch(() => null);
  await ok.click('#ok-acts'); d = await dl; await p.waitForTimeout(500);
  check(d && d.suggestedFilename() === 'Комплект_АОСР.docx' && /Сформировано актов: 5/.test(await ok.textContent('#report')), '«Сформировать акты» → комплект Word из данных окна');

  // правка в окне → журнал
  await ok.click('#ok-tabs [data-tab="days"]');
  const first = await ok.$('#ok-main tr[data-row]'); const did = await first.getAttribute('data-row');
  await ok.fill(`[data-p="d|${did}|weather"]`, 'Ясно, +5'); await ok.fill(`[data-p="d|${did}|vol"]`, '12,5'); await p.waitForTimeout(1200);
  st = await journal(ok);
  check(st.s4.find(r => r._id === did).weather === 'Ясно, +5' && st.s4.find(r => r._id === did).vol === '12,5', 'погода и объём из окна — в разделе 4 журнала');

  await ok.selectOption(`[data-p="d|${did}|jr"]`, 'Бетонные работы').catch(async () => {
    await ok.click(`tr.sub details[data-open="${did}"] summary`); await ok.selectOption(`[data-p="d|${did}|jr"]`, 'Бетонные работы');
  });
  await p.waitForTimeout(400);
  if (!(await ok.isVisible(`[data-p="sp|bj|${did}|cls"]`))) await ok.click(`details[data-open="${did}"] summary`);
  await ok.fill(`[data-p="sp|bj|${did}|cls"]`, 'C25/30, W6, F150'); await p.waitForTimeout(1200);
  st = await journal(ok);
  check(st.s4.find(r => r._id === did).jr === 'Бетонные работы' && st.bj[did] && st.bj[did].cls === 'C25/30, W6, F150', 'данные спецжурнала из окна — в журнале бетонных работ');

  await ok.click('#ok-main [data-add="d"]'); await p.waitForTimeout(1200);
  st = await journal(ok);
  check(st.s4.length === 7, 'запись, добавленная в окне, — в журнале (7)');

  // удалённая автозапись не возвращается
  const rows = await ok.$$('#ok-main tr[data-row]'); const last = rows[rows.length - 2];
  const lid = await last.getAttribute('data-row');
  const del = await last.$('[data-del]'); await del.click(); await del.click(); await p.waitForTimeout(1200);
  check(!/Записи журнала 7/.test(await tabsText(ok)) && /Записи журнала 6/.test(await tabsText(ok)) && !(await journal(ok)).s4.some(r => r._id === lid), 'удалённая запись не создаётся снова');

  // журнал во вкладке: только просмотр, ссылки работают
  await p.click('nav button:nth-child(2)'); await p.waitForTimeout(900);
  [ok, jf] = await frames(p);
  check(/собирается из «Окна объекта»/.test(await jf.textContent('#lockBar')), 'журнал: плашка «собирается из окна объекта»');
  await jf.click('nav.tabs [data-tab="s4"]'); await p.waitForTimeout(300);
  const dis = await jf.$$eval('#main textarea, #main select, #main input', xs => [xs.filter(x => x.disabled).length, xs.length]);
  check(dis[1] > 0 && dis[0] === dis[1] && !(await jf.isVisible('#bAdd')) && !(await jf.isVisible('#bOpen')), 'журнал: поля закрыты для правки, нет «Добавить» и «Открыть файл»');
  await jf.click('nav.tabs [data-tab="s3"]'); await p.waitForTimeout(300);
  const lnks = await jf.$$('tbody .lnk[data-goto="s4"]');
  if (lnks.length) { await lnks[0].click(); await p.waitForTimeout(300); }
  check(lnks.length >= 5 && await jf.$eval('nav.tabs [aria-selected="true"]', x => x.dataset.tab) === 's4' && !!(await jf.$('tbody tr.hl')), 'ссылки «акт ⇄ запись» работают');

  // повторная загрузка того же реестра: записи и их данные на месте
  await p.click('nav button:nth-child(1)');
  await p.setInputFiles('#file', REG); await p.waitForTimeout(2000);
  st = await journal(ok);
  check(/Записи журнала 6/.test(await tabsText(ok)) && st.s4.find(r => r._id === did) && st.s4.find(r => r._id === did).weather === 'Ясно, +5' && st.bj[did],
    'повторная загрузка реестра: записи не дублируются, погода и спецжурнал сохранены');

  // Excel: выгрузка и обратная загрузка
  dl = p.waitForEvent('download'); await ok.click('#ok-xlsout'); const xd = await dl; const xp = path.join(tmp, 'reg.xlsx'); await xd.saveAs(xp);
  await ok.setInputFiles('#ok-xls', xp); await p.waitForTimeout(2000);
  check(/Работы и акты 6/.test(await tabsText(ok)) && /Записи журнала 6/.test(await tabsText(ok)) && /ошибок нет/.test(await ok.textContent('#ok-check')),
    'выгрузка в Excel и загрузка обратно: те же 6 работ, без ошибок');

  // файл объекта
  dl = p.waitForEvent('download'); await ok.click('#ok-save'); const od = await dl; const op = path.join(tmp, 'obj.json'); await od.saveAs(op);
  await ok.click('#ok-clear'); await p.waitForTimeout(800);
  check(/Работы и акты 0/.test(await tabsText(ok)), '«Очистить» очищает окно');
  await ok.setInputFiles('#ok-open', op); await p.waitForTimeout(1500);
  st = await journal(ok);
  check(/Работы и акты 6/.test(await tabsText(ok)) && st.s4.find(r => r._id === did).weather === 'Ясно, +5', 'файл объекта: сохранить и открыть — всё на месте');

  await p.setInputFiles('#file', path.join(ROOT, 'm29/материалы.xlsx')); await p.waitForTimeout(3000);
  check(/не распознан/.test(await status()), 'чужой файл: понятное сообщение');
  await ctx.close();

  // ---------- 2. в браузере уже есть журнал, который вели вручную
  ({ ctx, p } = await open());
  await p.goto(PULT); await p.waitForTimeout(300);
  const manual = { title: { object: 'Мой объект', num: '7' }, meta: { v: 4 }, s1: [{ fio: 'Петров П.П.', pos: 'Прораб' }], s2: [{ name: 'Журнал бетонных работ' }],
    s3: [{ no: 'АОСР-001/26', name: 'старое' }, { no: 'Р-99', name: 'ручной акт' }],
    s4: [{ _id: 'm1', date: '2026-03-05', aosr: 'АОСР-001/26', work: 'Песчаная подготовка', weather: 'Пасмурно, +2' }, { _id: 'm2', date: '2026-03-06', work: 'Уборка' }],
    s5: [{ date: '2026-03-07', text: 'Замечание ТН' }], bj: {}, uh: {}, sv: {}, ak: {} };
  await (await frames(p))[0].evaluate(m => localStorage.setItem('zpr-journal-v2', JSON.stringify(m)), manual);
  await p.reload(); await p.waitForTimeout(1500);
  [ok] = await frames(p);
  check(await ok.isVisible('#ok-sync'), 'есть журнал, который вели вручную: окно спрашивает, что с ним делать');
  await p.setInputFiles('#file', REG); await p.waitForTimeout(2000);
  st = await journal(ok);
  check(st.title.object === 'Мой объект' && st.s4.length === 2 && !(st.meta && st.meta.source === 'okno'), 'пока не выбрали — журнал не тронут');
  await ok.click('#ok-jimport'); await p.waitForTimeout(1500);
  st = await journal(ok);
  const m1 = st.s4.find(r => r._id === 'm1');
  check(st.meta.source === 'okno' && m1 && m1.weather === 'Пасмурно, +2' && m1.aosr === 'АОСР-001/26' && st.s4.some(r => r.work === 'Уборка'),
    'перенос в окно: записи с погодой сохранены и привязаны к актам');
  check(st.title.num === '7' && st.title.object === 'Мой объект' && st.s1.some(x => x.fio === 'Петров П.П.') && st.s3.some(a => a.no === 'Р-99') && st.s5.length === 1 && st.s2.length === 1,
    'перенос в окно: титул, ИТР, ручной акт, замечания и спецжурналы сохранены');
  check(st.s4.length === 2, 'работы без записей в ручном журнале не получили автозаписей');

  // журнал отвязали в отдельной странице → окно больше не перезаписывает
  await (await frames(p))[1].evaluate(() => { const s = JSON.parse(localStorage.getItem('zpr-journal-v2')); s.meta.source = 'manual'; localStorage.setItem('zpr-journal-v2', JSON.stringify(s)); });
  await ok.click('#ok-tabs [data-tab="s5"]'); await ok.click('#ok-main [data-add="s5"]'); await p.waitForTimeout(1500);
  st = await journal(ok);
  check(st.meta.source === 'manual' && st.s5.length === 1 && await ok.isVisible('#ok-sync'), 'отвязанный журнал окно не перезаписывает и снова спрашивает');
  await ctx.close();

  if (errs.length) { bad++; console.log('Ошибки JS:', errs); }
  await b.close(); process.exit(bad ? 1 : 0);
})();

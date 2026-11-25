// Реестр -> журнал входного контроля и журнал производства работ; дополнение своего журнала.
// Запуск из pto-tools: node tests/test_journal.js
const { chromium } = require('playwright');
const path = require('path'); const fs = require('fs'); const os = require('os');
const ROOT = path.resolve(__dirname, '..');
const JOURNAL = path.resolve(ROOT, '../templates/zhurnal-proizvodstva-rabot/zhurnal-proizvodstva-rabot.html');
(async () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'jr-'));
  const b = await chromium.launch(); const ctx = await b.newContext({ acceptDownloads: true }); const p = await ctx.newPage();
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  let bad = 0; const check = (ok, msg) => { if (!ok) bad++; console.log(`${ok ? 'OK  ' : 'FAIL'} ${msg}`); };
  await p.goto('file://' + path.join(ROOT, 'dist', 'Генератор_АОСР.html'));
  let dl = p.waitForEvent('download'); await p.setInputFiles('#file', path.join(ROOT, 'tests/fixtures/aosr_с_датами.xlsx')); await dl;

  dl = p.waitForEvent('download'); await p.click('#dl-vhk'); await (await dl).saveAs(path.join(tmp, 'vhk.docx'));
  const vrep = await p.evaluate(() => document.getElementById('report').innerText);
  check(/Журнал входного контроля: 20 партий/.test(vrep) && /M19: нет № накладной/.test(vrep), 'журнал входного контроля: 20 партий, предупреждение о накладной');

  dl = p.waitForEvent('download'); await p.click('#dl-jr'); await (await dl).saveAs(path.join(tmp, 'jr.json'));
  const jr = JSON.parse(fs.readFileSync(path.join(tmp, 'jr.json'), 'utf8'));
  check(jr.s1.length === 3 && jr.s3.length === 6 && jr.s4.length === 6, `выгрузка: ИТР ${jr.s1.length}/3, актов ${jr.s3.length}/6, записей ${jr.s4.length}/6`);
  check(jr.s1.some(x => x.fio === 'Ким В.С.' && x.org === 'ТОО «ОтделкаСтрой»' && x.pos === 'Прораб'), 'ИТР разобран на ФИО, должность, организацию');

  if (fs.existsSync(JOURNAL)) {
    const j = await ctx.newPage(); j.on('pageerror', e => errs.push('журнал: ' + e));
    await j.goto('file://' + JOURNAL);
    await j.setInputFiles('#fOpen', path.join(tmp, 'jr.json')); await j.waitForTimeout(400);
    const st = await j.evaluate(() => JSON.parse(localStorage.getItem('zpr-journal-v2')));
    check(st && st.s3.length === 6 && st.s4.length === 6, 'журнал производства работ открыл выгрузку');
  } else console.log('SKIP журнал производства работ не найден рядом (templates/zhurnal-proizvodstva-rabot)');

  const mine = JSON.parse(JSON.stringify(jr));
  mine.s4[0].weather = 'Ясно, -2';
  mine.s4 = mine.s4.filter(r => !/гидроизоляц/i.test(r.work));
  mine.s3.find(a => a.no === 'АОСР-004/26').date = '2026-03-27';
  mine.s4.forEach(r => { delete r.aosr; });   // журнал, заведённый до связи с актами
  mine.s4[0].aosr = 'вручную';                 // номер, вписанный пользователем, не трогаем
  fs.writeFileSync(path.join(tmp, 'mine.json'), JSON.stringify(mine));
  dl = p.waitForEvent('download'); await p.setInputFiles('#jr-upd', path.join(tmp, 'mine.json')); await (await dl).saveAs(path.join(tmp, 'upd.json'));
  const rep = await p.evaluate(() => document.getElementById('report').innerText);
  const upd = JSON.parse(fs.readFileSync(path.join(tmp, 'upd.json'), 'utf8'));
  check(upd.s4.some(r => r.weather === 'Ясно, -2') && upd.s3.find(a => a.no === 'АОСР-004/26').date === '2026-03-27', 'дополнение не перезаписывает ручные данные');
  check(upd.s4.length === 6, 'недостающая запись о работах добавлена');
  check(upd.s4[0].aosr === 'вручную' && upd.s4.slice(1).every(r => r.aosr) && /связано записей с актами \+4/.test(rep),
    'старый журнал: записи связаны с актами, ручной номер не изменён');
  check(/АОСР-003\/26: в журнале нет записей/.test(rep) && /АОСР-004\/26: в журнале дата 27\.03\.2026, в реестре 26\.03\.2026/.test(rep), 'сверка находит расхождения');
  if (errs.length) { bad++; console.log('Ошибки JS:', errs); }
  await b.close(); process.exit(bad ? 1 : 0);
})();

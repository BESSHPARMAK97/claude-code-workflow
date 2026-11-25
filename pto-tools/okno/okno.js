/* Окно объекта: данные объекта вводятся один раз — акты (генератор АОСР) и журнал производства работ берут их отсюда.
   Работает внутри страницы генератора АОСР, когда она открыта во вкладке ПТО-пульта (window.name === "okno"). */
(function () {
  "use strict";
  if (window.name !== "okno") return;
  var C = window.AOSR_CORE;
  var KEY = "pto-object-v1", JKEY = "zpr-journal-v2";

  // ---------- справочники
  var CONCRETE = "Бетонные работы", WELDING = "Сварочные работы", AKZ = "Антикоррозионные работы";
  var JR_OPTS = ["", CONCRETE, WELDING, AKZ];
  var UNITS = ["", "м³", "м²", "м", "т", "шт", "компл"];
  var CARE = ["Укрытие плёнкой", "Полив водой", "Плёнкообразующий состав", "Термос (утеплённая опалубка)", "Электропрогрев", "Греющий провод", "Тепляк", "Обогрев тепловыми пушками"];
  var COMPACT = ["Глубинные вибраторы", "Поверхностные вибраторы", "Виброрейка", "Самоуплотняющаяся смесь"];
  var WELD = ["Ручная дуговая (РД)", "Механизированная в защитном газе (МП)", "Механизированная самозащитной проволокой (МПС)", "Автоматическая под флюсом (АФ)", "Ванная (ВП)", "Контактная точечная (КТ)", "Газовая"];
  var PREP = ["Sa 1", "Sa 2", "Sa 2½", "Sa 3", "St 2", "St 3"];
  var APPLY = ["Кисть", "Валик", "Безвоздушное распыление", "Пневматическое распыление", "Окунание"];
  // поля спецжурналов, которые вводятся вручную (те же ключи, что в журнале производства работ)
  var SPEC = [
    { id: "bj", kind: CONCRETE, title: "Журнал бетонных работ", cols: [
      ["cls", "Класс бетона, марки W и F"], ["slump", "Подвижность (осадка конуса, см / П)"], ["supplier", "Поставщик, № документа о качестве / накладной", "a"],
      ["tmix", "t смеси при укладке, °C"], ["compact", "Способ уплотнения", "s", COMPACT], ["samples", "Образцы: серий, маркировка"],
      ["rproj", "Проектная прочность, МПа"], ["r7", "Прочность 7 сут, МПа"], ["r28", "Прочность 28 сут, МПа"], ["strip", "Дата распалубки", "d"]] },
    { id: "uh", kind: CONCRETE, title: "Журнал ухода за бетоном", cols: [
      ["care", "Способ ухода / прогрева", "s", CARE], ["c1", "Начало ухода", "d"], ["c2", "Окончание ухода", "d"],
      ["tconc", "t бетона при выдерживании: мин / макс, °C"], ["rend", "Прочность к окончанию ухода, МПа (%)"],
      ["remarks", "Замечания (трещины, пересушивание и т. п.)", "a"], ["resp", "Ответственный (Ф.И.О., подпись)"]] },
    { id: "sv", kind: WELDING, title: "Журнал сварочных работ", cols: [
      ["steel", "Марка стали (арматуры) свариваемых элементов"], ["method", "Способ сварки", "s", WELD],
      ["cons", "Марка электродов / проволоки, № сертификата", "a"], ["heat", "Подогрев, защита от ветра и осадков"],
      ["welder", "Ф.И.О. сварщика"], ["cert", "№ удостоверения, клеймо"], ["visual", "Визуальный и измерительный контроль", "s", ["Годно", "Не годно — исправить"]],
      ["ndt", "Неразрушающий контроль: метод, № заключения, результат", "a"]] },
    { id: "ak", kind: AKZ, title: "Журнал антикоррозионной защиты", cols: [
      ["hum", "Отн. влажность воздуха, %"], ["tsurf", "t поверхности, °C"], ["prep", "Подготовка поверхности: способ, степень очистки", "s", PREP],
      ["degr", "Обезжиривание, обеспыливание"], ["mat", "Материал: марка, № партии, сертификат", "a"], ["layers", "Слой № / всего слоёв"],
      ["apply", "Способ нанесения", "s", APPLY], ["thick", "Толщина: слоя / общая, мкм"], ["adh", "Адгезия, сплошность: метод и результат", "a"]] }
  ];
  var TITLE = [["num", "Журнал №"], ["object", "Наименование объекта строительства"], ["address", "Адрес (местоположение) объекта"],
    ["customer", "Заказчик (наименование, БИН, адрес)"], ["contractor", "Генеральный подрядчик (наименование, БИН, № и дата лицензии)"],
    ["subcontractors", "Субподрядные организации и виды работ"], ["designer", "Генеральный проектировщик (наименование, № лицензии)"],
    ["project", "Шифр и наименование проекта (рабочей документации)"], ["expertise", "Положительное заключение экспертизы проекта (№, дата, организация)"],
    ["notice", "Уведомление о начале СМР (№, дата, орган ГАСК)"], ["technadzor", "Технический надзор (организация / эксперт, Ф.И.О., № аттестата)"],
    ["avtornadzor", "Авторский надзор (организация, Ф.И.О., № аттестата)"], ["start", "Начало работ (по договору / фактически)"],
    ["finish", "Окончание работ (по договору / фактически)"], ["responsible", "Ответственный за производство работ и ведение журнала (должность, Ф.И.О., № приказа)"],
    ["sheets", "В журнале пронумеровано, прошнуровано и скреплено печатью листов"]];
  var S1 = [["fio", "Фамилия, имя, отчество"], ["pos", "Должность"], ["org", "Организация"], ["area", "Участок (вид) работ"],
    ["d1", "Дата начала работы", "d"], ["d2", "Дата окончания работы", "d"], ["att", "Приказ о назначении, № аттестата"]];
  var S2 = [["name", "Наименование специального журнала"], ["who", "Организация, ведущая журнал, Ф.И.О. ответственного"],
    ["d1", "Дата начала ведения", "d"], ["d2", "Дата передачи заказчику", "d"], ["sign", "Подпись принявшего"]];
  var S5 = [["date", "Дата", "d"], ["text", "Замечания и указания", "a"], ["who", "Кто выдал (должность, Ф.И.О.)"],
    ["term", "Срок устранения", "d"], ["done", "Отметка об устранении (дата, Ф.И.О.)"], ["ack", "Ознакомлен"]];

  var DATE_H = /^(дата|начало|окончани|действ)/i, PERSON_H = /^представитель|субподряд/i, GEN_H = /^с?формировать/i,
    NUM_H = /^(количеств|цена)/i, MAT_H = /^материал/i, TPL_H = /^типов/i;
  function col(headers, re) { for (var i = 0; i < headers.length; i++) if (re.test(C.clean(headers[i]))) return i; return -1; }

  // ---------- состояние
  var S = null, A = null, saveT = null, syncT = null, lastJournal = "", takeOver = false, tab = "obj", openRows = {};
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function esc(v) { return String(v == null ? "" : v).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function today() { var d = new Date(); return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2); }
  function fmtD(v) { return /^\d{4}-\d{2}-\d{2}$/.test(v || "") ? v.split("-").reverse().join(".") : (v || ""); }
  function blank() {
    return { v: 1, objRows: [], people: [], works: { headers: [], rows: [] }, recv: { headers: [], rows: [] }, extra: {},
      days: [], noAuto: {}, title: {}, s1extra: [], s3extra: [], s2: [], s5: [], spec: { bj: {}, uh: {}, sv: {}, ak: {} }, linked: false };
  }
  function cellStr(v, isDate) {
    if (v === null || v === undefined) return "";
    if (isDate) { var d = C.toDate(v); return d ? C.iso(d) : C.clean(v); }
    if (v instanceof Date) return C.iso(C.toDate(v));
    return String(v).trim();
  }
  function sheetRows(wb, name) {
    var ws = wb.Sheets[name];
    return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false }) : null;
  }
  function table(rows) {
    var headers = (rows[0] || []).map(function (h) { return C.clean(h); });
    while (headers.length && !headers[headers.length - 1]) headers.pop();
    var out = [];
    rows.slice(1).forEach(function (r) {
      var c = headers.map(function (h, i) { return cellStr(r[i], DATE_H.test(h)); });
      if (c.some(function (x) { return x !== ""; })) out.push({ _id: uid(), c: c });
    });
    return { headers: headers, rows: out };
  }
  // реестр Excel → данные окна (ежедневные записи журнала, замечания и т. п. не трогаются)
  function fromWorkbook(wb, keep) {
    var st = keep ? JSON.parse(JSON.stringify(keep)) : blank();
    var o = sheetRows(wb, "Объект") || [[]];
    st.objRows = []; st.people = [];
    o.slice(1).forEach(function (r) {
      if (C.clean(r[0])) st.objRows.push([C.clean(r[0]), cellStr(r[1])]);
      if (C.clean(r[3])) st.people.push(C.clean(r[3]));
    });
    var old = keep ? keep.works : null;
    st.works = table(sheetRows(wb, "Работы") || [[]]);
    if (old) {
      var keyOf = function (t, r) {
        var g = function (re) { var i = col(t.headers, re); return i < 0 ? "" : C.norm(r.c[i]); };
        return g(/^№/) || [g(/^наименование/i), g(/^участок/i), g(/^место/i)].join("|");
      };
      var ids = {}; old.rows.forEach(function (r) { var k = keyOf(old, r); if (k && !ids[k]) ids[k] = r._id; });
      st.works.rows.forEach(function (r) { var k = keyOf(st.works, r); if (k && ids[k]) { r._id = ids[k]; delete ids[k]; } });
    }
    st.recv = table(sheetRows(wb, wb.Sheets["Поступления"] ? "Поступления" : "Материалы") || [[]]);
    st.extra = {};
    wb.SheetNames.forEach(function (n) {
      if (["Объект", "Работы", "Поступления", "Материалы"].indexOf(n) < 0) st.extra[n] = sheetRows(wb, n) || [];
    });
    // записи журнала, привязанные к удалённым работам, остаются со своим текстом
    var ids = {}; st.works.rows.forEach(function (r) { ids[r._id] = 1; });
    st.days.forEach(function (d) { if (d.wid && !ids[d.wid]) d.wid = ""; });
    return st;
  }
  function sampleWb() { return XLSX.read(C.b64ToBytes(window.AOSR_REGISTRY_B64), { type: "array" }); }
  function emptyState() {
    var st = fromWorkbook(sampleWb());
    st.objRows.forEach(function (r) { r[1] = ""; });
    st.people = []; st.works.rows = []; st.recv.rows = [];
    Object.keys(st.extra).forEach(function (n) { st.extra[n] = (st.extra[n] || []).slice(0, 1); });
    return st;
  }

  // данные окна → книга реестра (тот же формат, что читает генератор АОСР)
  function toWorkbook(forExcel) {
    var wb = XLSX.utils.book_new(), rowOf = {};
    var conv = function (h, v) {
      if (!forExcel || v === "") return v;
      if (DATE_H.test(h) && /^\d{4}-\d{2}-\d{2}$/.test(v)) { var p = v.split("-"); return new Date(+p[0], +p[1] - 1, +p[2]); }
      if (NUM_H.test(h) && /^-?\d+([.,]\d+)?$/.test(v)) return +v.replace(",", ".");
      return v;
    };
    var obj = [["Параметр", "Значение", "", "Представители (для выпадающих списков на листе «Работы»)"]];
    for (var i = 0; i < Math.max(S.objRows.length, S.people.length); i++) {
      var r = S.objRows[i] || ["", ""];
      obj.push([r[0], r[1], "", S.people[i] || ""]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(obj), "Объект");
    var sheet = function (t, name, track) {
      var aoa = [t.headers.slice()];
      t.rows.forEach(function (r) {
        if (!r.c.some(function (x) { return C.clean(x) !== ""; })) return;
        aoa.push(t.headers.map(function (h, i) { return conv(h, r.c[i] || ""); }));
        if (track) rowOf[r._id] = aoa.length;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa, { cellDates: true }), name);
    };
    sheet(S.works, "Работы", true);
    sheet(S.recv, "Поступления");
    Object.keys(S.extra).forEach(function (n) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(S.extra[n]), n); });
    return { wb: wb, rowOf: rowOf };
  }
  function wbBytes(wb) { return XLSX.write(wb, { type: "array", bookType: "xlsx" }); }

  // ---------- разбор теми же функциями, что у генератора
  function analyze() {
    var t = toWorkbook(false), res = { rowOf: t.rowOf, idxOfRow: {}, byRow: {}, data: null, check: null, stats: null, jr: null, error: "" };
    try {
      var data = C.load(t.wb), stats = C.prepare(data), check = C.validate(data);
      data.works.forEach(function (w) { (res.byRow[w.row] = res.byRow[w.row] || []).push(w); });
      res.data = data; res.check = check; res.stats = stats;
      S.works.rows.forEach(function (r, i) { if (t.rowOf[r._id]) res.idxOfRow[t.rowOf[r._id]] = i + 1; });
      res.jr = C.journalFromRegistry(data, check).state;
    } catch (e) { res.error = (e && e.message) || String(e); }
    return res;
  }
  function wIdx(re) { return col(S.works.headers, re); }
  function wv(row, re) { var i = wIdx(re); return i < 0 ? "" : C.clean(row.c[i]); }
  function workById(id) { return S.works.rows.filter(function (r) { return r._id === id; })[0]; }
  // что запись журнала берёт из своей работы
  function dayBase(d) {
    var w = d.wid && workById(d.wid);
    if (!w) return {};
    var list = (A && A.byRow[A.rowOf[w._id]]) || [], first = list[0];
    var obj = (A && A.data && A.data.obj) || {};
    var title = first ? C.clean(first.title) : wv(w, /^наименование/i);
    var s = C.toDate(wv(w, /^начало/i)), e = C.toDate(wv(w, /^окончани/i));
    var nos = list.map(function (x) { return C.clean(x.number); }).filter(function (n) { return n && !C.isBN(n); });
    var rep = first ? (A.data.rowReps ? C.clean(first.row_contractor_rep) : obj.contractor_rep) : wv(w, /^представитель смо/i);
    if (/^(нет|-|—|–)$/i.test(rep || "")) rep = "";
    return {
      date: s ? C.iso(s) : "",
      work: title + (s && e && C.iso(e) !== C.iso(s) ? " (работы с " + C.fmtShort(s) + " по " + C.fmtShort(e) + ")" : ""),
      place: [wv(w, /^участок/i), wv(w, /^место/i)].filter(Boolean).join(", "),
      exec: wv(w, /^(исполнитель|организация смо)/i) || C.clean(obj.contractor),
      resp: C.shortName(rep || ""), aosr: nos.join(", "), jr: C.guessJr(title)
    };
  }
  function dayVal(d, k, base) { var v = d[k]; return v !== undefined && v !== "" ? v : (base || dayBase(d))[k] || ""; }
  // у каждой работы с датой начала — своя запись журнала (если её не удаляли)
  function autoDays() {
    var has = {}; S.days.forEach(function (d) { if (d.wid) has[d.wid] = 1; });
    S.works.rows.forEach(function (w) {
      if (!has[w._id] && !S.noAuto[w._id] && C.toDate(wv(w, /^начало/i))) S.days.push({ _id: uid(), wid: w._id });
    });
  }

  // ---------- журнал производства работ из данных окна
  function journalState() {
    var jr = (A && A.jr) || { title: {}, s1: [], s3: [] };
    var st = { title: {}, meta: { v: 4, source: "okno", updated: new Date().toISOString() }, s1: [], s2: [], s3: [], s4: [], s5: [],
      bj: {}, uh: {}, sv: {}, ak: {} };
    C.TITLE_KEYS.forEach(function (k) { st.title[k] = C.clean(S.title[k]) || (jr.title && jr.title[k]) || ""; });
    var fios = {};
    (jr.s1 || []).concat(S.s1extra).forEach(function (p) {
      var k = C.norm(p.fio); if (!k || fios[k]) return; fios[k] = 1; st.s1.push(strip(p));
    });
    st.s2 = S.s2.map(strip);
    st.s3 = (jr.s3 || []).map(strip);
    var nos = {}; st.s3.forEach(function (a) { nos[C.norm(a.no)] = 1; });
    S.s3extra.forEach(function (a) { if (!nos[C.norm(a.no)]) st.s3.push(strip(a)); });
    S.days.forEach(function (d) {
      var b = dayBase(d), r = { _id: d._id };
      ["date", "shift", "jr", "aosr", "work", "place", "unit", "vol", "temp", "weather", "exec", "resp"].forEach(function (k) { r[k] = dayVal(d, k, b); });
      st.s4.push(r);
    });
    st.s4.sort(function (a, b) { return (a.date || "9").localeCompare(b.date || "9") || (a.shift || "").localeCompare(b.shift || ""); });
    st.s5 = S.s5.map(strip);
    SPEC.forEach(function (j) { st[j.id] = JSON.parse(JSON.stringify(S.spec[j.id] || {})); });
    return st;
  }
  function strip(o) { return Object.assign({}, o); }
  function tellPult(msg) { if (window.parent !== window) window.parent.postMessage(Object.assign({ tool: "aosr" }, msg), "*"); }
  function writeJournal() {
    if (!S.linked) return;
    // журнал отвязали («Вести журнал здесь вручную») или открыли чужой — не перезаписываем, спрашиваем
    var cur = null; try { cur = JSON.parse(localStorage.getItem(JKEY)); } catch (e) {}
    if (!takeOver && cur && cur.title && !(cur.meta && cur.meta.source === "okno")) { S.linked = false; lastJournal = ""; save(); renderSync(); return; }
    var st = journalState(), cmp = JSON.stringify(Object.assign({}, st, { meta: null }));
    if (cmp === lastJournal) return;
    lastJournal = cmp;
    try { localStorage.setItem(JKEY, JSON.stringify(st)); takeOver = false; } catch (e) { note("Не удалось записать журнал в браузер.", "bad"); return; }
    tellPult({ type: "pto-journal-updated", quiet: true });
  }
  // журнал, который вели вручную до окна, — в окно (ничего не теряется)
  function importJournal(j) {
    A = analyze();
    var num = {}; S.works.rows.forEach(function (w) {
      (A.byRow[A.rowOf[w._id]] || []).forEach(function (x) { if (C.clean(x.number)) num[C.norm(x.number)] = w._id; });
    });
    var used = {};
    S.days = (j.s4 || []).map(function (r) {
      var wid = "";
      String(r.aosr || "").split(/[,;]/).some(function (n) { var id = num[C.norm(n)]; if (id) { wid = id; return true; } return false; });
      var d = { _id: r._id || uid(), wid: wid }, b = wid ? dayBase(d) : {};
      ["date", "shift", "jr", "work", "place", "unit", "vol", "temp", "weather", "exec", "resp"].forEach(function (k) {
        var v = C.clean(r[k]); if (v && v !== C.clean(b[k])) d[k] = v;
      });
      if (wid) used[wid] = 1;
      return d;
    });
    // работы, у которых не было записи, — без автозаписи: журнал вели вручную
    S.works.rows.forEach(function (w) { if (!used[w._id]) S.noAuto[w._id] = 1; });
    var tj = (A.jr && A.jr.title) || {};
    Object.keys(j.title || {}).forEach(function (k) { var v = C.clean(j.title[k]); if (v && v !== C.clean(tj[k])) S.title[k] = v; });
    var derived = {}; ((A.jr && A.jr.s1) || []).forEach(function (p) { derived[C.norm(p.fio)] = 1; });
    S.s1extra = (j.s1 || []).filter(function (p) { return p.fio && !derived[C.norm(p.fio)]; }).map(function (p) { return Object.assign({ _id: uid() }, p); });
    var acts = {}; ((A.jr && A.jr.s3) || []).forEach(function (a) { acts[C.norm(a.no)] = 1; });
    S.s3extra = (j.s3 || []).filter(function (a) { return a.no && !acts[C.norm(a.no)]; });
    S.s2 = (j.s2 || []).map(function (p) { return Object.assign({ _id: uid() }, p); });
    S.s5 = (j.s5 || []).map(function (p) { return Object.assign({ _id: uid() }, p); });
    SPEC.forEach(function (s) { S.spec[s.id] = JSON.parse(JSON.stringify(j[s.id] || {})); });
  }

  // ---------- сохранение и пересчёт
  function save() {
    clearTimeout(saveT);
    saveT = setTimeout(function () { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch (e) { note("Не удалось сохранить данные в браузере — сохраните файл объекта.", "bad"); } }, 250);
  }
  function changed(now) {
    save(); clearTimeout(syncT);
    var run = function () { A = analyze(); autoDays(); renderCheck(); renderTabs(); writeJournal(); };
    if (now) run(); else syncT = setTimeout(run, 500);
  }

  // ---------- отрисовка
  var root;
  function note(text, kind) { var el = document.getElementById("ok-note"); if (!el) return; el.textContent = text || ""; el.className = "ok-note " + (kind || ""); el.hidden = !text; }
  function field(type, val, attrs, opts, ph) {
    attrs = attrs + (ph ? ' placeholder="' + esc(ph) + '"' : "");
    if (type === "d") return '<input type="date" ' + attrs + ' value="' + esc(val) + '">';
    if (type === "s") {
      var o = opts.slice(); if (o.indexOf("") < 0) o.unshift(""); if (val && o.indexOf(val) < 0) o.push(val);
      return "<select " + attrs + ">" + o.map(function (x) { return "<option" + (x === val ? " selected" : "") + ' value="' + esc(x) + '">' + esc(x || (ph ? "авто: " + ph : "")) + "</option>"; }).join("") + "</select>";
    }
    if (type === "list") return '<input type="text" ' + attrs + ' value="' + esc(val) + '">';
    return "<textarea rows=\"1\" " + attrs + ">" + esc(val) + "</textarea>";
  }
  function headerType(h) { return DATE_H.test(h) ? "d" : GEN_H.test(h) ? "gen" : PERSON_H.test(h) ? "person" : MAT_H.test(h) ? "mat" : TPL_H.test(h) ? "tpl" : "t"; }
  function headerField(h, val, attrs) {
    var t = headerType(h);
    if (t === "gen") return field("s", val, attrs, ["да", "нет"], "да");
    if (t === "person") return field("s", val, attrs + ' class="wide"', S.people.filter(function (x) { return C.clean(x); }).concat(["нет"]));
    if (t === "mat") return field("list", val, attrs + ' list="ok-codes"');
    if (t === "tpl") return field("list", val, attrs + ' list="ok-tpls"');
    return field(t, val, attrs);
  }
  var TABS = [["obj", "Объект и люди"], ["works", "Работы и акты"], ["recv", "Поступления"], ["days", "Записи журнала"], ["s5", "Замечания"], ["s2", "Спецжурналы"]];
  function count(id) {
    return id === "works" ? S.works.rows.length : id === "recv" ? S.recv.rows.length : id === "days" ? S.days.length :
      id === "s5" ? S.s5.length : id === "s2" ? S.s2.length : S.people.length;
  }
  function renderTabs() {
    document.getElementById("ok-tabs").innerHTML = TABS.map(function (t) {
      return '<button type="button" data-tab="' + t[0] + '" class="' + (tab === t[0] ? "on" : "") + '">' + esc(t[1]) + ' <span class="cnt">' + count(t[0]) + "</span></button>";
    }).join("");
  }
  function lists() {
    var codes = S.recv.rows.map(function (r) { return C.clean(r.c[0]); }).filter(Boolean);
    var tpl = (S.extra["Типовые работы"] || []).slice(1).map(function (r) { return C.clean(r[0]); }).filter(Boolean);
    return '<datalist id="ok-codes">' + codes.map(function (c) { return '<option value="' + esc(c) + '">'; }).join("") + "</datalist>" +
      '<datalist id="ok-tpls">' + tpl.map(function (c) { return '<option value="' + esc(c) + '">'; }).join("") + "</datalist>";
  }
  function grid(kind, cols, rows, opt) {
    // cols: [ключ|индекс, подпись, тип, варианты]; opt.cell(row, col) → html поля; opt.details(row) → html или ""
    var W = opt.widths || {};
    var head = "<tr><th>№</th>" + cols.map(function (c) { var w = W[c[0]]; return "<th" + (w ? ' style="width:' + w + 'px;min-width:' + w + 'px"' : "") + ">" + esc(c[1]) + "</th>"; }).join("") + "<th></th></tr>";
    var body = rows.map(function (r, i) {
      var det = opt.details ? opt.details(r) : "";
      return '<tr data-row="' + r._id + '"><td class="n">' + (i + 1) + "</td>" + cols.map(function (c) { return "<td>" + opt.cell(r, c) + "</td>"; }).join("") +
        '<td class="act"><button type="button" class="del" data-del="' + kind + "|" + r._id + '" title="Удалить строку">×</button></td></tr>' +
        (det ? '<tr class="sub"><td></td><td colspan="' + (cols.length + 1) + '"><details data-open="' + r._id + '"' + (openRows[r._id] ? " open" : "") + ">" + det + "</details></td></tr>" : "");
    }).join("");
    return '<div class="ok-wrap"><table class="ok-t"' + (opt.minWidth ? ' style="min-width:' + opt.minWidth + 'px"' : "") + "><thead>" + head + "</thead><tbody>" +
      (body || '<tr><td colspan="' + (cols.length + 2) + '" class="empty">' + esc(opt.empty || "Строк пока нет.") + "</td></tr>") + "</tbody></table></div>" +
      '<div class="row ok-bar2"><button type="button" class="primary" data-add="' + kind + '">' + esc(opt.add || "Добавить строку") + "</button>" + (opt.more || "") + "</div>";
  }
  function simpleGrid(kind, defs, rows, opt) {
    return grid(kind, defs, rows, Object.assign({ cell: function (r, c) { return field(c[2] || "t", r[c[0]] || "", 'data-p="' + kind + "|" + r._id + "|" + c[0] + '"', c[3] || []); } }, opt));
  }
  function regGrid(kind, t, summary, opt) {
    var idx = summary.map(function (re) { return col(t.headers, re); }).filter(function (i) { return i >= 0; });
    var cols = idx.map(function (i) { return [i, t.headers[i]]; }), widths = {};
    idx.forEach(function (i) { var h = t.headers[i]; widths[i] = /^наименование/i.test(h) ? 320 : DATE_H.test(h) ? 150 : /^№|^код/i.test(h) ? 140 : GEN_H.test(h) ? 90 : 130; });
    var rest = t.headers.map(function (h, i) { return i; }).filter(function (i) { return idx.indexOf(i) < 0 && t.headers[i]; });
    return grid(kind, cols, t.rows, Object.assign({
      widths: widths,
      cell: function (r, c) { return headerField(t.headers[c[0]], r.c[c[0]] || "", 'data-p="' + kind + "|" + r._id + "|" + c[0] + '"'); },
      details: function (r) {
        return "<summary>Остальные поля (" + rest.length + ")</summary><div class=\"ok-form\">" + rest.map(function (i) {
          return '<label class="f"><span>' + esc(t.headers[i]) + "</span>" + headerField(t.headers[i], r.c[i] || "", 'data-p="' + kind + "|" + r._id + "|" + i + '"') + "</label>";
        }).join("") + "</div>";
      }
    }, opt));
  }
  function renderMain() {
    var m = document.getElementById("ok-main"), h = "";
    if (tab === "obj") {
      h = '<div class="ok-sec"><h3>Объект</h3><div class="ok-form">' + S.objRows.map(function (r, i) {
        return '<label class="f"><span>' + esc(r[0]) + "</span>" + field("t", r[1], 'data-p="obj|' + i + '"') + "</label>";
      }).join("") + "</div></div>" +
      '<div class="ok-sec"><h3>Люди — представители для актов</h3><p class="hint">Формат: «Должность Организация: Фамилия И.О.», для надзора — с аттестатом. Из этого списка выбираются представители в работах; ИТР подрядчика попадают в раздел 1 журнала.</p>' +
        S.people.map(function (p, i) {
          return '<div class="row pl">' + field("t", p, 'data-p="ppl|' + i + '"') + '<button type="button" class="del" data-del="ppl|' + i + '" title="Удалить">×</button></div>';
        }).join("") + '<div class="row ok-bar2"><button type="button" data-add="ppl">Добавить человека</button></div></div>' +
      '<div class="ok-sec"><h3>Титул журнала производства работ</h3><p class="hint">Серым — что подставится из данных объекта. Впишите, чтобы заменить.</p><div class="ok-form">' +
        TITLE.map(function (t) {
          var auto = A && A.jr ? C.clean(A.jr.title[t[0]]) : "";
          return '<label class="f"><span>' + esc(t[1]) + "</span>" + field("t", S.title[t[0]] || "", 'data-p="ttl|' + t[0] + '"', null, auto) + "</label>";
        }).join("") + "</div></div>" +
      '<div class="ok-sec"><h3>Дополнительные ИТР для раздела 1 журнала</h3><p class="hint">ИТР подрядчика и субподрядчиков из работ попадают в раздел 1 сами. Здесь — остальные.</p>' +
        simpleGrid("s1", S1, S.s1extra, { add: "Добавить ИТР", empty: "Нет." }) + "</div>";
    } else if (tab === "works") {
      h = '<p class="hint">Одна строка — один акт скрытых работ. По датам начала работ сами создаются записи журнала; акты попадают в раздел 3 журнала.</p>' +
        regGrid("w", S.works, [/^№/, /^наименование/i, /^участок/i, /^начало/i, /^окончани/i, /^дата акта/i, GEN_H], { add: "Добавить работу" });
    } else if (tab === "recv") {
      h = '<p class="hint">Одна строка — партия материала по накладной с документом о качестве. Код партии указывается в работах (п. 3 акта); отсюда же — журнал входного контроля.</p>' +
        regGrid("r", S.recv, [/^код/i, /^наименование/i, /^ед/i, /^количеств/i, /^дата поступл/i, /^входн/i], { add: "Добавить партию" });
    } else if (tab === "days") {
      var wopts = S.works.rows.map(function (w) { return [w._id, (wv(w, /^№/) ? wv(w, /^№/) + " — " : "") + wv(w, /^наименование/i)]; });
      var cols = [["date", "Дата", "d"], ["shift", "Смена", "s", ["1", "2", "3"]], ["wid", "Работа (акт)"], ["work", "Описание работ"], ["place", "Место"],
        ["unit", "Ед.", "s", UNITS], ["vol", "Объём"], ["temp", "t воздуха, °C"], ["weather", "Погода"], ["exec", "Исполнитель"], ["resp", "Ответственный"]];
      var rows = S.days.slice().sort(function (a, b) { return (dayVal(a, "date") || "9").localeCompare(dayVal(b, "date") || "9"); });
      h = '<p class="hint">Записи раздела 4 журнала. Выберите работу — описание, место, исполнитель и № акта подставятся из неё (серым). Погоду, объём и данные спецжурналов вносите здесь.</p>' +
        grid("d", cols, rows, {
          cell: function (r, c) {
            var a = 'data-p="d|' + r._id + "|" + c[0] + '"', b = dayBase(r);
            if (c[0] === "wid") return '<select class="wide" ' + a + '><option value="">— без акта —</option>' + wopts.map(function (o) {
              return '<option value="' + esc(o[0]) + '"' + (o[0] === r.wid ? " selected" : "") + ">" + esc(o[1]) + "</option>"; }).join("") + "</select>" +
              (b.aosr ? '<div class="tag">АОСР № ' + esc(b.aosr) + "</div>" : "");
            return field(c[2] || "t", r[c[0]] || "", a, c[3] || [], b[c[0]]) +
              (c[0] === "date" && !r.date && b.date ? '<div class="tag">из работы: ' + esc(fmtD(b.date)) + "</div>" : "");
          },
          details: function (r) {
            var kind = dayVal(r, "jr"), js = SPEC.filter(function (j) { return j.kind === kind; });
            return "<summary>Спецжурнал: " + esc(kind || "нет") + (js.length ? " · данные для " + js.map(function (j) { return j.title.replace(/^Журнал /, ""); }).join(", ") : "") + "</summary>" +
              '<div class="ok-form"><label class="f"><span>В спецжурнал</span>' + field("s", r.jr || "", 'data-p="d|' + r._id + '|jr"', JR_OPTS, dayBase(r).jr || "нет") + "</label></div>" +
              js.map(function (j) {
                var own = (S.spec[j.id] || {})[r._id] || {};
                return "<h4>" + esc(j.title) + '</h4><div class="ok-form">' + j.cols.map(function (c) {
                  return '<label class="f"><span>' + esc(c[1]) + "</span>" + field(c[2] || "t", own[c[0]] || "", 'data-p="sp|' + j.id + "|" + r._id + "|" + c[0] + '"', c[3] || []) + "</label>";
                }).join("") + "</div>";
              }).join("");
          },
          widths: { date: 140, shift: 70, wid: 220, work: 280, place: 170, unit: 80, vol: 80, temp: 80, weather: 130, exec: 150, resp: 130 }, minWidth: 1700,
          add: "Добавить запись", empty: "Записей нет. Они появятся сами, когда у работ будут даты начала.",
          more: '<button type="button" data-add="dcopy">Повторить последнюю на сегодня</button>'
        });
    } else if (tab === "s5") {
      h = '<p class="hint">Раздел 5 журнала: замечания технического и авторского надзора.</p>' + simpleGrid("s5", S5, S.s5, { add: "Добавить замечание" });
    } else if (tab === "s2") {
      h = '<p class="hint">Раздел 2 журнала: перечень специальных журналов.</p>' + simpleGrid("s2", S2, S.s2, { add: "Добавить журнал" });
    }
    m.innerHTML = h + lists();
    m.querySelectorAll("textarea").forEach(grow);
  }
  function grow(t) { t.style.height = "auto"; t.style.height = Math.max(t.scrollHeight + 2, 34) + "px"; }
  function renderCheck() {
    var el = document.getElementById("ok-check");
    if (!S.works.rows.length) {
      el.className = "ok-check"; el.innerHTML = "<strong>С чего начать.</strong> Загрузите реестр Excel, нажмите «Заполнить примером» или заполните «Объект и люди» и добавьте работы.";
      return;
    }
    if (!A || A.error) { el.className = "ok-check bad"; el.textContent = "Проверка не выполнена: " + (A ? A.error : ""); return; }
    var lines = C.reportLines(A.data, A.check, [], "окна объекта", A.stats).filter(function (l) { return /✗|!|^Лист|^Строка/.test(l); });
    var errs = lines.filter(function (l) { return /✗/.test(l); }).length;
    // «Строка N» — строка книги реестра; в окне — номер строки на вкладке «Работы и акты»
    lines = lines.map(function (l) { return l.replace(/^Строка (\d+)/, function (m, n) { return "Работа " + (A.idxOfRow[n] || n); }); });
    el.className = "ok-check " + (errs ? "bad" : lines.length ? "mixed" : "ok");
    el.innerHTML = "<strong>" + (errs ? "Ошибок: " + errs + " — такие акты не сформируются" : lines.length ? "Ошибок нет, есть замечания" : "Проверка: ошибок нет, можно формировать акты") +
      "</strong>" + (lines.length ? "<div>" + lines.slice(0, 14).map(function (l) { return "<div class=\"" + (/✗/.test(l) ? "e" : /!/.test(l) ? "w" : "h") + "\">" + esc(l) + "</div>"; }).join("") +
      (lines.length > 14 ? "<div>…и ещё " + (lines.length - 14) + "</div>" : "") + "</div>" : "");
  }
  function renderSync() {
    var el = document.getElementById("ok-sync");
    if (S.linked) { el.hidden = true; return; }
    el.hidden = false;
    el.innerHTML = "<strong>В этом браузере уже есть журнал производства работ, который вели вручную.</strong> " +
      "Окно пока его не трогает. Выберите, что сделать:" +
      '<div class="row" style="margin-top:8px"><button type="button" class="primary" id="ok-jimport">Перенести записи журнала в окно</button>' +
      '<button type="button" id="ok-jreplace">Заменить журнал данными окна</button></div>';
  }
  function renderAll() { renderTabs(); renderMain(); renderCheck(); renderSync(); }

  // ---------- правка
  function setPath(p, v) {
    var a = p.split("|"), k = a[0];
    var find = function (list, id) { return list.filter(function (r) { return r._id === id; })[0]; };
    if (k === "obj") S.objRows[+a[1]][1] = v;
    else if (k === "ppl") S.people[+a[1]] = v;
    else if (k === "ttl") S.title[a[1]] = v;
    else if (k === "w") find(S.works.rows, a[1]).c[+a[2]] = v;
    else if (k === "r") find(S.recv.rows, a[1]).c[+a[2]] = v;
    else if (k === "d") { var d = find(S.days, a[1]); d[a[2]] = v; }
    else if (k === "s1") find(S.s1extra, a[1])[a[2]] = v;
    else if (k === "s2") find(S.s2, a[1])[a[2]] = v;
    else if (k === "s5") find(S.s5, a[1])[a[2]] = v;
    else if (k === "sp") { var o = (S.spec[a[1]] = S.spec[a[1]] || {}); (o[a[2]] = o[a[2]] || {})[a[3]] = v; }
  }
  function onEdit(e) {
    var t = e.target, p = t.dataset && t.dataset.p; if (!p) return;
    if (t.tagName === "TEXTAREA") grow(t);
    setPath(p, t.value);
    if (e.type === "change" && t.tagName === "SELECT" && /^d\|/.test(p)) { changed(true); renderMain(); }
    else changed(false);
  }
  function onClick(e) {
    var b = e.target.closest("button"); if (!b) return;
    if (b.dataset.tab) { tab = b.dataset.tab; try { localStorage.setItem(KEY + "-tab", tab); } catch (x) {} renderTabs(); renderMain(); return; }
    if (b.dataset.add) {
      var k = b.dataset.add, r = { _id: uid() };
      if (k === "ppl") S.people.push("");
      else if (k === "w") { r.c = S.works.headers.map(function () { return ""; }); var g = col(S.works.headers, GEN_H); if (g >= 0) r.c[g] = "да"; S.works.rows.push(r); }
      else if (k === "r") { r.c = S.recv.headers.map(function () { return ""; }); S.recv.rows.push(r); }
      else if (k === "d") { r.date = today(); S.days.push(r); }
      else if (k === "dcopy") {
        var last = S.days.slice().sort(function (a, b2) { return (dayVal(a, "date") || "").localeCompare(dayVal(b2, "date") || ""); }).pop();
        if (last) { r = Object.assign({}, last, { _id: uid(), date: today(), vol: "" }); S.days.push(r); }
      }
      else if (k === "s1") S.s1extra.push(r); else if (k === "s2") S.s2.push(r); else if (k === "s5") { r.date = today(); S.s5.push(r); }
      changed(true); renderMain();
      var f = document.querySelector('#ok-main [data-p^="' + (k === "dcopy" ? "d" : k) + "|" + (k === "ppl" ? S.people.length - 1 : r._id) + '"]');
      if (f) f.focus();
      return;
    }
    if (b.dataset.del) {
      if (!b.dataset.armed) {
        b.dataset.armed = "1"; b.textContent = "Удалить?"; b.classList.add("armed");
        setTimeout(function () { if (b.isConnected) { delete b.dataset.armed; b.textContent = "×"; b.classList.remove("armed"); } }, 3000);
        return;
      }
      var a = b.dataset.del.split("|"), drop = function (list) { return list.filter(function (x) { return x._id !== a[1]; }); };
      if (a[0] === "ppl") S.people.splice(+a[1], 1);
      else if (a[0] === "w") { S.works.rows = drop(S.works.rows); S.days.forEach(function (d) { if (d.wid === a[1]) d.wid = ""; }); }
      else if (a[0] === "r") S.recv.rows = drop(S.recv.rows);
      else if (a[0] === "d") {
        var gone = S.days.filter(function (x) { return x._id === a[1]; })[0];
        if (gone && gone.wid) S.noAuto[gone.wid] = 1;   // удалённая автозапись не появится снова
        S.days = drop(S.days); SPEC.forEach(function (j) { if (S.spec[j.id]) delete S.spec[j.id][a[1]]; });
      }
      else if (a[0] === "s1") S.s1extra = drop(S.s1extra); else if (a[0] === "s2") S.s2 = drop(S.s2); else if (a[0] === "s5") S.s5 = drop(S.s5);
      changed(true); renderMain();
    }
  }

  // ---------- файлы
  function readFile(f, asText, cb) {
    var r = new FileReader(); r.onload = function () { cb(r.result); };
    if (asText) r.readAsText(f); else r.readAsArrayBuffer(f);
  }
  function hasData() { return S.works.rows.length || S.recv.rows.length || S.days.length; }
  function loadWorkbook(wb, name) {
    if (!wb.Sheets["Работы"]) { note("«" + name + "» — не реестр работ: нет листа «Работы».", "bad"); return; }
    if (hasData() && !confirm("Заменить работы, поступления и данные объекта в окне данными из «" + name + "»?\nЗаписи журнала (погода, объёмы, спецжурналы) и замечания сохранятся.")) return;
    S = fromWorkbook(wb, S);
    tab = "works"; changed(true); renderAll();
    note("Загружен реестр «" + name + "»: работ — " + S.works.rows.length + ", партий — " + S.recv.rows.length + ", записей журнала — " + S.days.length + ".", "ok");
  }
  window.OKNO_IMPORT = loadWorkbook;
  function stamp() { var d = new Date(); return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2); }
  function objName() { var o = S.objRows[0] && C.clean(S.objRows[0][1]); return o ? "_" + o.replace(/[\\/:*?"<>|«»]+/g, "").slice(0, 40).trim().replace(/\s+/g, "_") : ""; }

  // ---------- запуск
  var CSS = "body.okno main{max-width:1280px;padding-top:20px}body.okno main>h1,body.okno main>.sub,body.okno #drop{display:none}" +
    "#okno h2{font-size:22px;margin:0 0 2px}#okno .sub2{color:var(--muted);margin:0 0 12px;font-size:14px}" +
    ".ok-bar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}.ok-bar .primary,.ok-bar2 .primary{background:var(--accent);border-color:var(--accent);color:var(--card)}" +
    ".ok-bar label{font:inherit;border:1px solid var(--line);background:var(--card);border-radius:8px;padding:8px 14px;cursor:pointer}.ok-bar label:hover{border-color:var(--accent)}.ok-bar input{display:none}" +
    ".ok-note{border-radius:10px;padding:10px 14px;margin-bottom:12px;font-size:14px;background:var(--accent-soft)}.ok-note.ok{background:var(--ok-bg)}.ok-note.bad{background:var(--bad-bg)}" +
    "#ok-sync{border:1px solid var(--mix-line);background:var(--mix-bg);border-radius:10px;padding:12px 14px;margin-bottom:12px;font-size:14px}" +
    "#ok-tabs{display:flex;flex-wrap:wrap;gap:4px;border-bottom:1px solid var(--line);margin-bottom:12px}" +
    "#ok-tabs button{border:1px solid transparent;border-bottom:0;border-radius:8px 8px 0 0;background:none;color:var(--muted);padding:8px 14px}" +
    "#ok-tabs button.on{background:var(--card);border-color:var(--line);color:var(--text);box-shadow:inset 0 2px 0 var(--accent)}#ok-tabs .cnt{font-size:12px;color:var(--muted)}" +
    "#ok-main{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:12px}" +
    ".ok-sec{margin-bottom:18px}.ok-sec h3{font-size:16px;margin:4px 0 6px}#ok-main h4{font-size:14px;margin:10px 0 4px}" +
    ".hint{color:var(--muted);font-size:13px;margin:0 0 10px}" +
    ".ok-form{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:8px 14px}" +
    ".f{display:flex;flex-direction:column;gap:3px;font-size:12.5px;color:var(--muted)}" +
    "#okno textarea,#okno input[type=text],#okno input[type=date],#okno select{font:inherit;font-size:14px;color:var(--text);background:var(--bg);border:1px solid var(--line);border-radius:6px;padding:6px 8px;width:100%;min-width:0;resize:vertical}" +
    "#okno textarea{overflow:hidden;min-height:34px}#okno select{min-width:70px}#okno select.wide{min-width:180px;max-width:100%}" +
    "#okno ::placeholder{color:var(--muted);opacity:.75}" +
    ".pl{margin-bottom:6px;flex-wrap:nowrap}.pl textarea{flex:1}" +
    ".ok-wrap{overflow-x:auto}.ok-t{border-collapse:collapse;width:100%;min-width:980px}" +
    ".ok-t th{font-size:12px;text-align:left;color:var(--muted);font-weight:600;padding:6px;border-bottom:1px solid var(--line);vertical-align:bottom}" +
    ".ok-t td{padding:4px;vertical-align:top;border-bottom:1px solid var(--line)}.ok-t td.n{color:var(--muted);font-size:12px;padding-top:12px;width:28px}" +
    ".ok-t tr.sub td{border-bottom:1px solid var(--line);padding-top:0}.ok-t tr[data-row] td{border-bottom:0}" +
    ".ok-t tr.sub details{position:sticky;left:4px;max-width:min(1180px,calc(100vw - 90px))}" +
    ".ok-t summary{cursor:pointer;color:var(--accent);font-size:13px;padding:2px 0 6px}.ok-t td.empty{color:var(--muted);text-align:center;padding:18px}" +
    ".del{border:0;background:none;color:var(--muted);font-size:18px;padding:4px 8px}.del.armed{color:var(--err);font-size:12px}" +
    ".ok-bar2{margin-top:10px;gap:8px}.tag{font-size:12px;color:var(--accent);margin-top:3px}" +
    ".ok-check{border-radius:12px;padding:12px 16px;margin-bottom:16px;font-size:13.5px;border:1px solid var(--line)}" +
    ".ok-check.ok{background:var(--ok-bg);border-color:var(--ok-line)}.ok-check.mixed{background:var(--mix-bg);border-color:var(--mix-line)}.ok-check.bad{background:var(--bad-bg);border-color:var(--bad-line)}" +
    ".ok-check .e{color:var(--err)}.ok-check .w{color:var(--warn)}.ok-check .h{font-weight:600;margin-top:4px}";

  document.addEventListener("DOMContentLoaded", function () {
    try { S = JSON.parse(localStorage.getItem(KEY)); } catch (e) { S = null; }
    if (!S || !S.works) S = emptyState();
    S = Object.assign(blank(), S);
    try { tab = localStorage.getItem(KEY + "-tab") || tab; } catch (e) {}
    if (!S.linked) {   // журнал, который вели вручную, не перезаписываем без спроса
      var j = null; try { j = JSON.parse(localStorage.getItem(JKEY)); } catch (e) {}
      S.linked = !(j && j.title && !(j.meta && j.meta.source === "okno"));
    }
    var st = document.createElement("style"); st.textContent = CSS; document.head.appendChild(st);
    document.body.classList.add("okno");
    ["jr-push", "dl-jr"].forEach(function (id) { var x = document.getElementById(id); if (x) x.hidden = true; });
    var upd = document.getElementById("jr-upd"); if (upd && upd.parentElement) upd.parentElement.hidden = true;
    var sampleLink = document.getElementById("sample"); if (sampleLink) sampleLink.textContent = "Скачать пустой реестр Excel с примером";
    root = document.createElement("section"); root.id = "okno";
    root.innerHTML = "<h2>Окно объекта</h2>" +
      '<p class="sub2">Данные вводятся один раз — акты скрытых работ и журнал производства работ берут их отсюда. Всё сохраняется в этом браузере; резервная копия — «Сохранить файл объекта».</p>' +
      '<div class="ok-bar"><button type="button" class="primary" id="ok-acts">Сформировать акты (Word)</button>' +
      '<button type="button" id="ok-save">Сохранить файл объекта</button><label>Открыть файл объекта<input type="file" id="ok-open" accept=".json"></label>' +
      '<label>Загрузить реестр Excel<input type="file" id="ok-xls" accept=".xlsx"></label><button type="button" id="ok-xlsout">Выгрузить в Excel</button>' +
      '<button type="button" id="ok-sample">Заполнить примером</button><button type="button" id="ok-clear">Очистить</button></div>' +
      '<div id="ok-note" class="ok-note" hidden></div><div id="ok-sync" hidden></div><nav id="ok-tabs"></nav><div id="ok-main"></div><div id="ok-check" class="ok-check"></div>';
    var main = document.querySelector("main"); main.insertBefore(root, document.getElementById("report"));
    var om = document.getElementById("ok-main");
    om.addEventListener("input", onEdit); om.addEventListener("change", onEdit);
    root.addEventListener("click", onClick);
    om.addEventListener("toggle", function (e) { var d = e.target; if (d.dataset && d.dataset.open) { if (d.open) { openRows[d.dataset.open] = 1; d.querySelectorAll("textarea").forEach(grow); } else delete openRows[d.dataset.open]; } }, true);

    document.getElementById("ok-acts").addEventListener("click", function () {
      if (!S.works.rows.length) { note("Добавьте работы на вкладке «Работы и акты» или загрузите реестр Excel.", "bad"); return; }
      var t = toWorkbook(false), f = new File([wbBytes(t.wb)], "данные окна объекта", { type: C.XLSXT });
      C.process(f, { okno: true });
      setTimeout(function () { var r = document.getElementById("report"); if (r && !r.hidden) r.scrollIntoView({ behavior: "smooth", block: "start" }); }, 300);
    });
    document.getElementById("ok-save").addEventListener("click", function () {
      C.download(new TextEncoder().encode(JSON.stringify(S, null, 1)), "Объект" + objName() + "_" + stamp() + ".json", "application/json");
    });
    document.getElementById("ok-open").addEventListener("change", function () {
      var f = this.files[0]; this.value = ""; if (!f) return;
      readFile(f, true, function (txt) {
        var x = null; try { x = JSON.parse(txt); } catch (e) {}
        if (!x || !x.works || !x.works.headers) { note("«" + f.name + "» — не файл окна объекта.", "bad"); return; }
        if (hasData() && !confirm("Заменить все данные окна данными из «" + f.name + "»?")) return;
        var linked = S.linked; S = Object.assign(blank(), x); S.linked = linked;
        changed(true); renderAll(); note("Открыт файл объекта «" + f.name + "».", "ok");
      });
    });
    document.getElementById("ok-xls").addEventListener("change", function () {
      var f = this.files[0]; this.value = ""; if (!f) return;
      readFile(f, false, function (buf) {
        try { loadWorkbook(XLSX.read(new Uint8Array(buf), { type: "array" }), f.name); }
        catch (e) { note("Не удалось прочитать «" + f.name + "»: " + ((e && e.message) || e), "bad"); }
      });
    });
    document.getElementById("ok-xlsout").addEventListener("click", function () {
      C.download(wbBytes(toWorkbook(true).wb), "реестр_работ" + objName() + "_" + stamp() + ".xlsx", C.XLSXT);
    });
    document.getElementById("ok-sample").addEventListener("click", function () {
      if (hasData() && !confirm("Заменить данные окна примером?")) return;
      var linked = S.linked; S = fromWorkbook(sampleWb()); S.linked = linked;
      tab = "works"; changed(true); renderAll(); note("Загружен пример объекта.", "ok");
    });
    document.getElementById("ok-clear").addEventListener("click", function () {
      if (!confirm("Очистить все данные окна? Сначала можно сохранить файл объекта.")) return;
      var linked = S.linked; S = emptyState(); S.linked = linked; tab = "obj"; changed(true); renderAll(); note("Окно очищено.", "ok");
    });
    root.addEventListener("click", function (e) {
      if (e.target.id === "ok-jimport") {
        var j = null; try { j = JSON.parse(localStorage.getItem(JKEY)); } catch (x) {}
        if (j) importJournal(j);
        S.linked = true; takeOver = true; changed(true); renderAll(); note("Записи журнала перенесены в окно: " + S.days.length + ". Дальше журнал собирается из окна.", "ok");
      } else if (e.target.id === "ok-jreplace") {
        if (!confirm("Журнал в этом браузере будет заменён данными окна. Сначала можно сохранить его файл кнопкой «Сохранить файл» в журнале.")) return;
        S.linked = true; takeOver = true; changed(true); renderAll();
      }
    });
    A = analyze(); autoDays(); save(); renderAll(); writeJournal();
  });
})();

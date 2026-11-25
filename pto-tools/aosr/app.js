/* Генератор АОСР (Приложение Д к СН РК 1.03-00-2022) — работает полностью в браузере, офлайн. */
(function () {
  "use strict";
  var MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
  var OBJ_KEYS = [
    [/^объект/i, "object"], [/^город/i, "city"], [/^строительно-монтажная/i, "contractor"],
    [/^проектная организация/i, "designer"], [/^представитель смо/i, "contractor_rep"],
    [/^представитель технадзора/i, "tech_rep"], [/^представитель авторского/i, "author_rep"],
    [/^первый номер/i, "first_number"], [/^материалы в приложение/i, "appendix_over"]
  ];
  var OBJ_NAMES = { object: "Объект", city: "Город", contractor: "Строительно-монтажная организация",
    designer: "Проектная организация", contractor_rep: "Представитель СМО", tech_rep: "Представитель технадзора" };

  // Колонки ищутся по названию — порядок в Excel может быть любым, старые реестры тоже читаются
  var WORK_SPEC = [
    ["generate", /^с?формировать/i], ["tpl", /^типов/i], ["number", /^№/], ["date", /^дата акта/i],
    ["place", /^участок/i], ["start", /^начало/i], ["end", /^окончани/i], ["location", /^место/i],
    ["title", /^наименование/i], ["drawings", /^черт/i], ["materials", /^материал/i], ["evidence", /^доп/i],
    ["deviations", /^отклонен/i], ["next_works", /^последующ/i], ["performer", /^(исполнитель|организация смо)/i],
    ["row_contractor_rep", /^представитель смо/i], ["row_tech_rep", /^представитель технадзора/i],
    ["row_author_rep", /^представитель авторского/i], ["sub_rep", /субподряд/i]
  ];
  var TPL_SPEC = [
    ["name", /^назван/i], ["title", /^наименование/i], ["location", /^место/i], ["drawings", /^черт/i],
    ["materials", /^материал/i], ["evidence", /^доп/i], ["deviations", /^отклонен/i], ["next_works", /^последующ/i],
    ["performer", /^исполнитель/i], ["sub_rep", /субподряд/i]
  ];
  var TPL_FIELDS = ["title", "location", "drawings", "materials", "evidence", "deviations", "next_works", "performer", "sub_rep"];
  // лист «Поступления» (партия + документ о качестве) или старый «Материалы»
  var MAT_SPEC = [
    ["code", /^код/i], ["name", /^наименование/i], ["unit", /^ед/i], ["qty", /^количеств/i],
    ["recv", /^дата поступл/i], ["invoice", /^№ наклад/i], ["doc_type", /^вид/i], ["doc_no", /^№/],
    ["date", /^дата/i], ["until", /^действ/i], ["supplier", /^поставщик/i], ["batch", /^парти/i], ["control", /^входн/i]
  ];
  var REQUIRED = { number: "№ акта", title: "Наименование работ", drawings: "Чертежи", next_works: "Последующие работы" };
  var BLANK_DATE = "«___» ___________ 20___ г.";   // пустая дата — заполняется от руки

  function $(id) { return document.getElementById(id); }
  function clean(v) { return v === null || v === undefined ? "" : String(v).trim(); }
  function norm(v) { return clean(v).toLowerCase().replace(/\s+/g, " "); }
  function lowerFirst(t) { return t ? t.charAt(0).toLowerCase() + t.slice(1) : t; }
  function b64ToBytes(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function pad(n) { return (n < 10 ? "0" : "") + n; }

  // Дата: {y,m,d} из числа Excel, строки «дд.мм.гггг» или Date
  function toDate(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") {
      var p = XLSX.SSF.parse_date_code(v);
      return p ? { y: p.y, m: p.m, d: p.d } : null;
    }
    if (v instanceof Date && !isNaN(v)) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    var s = String(v).trim(), m;
    if ((m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/))) {
      var y = +m[3]; if (y < 100) y += 2000;
      return { y: y, m: +m[2], d: +m[1] };
    }
    if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return { y: +m[1], m: +m[2], d: +m[3] };
    return null;
  }
  function key(d) { return d.y * 10000 + d.m * 100 + d.d; }
  function fmtShort(d) { return pad(d.d) + "." + pad(d.m) + "." + d.y; }
  function fmtLong(d, blank) { return d ? "«" + pad(d.d) + "» " + MONTHS[d.m - 1] + " " + d.y + " г." : (blank || BLANK_DATE); }
  var BLANK_SHORT = "«___» ________ 20___ г.";   // для п. 6, чтобы обе даты поместились в строку
  function shortName(rep) { if (!rep) return ""; var i = rep.lastIndexOf(":"); return (i >= 0 ? rep.slice(i + 1) : rep).trim(); }
  function safeName(s) { return s.replace(/[\\/:*?"<>|]+/g, "_").slice(0, 80).trim(); }

  function rows(wb, name, optional) {
    var ws = wb.Sheets[name];
    if (!ws) { if (optional) return null; throw new Error("в файле нет листа «" + name + "»"); }
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  }
  function mapHeaders(header, spec) {
    var map = {}, used = {};
    spec.forEach(function (sp) {
      for (var i = 0; i < header.length; i++) {
        if (!used[i] && sp[1].test(clean(header[i]))) { map[sp[0]] = i; used[i] = true; return; }
      }
    });
    return map;
  }
  function pick(r, map) {
    var o = {};
    Object.keys(map).forEach(function (f) { var v = r[map[f]]; o[f] = v === undefined ? null : v; });
    return o;
  }

  function load(wb) {
    var obj = {}, loadErrors = [];
    rows(wb, "Объект").slice(1).forEach(function (r) {
      var label = clean(r[0]);
      for (var i = 0; i < OBJ_KEYS.length; i++) if (OBJ_KEYS[i][0].test(label)) { obj[OBJ_KEYS[i][1]] = clean(r[1]); break; }
    });

    var materials = {};
    var matSheet = wb.Sheets["Поступления"] ? "Поступления" : "Материалы";
    var mr = rows(wb, matSheet), mmap = mapHeaders(mr[0] || [], MAT_SPEC);
    if (mmap.code === undefined) { mmap = { code: 0, name: 1, doc_type: 2, doc_no: 3, date: 4, supplier: 5, batch: 6, until: 7 }; }
    mr.slice(1).forEach(function (r) {
      var m = pick(r, mmap);
      if (!clean(m.code)) return;
      materials[clean(m.code)] = { name: clean(m.name), doc_type: clean(m.doc_type), doc_no: clean(m.doc_no), date: toDate(m.date),
        supplier: clean(m.supplier), until: toDate(m.until), untilRaw: m.until, recv: toDate(m.recv), invoice: clean(m.invoice),
        qty: m.qty, unit: clean(m.unit), control: clean(m.control),
        batch: clean(m.batch) || [clean(m.qty).replace(".", ","), clean(m.unit)].filter(Boolean).join(" ") };
    });

    var templates = {}, tr = rows(wb, "Типовые работы", true);
    if (tr) {
      var tmap = mapHeaders(tr[0] || [], TPL_SPEC);
      if (tmap.name === undefined) tmap.name = 0;
      tr.slice(1).forEach(function (r) {
        var t = pick(r, tmap);
        if (!clean(t.name)) return;
        TPL_FIELDS.forEach(function (f) { t[f] = clean(t[f]); });
        templates[norm(t.name)] = t;
      });
    }

    var floors = {}, fr = rows(wb, "Этажи", true);
    if (fr) fr.slice(1).forEach(function (r) {
      if (clean(r[1]) === "" || clean(r[2]) === "") return;
      floors[norm(r[0]) + "|" + clean(r[1])] = clean(r[2]);
    });

    var wr = rows(wb, "Работы"), header = wr[0] || [], wmap = mapHeaders(header, WORK_SPEC);
    if (wmap.title === undefined && wmap.tpl === undefined)
      loadErrors.push("на листе «Работы» нет колонки «Наименование работ»");
    var works = [];
    wr.forEach(function (r, i) {
      if (i === 0 || !r.some(function (v) { return clean(v) !== ""; })) return;
      var w = pick(r, wmap);
      // строка, где заполнено только «Формировать», — пустая
      if (!Object.keys(w).some(function (f) { return f !== "generate" && clean(w[f]) !== ""; })) return;
      w.row = i + 1;
      works.push(w);
    });
    return { obj: obj, materials: materials, matSheet: matSheet, works: works, floors: floors, templates: templates, hasFloors: !!fr,
      header: header, wmap: wmap, loadErrors: loadErrors,
      rowReps: wmap.row_contractor_rep !== undefined || wmap.row_tech_rep !== undefined };
  }

  function core(x) { return clean(x).replace(/^(ТОО|ИП|АО|ООО)\s*/i, "").replace(/[«»"']/g, "").toLowerCase(); }
  function isNone(v) { return /^(нет|-|—|–)$/i.test(clean(v)); }
  // значение для акта: из строки (если заполнено), иначе с листа «Объект»; «нет» — убрать
  function pickRep(rowVal, objVal) { var v = clean(rowVal); if (!v) return objVal || ""; return isNone(v) ? "" : v; }
  function codesOf(w) { return clean(w.materials).split(";").map(function (s) { return s.trim(); }).filter(Boolean); }
  function batchesOf(data, c) {
    if (data.materials[c]) return [c];
    return Object.keys(data.materials).filter(function (k) { return k.indexOf(c + "-") === 0; }).sort();
  }
  function matsOf(data, w) {
    var out = [];
    codesOf(w).forEach(function (c) { batchesOf(data, c).forEach(function (k) { if (out.indexOf(k) < 0) out.push(k); }); });
    return out.map(function (k) { return data.materials[k]; });
  }
  function isBN(num) { return /^б\/н$/i.test(clean(num)); }   // «б/н» — заготовка без номера

  // «Блок 1, этажи 1–3, 5» -> {prefix: "Блок 1", floors: ["1","2","3","5"]}
  function parseFloors(place) {
    var m = clean(place).match(/^(.*?)[,;\s]*этаж(?:и|ах)\s+([\d\s,;–—-]+)$/i);
    if (!m) return null;
    var list = [];
    m[2].split(/[,;]/).forEach(function (part) {
      part = part.trim(); if (!part) return;
      var r = part.match(/^(\d+)\s*[–—-]\s*(\d+)$/);
      if (r) {
        var a = +r[1], b = +r[2], step = a <= b ? 1 : -1;
        for (var f = a; step > 0 ? f <= b : f >= b; f += step) list.push(String(f));
      } else if (/^\d+$/.test(part)) list.push(String(+part));
    });
    return list.length ? { prefix: m[1].trim(), floors: list.slice(0, 300) } : null;
  }

  // АОСР-Блок S1-001/26 + 2 -> АОСР-Блок S1-003/26 (увеличивается число перед «/», иначе последнее число)
  function incNumber(num, k) {
    if (!k || isBN(num)) return num;
    var slash = num.indexOf("/"), head = slash >= 0 ? num.slice(0, slash) : num, tail = slash >= 0 ? num.slice(slash) : "";
    var m = head.match(/(\d+)(?!.*\d)/);
    if (!m) return num + "-" + (k + 1);
    var n = String(+m[1] + k); while (n.length < m[1].length) n = "0" + n;
    return head.slice(0, m.index) + n + head.slice(m.index + m[1].length) + tail;
  }

  // справочник -> нумерация -> разворот по этажам -> дата акта -> последующие работы
  function prepare(data) {
    var out = [], stats = { srcRows: 0, expanded: 0, blankDate: 0, autoNext: 0, fromTpl: 0, autoNum: 0 };
    var lastNum = "";
    data.works.forEach(function (w) {
      w.pre0 = [];
      if (clean(w.tpl)) {
        var t = data.templates[norm(w.tpl)];
        if (!t) w.pre0.push(["e", "типовая работа «" + clean(w.tpl) + "» не найдена на листе «Типовые работы»"]);
        else { TPL_FIELDS.forEach(function (f) { if (clean(w[f]) === "" && t[f]) w[f] = t[f]; }); stats.fromTpl++; }
      }
      var base = clean(w.number);
      if (!base) {
        base = lastNum ? incNumber(lastNum, 1) : clean(data.obj.first_number);
        if (base) stats.autoNum++;
      }
      var pf = parseFloors(w.place);
      var list = pf ? pf.floors : [null];
      if (pf) stats.srcRows++;
      list.forEach(function (fl, k) {
        var nw = Object.assign({}, w, { pre: w.pre0.slice(), number: base ? incNumber(base, k) : "" });
        if (fl === null) { nw.label = "Строка " + w.row; out.push(nw); }
        else {
          var mark = data.floors[norm(pf.prefix) + "|" + fl] || data.floors["|" + fl] || "";
          nw.label = "Строка " + w.row + ", " + fl + " этаж";
          nw.place = (pf.prefix ? pf.prefix + ", " : "") + fl + " этаж";
          var t = clean(w.title), loc = clean(w.location);
          if (/\{(этаж|отметка)\}/.test(t + loc)) {
            if (/\{отметка\}/.test(t + loc) && !mark) nw.pre.push(["e", "нет отметки для " + fl + " этажа на листе «Этажи»"]);
            var sub = function (x) { return x.replace(/\{этаж\}/g, fl).replace(/\{отметка\}/g, mark); };
            nw.title = sub(t); nw.location = sub(loc);
          } else {
            if (!mark && data.hasFloors) nw.pre.push(["w", "нет отметки для " + fl + " этажа на листе «Этажи» — указан только этаж"]);
            nw.location = loc + (loc ? ", " : "") + (mark ? "с отм. " + mark + " " : "") + "(" + fl + " этаж)";
          }
          stats.expanded++;
          out.push(nw);
        }
        if (nw.number && !isBN(nw.number)) lastNum = nw.number;
      });
    });
    out.forEach(function (w, i) {
      w.id = i + 1;
      w.sel = clean(w.generate).toLowerCase() !== "нет";
      if (w.sel && clean(w.date) === "") stats.blankDate++;
    });
    out.forEach(function (w, i) {
      if (clean(w.next_works) !== "" || norm(w.place) === "") return;
      for (var j = i + 1; j < out.length; j++) {
        if (norm(out[j].place) === norm(w.place) && clean(out[j].title)) {
          w.next_works = lowerFirst(clean(out[j].title)); stats.autoNext++; return;
        }
      }
    });
    data.works = out;
    return stats;
  }

  function validate(data) {
    var errors = {}, warns = {};
    function err(r, m) { (errors[r] = errors[r] || []).push(m); }
    function warn(r, m) { (warns[r] = warns[r] || []).push(m); }
    data.loadErrors.forEach(function (m) { err("obj", m); });
    Object.keys(OBJ_NAMES).forEach(function (k) {
      if (data.rowReps && (k === "contractor_rep" || k === "tech_rep")) return;   // представители — в строках «Работ»
      if (!data.obj[k]) err("obj", "не заполнено «" + OBJ_NAMES[k] + "»");
    });
    (data.rowReps ? [] : ["contractor_rep", "tech_rep", "author_rep"]).forEach(function (k) {
      var v = data.obj[k];
      if (v && v.indexOf(":") < 0) warn("obj", "«" + (OBJ_NAMES[k] || "Представитель авторского надзора") + "»: нет двоеточия перед ФИО — расшифровка подписи будет целой строкой");
    });
    var seen = {}, prev = null, prevByPlace = {};
    data.works.forEach(function (w) {
      var r = w.id;
      (w.pre || []).forEach(function (x) { (x[0] === "e" ? err : warn)(r, x[1]); });
      Object.keys(REQUIRED).forEach(function (k) {
        if (clean(w[k]) !== "") return;
        if (k === "next_works") err(r, "не заполнено «Последующие работы», и на этом участке нет следующего акта, откуда их взять");
        else if (k === "number") err(r, "нет № акта: укажите его или «Первый номер акта» на листе «Объект»");
        else err(r, "не заполнено «" + REQUIRED[k] + "»");
      });
      ["date", "start", "end"].forEach(function (k) {
        if (clean(w[k]) !== "" && !toDate(w[k])) err(r, "не распознана дата: «" + w[k] + "»");
      });
      var d = toDate(w.date), s = toDate(w.start), e = toDate(w.end);
      if (s && e && key(s) > key(e)) err(r, "начало работ " + fmtShort(s) + " позже окончания " + fmtShort(e));
      if (e && d && key(e) > key(d)) err(r, "акт от " + fmtShort(d) + " составлен раньше окончания работ " + fmtShort(e));
      var num = clean(w.number);
      if (num && !isBN(num)) { if (seen[num]) err(r, "№ акта " + num + " повторяется (" + seen[num].toLowerCase() + ")"); else seen[num] = w.label; }
      if (prev && d && prev.d && key(d) < key(prev.d)) warn(r, "дата акта " + fmtShort(d) + " раньше, чем у предыдущего акта № " + prev.n + " (" + fmtShort(prev.d) + ")");
      prev = { n: num, d: d };
      var pk = norm(w.place), pp = pk && prevByPlace[pk];
      if (pp && s && pp.e && key(s) < key(pp.e)) warn(r, "работы начаты " + fmtShort(s) + " до окончания предыдущих на этом участке (акт № " + pp.n + ", " + fmtShort(pp.e) + ")");
      if (pk) prevByPlace[pk] = { n: num, e: e };
      if (data.rowReps) {
        if (!clean(w.row_contractor_rep) || isNone(w.row_contractor_rep)) err(r, "не указан «Представитель СМО» — без него акт недействителен");
        if (!clean(w.row_tech_rep) || isNone(w.row_tech_rep)) err(r, "не указан «Представитель технадзора» — без него акт недействителен");
        // «Организация СМО» не указана, а представитель — из другой организации
        var orgCore = core(data.obj.contractor), crep = clean(w.row_contractor_rep);
        if (!clean(w.performer) && orgCore && crep && !isNone(crep) && core(crep).indexOf(orgCore) < 0)
          warn(r, "представитель СМО не из «" + clean(data.obj.contractor) + "», а «Организация СМО» в строке пустая — в акте будет «" +
            clean(data.obj.contractor) + "». Проверьте");
      }
      [["row_contractor_rep", "Представитель СМО"], ["row_tech_rep", "Представитель технадзора"],
       ["row_author_rep", "Представитель авторского надзора"], ["sub_rep", "Представитель субподрядчика"]].forEach(function (x) {
        var v = clean(w[x[0]]);
        if (v && !isNone(v) && v.indexOf(":") < 0) warn(r, "«" + x[1] + "»: нет двоеточия перед ФИО — расшифровка подписи будет целой строкой");
      });
      var codes = codesOf(w);
      if (!codes.length) warn(r, "не указаны материалы");
      var all = [];
      codes.forEach(function (c0) {
        var list = batchesOf(data, c0);
        if (!list.length) err(r, "материал «" + c0 + "» не найден на листе «" + data.matSheet + "»");
        list.forEach(function (k) { if (all.indexOf(k) < 0) all.push(k); });
      });
      all.forEach(function (c) {
        var m = data.materials[c];
        if (m.recv && e && key(m.recv) > key(e)) err(r, "партия " + c + " поступила " + fmtShort(m.recv) + " — после окончания работ " + fmtShort(e));
        else if (m.recv && s && key(m.recv) > key(s)) warn(r, "партия " + c + " поступила " + fmtShort(m.recv) + " — после начала работ " + fmtShort(s));
        if (!m.doc_no || !m.date) err(r, "у материала " + c + " нет номера или даты документа о качестве");
        else if (d && key(m.date) > key(d)) err(r, "документ на " + c + " (" + m.doc_no + ") датирован " + fmtShort(m.date) + " — позже акта");
        else if (!d && e && key(m.date) > key(e)) err(r, "документ на " + c + " (" + m.doc_no + ") датирован " + fmtShort(m.date) + " — позже окончания работ");
        else if (d && e && key(m.date) > key(e)) warn(r, "документ на " + c + " датирован " + fmtShort(m.date) + " — после окончания работ");
        if (clean(m.untilRaw) !== "" && !m.until) err(r, "у материала " + c + " не распознана дата «Действует до»: «" + m.untilRaw + "»");
        if (m.until) {
          if (s && key(m.until) < key(s)) err(r, "документ на " + c + " (" + m.doc_no + ") действовал до " + fmtShort(m.until) + " — истёк до начала работ " + fmtShort(s));
          else if (e && key(m.until) < key(e)) warn(r, "документ на " + c + " (" + m.doc_no + ") истёк " + fmtShort(m.until) + ", в период работ (окончание " + fmtShort(e) + ")");
        } else if (m.date && d && (Date.UTC(d.y, d.m - 1, d.d) - Date.UTC(m.date.y, m.date.m - 1, m.date.d)) / 864e5 > 365) {
          warn(r, "документу на " + c + " больше года, а срок действия («Действует до») не указан — проверьте");
        }
      });
    });
    return { errors: errors, warns: warns };
  }

  // ---------- Word
  function esc(t) { return String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  var RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';
  function xp(text, opt) {
    opt = opt || {};
    return '<w:p><w:pPr><w:spacing w:after="' + (opt.after || 0) + '"/><w:jc w:val="' + (opt.jc || "center") + '"/></w:pPr>' +
      '<w:r>' + RPR + (opt.bold ? "<w:b/>" : "") + '<w:sz w:val="' + (opt.sz || 20) + '"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
  }
  function xcell(text, w, bold, jc) {
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="' + jc + '"/></w:pPr><w:r>' + RPR + (bold ? "<w:b/>" : "") +
      '<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p></w:tc>';
  }
  function xtable(widths, headRow, bodyRows, centerCols) {
    var b = '<w:tblBorders>' + ["top", "left", "bottom", "right", "insideH", "insideV"].map(function (x) {
      return '<w:' + x + ' w:val="single" w:sz="4" w:space="0" w:color="000000"/>'; }).join("") + '</w:tblBorders>';
    var row = function (cells, bold) {
      return '<w:tr><w:trPr><w:cantSplit/>' + (bold ? '<w:tblHeader/>' : '') + '</w:trPr>' + cells.map(function (c, i) {
        return xcell(c, widths[i], bold, bold || centerCols.indexOf(i) >= 0 ? "center" : "left"); }).join("") + '</w:tr>';
    };
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/>' + b + '</w:tblPr><w:tblGrid>' +
      widths.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join("") + '</w:tblGrid>' +
      row(headRow, true) + bodyRows.map(function (r) { return row(r, false); }).join("") + '</w:tbl>';
  }

  function appendixXml(ctx, matList) {
    var head = '<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="right"/></w:pPr><w:r><w:br w:type="page"/></w:r><w:r>' + RPR +
      '<w:i/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">' + esc("Приложение к акту № " + ctx.number) + '</w:t></w:r></w:p>' +
      xp("освидетельствования скрытых работ от " + ctx.date_text, { jc: "right", after: 160 }) +
      xp("Перечень материалов, конструкций и изделий, применённых при выполнении работ", { bold: true, sz: 24, after: 60 }) +
      xp(ctx.works, { after: 160 });
    var tbl = xtable([500, 3300, 2900, 2200, 1100],
      ["№ п/п", "Наименование материала, конструкции, изделия", "Документ, подтверждающий качество", "Поставщик", "Партия / количество"],
      matList.map(function (m, i) {
        return [String(i + 1), m.name, m.doc_type + " № " + m.doc_no + " от " + fmtShort(m.date) + (m.until ? ", действует до " + fmtShort(m.until) : ""), m.supplier, m.batch];
      }), [0]);
    var sign = '<w:p><w:pPr><w:spacing w:before="360" w:after="0"/></w:pPr><w:r>' + RPR + '<w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">' +
      esc("Представитель строительно-монтажной организации   ____________   " + ctx.contractor_sign) + '</w:t></w:r></w:p>';
    return head + tbl + sign;
  }

  function render(data, check, templateBytes) {
    var obj = data.obj, over = parseInt(obj.appendix_over, 10);
    if (isNaN(over)) over = 4;
    var base = {
      object: obj.object, city: obj.city, contractor_rep: obj.contractor_rep, tech_rep: obj.tech_rep,
      author_rep: obj.author_rep || "", contractor_sign: shortName(obj.contractor_rep),
      tech_sign: shortName(obj.tech_rep), author_sign: shortName(obj.author_rep)
    };
    var out = [];
    data.works.forEach(function (w) {
      if (!w.sel || check.errors[w.id]) return;
      var d = toDate(w.date), s = toDate(w.start), e = toDate(w.end);
      var matList = matsOf(data, w);
      var mats = matList.map(function (m) { return m.name + ", " + m.doc_type + " № " + m.doc_no + " от " + fmtShort(m.date); });
      var toAppendix = over > 0 && matList.length > over;
      var title = clean(w.title), loc = clean(w.location), sub = clean(w.sub_rep);
      // есть колонки представителей — в акте ровно то, что в строке; нет колонок (старый реестр) — с листа «Объект»
      var own = function (v) { return isNone(v) ? "" : clean(v); };
      var cRep = data.rowReps ? own(w.row_contractor_rep) : obj.contractor_rep,
        tRep = data.rowReps ? own(w.row_tech_rep) : obj.tech_rep,
        aRep = data.rowReps ? own(w.row_author_rep) : pickRep(w.row_author_rep, obj.author_rep);
      var ctx = Object.assign({}, base, {
        contractor_rep: cRep, contractor_sign: shortName(cRep), tech_rep: tRep, tech_sign: shortName(tRep),
        author_rep: aRep, author_sign: shortName(aRep),
        number: clean(w.number), place: clean(w.place), works: (title + " " + loc).trim(),
        design_docs: [obj.designer, clean(w.drawings)].filter(Boolean).join(", "),
        materials_text: toAppendix ? "согласно приложению к настоящему акту (" + matList.length + " поз.)" : (mats.join("; ") || "не применялись"),
        evidence: clean(w.evidence) || "не приложены", deviations: clean(w.deviations) || "отсутствуют",
        date_text: fmtLong(d), start_text: fmtLong(s, BLANK_SHORT), end_text: fmtLong(e, BLANK_SHORT),
        next_works: clean(w.next_works), performer: clean(w.performer) || obj.contractor,
        sub_rep: isNone(sub) ? "" : sub, sub_sign: isNone(sub) ? "" : shortName(sub)
      });
      var doc = new Docxtemplater(new PizZip(templateBytes), { paragraphLoop: true, linebreaks: true, nullGetter: function () { return ""; } });
      doc.render(ctx);
      var z = doc.getZip();
      if (toAppendix) {
        var x = z.file("word/document.xml").asText(), cut = x.lastIndexOf("<w:sectPr");
        z.file("word/document.xml", x.slice(0, cut) + appendixXml(ctx, matList) + x.slice(cut));
      }
      out.push({
        name: "АОСР_" + safeName(ctx.number) + "_" + safeName(title) + ".docx",
        bytes: z.generate({ type: "uint8array", compression: "DEFLATE" }),
        zip: z, number: ctx.number, title: title, works: ctx.works, place: ctx.place, date: d ? fmtShort(d) : "", appendix: toAppendix
      });
    });
    return out;
  }

  function bodyOf(xmlText) {
    var start = xmlText.indexOf(">", xmlText.indexOf("<w:body")) + 1;
    return xmlText.slice(start, xmlText.lastIndexOf("<w:sectPr"));
  }
  function withPageBreak(bodyXml) {
    // разрыв страницы в начале первого абзаца акта — каждый акт с новой страницы, без пустых листов
    return bodyXml.replace(/(<w:p(?:\s[^>]*)?>)(<w:pPr>[\s\S]*?<\/w:pPr>)?/, function (m, p, ppr) {
      return p + (ppr || "") + '<w:r><w:br w:type="page"/></w:r>';
    });
  }
  function buildKit(obj, acts) {
    var base = acts[0].zip;
    var docXml = base.file("word/document.xml").asText();
    var start = docXml.indexOf(">", docXml.indexOf("<w:body")) + 1;
    var sect = docXml.lastIndexOf("<w:sectPr");
    var front = xp("Комплект исполнительной документации", { bold: true, sz: 26, after: 80 }) +
      xp(obj.object || "", { after: 40 }) +
      xp("Актов освидетельствования скрытых работ в комплекте: " + acts.length, { after: 160 }) +
      xtable([680, 1900, 1250, 4400, 1900], ["№ п/п", "№ акта", "Дата", "Наименование работ", "Участок"],
        acts.map(function (a, i) { return [String(i + 1), a.number, a.date, a.works, a.place || ""]; }), [0, 2]);
    var bodies = acts.map(function (a) { return withPageBreak(bodyOf(a.zip.file("word/document.xml").asText())); }).join("");
    var kit = new PizZip(base.generate({ type: "uint8array" }));
    kit.file("word/document.xml", docXml.slice(0, start) + front + bodies + docXml.slice(sect));
    return kit.generate({ type: "uint8array", compression: "DEFLATE" });
  }

  // ---------- Excel
  function filledRegistry(wb, data) {
    var header = data.header.slice(), map = Object.assign({}, data.wmap);
    if (map.number === undefined) { map.number = header.length; header.push("№ акта"); }
    var aoa = [header];
    data.works.forEach(function (w) {
      var row = header.map(function () { return ""; });
      Object.keys(map).forEach(function (f) {
        var v = w[f];
        if (f === "date" || f === "start" || f === "end") { var d = toDate(v); v = d ? new Date(d.y, d.m - 1, d.d) : v; }
        row[map[f]] = v === null || v === undefined ? "" : v;
      });
      aoa.push(row);
    });
    // встроенный формат 14 — Excel показывает его как короткую дату системы (в RU/KZ: ДД.ММ.ГГГГ)
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: XLSX.SSF.get_table()[14] });
    ws["!cols"] = header.map(function () { return { wch: 18 }; });
    var out = XLSX.utils.book_new();
    wb.SheetNames.forEach(function (n) { XLSX.utils.book_append_sheet(out, n === "Работы" ? ws : wb.Sheets[n], n); });
    return new Uint8Array(XLSX.write(out, { type: "array", bookType: "xlsx" }));
  }

  function registerXlsx(obj, acts) {
    var aoa = [["Реестр исполнительной документации — " + (obj.object || "")], [],
      ["№ п/п", "Наименование документа", "№ документа", "Дата", "Кол-во листов", "Примечание"]];
    acts.forEach(function (a, i) {
      aoa.push([i + 1, "Акт освидетельствования скрытых работ: " + a.works, a.number, a.date, a.appendix ? 2 : 1, a.appendix ? "с приложением" : ""]);
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 7 }, { wch: 70 }, { wch: 16 }, { wch: 12 }, { wch: 10 }, { wch: 20 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Реестр ИД");
    return new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
  }

  // ---------- отчёт
  function reportLines(data, check, acts, fileName, stats) {
    var now = new Date();
    var lines = ["Проверка реестра " + fileName + " — " + pad(now.getDate()) + "." + pad(now.getMonth() + 1) + "." + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes()), ""];
    var byId = {}; data.works.forEach(function (w) { byId[w.id] = w; });
    var ids = Object.keys(Object.assign({}, check.errors, check.warns)).filter(function (k) { return k === "obj" || byId[k].sel; });
    ids.sort(function (a, b) { return (a === "obj" ? -1 : +a) - (b === "obj" ? -1 : +b); });
    ids.forEach(function (r) {
      lines.push(r === "obj" ? "Лист «Объект»:" : byId[r].label + " (акт № " + (clean(byId[r].number) || "?") + "):");
      (check.errors[r] || []).forEach(function (m) { lines.push("  ✗ ОШИБКА: " + m); });
      (check.warns[r] || []).forEach(function (m) { lines.push("  ! внимание: " + m); });
    });
    var skipped = data.works.filter(function (w) { return !w.sel; }).length;
    var failed = Object.keys(check.errors).filter(function (k) { return k !== "obj" && byId[k].sel; }).length;
    var withApp = acts.filter(function (a) { return a.appendix; }).length;
    if (ids.length) lines.push("");
    if (stats.fromTpl) lines.push("Заполнено из «Типовых работ»: " + stats.fromTpl + " строк(и)");
    if (stats.srcRows) lines.push("Развёрнуто по этажам: " + stats.srcRows + " строк(и) → " + stats.expanded + " актов");
    if (stats.autoNum) lines.push("№ актов проставлены автоматически: " + stats.autoNum + " строк(и)");
    if (stats.blankDate) lines.push("Без даты акта (впишете от руки): " + stats.blankDate);
    if (stats.autoNext) lines.push("«Последующие работы» взяты из следующего акта на участке: " + stats.autoNext);
    if (withApp) lines.push("Материалы вынесены в приложение: " + withApp + " акт(ов)");
    if (skipped) lines.push("Не формировались («Формировать» = нет): " + skipped);
    lines.push("Сформировано актов: " + acts.length, "Пропущено из-за ошибок: " + failed);
    return lines;
  }

  function download(bytes, name, type) {
    var blob = new Blob([bytes], { type: type });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ---------- журналы: входной контроль (Word) и выгрузка в журнал производства работ (.json)
  function iso(d) { return d ? d.y + "-" + pad(d.m) + "-" + pad(d.d) : ""; }
  function isoKey(s) { var d = toDate(s); return d ? key(d) : 0; }

  // «Прораб ТОО «ОтделкаСтрой», аттестат № 1: Ким В.С.» -> {fio, pos, org, att}
  function parseRep(rep) {
    var v = clean(rep), i = v.lastIndexOf(":"), head = i >= 0 ? v.slice(0, i) : "", fio = (i >= 0 ? v.slice(i + 1) : v).trim();
    var att = (head.match(/,?\s*(аттестат[^,]*)/i) || [])[1] || "";
    head = head.replace(/,?\s*аттестат[^,]*/i, "").trim();
    var m = head.match(/(ТОО|ИП|АО|ООО|ГКП|РГП|КГП|ПК)\s*[«"]/);
    var pos = m ? head.slice(0, m.index).trim() : head, org = m ? head.slice(m.index).trim() : "";
    return { fio: fio, pos: pos, org: org, att: att };
  }

  // спецжурнал по описанию работ — те же правила, что в самом журнале
  function guessJr(text) {
    var t = String(text || "").toLowerCase();
    if (/антикорроз|огрунтов|грунтовк|грунтован|защитн\S* покрыт|окраск\S* (металл|конструкц)|лакокрас/.test(t)) return "Антикоррозионные работы";
    if (/свар|прихват|наплавк/.test(t)) return "Сварочные работы";
    if (/бетонир|укладк\S* бетон|бетонн\S* смес|подлив\S* бетон/.test(t)) return "Бетонные работы";
    return "";
  }
  var TITLE_KEYS = ["num", "object", "address", "customer", "contractor", "subcontractors", "designer", "project", "expertise",
    "notice", "technadzor", "avtornadzor", "start", "finish", "responsible", "sheets"];

  function repsOf(data, w) {
    var own = function (v) { return isNone(v) ? "" : clean(v); };
    return {
      c: data.rowReps ? own(w.row_contractor_rep) : data.obj.contractor_rep || "",
      t: data.rowReps ? own(w.row_tech_rep) : data.obj.tech_rep || "",
      a: data.rowReps ? own(w.row_author_rep) : pickRep(w.row_author_rep, data.obj.author_rep),
      s: isNone(w.sub_rep) ? "" : clean(w.sub_rep)
    };
  }
  function uniq(list) { var o = []; list.forEach(function (x) { if (x && o.indexOf(x) < 0) o.push(x); }); return o; }

  // данные для журнала из реестра (акты с ошибками не попадают)
  function journalFromRegistry(data, check) {
    var st = { title: {}, meta: { created: new Date().toISOString(), v: 4, source: "Генератор АОСР" }, bj: {}, uh: {}, sv: {}, ak: {},
      s1: [], s2: [], s3: [], s4: [], s5: [] };
    TITLE_KEYS.forEach(function (k) { st.title[k] = ""; });
    var ok = data.works.filter(function (w) { return !check.errors[w.id]; });
    var reps = ok.map(function (w) { return repsOf(data, w); });
    st.title.object = clean(data.obj.object);
    st.title.contractor = clean(data.obj.contractor);
    st.title.designer = clean(data.obj.designer);
    st.title.technadzor = uniq(reps.map(function (r) { return r.t; })).join("; ");
    st.title.avtornadzor = uniq(reps.map(function (r) { return r.a; })).join("; ");
    st.title.subcontractors = uniq(reps.map(function (r) { return parseRep(r.s).org; })).join("; ");
    // раздел 1 — ИТР подрядчика и субподрядчиков (надзор — в титуле, не в ИТР)
    var seen = {};
    ok.forEach(function (w, i) {
      [reps[i].c, reps[i].s].forEach(function (rep) {
        if (!rep) return;
        var p = parseRep(rep), k = norm(p.fio), s = toDate(w.start);
        if (!seen[k]) { seen[k] = { fio: p.fio, pos: p.pos, org: p.org || clean(w.performer) || clean(data.obj.contractor), area: "", d1: "", d2: "", att: p.att }; st.s1.push(seen[k]); }
        if (s && (!seen[k].d1 || iso(s) < seen[k].d1)) seen[k].d1 = iso(s);
      });
    });
    // раздел 3 — перечень актов
    ok.forEach(function (w, i) {
      if (!clean(w.number) || isBN(w.number)) return;
      var r = reps[i];
      st.s3.push({ no: clean(w.number), name: (clean(w.title) + " " + clean(w.location)).trim() + (clean(w.place) ? " (" + clean(w.place) + ")" : ""),
        date: iso(toDate(w.date)), who: [shortName(r.c) && shortName(r.c) + " (СМО)", shortName(r.t) && shortName(r.t) + " (ТН)",
          shortName(r.a) && shortName(r.a) + " (АН)"].filter(Boolean).join(", "), note: "" });
    });
    // раздел 4 — работы с датами (одна запись на дату начала; период — в тексте)
    var noDate = 0;
    ok.forEach(function (w, i) {
      var s = toDate(w.start) || toDate(w.end), e = toDate(w.end);
      if (!s) { noDate++; return; }
      var work = clean(w.title) + (e && key(e) !== key(s) ? " (работы с " + fmtShort(s) + " по " + fmtShort(e) + ")" : "");
      st.s4.push({ date: iso(s), shift: "", jr: guessJr(w.title), aosr: clean(w.number) && !isBN(w.number) ? clean(w.number) : "",
        work: work, place: [clean(w.place), clean(w.location)].filter(Boolean).join(", "),
        unit: "", vol: "", temp: "", weather: "", exec: clean(w.performer) || clean(data.obj.contractor), resp: shortName(reps[i].c) });
    });
    st.s4.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    return { state: st, noDate: noDate };
  }

  // дополнить существующий журнал: только недостающее, ничего не перезаписывать
  function mergeJournal(old, fresh) {
    var st = JSON.parse(JSON.stringify(old)), add = { s1: 0, s3: 0, s4: 0, title: 0 }, notes = [];
    st.title = st.title || {};
    TITLE_KEYS.forEach(function (k) { if (!clean(st.title[k]) && clean(fresh.title[k])) { st.title[k] = fresh.title[k]; add.title++; } });
    ["s1", "s2", "s3", "s4", "s5"].forEach(function (k) { if (!Array.isArray(st[k])) st[k] = []; });
    fresh.s1.forEach(function (p) {
      if (!st.s1.some(function (x) { return norm(x.fio) === norm(p.fio); })) { st.s1.push(p); add.s1++; }
    });
    fresh.s3.forEach(function (a) {
      var ex = st.s3.filter(function (x) { return norm(x.no) === norm(a.no); })[0];
      if (!ex) { st.s3.push(a); add.s3++; return; }
      if (!ex.date && a.date) ex.date = a.date;
      else if (ex.date && a.date && ex.date !== a.date)
        notes.push("акт № " + a.no + ": в журнале дата " + fmtShort(toDate(ex.date)) + ", в реестре " + fmtShort(toDate(a.date)));
    });
    var k4 = function (r) { return r.date + "|" + norm(String(r.work).replace(/\s*\(работы с .*\)$/, "")) + "|" + norm(r.place); };
    var have = {}; st.s4.forEach(function (r) { have[k4(r)] = r; });
    add.link = 0;
    fresh.s4.forEach(function (r) {
      var ex = have[k4(r)];
      if (!ex) { st.s4.push(r); add.s4++; return; }
      if (!clean(ex.aosr) && r.aosr) { ex.aosr = r.aosr; add.link++; }   // связь с актом; ручной номер не трогаем
    });
    st.s4.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    st.meta = st.meta || {}; st.meta.updatedFromRegistry = new Date().toISOString();
    return { state: st, add: add, notes: notes };
  }

  // сверка: за период работ по акту в журнале должны быть записи
  function crossCheck(old, data, check) {
    var lines = [], recs = (old.s4 || []).map(function (r) { return isoKey(r.date); }).filter(Boolean);
    data.works.forEach(function (w) {
      if (check.errors[w.id] || !clean(w.number) || isBN(w.number)) return;
      var s = toDate(w.start), e = toDate(w.end) || s;
      if (!s) return;
      var has = recs.some(function (k) { return k >= key(s) && k <= key(e); });
      if (!has) lines.push("акт № " + clean(w.number) + ": в журнале нет записей за период работ " + fmtShort(s) + (e && key(e) !== key(s) ? "–" + fmtShort(e) : ""));
    });
    var regNos = {}; data.works.forEach(function (w) { if (clean(w.number)) regNos[norm(w.number)] = 1; });
    (old.s3 || []).forEach(function (a) { if (a.no && !regNos[norm(a.no)]) lines.push("акт № " + a.no + " есть в журнале, но нет в реестре"); });
    return lines;
  }

  // журнал входного контроля — Word, альбомный лист
  function vhkDocx(data, tplBytes) {
    var list = Object.keys(data.materials).map(function (k) { return Object.assign({ code: k }, data.materials[k]); });
    list.sort(function (a, b) { return (a.recv ? key(a.recv) : 0) - (b.recv ? key(b.recv) : 0) || (a.code < b.code ? -1 : 1); });
    var warns = [];
    list.forEach(function (m) {
      if (!m.recv) warns.push(m.code + ": нет даты поступления");
      if (!m.invoice) warns.push(m.code + ": нет № накладной");
      if (m.until && m.recv && key(m.until) < key(m.recv)) warns.push(m.code + ": документ о качестве истёк " + fmtShort(m.until) + " — до поступления " + fmtShort(m.recv));
      if (!clean(m.control)) warns.push(m.code + ": не заполнен результат входного контроля");
    });
    var body = xp("ЖУРНАЛ", { bold: true, sz: 28, after: 0 }) +
      xp("входного контроля материалов, изделий и конструкций", { bold: true, sz: 24, after: 120 }) +
      xp("Объект: " + clean(data.obj.object), { after: 40 }) +
      xp("Строительно-монтажная организация: " + clean(data.obj.contractor), { after: 160 }) +
      xtable([500, 1400, 3400, 650, 900, 2400, 3150, 1350, 1150],
        ["№ п/п", "Дата поступления", "Наименование материала, изделия, конструкции", "Ед. изм.", "Кол-во",
          "Поставщик, № накладной", "Документ о качестве (вид, №, дата, срок действия)", "Результат входного контроля", "Подпись"],
        list.map(function (m, i) {
          return [String(i + 1), m.recv ? fmtShort(m.recv) : "", m.name + " [" + m.code + "]", m.unit || "", m.qty === null || m.qty === undefined ? "" : String(m.qty).replace(".", ","),
            [m.supplier, m.invoice && "накл. № " + m.invoice].filter(Boolean).join(", "),
            [m.doc_type, m.doc_no && "№ " + m.doc_no, m.date && "от " + fmtShort(m.date), m.until && "действует до " + fmtShort(m.until)].filter(Boolean).join(" "),
            clean(m.control), ""];
        }), [0, 1, 3, 4]);
    var z = new PizZip(tplBytes), x = z.file("word/document.xml").asText();
    var start = x.indexOf(">", x.indexOf("<w:body")) + 1, end = x.lastIndexOf("</w:body>");
    var sect = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
      '<w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>';
    z.file("word/document.xml", x.slice(0, start) + body + sect + x.slice(end));
    return { bytes: z.generate({ type: "uint8array", compression: "DEFLATE" }), count: list.length, warns: warns };
  }

  var customTemplate = null, last = null;
  var DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  var XLSXT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  function showReport(lines, state) {
    var box = $("report");
    box.hidden = false;
    box.className = "report " + state;
    box.innerHTML = "";
    lines.forEach(function (l) {
      var div = document.createElement("div");
      div.textContent = l || " ";
      if (/✗/.test(l)) div.className = "e"; else if (/^\s+!/.test(l)) div.className = "w";
      else if (/^(Строка|Лист)/.test(l)) div.className = "h";
      box.appendChild(div);
    });
  }

  function tellPult(msg) { if (window.parent !== window) window.parent.postMessage(Object.assign({ tool: "aosr" }, msg), "*"); }
  var JR_KEY = "zpr-journal-v2";   // ключ хранения журнала производства работ (templates/zhurnal-proizvodstva-rabot)

  // дописать данные реестра в журнал производства работ (общее хранилище браузера в пульте)
  function pushJournal(auto) {
    if (!last) return;
    var old = null;
    try { old = JSON.parse(localStorage.getItem(JR_KEY)); } catch (e) {}
    var out, issues = [];
    if (old && old.title) {
      var m = mergeJournal(old, last.jr.state);
      issues = crossCheck(old, last.data, last.check).concat(m.notes);
      try { localStorage.setItem(JR_KEY, JSON.stringify(m.state)); } catch (e) { showReport(last.lines.concat(["", "  ✗ Не удалось записать журнал в браузер"]), "bad"); return; }
      out = ["", (auto ? "Журнал работ обновлён автоматически: ИТР +" : "Журнал дополнен: ИТР +") + m.add.s1 + ", акты +" + m.add.s3 + ", записи о работах +" + m.add.s4 +
        ", поля титула +" + m.add.title + (m.add.link ? ", связано записей с актами +" + m.add.link : "") + ". Существующие записи не изменены."];
    } else {
      try { localStorage.setItem(JR_KEY, JSON.stringify(last.jr.state)); } catch (e) { showReport(last.lines.concat(["", "  ✗ Не удалось записать журнал в браузер"]), "bad"); return; }
      out = ["", "Журнал создан из реестра: ИТР — " + last.jr.state.s1.length + ", актов — " + last.jr.state.s3.length +
        ", записей о работах — " + last.jr.state.s4.length + "."];
    }
    if (issues.length) out = out.concat(["", "Сверка журнала с реестром:"], issues.map(function (x) { return "  ! " + x; }));
    else if (old && old.title) out.push("Сверка журнала с реестром: расхождений нет.");
    out.push("Журнал — на вкладке «Журнал работ»" + (auto ? "; повторить передачу — кнопка «Передать в журнал работ»." : "."));
    showReport(last.lines.concat(out), issues.length ? "mixed" : "ok");
    tellPult({ type: "pto-journal-updated" });
  }

  function process(file, opts) {
    opts = opts || {};
    var reader = new FileReader();
    reader.onload = function () {
      var handled = false;
      try {
        var wb = XLSX.read(new Uint8Array(reader.result), { type: "array" });
        if (opts.pult && !wb.Sheets["Работы"]) { tellPult({ type: "pto-ignored" }); return; }
        tellPult({ type: "pto-handled" }); handled = true;
        $("extra").hidden = true; last = null;
        var data = load(wb);
        var stats = prepare(data);
        var check = validate(data);
        if (check.errors.obj) {
          showReport(["Акты не сформированы — исправьте лист «Объект» или структуру реестра:"].concat(check.errors.obj.map(function (m) { return "  ✗ " + m; })), "bad");
          return;
        }
        var tpl = customTemplate || b64ToBytes(window.AOSR_TEMPLATE_B64);
        var acts = render(data, check, tpl);
        var lines = reportLines(data, check, acts, file.name, stats);
        var hasErr = lines.some(function (l) { return /✗/.test(l); });
        showReport(lines, !acts.length ? "bad" : hasErr ? "mixed" : "ok");
        if (!acts.length) return;
        var vhk = vhkDocx(data, tpl), jr = journalFromRegistry(data, check);
        last = {
          kit: buildKit(data.obj, acts), acts: acts,
          reg: registerXlsx(data.obj, acts), report: "﻿" + lines.join("\r\n"),
          filled: filledRegistry(wb, data), vhk: vhk, jr: jr, data: data, check: check, lines: lines
        };
        download(last.kit, "Комплект_АОСР.docx", DOCX);
        $("extra").hidden = false;
        if (window.parent !== window) pushJournal(true);   // в пульте журнал обновляется сам
      } catch (e) {
        if (opts.pult && !handled) { tellPult({ type: "pto-ignored" }); return; }
        var msg = e && e.properties && e.properties.errors
          ? e.properties.errors.map(function (x) { return x.properties.explanation; }).join("; ")
          : (e && e.message) || String(e);
        showReport(["Не удалось обработать файл: " + msg], "bad");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  document.addEventListener("DOMContentLoaded", function () {
    var drop = $("drop"), input = $("file");
    input.addEventListener("change", function () { if (input.files[0]) process(input.files[0]); input.value = ""; });
    ["dragenter", "dragover"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add("over"); }); });
    ["dragleave", "drop"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove("over"); }); });
    drop.addEventListener("drop", function (e) { if (e.dataTransfer.files[0]) process(e.dataTransfer.files[0]); });
    $("dl-kit").addEventListener("click", function () { if (last) download(last.kit, "Комплект_АОСР.docx", DOCX); });
    $("dl-reg").addEventListener("click", function () { if (last) download(last.reg, "Реестр_ИД.xlsx", XLSXT); });
    $("dl-rep").addEventListener("click", function () { if (last) download(new TextEncoder().encode(last.report), "Отчёт_проверки.txt", "text/plain;charset=utf-8"); });
    // в пульте: дописать данные реестра прямо в журнал на соседней вкладке (без .json)
    if (window.parent !== window) $("jr-push").hidden = false;
    $("jr-push").addEventListener("click", function () { pushJournal(); });
    window.addEventListener("message", function (e) {
      if (e.data && e.data.type === "pto-file") process(new File([e.data.buf], e.data.name), { pult: true });
    });
    $("dl-vhk").addEventListener("click", function () {
      if (!last) return;
      download(last.vhk.bytes, "Журнал_входного_контроля.docx", DOCX);
      showReport(last.lines.concat(["", "Журнал входного контроля: " + last.vhk.count + " партий"],
        last.vhk.warns.map(function (w) { return "  ! " + w; })), last.vhk.warns.length ? "mixed" : "ok");
    });
    $("dl-jr").addEventListener("click", function () {
      if (!last) return;
      var st = last.jr.state;
      download(new TextEncoder().encode(JSON.stringify(st, null, 1)), "журнал_работ_из_реестра.json", "application/json");
      showReport(last.lines.concat(["", "Для журнала производства работ: ИТР — " + st.s1.length + ", актов — " + st.s3.length +
        ", записей о работах — " + st.s4.length + (last.jr.noDate ? " (без дат пропущено: " + last.jr.noDate + ")" : ""),
        "Откройте файл в журнале кнопкой «Открыть файл»."]), "ok");
    });
    $("jr-upd").addEventListener("change", function () {
      var f = $("jr-upd").files[0]; $("jr-upd").value = "";
      if (!f || !last) return;
      var r = new FileReader();
      r.onload = function () {
        var old;
        try { old = JSON.parse(r.result); if (!old || !old.title) throw 0; }
        catch (e) { showReport(last.lines.concat(["", "  ✗ " + f.name + " — не файл журнала производства работ (.json)"]), "bad"); return; }
        var cc = crossCheck(old, last.data, last.check), m = mergeJournal(old, last.jr.state);
        download(new TextEncoder().encode(JSON.stringify(m.state, null, 1)), f.name.replace(/\.json$/i, "") + "_обновлён.json", "application/json");
        var out = last.lines.concat(["", "Журнал «" + f.name + "» дополнен: ИТР +" + m.add.s1 + ", акты +" + m.add.s3 +
          ", записи о работах +" + m.add.s4 + ", поля титула +" + m.add.title + (m.add.link ? ", связано записей с актами +" + m.add.link : "") +
          ". Существующие записи не изменены."]);
        var issues = cc.concat(m.notes);
        if (issues.length) out = out.concat(["", "Сверка журнала с реестром:"], issues.map(function (x) { return "  ! " + x; }));
        else out.push("Сверка журнала с реестром: расхождений нет.");
        showReport(out, issues.length ? "mixed" : "ok");
      };
      r.readAsText(f);
    });
    $("dl-filled").addEventListener("click", function () { if (last) download(last.filled, "реестр_работ_заполненный.xlsx", XLSXT); });
    $("dl-acts").addEventListener("click", function () {
      if (!last) return;
      last.acts.forEach(function (a, i) { setTimeout(function () { download(a.bytes, a.name, DOCX); }, i * 400); });
    });
    $("sample").addEventListener("click", function () { download(b64ToBytes(window.AOSR_REGISTRY_B64), "реестр_работ.xlsx", XLSXT); });
    $("tpl-download").addEventListener("click", function () { download(b64ToBytes(window.AOSR_TEMPLATE_B64), "шаблон_АОСР.docx", DOCX); });
    $("tpl").addEventListener("change", function () {
      var f = $("tpl").files[0]; if (!f) return;
      var r = new FileReader();
      r.onload = function () { customTemplate = new Uint8Array(r.result); $("tpl-name").textContent = "· свой шаблон: " + f.name; };
      r.readAsArrayBuffer(f);
    });
  });
})();

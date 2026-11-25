/* Формы 2 и 3 из план-факта — работает полностью в браузере, офлайн. */
(function () {
  "use strict";
  var MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"];
  var MONEY = '#,##0.00', QTY = '#,##0.000';
  var SMETA_SPEC = [["pos", /^№/], ["section", /^раздел/i], ["code", /^шифр/i], ["name", /^наимен/i],
    ["unit", /^ед/i], ["qty", /^кол/i], ["price", /^цена/i]];
  var WORK_SPEC = [["date", /^дата/i], ["pos", /^№ ?поз/i], ["block", /^блок/i], ["floor", /^этаж/i],
    ["qty", /^объ[её]м/i], ["aosr", /аоср/i]];

  function $(id) { return document.getElementById(id); }
  function clean(v) { return v === null || v === undefined ? "" : String(v).trim(); }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function fnum(x, d) {
    d = d === undefined ? 2 : d;
    var s = Math.abs(x).toFixed(d).split("."), i = s[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
    return (x < 0 ? "−" : "") + i + (d ? "," + s[1] : "");
  }

  // значение ячейки ExcelJS: формула -> результат, rich text -> текст
  function val(c) {
    var v = c && c.value;
    if (v && typeof v === "object" && !(v instanceof Date)) {
      if ("result" in v) return v.result;
      if (v.richText) return v.richText.map(function (t) { return t.text; }).join("");
      if (v.text) return v.text;
    }
    return v === undefined ? null : v;
  }
  function toDate(v) {
    if (v === null || v === undefined || v === "") return null;
    if (v instanceof Date && !isNaN(v)) return { y: v.getUTCFullYear(), m: v.getUTCMonth() + 1, d: v.getUTCDate() };
    if (typeof v === "number") { var dt = new Date(Math.round((v - 25569) * 864e5)); return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() }; }
    var s = String(v).trim(), m;
    if ((m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/))) return { y: +m[3], m: +m[2], d: +m[1] };
    if ((m = s.match(/^(\d{1,2})\.(\d{4})$/))) return { y: +m[2], m: +m[1], d: 1 };
    if ((m = s.match(/^(\d{4})-(\d{2})(?:-(\d{2}))?/))) return { y: +m[1], m: +m[2], d: +(m[3] || 1) };
    return null;
  }
  function num(v) {
    if (typeof v === "number") return v;
    if (v === null || v === undefined || v === "") return null;
    var n = parseFloat(String(v).replace(/\s/g, "").replace(",", "."));
    return isNaN(n) ? null : n;
  }
  function mk(d) { return d.y * 100 + d.m; }

  function sheetRows(wb, name) {
    var ws = wb.getWorksheet(name);
    if (!ws) throw new Error("в файле нет листа «" + name + "»");
    var out = [];
    ws.eachRow({ includeEmpty: false }, function (row, rn) {
      var arr = [];
      for (var c = 1; c <= Math.max(row.cellCount, 12); c++) arr.push(val(row.getCell(c)));
      arr.rn = rn;
      out.push(arr);
    });
    return out;
  }
  function mapHeaders(header, spec) {
    var map = {}, used = {};
    spec.forEach(function (sp) {
      for (var i = 0; i < header.length; i++) if (!used[i] && sp[1].test(clean(header[i]))) { map[sp[0]] = i; used[i] = true; return; }
    });
    return map;
  }

  // ---------------------------------------------------------- загрузка и проверка
  function load(wb) {
    var settings = {};
    sheetRows(wb, "Настройки").slice(1).forEach(function (r) { if (clean(r[0])) settings[clean(r[0])] = r[1]; });

    var sr = sheetRows(wb, "Смета"), smap = mapHeaders(sr[0] || [], SMETA_SPEC), smeta = {}, order = [];
    sr.slice(1).forEach(function (r) {
      var pos = num(r[smap.pos]);
      if (pos === null) return;
      smeta[pos] = { pos: pos, section: clean(r[smap.section]), code: clean(r[smap.code]), name: clean(r[smap.name]),
        unit: clean(r[smap.unit]), qty: num(r[smap.qty]) || 0, price: num(r[smap.price]) || 0 };
      order.push(pos);
    });

    var wr = sheetRows(wb, "Выполнение"), wmap = mapHeaders(wr[0] || [], WORK_SPEC), rows = [];
    wr.slice(1).forEach(function (r) {
      var g = function (f) { return wmap[f] === undefined ? null : r[wmap[f]]; };
      if ([g("date"), g("pos"), g("block"), g("floor"), g("qty"), g("aosr")].every(function (v) { return clean(v) === ""; })) return;
      rows.push({ row: r.rn, date: toDate(g("date")), dateRaw: g("date"), pos: num(g("pos")), posRaw: g("pos"),
        block: clean(g("block")), floor: clean(g("floor")), qty: num(g("qty")), qtyRaw: g("qty"), aosr: clean(g("aosr")) });
    });
    return { settings: settings, smeta: smeta, order: order, rows: rows };
  }

  function validate(data) {
    var errors = [], warns = [], st = data.settings;
    var period = toDate(st["Отчётный период"]);
    if (!period) errors.push("Настройки: не задан или не распознан «Отчётный период»");
    var vat = num(st["Ставка НДС"]);
    if (vat === null) errors.push("Настройки: не задана «Ставка НДС»");
    else if (vat > 1) vat = vat / 100;
    if (!data.order.length) errors.push("Смета: нет ни одной позиции");

    var good = [];
    data.rows.forEach(function (r) {
      var where = "Выполнение, строка " + r.row, bad = false;
      if (!r.date) { errors.push(where + ": не распознана дата «" + clean(r.dateRaw) + "»"); bad = true; }
      if (r.pos === null || !data.smeta[r.pos]) { errors.push(where + ": позиции № " + (clean(r.posRaw) || "—") + " нет в смете"); bad = true; }
      if (r.qty === null || r.qty <= 0) { errors.push(where + ": объём должен быть больше нуля (сейчас «" + clean(r.qtyRaw) + "»)"); bad = true; }
      if (!r.block || !r.floor) warns.push(where + ": не указан блок или этаж");
      if (!r.aosr) warns.push(where + ": не указан № АОСР");
      if (!bad) good.push(r);
    });

    if (period) {
      var totals = {};
      good.forEach(function (r) { if (mk(r.date) <= mk(period)) totals[r.pos] = (totals[r.pos] || 0) + r.qty; });
      Object.keys(totals).forEach(function (p) {
        var s = data.smeta[p], q = totals[p];
        if (q > s.qty + 1e-9) errors.push("Позиция " + p + " «" + s.name + "»: выполнено " + fnum(q, 3) + " " + s.unit +
          " при смете " + fnum(s.qty, 3) + " — перевыполнение " + fnum(q - s.qty, 3) + ". Нужна доп. смета или исправление объёмов");
      });
    }
    return { period: period, vat: vat, good: good, errors: errors, warns: warns };
  }

  function compute(data, chk) {
    var pk = mk(chk.period), acc = {}, detail = [];
    data.order.forEach(function (p) { acc[p] = { total: 0, year: 0, month: 0 }; });
    chk.good.forEach(function (r) {
      var k = mk(r.date);
      if (k > pk) return;  // будущие периоды не учитываем
      var a = acc[r.pos];
      a.total += r.qty;
      if (r.date.y === chk.period.y) a.year += r.qty;
      if (k === pk) { a.month += r.qty; detail.push(r); }
    });
    return { acc: acc, detail: detail };
  }

  // ---------------------------------------------------------- Excel с оформлением
  var THIN = { style: "thin", color: { argb: "FF000000" } };
  var BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  var HEAD_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };
  var FONT = "Times New Roman";

  function titleBlock(ws, st, title, period, ncols) {
    var lines = [[title, { bold: true, size: 13 }], ["за " + MONTHS[period.m - 1] + " " + period.y + " г.", { bold: true, size: 11 }],
      ["Объект: " + clean(st["Объект"])], ["Заказчик: " + clean(st["Заказчик"])], ["Подрядчик: " + clean(st["Подрядчик"])],
      ["Договор: " + clean(st["Договор"])]];
    lines.forEach(function (l, i) {
      var r = i + 1;
      ws.mergeCells(r, 1, r, ncols);
      var c = ws.getCell(r, 1);
      c.value = l[0];
      c.font = Object.assign({ name: FONT, size: 10 }, l[1] || {});
      c.alignment = { horizontal: i < 2 ? "center" : "left", wrapText: true };
    });
    return lines.length + 2;
  }
  function tableHead(ws, row, groups, widths) {
    var col = 1;
    groups.forEach(function (g) {
      if (g[1].length) {
        ws.mergeCells(row, col, row, col + g[1].length - 1);
        ws.getCell(row, col).value = g[0];
        g[1].forEach(function (s, k) { ws.getCell(row + 1, col + k).value = s; });
        col += g[1].length;
      } else {
        ws.mergeCells(row, col, row + 1, col);
        ws.getCell(row, col).value = g[0];
        col += 1;
      }
    });
    for (var r = row; r <= row + 1; r++) for (var c = 1; c < col; c++) {
      var cell = ws.getCell(r, c);
      cell.font = { name: FONT, bold: true, size: 9 }; cell.fill = HEAD_FILL; cell.border = BORDER;
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    }
    widths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    return row + 2;
  }
  function putRow(ws, r, values, fmts, bold) {
    values.forEach(function (v, i) {
      var c = ws.getCell(r, i + 1);
      c.value = v === "" ? null : v;
      c.border = BORDER;
      c.font = { name: FONT, size: 10, bold: !!bold };
      c.alignment = { vertical: "top", wrapText: true };
      if (fmts && fmts[i + 1]) c.numFmt = fmts[i + 1];
    });
  }
  function vatLines(ws, r, ncols, labelCol, sums, vat) {
    [["Итого без НДС", 1], ["НДС " + Math.round(vat * 100) + "%", vat], ["Всего с НДС", 1 + vat]].forEach(function (l) {
      var vals = [], fm = {};
      for (var i = 0; i < ncols; i++) vals.push("");
      vals[labelCol - 1] = l[0];
      sums.forEach(function (s) { vals[s[0] - 1] = Math.round(s[1] * l[1] * 100) / 100; fm[s[0]] = MONEY; });
      putRow(ws, r, vals, fm, true);
      r++;
    });
    return r;
  }
  function signatures(ws, r) {
    r++;
    ["Сдал: представитель подрядчика", "Принял: представитель заказчика"].forEach(function (t) {
      ws.getCell(r, 2).value = t; ws.getCell(r, 2).font = { name: FONT, bold: true, size: 10 };
      ws.getCell(r, 5).value = "_______________ / _______________ /"; ws.getCell(r, 5).font = { name: FONT, size: 10 };
      r += 2;
    });
  }
  function landscape(ws) {
    ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      horizontalCentered: true, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
  }

  function form2(wb, data, res, chk) {
    var ws = wb.addWorksheet("Форма 2");
    var groups = [["№ п/п", []], ["№ поз. по смете", []], ["Шифр", []], ["Наименование работ", []], ["Ед. изм.", []],
      ["Цена за ед., тг", []], ["По смете", ["кол-во", "стоимость, тг"]],
      ["Выполнено с начала строительства", ["кол-во", "стоимость, тг"]], ["в т.ч. за отчётный период", ["кол-во", "стоимость, тг"]]];
    var widths = [6, 8, 13, 46, 7, 13, 11, 15, 11, 15, 11, 15];
    var r = titleBlock(ws, data.settings, "АКТ ВЫПОЛНЕННЫХ РАБОТ (форма 2)", chk.period, widths.length);
    var headRow = r;
    r = tableHead(ws, r, groups, widths);
    var fm = { 6: MONEY, 7: QTY, 8: MONEY, 9: QTY, 10: MONEY, 11: QTY, 12: MONEY };
    var n = 0, tS = 0, tT = 0, tM = 0;
    data.order.forEach(function (p) {
      var s = data.smeta[p], a = res.acc[p];
      if (!a.month) return;  // в форму 2 — только позиции, выполненные в отчётном месяце
      n++;
      var vs = s.qty * s.price, vt = a.total * s.price, vm = a.month * s.price;
      tS += vs; tT += vt; tM += vm;
      putRow(ws, r++, [n, p, s.code, s.name, s.unit, s.price, s.qty, round2(vs), a.total, round2(vt), a.month, round2(vm)], fm);
    });
    r = vatLines(ws, r, widths.length, 4, [[8, tS], [10, tT], [12, tM]], chk.vat);
    signatures(ws, r);
    landscape(ws);
    ws.pageSetup.printTitlesRow = headRow + ":" + (headRow + 1);
    return { n: n, month: tM };
  }

  function form3(wb, data, res, chk) {
    var ws = wb.addWorksheet("Форма 3");
    var groups = [["№ п/п", []], ["Наименование разделов и работ", []], ["Сметная стоимость, тг", []],
      ["Стоимость выполненных работ, тг", ["с начала строительства", "с начала года", "за отчётный месяц"]]];
    var widths = [6, 52, 18, 18, 18, 18];
    var r = titleBlock(ws, data.settings, "СПРАВКА О СТОИМОСТИ ВЫПОЛНЕННЫХ РАБОТ (форма 3)", chk.period, widths.length);
    r = tableHead(ws, r, groups, widths);
    var sec = {}, order = [];
    data.order.forEach(function (p) {
      var s = data.smeta[p], a = res.acc[p], k = s.section || "Без раздела";
      if (!sec[k]) { sec[k] = [0, 0, 0, 0]; order.push(k); }
      sec[k][0] += s.qty * s.price; sec[k][1] += a.total * s.price; sec[k][2] += a.year * s.price; sec[k][3] += a.month * s.price;
    });
    var tot = [0, 0, 0, 0], fm = { 3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY };
    order.forEach(function (k, i) {
      var v = sec[k];
      tot = tot.map(function (x, j) { return x + v[j]; });
      putRow(ws, r++, [i + 1, k].concat(v.map(round2)), fm);
    });
    r = vatLines(ws, r, widths.length, 2, [[3, tot[0]], [4, tot[1]], [5, tot[2]], [6, tot[3]]], chk.vat);
    signatures(ws, r);
    landscape(ws);
    return tot;
  }

  function breakdown(wb, data, res, chk) {
    var ws = wb.addWorksheet("Расшифровка по участкам");
    var groups = [["№ поз.", []], ["Наименование работ", []], ["Блок", []], ["Этаж", []], ["Ед. изм.", []],
      ["Объём", []], ["Стоимость, тг", []], ["№ АОСР", []]];
    var widths = [7, 50, 10, 9, 7, 12, 15, 18];
    var r = titleBlock(ws, data.settings, "РАСШИФРОВКА ОБЪЁМОВ ПО УЧАСТКАМ", chk.period, widths.length);
    r = tableHead(ws, r, groups, widths);
    var fm = { 6: QTY, 7: MONEY }, byPos = {};
    res.detail.forEach(function (d) { (byPos[d.pos] = byPos[d.pos] || []).push(d); });
    data.order.forEach(function (p) {
      if (!byPos[p]) return;
      var s = data.smeta[p], q = 0;
      byPos[p].sort(function (a, b) { return (a.block + "|" + a.floor).localeCompare(b.block + "|" + b.floor, "ru", { numeric: true }); })
        .forEach(function (d) {
          q += d.qty;
          putRow(ws, r++, [p, s.name, d.block, d.floor, s.unit, d.qty, round2(d.qty * s.price), d.aosr || "— нет акта —"], fm);
        });
      putRow(ws, r++, [p, "Итого по позиции " + p, "", "", s.unit, q, round2(q * s.price), ""], fm, true);
    });
    landscape(ws);
  }
  function round2(x) { return Math.round(x * 100) / 100; }

  // ---------------------------------------------------------- UI
  var last = null;
  var XLSXT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  function download(bytes, name, type) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([bytes], { type: type })); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function showReport(lines, state) {
    var box = $("report");
    box.hidden = false; box.className = "report " + state; box.innerHTML = "";
    lines.forEach(function (l) {
      var div = document.createElement("div");
      div.textContent = l || " ";
      if (/✗/.test(l)) div.className = "e"; else if (/^\s+!/.test(l)) div.className = "w";
      else if (/^(ОШИБКИ|Предупреждения)/.test(l)) div.className = "h";
      box.appendChild(div);
    });
  }

  function process(file) {
    $("extra").hidden = true; last = null;
    var reader = new FileReader();
    reader.onload = function () {
      var wb = new ExcelJS.Workbook();
      wb.xlsx.load(reader.result).then(function () {
        var data = load(wb), chk = validate(data);
        var now = new Date();
        var lines = ["Проверка " + file.name + " — " + pad(now.getDate()) + "." + pad(now.getMonth() + 1) + "." + now.getFullYear() +
          " " + pad(now.getHours()) + ":" + pad(now.getMinutes())];
        if (chk.period) lines.push("Отчётный период: " + MONTHS[chk.period.m - 1] + " " + chk.period.y + ", НДС " + Math.round((chk.vat || 0) * 100) + "%");
        if (chk.errors.length) {
          lines.push("", "ОШИБКИ — формы не сформированы, исправьте и загрузите снова:");
          chk.errors.forEach(function (e) { lines.push("  ✗ " + e); });
        }
        if (chk.warns.length) { lines.push("", "Предупреждения:"); chk.warns.forEach(function (w) { lines.push("  ! " + w); }); }
        if (chk.errors.length) { showReport(lines, "bad"); return; }

        var res = compute(data, chk), out = new ExcelJS.Workbook();
        out.creator = "Генератор форм 2 и 3";
        var f2 = form2(out, data, res, chk), tot = form3(out, data, res, chk);
        breakdown(out, data, res, chk);
        lines.push("", "Форма 2: позиций — " + f2.n + ", за месяц без НДС — " + fnum(tot[3]) + " тг, с НДС — " + fnum(tot[3] * (1 + chk.vat)) + " тг",
          "С начала строительства без НДС — " + fnum(tot[1]) + " тг из " + fnum(tot[0]) + " тг по смете" +
          (tot[0] ? " (" + (tot[1] / tot[0] * 100).toFixed(1).replace(".", ",") + "%)" : ""));
        if (!f2.n) lines.push("  ! В отчётном месяце нет выполненных объёмов — форма 2 пустая");
        showReport(lines, chk.warns.length || !f2.n ? "mixed" : "ok");
        var name = "Формы_2_3_" + chk.period.y + "-" + pad(chk.period.m) + ".xlsx";
        return out.xlsx.writeBuffer().then(function (buf) {
          last = { bytes: buf, name: name, report: "﻿" + lines.join("\r\n") };
          download(buf, name, XLSXT);
          $("extra").hidden = false;
        });
      }).catch(function (e) {
        showReport(["Не удалось обработать файл: " + ((e && e.message) || e)], "bad");
      });
    };
    reader.readAsArrayBuffer(file);
  }

  function b64ToBytes(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  document.addEventListener("DOMContentLoaded", function () {
    var drop = $("drop"), input = $("file");
    input.addEventListener("change", function () { if (input.files[0]) process(input.files[0]); input.value = ""; });
    ["dragenter", "dragover"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add("over"); }); });
    ["dragleave", "drop"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove("over"); }); });
    drop.addEventListener("drop", function (e) { if (e.dataTransfer.files[0]) process(e.dataTransfer.files[0]); });
    $("dl-again").addEventListener("click", function () { if (last) download(last.bytes, last.name, XLSXT); });
    $("dl-rep").addEventListener("click", function () { if (last) download(new TextEncoder().encode(last.report), "Отчёт_проверки.txt", "text/plain;charset=utf-8"); });
    $("sample").addEventListener("click", function () { download(b64ToBytes(window.KS_SAMPLE_B64), "план_факт.xlsx", XLSXT); });
  });
})();

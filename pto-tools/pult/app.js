/* ПТО-пульт: вкладки с инструментами (каждый — свой HTML во фрейме), общая загрузка файлов, связь между вкладками. */
(function () {
  "use strict";
  var TOOLS = [
    // «Окно объекта» — генератор АОСР в режиме ввода данных (okno/okno.js): данные объекта, акты и журнал
    { id: "aosr", title: "Окно объекта", takesFiles: true, what: "реестр работ", name: "okno" },
    { id: "journal", title: "Журнал работ", takesFiles: false },
    { id: "forms", title: "Формы 2 и 3", takesFiles: true, what: "план-факт" }
  ];
  var frames = {}, buttons = {}, current = null, pending = null;

  function $(id) { return document.getElementById(id); }
  function decode(b64) {
    var bin = atob(b64), bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new TextDecoder().decode(bytes);
  }
  function status(text, kind) { var s = $("status"); s.textContent = text; s.className = kind || ""; }
  function remember(id) { try { localStorage.setItem("pto-pult-tab", id); } catch (e) {} }

  function show(id) {
    current = id;
    TOOLS.forEach(function (t) {
      if (!frames[t.id]) return;
      frames[t.id].hidden = t.id !== id;
      buttons[t.id].classList.toggle("active", t.id === id);
    });
    var dot = buttons[id] && buttons[id].querySelector(".dot");
    if (dot) dot.remove();
    remember(id);
  }

  function load(id) { frames[id].srcdoc = decode(window.PULT_APPS[id]); }

  function build() {
    TOOLS.forEach(function (t) {
      if (!window.PULT_APPS[t.id]) return;   // инструмент не собран — вкладки нет
      var f = document.createElement("iframe");
      f.title = t.title; f.hidden = true; if (t.name) f.name = t.name;
      $("frames").appendChild(f);
      frames[t.id] = f;
      var b = document.createElement("button");
      b.type = "button"; b.textContent = t.title;
      b.addEventListener("click", function () { show(t.id); });
      $("tabs").appendChild(b);
      buttons[t.id] = b;
      load(t.id);
    });
    var saved = null;
    try { saved = localStorage.getItem("pto-pult-tab"); } catch (e) {}
    show(frames[saved] ? saved : TOOLS[0].id);
  }

  // файл отправляется всем инструментам, принимающим файлы; берёт тот, кому он подходит
  function route(file) {
    var takers = TOOLS.filter(function (t) { return t.takesFiles && frames[t.id]; });
    status("Открываю «" + file.name + "»…");
    var reader = new FileReader();
    reader.onload = function () {
      pending = { name: file.name, left: takers.length, handled: false };
      takers.forEach(function (t) {
        frames[t.id].contentWindow.postMessage({ type: "pto-file", name: file.name, buf: reader.result.slice(0) }, "*");
      });
      setTimeout(function () {
        if (pending && !pending.handled) status("«" + file.name + "» не распознан: нужен реестр работ (лист «Работы») или план-факт (листы «Смета» и «Выполнение»).", "bad");
      }, 8000);
    };
    reader.readAsArrayBuffer(file);
  }

  window.addEventListener("message", function (e) {
    var d = e.data || {}, tool = TOOLS.filter(function (t) { return t.id === d.tool; })[0];
    if (!tool) return;
    if (d.type === "pto-handled" && pending) {
      pending.handled = true;
      show(tool.id);
      status("«" + pending.name + "» — " + tool.what + ", открыт во вкладке «" + tool.title + "».", "ok");
    } else if (d.type === "pto-ignored" && pending) {
      pending.left--;
      if (pending.left <= 0 && !pending.handled)
        status("«" + pending.name + "» не распознан: нужен реестр работ (лист «Работы») или план-факт (листы «Смета» и «Выполнение»).", "bad");
    } else if (d.type === "pto-journal-updated" && frames.journal) {
      load("journal");   // журнал перечитает данные из хранилища браузера
      if (d.quiet) return;   // окно объекта обновляет журнал при каждой правке — без сообщений
      if (current !== "journal" && !buttons.journal.querySelector(".dot")) {
        var dot = document.createElement("span"); dot.className = "dot"; dot.title = "обновлён";
        buttons.journal.appendChild(dot);
      }
      // сразу после загрузки реестра — дописываем к сообщению о файле, а не заменяем его
      if (pending && pending.handled && !pending.jr) { pending.jr = true; status($("status").textContent + " Журнал работ обновлён автоматически.", "ok"); }
      else status("Журнал работ обновлён данными реестра — вкладка «Журнал работ».", "ok");
    }
  });

  document.addEventListener("DOMContentLoaded", function () {
    build();
    var drop = $("drop"), input = $("file");
    input.addEventListener("change", function () { if (input.files[0]) route(input.files[0]); input.value = ""; });
    ["dragenter", "dragover"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add("over"); }); });
    ["dragleave", "drop"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove("over"); }); });
    drop.addEventListener("drop", function (e) { if (e.dataTransfer.files[0]) route(e.dataTransfer.files[0]); });
  });
})();

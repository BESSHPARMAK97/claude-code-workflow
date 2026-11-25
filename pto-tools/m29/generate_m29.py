"""
Макет отчёта о расходе основных материалов (по логике формы М-29).

Берёт объёмы выполненных работ из план-факта (../ks_generator/план_факт.xlsx),
нормы, поступление и остатки — из «материалы.xlsx», и формирует за отчётный месяц:
  М-29.xlsx  — Раздел 1: расход по нормам на выполненные объёмы;
               Раздел 2: сопоставление нормативного и фактического расхода по материалам;
               лист «По участкам»: нормативный расход по блокам и этажам.
  Отчёт_проверки.txt

Фактический расход = остаток на начало + поступило за месяц − фактический остаток на конец.

Запуск:  python generate_m29.py [путь_к_материалы.xlsx]
"""
import datetime as dt
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "материалы.xlsx"

MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
          "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
MONEY, QTY = '#,##0.00', '#,##0.000'
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
RED = PatternFill("solid", fgColor="F8CBAD")
GREEN = PatternFill("solid", fgColor="E2EFDA")


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for f in ("%d.%m.%Y", "%m.%Y", "%Y-%m-%d", "%Y-%m"):
            try:
                return dt.datetime.strptime(v.strip(), f).date()
            except ValueError:
                pass
    return None


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return None


def clean(v):
    return str(v).strip() if v is not None else ""


def fnum(x, digits=2):
    return f"{x:,.{digits}f}".replace(",", " ").replace(".", ",")


def mk(d):
    return (d.year, d.month)


def prev_mk(k):
    return (k[0] - 1, 12) if k[1] == 1 else (k[0], k[1] - 1)


# ---------------------------------------------------------------- загрузка
def load(errors, warns):
    wb = load_workbook(SRC, data_only=True)
    settings = {clean(r[0]): r[1] for r in wb["Настройки"].iter_rows(min_row=2, values_only=True) if r[0]}

    mats = {}
    for r in wb["Материалы"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            mats[clean(r[0])] = {"code": clean(r[0]), "name": clean(r[1]), "unit": clean(r[2]), "price": num(r[3]) or 0.0}

    norms = defaultdict(list)
    for i, r in enumerate(wb["Нормы"].iter_rows(min_row=2, values_only=True), start=2):
        if not any(r[:3]):
            continue
        pos, code, n = num(r[0]), clean(r[1]), num(r[2])
        if pos is None or n is None or n < 0:
            errors.append(f"Нормы, строка {i}: не заполнен № позиции или норма")
            continue
        if code not in mats:
            errors.append(f"Нормы, строка {i}: материала «{code}» нет на листе «Материалы»")
            continue
        norms[int(pos)].append((code, n))

    receipts = []
    for i, r in enumerate(wb["Поступление"].iter_rows(min_row=2, values_only=True), start=2):
        if not any(r[:3]):
            continue
        d, code, q = to_date(r[0]), clean(r[1]), num(r[2])
        if d is None or q is None or code not in mats:
            errors.append(f"Поступление, строка {i}: проверьте дату, код материала и количество")
            continue
        if not clean(r[3]):
            warns.append(f"Поступление, строка {i}: {mats[code]['name']} — нет № накладной")
        receipts.append({"date": d, "code": code, "qty": q, "doc": clean(r[3])})

    balances = {}
    for i, r in enumerate(wb["Остатки"].iter_rows(min_row=2, values_only=True), start=2):
        if not any(r[:3]):
            continue
        d, code, q = to_date(r[0]), clean(r[1]), num(r[2])
        if d is None or q is None or code not in mats:
            errors.append(f"Остатки, строка {i}: проверьте дату, код материала и количество")
            continue
        if q < 0:
            errors.append(f"Остатки, строка {i}: отрицательный остаток")
            continue
        balances[(mk(d), code)] = q
    return settings, mats, norms, receipts, balances


def load_volumes(settings, period, errors):
    """Возвращает (smeta, rows за отчётный месяц) из план-факта."""
    p = Path(clean(settings.get("Файл план-факта")))
    if not p.is_absolute():
        p = (SRC.parent / p).resolve()
    if not p.exists():
        errors.append(f"Не найден файл план-факта: {p}")
        return {}, []
    wb = load_workbook(p, data_only=True)
    smeta = {}
    for r in wb["Смета"].iter_rows(min_row=2, values_only=True):
        pos = num(r[0])
        if pos is not None:
            smeta[int(pos)] = {"name": clean(r[3]), "unit": clean(r[4])}
    rows = []
    for r in wb["Выполнение"].iter_rows(min_row=2, values_only=True):
        d, pos, q = to_date(r[0]), num(r[1]), num(r[5])
        if d is None or pos is None or q is None or mk(d) != mk(period) or int(pos) not in smeta:
            continue  # ошибки в объёмах ловит генератор форм 2 и 3
        rows.append({"pos": int(pos), "block": clean(r[3]), "floor": clean(r[4]), "qty": q})
    return smeta, rows


# ---------------------------------------------------------------- расчёт
def compute(period, mats, norms, receipts, balances, smeta, vol_rows, tolerance, errors, warns):
    pk = mk(period)
    by_pos = defaultdict(float)
    for r in vol_rows:
        by_pos[r["pos"]] += r["qty"]

    sec1 = []                      # раздел 1: работа × материал
    norm_use = defaultdict(float)  # материал -> расход по норме
    for pos in sorted(by_pos):
        if pos not in norms:
            warns.append(f"Позиция {pos} «{smeta[pos]['name']}» выполнена в месяце, но для неё не заданы нормы — "
                         f"материалы по ней не списываются")
            continue
        for code, n in norms[pos]:
            use = by_pos[pos] * n
            norm_use[code] += use
            sec1.append({"pos": pos, "work": smeta[pos]["name"], "wunit": smeta[pos]["unit"], "wqty": by_pos[pos],
                         "code": code, "norm": n, "use": use})

    by_site = defaultdict(float)   # (блок, этаж, материал) -> расход по норме
    for r in vol_rows:
        for code, n in norms.get(r["pos"], []):
            by_site[(r["block"], r["floor"], code)] += r["qty"] * n

    rec = defaultdict(float)
    for r in receipts:
        if mk(r["date"]) == pk:
            rec[r["code"]] += r["qty"]

    sec2 = []
    codes = sorted(set(norm_use) | set(rec) | {c for (k, c) in balances if k in (pk, prev_mk(pk))})
    for code in codes:
        m = mats[code]
        start = balances.get((prev_mk(pk), code))
        end = balances.get((pk, code))
        item = {"code": code, "name": m["name"], "unit": m["unit"], "price": m["price"], "start": start,
                "rec": rec[code], "norm": norm_use[code], "end": end, "fact": None, "dev": None, "pct": None, "flag": ""}
        if start is None:
            errors.append(f"{m['name']}: нет фактического остатка на конец предыдущего месяца (лист «Остатки»)")
        if end is None:
            errors.append(f"{m['name']}: нет фактического остатка на конец отчётного месяца (лист «Остатки»)")
        if start is not None and end is not None:
            fact = start + item["rec"] - end
            item["fact"] = fact
            if fact < -1e-9:
                errors.append(f"{m['name']}: остаток на конец ({fnum(end, 3)}) больше, чем было с поступлением "
                              f"({fnum(start + item['rec'], 3)}) — не внесено поступление или ошибка в остатке")
            item["dev"] = fact - item["norm"]
            if item["norm"] > 0:
                item["pct"] = item["dev"] / item["norm"]
                if item["pct"] > tolerance:
                    item["flag"] = "ПЕРЕРАСХОД"
                    warns.append(f"{m['name']}: перерасход {fnum(item['dev'], 3)} {m['unit']} ({item['pct']:.1%}) "
                                 f"≈ {fnum(item['dev'] * m['price'])} тг — нужна объяснительная или акт")
                elif item["pct"] < -tolerance:
                    item["flag"] = "экономия"
            elif fact > 1e-9:
                item["flag"] = "РАСХОД БЕЗ РАБОТ"
                warns.append(f"{m['name']}: израсходовано {fnum(fact, 3)} {m['unit']}, но работ с этим материалом "
                             f"в месяце нет — проверьте объёмы или нормы")
        sec2.append(item)
    return sec1, sec2, by_site


# ---------------------------------------------------------------- вывод
def head(ws, row, cols, widths):
    for c, (title, w) in enumerate(zip(cols, widths), start=1):
        cell = ws.cell(row, c, title)
        cell.font = Font(bold=True, size=9); cell.fill = HEAD_FILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = max(ws.column_dimensions[get_column_letter(c)].width or 0, w)
    ws.row_dimensions[row].height = 42
    return row + 1


def put(ws, r, values, fmts=None, bold=False, fill=None):
    for c, v in enumerate(values, start=1):
        cell = ws.cell(r, c, v)
        cell.border = BORDER; cell.font = Font(bold=bold, size=10)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        if fmts and c in fmts:
            cell.number_format = fmts[c]
        if fill:
            cell.fill = fill


def title_block(ws, settings, period, title, ncols):
    lines = [(title, Font(bold=True, size=13)),
             (f"за {MONTHS[period.month - 1]} {period.year} г.", Font(bold=True)),
             (f"Объект: {clean(settings.get('Объект'))}", None),
             (f"Подрядчик: {clean(settings.get('Подрядчик'))}", None)]
    for i, (t, f) in enumerate(lines, start=1):
        ws.cell(i, 1, t)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=ncols)
        ws.cell(i, 1).alignment = Alignment(horizontal="center" if i <= 2 else "left", wrap_text=True)
        if f:
            ws.cell(i, 1).font = f
    return len(lines) + 2


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def write(settings, period, mats, sec1, sec2, by_site, out):
    wb = Workbook()
    ws = wb.active; ws.title = "М-29"
    ncols = 11
    r = title_block(ws, settings, period, "ОТЧЁТ О РАСХОДЕ ОСНОВНЫХ МАТЕРИАЛОВ (М-29) — макет", ncols)

    ws.cell(r, 1, "Раздел 1. Расход материалов по нормам на выполненные работы").font = Font(bold=True); r += 1
    r = head(ws, r, ["№ поз.", "Наименование работ", "Ед. изм.", "Объём за месяц", "Код", "Материал", "Ед. изм.",
                     "Норма на ед.", "Расход по норме"], [7, 40, 8, 12, 9, 34, 8, 11, 14])
    for s in sec1:
        m = mats[s["code"]]
        put(ws, r, [s["pos"], s["work"], s["wunit"], s["wqty"], s["code"], m["name"], m["unit"], s["norm"], s["use"]],
            {4: QTY, 8: QTY, 9: QTY})
        r += 1

    r += 1
    ws.cell(r, 1, "Раздел 2. Сопоставление расхода по нормам с фактическим").font = Font(bold=True); r += 1
    r = head(ws, r, ["Код", "Материал", "Ед. изм.", "Остаток на начало", "Поступило", "Расход по нормам",
                     "Фактический расход", "Отклонение (+ перерасход / − экономия)", "Отклонение, %",
                     "Остаток на конец (факт)", "Перерасход, тг"], [7, 40, 8, 12, 12, 12, 12, 14, 10, 12, 14])
    total_over = 0.0
    for s in sec2:
        # сумма — только по перерасходу сверх допуска, чтобы колонка сходилась с итогом
        over_tg = s["dev"] * s["price"] if s["flag"] == "ПЕРЕРАСХОД" else 0.0
        total_over += over_tg
        fill = RED if s["flag"] in ("ПЕРЕРАСХОД", "РАСХОД БЕЗ РАБОТ") else GREEN if s["flag"] == "экономия" else None
        put(ws, r, [s["code"], s["name"] + (f"  [{s['flag']}]" if s["flag"] else ""), s["unit"], s["start"], s["rec"],
                    s["norm"], s["fact"], s["dev"], s["pct"], s["end"], round(over_tg, 2) if over_tg else None],
            {4: QTY, 5: QTY, 6: QTY, 7: QTY, 8: QTY, 9: "0.0%", 10: QTY, 11: MONEY}, fill=fill)
        r += 1
    put(ws, r, ["", "Итого перерасход сверх допуска, тг", "", "", "", "", "", "", "", "", round(total_over, 2)],
        {11: MONEY}, bold=True)
    r += 3
    for label, key in (("Отчёт составил", "Материально-ответственное лицо"), ("Проверил", "Инженер ПТО"),
                       ("Утверждаю", "Утверждает")):
        ws.cell(r, 2, f"{label}:").font = Font(bold=True)
        ws.cell(r, 4, "_______________")
        ws.cell(r, 6, clean(settings.get(key)))
        r += 2
    landscape(ws)

    ws2 = wb.create_sheet("По участкам")
    r = title_block(ws2, settings, period, "РАСХОД МАТЕРИАЛОВ ПО НОРМАМ — ПО БЛОКАМ И ЭТАЖАМ", 6)
    r = head(ws2, r, ["Блок", "Этаж", "Код", "Материал", "Ед. изм.", "Расход по норме"], [12, 9, 9, 44, 9, 14])
    for (blk, fl, code), q in sorted(by_site.items()):
        put(ws2, r, [blk, fl, code, mats[code]["name"], mats[code]["unit"], q], {6: QTY})
        r += 1
    landscape(ws2)
    wb.save(out / "М-29.xlsx")
    return total_over


def main():
    if not SRC.exists():
        sys.exit(f"Не найден файл: {SRC}")
    errors, warns = [], []
    settings, mats, norms, receipts, balances = load(errors, warns)
    period = to_date(settings.get("Отчётный период"))
    if not period:
        sys.exit("Настройки: не задан «Отчётный период»")
    tol = num(settings.get("Допустимый перерасход")) or 0.0
    tol = tol / 100 if tol > 1 else tol
    smeta, vol_rows = load_volumes(settings, period, errors)
    sec1, sec2, by_site = compute(period, mats, norms, receipts, balances, smeta, vol_rows, tol, errors, warns)

    out = BASE / f"М-29_{period:%Y-%m}"
    out.mkdir(exist_ok=True)
    lines = [f"Проверка {SRC.name} — {dt.datetime.now():%d.%m.%Y %H:%M}",
             f"Отчётный период: {MONTHS[period.month - 1]} {period.year}, допуск перерасхода {tol:.0%}", ""]
    if errors:
        lines += ["ОШИБКИ — отчёт НЕ сформирован, исправьте и запустите снова:"] + [f"  ✗ {e}" for e in errors]
    if warns:
        lines += ["", "Внимание:"] + [f"  ! {w}" for w in warns]
    if not errors:
        total_over = write(settings, period, mats, sec1, sec2, by_site, out)
        lines += ["", f"Материалов в отчёте: {len(sec2)}; перерасход сверх допуска ≈ {fnum(total_over)} тг",
                  f"Файл: {out.name}/М-29.xlsx"]
    text = "\n".join(lines)
    (out / "Отчёт_проверки.txt").write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()

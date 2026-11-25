"""Рисует листы реестра (данные из xlsx) в виде таблицы-электронной таблицы для кадров видео."""
import datetime as dt
import html
import json
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from pathlib import Path
V = str(Path(__file__).resolve().parent / "work") + "/"
wb = load_workbook(V + "assets/реестр_работ.xlsx")


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%d.%m.%Y")
    return str(v)


def render(sheet, max_row, max_col, hl_cols=(), hl_rows=(), dim_others=False, name="x", col_scale=7.2, only=None):
    ws = wb[sheet]
    show = list(only) if only else list(range(1, max_col + 1))
    cols = {}
    for c in show:
        w = ws.column_dimensions[get_column_letter(c)].width or 10
        cols[c] = max(60, int(w * col_scale))
    rows_html = []
    # буквы колонок
    head = '<tr><th class="corner"></th>' + "".join(
        f'<th class="colhdr{" hl" if c in hl_cols else ""}" style="width:{cols[c]}px">{get_column_letter(c)}</th>'
        for c in show) + "</tr>"
    rows_html.append(head)
    for r in range(1, max_row + 1):
        cells = []
        for c in show:
            cell = ws.cell(r, c)
            v = html.escape(fmt(cell.value))
            cls = []
            style = ""
            if r == 1 and cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb not in (None, "00000000"):
                rgb = str(cell.fill.fgColor.rgb)[-6:]
                fc = str(cell.font.color.rgb)[-6:] if cell.font and cell.font.color and cell.font.color.rgb else "000000"
                style = f"background:#{rgb};color:#{fc};font-weight:{'700' if cell.font.bold else '400'};font-style:{'italic' if cell.font.italic else 'normal'};"
                cls.append("hdr")
            elif cell.font and cell.font.bold:
                style = "font-weight:700;"
            if c in hl_cols or r in hl_rows:
                cls.append("hl")
            elif dim_others and r > 1:
                cls.append("dim")
            cells.append(f'<td class="{" ".join(cls)}" style="{style}">{v}</td>')
        rows_html.append(f'<tr><th class="rowhdr{" hl" if r in hl_rows else ""}">{r}</th>' + "".join(cells) + "</tr>")
    tabs = "".join(f'<span class="tab{" active" if s == sheet else ""}">{s}</span>' for s in wb.sheetnames)
    page = f"""<!doctype html><html><head><meta charset="utf-8"><style>
body {{ margin:0; background:#fff; font-family: "Liberation Sans", Arial, sans-serif; }}
.wrap {{ display:inline-block; border:1px solid #c8c8c8; }}
.bar {{ background:#217346; color:#fff; padding:8px 14px; font-size:15px; }}
.fx {{ border-bottom:1px solid #d4d4d4; padding:5px 10px; font-size:13px; color:#555; background:#fafafa; }}
table {{ border-collapse:collapse; table-layout:fixed; font-size:13px; }}
td, th {{ border:1px solid #d4d4d4; padding:4px 6px; vertical-align:top; overflow:hidden; word-wrap:break-word; }}
th.colhdr, th.rowhdr, th.corner {{ background:#f0f0f0; color:#555; font-weight:400; font-size:12px; text-align:center; }}
th.rowhdr {{ width:34px; }}
td.hdr {{ text-align:center; vertical-align:middle; font-size:12px; }}
td.hl {{ outline:3px solid #e8590c; outline-offset:-2px; background-color: rgba(255,230,200,.35); }}
th.hl {{ background:#ffd8a8; color:#000; }}
td.dim {{ color:#999; }}
.tabs {{ background:#f0f0f0; border-top:1px solid #c8c8c8; padding:0 8px; font-size:13px; }}
.tab {{ display:inline-block; padding:6px 16px; color:#444; }}
.tab.active {{ background:#fff; color:#217346; font-weight:700; border-bottom:3px solid #217346; }}
</style></head><body><div class="wrap" id="sheet">
<div class="bar">реестр_работ.xlsx</div><div class="fx">&nbsp;</div>
<table>{"".join(rows_html)}</table><div class="tabs">{tabs}</div></div></body></html>"""
    open(V + f"assets/sheet_{name}.html", "w", encoding="utf-8").write(page)


render("Объект", 10, 2, name="obj", col_scale=9)
render("Объект", 10, 2, hl_rows=(6, 7, 8), name="obj_reps", col_scale=9)
render("Объект", 10, 2, hl_rows=(9, 10), name="obj_num", col_scale=9)
render("Работы", 7, 8, hl_cols=(2, 3, 4, 5, 6), name="works_main", col_scale=6.6)
render("Работы", 7, 8, hl_cols=(1,), name="works_gen", col_scale=6.6)
# необязательные колонки: показываем I..N (сдвиг через отдельный рендер)
ws = wb["Работы"]
render("Работы", 7, 14, hl_cols=(7, 8, 9, 10, 11, 12, 13, 14), name="works_opt", col_scale=5.6, only=[1, 2, 7, 8, 9, 10, 11, 12, 13, 14])
render("Материалы", 11, 8, name="mat", col_scale=5.8)
render("Материалы", 11, 8, hl_cols=(1, 6), name="mat_hl", col_scale=5.8)
print("ok")

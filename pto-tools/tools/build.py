"""Собирает автономные HTML-файлы в dist/ из исходников.

    npm install          # один раз
    python3 tools/build.py
"""
import base64
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NM = ROOT / "node_modules"
DIST = ROOT / "dist"
BUILD = ROOT / "build"


def js(path):
    # </script внутри библиотек оборвал бы тег — экранируем
    return Path(path).read_text(encoding="utf-8").replace("</script", "<\\/script")


def b64(path):
    return base64.b64encode(Path(path).read_bytes()).decode()


def fill(shell, parts):
    html = Path(shell).read_text(encoding="utf-8")
    for ph, content in parts:
        assert html.count(ph) == 1, f"{ph} в {shell}"
        i = html.index(ph)
        html = html[:i] + content + html[i + len(ph):]
    return html


def main():
    DIST.mkdir(exist_ok=True)
    BUILD.mkdir(exist_ok=True)
    subprocess.run([str(NM / ".bin" / "esbuild"), str(ROOT / "aosr" / "entry.js"), "--bundle", "--minify", "--format=iife",
                    "--platform=browser", "--log-level=warning", f"--outfile={BUILD / 'dt.bundle.js'}"], check=True)

    aosr = fill(ROOT / "aosr" / "shell.html", [
        ("/*LIB_XLSX*/", js(NM / "xlsx" / "dist" / "xlsx.core.min.js")),
        ("/*LIB_DT*/", js(BUILD / "dt.bundle.js")),
        ("/*DATA*/", f'window.AOSR_TEMPLATE_B64="{b64(ROOT / "aosr" / "template.docx")}";'
                     f'window.AOSR_REGISTRY_B64="{b64(ROOT / "aosr" / "реестр_работ.xlsx")}";'),
        ("/*APP*/", js(ROOT / "aosr" / "app.js")),
    ])
    (DIST / "Генератор_АОСР.html").write_text(aosr, encoding="utf-8")

    forms = fill(ROOT / "forms_2_3" / "shell.html", [
        ("/*LIB_EXCELJS*/", js(NM / "exceljs" / "dist" / "exceljs.min.js")),
        ("/*DATA*/", f'window.KS_SAMPLE_B64="{b64(ROOT / "forms_2_3" / "план_факт.xlsx")}";'),
        ("/*APP*/", js(ROOT / "forms_2_3" / "app.js")),
    ])
    (DIST / "Формы_2_и_3.html").write_text(forms, encoding="utf-8")

    for name in ("реестр_работ.xlsx",):
        (DIST / name).write_bytes((ROOT / "aosr" / name).read_bytes())
    (DIST / "план_факт.xlsx").write_bytes((ROOT / "forms_2_3" / "план_факт.xlsx").read_bytes())

    # ПТО-пульт: готовые инструменты во вкладках; журнал — из templates/ этого репозитория, если он есть
    apps = {"aosr": aosr, "forms": forms}
    journal = ROOT.parent / "templates" / "zhurnal-proizvodstva-rabot" / "zhurnal-proizvodstva-rabot.html"
    if journal.exists():
        apps["journal"] = journal.read_text(encoding="utf-8")
    else:
        print("  ! журнал производства работ не найден — вкладки «Журнал работ» в пульте не будет")
    data = "window.PULT_APPS={" + ",".join(
        f'"{k}":"{base64.b64encode(v.encode("utf-8")).decode()}"' for k, v in apps.items()) + "};"
    pult = fill(ROOT / "pult" / "shell.html", [("/*DATA*/", data), ("/*APP*/", js(ROOT / "pult" / "app.js"))])
    (DIST / "ПТО_пульт.html").write_text(pult, encoding="utf-8")

    for f in sorted(DIST.iterdir()):
        print(f"{f.stat().st_size // 1024:>7} КБ  dist/{f.name}")


if __name__ == "__main__":
    main()

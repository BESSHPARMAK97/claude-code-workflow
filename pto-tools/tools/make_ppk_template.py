"""Шаблон акта промежуточной приёмки ответственных конструкций (aosr/template_ppk.docx)
из шаблона АОСР (aosr/template.docx): то же оформление, текст — по форме акта промежуточной приёмки
СН РК 1.03-00 (в редакции 2011 г. — приложение Д, в редакции 2022 г. — приложение Е).

    python3 tools/make_ppk_template.py
"""
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC, DST = ROOT / "aosr" / "template.docx", ROOT / "aosr" / "template_ppk.docx"

# номер абзаца в шаблоне АОСР → новый текст (None — удалить абзац)
EDIT = {
    0: "Приложение Е",
    3: "промежуточной приёмки ответственных конструкций",
    8: "представителя строительно-монтажной организации",
    11: None, 12: None, 13: None,                      # субподрядчик — в форме промежуточной приёмки его нет
    17: "представителя проектной организации", 18: None,
    20: "(фамилия, имя, отчество (при его наличии), организация, должность)",
    21: "произвела осмотр конструкций и проверку качества работ, выполненных",
    25: "1. К приёмке предъявлены следующие конструкции",
    27: "перечень и краткая характеристика конструкций",
    28: "2. Работы выполнены по проектной документации",
    30: "наименование проектной организации, номер чертежей и дата их составления",
    31: None, 32: None, 33: None, 34: None, 35: None, 36: None,   # материалы и доп. доказательства — не в этой форме
    37: "3. Отклонения от проектной документации",
    40: "4. Даты производства работ:   начало — {start_text};   окончание — {end_text}",
    41: "5. Решение комиссии",
    42: "Работы выполнены в соответствии с проектной документацией, требованиями действующих нормативных документов. Конструкции принимаются.",
    43: "На основании изложенного разрешается производство последующих работ по устройству (монтажу):",
    49: None, 50: None, 51: None,
    55: "Представитель проектной организации:",
    57: "(подпись расшифровка подписи)",
}


def set_text(p, text):
    # текст — в первый <w:t>, остальные пустые: оформление первого фрагмента сохраняется
    n = [0]

    def sub(m):
        n[0] += 1
        t = text if n[0] == 1 else ""
        return m.group(1) + t.replace("&", "&amp;").replace("<", "&lt;") + "</w:t>"
    out = re.sub(r'(<w:t(?:\s[^>]*)?>)[^<]*</w:t>', sub, p)
    return out.replace("<w:t>", '<w:t xml:space="preserve">', 1) if n[0] else out


def main():
    zin = zipfile.ZipFile(SRC)
    xml = zin.read("word/document.xml").decode("utf-8")
    paras = list(re.finditer(r'<w:p[ >].*?</w:p>', xml, flags=re.S))
    assert len(paras) == 58, f"в шаблоне АОСР {len(paras)} абзацев — проверьте EDIT"
    out, pos = [], 0
    for i, m in enumerate(paras):
        out.append(xml[pos:m.start()]); pos = m.end()
        if i in EDIT:
            if EDIT[i] is not None:
                out.append(set_text(m.group(0), EDIT[i]))
        else:
            out.append(m.group(0))
    out.append(xml[pos:])
    new = "".join(out)
    # после подписей (после их таблицы, в конце документа) — «Приложения»: исп. схема, акты скрытых работ, протоколы;
    # нет данных — блока нет
    tail = set_text(paras[41].group(0), "{#attach_text}Приложения:") + set_text(paras[43].group(0), "{attach_text}{/attach_text}")
    cut = new.rfind("<w:sectPr")
    new = new[:cut] + tail + new[cut:]
    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            z.writestr(item, new.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename))
    print("готово:", DST.relative_to(ROOT))


if __name__ == "__main__":
    main()

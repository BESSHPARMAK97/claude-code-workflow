import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

# ---- спокойная палитра ----
F="Arial"
INK="2F3B4A"; MUTED="6B7785"; TITLE="1F3A5F"
LINE="D3DAE2"; LINE_H="B8C3CF"
HDR="E9EEF4"; HDR_AUTO="E4EEE8"; NUMROW="F4F6F9"
AUTO="F3F7F4"; INPUT="FBFAF4"; EXF="FFF7E6"; ZEBRA="F8FAFC"
sl=Side(style="thin",color=LINE); sh=Side(style="thin",color=LINE_H)
BD=Border(left=sl,right=sl,top=sl,bottom=sl); BH=Border(left=sh,right=sh,top=sh,bottom=sh)
fill=lambda c:PatternFill("solid",fgColor=c)
font=lambda **k:Font(name=F,color=k.pop("color",INK),**k)
wrap=Alignment(wrap_text=True,vertical="center"); cen=Alignment(horizontal="center",vertical="center",wrap_text=True)
left_top=Alignment(wrap_text=True,vertical="top")

wb=Workbook()

# ---------- Справочник ----------
def ref_sheet():
    ws=wb.create_sheet("Справочник")
    lists={"A":("Виды работ",["Подготовительные работы","Земляные работы","Бетонные работы","Арматурные работы","Опалубочные работы",
                 "Монтажные работы","Сварочные работы","Каменные работы","Кровельные работы","Отделочные работы","Инженерные сети","Прочие работы"]),
           "C":("Способы ухода за бетоном",["Укрытие плёнкой","Полив водой","Плёнкообразующий состав","Термос (утеплённая опалубка)",
                 "Электропрогрев","Греющий провод","Тепляк","Обогрев тепловыми пушками"]),
           "E":("Способы уплотнения",["Глубинные вибраторы","Поверхностные вибраторы","Виброрейка","Самоуплотняющаяся смесь"]),
           "G":("Способы сварки",["Ручная дуговая (РД)","Механизированная в защитном газе (МП)","Механизированная самозащитной проволокой (МПС)",
                 "Автоматическая под флюсом (АФ)","Ванная (ВП)","Контактная точечная (КТ)","Газовая"])}
    for col,(h,vals) in lists.items():
        c=ws[f"{col}1"]; c.value=h; c.font=font(bold=True,size=10); c.fill=fill(HDR); c.border=BD
        ws.column_dimensions[col].width=32
        for i,v in enumerate(vals,2):
            x=ws[f"{col}{i}"]; x.value=v; x.font=font(size=10); x.border=BD
    ws["I1"]="Списки можно дополнять: добавляйте значения в конец столбца (до 30-й строки)."
    ws["I1"].font=font(size=9,italic=True,color=MUTED)
    ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="C9D3DE"
    return ws

N=300; FIRST=5; LAST=FIRST+N-1   # строки разделов
NJ=200; JLAST=FIRST+NJ-1        # строки журналов бетона
P4="'4 Производство работ'"

def setup_print(ws,last_col,last_row,land=True):
    ws.print_title_rows="3:4"; ws.print_area=f"A1:{L(last_col)}{last_row}"
    ws.page_setup.orientation="landscape" if land else "portrait"; ws.page_setup.paperSize=9
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_margins.left=ws.page_margins.right=0.4; ws.page_margins.top=ws.page_margins.bottom=0.5
    ws.oddFooter.center.text="Лист &P из &N"; ws.oddFooter.center.size=8

def table(ws,title,subtitle,heads,nrows):
    """heads: (заголовок, ширина, тип, авто?)  тип: n,t,d,s,num,f"""
    nc=len(heads)
    ws.sheet_view.showGridLines=False
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=nc)
    c=ws.cell(1,1,title); c.font=font(bold=True,size=13,color=TITLE); c.alignment=Alignment(vertical="center",wrap_text=True)
    ws.row_dimensions[1].height=30
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=nc)
    c=ws.cell(2,1,subtitle); c.font=font(size=9,italic=True,color=MUTED); ws.row_dimensions[2].height=18
    last=FIRST+nrows-1
    for j,(h,w,t,auto) in enumerate(heads,1):
        c=ws.cell(3,j,h); c.font=font(bold=True,size=9); c.alignment=cen; c.border=BH
        c.fill=fill(HDR_AUTO if auto else HDR)
        n=ws.cell(4,j,j); n.font=font(size=8,color=MUTED); n.alignment=cen; n.border=BH; n.fill=fill(NUMROW)
        ws.column_dimensions[L(j)].width=w
        for r in range(FIRST,last+1):
            x=ws.cell(r,j); x.border=BD
            x.font=font(size=10,color="3E5A48" if auto else INK)
            x.alignment=cen if t in("n","s","d","num","f") else wrap
            if t=="d": x.number_format="DD.MM.YYYY"
            if t=="num": x.number_format="#,##0.0#"
            if auto: x.fill=fill(AUTO)
    ws.row_dimensions[3].height=54; ws.row_dimensions[4].height=14
    for r in range(FIRST,last+1): ws.row_dimensions[r].height=30
    # мягкая «зебра» только для ручных столбцов
    for j,(h,w,t,auto) in enumerate(heads,1):
        if not auto:
            ws.conditional_formatting.add(f"{L(j)}{FIRST}:{L(j)}{last}",FormulaRule(formula=[f"AND(MOD(ROW(),2)=0,ROW()>{FIRST+1})"],fill=fill(ZEBRA)))
    ws.freeze_panes=ws.cell(FIRST,2)
    setup_print(ws,nc,last)
    return last

def dv_date(ws,col,last):
    dv=DataValidation(type="date",operator="between",formula1="DATE(2000,1,1)",formula2="DATE(2100,12,31)",allow_blank=True,
                      showErrorMessage=True,errorTitle="Неверная дата",error="Введите дату в формате ДД.ММ.ГГГГ")
    ws.add_data_validation(dv); dv.add(f"{col}{FIRST}:{col}{last}")
def dv_list(ws,col,last,src):
    dv=DataValidation(type="list",formula1=src,allow_blank=True,showErrorMessage=False)
    ws.add_data_validation(dv); dv.add(f"{col}{FIRST}:{col}{last}")
def example(ws,row,nc,note_col):
    for j in range(1,nc+1):
        c=ws.cell(row,j); c.fill=fill(EXF); c.font=Font(name=F,size=10,italic=True,color="7A5A1E")
    n=ws.cell(row,note_col,"← пример, замените своими данными"); n.font=font(size=9,italic=True,color="A07A2E")
    ws.column_dimensions[L(note_col)].width=30
def numcol(ws,last,keycol="B"):
    for r in range(FIRST,last+1):
        ws.cell(r,1,f'=IF({keycol}{r}="","",COUNTA(${keycol}${FIRST}:{keycol}{r}))')

# ---------- Инструкция ----------
ws=wb.active; ws.title="Инструкция"; ws.sheet_view.showGridLines=False
ws.column_dimensions["A"].width=3; ws.column_dimensions["B"].width=6; ws.column_dimensions["C"].width=110
ws["B2"]="Журнал производства работ"; ws["B2"].font=font(bold=True,size=16,color=TITLE)
ws["B3"]="Форма по Приложению В к СН РК 1.03-00-2022 «Строительное производство. Организация строительства предприятий, зданий и сооружений»"
ws["B3"].font=font(size=10,color=MUTED); ws.merge_cells("B3:C3"); ws["B3"].alignment=left_top; ws.row_dimensions[3].height=28
r=5; ws.cell(r,2,"Как заполнять").font=font(bold=True,size=11,color=TITLE); r+=1
steps=["Заполните «Титульный лист».",
"Разделы 1–5 — отдельные листы. Записи вносите подряд, начиная с первой пустой строки; № п/п проставляется сам.",
"В разделе 4 справа от основной таблицы есть дополнительные столбцы. Они не печатаются, и в них вносятся данные для журналов бетонных работ, ухода за бетоном, сварочных работ и антикоррозионной защиты.",
"Журнал для записи определяется сам по описанию работ: «бетонирование» → журналы бетона, «сварка» → журнал сварочных работ, «антикоррозионная защита», «огрунтовка», «грунтовка» → журнал антикоррозионной защиты. Если определилось неверно, выберите вручную в столбце «В спецжурнал: выбор вручную» (в том числе «Не переносить»).",
"Листы «Ж бетонных работ», «Ж ухода за бетоном», «Ж сварочных работ» и «Ж антикоррозионной защиты» заполняются полностью автоматически. На них ничего вводить не нужно, их можно печатать.",
"Записи раздела 4 можно сортировать и вставлять в любом месте: все данные журнала хранятся в той же строке, поэтому ничего не перепутается.",
"Даты — в формате ДД.ММ.ГГГГ. Выпадающие списки можно дополнить на листе «Справочник».",
"Печать: листы настроены на A4 по ширине страницы, шапка таблицы повторяется на каждом листе. Чтобы не печатать пустые строки, выделите заполненные и выберите «Печать выделенного».",
"После распечатки журнал нумеруется, прошнуровывается, скрепляется печатью и подписью руководителя подрядной организации."]
for i,s in enumerate(steps,1):
    ws.cell(r,2,f"{i}.").font=font(size=10,color=MUTED); ws.cell(r,2).alignment=Alignment(vertical="top",horizontal="right")
    c=ws.cell(r,3,s); c.font=font(size=10); c.alignment=left_top; ws.row_dimensions[r].height=30 if len(s)>110 else 17; r+=1
r+=1; ws.cell(r,2,"Обозначения").font=font(bold=True,size=11,color=TITLE); r+=1
for col,txt in [(INPUT,"Поле для заполнения (титульный лист)"),(AUTO,"Заполняется автоматически — не редактировать"),
                (EXF,"Строка-пример — замените своими данными или очистите"),("F6E1DF","Требует внимания: просрочено или нарушена хронология")]:
    s=ws.cell(r,2); s.fill=fill(col); s.border=BD; c=ws.cell(r,3,txt); c.font=font(size=10); ws.row_dimensions[r].height=18; r+=1
r+=1; ws.cell(r,2,"Важно").font=font(bold=True,size=11,color=TITLE); r+=1
for s in ["Названия граф разделов 1–5 составлены по форме журнала СН РК 1.03-00-2011 и практике её применения; перед сдачей сверьте их с официальным текстом Приложения В СН РК 1.03-00-2022.",
          "Сортируйте раздел 4 только вместе со столбцами справа (выделяйте всю строку), иначе данные журналов отделятся от своей записи."]:
    c=ws.cell(r,3,s); c.font=font(size=10); c.alignment=left_top; ws.row_dimensions[r].height=30; r+=1
ws.sheet_properties.tabColor="8FA6C1"
ws.page_setup.orientation="landscape"; ws.page_setup.paperSize=9; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1
ws.sheet_properties.pageSetUpPr.fitToPage=True; ws.print_area=f"A1:C{r}"

# ---------- Титульный лист ----------
ws=wb.create_sheet("Титульный лист"); ws.sheet_view.showGridLines=False
ws.column_dimensions["A"].width=46; ws.column_dimensions["B"].width=60
ws.merge_cells("A2:B2"); ws["A2"]="ЖУРНАЛ ПРОИЗВОДСТВА РАБОТ"; ws["A2"].font=font(bold=True,size=18,color=TITLE); ws["A2"].alignment=cen
ws.row_dimensions[2].height=34
ws["A3"]="№"; ws["A3"].alignment=Alignment(horizontal="right",vertical="center"); ws["A3"].font=font(size=13)
ws["B3"].fill=fill(INPUT); ws["B3"].border=BD; ws["B3"].font=font(size=13); ws.row_dimensions[3].height=22
ws.merge_cells("A4:B4"); ws["A4"]="(по форме Приложения В к СН РК 1.03-00-2022)"; ws["A4"].alignment=cen; ws["A4"].font=font(size=9,italic=True,color=MUTED)
fields=["Наименование объекта строительства","Адрес (местоположение) объекта","Заказчик (наименование, БИН, адрес)",
"Генеральный подрядчик (наименование, БИН, № и дата лицензии)","Субподрядные организации и виды работ",
"Генеральный проектировщик (наименование, № лицензии)","Шифр и наименование проекта (рабочей документации)",
"Положительное заключение экспертизы проекта (№, дата, организация)","Уведомление о начале строительно-монтажных работ (№, дата, орган ГАСК)",
"Технический надзор (организация / эксперт, Ф.И.О., № аттестата)","Авторский надзор (организация, Ф.И.О., № аттестата)",
"Начало работ (по договору / фактически)","Окончание работ (по договору / фактически)",
"Ответственный за производство работ и ведение журнала (должность, Ф.И.О., № приказа)",
"В журнале пронумеровано, прошнуровано и скреплено печатью листов"]
r=6
for f in fields:
    a=ws.cell(r,1,f); a.font=font(size=10,color="44505E"); a.alignment=wrap; a.border=BD
    b=ws.cell(r,2); b.fill=fill(INPUT); b.border=BD; b.alignment=wrap; b.font=font(size=10)
    ws.row_dimensions[r].height=36; r+=1
r+=2
ws.cell(r,1,"Руководитель подрядной организации").font=font(size=10)
ws.cell(r,2,"__________________ / __________________ /").font=font(size=10)
ws.cell(r+2,1,"М.П.").font=font(size=10,color=MUTED)
ws.page_setup.orientation="portrait"; ws.page_setup.paperSize=9; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1
ws.sheet_properties.pageSetUpPr.fitToPage=True; ws.print_area=f"A1:B{r+2}"
ws.page_margins.left=0.98; ws.page_margins.right=0.6; ws.sheet_properties.tabColor="8FA6C1"

TAB="B9C7D6"
# ---------- Раздел 1 ----------
ws=wb.create_sheet("1 ИТР")
H=[("№ п/п",6,"n",0),("Фамилия, имя, отчество",28,"t",0),("Должность",18,"t",0),("Организация",22,"t",0),("Участок (вид) работ",22,"t",0),
   ("Дата начала работы на объекте",14,"d",0),("Дата окончания работы на объекте",14,"d",0),("Приказ о назначении, № аттестата / удостоверения",26,"t",0)]
last=table(ws,"Раздел 1. Список инженерно-технического персонала, занятого на строительстве объекта","Вносятся все ИТР генподрядчика и субподрядчиков, допущенные к руководству работами, включая смену лиц.",H,N)
numcol(ws,last); dv_date(ws,"F",last); dv_date(ws,"G",last)
for j,v in enumerate(["Иванов Сергей Петрович","Прораб","ТОО «СтройМонтаж»","Общестроительные работы",dt.date(2026,3,2),None,"Приказ №12-П от 27.02.2026"],2): ws.cell(FIRST,j,v)
example(ws,FIRST,len(H),len(H)+1); ws.sheet_properties.tabColor=TAB

# ---------- Раздел 2 ----------
ws=wb.create_sheet("2 Спецжурналы")
H=[("№ п/п",6,"n",0),("Наименование специального журнала",40,"t",0),("Организация, ведущая журнал, Ф.И.О. ответственного",36,"t",0),
   ("Дата начала ведения",14,"d",0),("Дата передачи журнала заказчику",16,"d",0),("Подпись принявшего",18,"t",0)]
last=table(ws,"Раздел 2. Перечень специальных журналов работ","Журналы бетонных, сварочных, антикоррозионных работ и ухода за бетоном ведутся в этой книге на отдельных листах.",H,60)
numcol(ws,last); dv_date(ws,"D",last); dv_date(ws,"E",last)
for i,(n,d) in enumerate([("Журнал бетонных работ",dt.date(2026,3,12)),("Журнал ухода за бетоном",dt.date(2026,3,12)),("Журнал сварочных работ",dt.date(2026,3,16)),("Журнал антикоррозионной защиты",dt.date(2026,3,17))]):
    ws.cell(FIRST+i,2,n); ws.cell(FIRST+i,3,"ТОО «СтройМонтаж», Иванов С.П."); ws.cell(FIRST+i,4,d)
for i in range(4): example(ws,FIRST+i,len(H),len(H)+1)
ws.sheet_properties.tabColor=TAB

# ---------- Раздел 3 ----------
ws=wb.create_sheet("3 Акты")
H=[("№ п/п",6,"n",0),("№ акта",12,"t",0),("Наименование акта (вид работ, конструкция, место: оси, отметки)",60,"t",0),("Дата подписания",14,"d",0),
   ("Кто подписал (подрядчик / технадзор / автонадзор)",32,"t",0),("Примечание",20,"t",0)]
last=table(ws,"Раздел 3. Перечень актов освидетельствования скрытых работ и промежуточной приёмки ответственных конструкций",
           "Акты оформляются по формам приложений Д и Е СН РК 1.03-00-2022; здесь ведётся их перечень.",H,N)
numcol(ws,last); dv_date(ws,"D",last)
for j,v in enumerate(["АСР-1","Устройство песчано-гравийного основания под фундаментную плиту в осях 1–6 / А–Г, отм. −3,450",dt.date(2026,3,9),"Иванов С.П., Ахметов Е.К., Ли В.А."],2): ws.cell(FIRST,j,v)
example(ws,FIRST,len(H),len(H)+1); ws.sheet_properties.tabColor=TAB

# ---------- Раздел 4 ----------
ws=wb.create_sheet("4 Производство работ")
H=[("№ п/п",6,"n",0),("Дата",12,"d",0),("Смена",7,"s",0),("Наименование и краткое описание работ, конструкция, методы производства",46,"t",0),
   ("Место (оси, отметки, этаж)",20,"t",0),("Ед. изм.",8,"s",0),("Объём",10,"num",0),("Температура воздуха, °C",11,"num",0),("Погодные условия",14,"t",0),
   ("Исполнитель (организация, бригада)",22,"t",0),("Ответственный: Ф.И.О., подпись",20,"t",0)]
last=table(ws,"Раздел 4. Сведения о производстве работ",
           "Заполняется ежедневно по сменам. Справа от таблицы (в печать не входит) — данные для спецжурналов: бетонных, сварочных, антикоррозионных работ и ухода за бетоном.",H,N)
numcol(ws,last); dv_date(ws,"B",last)
dv_list(ws,"C",last,'"1,2,3"'); dv_list(ws,"F",last,'"м³,м²,м,т,шт,компл"')
PRINT_COLS=len(H)
# --- дополнительные столбцы справа (не печатаются) ---
EXTRA=[ # (ключ, заголовок, ширина, тип, группа)
 ("pick","В спецжурнал: выбор вручную (если не определилось)",17,"s","sel"),
 ("jr","Спецжурнал (определяется автоматически)",17,"f","sel"),
 ("ib","служ. № бетон.",4,"f","hid"),("iw","служ. № свар.",4,"f","hid"),("ia","служ. № АКЗ",4,"f","hid"),
 ("cls","Класс бетона, марки W и F",15,"t","bj"),("slump","Подвижность (осадка конуса, см / П)",13,"t","bj"),
 ("supplier","Поставщик смеси, № документа о качестве / накладной",24,"t","bj"),("tmix","t смеси при укладке, °C",10,"num","bj"),
 ("compact","Способ уплотнения",18,"s","bj"),("samples","Отобрано образцов: серий, маркировка",15,"t","bj"),
 ("rproj","Проектная прочность, МПа",10,"num","bj"),("r7","Прочность 7 сут, МПа",10,"num","bj"),("r28","Прочность 28 сут, МПа",10,"num","bj"),
 ("strip","Дата распалубки",12,"d","bj"),
 ("care","Способ ухода / прогрева",20,"s","uh"),("c1","Начало ухода",12,"d","uh"),("c2","Окончание ухода",12,"d","uh"),
 ("tconc","t бетона при выдерживании: мин / макс, °C",14,"t","uh"),("rend","Прочность к окончанию ухода, МПа (%)",14,"t","uh"),
 ("remarks","Замечания по уходу",20,"t","uh"),("cresp","Ответственный за уход",16,"t","uh"),
 ("steel","Марка стали (арматуры)",14,"t","sv"),("method","Способ сварки",22,"s","sv"),("cons","Марка электродов / проволоки, № сертификата",22,"t","sv"),
 ("heat","Подогрев, защита от ветра и осадков",16,"t","sv"),("welder","Ф.И.О. сварщика",18,"t","sv"),("cert","№ удостоверения, клеймо",16,"t","sv"),
 ("visual","Визуальный и измерительный контроль",16,"s","sv"),("ndt","Неразрушающий контроль: метод, № заключения, результат",22,"t","sv"),
 ("hum","Отн. влажность воздуха, %",10,"num","ak"),("tsurf","t поверхности, °C",10,"num","ak"),("prep","Степень очистки поверхности",12,"s","ak"),
 ("degr","Обезжиривание, обеспыливание",16,"t","ak"),("mat","Материал: марка, № партии, сертификат",22,"t","ak"),("layers","Слой № / всего слоёв",9,"t","ak"),
 ("apply","Способ нанесения",18,"s","ak"),("thick","Толщина: слоя / общая, мкм",11,"t","ak"),("adh","Адгезия, сплошность: метод и результат",22,"t","ak")]
GROUPS={"sel":("Перенос в спецжурналы","DDE5EE"),"bj":("Для журнала бетонных работ","DCEBE1"),"uh":("Для журнала ухода за бетоном","E3EDDC"),"sv":("Для журнала сварочных работ","E8E3EF"),"ak":("Для журнала антикоррозионной защиты","F0E6DC")}
XC={}
c0=PRINT_COLS+1
for idx,(k,h,w,t,g) in enumerate(EXTRA):
    j=c0+idx; XC[k]=L(j)
    c=ws.cell(3,j,h); c.font=font(bold=True,size=9); c.alignment=cen; c.border=BH
    if g in GROUPS: c.fill=fill(GROUPS[g][1])
    n=ws.cell(4,j,"—"); n.font=font(size=8,color=MUTED); n.alignment=cen; n.border=BH; n.fill=fill(NUMROW)
    ws.column_dimensions[L(j)].width=w
    if g=="hid": ws.column_dimensions[L(j)].hidden=True
    for r in range(FIRST,last+1):
        x=ws.cell(r,j); x.border=BD; x.font=font(size=10,color="3E5A48" if t=="f" else INK)
        x.alignment=cen if t in("s","d","num","f") else wrap
        if t=="d": x.number_format="DD.MM.YYYY"
        if t=="f": x.fill=fill(AUTO)
    if t=="d": dv_date(ws,L(j),last)
# подписи групп над столбцами (строка 2)
gstart={}
for idx,(k,h,w,t,g) in enumerate(EXTRA):
    if g in GROUPS and g not in gstart: gstart[g]=c0+idx
for g,(title,col) in GROUPS.items():
    j=gstart[g]; jl=max(c0+i for i,e in enumerate(EXTRA) if e[4]==g)
    ws.merge_cells(start_row=2,start_column=j,end_row=2,end_column=jl)
    x=ws.cell(2,j,title+(" · в печать не входит" if g=="sel" else "")); x.font=font(bold=True,size=9,color=TITLE); x.fill=fill(col); x.alignment=cen
# формулы определения спецжурнала
P,J,IB,IW,IA=XC["pick"],XC["jr"],XC["ib"],XC["iw"],XC["ia"]
for r in range(FIRST,last+1):
    ws[f"{J}{r}"]=(f'=IF(B{r}="","",IF({P}{r}="Не переносить","",IF({P}{r}<>"",{P}{r},'
                  f'IF(OR(ISNUMBER(SEARCH("антикорроз",D{r})),ISNUMBER(SEARCH("огрунтов",D{r})),ISNUMBER(SEARCH("грунтовк",D{r})),ISNUMBER(SEARCH("грунтован",D{r})),ISNUMBER(SEARCH("лакокрас",D{r}))),"Антикоррозионные работы",'
                  f'IF(ISNUMBER(SEARCH("свар",D{r})),"Сварочные работы",'
                  f'IF(OR(ISNUMBER(SEARCH("бетонир",D{r})),ISNUMBER(SEARCH("укладка бетон",D{r}))),"Бетонные работы",""))))))')
    ws[f"{IA}{r}"]=f'=IF({J}{r}="Антикоррозионные работы",COUNTIF(${J}${FIRST}:{J}{r},"Антикоррозионные работы"),"")'
    ws[f"{IB}{r}"]=f'=IF({J}{r}="Бетонные работы",COUNTIF(${J}${FIRST}:{J}{r},"Бетонные работы"),"")'
    ws[f"{IW}{r}"]=f'=IF({J}{r}="Сварочные работы",COUNTIF(${J}${FIRST}:{J}{r},"Сварочные работы"),"")'
dv_list(ws,P,last,'"Бетонные работы,Сварочные работы,Антикоррозионные работы,Не переносить"')
dv_list(ws,XC["prep"],last,'"Sa 1,Sa 2,Sa 2½,Sa 3,St 2,St 3"'); dv_list(ws,XC["apply"],last,'"Кисть,Валик,Безвоздушное распыление,Пневматическое распыление,Окунание"')
dv_list(ws,XC["compact"],last,"'Справочник'!$E$2:$E$30"); dv_list(ws,XC["care"],last,"'Справочник'!$C$2:$C$30")
dv_list(ws,XC["method"],last,"'Справочник'!$G$2:$G$30"); dv_list(ws,XC["visual"],last,'"Годно,Не годно — исправить"')
# подсветка столбцов чужого журнала: серым, если запись не относится к нему
for g,kind in (("bj","Бетонные работы"),("uh","Бетонные работы"),("sv","Сварочные работы"),("ak","Антикоррозионные работы")):
    cols=[XC[e[0]] for e in EXTRA if e[4]==g]
    ws.conditional_formatting.add(f"{cols[0]}{FIRST}:{cols[-1]}{last}",
        FormulaRule(formula=[f'AND(${J}{FIRST}<>"{kind}",$B{FIRST}<>"")'],fill=fill("F1F2F4"),font=Font(name=F,color="A0A8B2")))
ws.conditional_formatting.add(f"B{FIRST+1}:B{last}",FormulaRule(formula=[f'AND($B{FIRST+1}<>"",$B{FIRST}<>"",$B{FIRST+1}<$B{FIRST})'],fill=fill("F6E1DF"),font=Font(name=F,color="9B2C2C",bold=True)))
ws.print_area=f"A1:{L(PRINT_COLS)}{last}"
ex=[[dt.date(2026,3,11),"1","Армирование фундаментной плиты Фп-1: нижняя и верхняя сетки А500С Ø16 шаг 200","оси 1–6 / А–Г, отм. −3,250","т",7.8,5,"Ясно","ТОО «СтройМонтаж», бригада Нурланова","Иванов С.П."],
    [dt.date(2026,3,12),"1","Бетонирование фундаментной плиты Фп-1","оси 1–6 / А–Г, отм. −3,250","м³",86.4,6,"Без осадков","ТОО «СтройМонтаж», бригада Сейткалиева","Иванов С.П."],
    [dt.date(2026,3,16),"1","Сварка закладных деталей МН-1 колонн каркаса, тавровые соединения","оси 2–5 / Б, отм. +0,000","шт",24,3,"Ветер 6 м/с, без осадков","ТОО «СтройМонтаж», звено сварщиков","Иванов С.П."],
    [dt.date(2026,3,17),"1","Антикоррозионная защита закладных деталей МН-1 и сварных швов: грунтовка ГФ-021 в 2 слоя","оси 2–5 / Б, отм. +0,000","м²",6.5,8,"Без осадков","ТОО «СтройМонтаж», звено маляров","Иванов С.П."]]
exx=[{},
     dict(cls="С25/30, W6, F150",slump="ОК 14 см / П3",supplier="ТОО «БетонСервис», паспорт №418 от 12.03.2026",tmix=14,compact="Глубинные вибраторы",
          samples="3 серии: Фп1-1…Фп1-3",rproj=30,r7=21.4,r28=32.8,strip=dt.date(2026,3,19),care="Укрытие плёнкой",c1=dt.date(2026,3,12),c2=dt.date(2026,3,19),
          tconc="+8 / +17",rend="21,4 МПа (71%)",remarks="Нет",cresp="Иванов С.П."),
     dict(steel="С255; А500С",method="Ручная дуговая (РД)",cons="Э50А, УОНИ-13/55 Ø4, сертификат №1123",heat="Ветрозащитный экран",
          welder="Касымов Б.Т.",cert="Удостоверение №0457, клеймо К12",visual="Годно",ndt="УЗК 10% соединений, заключение №7 — годно"),
     dict(hum=65,tsurf=9,prep="St 3",degr="Уайт-спирит, обеспылено",mat="ГФ-021, партия 2211, сертификат №5520",layers="2 / 2",apply="Кисть",
          thick="25 / 50",adh="Решётчатый надрез — балл 1; сплошность без пропусков")]
NOTE_COL=c0+len(EXTRA)
for i,row in enumerate(ex):
    for j,v in enumerate(row,2): ws.cell(FIRST+i,j,v)
    for k,v in exx[i].items(): ws[f"{XC[k]}{FIRST+i}"]=v
    for j in range(1,NOTE_COL):
        if ws.column_dimensions[L(j)].hidden: continue
        c=ws.cell(FIRST+i,j); c.fill=fill(EXF); c.font=Font(name=F,size=10,italic=True,color="7A5A1E")
    ws.cell(FIRST+i,NOTE_COL,"← пример").font=font(size=9,italic=True,color="A07A2E")
ws.sheet_properties.tabColor=TAB

# ---------- Раздел 5 ----------
ws=wb.create_sheet("5 Замечания")
H=[("№ п/п",6,"n",0),("Дата",12,"d",0),("Замечания и указания",50,"t",0),("Кто выдал (должность, Ф.И.О., подпись)",24,"t",0),("Срок устранения",13,"d",0),
   ("Отметка об устранении (дата, Ф.И.О., подпись)",26,"t",0),("Ознакомлен (подпись)",14,"t",0),("Статус",13,"f",1)]
last=table(ws,"Раздел 5. Замечания технического и авторского надзора, контролирующих органов","Статус считается автоматически по сроку и отметке об устранении.",H,N)
numcol(ws,last); dv_date(ws,"B",last); dv_date(ws,"E",last)
for r in range(FIRST,last+1):
    ws.cell(r,8,f'=IF(C{r}="","",IF(F{r}<>"","Устранено",IF(AND(E{r}<>"",E{r}<TODAY()),"Просрочено","В работе")))')
ws.conditional_formatting.add(f"H{FIRST}:H{last}",FormulaRule(formula=[f'$H{FIRST}="Просрочено"'],fill=fill("F6E1DF"),font=Font(name=F,bold=True,color="9B2C2C"),stopIfTrue=True))
ws.conditional_formatting.add(f"H{FIRST}:H{last}",FormulaRule(formula=[f'$H{FIRST}="Устранено"'],font=Font(name=F,color="3E7A52"),stopIfTrue=True))
for j,v in enumerate([dt.date(2026,3,11),"Не представлены документы о качестве на арматурную сталь А500С Ø16. Приостановить армирование до представления сертификатов.","Технадзор Ахметов Е.К.",dt.date(2026,3,12),"12.03.2026 представлены, Иванов С.П."],2): ws.cell(FIRST,j,v)
example(ws,FIRST,len(H),len(H)+1); ws.sheet_properties.tabColor=TAB
# ↑ conditional formatting order: zebra added first; status rules after — fine

# ---------- Спецжурналы (заполняются автоматически) ----------
def pull(col,idx):
    k=f"ROW()-{FIRST-1}"
    src=f"INDEX({P4}!${col}${FIRST}:${col}${LAST},MATCH({k},{P4}!${idx}${FIRST}:${idx}${LAST},0))"
    return f'=IFERROR(IF({src}="","",{src}),"")'
def journal(name,title,idx,spec,color):
    ws=wb.create_sheet(name)
    heads=[("№ п/п",6,"n",0)]+[(h,w,t,1) for h,w,t,_ in spec]
    last=table(ws,title,"Лист заполняется автоматически из раздела 4 (записи, отнесённые к этому журналу). Данные вносите на листе «4 Производство работ» в столбцах справа.",heads,NJ)
    for r in range(FIRST,last+1):
        ws.cell(r,1,f'=IF(B{r}="","",ROW()-{FIRST-1})')
        for j,(h,w,t,src) in enumerate(spec,2):
            ws.cell(r,j,src(r) if callable(src) else pull(src,idx))
            if t=="pct": ws.cell(r,j).number_format="0%"
            if t=="int": ws.cell(r,j).number_format="0"
    ws.sheet_properties.tabColor=color
    return ws,last
X=XC; TABC="B7D0BF"
ws,last=journal("Ж бетонных работ","Журнал бетонных работ",IB,[
 ("Дата бетонирования",12,"d","B"),("Бетонируемая конструкция",30,"t","D"),("Место (оси, отметки)",18,"t","E"),("Объём уложенной смеси, м³",11,"num","G"),
 ("Температура воздуха, °C",11,"num","H"),("Класс бетона, марки по W и F",15,"t",X["cls"]),("Подвижность смеси (осадка конуса, см / П)",13,"t",X["slump"]),
 ("Поставщик смеси, № документа о качестве / накладной",24,"t",X["supplier"]),("Температура смеси при укладке, °C",11,"num",X["tmix"]),
 ("Способ уплотнения",16,"t",X["compact"]),("Отобрано образцов: серий, маркировка",15,"t",X["samples"]),("Проектная прочность, МПа",11,"num",X["rproj"]),
 ("Прочность в 7 сут, МПа",10,"num",X["r7"]),("Прочность в 28 сут, МПа",10,"num",X["r28"]),
 ("% от проектной (28 сут)",10,"pct",lambda r:f'=IF(AND(ISNUMBER(O{r}),ISNUMBER(M{r}),N(M{r})>0),O{r}/M{r},"")'),
 ("Дата распалубки",12,"d",X["strip"]),("Ответственный (Ф.И.О., подпись)",20,"t","K")],TABC)
ws.conditional_formatting.add(f"P{FIRST}:P{last}",FormulaRule(formula=[f'AND(ISNUMBER($P{FIRST}),$P{FIRST}<1)'],fill=fill("F6E1DF"),font=Font(name=F,bold=True,color="9B2C2C")))
ws,last=journal("Ж ухода за бетоном","Журнал ухода за бетоном",IB,[
 ("Дата бетонирования",12,"d","B"),("Конструкция",30,"t","D"),("Место (оси, отметки)",18,"t","E"),("Объём, м³",10,"num","G"),
 ("Температура воздуха при бетонировании, °C",13,"num","H"),("Способ ухода / прогрева",22,"t",X["care"]),("Начало ухода",12,"d",X["c1"]),("Окончание ухода",12,"d",X["c2"]),
 ("Продолжи-тельность, сут",11,"int",lambda r:f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r})),I{r}-H{r},"")'),
 ("Температура бетона при выдерживании: мин / макс, °C",15,"t",X["tconc"]),("Прочность к окончанию ухода, МПа (%)",14,"t",X["rend"]),
 ("Замечания (трещины, пересушивание и т. п.)",24,"t",X["remarks"]),("Ответственный (Ф.И.О., подпись)",20,"t",X["cresp"])],TABC)
ws.conditional_formatting.add(f"J{FIRST}:J{last}",FormulaRule(formula=[f'AND(ISNUMBER($J{FIRST}),$J{FIRST}<0)'],fill=fill("F6E1DF"),font=Font(name=F,bold=True,color="9B2C2C")))
ws,last=journal("Ж сварочных работ","Журнал сварочных работ",IW,[
 ("Дата сварки",12,"d","B"),("Конструкция, узел, соединение",30,"t","D"),("Место (оси, отметки)",18,"t","E"),("Ед. изм.",7,"t","F"),("Объём",9,"num","G"),
 ("Температура воздуха, °C",10,"num","H"),("Марка стали (арматуры) свариваемых элементов",14,"t",X["steel"]),("Способ сварки",20,"t",X["method"]),
 ("Марка электродов / проволоки, № сертификата",22,"t",X["cons"]),("Подогрев, защита от ветра и осадков",16,"t",X["heat"]),
 ("Ф.И.О. сварщика",18,"t",X["welder"]),("№ удостоверения, клеймо сварщика",16,"t",X["cert"]),("Визуальный и измерительный контроль",14,"t",X["visual"]),
 ("Неразрушающий контроль: метод, № заключения, результат",22,"t",X["ndt"]),("Ответственный (Ф.И.О., подпись)",18,"t","K")],"C9C0D9")
ws.conditional_formatting.add(f"N{FIRST}:N{last}",FormulaRule(formula=[f'LEFT($N{FIRST},3)="Не "'],fill=fill("F6E1DF"),font=Font(name=F,bold=True,color="9B2C2C")))
ws,last=journal("Ж антикоррозионной защиты","Журнал антикоррозионной защиты",IA,[
 ("Дата",12,"d","B"),("Конструкция, элемент, вид покрытия",30,"t","D"),("Место (оси, отметки)",16,"t","E"),("Ед. изм.",7,"t","F"),("Площадь / объём",9,"num","G"),
 ("Температура воздуха, °C",10,"num","H"),("Отн. влажность воздуха, %",10,"num",X["hum"]),("Температура поверхности, °C",10,"num",X["tsurf"]),
 ("Подготовка поверхности: степень очистки",12,"t",X["prep"]),("Обезжиривание, обеспыливание",16,"t",X["degr"]),
 ("Материал: марка, № партии, сертификат",22,"t",X["mat"]),("Слой № / всего слоёв",9,"t",X["layers"]),("Способ нанесения",16,"t",X["apply"]),
 ("Толщина: слоя / общая, мкм",11,"t",X["thick"]),("Адгезия, сплошность: метод и результат",22,"t",X["adh"]),("Ответственный (Ф.И.О., подпись)",18,"t","K")],"DCC7B0")
ws.conditional_formatting.add(f"H{FIRST}:H{last}",FormulaRule(formula=[f'AND(ISNUMBER($H{FIRST}),$H{FIRST}>80)'],fill=fill("F6E1DF"),font=Font(name=F,bold=True,color="9B2C2C")))
ws.cell(2,1).value+=" Красным: влажность выше 80% — проверьте допустимость по паспорту материала."

ref_sheet()
import os; out=os.path.join(os.path.dirname(os.path.abspath(__file__)),"zhurnal-proizvodstva-rabot.xlsx"); wb.save(out); print(out)

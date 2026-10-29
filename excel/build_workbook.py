"""Build the supporting-calculations workbook with live Excel formulas.

    python -m excel.build_workbook          # writes the workbook (no cached values yet)
    powershell excel/recalc.ps1             # Windows + Excel: recalculates and saves cached values

All figures in the workbook are formulas over the `Data` sheet, and the store to improve and the
categories behind its gap are chosen by formulas too, so editing a number in `Data` re-points the
workbook. The `Checks` sheet puts each Excel result next to the Python result.
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

from dm_analyst import config
from dm_analyst.load import load_raw
from dm_analyst.run_analysis import compute_all

OUT = config.EXCEL_PATH
df = load_raw()
R = compute_all()  # the Python results the Checks sheet compares against
N = len(df)
LAST = N + 1

F = 'Arial'
base = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
blue = Font(name=F, size=10, color='0000FF')
hdr_font = Font(name=F, size=10, bold=True, color='FFFFFF')
title_font = Font(name=F, size=14, bold=True)
hdr_fill = PatternFill('solid', fgColor='1F3A5F')
key_fill = PatternFill('solid', fgColor='FFF2CC')
note_font = Font(name=F, size=9, italic=True, color='595959')

wb = Workbook()

# ---------------- Data sheet ----------------
wd = wb.active
wd.title = 'Data'
raw_cols = list(df.columns)
helpers = ['Dup', 'Incl', 'NetSales', 'NetProfit', 'Qtr', 'SameBill', 'IsLast', 'BillRun', 'BillTotal',
           'BigRun', 'BillTotalExBig', 'Cat1Run', 'Cat2Run', 'PriciestRun', 'BigLine', 'DiscOver',
           'CostGtPrice', 'BelowCost', 'Lines', 'PriciestFlag']
heads = raw_cols + helpers
col = {h: L(len(raw_cols) + 1 + i) for i, h in enumerate(helpers)}
assert col['Dup'] == 'T' and col['IsLast'] == 'Z' and col['BillTotal'] == 'AB'
wd.append(heads)
for c in range(1, len(heads) + 1):
    cell = wd.cell(1, c)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal='center')

c_ = col  # short alias used in the row formulas
for i, row in enumerate(df.values.tolist()):
    r = i + 2
    out = []
    for v in row:
        if isinstance(v, pd.Timestamp):
            v = v.to_pydatetime()
        elif v != v:  # NaN
            v = None
        out.append(v)
    run = lambda name, cond: f'=IF({c_["SameBill"]}{r}=1,{c_[name]}{r-1},0)+IF(AND({c_["Incl"]}{r}=1,{cond}),1,0)'
    f = [
        f'=IF(AND(D{r}=D{r-1},B{r}=B{r-1},C{r}=C{r-1}),1,0)',                          # Dup
        f'=1-{c_["Dup"]}{r}',                                                           # Incl
        f'=Q{r}*{c_["Incl"]}{r}',                                                       # NetSales
        f'=(Q{r}-R{r})*{c_["Incl"]}{r}',                                                # NetProfit
        f'=ROUNDUP(MONTH(A{r})/3,0)',                                                   # Qtr
        f'=IF(AND(D{r}=D{r-1},B{r}=B{r-1}),1,0)',                                       # SameBill
        f'=IF(AND(D{r}=D{r+1},B{r}=B{r+1}),0,1)',                                       # IsLast
        f'=IF({c_["SameBill"]}{r}=1,{c_["BillRun"]}{r-1},0)+{c_["NetSales"]}{r}',       # BillRun
        f'=IF({c_["IsLast"]}{r}=1,{c_["BillRun"]}{r},"")',                              # BillTotal
        run('BigRun', f'M{r}>=Assumptions!$B$4'),                                       # BigRun
        f'=IF(AND({c_["IsLast"]}{r}=1,{c_["BigRun"]}{r}=0),{c_["BillRun"]}{r},"")',     # BillTotalExBig
        run('Cat1Run', f'K{r}=Assumptions!$B$7'),                                       # Cat1Run
        run('Cat2Run', f'K{r}=Assumptions!$B$8'),                                       # Cat2Run
        run('PriciestRun', f'K{r}=Assumptions!$B$9'),                                   # PriciestRun
        f'=IF(AND({c_["Incl"]}{r}=1,M{r}>=Assumptions!$B$4),1,0)',                      # BigLine
        f'=IF(AND({c_["Incl"]}{r}=1,M{r}>0,P{r}>0,P{r}>=M{r}*N{r}),1,0)',               # DiscOver
        f'=IF(AND({c_["Incl"]}{r}=1,O{r}>N{r}),1,0)',                                   # CostGtPrice
        f'=IF(AND({c_["Incl"]}{r}=1,M{r}>0,Q{r}<R{r}),1,0)',                            # BelowCost
        f'=IF({c_["IsLast"]}{r}=1,C{r},"")',                                            # Lines (LineNum of the last line)
        f'=IF({c_["IsLast"]}{r}=1,IF({c_["PriciestRun"]}{r}>0,1,0),"")',                # PriciestFlag
    ]
    wd.append(out + f)
for r in range(2, LAST + 1):
    wd[f'A{r}'].number_format = 'yyyy-mm-dd'
wd.freeze_panes = 'A2'
for c in range(1, len(heads) + 1):
    wd.column_dimensions[L(c)].width = 13
wd.column_dimensions['I'].width = 22


def rng(name_or_letter):
    letter = col.get(name_or_letter, name_or_letter)
    return f'Data!${letter}$2:${letter}${LAST}'


STORE, CAT, MQTY, SALESCH = rng('E'), rng('K'), rng('M'), rng('S')
SALES, PROFIT, QTR, INCL = rng('NetSales'), rng('NetProfit'), rng('Qtr'), rng('Incl')
BILLTOT, ISLAST, EXBIG = rng('BillTotal'), rng('IsLast'), rng('BillTotalExBig')
CAT1, CAT2, PRI = rng('Cat1Run'), rng('Cat2Run'), rng('PriciestRun')
LINES, PRIFLAG = rng('Lines'), rng('PriciestFlag')
CUST = rng('G')

# ---------------- helpers for writing ----------------
def put(ws, ref, v, font=base, fmt=None, fill=None, align=None):
    c = ws[ref]
    c.value = v
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    return c


def header(ws, row, labels, start_col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row, start_col + i, t)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def title(ws, text, sub=None):
    put(ws, 'A1', text, title_font)
    if sub:
        put(ws, 'A2', sub, note_font)


def setup(ws, widths):
    ws.sheet_view.rightToLeft = True
    for k, v in widths.items():
        ws.column_dimensions[k].width = v


NUM, NUM2, PCT = '#,##0;(#,##0);-', '#,##0.0', '0.0%'

stores = list(df.drop_duplicates('StoreCode').sort_values('StoreCode').StoreName)
cats = list(R['q1']['categories'].index)          # categories by sales, from the data
nS, nC = len(stores), len(cats)
SC = [L(2 + j) for j in range(nS)]                # store columns B..F
SFIRST, SLAST = SC[0], SC[-1]
SUBJ = 'Assumptions!$B$6'                         # the store to improve (a formula)

# ================= Assumptions (parameters + data quality) =================
wa = wb.create_sheet('Assumptions', 0)
setup(wa, {'A': 50, 'B': 18, 'C': 80})
title(wa, 'הנחות, בדיקות איכות נתונים ושימוש ב-AI',
      'כחול = קלט שאפשר לשנות; שחור = נוסחה. הסניף והקטגוריות נבחרים מהמספרים ולא נכתבים ידנית')
header(wa, 3, ['פרמטר', 'ערך', 'הערה'])
params = [
    (4, 'סף כמות לשורה חריגה (יחידות)', config.BULK_QTY, 'blue',
     'הכמות הרגילה היא 1–2; שורות עם 20–45 יחידות הן קבוצה נפרדת (בדיקת הפער ב-Python: quality.py)'),
    (5, 'אחוזון לניתוח ריכוז קבלות', 0.9, 'blue', 'משמש בשאלה 2'),
    (6, 'הסניף לשיפור (פוטנציאל הרווח הגדול ביותר)',
     f'=INDEX(Q1_Stores!{SFIRST}4:{SLAST}4,MATCH(MAX(Q1_Stores!{SFIRST}22:{SLAST}22),Q1_Stores!{SFIRST}22:{SLAST}22,0))', 'calc',
     'נבחר בנוסחה; אפשר להקליד שם סניף ידנית כדי לבדוק סניף אחר'),
    (7, 'קטגוריה חסרה בסניף - מס׳ 1',
     f'=INDEX(Q1_Stores!A27:A{26+nC},MATCH(LARGE(Q1_Stores!I27:I{26+nC},1),Q1_Stores!I27:I{26+nC},0))', 'calc',
     'הקטגוריה עם התרומה החיובית הגדולה ביותר לפער הרווח של הסניף'),
    (8, 'קטגוריה חסרה בסניף - מס׳ 2',
     f'=INDEX(Q1_Stores!A27:A{26+nC},MATCH(LARGE(Q1_Stores!I27:I{26+nC},2),Q1_Stores!I27:I{26+nC},0))', 'calc', 'כנ"ל, השנייה'),
    (9, 'הקטגוריה היקרה ביותר (ערך שורה ממוצע)',
     f'=INDEX(Q2_MeanMedian!A31:A{30+nC},MATCH(MAX(Q2_MeanMedian!B31:B{30+nC}),Q2_MeanMedian!B31:B{30+nC},0))', 'calc',
     'משמשת בשאלה 2 כמשתנה מסביר לגודל קבלה'),
]
for r, a, v, kind, note in params:
    put(wa, f'A{r}', a)
    put(wa, f'B{r}', v, blue if kind == 'blue' else base, '0%' if r == 5 else None, key_fill if kind == 'blue' else None)
    put(wa, f'C{r}', note, note_font)

header(wa, 11, ['בדיקת איכות נתונים', 'כמות', 'טיפול והשפעה'])
dq = [
    ('שורות כפולות (סניף+קבלה+שורה זהים)', f"=SUM({rng('Dup')})", 'הוסרו מכל החישובים דרך עמודת Incl. כל הכפילויות צמודות (נבדק).'),
    ('מכירות בשורות הכפולות (₪)', f"=SUMIFS({rng('Q')},{rng('Dup')},1)", 'כ-0.16% מהמכירות; לא משנה את דירוג הסניפים.'),
    ('שורות החזרה (כמות שלילית)', f"=COUNTIFS({MQTY},\"<0\",{INCL},1)", 'נשארו, כי הן מכירות נטו תקינות עסקית.'),
    ('שורות כמות חריגה (מעל הסף)', f"=SUM({rng('BigLine')})", 'נשארו (מכירות לגיטימיות); הרגישות נבדקה בשאלה 2.'),
    ('מכירה עם הנחה שווה או גדולה למחיר ברוטו', f"=SUM({rng('DiscOver')})", 'מכירות שורה אפסיות/שליליות. נספרות ב-KPI 2.'),
    ('עלות יחידה גבוהה ממחיר היחידה', f"=SUM({rng('CostGtPrice')})", 'עלות חריגה פי 1.4–2.5 מהרגיל באותו מוצר (בדיקה ב-Python). נספרות ב-KPI 2.'),
    ('CustomerCode ריק', f"=COUNTBLANK({CUST})", 'לא הוחלף; לא משפיע על חישובי מכירות.'),
    ('CustomerCode = 0 (לקוח לא מזוהה)', f"=COUNTIFS({CUST},0,{INCL},1)", 'כ-16% מהשורות; מנותח כקבוצה נפרדת, לא נמחק.'),
]
for i, (a, b, c) in enumerate(dq):
    r = 12 + i
    put(wa, f'A{r}', a); put(wa, f'B{r}', b, base, NUM); put(wa, f'C{r}', c, note_font)
put(wa, 'A22', 'הנחות מרכזיות', bold)
for i, t in enumerate([
    'קבלה = StoreCode + BillNum (מספרי קבלה אינם ייחודיים בין סניפים). סכום קבלה = סכום LineSales של שורותיה.',
    'LineSales כבר נטו אחרי הנחה; ההנחה לא מופחתת שוב. רווח גולמי = LineSales - LineCost.',
    'SalesChannel נרשם ברמת השורה; קבלה אחת יכולה לכלול שורות Store ו-Online, ולכן אין "ערוץ של קבלה".',
    'קיבוץ מוצרים לפי ItemCode. המחיר קבוע לכל קוד לאורך השנה בכל הסניפים (נבדק).',
    'כל הסכומים בש"ח לפני מע"מ, כפי שהוגדר בגיליון ההנחיות. חמש השאלות עצמאיות זו מזו.',
]):
    put(wa, f'A{23+i}', t)
put(wa, 'A29', 'שימוש ב-AI ואימות', bold)
put(wa, 'A30', 'הניתוח וכתיבת הנוסחאות בוצעו בסיוע Claude. אימות: כל תוצאה מרכזית חושבה ב-Python, ב-SQL ובנוסחאות Excel והושוותה (גיליון Checks).')

# ================= Q1: stores =================
w1 = wb.create_sheet('Q1_Stores')
setup(w1, {'A': 44, **{L(i): 15 for i in range(2, 11)}})
title(w1, 'שאלה 1 - ביצועי הסניפים: השוואה, הסניף לשיפור והסיבה',
      'מכירות/רווח נטו אחרי הסרת כפילויות; קבלות = קבלות ייחודיות; מגמה = רבעון 4 מול רבעון 1')
header(w1, 4, ['מדד'] + stores + ['סה"כ רשת'])
for j, s in enumerate(stores):
    put(w1, f'{SC[j]}4', s, hdr_font, fill=hdr_fill, align=Alignment(horizontal='center'))
G = L(2 + nS)  # totals column
rows1 = [
    (5, 'מכירות (₪)', lambda c: f'=SUMIFS({SALES},{STORE},{c}$4)', NUM, f'=SUM({SFIRST}5:{SLAST}5)'),
    (6, 'רווח גולמי (₪)', lambda c: f'=SUMIFS({PROFIT},{STORE},{c}$4)', NUM, f'=SUM({SFIRST}6:{SLAST}6)'),
    (7, 'שולי רווח', lambda c: f'={c}6/{c}5', PCT, f'={G}6/{G}5'),
    (8, 'קבלות', lambda c: f'=COUNTIFS({STORE},{c}$4,{ISLAST},1)', NUM, f'=SUM({SFIRST}8:{SLAST}8)'),
    (9, 'סכום קבלה ממוצע (₪)', lambda c: f'={c}5/{c}8', NUM2, f'={G}5/{G}8'),
    (10, 'שורות מוצר', lambda c: f'=SUMIFS({INCL},{STORE},{c}$4)', NUM, f'=SUM({SFIRST}10:{SLAST}10)'),
    (11, 'יחידות לקבלה', lambda c: f'=SUMIFS({MQTY},{STORE},{c}$4,{INCL},1)/{c}8', '0.00', f'=SUMIFS({MQTY},{INCL},1)/{G}8'),
    (12, 'מכירות Online (% מהמכירות)', lambda c: f'=SUMIFS({SALES},{STORE},{c}$4,{SALESCH},"Online")/{c}5', PCT, f'=SUMIFS({SALES},{SALESCH},"Online")/{G}5'),
    (13, 'מכירות ללקוחות לא מזוהים (%)', lambda c: f'=SUMIFS({SALES},{STORE},{c}$4,{CUST},0)/{c}5', PCT, f'=SUMIFS({SALES},{CUST},0)/{G}5'),
    (14, 'חלק ממכירות הרשת', lambda c: f'={c}5/${G}$5', PCT, f'={G}5/${G}$5'),
    (15, 'מכירות רבעון 1 (₪)', lambda c: f'=SUMIFS({SALES},{STORE},{c}$4,{QTR},1)', NUM, f'=SUM({SFIRST}15:{SLAST}15)'),
    (16, 'מכירות רבעון 4 (₪)', lambda c: f'=SUMIFS({SALES},{STORE},{c}$4,{QTR},4)', NUM, f'=SUM({SFIRST}16:{SLAST}16)'),
    (17, 'שינוי מכירות Q4 מול Q1', lambda c: f'={c}16/{c}15-1', PCT, f'={G}16/{G}15-1'),
    (18, 'רווח רבעון 1 (₪)', lambda c: f'=SUMIFS({PROFIT},{STORE},{c}$4,{QTR},1)', NUM, f'=SUM({SFIRST}18:{SLAST}18)'),
    (19, 'רווח רבעון 4 (₪)', lambda c: f'=SUMIFS({PROFIT},{STORE},{c}$4,{QTR},4)', NUM, f'=SUM({SFIRST}19:{SLAST}19)'),
    (20, 'שינוי רווח Q4 מול Q1', lambda c: f'={c}19/{c}18-1', PCT, f'={G}19/{G}18-1'),
    (21, 'שולי רווח שאר הרשת (מצטבר)', lambda c: f'=(${G}6-{c}6)/(${G}5-{c}5)', PCT, None),
    (22, 'פוטנציאל רווח: מכירות × (שולי שאר הרשת − שוליים)', lambda c: f'={c}5*({c}21-{c}7)', NUM, None),
]
for r, label, fn, fmt, tot in rows1:
    put(w1, f'A{r}', label, bold if r in (7, 9, 17, 22) else base)
    for c in SC:
        put(w1, f'{c}{r}', fn(c), base, fmt, key_fill if r == 22 else None)
    if tot:
        put(w1, f'{G}{r}', tot, base, fmt)
put(w1, 'A23', 'הסניף לשיפור נבחר לפי הפוטנציאל בשורה 22 (גיליון Assumptions, תא B6).', note_font)

put(w1, 'A25', 'למה? פירוק פער הרווח של הסניף לפי קטגוריה מול שאר הרשת', bold)
header(w1, 26, ['קטגוריה', 'שולי רווח - הסניף', 'שולי רווח - שאר הרשת', 'משקל - הסניף', 'משקל - שאר הרשת',
                'מכירות הסניף (₪)', 'מכירות שאר (₪)', 'רווח הסניף (₪)', 'תרומה לפער (₪)', 'רווח שאר (₪)'])
for i, k in enumerate(cats):
    r = 27 + i
    put(w1, f'A{r}', k)
    put(w1, f'F{r}', f'=SUMIFS({SALES},{CAT},$A{r},{STORE},{SUBJ})', fmt=NUM)
    put(w1, f'G{r}', f'=SUMIFS({SALES},{CAT},$A{r},{STORE},"<>"&{SUBJ})', fmt=NUM)
    put(w1, f'H{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{STORE},{SUBJ})', fmt=NUM)
    put(w1, f'J{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{STORE},"<>"&{SUBJ})', fmt=NUM)
    put(w1, f'B{r}', f'=IF(F{r}=0,C{r},H{r}/F{r})', fmt=PCT)
    put(w1, f'C{r}', f'=J{r}/G{r}', fmt=PCT)
    put(w1, f'D{r}', f'=F{r}/F${27+nC}', fmt=PCT)
    put(w1, f'E{r}', f'=G{r}/G${27+nC}', fmt=PCT)
    put(w1, f'I{r}', f'=F${27+nC}*(E{r}-D{r})*B{r}', fmt=NUM)
T = 27 + nC
put(w1, f'A{T}', 'סה"כ', bold)
for c in 'FGHIJ':
    put(w1, f'{c}{T}', f'=SUM({c}27:{c}{T-1})', bold, NUM)
put(w1, f'A{T+2}', 'פירוק פוטנציאל הרווח של הסניף', bold)
put(w1, f'A{T+3}', 'שולי רווח בפועל - הסניף'); put(w1, f'B{T+3}', f'=H{T}/F{T}', fmt=PCT)
put(w1, f'A{T+4}', 'שולי רווח בתמהיל שאר הרשת'); put(w1, f'B{T+4}', f'=SUMPRODUCT(B27:B{T-1},E27:E{T-1})', fmt=PCT, fill=key_fill)
put(w1, f'A{T+5}', 'השפעת תמהיל (₪) = סכום התרומות'); put(w1, f'B{T+5}', f'=F{T}*(B{T+4}-B{T+3})', bold, NUM, key_fill)
put(w1, f'A{T+6}', 'השפעת שולי רווח בתוך קטגוריה (₪)'); put(w1, f'B{T+6}', f'=F{T}*SUMPRODUCT(E27:E{T-1},C27:C{T-1}-B27:B{T-1})', fmt=NUM)
put(w1, f'A{T+7}', 'סה"כ פוטנציאל (בדיקה מול שורה 22)'); put(w1, f'B{T+7}', f'=B{T+5}+B{T+6}', bold, NUM)
put(w1, f'A{T+8}', 'סימולציה באותן מכירות; אינה תחזית ואינה מביאה בחשבון שינוי בביקוש.', note_font)
MIX_ROW, RATE_ROW, MARGIN_ROW, MIXMARGIN_ROW = T + 5, T + 6, T + 3, T + 4

# ================= Q2: mean / median =================
w2 = wb.create_sheet('Q2_MeanMedian')
setup(w2, {'A': 58, 'B': 16, 'C': 16, 'D': 16, 'E': 16, 'F': 40})
title(w2, 'שאלה 2 - ממוצע וחציון של סכום קבלה',
      'סכום קבלה = סכום שורות הקבלה (אחרי הסרת כפילויות), עמודת BillTotal בגיליון Data')
header(w2, 4, ['מדד', 'ערך', 'הערה'])
rows2 = [
    ('מספר קבלות', f'=COUNT({BILLTOT})', NUM, ''),
    ('ממוצע סכום קבלה (₪)', f'=AVERAGE({BILLTOT})', NUM2, 'עונה: כמה הכנסה מניבה קבלה? (תחזיות, תכנון הכנסות)'),
    ('חציון סכום קבלה (₪)', f'=MEDIAN({BILLTOT})', NUM2, 'עונה: מהי הקבלה הטיפוסית? (חוויית הלקוח הרגיל)'),
    ('פער ממוצע - חציון (₪)', '=B6-B7', NUM2, ''),
    ('פער באחוזים מהחציון', '=B8/B7', PCT, ''),
    ('סטיית תקן', f'=STDEV({BILLTOT})', NUM2, ''),
    ('אסימטריה (Skew)', f'=SKEW({BILLTOT})', '0.0', 'חיובית וגבוהה = זנב ימני ארוך'),
    ('אחוז הקבלות מתחת לממוצע', f'=COUNTIF({BILLTOT},"<"&B6)/B5', PCT, ''),
    ('אחוזון 90 (₪)', f'=PERCENTILE({BILLTOT},Assumptions!B5)', NUM2, ''),
    ('אחוזון 99 (₪)', f'=PERCENTILE({BILLTOT},0.99)', NUM2, ''),
    ('מקסימום (₪)', f'=MAX({BILLTOT})', NUM2, ''),
    ('חלק העשירון העליון מסך המכירות', f'=SUMIF({BILLTOT},">="&B13)/SUM({BILLTOT})', PCT, ''),
    ('קבלות עם שורה חריגה (מעל הסף)', f'=B5-COUNT({EXBIG})', NUM, ''),
    ('ממוצע ללא קבלות אלה (₪)', f'=AVERAGE({EXBIG})', NUM2, ''),
    ('חציון ללא קבלות אלה (₪)', f'=MEDIAN({EXBIG})', NUM2, ''),
]
for i, (a, b, fm, c) in enumerate(rows2):
    r = 5 + i
    put(w2, f'A{r}', a, bold if i in (1, 2) else base)
    put(w2, f'B{r}', b, base, fm, key_fill if i in (1, 2) else None)
    put(w2, f'C{r}', c, note_font)

put(w2, 'A21', 'מה מסביר את הפער? הפער שנשאר בתוך קבלות דומות (ממוצע פחות חציון)', bold)
header(w2, 22, ['השוואה', 'הפער הנותר (₪)', 'חלק הפער שהוסר'])
put(w2, 'A23', 'כל הקבלות'); put(w2, 'B23', '=B8', fmt=NUM2); put(w2, 'C23', '=1-B23/$B$23', fmt=PCT)
put(w2, 'A24', 'ללא קבלות עם שורה חריגה'); put(w2, 'B24', '=B18-B19', fmt=NUM2); put(w2, 'C24', '=1-B24/$B$23', fmt=PCT)
put(w2, 'A25', 'בתוך קבלות עם אותו מספר שורות'); put(w2, 'C25', '=1-B25/$B$23', fmt=PCT)
put(w2, 'A26', 'בתוך אותו מספר שורות ואותה נוכחות של הקטגוריה היקרה, ללא חריגות'); put(w2, 'C26', '=1-B26/$B$23', fmt=PCT)
put(w2, 'A27', '="הקטגוריה היקרה: "&Assumptions!B9', note_font)

# comparable-receipt cells: lines (2..5) x has-priciest (0/1), no bulk receipts
cell_rows = []
r0 = 100
header(w2, r0 - 1, ['תא השוואה', 'שורות', 'כולל היקרה', 'קבלות', 'ממוצע', 'חציון'])
put(w2, f'G{r0-1}', 'פער', hdr_font, fill=hdr_fill)
line_vals = sorted(df.groupby(['StoreCode', 'BillNum']).LineNum.max().unique())
line_vals = [int(x) for x in line_vals]
for n in line_vals:
    for flag in (0, 1):
        r = r0 + len(cell_rows)
        put(w2, f'A{r}', f'{n} שורות, קטגוריה יקרה={flag}')
        put(w2, f'B{r}', n, blue); put(w2, f'C{r}', flag, blue)
        put(w2, f'D{r}', f'=COUNTIFS({LINES},B{r},{PRIFLAG},C{r},{EXBIG},">-1E+99")', fmt=NUM)
        put(w2, f'E{r}', f'=IF(D{r}=0,0,AVERAGEIFS({EXBIG},{LINES},B{r},{PRIFLAG},C{r}))', fmt=NUM2)
        w2[f'F{r}'] = ArrayFormula(f'F{r}', f'=IF(D{r}=0,0,MEDIAN(IF(({LINES}=B{r})*({PRIFLAG}=C{r})*ISNUMBER({EXBIG}),{EXBIG})))')
        w2[f'F{r}'].font = base; w2[f'F{r}'].number_format = NUM2
        put(w2, f'G{r}', f'=E{r}-F{r}', fmt=NUM2)
        cell_rows.append(r)
r1, r2 = cell_rows[0], cell_rows[-1]
put(w2, 'B26', f'=SUMPRODUCT(D{r1}:D{r2},G{r1}:G{r2})/SUM(D{r1}:D{r2})', fmt=NUM2, fill=key_fill)
# same number of lines only (all receipts, bulk included)
lr0 = r2 + 3
header(w2, lr0 - 1, ['מספר שורות', 'קבלות', 'ממוצע', 'חציון', 'פער'])
lrows = []
for k, n in enumerate(line_vals):
    r = lr0 + k
    put(w2, f'A{r}', n, blue)
    put(w2, f'B{r}', f'=COUNTIFS({LINES},A{r},{BILLTOT},">-1E+99")', fmt=NUM)
    put(w2, f'C{r}', f'=AVERAGEIFS({BILLTOT},{LINES},A{r})', fmt=NUM2)
    w2[f'D{r}'] = ArrayFormula(f'D{r}', f'=MEDIAN(IF(({LINES}=A{r})*ISNUMBER({BILLTOT}),{BILLTOT}))')
    w2[f'D{r}'].font = base; w2[f'D{r}'].number_format = NUM2
    put(w2, f'E{r}', f'=C{r}-D{r}', fmt=NUM2)
    lrows.append(r)
put(w2, 'B25', f'=SUMPRODUCT(B{lrows[0]}:B{lrows[-1]},E{lrows[0]}:E{lrows[-1]})/SUM(B{lrows[0]}:B{lrows[-1]})', fmt=NUM2, fill=key_fill)

put(w2, 'A30', 'ערך שורה ממוצע לפי קטגוריה (לבחירת הקטגוריה היקרה)', bold)
for i, k in enumerate(cats):
    r = 31 + i
    put(w2, f'A{r}', k)
    put(w2, f'B{r}', f'=SUMIFS({SALES},{CAT},A{r})/COUNTIFS({CAT},A{r},{INCL},1)', fmt=NUM2)

# ================= Q3: categories, chain-wide =================
w3 = wb.create_sheet('Q3_Categories')
setup(w3, {'A': 18, **{L(i): 14 for i in range(2, 13)}})
title(w3, 'שאלה 3 - קטגוריות: רבעון 1 מול רבעון 4 (כלל הרשת)', 'מכירות נטו (₪) ורווח גולמי לפי קטגוריה')
header(w3, 4, ['קטגוריה', 'מכירות Q1', 'מכירות Q4', 'שינוי (₪)', 'שינוי %', 'משקל Q1', 'משקל Q4', 'רווח Q1', 'רווח Q4',
               'שולי רווח Q1', 'שולי רווח Q4', 'שינוי רווח (₪)'])
q3_first, q3_last, q3_tot = 5, 4 + nC, 5 + nC
for i, k in enumerate(cats):
    r = 5 + i
    put(w3, f'A{r}', k)
    put(w3, f'B{r}', f'=SUMIFS({SALES},{CAT},$A{r},{QTR},1)', fmt=NUM)
    put(w3, f'C{r}', f'=SUMIFS({SALES},{CAT},$A{r},{QTR},4)', fmt=NUM)
    put(w3, f'D{r}', f'=C{r}-B{r}', fmt=NUM)
    put(w3, f'E{r}', f'=C{r}/B{r}-1', fmt=PCT)
    put(w3, f'F{r}', f'=B{r}/B${q3_tot}', fmt=PCT)
    put(w3, f'G{r}', f'=C{r}/C${q3_tot}', fmt=PCT)
    put(w3, f'H{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{QTR},1)', fmt=NUM)
    put(w3, f'I{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{QTR},4)', fmt=NUM)
    put(w3, f'J{r}', f'=H{r}/B{r}', fmt=PCT)
    put(w3, f'K{r}', f'=I{r}/C{r}', fmt=PCT)
    put(w3, f'L{r}', f'=I{r}-H{r}', fmt=NUM)
put(w3, f'A{q3_tot}', 'סה"כ', bold)
for c in 'BCDHIL':
    put(w3, f'{c}{q3_tot}', f'=SUM({c}{q3_first}:{c}{q3_last})', bold, NUM)
put(w3, f'E{q3_tot}', f'=C{q3_tot}/B{q3_tot}-1', bold, PCT)
put(w3, f'F{q3_tot}', f'=SUM(F{q3_first}:F{q3_last})', bold, PCT); put(w3, f'G{q3_tot}', f'=SUM(G{q3_first}:G{q3_last})', bold, PCT)
put(w3, f'J{q3_tot}', f'=H{q3_tot}/B{q3_tot}', bold, PCT); put(w3, f'K{q3_tot}', f'=I{q3_tot}/C{q3_tot}', bold, PCT)
put(w3, f'A{q3_tot+2}', 'הקטגוריה שירדה הכי הרבה (₪)'); put(w3, f'B{q3_tot+2}', f'=INDEX(A{q3_first}:A{q3_last},MATCH(MIN(D{q3_first}:D{q3_last}),D{q3_first}:D{q3_last},0))', bold)
put(w3, f'A{q3_tot+3}', 'הקטגוריה שגדלה הכי הרבה (₪)'); put(w3, f'B{q3_tot+3}', f'=INDEX(A{q3_first}:A{q3_last},MATCH(MAX(D{q3_first}:D{q3_last}),D{q3_first}:D{q3_last},0))', bold)
DECL, GROW = f'Q3_Categories!$B${q3_tot+2}', f'Q3_Categories!$B${q3_tot+3}'

# ================= Q4: open investigation =================
w4 = wb.create_sheet('Q4_Investigation')
setup(w4, {'A': 52, **{L(i): 14 for i in range(2, 11)}})
title(w4, 'שאלה 4 - חקירה פתוחה: הסניף עם הפער הגדול והסבריו',
      'הנושא נבחר כאן באותו כלל כמו בשאלה 1 (Assumptions!B6), לא נשען על תוצאה של גיליון אחר')
header(w4, 4, ['מדד'] + stores + ['שאר הרשת (מצטבר)'])
for j, s in enumerate(stores):
    put(w4, f'{SC[j]}4', s, hdr_font, fill=hdr_fill, align=Alignment(horizontal='center'))
RG = L(2 + nS)  # rest-of-chain column
def attach_row(r, label_formula, run_range):
    put(w4, f'A{r}', label_formula)
    for c in SC:
        put(w4, f'{c}{r}', f'=COUNTIFS({STORE},{c}$4,{ISLAST},1,{run_range},">0")/Q1_Stores!{c}$8', fmt=PCT)
    sub = f'INDEX({SFIRST}{r}:{SLAST}{r},MATCH({SUBJ},{SFIRST}$4:{SLAST}$4,0))'
    subb = f'INDEX(Q1_Stores!{SFIRST}$8:{SLAST}$8,MATCH({SUBJ},Q1_Stores!{SFIRST}$4:{SLAST}$4,0))'
    put(w4, f'{RG}{r}', f'=(SUMPRODUCT({SFIRST}{r}:{SLAST}{r},Q1_Stores!{SFIRST}$8:{SLAST}$8)-{sub}*{subb})/(Q1_Stores!{G}$8-{subb})', fmt=PCT)
attach_row(5, '="% קבלות עם "&Assumptions!B7', CAT1)
attach_row(6, '="% קבלות עם "&Assumptions!B8', CAT2)
attach_row(7, '="% קבלות עם "&Assumptions!B9&" (הקטגוריה היקרה)"', PRI)
put(w4, 'A8', 'שורות מוצר לקבלה')
put(w4, 'A9', 'שיעור הנחה (הנחה / מחיר ברוטו)')
for c in SC:
    put(w4, f'{c}8', f'=Q1_Stores!{c}10/Q1_Stores!{c}8', fmt='0.00')
    put(w4, f'{c}9', f'=SUMIFS({rng("P")},{STORE},{c}$4,{INCL},1)/SUMPRODUCT(({STORE}={c}$4)*{INCL}*({MQTY}>0)*{MQTY}*{rng("N")})', fmt=PCT)

put(w4, 'A12', 'האם הסניף מפגר במגמה הגדולה של הרשת? משקל הקטגוריה במכירות הסניף, לפי רבעון', bold)
header(w4, 13, ['קבוצה / קטגוריה', 'Q1', 'Q2', 'Q3', 'Q4'])
tr = 14
for name, cat_ref in (('ירדה הכי הרבה', DECL), ('גדלה הכי הרבה', GROW)):
    for grp, crit in (('הסניף', SUBJ), ('שאר הרשת', f'"<>"&{SUBJ}')):
        put(w4, f'A{tr}', f'="{grp} - {name}: "&{cat_ref}')
        for q in range(1, 5):
            put(w4, f'{L(1+q)}{tr}',
                f'=SUMIFS({SALES},{STORE},{crit},{QTR},{q},{CAT},{cat_ref})/SUMIFS({SALES},{STORE},{crit},{QTR},{q})', fmt=PCT)
        tr += 1
put(w4, f'A{tr+1}', 'מסקנות (ללא מספרים, כדי שלא יתיישנו - המספרים בגיליון Checks ובשורות למעלה)', bold)
for i, t in enumerate([
    'מחיר כל פריט זהה בכל הסניפים ושיעור ההנחה דומה: הפער אינו תמחור.',
    'מספר שורות לקבלה דומה: הלקוחות לא קונים פחות, אלא קונים תמהיל אחר.',
    'לקוחות שקנו גם בסניף וגם בסניפים אחרים קונים תמהיל שונה בסניף (גיליון Checks): ההסבר הוא בסניף ולא בלקוחות.',
    'השערה שלא נבדקה: זמינות מלאי, שטח מדף או שיווק של הקטגוריות החסרות. אין בנתונים מידע על מלאי או מדפים.',
]):
    put(w4, f'A{tr+2+i}', t)

# ================= Q5: KPIs =================
w5 = wb.create_sheet('Q5_KPIs')
setup(w5, {'A': 46, 'B': 50, 'C': 16, 'D': 14, 'E': 12, 'F': 44})
title(w5, 'שאלה 5 - המלצות ומדדי KPI', 'מצב התחלתי מחושב בנוסחאות; הסניף והקטגוריות נבחרים מהנתונים')
header(w5, 4, ['פעולה / KPI', 'חישוב ה-KPI', 'מצב התחלתי', 'כיוון רצוי', 'תדירות', 'ייחוס / הערה'])
pos = lambda row: f'INDEX(Q4_Investigation!{SFIRST}{row}:{SLAST}{row},MATCH({SUBJ},Q4_Investigation!{SFIRST}$4:{SLAST}$4,0))'
put(w5, 'A5', '="1. הגדלת "&Assumptions!B7&" ו-"&Assumptions!B8&" ב"&Assumptions!B6', bold)
put(w5, 'B5', '="אחוז קבלות ב"&Assumptions!B6&" שכוללות "&Assumptions!B7'); put(w5, 'C5', '=' + pos(5), fmt=PCT, fill=key_fill)
put(w5, 'D5', 'עלייה'); put(w5, 'E5', 'חודשי'); put(w5, 'F5', f'="שאר הרשת: "&TEXT(Q4_Investigation!{RG}5,"0.0%")')
put(w5, 'B6', '="אחוז קבלות ב"&Assumptions!B6&" שכוללות "&Assumptions!B8'); put(w5, 'C6', '=' + pos(6), fmt=PCT, fill=key_fill)
put(w5, 'D6', 'עלייה'); put(w5, 'E6', 'חודשי'); put(w5, 'F6', f'="שאר הרשת: "&TEXT(Q4_Investigation!{RG}6,"0.0%")')
put(w5, 'B7', 'שולי רווח של הסניף'); put(w5, 'C7', f'=INDEX(Q1_Stores!{SFIRST}7:{SLAST}7,MATCH({SUBJ},Q1_Stores!{SFIRST}4:{SLAST}4,0))', fmt=PCT, fill=key_fill)
put(w5, 'D7', 'עלייה'); put(w5, 'E7', 'חודשי'); put(w5, 'F7', f'="שאר הרשת: "&TEXT(INDEX(Q1_Stores!{SFIRST}21:{SLAST}21,MATCH({SUBJ},Q1_Stores!{SFIRST}4:{SLAST}4,0)),"0.0%")')
put(w5, 'A9', '2. חסימת מכירה מתחת לעלות (מחיר / הנחה / נתוני עלות)', bold)
put(w5, 'B9', 'מספר שורות מכירה בחודש עם LineSales < LineCost (כמות חיובית)')
put(w5, 'C9', f'=SUM({rng("BelowCost")})/12', fmt='0.0', fill=key_fill)
put(w5, 'D9', 'ירידה'); put(w5, 'E9', 'חודשי'); put(w5, 'F9', f'="הפסד שנתי: "&TEXT(SUMIFS({PROFIT},{rng("BelowCost")},1),"#,##0")&" ₪"')
put(w5, 'A11', 'עדיפות', bold)
put(w5, 'A12', f'="פעולה 1 מקבלת עדיפות: פוטנציאל הרווח שלה "&TEXT(Q1_Stores!B{MIX_ROW},"#,##0")&" ₪ בשנה (השפעת תמהיל) מול ההפסד השנתי הקטן שבפעולה 2."')

# ================= Checks =================
wc = wb.create_sheet('Checks')
setup(wc, {'A': 56, 'B': 18, 'C': 18, 'D': 12})
title(wc, 'אימות: תוצאות Excel מול Python (pandas)', 'ערכי Python נמשכים מהחישוב ב-Python (כחול); ההפרש מחושב')
header(wc, 4, ['בדיקה', 'Excel', 'Python', 'הפרש'])
q1, q2, q3, q4, q5 = R['q1'], R['q2'], R['q3'], R['q4'], R['q5']
subj = q1['focus']
look = lambda sheet, row, key: f'INDEX({sheet}!{SFIRST}{row}:{SLAST}{row},MATCH("{key}",{sheet}!{SFIRST}4:{SLAST}4,0))'
cat_row = lambda name: q3_first + cats.index(name)
checks = [
    ('סה"כ מכירות נטו (₪)', f'=Q1_Stores!{G}5', q1['stores'].Sales.sum()),
    ('סה"כ רווח גולמי (₪)', f'=Q1_Stores!{G}6', q1['stores'].Profit.sum()),
    ('מספר קבלות', f'=Q1_Stores!{G}8', q2['bills']),
    ('הסניף לשיפור - אינדקס בין הסניפים', f'=MATCH({SUBJ},Q1_Stores!{SFIRST}4:{SLAST}4,0)', stores.index(subj) + 1),
    ('מכירות הסניף לשיפור (₪)', '=' + look('Q1_Stores', 5, subj), q1['stores'].loc[subj, 'Sales']),
    ('שולי רווח הסניף לשיפור', '=' + look('Q1_Stores', 7, subj), q1['stores'].loc[subj, 'Margin']),
    ('שינוי מכירות Q4 מול Q1 - הסניף לשיפור', '=' + look('Q1_Stores', 17, subj), q1['stores'].loc[subj, 'SalesChangeQ4vsQ1']),
    ('פוטנציאל רווח הסניף (₪)', '=' + look('Q1_Stores', 22, subj), q1['stores'].loc[subj, 'ProfitUpside']),
    ('השפעת תמהיל (₪)', f'=Q1_Stores!B{MIX_ROW}', q1['focus_extra_profit_same_sales']),
    ('השפעת שולי רווח בתוך קטגוריה (₪)', f'=Q1_Stores!B{RATE_ROW}', q1['focus_rate_effect']),
    ('ממוצע סכום קבלה', '=Q2_MeanMedian!B6', q2['mean']),
    ('חציון סכום קבלה', '=Q2_MeanMedian!B7', q2['median']),
    ('ממוצע ללא קבלות עם שורה חריגה', '=Q2_MeanMedian!B18', q2['mean_ex_bulk']),
    ('פער בתוך קבלות עם אותו מספר שורות', '=Q2_MeanMedian!B25', q2['gap_steps'].Gap.iloc[2]),
    ('פער בתוך קבלות דומות (שורות + קטגוריה יקרה, ללא חריגות)', '=Q2_MeanMedian!B26', q2['gap_steps'].Gap.iloc[3]),
    (f'מכירות Q1 של {q3["biggest_decline"]}', f'=Q3_Categories!B{cat_row(q3["biggest_decline"])}', q3['categories'].loc[q3['biggest_decline'], 'Sales_Q1']),
    (f'מכירות Q4 של {q3["biggest_growth"]}', f'=Q3_Categories!C{cat_row(q3["biggest_growth"])}', q3['categories'].loc[q3['biggest_growth'], 'Sales_Q4']),
    ('שינוי רווח הרשת Q4 מול Q1', f'=Q3_Categories!I{q3_tot}/Q3_Categories!H{q3_tot}-1', q3['total_profit_change_pct']),
    (f'% קבלות עם {q1["driver_categories"][0]} בסניף לשיפור', '=' + pos(5), q4['attach_rates'].loc[subj, q1['driver_categories'][0]]),
    (f'% קבלות עם {q1["driver_categories"][0]} בשאר הרשת', f'=Q4_Investigation!{RG}5', q4['rest_attach_rates'][q1['driver_categories'][0]]),
    ('שורות מכירה מתחת לעלות בחודש', '=Q5_KPIs!C9', q5.set_index('KPI').Baseline.iloc[-1]),
]
for i, (a, f, pv) in enumerate(checks):
    r = 5 + i
    put(wc, f'A{r}', a); put(wc, f'B{r}', f, fmt='#,##0.000'); put(wc, f'C{r}', float(pv), blue, '#,##0.000')
    put(wc, f'D{r}', f'=B{r}-C{r}', fmt='0.000')
last_check = 5 + len(checks)
put(wc, f'A{last_check+1}', 'לקוחות שקנו גם בסניף לשיפור וגם בסניף אחר (בדיקת Python בלבד)', bold)
put(wc, f'A{last_check+2}', 'מספר לקוחות משותפים'); put(wc, f'C{last_check+2}', q4['common_customers'], blue)
mix = q4['common_customer_category_mix']
cat0 = cats[0]
put(wc, f'A{last_check+3}', f'משקל {cat0} של אותם לקוחות בסניף לשיפור'); put(wc, f'C{last_check+3}', float(mix.loc[subj, cat0]), blue, PCT)
put(wc, f'A{last_check+4}', f'משקל {cat0} של אותם לקוחות בשאר הרשת'); put(wc, f'C{last_check+4}', float(mix.loc['Rest of chain', cat0]), blue, PCT)
put(wc, f'A{last_check+5}', 'מקור: pandas על הנתונים המנוקים; לא משוחזר בנוסחאות כי דורש קיבוץ לפי לקוח.', note_font)

wb.move_sheet('Data', offset=len(wb.sheetnames))
for ws in wb.worksheets:
    if ws.title != 'Data':
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and c.font.name != F:
                    c.font = base
wb.save(OUT)
print('saved', OUT)

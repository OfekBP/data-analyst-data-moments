"""Build the supporting-calculations workbook with live Excel formulas.

    python -m excel.build_workbook          # writes the workbook (no cached values yet)
    powershell excel/recalc.ps1             # Windows + Excel: recalculates and saves cached values

All figures in the workbook are formulas over the `Data` sheet. The `Checks` sheet puts the
Excel result next to the Python result for the main figures.
"""
import pandas as pd
from openpyxl import Workbook
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
           'BigRun', 'BillTotalExBig', 'DryRun', 'WetRun', 'SnackRun', 'BigLine', 'DiscOver', 'CostGtPrice']
heads = raw_cols + helpers
wd.append(heads)
for c in range(1, len(heads) + 1):
    cell = wd.cell(1, c)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal='center')

vals = df.values.tolist()
for i, row in enumerate(vals):
    r = i + 2
    out = []
    for v in row:
        if isinstance(v, pd.Timestamp):
            v = v.to_pydatetime()
        elif v != v:  # NaN
            v = None
        out.append(v)
    f = [
        f'=IF(AND(D{r}=D{r-1},B{r}=B{r-1},C{r}=C{r-1}),1,0)',            # T Dup
        f'=1-T{r}',                                                       # U Incl
        f'=Q{r}*U{r}',                                                    # V NetSales
        f'=(Q{r}-R{r})*U{r}',                                             # W NetProfit
        f'=ROUNDUP(MONTH(A{r})/3,0)',                                     # X Qtr
        f'=IF(AND(D{r}=D{r-1},B{r}=B{r-1}),1,0)',                         # Y SameBill
        f'=IF(AND(D{r}=D{r+1},B{r}=B{r+1}),0,1)',                         # Z IsLast
        f'=IF(Y{r}=1,AA{r-1},0)+V{r}',                                    # AA BillRun
        f'=IF(Z{r}=1,AA{r},"")',                                          # AB BillTotal
        f'=IF(Y{r}=1,AC{r-1},0)+AH{r}',                                   # AC BigRun
        f'=IF(AND(Z{r}=1,AC{r}=0),AA{r},"")',                             # AD BillTotalExBig
        f'=IF(Y{r}=1,AE{r-1},0)+IF(AND(U{r}=1,K{r}="Dry Food"),1,0)',     # AE DryRun
        f'=IF(Y{r}=1,AF{r-1},0)+IF(AND(U{r}=1,K{r}="Wet Food"),1,0)',     # AF WetRun
        f'=IF(Y{r}=1,AG{r-1},0)+IF(AND(U{r}=1,K{r}="Snacks"),1,0)',       # AG SnackRun
        f'=IF(AND(U{r}=1,M{r}>=Assumptions!$B$4),1,0)',                   # AH BigLine
        f'=IF(AND(U{r}=1,P{r}>0,P{r}>=M{r}*N{r}),1,0)',                   # AI DiscOver
        f'=IF(AND(U{r}=1,O{r}>N{r}),1,0)',                                # AJ CostGtPrice
    ]
    wd.append(out + f)

col = {h: L(len(raw_cols) + 1 + i) for i, h in enumerate(helpers)}
assert col['Dup'] == 'T' and col['BigRun'] == 'AC' and col['BigLine'] == 'AH' and col['DiscOver'] == 'AI'
for r in range(2, LAST + 1):
    wd[f'A{r}'].number_format = 'yyyy-mm-dd'
wd.freeze_panes = 'A2'
for c in range(1, len(heads) + 1):
    wd.column_dimensions[L(c)].width = 13
wd.column_dimensions['I'].width = 22

def rng(c):
    return f'Data!${c}$2:${c}${LAST}'

SALES, PROFIT, STORE, CAT, QTR = rng('V'), rng('W'), rng('E'), rng('K'), rng('X')
BILLTOT, ISLAST = rng('AB'), rng('Z')
EXBIG = rng('AD')
BIGRUN, DRY, WET, SNK = rng('AC'), rng('AE'), rng('AF'), rng('AG')
INCL = rng('U')

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

NUM = '#,##0;(#,##0);-'
NUM2 = '#,##0.0'
PCT = '0.0%'

stores = ['Kiryat Ono', 'Haifa', 'Ashdod', 'Holon', 'Jerusalem']
cats = ['Dry Food', 'Wet Food', 'Snacks', 'Accessories', 'Grooming', 'Litter', 'Toys']

# ---------------- Assumptions ----------------
wa = wb.create_sheet('Assumptions', 0)
setup(wa, {'A': 46, 'B': 16, 'C': 80})
title(wa, 'הנחות, בדיקות איכות נתונים ושימוש ב-AI', 'ערכים בכחול הם קלטים; כל שאר הגיליונות מחושבים בנוסחאות מגיליון Data')
header(wa, 3, ['פרמטר', 'ערך', 'הערה'])
put(wa, 'A4', 'סף כמות לשורה חריגה (יחידות)'); put(wa, 'B4', 20, blue, fill=key_fill)
put(wa, 'C4', 'בנתונים הכמות הרגילה היא 1–2; שורות עם 20–45 יחידות הן קבוצה נפרדת (77 שורות)')
put(wa, 'A5', 'אחוז עשירון עליון לניתוח ריכוז'); put(wa, 'B5', 0.9, blue, '0%', key_fill)
put(wa, 'C5', 'משמש בשאלה 2')
put(wa, 'A6', 'סניף המיקוד'); put(wa, 'B6', 'Haifa', blue, fill=key_fill)
put(wa, 'C6', 'משמש בשאלות 1 ו-4 להשוואה מול שאר הרשת')

header(wa, 8, ['בדיקת איכות נתונים', 'כמות', 'טיפול והשפעה'])
dq = [
    ('שורות כפולות (סניף+קבלה+שורה זהים)', f"=SUM({rng('T')})", 'הוסרו מכל החישובים דרך עמודת Incl. כל הכפילויות צמודות (נבדק).'),
    ('מכירות בשורות הכפולות (₪)', f"=SUMIFS({rng('Q')},{rng('T')},1)", 'כ-0.16% מהמכירות; לא משנה את דירוג הסניפים.'),
    ('שורות החזרה (כמות שלילית)', f"=COUNTIFS({rng('M')},\"<0\",{INCL},1)", 'נשארו, כי הן מכירות נטו תקינות עסקית.'),
    ('שורות כמות חריגה (מעל הסף)', f"=SUM({rng('AH')})", 'נשארו; הרגישות נבדקה בשאלה 2 (ממוצע/חציון ללא קבלות אלה).'),
    ('הנחה שווה או גדולה מהמחיר ברוטו', f"=SUM({rng('AI')})", 'מכירות שורה אפסיות/שליליות. סומנו ולא תוקנו, השפעה זניחה.'),
    ('עלות יחידה גבוהה ממחיר היחידה', f"=SUM({rng('AJ')})", 'סומנו; עלות חריגה פי 1.4–2.5 מהרגיל באותו מוצר (בדיקה ב-Python). לא תוקנו.'),
    ('CustomerCode ריק', f"=COUNTBLANK({rng('G')})", 'לא הוחלף; לא משפיע על חישובי מכירות.'),
    ('CustomerCode = 0 (לקוח לא מזוהה)', f"=COUNTIFS({rng('G')},0,{INCL},1)", 'כ-16% מהשורות; מנותח כקבוצה נפרדת, לא נמחק.'),
]
for i, (a, b, c) in enumerate(dq):
    r = 9 + i
    put(wa, f'A{r}', a); put(wa, f'B{r}', b, base, NUM); put(wa, f'C{r}', c)
put(wa, 'A19', 'הנחות מרכזיות', bold)
notes = [
    'קבלה = StoreCode + BillNum (מספרי קבלה אינם ייחודיים בין סניפים). סכום קבלה = סכום LineSales של שורותיה.',
    'LineSales כבר נטו אחרי הנחה; ההנחה לא מופחתת שוב. רווח גולמי = LineSales - LineCost.',
    'SalesChannel נרשם ברמת השורה; קבלה אחת יכולה לכלול שורות Store ו-Online, ולכן אין "ערוץ של קבלה".',
    'קיבוץ מוצרים נעשה לפי ItemCode (ל-ItemName יש 12 קודים). המחירים קבועים לכל קוד לאורך השנה בכל הסניפים.',
    'כל הסכומים בש"ח לפני מע"מ, כפי שהוגדר בגיליון ההנחיות.',
]
for i, t in enumerate(notes):
    put(wa, f'A{20+i}', t)
put(wa, 'A26', 'שימוש ב-AI ואימות', bold)
put(wa, 'A27', 'ניתוח ראשוני וכתיבת הנוסחאות בוצעו בסיוע Claude. אימות: כל תוצאה מרכזית חושבה פעמיים, פעם ב-Python (pandas) ופעם בנוסחאות Excel, והתוצאות הושוו (ראו גיליון Checks).')

# ---------------- Q1 Stores ----------------
w1 = wb.create_sheet('Q1_Stores')
setup(w1, {'A': 30, **{L(i): 15 for i in range(2, 9)}})
title(w1, 'שאלה 1 - השוואת סניפים', 'מכירות/רווח נטו אחרי הסרת כפילויות. קבלות = קבלות ייחודיות')
header(w1, 4, ['מדד'] + stores + ['סה"כ רשת'])
labels = ['מכירות (₪)', 'רווח גולמי (₪)', 'שולי רווח', 'קבלות', 'סכום קבלה ממוצע (₪)', 'שורות מוצר', 'יחידות לקבלה',
          'מכירות מ-Online (%)', 'אחוז מכירות מלקוחות לא מזוהים', 'שיעור מכירות ברשת']
for i, t in enumerate(labels):
    put(w1, f'A{5+i}', t, bold if i in (2, 4) else base)
for j, s in enumerate(stores):
    c = L(2 + j)
    put(w1, f'{c}4', s, hdr_font, fill=hdr_fill, align=Alignment(horizontal='center'))
    put(w1, f'{c}5', f'=SUMIFS({SALES},{STORE},{c}$4)', fmt=NUM)
    put(w1, f'{c}6', f'=SUMIFS({PROFIT},{STORE},{c}$4)', fmt=NUM)
    put(w1, f'{c}7', f'={c}6/{c}5', fmt=PCT)
    put(w1, f'{c}8', f'=COUNTIFS({STORE},{c}$4,{ISLAST},1)', fmt=NUM)
    put(w1, f'{c}9', f'={c}5/{c}8', fmt=NUM2)
    put(w1, f'{c}10', f'=SUMIFS({INCL},{STORE},{c}$4)', fmt=NUM)
    put(w1, f'{c}11', f'=SUMIFS({rng("M")},{STORE},{c}$4,{INCL},1)/{c}8', fmt='0.00')
    put(w1, f'{c}12', f'=SUMIFS({SALES},{STORE},{c}$4,{rng("S")},"Online")/{c}5', fmt=PCT)
    put(w1, f'{c}13', f'=SUMIFS({SALES},{STORE},{c}$4,{rng("G")},0)/{c}5', fmt=PCT)
    put(w1, f'{c}14', f'={c}5/$G$5', fmt=PCT)
put(w1, 'G4', 'סה"כ רשת', hdr_font, fill=hdr_fill, align=Alignment(horizontal='center'))
put(w1, 'G5', '=SUM(B5:F5)', fmt=NUM); put(w1, 'G6', '=SUM(B6:F6)', fmt=NUM)
put(w1, 'G7', '=G6/G5', fmt=PCT); put(w1, 'G8', '=SUM(B8:F8)', fmt=NUM); put(w1, 'G9', '=G5/G8', fmt=NUM2)
put(w1, 'G10', '=SUM(B10:F10)', fmt=NUM)
put(w1, 'G11', f'=SUMIFS({rng("M")},{INCL},1)/G8', fmt='0.00')
put(w1, 'G12', f'=SUMIFS({SALES},{rng("S")},"Online")/G5', fmt=PCT)
put(w1, 'G13', f'=SUMIFS({SALES},{rng("G")},0)/G5', fmt=PCT)
put(w1, 'G14', '=G5/$G$5', fmt=PCT)

put(w1, 'A17', 'מדוע חיפה? שולי רווח לפי קטגוריה ומשקל הקטגוריה במכירות', bold)
header(w1, 18, ['קטגוריה', 'שולי רווח חיפה', 'שולי רווח שאר הרשת', 'משקל במכירות - חיפה', 'משקל במכירות - שאר', 'מכירות חיפה (₪)', 'מכירות שאר (₪)', 'רווח חיפה (₪)'])
for i, k in enumerate(cats):
    r = 19 + i
    put(w1, f'A{r}', k)
    put(w1, f'F{r}', f'=SUMIFS({SALES},{CAT},$A{r},{STORE},Assumptions!$B$6)', fmt=NUM)
    put(w1, f'G{r}', f'=SUMIFS({SALES},{CAT},$A{r},{STORE},"<>"&Assumptions!$B$6)', fmt=NUM)
    put(w1, f'H{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{STORE},Assumptions!$B$6)', fmt=NUM)
    put(w1, f'B{r}', f'=H{r}/F{r}', fmt=PCT)
    put(w1, f'C{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{STORE},"<>"&Assumptions!$B$6)/G{r}', fmt=PCT)
    put(w1, f'D{r}', f'=F{r}/F$26', fmt=PCT)
    put(w1, f'E{r}', f'=G{r}/G$26', fmt=PCT)
put(w1, 'A26', 'סה"כ', bold)
put(w1, 'F26', '=SUM(F19:F25)', bold, NUM); put(w1, 'G26', '=SUM(G19:G25)', bold, NUM); put(w1, 'H26', '=SUM(H19:H25)', bold, NUM)

put(w1, 'A28', 'סימולציה: שולי רווח של חיפה אילו היה לה תמהיל שאר הרשת', bold)
put(w1, 'A29', 'שולי רווח בפועל - חיפה'); put(w1, 'B29', '=H26/F26', fmt=PCT)
put(w1, 'A30', 'שולי רווח בתמהיל שאר הרשת'); put(w1, 'B30', '=SUMPRODUCT(B19:B25,E19:E25)', fmt=PCT, fill=key_fill)
put(w1, 'A31', 'שולי רווח שאר הרשת'); put(w1, 'B31', '=SUMPRODUCT(C19:C25,E19:E25)', fmt=PCT)
put(w1, 'A32', 'רווח נוסף אפשרי באותן מכירות (₪)'); put(w1, 'B32', '=F26*(B30-B29)', bold, NUM, key_fill)
put(w1, 'A33', 'הערה: סימולציה עם מכירות קבועות; אינה תחזית ואינה מביאה בחשבון שינוי בביקוש.', note_font)

# ---------------- Q2 Mean/Median ----------------
w2 = wb.create_sheet('Q2_MeanMedian')
setup(w2, {'A': 52, 'B': 16, 'C': 60})
title(w2, 'שאלה 2 - ממוצע וחציון של סכום קבלה', 'סכום קבלה = סכום שורות הקבלה (אחרי הסרת כפילויות). מחושב בעמודת BillTotal בגיליון Data')
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
    ('קבלות עם שורה חריגה (מעל הסף)', f'=B5-COUNT({EXBIG})', NUM, 'בדיקת ההסבר: האם קבלות גדולות מושכות את הממוצע?'),
    ('ממוצע ללא קבלות אלה (₪)', f'=AVERAGE({EXBIG})', NUM2, ''),
    ('חציון ללא קבלות אלה (₪)', f'=MEDIAN({EXBIG})', NUM2, ''),
    ('פער ללא קבלות אלה (₪)', '=B18-B19', NUM2, ''),
    ('חלק הפער המוסבר ע"י הקבלות החריגות', '=1-B20/B8', PCT, 'השאר הוא הטיה טבעית ימינה (קבלות גדולות בגלל יותר שורות)'),
]
for i, (a, b, fm, c) in enumerate(rows2):
    r = 5 + i
    put(w2, f'A{r}', a, bold if i in (1, 2) else base)
    put(w2, f'B{r}', b, base, fm, key_fill if i in (1, 2) else None)
    put(w2, f'C{r}', c, note_font)
put(w2, 'A24', 'סכום קבלה ממוצע לפי מספר שורות בקבלה', bold)
header(w2, 25, ['שורות בקבלה', 'מספר קבלות', 'ממוצע (₪)'])
# lines in bill: BillRun-type count; compute via COUNTIFS on IsLast with line number = LineNum of last line (LineNum is sequential)
for n in range(2, 6):
    r = 24 + n
    put(w2, f'A{r}', n, blue)
    put(w2, f'B{r}', f'=COUNTIFS({ISLAST},1,{rng("C")},A{r})', fmt=NUM)
    put(w2, f'C{r}', f'=AVERAGEIFS({BILLTOT},{rng("C")},A{r},{ISLAST},1)', fmt=NUM2)
put(w2, 'A31', 'הערה: מספר שורות בקבלה = LineNum של השורה האחרונה (המספור רציף).', note_font)

# ---------------- Q3 Categories ----------------
w3 = wb.create_sheet('Q3_Categories')
setup(w3, {'A': 18, **{L(i): 14 for i in range(2, 13)}})
title(w3, 'שאלה 3 - קטגוריות: רבעון 1 מול רבעון 4', 'מכירות נטו (₪) ורווח גולמי לפי קטגוריה')
header(w3, 4, ['קטגוריה', 'מכירות Q1', 'מכירות Q4', 'שינוי (₪)', 'שינוי %', 'משקל Q1', 'משקל Q4', 'רווח Q1', 'רווח Q4', 'שולי רווח Q1', 'שולי רווח Q4', 'שינוי רווח (₪)'])
for i, k in enumerate(cats):
    r = 5 + i
    put(w3, f'A{r}', k)
    put(w3, f'B{r}', f'=SUMIFS({SALES},{CAT},$A{r},{QTR},1)', fmt=NUM)
    put(w3, f'C{r}', f'=SUMIFS({SALES},{CAT},$A{r},{QTR},4)', fmt=NUM)
    put(w3, f'D{r}', f'=C{r}-B{r}', fmt=NUM)
    put(w3, f'E{r}', f'=C{r}/B{r}-1', fmt=PCT)
    put(w3, f'F{r}', f'=B{r}/B$12', fmt=PCT)
    put(w3, f'G{r}', f'=C{r}/C$12', fmt=PCT)
    put(w3, f'H{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{QTR},1)', fmt=NUM)
    put(w3, f'I{r}', f'=SUMIFS({PROFIT},{CAT},$A{r},{QTR},4)', fmt=NUM)
    put(w3, f'J{r}', f'=H{r}/B{r}', fmt=PCT)
    put(w3, f'K{r}', f'=I{r}/C{r}', fmt=PCT)
    put(w3, f'L{r}', f'=I{r}-H{r}', fmt=NUM)
put(w3, 'A12', 'סה"כ', bold)
for c in 'BCDHIL':
    put(w3, f'{c}12', f'=SUM({c}5:{c}11)', bold, NUM)
put(w3, 'E12', '=C12/B12-1', bold, PCT); put(w3, 'F12', '=SUM(F5:F11)', bold, PCT); put(w3, 'G12', '=SUM(G5:G11)', bold, PCT)
put(w3, 'J12', '=H12/B12', bold, PCT); put(w3, 'K12', '=I12/C12', bold, PCT)
put(w3, 'A14', 'פירוק: חיפה מול שאר הרשת - משקל דרי פוד ווט פוד במכירות', bold)
header(w3, 15, ['קבוצה', 'דרי Q1', 'דרי Q2', 'דרי Q3', 'דרי Q4', 'וט Q1', 'וט Q2', 'וט Q3', 'וט Q4'])
for i, (name, op) in enumerate([('Haifa', '='), ('שאר הרשת', '<>')]):
    r = 16 + i
    put(w3, f'A{r}', name)
    crit = f'"{op}"&Assumptions!$B$6' if op == '<>' else 'Assumptions!$B$6'
    for q in range(1, 5):
        tot = f'SUMIFS({SALES},{STORE},{crit},{QTR},{q})'
        put(w3, f'{L(1+q)}{r}', f'=SUMIFS({SALES},{STORE},{crit},{QTR},{q},{CAT},"Dry Food")/{tot}', fmt=PCT)
        put(w3, f'{L(5+q)}{r}', f'=SUMIFS({SALES},{STORE},{crit},{QTR},{q},{CAT},"Wet Food")/{tot}', fmt=PCT)
put(w3, 'A19', 'מכירות רבעוניות לפי סניף', bold)
header(w3, 20, ['סניף', 'Q1', 'Q2', 'Q3', 'Q4'])
for i, s in enumerate(stores):
    r = 21 + i
    put(w3, f'A{r}', s)
    for q in range(1, 5):
        put(w3, f'{L(1+q)}{r}', f'=SUMIFS({SALES},{STORE},$A{r},{QTR},{q})', fmt=NUM)

# ---------------- Q4 Haifa ----------------
w4 = wb.create_sheet('Q4_Haifa_Mix')
setup(w4, {'A': 40, **{L(i): 14 for i in range(2, 8)}})
title(w4, 'שאלה 4 - הסבר פער הרווחיות של חיפה', 'האם הפער נובע ממחיר, ערוץ, או תמהיל? אחוז קבלות עם קטגוריה (attach rate)')
header(w4, 4, ['מדד'] + stores)
for j, s in enumerate(stores):
    put(w4, f'{L(2+j)}4', s, hdr_font, fill=hdr_fill, align=Alignment(horizontal='center'))
at = [('דרי פוד', DRY), ('וט פוד', WET), ('חטיפים', SNK)]
for i, (nm, rg) in enumerate(at):
    r = 5 + i
    put(w4, f'A{r}', f'% קבלות עם {nm}')
    for j in range(5):
        c = L(2 + j)
        put(w4, f'{c}{r}', f'=COUNTIFS({STORE},{c}$4,{ISLAST},1,{rg},">0")/COUNTIFS({STORE},{c}$4,{ISLAST},1)', fmt=PCT)
put(w4, 'A8', 'שורות מוצר לקבלה')
for j in range(5):
    c = L(2 + j)
    put(w4, f'{c}8', f'=SUMIFS({INCL},{STORE},{c}$4)/COUNTIFS({STORE},{c}$4,{ISLAST},1)', fmt='0.00')
put(w4, 'A9', 'מחיר ממוצע ליחידת דרי פוד (₪)')
for j in range(5):
    c = L(2 + j)
    put(w4, f'{c}9', f'=SUMIFS({SALES},{STORE},{c}$4,{CAT},"Dry Food")/SUMIFS({rng("M")},{STORE},{c}$4,{CAT},"Dry Food",{INCL},1)', fmt=NUM2)
put(w4, 'A10', 'שיעור הנחה (הנחה/מחיר ברוטו)')
for j in range(5):
    c = L(2 + j)
    put(w4, f'{c}10', f'=SUMIFS({rng("P")},{STORE},{c}$4,{INCL},1)/SUMPRODUCT(({STORE}={c}$4)*{INCL}*({rng("M")}>0)*{rng("M")}*{rng("N")})', fmt=PCT)
put(w4, 'A12', 'מסקנות', bold)
for i, t in enumerate([
    'מחיר כל פריט זהה בכל הסניפים, ושיעור ההנחה דומה (1.5%–1.8%): הפער אינו תמחור. הבדלים קלים במחיר ממוצע ליחידה נובעים מתמהיל הפריטים.',
    'מספר שורות בקבלה זהה: הלקוח בחיפה לא קונה פחות, אלא קונה דרי פוד במקום וט פוד וחטיפים.',
    'בדיקת הלקוחות המשותפים (ראו Checks): אותם לקוחות קונים 79% דרי פוד בחיפה ו-50% בשאר - לכן זה מאפיין של הסניף ולא של הלקוחות.',
    'השערה שלא נבדקה: זמינות מלאי, שטח מדף או שיווק של וט פוד וחטיפים בחיפה. אין בנתונים מידע על מלאי או מדפים.',
]):
    put(w4, f'A{13+i}', t)

# ---------------- Q5 KPI ----------------
w5 = wb.create_sheet('Q5_KPIs')
setup(w5, {'A': 34, 'B': 44, 'C': 16, 'D': 18, 'E': 16, 'F': 40})
title(w5, 'שאלה 5 - המלצות ומדדי KPI', 'מצב התחלתי מחושב מגיליונות הניתוח')
header(w5, 4, ['המלצה', 'חישוב ה-KPI', 'מצב התחלתי', 'כיוון רצוי', 'תדירות', 'הערות / נימוק'])
put(w5, 'A5', '1. הגדלת נוכחות וט פוד וחטיפים בחיפה', bold)
put(w5, 'B5', 'אחוז קבלות בחיפה שכוללות וט פוד'); put(w5, 'C5', "=Q4_Haifa_Mix!C6", fmt=PCT, fill=key_fill)
put(w5, 'D5', 'עלייה'); put(w5, 'E5', 'חודשי'); put(w5, 'F5', 'בשאר הרשת כ-66.5%: ייחוס טבעי ליעד')
put(w5, 'B6', 'אחוז קבלות בחיפה שכוללות חטיפים'); put(w5, 'C6', "=Q4_Haifa_Mix!C7", fmt=PCT, fill=key_fill)
put(w5, 'D6', 'עלייה'); put(w5, 'E6', 'חודשי'); put(w5, 'F6', 'בשאר הרשת כ-46.5%')
put(w5, 'B7', 'שולי רווח בחיפה'); put(w5, 'C7', "=Q1_Stores!B29", fmt=PCT, fill=key_fill)
put(w5, 'D7', 'עלייה'); put(w5, 'E7', 'חודשי'); put(w5, 'F7', 'פוטנציאל רווח נוסף: ראו Q1_Stores!B32')
put(w5, 'A9', '2. בקרת שורות חריגות', bold)
put(w5, 'B9', 'מספר שורות חריגות בחודש (כמות גדולה, הנחה ≥ מחיר ברוטו, עלות>מחיר)')
put(w5, 'C9', '=(Assumptions!B12+Assumptions!B13+Assumptions!B14)/12', fmt='0.0', fill=key_fill)
put(w5, 'D9', 'ירידה'); put(w5, 'E9', 'חודשי'); put(w5, 'F9', 'ממוצע חודשי בשנת 2025')
put(w5, 'A11', 'עדיפות', bold)
put(w5, 'A12', 'פעולה 1 מקבלת עדיפות: היא מטפלת בפער הרווחיות הגדול ביותר (Q1_Stores!B32). פעולה 2 היא בקרה תפעולית עם השפעה כספית קטנה.')

# ---------------- Checks ----------------
wc = wb.create_sheet('Checks')
setup(wc, {'A': 52, 'B': 18, 'C': 18, 'D': 12})
title(wc, 'אימות: תוצאות Excel מול Python (pandas)', 'ערכי Python הם קלט קשיח (כחול) שחושב על אותם נתונים; ההפרש מחושב')
header(wc, 4, ['בדיקה', 'Excel', 'Python', 'הפרש'])
checks = [
    ('סה"כ מכירות נטו (₪)', '=Q1_Stores!G5', R['q1']['stores'].Sales.sum()),
    ('סה"כ רווח גולמי (₪)', '=Q1_Stores!G6', R['q1']['stores'].Profit.sum()),
    ('מספר קבלות', '=Q1_Stores!G8', R['q2']['bills']),
    ('ממוצע סכום קבלה', '=Q2_MeanMedian!B6', R['q2']['mean']),
    ('חציון סכום קבלה', '=Q2_MeanMedian!B7', R['q2']['median']),
    ('ממוצע ללא קבלות עם שורה חריגה', '=Q2_MeanMedian!B18', R['q2']['mean_ex_bulk']),
    ('מכירות חיפה (₪)', '=Q1_Stores!C5', R['q1']['stores'].loc['Haifa', 'Sales']),
    ('שולי רווח חיפה', '=Q1_Stores!C7', R['q1']['stores'].loc['Haifa', 'Margin']),
    ('שולי רווח חיפה בתמהיל שאר הרשת', '=Q1_Stores!B30', R['q1']['focus_margin_at_rest_mix']),
    ('מכירות דרי פוד Q1', '=Q3_Categories!B5', R['q3']['categories'].loc['Dry Food', 'Sales_Q1']),
    ('מכירות וט פוד Q4', '=Q3_Categories!C6', R['q3']['categories'].loc['Wet Food', 'Sales_Q4']),
    ('% קבלות חיפה עם וט פוד', '=Q4_Haifa_Mix!C6', R['q4']['attach_rates'].loc['Haifa', 'Wet']),
]
for i, (a, f, pv) in enumerate(checks):
    r = 5 + i
    put(wc, f'A{r}', a); put(wc, f'B{r}', f, fmt='#,##0.000'); put(wc, f'C{r}', pv, blue, '#,##0.000')
    put(wc, f'D{r}', f'=B{r}-C{r}', fmt='0.000')
put(wc, 'A19', 'לקוחות שקנו גם בחיפה וגם בסניף אחר (בדיקת Python בלבד)', bold)
put(wc, 'A20', 'מספר לקוחות משותפים'); put(wc, 'C20', R['q4']['common_customers'], blue)
put(wc, 'A21', 'משקל דרי פוד של אותם לקוחות בחיפה'); put(wc, 'C21', R['q4']['common_customer_category_mix'].loc['Haifa', 'Dry Food'], blue, PCT)
put(wc, 'A22', 'משקל דרי פוד של אותם לקוחות בשאר הרשת'); put(wc, 'C22', R['q4']['common_customer_category_mix'].loc['Rest of chain', 'Dry Food'], blue, PCT)
put(wc, 'A23', 'מקור: pandas על הנתונים המנוקים; לא משוחזר בנוסחאות כי דורש קיבוץ לפי לקוח.', note_font)

wb.move_sheet('Data', offset=len(wb.sheetnames))
for ws in wb.worksheets:
    if ws.title != 'Data':
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and c.font.name != F:
                    c.font = base
wb.save(OUT)
print('saved', OUT)

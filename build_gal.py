import openpyxl, datetime as dt, json
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.comments import Comment

SRC = '/home/user/workspace/uploaded_attachments/36403e9f652f4ea2bb9d74a576ea5b2e/By-Day-pivot-38.xlsx'
OUT = '/home/user/workspace/gal/Longevity-GAL-September-2026.xlsx'

# ---------- BI (By-Day pivot 38) ----------
wb0 = openpyxl.load_workbook(SRC, data_only=True)
ws0 = wb0['pivot']
HDR = [c.value for c in ws0[1]]           # 21 board columns, untouched
bi = {}
for r in ws0.iter_rows(min_row=2, values_only=True):
    if isinstance(r[0], dt.datetime):
        bi[r[0].date()] = list(r)
bi_total = [c.value for c in ws0[22]]

# ---------- Meta daily spend & results (Ads API pull, account 1459085242361281, 1–28.9.2026) ----------
# campaign -> {date: (spend, atc, checkout, purchase, revenue)}
LEADS = "LGV_EN_FB_Demographics_FB Leads_2026-08-03_#117186"
LLA   = "LGV_EN_FB_Demographics_Longevity Life Academy_2026-05-04_#116622"
EC1   = "LGV_EN_PPC_ecomm-01_2026-09-02_#118148"
EC2   = "LGV_EN_PPC_ecomm-02_2026-09-02_#118149"
leads_spend = {1:1776.6,2:1024.14,3:557.67,4:929.38,5:225.37,6:251.03,7:708.42,8:1483.97,9:1035.48,10:1079.07,11:1000.92,12:358.56,13:368.4,14:546.24,15:450.02,16:484.04,17:442.91,18:442.35,19:378.27,20:340.46,21:269.38,22:357.15,23:435.6,24:476.13,25:380.37,26:174.5,27:183.58,28:81.09}
lla_spend = {1:594.95,2:597.62,3:437.62,4:274.61,5:102.24,6:113.58,7:103.69,8:97.47,9:126.91,10:135.09,11:68.79,12:50.85,13:149.54,14:138.06,15:108.45,16:147.89,17:73.39,18:117.53,19:53.29,20:106.34,21:70.79,22:82.46,23:90.88,24:132.22,25:161.24,26:94.93,27:156.13,28:81.51}
lla_atc = {1:13,2:10,3:1,4:1,5:2,8:1,9:4,10:4,11:3,13:7,14:1,15:5,16:4,17:2,20:1,21:3,22:3}
# ecomm-01: day -> (spend, atc, ic, purch, rev)
ec1 = {3:(293.62,0,0,0,0),4:(345.27,0,0,0,0),5:(374.91,0,0,0,0),6:(466.09,0,0,0,0),7:(388.04,1,0,0,0),8:(232.46,0,0,0,0),
       9:(393.75,1,1,0,0),10:(369.34,1,1,1,179),11:(250.66,0,0,0,0),12:(274.27,1,0,0,0),13:(471.76,1,1,1,179),14:(393.67,0,0,0,0),
       15:(511.01,2,1,0,0),16:(851.88,3,1,1,179),17:(850.68,0,0,0,0),18:(805.53,2,1,0,0),19:(628.5,0,0,0,0),20:(977.73,13,7,0,0),
       21:(692.24,9,1,0,0),22:(682.9,2,0,0,0),23:(526.57,2,0,0,0),24:(820.19,15,11,0,0),25:(443.3,10,0,0,0),26:(267.92,4,3,0,0),
       27:(477.29,3,0,0,0),28:(224.41,3,2,0,0)}
ec2 = {28:(293.47,0,0,0,0)}

def ec(d):
    a = ec1.get(d,(0,0,0,0,0)); b = ec2.get(d,(0,0,0,0,0))
    return tuple(x+y for x,y in zip(a,b))

# reconciliation vs uploaded Meta CSV (7–28.9)
rec = {
 'Leads #117186': (sum(v for k,v in leads_spend.items() if k>=7), 11476.91),
 'LLA/ATC #116622': (sum(v for k,v in lla_spend.items() if k>=7), 2347.24),
 'ecomm-01 #118148': (sum(v[0] for k,v in ec1.items() if k>=7), 11533.53),
 'ecomm-02 #118149': (sum(v[0] for k,v in ec2.items() if k>=7), 292.57),
}
for _ws_ in []: pass
print('RECON', json.dumps(rec, indent=1))

# hidden BI "contacts" denominator used by %Cont/Calls and %Regs/Cont: = Q * Calls
hidden = {}
for d,row in bi.items():
    q, calls = row[16], row[14]
    hidden[d] = round(q*calls) if (q and calls) else 0
hidden[dt.date(2026,9,28)] = 1  # BI grand total implies 469; 28.9 row has 1 contact from leads, no calls yet
print('hidden sum', sum(hidden.values()), 'BI grand-total implied', round(bi_total[3]/bi_total[13]))

# ---------- styles ----------
NAVY='0B1F3A'; INK='1F2937'; MID='4B5563'; LINE='D6DBE3'; SOFT='F3F5F9'; SOFT2='E8EDF5'
TEAL='0E7C86'; TEAL_L='E3F4F5'; AMBER='B45309'; AMBER_L='FDF3E1'; PLUM='6D28D9'; PLUM_L='EFE9FC'; GREEN='0F6E3D'; GREEN_L='E6F4EC'
thin = Side(style='thin', color=LINE); med = Side(style='medium', color=NAVY); hair=Side(style='hair', color=LINE)
F = lambda **k: Font(name='Calibri', **{**k, 'size': k.get('size',10)+8})
fill = lambda c: PatternFill('solid', start_color=c, end_color=c)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
right = Alignment(horizontal='right', vertical='center')
left = Alignment(horizontal='left', vertical='center')

wb = openpyxl.Workbook()
ws = wb.active; ws.title = 'GAL Sep-26'
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = NAVY

# ---------- column map ----------
# A..U = 21 original BI/board columns (unchanged names, order, formats)
# V spacer
# Added block 1: Meta spend (lead-gen)
ADD = [
 ('Spend Leads camp. (USD)', 'spend_leads', '#,##0', TEAL),
 ('Spend LLA/ATC camp. (USD)', 'spend_lla', '#,##0', TEAL),
 ('Lead-gen Spend (USD)', 'spend_lg', '#,##0', TEAL),
 ('CPL (USD)', 'cpl', '#,##0.0', TEAL),
 ('Cost / Acq (USD)', 'cpa', '#,##0', TEAL),
 ('Cost / Cont (USD)', 'cpc', '#,##0.0', TEAL),
 (None,None,None,None),  # spacer
 ('E-comm Spend (USD)', 'spend_ec', '#,##0', PLUM),
 ('E-comm Adds to Cart', 'ec_atc', '#,##0', PLUM),
 ('E-comm Checkouts', 'ec_ic', '#,##0', PLUM),
 ('E-comm Purchases', 'ec_p', '#,##0', PLUM),
 ('E-comm Revenue (USD)', 'ec_rev', '#,##0', PLUM),
 ('Cost / ATC (USD)', 'ec_cpatc', '#,##0.0', PLUM),
 ('Cost / Purchase (USD)', 'ec_cpp', '#,##0', PLUM),
 ('ROAS', 'ec_roas', '0.00', PLUM),
 (None,None,None,None),
 ('TOTAL Meta Spend (USD)', 'spend_total', '#,##0', NAVY),
 ('Blended Cost / Acq (USD)', 'bcpa', '#,##0', NAVY),
]
col = {}
c = 23  # W
for name,key,fmt,color in ADD:
    if key: col[key] = c
    c += 1
LASTCOL = c-1

# original col formats (mirror BI look)
ORIG_FMT = {1:'d.m.yyyy',2:'#,##0',3:'#,##0',4:'#,##0',5:'0%',6:'#,##0',7:'0.0%',8:'#,##0',9:'0.0%',10:'#,##0',11:'0.0%',
            12:'0.0%',13:'0.0%',14:'0.0%',15:'#,##0',16:'0.0',17:'0.0%',18:'#,##0',19:'0.0',20:'#,##0.0;-#,##0.0;0',21:'#,##0.0;-#,##0.0;0'}

# ---------- title block ----------
ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=LASTCOL)
ws['A1'] = 'Longevity Life Academy — GAL Board Report · September 2026'
ws['A1'].font = F(size=18, bold=True, color=NAVY); ws['A1'].alignment = left
ws.row_dimensions[1].height = 50
ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=LASTCOL)
ws['A2'] = ('PINK rows = weeks 1–2 (1–11.9) exactly as already reported, not touched.  Weeks 3–5 (14–28.9) completed from By-Day pivot 38 (BI fields A–U) + Meta Ads spend & e-commerce results by campaign type '
            '(Ads API, account 1459085242361281, daily, USD).  Masterclass campaign not included in spend.')
ws['A2'].font = F(size=10, italic=True, color=MID); ws['A2'].alignment = Alignment(wrap_text=True, vertical='center')
ws.row_dimensions[2].height = 52

# ---------- group header (row 4) ----------
GH = 4; H = 5
def group(c1,c2,text,color,fg='FFFFFF'):
    ws.merge_cells(start_row=GH,start_column=c1,end_row=GH,end_column=c2)
    cell = ws.cell(GH,c1,text); cell.font=F(bold=True,color=fg,size=10); cell.fill=fill(color); cell.alignment=center
    for cc in range(c1,c2+1): ws.cell(GH,cc).fill=fill(color)
group(1,21,'BOARD FIELDS (weeks 1–2 as reported · weeks 3–5 from By-Day pivot 38)', NAVY)
group(col['spend_leads'],col['cpc'],'META SPEND — LEAD-GEN CAMPAIGNS (Leads #117186 · LLA/ATC #116622)', TEAL)
group(col['spend_ec'],col['ec_roas'],'META SPEND & RESULTS — E-COMMERCE BLUEPRINT (ecomm-01 #118148 · ecomm-02 #118149)', PLUM)
group(col['spend_total'],col['bcpa'],'TOTAL', NAVY)
ws.row_dimensions[GH].height = 34

# ---------- header row ----------
for i,h in enumerate(HDR, start=1):
    cell = ws.cell(H,i,h); cell.font=F(bold=True,color='FFFFFF',size=10); cell.fill=fill(INK); cell.alignment=center
    cell.border=Border(bottom=med)
c = 23
for name,key,fmt,color in ADD:
    if key:
        cell = ws.cell(H,c,name); cell.font=F(bold=True,color='FFFFFF',size=10); cell.fill=fill(color); cell.alignment=center; cell.border=Border(bottom=med)
    c+=1
ws.row_dimensions[H].height = 84
ws.sheet_format.defaultRowHeight = 30
ws.sheet_view.zoomScale = 110

# ---------- rows ----------
weeks = [  # (label, weekday dates, weekend dates, locked)
 ('Week 1 · 1–6.9', [1,2,3,4], [], True),
 ('Week 2 · 7–13.9', [7,8,9,10,11], [], True),
 ('Week 3 · 14–20.9', [14,15,16,17,18], [19,20], False),
 ('Week 4 · 21–27.9', [21,22,23,24,25], [26,27], False),
 ('Week 5 · 28.9 (to date)', [28], [], False),
]
D = lambda d: dt.date(2026,9,d)
row = H+1
day_rows = []      # rows holding daily data (weekday + weekend)
week_total_rows = []
hidden_col = LASTCOL+2  # helper column (hidden): BI contacts denominator
ws.cell(H,hidden_col,'BI cont. denom (helper)').font=F(size=8,color=MID)

def ratio(num, den):
    return f'=IF({den}=0,0,{num}/{den})'

def write_day(r, d, weekend=False, locked=False):
    date = D(d)
    ws.cell(r,1,date).number_format='d.m.yyyy'
    if not weekend and date in bi:
        vals = bi[date]
        for i in range(1,21):
            v = vals[i]
            if v is not None:
                ws.cell(r,i+1,v)
        ws.cell(r,hidden_col,hidden[date])
    else:
        ws.cell(r,hidden_col,0)
    if locked:
        return
    ls = leads_spend.get(d,0); ll = lla_spend.get(d,0); e = ec(d)
    ws.cell(r,col['spend_leads'],ls); ws.cell(r,col['spend_lla'],ll)
    ws.cell(r,col['spend_lg'],f'={L(col["spend_leads"])}{r}+{L(col["spend_lla"])}{r}')
    ws.cell(r,col['cpl'],ratio(f'{L(col["spend_lg"])}{r}',f'B{r}'))
    ws.cell(r,col['cpa'],ratio(f'{L(col["spend_lg"])}{r}',f'C{r}'))
    ws.cell(r,col['cpc'],ratio(f'{L(col["spend_lg"])}{r}',f'H{r}'))
    ws.cell(r,col['spend_ec'],e[0]); ws.cell(r,col['ec_atc'],e[1]); ws.cell(r,col['ec_ic'],e[2]); ws.cell(r,col['ec_p'],e[3]); ws.cell(r,col['ec_rev'],e[4])
    ws.cell(r,col['ec_cpatc'],ratio(f'{L(col["spend_ec"])}{r}',f'{L(col["ec_atc"])}{r}'))
    ws.cell(r,col['ec_cpp'],ratio(f'{L(col["spend_ec"])}{r}',f'{L(col["ec_p"])}{r}'))
    ws.cell(r,col['ec_roas'],ratio(f'{L(col["ec_rev"])}{r}',f'{L(col["spend_ec"])}{r}'))
    ws.cell(r,col['spend_total'],f'={L(col["spend_lg"])}{r}+{L(col["spend_ec"])}{r}')
    ws.cell(r,col['bcpa'],ratio(f'{L(col["spend_total"])}{r}',f'C{r}'))

def write_total(r, label, rows, style):
    """rows: list of row numbers to aggregate. Sums for count fields, ratio-of-sums for rates (mirrors BI grand total)."""
    rng = lambda c: ','.join(f'{L(c)}{x}' for x in rows) if len(rows)<=8 else None
    def S(c):
        # contiguous? use range, else list
        if rows == list(range(rows[0], rows[-1]+1)):
            return f'SUM({L(c)}{rows[0]}:{L(c)}{rows[-1]})'
        return 'SUM(' + ','.join(f'{L(c)}{x}' for x in rows) + ')'
    ws.cell(r,1,label)
    for cidx in [2,3,4,6,8,10,15,18,20,21]:
        ws.cell(r,cidx,f'={S(cidx)}')
    ws.cell(r,5,ratio(f'D{r}',f'C{r}'))          # Ratio = Regs/Acqs
    ws.cell(r,7,ratio(f'F{r}',f'B{r}'))          # %Cont from Lds
    ws.cell(r,9,ratio(f'H{r}',f'B{r}'))          # Cont%
    ws.cell(r,11,ratio(f'J{r}',f'B{r}'))         # % Phone Comm
    ws.cell(r,12,ratio(f'C{r}',f'B{r}'))         # % Acqs/Leads
    ws.cell(r,13,ratio(f'C{r}',f'H{r}'))         # % Acqs/Cont
    ws.cell(r,hidden_col,f'={S(hidden_col)}')
    ws.cell(r,14,ratio(f'D{r}',f'{L(hidden_col)}{r}'))   # % Regs/Cont (BI denominator)
    ws.cell(r,16,ratio(f'O{r}',f'B{r}'))         # Calls per Lead
    ws.cell(r,17,ratio(f'{L(hidden_col)}{r}',f'O{r}'))   # %Cont/Calls (BI denominator)
    ws.cell(r,19,ratio(f'R{r}',f'B{r}'))         # Retries per Lead
    for key in ['spend_leads','spend_lla','spend_ec','ec_atc','ec_ic','ec_p','ec_rev']:
        ws.cell(r,col[key],f'={S(col[key])}')
    ws.cell(r,col['spend_lg'],f'={L(col["spend_leads"])}{r}+{L(col["spend_lla"])}{r}')
    ws.cell(r,col['cpl'],ratio(f'{L(col["spend_lg"])}{r}',f'B{r}'))
    ws.cell(r,col['cpa'],ratio(f'{L(col["spend_lg"])}{r}',f'C{r}'))
    ws.cell(r,col['cpc'],ratio(f'{L(col["spend_lg"])}{r}',f'H{r}'))
    ws.cell(r,col['ec_cpatc'],ratio(f'{L(col["spend_ec"])}{r}',f'{L(col["ec_atc"])}{r}'))
    ws.cell(r,col['ec_cpp'],ratio(f'{L(col["spend_ec"])}{r}',f'{L(col["ec_p"])}{r}'))
    ws.cell(r,col['ec_roas'],ratio(f'{L(col["ec_rev"])}{r}',f'{L(col["spend_ec"])}{r}'))
    ws.cell(r,col['spend_total'],f'={L(col["spend_lg"])}{r}+{L(col["spend_ec"])}{r}')
    ws.cell(r,col['bcpa'],ratio(f'{L(col["spend_total"])}{r}',f'C{r}'))
    # style
    bg, fg, bold = style
    for cc in range(1,LASTCOL+1):
        cell = ws.cell(r,cc); cell.fill=fill(bg); cell.font=F(bold=bold,color=fg,size=10)
        cell.border=Border(top=thin,bottom=thin)
    ws.cell(r,1).alignment=left


PINK='FCE4EC'; PINK_D='AD1457'
def week_headline(r, label, locked):
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=LASTCOL)
    c=ws.cell(r,1,label + ('   ·   ALREADY REPORTED — NOT TOUCHED' if locked else '   ·   COMPLETED IN THIS UPDATE (results + Meta spend)'))
    c.font=F(size=13,bold=True,color=(PINK_D if locked else 'FFFFFF')); c.alignment=left
    for cc in range(1,LASTCOL+1):
        ws.cell(r,cc).fill=fill(PINK if locked else NAVY)
        ws.cell(r,cc).border=Border(top=med)
    ws.row_dimensions[r].height=30

pre_rows=[]; post_rows=[]; all_rows=[]
locked_rows=[]
week_starts=[]
for label, wd, we, locked in weeks:
    week_headline(row, label, locked); row+=1
    week_starts.append(row)
    wrows=[]
    for d in wd:
        write_day(row,d,locked=locked); wrows.append(row); all_rows.append(row)
        (pre_rows if locked else post_rows).append(row)
        if locked: locked_rows.append(row)
        row+=1
    if locked:
        write_total(row, f'{label} · total (computed from reported days)', wrows, (PINK, PINK_D, True))
        for key in col: ws.cell(row,col[key]).value=None
        week_total_rows.append(row); row+=1
        continue
    if we:
        # weekend row: Meta spend only (BI has no weekend rows)
        r=row
        ws.cell(r,1,f'Sat–Sun {we[0]}–{we[1]}.9 (spend only)')
        for i in range(2,22): ws.cell(r,i,None)
        ws.cell(r,hidden_col,0)
        for key,src in [('spend_leads',lambda d: leads_spend.get(d,0)),('spend_lla',lambda d: lla_spend.get(d,0))]:
            ws.cell(r,col[key],sum(src(d) for d in we))
        ws.cell(r,col['spend_lg'],f'={L(col["spend_leads"])}{r}+{L(col["spend_lla"])}{r}')
        e=[sum(ec(d)[i] for d in we) for i in range(5)]
        for key,v in zip(['spend_ec','ec_atc','ec_ic','ec_p','ec_rev'],e): ws.cell(r,col[key],v)
        ws.cell(r,col['ec_cpatc'],ratio(f'{L(col["spend_ec"])}{r}',f'{L(col["ec_atc"])}{r}'))
        ws.cell(r,col['ec_cpp'],ratio(f'{L(col["spend_ec"])}{r}',f'{L(col["ec_p"])}{r}'))
        ws.cell(r,col['ec_roas'],ratio(f'{L(col["ec_rev"])}{r}',f'{L(col["spend_ec"])}{r}'))
        ws.cell(r,col['spend_total'],f'={L(col["spend_lg"])}{r}+{L(col["spend_ec"])}{r}')
        for cc in range(1,LASTCOL+1):
            cell=ws.cell(r,cc); cell.font=F(italic=True,color=MID,size=9); cell.fill=fill(SOFT)
        wrows.append(r); all_rows.append(r); post_rows.append(r)
        row+=1
    write_total(row, f'{label} · total', wrows, (SOFT2, NAVY, True))
    week_total_rows.append(row); row+=1

row+=1
# period totals
write_total(row,'WEEKS 1–2 · 1–11.9 · as reported (no spend reported)', pre_rows, (AMBER_L, AMBER, True)); pre_tot=row
for key in col: ws.cell(row,col[key]).value=None
row+=1
write_total(row,'E-COMM WEEKS 3–5 · 14–28.9 · results + Meta spend', post_rows, (PLUM_L, PLUM, True)); post_tot=row; row+=1
write_total(row,'GRAND TOTAL · SEPTEMBER 2026 (spend = 14–28.9 only)', all_rows, (NAVY, 'FFFFFF', True)); grand=row
for key in ['cpl','cpa','cpc','bcpa']: ws.cell(grand,col[key]).value='n/a'; ws.cell(grand,col[key]).alignment=right
for cc in range(1,LASTCOL+1):
    ws.cell(grand,cc).border=Border(top=med,bottom=med)
ws.row_dimensions[grand].height=36
row+=1

# BI control line (from pivot Grand Total) for auditability
row+=1
ws.cell(row,1,'Control · BI pivot Grand Total (as exported)').font=F(italic=True,size=9,color=MID)
for i in range(1,21):
    v=bi_total[i]
    if v is not None:
        c=ws.cell(row,i+1,v); c.font=F(italic=True,size=9,color=MID); c.number_format=ORIG_FMT[i+1]
ctrl=row; row+=1
ws.cell(row,1,'Check · Grand Total − BI control (must be 0)').font=F(italic=True,size=9,color=MID)
for i in [2,3,4,6,8,10,15,18,20,21]:
    c=ws.cell(row,i,f'=ROUND({L(i)}{grand}-{L(i)}{ctrl},2)'); c.font=F(italic=True,size=9,color=MID); c.number_format='0.00;[Red]-0.00'
chk=row

# ---------- number formats, zebra, borders for day rows ----------
for r in range(H+1, grand+1):
    ws.row_dimensions[r].height = 30
    for cidx,fmt in ORIG_FMT.items():
        ws.cell(r,cidx).number_format=fmt
    c=23
    for name,key,fmt,color in ADD:
        if key: ws.cell(r,c).number_format=fmt
        c+=1
    ws.cell(r,hidden_col).number_format='0'
for r in all_rows:
    if not ws.cell(r,1).fill.start_color.rgb.endswith(SOFT):
        pass
    for cc in range(1,LASTCOL+1):
        cell=ws.cell(r,cc)
        if cell.font.italic: continue
        cell.font=F(size=10,color=INK)
        cell.border=Border(bottom=hair) if cell.border.top.style is None else Border(top=cell.border.top,bottom=hair)
    ws.cell(r,1).alignment=left
    if not ws.cell(r,1).font.italic: ws.cell(r,1).font=F(size=10,bold=True,color=INK)
# zero ratios -> show dash
for r in range(H+1, grand+1):
    for cidx in [5,7,9,11,12,13,14,16,17,19]:
        ws.cell(r,cidx).number_format = ORIG_FMT[cidx]+';-'+ORIG_FMT[cidx]+';"–"'
    for key in ['cpl','cpa','cpc','ec_cpatc','ec_cpp','ec_roas','bcpa']:
        f=ws.cell(r,col[key]).number_format; ws.cell(r,col[key]).number_format=f+';-'+f+';"–"'
    for key in ['ec_atc','ec_ic','ec_p','ec_rev','spend_lla','spend_ec']:
        f=ws.cell(r,col[key]).number_format.split(';')[0]; ws.cell(r,col[key]).number_format=f+';-'+f+';"–"'

# untouched rows -> pink
for r in locked_rows:
    for cc in range(1,22):
        cell=ws.cell(r,cc); cell.fill=fill(PINK); cell.font=F(size=10,bold=(cc==1),color=PINK_D)
# tint added blocks lightly on day rows
for r in all_rows:
    if ws.cell(r,1).font.italic or r in pre_rows: continue
    for key in ['spend_leads','spend_lla','spend_lg','cpl','cpa','cpc']: ws.cell(r,col[key]).fill=fill(TEAL_L)
    for key in ['spend_ec','ec_atc','ec_ic','ec_p','ec_rev','ec_cpatc','ec_cpp','ec_roas']: ws.cell(r,col[key]).fill=fill(PLUM_L)
    for key in ['spend_total','bcpa']: ws.cell(r,col[key]).fill=fill(SOFT2)
    ws.cell(r,col['spend_total']).font=F(size=10,bold=True,color=NAVY)

# data bars on Total Spend & Leads (day rows only)
first,last=post_rows[0],all_rows[-1]
ws.conditional_formatting.add(f'{L(col["spend_total"])}{first}:{L(col["spend_total"])}{last}', DataBarRule(start_type='num',start_value=0,end_type='max',color='9DB4D6',showValue=True))
ws.conditional_formatting.add(f'B{first}:B{last}', DataBarRule(start_type='num',start_value=0,end_type='max',color='C7D2E3',showValue=True))
# highlight acquisitions
ws.conditional_formatting.add(f'C{first}:C{last}', CellIsRule(operator='greaterThan', formula=['0'], font=Font(bold=True,color=GREEN), fill=fill(GREEN_L)))
ws.conditional_formatting.add(f'{L(col["ec_p"])}{first}:{L(col["ec_p"])}{last}', CellIsRule(operator='greaterThan', formula=['0'], font=Font(bold=True,color=GREEN), fill=fill(GREEN_L)))

# widths
widths = {1:66,2:13,3:12,4:12,5:12,6:14,7:14,8:12,9:13,10:13,11:14,12:17,13:17,14:17,15:13,16:14,17:14,18:14,19:14,20:15,21:15,22:2}
for k,v in widths.items(): ws.column_dimensions[L(k)].width=v
c=23
for name,key,fmt,color in ADD:
    ws.column_dimensions[L(c)].width = 18 if key else 2
    c+=1
ws.column_dimensions[L(hidden_col)].hidden=True
ws.column_dimensions[L(LASTCOL+1)].width=2
ws.freeze_panes = f'B{H+1}'
ws.sheet_view.rightToLeft = False
ws.print_title_rows=f'{GH}:{H}'
ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=8; ws.page_setup.fitToWidth=3; ws.page_setup.fitToHeight=0; ws.print_title_cols='A:A'
ws.sheet_properties.pageSetUpPr.fitToPage=True

# ================= Sheet 2: campaign x period summary =================
s2 = wb.create_sheet('Spend by Campaign'); s2.sheet_view.showGridLines=False; s2.sheet_properties.tabColor=TEAL
s2['A1']='Meta Spend by Campaign × Week — 14–28.9.2026 (USD)'; s2['A1'].font=F(size=16,bold=True,color=NAVY)
s2['A2']='Source: Meta Ads API daily pull (account 1459085242361281). Weeks are Mon–Sun incl. weekend spend. Weeks 1–2 are not restated. ecomm-02 relaunched 24.9, first spend 28.9. Masterclass excluded.'
s2['A2'].font=F(size=10,italic=True,color=MID)
periods = [('Week 3\n14–20.9',list(range(14,21))),('Week 4\n21–27.9',list(range(21,28))),('Week 5\n28.9',[28]),('TOTAL\n14–28.9',list(range(14,29)))]
camps = [
 ('Lead-gen','Leads campaign · FB Leads #117186', lambda d: leads_spend.get(d,0)),
 ('Lead-gen','LLA / Adds-to-cart (leads) campaign · #116622', lambda d: lla_spend.get(d,0)),
 ('E-commerce','Blueprint e-comm · ecomm-01 #118148', lambda d: ec1.get(d,(0,))[0]),
 ('E-commerce','Blueprint e-comm · ecomm-02 #118149', lambda d: ec2.get(d,(0,))[0]),
]
hr=4
s2.cell(hr,1,'Type'); s2.cell(hr,2,'Campaign')
for j,(p,_) in enumerate(periods): s2.cell(hr,3+j,p)
for cc in range(1,3+len(periods)):
    cell=s2.cell(hr,cc); cell.font=F(bold=True,color='FFFFFF'); cell.fill=fill(INK); cell.alignment=center
s2.row_dimensions[hr].height=34
r=hr+1
for typ,name,fn in camps:
    s2.cell(r,1,typ); s2.cell(r,2,name)
    for j,(p,days) in enumerate(periods):
        c=s2.cell(r,3+j,round(sum(fn(d) for d in days),2)); c.number_format='#,##0;-#,##0;"–"'
    for cc in range(1,3+len(periods)):
        s2.cell(r,cc).font=F(size=10,color=INK); s2.cell(r,cc).border=Border(bottom=hair)
        if typ=='E-commerce': s2.cell(r,cc).fill=fill(PLUM_L)
        else: s2.cell(r,cc).fill=fill(TEAL_L)
    r+=1
# subtotal rows
def s2_total(r,label,rows,bg,fg):
    s2.cell(r,1,label)
    s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    for j in range(len(periods)):
        cl=L(3+j); c=s2.cell(r,3+j,'='+'+'.join(f'{cl}{x}' for x in rows)); c.number_format='#,##0;-#,##0;"–"'
    for cc in range(1,3+len(periods)):
        s2.cell(r,cc).font=F(bold=True,color=fg); s2.cell(r,cc).fill=fill(bg); s2.cell(r,cc).border=Border(top=thin,bottom=thin)
s2_total(r,'Lead-gen spend',[hr+1,hr+2],SOFT2,TEAL); r+=1
s2_total(r,'E-commerce spend',[hr+3,hr+4],SOFT2,PLUM); r+=1
s2_total(r,'TOTAL META SPEND',[hr+1,hr+2,hr+3,hr+4],NAVY,'FFFFFF'); tot=r; r+=1
s2.cell(r,1,'E-comm share of spend'); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
for j in range(len(periods)):
    cl=L(3+j); c=s2.cell(r,3+j,f'=IF({cl}{tot}=0,0,{cl}{tot-1}/{cl}{tot})'); c.number_format='0%;-0%;"–"'; c.font=F(italic=True,color=MID)
s2.cell(r,1).font=F(italic=True,color=MID); r+=2

# e-comm results block
s2.cell(r,1,'E-commerce Blueprint results (Meta-attributed, ecomm-01 + ecomm-02)').font=F(bold=True,color=PLUM,size=11); r+=1
s2.cell(r,1,'Metric'); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
for j,(p,_) in enumerate(periods): s2.cell(r,3+j,p)
for cc in range(1,3+len(periods)):
    cell=s2.cell(r,cc); cell.font=F(bold=True,color='FFFFFF'); cell.fill=fill(PLUM); cell.alignment=center
s2.row_dimensions[r].height=34; r+=1
metrics=[('Spend (USD)',0,'#,##0'),('Adds to cart',1,'#,##0'),('Checkouts initiated',2,'#,##0'),('Purchases',3,'#,##0'),('Revenue (USD)',4,'#,##0')]
base=r
for name,idx,fmt in metrics:
    s2.cell(r,1,name); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    for j,(p,days) in enumerate(periods):
        c=s2.cell(r,3+j,round(sum(ec(d)[idx] for d in days),2)); c.number_format=fmt+';-'+fmt+';"–"'
    for cc in range(1,3+len(periods)): s2.cell(r,cc).font=F(size=10,color=INK); s2.cell(r,cc).border=Border(bottom=hair)
    r+=1
for name,num,den,fmt in [('Cost / Add to cart (USD)',base+1,base,'#,##0.0'),('Cost / Checkout (USD)',base+2,base,'#,##0.0'),('Cost / Purchase (USD)',base+3,base,'#,##0'),('ROAS',base+4,base,'0.00')]:
    s2.cell(r,1,name); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    for j in range(len(periods)):
        cl=L(3+j)
        if name=='ROAS': f=f'=IF({cl}{den}=0,0,{cl}{num}/{cl}{den})'
        else: f=f'=IF({cl}{num}=0,0,{cl}{den}/{cl}{num})'
        c=s2.cell(r,3+j,f); c.number_format=fmt+';-'+fmt+';"–"'
    for cc in range(1,3+len(periods)): s2.cell(r,cc).font=F(bold=True,size=10,color=PLUM); s2.cell(r,cc).fill=fill(PLUM_L); s2.cell(r,cc).border=Border(bottom=hair)
    r+=1
r+=1
s2.cell(r,1,'Lead-gen efficiency (BI results ÷ lead-gen Meta spend)').font=F(bold=True,color=TEAL,size=11); r+=1
s2.cell(r,1,'Metric'); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
for j,(p,_) in enumerate(periods): s2.cell(r,3+j,p)
for cc in range(1,3+len(periods)):
    cell=s2.cell(r,cc); cell.font=F(bold=True,color='FFFFFF'); cell.fill=fill(TEAL); cell.alignment=center
s2.row_dimensions[r].height=34; r+=1
def bi_sum(days, idx):
    return sum((bi[D(d)][idx] or 0) for d in days if D(d) in bi)
lg_base=r
for name,idx,fmt in [('Leads (BI)',1,'#,##0'),('Acqs (BI)',2,'#,##0'),('Cont (BI)',7,'#,##0'),('Contract Price (BI)',19,'#,##0')]:
    s2.cell(r,1,name); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    for j,(p,days) in enumerate(periods):
        c=s2.cell(r,3+j,bi_sum(days,idx)); c.number_format=fmt+';-'+fmt+';"–"'
    for cc in range(1,3+len(periods)): s2.cell(r,cc).font=F(size=10,color=INK); s2.cell(r,cc).border=Border(bottom=hair)
    r+=1
lgspend=hr+5
for name,num,fmt in [('CPL (USD)',lg_base,'#,##0.0'),('Cost / Acq (USD)',lg_base+1,'#,##0'),('Cost / Cont (USD)',lg_base+2,'#,##0.0')]:
    s2.cell(r,1,name); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    for j in range(len(periods)):
        cl=L(3+j); c=s2.cell(r,3+j,f'=IF({cl}{num}=0,0,{cl}{lgspend}/{cl}{num})'); c.number_format=fmt+';-'+fmt+';"–"'
    for cc in range(1,3+len(periods)): s2.cell(r,cc).font=F(bold=True,size=10,color=TEAL); s2.cell(r,cc).fill=fill(TEAL_L); s2.cell(r,cc).border=Border(bottom=hair)
    r+=1
s2.cell(r,1,'Blended Cost / Acq (total Meta spend ÷ BI Acqs)'); s2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
for j in range(len(periods)):
    cl=L(3+j); c=s2.cell(r,3+j,f'=IF({cl}{lg_base+1}=0,0,{cl}{tot}/{cl}{lg_base+1})'); c.number_format='#,##0;-#,##0;"–"'
for cc in range(1,3+len(periods)): s2.cell(r,cc).font=F(bold=True,size=10,color=NAVY); s2.cell(r,cc).fill=fill(SOFT2); s2.cell(r,cc).border=Border(top=thin,bottom=thin)
r+=2
s2.cell(r,1,'Note: BI reports business days only; Meta spend includes weekends. Blueprint purchases are Meta-attributed pixel events and are not reconciled to settled payments.').font=F(size=9,italic=True,color=MID)
s2.column_dimensions['A'].width=18; s2.column_dimensions['B'].width=60
for j in range(len(periods)): s2.column_dimensions[L(3+j)].width=17
s2.freeze_panes='C5'
s2.sheet_view.rightToLeft=False

# ================= Sheet 3: sources & reconciliation =================
s3=wb.create_sheet('Sources & Checks'); s3.sheet_view.showGridLines=False; s3.sheet_properties.tabColor=MID
s3['A1']='Sources, definitions & reconciliation'; s3['A1'].font=F(size=16,bold=True,color=NAVY)
rows3=[
 ('Board fields (cols A–U)','Weeks 1–2 (1–11.9) left exactly as in the uploaded GAL file. Weeks 3–5 (14–28.9) filled from By-Day-pivot-38.xlsx · sheet "pivot" · copied 1:1, no field renamed, reordered or recalculated. Weekly/period totals use sums for counts and ratio-of-sums for rates (same method as the BI Grand Total; % Regs/Cont and %Cont/Calls use the BI contact denominator implied by the pivot).'),
 ('Meta spend & e-comm results','Meta Ads API, ad account 1459085242361281 (Longevity Life Academy · ETeacher Group), campaign level, daily (time_increment=1), 1–28.9.2026, USD, pulled 28.9.2026 12:00 IDT. 28.9 is a partial day.'),
 ('Campaign mapping','Leads = LGV_EN_FB_Demographics_FB Leads_2026-08-03_#117186 · LLA/Adds-to-cart (leads) = LGV_EN_FB_Demographics_Longevity Life Academy_2026-05-04_#116622 · E-commerce Blueprint = LGV_EN_PPC_ecomm-01_2026-09-02_#118148 (first spend 3.9) + LGV_EN_PPC_ecomm-02_2026-09-02_#118149 (relaunched 24.9, first spend 28.9). Two inactive campaigns had $0 spend. Masterclass campaign: not included.'),
 ('Period split','Weeks 1–2 = as previously reported (results only, no spend added). Weeks 3–5 = completed with results + Meta spend, split lead-gen vs e-commerce Blueprint. Weeks = Mon–Sun; weekend spend on separate "Sat–Sun (spend only)" rows so weekly spend reconciles exactly to Ads Manager. Note: Ads Manager shows Blueprint (ecomm-01) spend from 3.9.2026; per instruction weeks 1–2 were not restated.'),
 ('Test exclusion','ecomm-02 contains two TEST ad sets created 28.9 (controlled-purchase, TLV) with $0 spend and 0 results — nothing from them is in this file.'),
 ('Not verified','Blueprint purchases (3 × $179, on 10/13/16.9) are Meta-attributed pixel events, not reconciled to settled payments. BI weekend leads are not in the pivot; Meta shows leads on weekends, so CPL is computed on business-day leads only.'),
]
r=3
for k,v in rows3:
    s3.cell(r,1,k).font=F(bold=True,color=NAVY); s3.cell(r,2,v).font=F(size=10,color=INK); s3.cell(r,2).alignment=Alignment(wrap_text=True,vertical='top'); s3.cell(r,1).alignment=Alignment(vertical='top')
    s3.row_dimensions[r].height=95; r+=1
r+=1
s3.cell(r,1,'Reconciliation · Meta API daily pull vs uploaded Ads Manager CSV (7–28.9.2026)').font=F(bold=True,color=NAVY,size=11); r+=1
for j,h in enumerate(['Campaign','API daily sum (USD)','Ads Manager CSV (USD)','Δ (USD)']):
    c=s3.cell(r,1+j,h); c.font=F(bold=True,color='FFFFFF'); c.fill=fill(INK); c.alignment=center
r+=1
for k,(a,b) in rec.items():
    s3.cell(r,1,k); s3.cell(r,2,round(a,2)).number_format='#,##0.00'; s3.cell(r,3,b).number_format='#,##0.00'; s3.cell(r,4,f'=B{r}-C{r}').number_format='0.00;[Red]-0.00'
    for cc in range(1,5): s3.cell(r,cc).border=Border(bottom=hair); s3.cell(r,cc).font=F(size=10)
    r+=1
s3.cell(r,1,'Δ within ±$1 = live-day drift between the CSV export time and the API pull (ecomm-01/-02 still delivering).').font=F(size=9,italic=True,color=MID)
s3.sheet_view.rightToLeft=False
s3.column_dimensions['A'].width=34; s3.column_dimensions['B'].width=130; s3.column_dimensions['C'].width=22; s3.column_dimensions['D'].width=12

wb.save(OUT)
print('saved', OUT, 'grand row', grand, 'lastcol', LASTCOL)

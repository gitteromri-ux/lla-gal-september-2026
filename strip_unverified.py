import openpyxl, warnings
warnings.filterwarnings('ignore')
src='Longevity GAL September 2026.xlsx'
wb=openpyxl.load_workbook(src)
ws=wb['GAL Sep-26']
for c in ['AB','AE','AF','AG','AH','AI','AJ','AK']:
    ws.column_dimensions[c].hidden=True
for r in range(1,6):
    for c in range(1,ws.max_column+1):
        v=ws.cell(r,c).value
        if isinstance(v,str) and v.startswith('META SPEND & RESULTS — E-COMMERCE BLUEPRINT'):
            ws.cell(r,c).value='META SPEND — E-COMMERCE BLUEPRINT (spend only · pixel ATC/checkout/purchase EXCLUDED: not verified, may include tracking tests)'
        if isinstance(v,str) and v.startswith('PINK rows'):
            ws.cell(r,c).value=v+' Pixel e-commerce results and Cost/Cont removed — unverified.'
s2=wb['Spend by Campaign']
s2.cell(14,1).value='E-commerce Blueprint results — EXCLUDED (pixel events not verified; may include tracking tests; not reconciled to settled payments)'
for r in list(range(15,25))+[30,34]:
    s2.row_dimensions[r].hidden=True
s2.cell(37,1).value='Note: BI reports business days only; Meta spend includes weekends. Pixel-attributed Blueprint ATC/checkout/purchase counts are excluded from this report: unverified, may include tracking tests, not reconciled to settled payments.'
s3=wb['Sources & Checks']
for row in s3.iter_rows():
    for cell in row:
        if isinstance(cell.value,str) and cell.value.startswith('Blueprint purchases (3'):
            cell.value='Pixel e-commerce results (ATC, checkouts, purchases, revenue) are EXCLUDED from the board sheet: Meta-attributed pixel events, possibly including tracking tests, not reconciled to settled payments. Only Meta spend is reported for the Blueprint campaigns.'
wb.save('Longevity-GAL-September-2026.xlsx')
print('stripped')

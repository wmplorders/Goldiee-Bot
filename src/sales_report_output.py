"""Portable Excel and PNG rendering for the monthly sales/payment report."""
import calendar
import os
import tempfile
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

INDIAN_FORMAT='[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;#,##0'
DARK='35464A';TEXT='293237';GRID='E2E7E9';STRIPE='F3F5F6'
CATEGORIES=[('SPICES SALE','F3EDDC','68542E'),('DET. SALE','E7EFF5','3B5B73'),('PAYMENT','E5F0E8','316246')]


def indian(value):
    if value is None:return ''
    rounded=int(Decimal(str(value)).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
    sign='-' if rounded<0 else '';digits=str(abs(rounded))
    if len(digits)<=3:return sign+digits
    tail=digits[-3:];head=digits[:-3];parts=[]
    while head:parts.insert(0,head[-2:]);head=head[:-2]
    return sign+','.join(parts+[tail])


def excel(payload,path):
    end=date.fromisoformat(payload['end']);days=payload['days'];rows=payload['rows'];total=len(rows)+3
    lastcol=2+days*3+4;wb=Workbook();ws=wb.active;ws.title='Sales and Payment'
    ws.sheet_view.showGridLines=False;ws.freeze_panes='C3'
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=2)
    ws.cell(1,1,'SALE & PAYMENT REPORT - '+end.strftime('%B %Y').upper())
    ws.cell(2,1,'PARTY NAME');ws.cell(2,2,'POST CODE')
    for day in range(1,days+1):
        first=3+(day-1)*3
        ws.merge_cells(start_row=1,start_column=first,end_row=1,end_column=first+2)
        ws.cell(1,first,date(end.year,end.month,day));ws.cell(1,first).number_format='d-mmm'
        for j,(label,_,_) in enumerate(CATEGORIES):ws.cell(2,first+j,label)
    totals_first=3+days*3
    for j,label in enumerate(('SPICES TOTAL','DET. TOTAL','DEBIT BALANCE','CREDIT BALANCE')):
        ws.merge_cells(start_row=1,start_column=totals_first+j,end_row=2,end_column=totals_first+j)
        ws.cell(1,totals_first+j,label)
    start=date.fromisoformat(payload['start'])
    for r,record in enumerate(rows,3):
        ws.cell(r,1,record['name']);ws.cell(r,2,record['town'])
        for c,value in enumerate(record['values'],3):
            if value is not None:ws.cell(r,c,value)
        for j in (0,1):
            cells=','.join(f'{get_column_letter(3+i*3+j)}{r}' for i in range(days))
            ws.cell(r,totals_first+j,'=SUM('+cells+')')
        if record['debit'] is not None:ws.cell(r,totals_first+2,record['debit'])
        if record['credit'] is not None:ws.cell(r,totals_first+3,record['credit'])
    ws.merge_cells(start_row=total,start_column=1,end_row=total,end_column=2);ws.cell(total,1,'Grand Total')
    for c in range(3,lastcol+1):
        if c<totals_first and not start.day<=((c-3)//3+1)<=end.day:continue
        letter=get_column_letter(c);ws.cell(total,c,f'=SUM({letter}3:{letter}{total-1})')
    faint=Side(style='thin',color=GRID);bold=Side(style='medium',color='52656C')
    for row in ws:
        ws.row_dimensions[row[0].row].height=15
        for cell in row:
            cell.font=Font(name='Arial',size=10,color=TEXT)
            cell.alignment=Alignment(horizontal='left' if cell.column<3 else 'right',vertical='center')
            cell.border=Border(left=faint,right=faint,top=faint,bottom=faint)
            cell.fill=PatternFill('solid',fgColor=STRIPE if cell.row%2==0 and cell.row>2 else 'FFFFFF')
            if cell.column>=3 and cell.row>=3:cell.number_format=INDIAN_FORMAT
    for day in range(1,days+1):
        first=3+(day-1)*3
        for c in range(first,first+3):
            for r in (1,2):ws.cell(r,c).alignment=Alignment(horizontal='center',vertical='center')
            _,fill,color=CATEGORIES[c-first]
            ws.cell(2,c).fill=PatternFill('solid',fgColor=fill);ws.cell(2,c).font=Font(name='Arial',size=10,bold=True,color=color)
            if c-first==2:
                for r in range(3,total):ws.cell(r,c).font=Font(name='Arial',size=10,color=color)
            if not start.day<=day<=end.day:
                for r in range(3,total):ws.cell(r,c).fill=PatternFill('solid',fgColor='F7F8F9')
        ws.cell(1,first).fill=PatternFill('solid',fgColor='E9DDBB' if day==end.day else 'E7ECEE')
        ws.cell(1,first).font=Font(name='Arial',size=10,bold=True,color=DARK)
    fills=['F3EDDC','E7EFF5','F4E9E6','E5F0E8'];colors=['68542E','3B5B73','784D43','316246']
    for j in range(4):
        c=totals_first+j
        for r in (1,2):
            ws.cell(r,c).fill=PatternFill('solid',fgColor=fills[j]);ws.cell(r,c).font=Font(name='Arial',size=10,bold=True,color=colors[j]);ws.cell(r,c).alignment=Alignment(horizontal='center',vertical='center')
    for c in (1,2):
        ws.cell(1,c).fill=PatternFill('solid',fgColor=DARK)
        ws.cell(1,c).font=Font(name='Arial',size=11,bold=True,color='FFFFFF')
        ws.cell(2,c).fill=PatternFill('solid',fgColor='E7ECEE');ws.cell(2,c).font=Font(name='Arial',size=10,bold=True,color=DARK)
    for cell in ws[total]:cell.fill=PatternFill('solid',fgColor=DARK);cell.font=Font(name='Arial',size=10,bold=True,color='FFFFFF')
    for day in range(1,days+1):
        a=3+(day-1)*3;b=a+2
        for r in range(1,total+1):
            for c in range(a,b+1):
                cell=ws.cell(r,c);old=cell.border
                cell.border=Border(left=bold if c==a else old.left,right=bold if c==b else old.right,
                    top=bold if r==1 else old.top,bottom=bold if r==total else old.bottom)
    ws.column_dimensions['A'].width=64;ws.column_dimensions['B'].width=21
    for c in range(3,lastcol+1):ws.column_dimensions[get_column_letter(c)].width=13 if c<totals_first else 20
    # Make untrusted ledger/town labels literal strings, not spreadsheet formulas.
    for r in range(3,total):
        for c in (1,2):ws.cell(r,c).data_type='s'
    wb.calculation.fullCalcOnLoad=True
    wb.save(path)


def snapshot_rows(payload):
    end=date.fromisoformat(payload['end']);previous=end-timedelta(days=1);start=date.fromisoformat(payload['start'])
    result=[]
    for record in payload['rows']:
        daily=[]
        for day in (previous,end):
            if start<=day<=end:
                i=(day.day-1)*3;daily.extend(record['values'][i:i+3])
            else:daily.extend([None]*3)
        spices=sum(Decimal(str(x or 0)) for x in record['values'][0::3])
        detergent=sum(Decimal(str(x or 0)) for x in record['values'][1::3])
        result.append([record['name'],record['town'],*daily,float(spices),float(detergent),record['debit'],record['credit']])
    totals=['Grand Total','']
    for c in range(2,12):
        # Keep the preceding date blank if it is outside the selected range.
        if c<5 and previous<start:totals.append(None)
        else:totals.append(float(sum((Decimal(str(r[c] or 0)) for r in result),Decimal(0))))
    return previous,end,result+[totals]


def png(payload,path):
    from PIL import Image, ImageDraw, ImageFont
    previous,end,rows=snapshot_rows(payload)
    def font(size,bold=False):
        candidates=[Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'/('arialbd.ttf' if bold else 'arial.ttf'),
                    Path('/usr/share/fonts/truetype/dejavu')/('DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf')]
        for f in candidates:
            if f.exists():return ImageFont.truetype(str(f),size)
        return ImageFont.load_default(size=size)
    normal,bold,title=font(18),font(18,True),font(26,True)
    probe=ImageDraw.Draw(Image.new('RGB',(1,1)))
    widths=[max(590,int(max(probe.textlength(r[0],font=normal) for r in rows))+24),
            max(180,int(max(probe.textlength(r[1],font=normal) for r in rows))+24)]+[140]*6+[180]*4
    xs=[0]
    for w in widths:xs.append(xs[-1]+w)
    heading,subtitle,header,rowheight=48,38,64,32
    top=heading+subtitle+header;height=top+rowheight*len(rows)
    image=Image.new('RGB',(xs[-1]+2,height+2),'white');draw=ImageDraw.Draw(image)
    draw.rectangle((0,0,xs[-1],heading),fill='#'+DARK)
    draw.text((10,10),'GOLDIEE - SALES AND PAYMENT SNAPSHOT',font=title,fill='white')
    sub=f'Daily detail: {previous:%d-%b-%Y} and {end:%d-%b-%Y}. Totals: {date.fromisoformat(payload["start"]):%d-%b-%Y} to {end:%d-%b-%Y}. Balances: {end:%d-%b-%Y}.'
    draw.text((10,heading+10),sub,font=normal,fill='#58676D')
    y=heading+subtitle
    draw.rectangle((0,y,xs[-1],top),fill='#E7ECEE')
    for i,label in [(0,'PARTY NAME'),(1,'POST CODE'),(8,'SPICES TOTAL'),(9,'DET. TOTAL'),(10,'DEBIT BALANCE'),(11,'CREDIT BALANCE')]:
        draw.text((xs[i]+8,y+24),label,font=bold,fill='#'+DARK)
    for day,first in [(previous,2),(end,5)]:
        draw.rectangle((xs[first],y,xs[first+3],y+32),fill='#E9DDBB' if first==5 else '#E7ECEE')
        label=day.strftime('%d-%b-%Y');draw.text(((xs[first]+xs[first+3]-draw.textlength(label,font=bold))/2,y+7),label,font=bold,fill='#'+DARK)
        for j,(label,fill,color) in enumerate(CATEGORIES):
            c=first+j;draw.rectangle((xs[c],y+32,xs[c+1],top),fill='#'+fill)
            short=['SPICES','DETERGENT','PAYMENT'][j]
            draw.text((xs[c]+8,y+39),short,font=bold,fill='#'+color)
    for i,row in enumerate(rows):
        y=top+i*rowheight;is_total=i==len(rows)-1
        draw.rectangle((0,y,xs[-1],y+rowheight),fill='#'+DARK if is_total else ('#'+STRIPE if i%2 else 'white'))
        for c,value in enumerate(row):
            label=str(value or '') if c<2 else indian(value)
            f=bold if is_total else normal
            color='white' if is_total else ('#316246' if c in (4,7) else '#'+TEXT)
            x=xs[c]+8 if c<2 else xs[c+1]-draw.textlength(label,font=f)-8
            draw.text((x,y+6),label,font=f,fill=color)
        draw.line((0,y+rowheight,xs[-1],y+rowheight),fill='#'+GRID,width=1)
    for x in xs:draw.line((x,top,x,height),fill='#'+GRID,width=1)
    for first in (2,5):draw.rectangle((xs[first],heading+subtitle,xs[first+3],height),outline='#52656C',width=3)
    image.save(path,format='PNG')


def build(payload,destination,stem):
    destination=Path(destination);destination.mkdir(parents=True,exist_ok=True)
    xlsx=destination/(stem+'.xlsx');image=destination/(stem+' Snapshot.png')
    with tempfile.TemporaryDirectory(prefix='sales-report-',dir=destination) as tmp:
        staged_xlsx=Path(tmp)/xlsx.name;staged_image=Path(tmp)/image.name
        excel(payload,staged_xlsx);png(payload,staged_image)
        staged_xlsx.replace(xlsx);staged_image.replace(image)
    return xlsx,image

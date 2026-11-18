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


def _wrap_snapshot_text(value,width,font,draw):
    """Wrap full labels, including unbroken ledger codes, without truncation."""
    lines=[];current=''
    for word in str(value).split():
        candidate=(current+' '+word).strip()
        if draw.textlength(candidate,font=font)<=width:
            current=candidate;continue
        if current:lines.append(current);current=''
        while draw.textlength(word,font=font)>width:
            split=1
            while split<len(word) and draw.textlength(word[:split+1],font=font)<=width:split+=1
            lines.append(word[:split]);word=word[split:]
        current=word
    if current:lines.append(current)
    return lines or ['']


def png(payload,path):
    from PIL import Image, ImageDraw, ImageFont
    import math
    previous,end,rows=snapshot_rows(payload)
    scale=3
    def font(size,bold=False):
        pixels=round(size*scale)
        candidates=[Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'/('arialbd.ttf' if bold else 'arial.ttf'),
                    Path('/usr/share/fonts/truetype/dejavu')/('DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf')]
        for f in candidates:
            if f.exists():return ImageFont.truetype(str(f),pixels)
        return ImageFont.load_default(size=pixels)
    regular,bold,small,smallbold,title=font(15),font(15,True),font(12.5),font(12.5,True),font(20,True)
    probe=ImageDraw.Draw(Image.new('RGB',(1,1)))
    widths=[305,108,90,82,90,90,82,90,102,92,104,94]
    # Widen only exceptional numeric columns; normal reports retain the compact layout.
    for c in range(2,12):
        widest=max(probe.textlength(indian(row[c]),font=f) for row in rows for f in (regular,bold))
        widths[c]=max(widths[c],math.ceil(widest/scale)+10)
    xs=[0]
    for width in widths:xs.append(xs[-1]+width)
    names=[_wrap_snapshot_text(r[0],(widths[0]-12)*scale,smallbold,probe) for r in rows]
    towns=[_wrap_snapshot_text(r[1],(widths[1]-10)*scale,small,probe) for r in rows]
    heights=[max(32,9+max(len(n),len(t))*15) for n,t in zip(names,towns)]
    heading,subtitle,header=36,22,50
    top=heading+subtitle+header;height=top+sum(heights)
    image=Image.new('RGB',(xs[-1]*scale+2,height*scale+2),'white');draw=ImageDraw.Draw(image)
    def rect(box,fill=None,outline=None,width=1):
        draw.rectangle(tuple(round(x*scale) for x in box),fill=fill,outline=outline,width=round(width*scale))
    def line(box,fill,width=1):draw.line(tuple(round(x*scale) for x in box),fill=fill,width=round(width*scale))
    def label(x,y,value,f,color='#293237',align='left'):
        value=str(value)
        if align=='right':x-=draw.textlength(value,font=f)/scale
        elif align=='center':x-=draw.textlength(value,font=f)/scale/2
        draw.text((round(x*scale),round(y*scale)),value,font=f,fill=color)
    rect((0,0,xs[-1],heading),fill='#'+DARK)
    label(8,6,'GOLDIEE - SALES & PAYMENTS',title,'white')
    start=date.fromisoformat(payload['start'])
    label(8,heading+4,f'Totals: {start:%d}-{end:%d %b %Y}. Closing balances as at {end:%d %b %Y}.',small,'#58676D')
    y=heading+subtitle
    rect((0,y,xs[-1],top),fill='#E7ECEE')
    label(6,y+18,'PARTY NAME',smallbold)
    label(xs[1]+5,y+18,'POST CODE',smallbold)
    for day,first in [(previous,2),(end,5)]:
        rect((xs[first],y,xs[first+3],y+24),fill='#E9DDBB' if first==5 else '#E7ECEE')
        label((xs[first]+xs[first+3])/2,y+4,day.strftime('%d %b'),bold,align='center')
        for j,(_,fill,color) in enumerate(CATEGORIES):
            c=first+j;rect((xs[c],y+24,xs[c+1],top),fill='#'+fill)
            label((xs[c]+xs[c+1])/2,y+29,['SPICES','DET.','PAYMENT'][j],smallbold,'#'+color,align='center')
    for c,labels,fill,color in [(8,['SPICES','TOTAL'],'F3EDDC','68542E'),(9,['DET.','TOTAL'],'E7EFF5','3B5B73'),
                               (10,['DEBIT','BALANCE'],'F4E9E6','784D43'),(11,['CREDIT','BALANCE'],'E5F0E8','316246')]:
        rect((xs[c],y,xs[c+1],top),fill='#'+fill)
        for k,value in enumerate(labels):label((xs[c]+xs[c+1])/2,y+8+k*17,value,smallbold,'#'+color,align='center')
    y=top
    for i,(row,rowheight,name,town) in enumerate(zip(rows,heights,names,towns)):
        is_total=i==len(rows)-1
        rect((0,y,xs[-1],y+rowheight),fill='#'+DARK if is_total else ('#F2F5F6' if i%2 else 'white'))
        color='white' if is_total else '#'+TEXT
        for lines,x,f in [(name,6,smallbold),(town,xs[1]+5,small)]:
            for k,value in enumerate(lines):label(x,y+(rowheight-len(lines)*15)/2+k*15-1,value,f,color)
        for c in range(2,12):
            label(xs[c+1]-5,y+(rowheight-18)/2-1,indian(row[c]),bold if is_total else regular,
                  'white' if is_total else ('#316246' if c in (4,7) else '#'+TEXT),align='right')
        line((0,y+rowheight,xs[-1],y+rowheight),'#DCE3E6',.6);y+=rowheight
    for x in xs:line((x,top,x,height),'#DCE3E6',.6)
    for first in (2,5):rect((xs[first],heading+subtitle,xs[first+3],height),outline='#52656C',width=1.5)
    image.save(path,format='PNG',optimize=True,dpi=(300,300))
    image.close()


def build(payload,destination,stem):
    destination=Path(destination);destination.mkdir(parents=True,exist_ok=True)
    xlsx=destination/(stem+'.xlsx');image=destination/(stem+' Snapshot.png')
    with tempfile.TemporaryDirectory(prefix='sales-report-',dir=destination) as tmp:
        staged_xlsx=Path(tmp)/xlsx.name;staged_image=Path(tmp)/image.name
        excel(payload,staged_xlsx);png(payload,staged_image)
        staged_xlsx.replace(xlsx);staged_image.replace(image)
    return xlsx,image

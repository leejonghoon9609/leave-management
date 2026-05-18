from http.server import BaseHTTPRequestHandler
import json, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

GREEN='1D9E75'; GREEN_D='0F6E56'; GREEN_L='E1F5EE'
BLUE='378ADD';  BLUE_D='185FA5';  BLUE_L='E6F1FB'
RED='E24B4A';   RED_D='A32D2D';   RED_L='FCEBEB'
PURPLE='7F77DD';PURPLE_D='534AB7';PURPLE_L='EEEDFE'
AMBER_D='854F0B';AMBER_L='FAEEDA'
HEADER='2E7D52';WHITE='FFFFFF';GRAY_L='F3F9F6';NAVY='1E3A5F'

def F(bold=False,sz=11,color='000000'): return Font(bold=bold,size=sz,color=color)
def A(h='center',v='center',wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
def FL(color): return PatternFill('solid',fgColor=color)
def BD(color='CCCCCC'):
    s=Side(style='thin',color=color)
    return Border(top=s,bottom=s,left=s,right=s)

def make_excel(payload):
    data=payload['data']; daily=payload['daily']
    title=payload['title']; month_label=payload['monthLabel']
    wb=Workbook(); wb.remove(wb.active)

    # ── 시트1: 월별현황 ──────────────────────────────────────────
    ws=wb.create_sheet('월별현황')
    BAR=20
    max_v=max((max(r['p'],r['l'],r['a'],r['ot']) for r in data),default=1) or 1

    ws.column_dimensions['A'].width=9
    for i in range(2,BAR+2): ws.column_dimensions[get_column_letter(i)].width=1.6
    ws.column_dimensions[get_column_letter(BAR+2)].width=5
    ws.column_dimensions[get_column_letter(BAR+3)].width=2
    sc=BAR+4
    for i,w in enumerate([8,14,7,9,7,7,7,7,7]):
        ws.column_dimensions[get_column_letter(sc+i)].width=w

    ws.row_dimensions[1].height=20; ws.row_dimensions[2].height=16

    # 제목
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=BAR+2)
    c=ws.cell(1,1,title+' - 그래프'); c.font=F(True,12,HEADER); c.alignment=A('left')
    ws.merge_cells(start_row=1,start_column=sc,end_row=1,end_column=sc+8)
    c=ws.cell(1,sc,title); c.font=F(True,12,HEADER); c.alignment=A('left')

    # 범례
    for i,(lb,col) in enumerate([('■ 정상출근',GREEN),('■ 지각',BLUE),('■ 결근',RED),('■ 특근',PURPLE)]):
        c=ws.cell(2,1+i*5,lb); c.font=F(True,10,col); c.alignment=A('left')
        c.border=Border(bottom=Side(style='thin',color='CCCCCC'))

    # 팀통계 헤더
    for i,h in enumerate(['이름','부서','평일수','정상출근','지각','결근','연차','특근','지각률']):
        c=ws.cell(2,sc+i,h); c.font=F(True,10,WHITE); c.fill=FL(HEADER)
        c.alignment=A('center'); c.border=BD('AAAAAA')

    bar_defs=[('정상출근',GREEN,GREEN_L),('지각',BLUE,BLUE_L),('결근',RED,RED_L),('특근',PURPLE,PURPLE_L)]
    keys=['p','l','a','ot']

    row=3
    for ri,emp in enumerate(data):
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=BAR+2)
        c=ws.cell(row,1,emp['name']); c.font=F(True,12,NAVY); c.alignment=A('left','center')
        ws.row_dimensions[row].height=18

        # 팀통계 이름행
        stat_row=row
        vals=[emp['name'],emp['dept'],emp['wd'],emp['p'],emp['l'],emp['a'],emp['lv'],emp['ot'],emp['lateRate']]
        fcs=[None,None,None,GREEN_D,AMBER_D,RED_D,None,PURPLE_D,None]
        bds=[True,False,False,True,True,True,False,True,False]
        for i,(v,fc,bd) in enumerate(zip(vals,fcs,bds)):
            c=ws.cell(stat_row,sc+i,v)
            c.font=F(bd,9 if i<2 else 10,fc or '000000')
            c.alignment=A('left' if i==0 else 'center'); c.border=BD()
        row+=1

        # 일수 헤더
        ws.row_dimensions[row].height=14
        c=ws.cell(row,1,'일수'); c.font=F(True,9,NAVY); c.alignment=A('center')
        c.border=Border(bottom=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'))
        for d in range(1,BAR+1):
            c=ws.cell(row,1+d,d); c.font=F(False,9); c.alignment=A('center')
            c.border=Border(bottom=Side(style='thin',color='CCCCCC'))
        c=ws.cell(row,BAR+2,'합계'); c.font=F(False,9); c.alignment=A('center')
        c.border=Border(bottom=Side(style='thin',color='CCCCCC'),left=Side(style='thin',color='CCCCCC'))
        row+=1

        for ki,(lb,col,col_l) in enumerate(bar_defs):
            ws.row_dimensions[row].height=16
            val=emp[keys[ki]]; bl=round((val/max_v)*BAR) if max_v>0 else 0
            c=ws.cell(row,1,lb); c.font=F(False,10,col); c.alignment=A('right','center')
            c.border=Border(bottom=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'))
            for bi in range(1,BAR+1):
                c=ws.cell(row,1+bi,''); c.fill=FL(col if bi<=bl else 'FAFAFA')
                c.border=Border(bottom=Side(style='thin',color='CCCCCC'))
            c=ws.cell(row,BAR+2,val); c.font=F(True,11,col); c.alignment=A('center')
            c.border=Border(bottom=Side(style='thin',color='CCCCCC'),left=Side(style='thin',color='CCCCCC'))
            row+=1

        ws.row_dimensions[row].height=8; row+=1

    # 합계행 (팀통계)
    tot={k:sum(r[k] for r in data) for k in ['wd','p','l','a','lv','ot']}
    ta=tot['p']+tot['l']; tl=f"{round(tot['l']/ta*100)}%" if ta>0 else '0%'
    sr=3+(len(data)-1)*6
    for i,v in enumerate(['합계','',tot['wd'],tot['p'],tot['l'],tot['a'],tot['lv'],tot['ot'],tl]):
        c=ws.cell(sr,sc+i,v); c.font=F(True,10,WHITE); c.fill=FL(HEADER)
        c.alignment=A('center'); c.border=BD()

    # ── 시트2: 팀 통계 ──────────────────────────────────────────
    ws2=wb.create_sheet('팀 통계')
    for col,w in zip('ABCDEFGHI',[10,16,8,10,8,8,8,8,8]): ws2.column_dimensions[col].width=w
    ws2.row_dimensions[1].height=22; ws2.row_dimensions[2].height=20
    ws2.merge_cells('A1:I1')
    c=ws2.cell(1,1,title); c.font=F(True,13,HEADER); c.alignment=A('left')
    for i,h in enumerate(['이름','부서','평일수','정상출근','지각','결근','연차','특근','지각률']):
        c=ws2.cell(2,i+1,h); c.font=F(True,11,WHITE); c.fill=FL(HEADER)
        c.alignment=A('center'); c.border=BD('AAAAAA')
    for ri,emp in enumerate(data):
        r=ri+3; ws2.row_dimensions[r].height=18
        for i,(v,fc,bd) in enumerate(zip(
            [emp['name'],emp['dept'],emp['wd'],emp['p'],emp['l'],emp['a'],emp['lv'],emp['ot'],emp['lateRate']],
            [None,None,None,GREEN_D,AMBER_D,RED_D,None,PURPLE_D,None],
            [True,False,False,True,True,True,False,True,False])):
            c=ws2.cell(r,i+1,v); c.font=F(bd,11,fc or '000000')
            c.alignment=A('left' if i==0 else 'center'); c.border=BD()
    sr2=len(data)+3; ws2.row_dimensions[sr2].height=18
    for i,v in enumerate(['합계','',tot['wd'],tot['p'],tot['l'],tot['a'],tot['lv'],tot['ot'],tl]):
        c=ws2.cell(sr2,i+1,v); c.font=F(True,11,WHITE); c.fill=FL(HEADER)
        c.alignment=A('center'); c.border=BD()

    # ── 시트3~N: 직원별 일별 ─────────────────────────────────────
    for emp_d in daily:
        sn=emp_d['name'][:10]+'_일별'; wd=wb.create_sheet(sn)
        for col,w in zip('ABCDEF',[13,9,8,12,11,24]): wd.column_dimensions[col].width=w
        wd.row_dimensions[1].height=22; wd.row_dimensions[2].height=6; wd.row_dimensions[3].height=20
        wd.merge_cells('A1:F1')
        c=wd.cell(1,1,f"{emp_d['name']}  |  부서: {emp_d['dept']}  |  기간: {emp_d['month']}")
        c.font=F(True,13,HEADER); c.alignment=A('left')
        for i,h in enumerate(['날짜','구분','유형','시간','상태','현장메모']):
            c=wd.cell(3,i+1,h); c.font=F(True,11,WHITE); c.fill=FL(HEADER)
            c.alignment=A('center'); c.border=BD('AAAAAA')
        last_date=None; bg_tog=True
        for rd in emp_d['rows']:
            r=wd.max_row+1; wd.row_dimensions[r].height=16
            dv=rd.get('date','')
            if dv and dv!=last_date: last_date=dv; bg_tog=not bg_tog
            bg=GRAY_L if bg_tog else 'FFFFFF'
            gubun=rd.get('gubun',''); status=rd.get('status','')
            if gubun=='현장': gfg,gbg=BLUE_D,BLUE_L
            elif gubun=='특근': gfg,gbg=PURPLE_D,PURPLE_L
            elif gubun=='결근': gfg,gbg=RED_D,RED_L
            elif gubun in('연차','반차','병가','예비군/민방위'): gfg,gbg=GREEN_D,GREEN_L
            else: gfg,gbg=GREEN_D,bg
            if status in('정상출근','정상퇴근'): sfg,sbg=GREEN_D,GREEN_L
            elif status=='지각': sfg,sbg=AMBER_D,AMBER_L
            elif status in('결근','조기퇴근'): sfg,sbg=RED_D,RED_L
            elif status in('특근출근','특근퇴근'): sfg,sbg=PURPLE_D,PURPLE_L
            else: sfg,sbg='888888',bg
            for ci,v in enumerate([dv,gubun,rd.get('type',''),rd.get('time',''),status,rd.get('memo','')]):
                c=wd.cell(r,ci+1,v)
                if ci==1: c.font=F(True,10,gfg); c.fill=FL(gbg); c.alignment=A('center')
                elif ci==4: c.font=F(True,10,sfg); c.fill=FL(sbg); c.alignment=A('center')
                elif ci==5: c.font=F(False,11); c.fill=FL(bg); c.alignment=A('left')
                else: c.font=F(False,11); c.fill=FL(bg); c.alignment=A('center')
                c.border=BD()

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length=int(self.headers.get('Content-Length',0))
            body=self.rfile.read(length)
            payload=json.loads(body)
            excel_bytes=make_excel(payload)
            self.send_response(200)
            self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition','attachment; filename="team_att.xlsx"')
            self.send_header('Content-Length',str(len(excel_bytes)))
            self.send_header('Access-Control-Allow-Origin','*')
            self.end_headers(); self.wfile.write(excel_bytes)
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type','application/json')
            self.send_header('Access-Control-Allow-Origin','*')
            self.end_headers(); self.wfile.write(json.dumps({'error':str(e)}).encode())
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Methods','POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers','Content-Type')
        self.end_headers()

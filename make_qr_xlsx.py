"""薬袋QR入りExcelを作る。使い方:
  python make_qr_xlsx.py            -> サンプルから 薬袋QR.xlsx を新規作成
  python make_qr_xlsx.py 薬袋QR.xlsx -> 「処方」シートを編集後に再実行すると QRシートを作り直す
必要: pip install openpyxl qrcode pillow
"""
import sys, io, os, datetime as dt
import qrcode
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

F = "Arial"; OUT = "薬袋QR.xlsx"
today = dt.date.today()
SAMPLE = {
 "利用者": [("R001","佐藤 花子","101"),("R002","鈴木 一郎","102"),("R003","田中 みどり","103")],
 "薬": [("M001","アムロジピン錠5mg"),("M002","メトホルミン錠250mg"),("M003","ワーファリン錠1mg"),("M004","マグミット錠330mg"),("M005","ガスター錠20mg")],
 "処方": [("P001","R001","M001","1錠","朝"),("P002","R001","M002","1錠","夕"),("P003","R002","M003","1錠","朝"),
          ("P004","R002","M004","2錠","就寝前"),("P005","R003","M005","1錠","昼"),("P006","R003","M001","1錠","夕"),
          ("P007","R003","M002","1錠","朝")],
}
thin = Side(style="thin", color="888888"); box = Border(left=thin,right=thin,top=thin,bottom=thin)
def head(ws, cols):
    for i,c in enumerate(cols,1):
        x=ws.cell(1,i,c); x.font=Font(name=F,bold=True,color="FFFFFF"); x.fill=PatternFill("solid",fgColor="0B4F9C"); x.alignment=Alignment(horizontal="center")
def body(ws):
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.font.name!=F or not c.font.color: c.font=Font(name=F, color=c.font.color.rgb if c.font and c.font.color and isinstance(c.font.color.rgb,str) else "000000")
            c.border=box

def load(path):
    wb = load_workbook(path, data_only=False); d={}
    for name,n in (("利用者",3),("薬",2),("処方",5)):
        d[name]=[tuple(str(v) for v in row[:n]) for row in wb[name].iter_rows(min_row=2,values_only=True) if row[0]]
    return d

def build(d):
    wb = Workbook(); ws = wb.active; ws.title = "使い方"
    for i,t in enumerate([
      "薬袋QR 作成ブック","",
      "1. 「利用者」「薬」「処方」シートの青字セルを編集（処方IDは P001 の形式で一意に）",
      "2. このブックと同じ場所で  python make_qr_xlsx.py 薬袋QR.xlsx  を実行 → 「QR」シートが作り直されます",
      "3. 「QR」シートを印刷し、薬袋に貼る（QRの中身は処方IDだけ。氏名などの個人情報は入っていません）",
      "4. アプリ側のデータ(index.html内のseed)も同じ内容に合わせる",
      "※ 1行目(P001)は入力例です。サンプルのP007は処方期間外のテスト用（アプリ側で期限切れ設定）。"],1):
        ws.cell(i,1,t).font=Font(name=F,bold=(i==1),size=14 if i==1 else 11)
    ws.column_dimensions["A"].width=110
    blue = Font(name=F, color="0000FF")
    w1=wb.create_sheet("利用者"); head(w1,["利用者ID","氏名","部屋"])
    for r in d["利用者"]: w1.append(list(r))
    w2=wb.create_sheet("薬"); head(w2,["薬ID","薬名"])
    for r in d["薬"]: w2.append(list(r))
    w3=wb.create_sheet("処方"); head(w3,["処方ID(QR)","利用者ID","薬ID","用量","時間帯","利用者名","薬名"])
    for i,r in enumerate(d["処方"],2):
        w3.append(list(r))
        w3.cell(i,6,f'=IFERROR(INDEX(利用者!$B:$B,MATCH(B{i},利用者!$A:$A,0)),"※IDなし")')
        w3.cell(i,7,f'=IFERROR(INDEX(薬!$B:$B,MATCH(C{i},薬!$A:$A,0)),"※IDなし")')
    for w in (w1,w2,w3):
        for row in w.iter_rows(min_row=2):
            for c in row:
                isf = isinstance(c.value,str) and c.value.startswith("=")
                c.font = Font(name=F, color="000000" if isf else "0000FF"); c.border=box
        for col in "ABCDEFG": w.column_dimensions[col].width=22
    q=wb.create_sheet("QR"); head(q,["処方ID","利用者名","薬名 / 用量 / 時間帯","QR(印刷用)"])
    q.column_dimensions["A"].width=12; q.column_dimensions["B"].width=18; q.column_dimensions["C"].width=34; q.column_dimensions["D"].width=24
    names={r[0]:r[1] for r in d["利用者"]}; meds={r[0]:r[1] for r in d["薬"]}
    for i,(pid,rid,mid,dose,slot) in enumerate(d["処方"],2):
        q.cell(i,1,f"=処方!A{i}"); q.cell(i,2,f"=処方!F{i}"); q.cell(i,3,f'=処方!G{i}&" / "&処方!D{i}&" / "&処方!E{i}')
        for c in range(1,5):
            q.cell(i,c).font=Font(name=F,size=14,bold=True); q.cell(i,c).border=box
            q.cell(i,c).alignment=Alignment(vertical="center",wrap_text=True)
        img=qrcode.make(pid, box_size=8, border=2); b=io.BytesIO(); img.save(b,"PNG"); b.seek(0)
        im=Image(b); im.width=im.height=120; q.add_image(im,f"D{i}"); q.row_dimensions[i].height=95
    q.page_setup.orientation="portrait"; q.page_setup.fitToWidth=1; q.sheet_properties.pageSetUpPr.fitToPage=True; q.page_setup.fitToHeight=0
    wb.save(OUT); print("saved", OUT)

if __name__=="__main__":
    build(load(sys.argv[1]) if len(sys.argv)>1 and os.path.exists(sys.argv[1]) else SAMPLE)

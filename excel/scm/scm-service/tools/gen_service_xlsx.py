# -*- coding: utf-8 -*-
# SCM ServiceObject / ServiceGroup Excel 生成
import sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

SRC = '/mnt/project-files/scm-excel/SCM_AddressObject.xlsx'
OUT = sys.argv[1]

SE, SI, GE, GI = 'SvcObj_既存一覧', 'SvcObj_登録変更削除', 'SvcGroupObj_既存一覧', 'SvcGroupObj_登録変更削除'
SJ, SD, GD, SL = 'JSON_CURL', 'SvcObj_差分確認', 'SvcGroupObj_差分確認', '実行ログ'
FIRST, LAST, EXL = 8, 507, 5000

FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
F_TITLE, F_HEAD, F_IN, F_FX, F_VBA, F_BTN = fill('1F3864'), fill('2F5597'), fill('FFF2CC'), fill('F2F2F2'), fill('E2EFDA'), fill('C55A11')
F_REQ, F_OPT = fill('C00000'), fill('7F7F7F')

wb = openpyxl.load_workbook(SRC)
for n in ['AdrObj_既存一覧', 'AdrObj_登録変更削除', 'AdrGroupObj_既存一覧', 'AdrGroupObj_登録変更削除', 'AdrObj_差分確認', 'AdrGroupObj_差分確認']:
    del wb[n]
for n in ['EX_FETCHED_AT', 'EX_NAMES', 'EXG_FETCHED_AT', 'EXG_NAMES']:
    del wb.defined_names[n]

def sheet(title, idx):
    ws = wb.create_sheet(title, idx)
    ws.sheet_view.showGridLines = False
    return ws

def put(ws, addr, v, font=None, fl=None, al=None, border=True, nf=None):
    c = ws[addr]
    c.value = v
    c.font = font or Font(name=FN, size=10)
    if fl: c.fill = fl
    if al: c.alignment = al
    if border: c.border = BOX
    if nf: c.number_format = nf
    return c

def header(ws, title, desc, lastcol, h2=44):
    lc = CL(lastcol)
    ws.merge_cells(f'A1:{lc}1'); ws.merge_cells(f'A2:{lc}2')
    put(ws, 'A1', title, Font(name=FN, size=14, bold=True, color='FFFFFF'), F_TITLE, Alignment(vertical='center'), False)
    put(ws, 'A2', desc, Font(name=FN, size=9, color='595959'), None, Alignment(vertical='center', wrap_text=True), False)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = h2

def btn(ws, rng, text):
    ws.merge_cells(rng)
    put(ws, rng.split(':')[0], text, Font(name=FN, size=10, bold=True, color='FFFFFF'), F_BTN,
        Alignment(horizontal='center', vertical='center'), False)

def heads(ws, row, names, widths):
    for i, (n, w) in enumerate(zip(names, widths), 1):
        put(ws, f'{CL(i)}{row}', n, Font(name=FN, size=10, bold=True, color='FFFFFF'), F_HEAD,
            Alignment(horizontal='center', vertical='center', wrap_text=True))
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[row].height = 30

def labels(ws, row, items):
    # items: {col: '必須'/'任意'/'必須※'}
    for col, t in items.items():
        put(ws, f'{col}{row}', t, Font(name=FN, size=9, bold=True, color='FFFFFF'),
            F_REQ if t.startswith('必須') else F_OPT, Alignment(horizontal='center', vertical='center'), False)

def stat(ws, addr, label, formula, fl=F_FX):
    put(ws, addr[0], label, Font(name=FN, size=10, bold=True))
    put(ws, addr[1], formula, Font(name=FN, size=10, bold=True), fl)

def cf(ws, rng, formula, color):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(color), stopIfTrue=False))

Q = lambda s: "'" + s + "'"   # シート名の引用
def ex(sh, col, last=EXL): return f"{Q(sh)}!${col}${FIRST}:${col}${last}"

# ---------------- 共通の数式部品 ----------------
NAMECHK = ('IF(SUMPRODUCT(--ISERROR(FIND(MID($C{r}&REPT(" ",63),ROW($1:$63),1),'
           '" 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")))>0,'
           '"NG:nameに使用不可の文字(英数字 空白 . _ - のみ)",')
RESERVED = ('IF(AND($B{r}="登録",OR($C{r}="any",$C{r}="application-default",$C{r}="service-http",$C{r}="service-https")),'
            '"NG:予約名(any / application-default / service-http / service-https)は使用できません",')

def portchk(s, label):
    R = 'ROW($1:$1023)'
    seg = f'MID({s},{R},12)'
    E = f'LEFT({seg},FIND(",",{seg}&",")-1)'
    A = f'LEFT({E},FIND("-",{E}&"-")-1)'
    B = f'MID({E},FIND("-",{E}&"-")+1,12)'
    start = f'--(MID(","&{s},{R},1)=","),--({R}<=LEN({s}))'
    parts = [
        (f'LENB({s})<>LEN({s})', f'NG:{label}に全角文字が含まれます'),
        (f'LEN({s})>1023', f'NG:{label}が1023文字超'),
        (f'SUMPRODUCT(--ISERROR(FIND(MID({s}&REPT("0",1023),{R},1),"0123456789,-")))>0',
         f'NG:{label}は数字・カンマ・ハイフンのみ(空白不可。例 80,443,8000-8080)'),
        (f'OR(LEFT({s},1)=",",RIGHT({s},1)=",",ISNUMBER(FIND(",,",{s})))',
         f'NG:{label}の先頭・末尾のカンマや連続したカンマを削除してください'),
        (f'SUMPRODUCT(--(LEN({seg})=12),--ISERROR(FIND(",",{seg})))>0',
         f'NG:{label}に形式不正の要素があります(1要素は 80 または 8000-8080)'),
        (f'SUMPRODUCT({start},--(((LEN({E})-LEN(SUBSTITUTE({E},"-","")))>1)+(LEFT({E},1)="-")+(RIGHT({E},1)="-")>0))>0',
         f'NG:{label}の範囲は 開始-終了 の形式(例 8000-8080)'),
        (f'SUMPRODUCT({start},--(((--(0&{A}))>65535)+((--(0&{B}))>65535)>0))>0',
         f'NG:{label}の番号は0〜65535'),
        (f'SUMPRODUCT({start},--ISNUMBER(FIND("-",{E})),--((--(0&{A}))>(--(0&{B}))))>0',
         f'NG:{label}の範囲の開始番号が終了番号より大きい'),
    ]
    f = '""'
    for cond, msg in reversed(parts):
        f = f'IF({cond},"{msg}",{f})'
    return f

def numchk(c, mx, label):
    m = f'"NG:{label}は1〜{mx}の整数"'
    return f'IF({c}="","",IF(ISERROR(--{c}),{m},IF(OR(--{c}<1,--{c}>{mx},--{c}<>INT(--{c})),{m},"")))'

def nest(conds, last):
    f = last
    for cond, val in reversed(conds):
        f = f'IF({cond},{val},{f})'
    return f

# =========================================================
# 基本設定 / JSON_CURL / 実行ログ（Addressツールから流用して文言を変更）
# =========================================================
cfg = wb['基本設定']
cfg['B7'] = '/sse/config/v1/services'
cfg['D7'] = 'ServiceObject（ServiceGroupは /sse/config/v1/service-groups をVBA内で使用）'
cfg['B19'] = None; cfg['B20'] = None; cfg['B21'] = None

js = wb[SJ]
js['A2'] = ('③⑦のJSON発行ボタンでVBAが10行目以降を生成します（緑=VBA書込み）。curlはbash形式。先に5行目でトークン取得→6行目でTOKEN変数を設定してから各行を実行。'
            '「API実行」ボタンでExcelから直接実行も可能（結果は⑤実行ログへ）。JSON発行は③と⑦のどちらのボタンでも両方から一覧を作り直し、'
            'Service登録・変更→Group登録・変更→Group削除→Service削除の順に並べます。削除するService／GroupがほかのGroupのメンバーなら、'
            'Groupから外す変更（変更(Group)自動）を先に追加します。Groupの中にGroupを入れる場合は、中のGroupを先に登録します。')
for row in js.iter_rows(min_row=10, max_row=js.max_row):
    for c in row: c.value = None

lg = wb[SL]
for row in lg.iter_rows(min_row=6, max_row=lg.max_row):
    for c in row:
        c.value = None; c.border = Border(); c.number_format = 'General'

# =========================================================
# ② SvcObj_既存一覧
# =========================================================
ws = sheet(SE, 1)
header(ws, '② 既存情報（SCMから取得した ServiceObject 一覧）',
       'ボタンを押すとAPI（GET）でServiceとServiceGroupの両方を取得し、8行目以降に表を再生成します（B〜L列=VBA書込み、A・M列=関数）。直接編集はしないでください。'
       'timeout系はSCMで上書き（override）設定がある場合のみ値が入ります。', 13)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{EXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
heads(ws, 7, ['No', 'id', 'name', 'protocol', 'port', 'source_port', 'description', 'tag', 'timeout', 'halfclose_timeout',
              'timewait_timeout', 'コンテナ', '定義(関数)'],
      [6, 38, 20, 9, 22, 14, 24, 14, 9, 10, 10, 16, 28])
svc_exist = [
    ('svc-web-8080', 'tcp', '8080', '', 'Web管理画面', 'prod', '', '', ''),
    ('svc-https-alt', 'tcp', '8443', '', 'HTTPS代替ポート', 'prod', '', '', ''),
    ('svc-dns-udp', 'udp', '53', '', 'DNS', '', '', '', ''),
    ('svc-ssh', 'tcp', '22', '', 'SSH', 'mgmt', '', '', ''),
    ('svc-app-range', 'tcp', '9000-9010,9100', '', 'アプリ', 'prod', '7200', '', ''),
    ('svc-old-ftp', 'tcp', '21', '', '旧FTP（廃止予定）', 'mgmt', '', '', ''),
]
for i, v in enumerate(svc_exist):
    r = FIRST + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)')
    put(ws, f'B{r}', f'5c1e2a30-000{i+1}-4d2e-8a11-00000000000{i+1}', nf='@')
    for j, x in enumerate(v):
        put(ws, f'{CL(3+j)}{r}', x if x != '' else None, nf='@')
    put(ws, f'L{r}', 'folder=Shared', nf='@')
    put(ws, f'M{r}', f'=IF(C{r}="","",D{r}&"/"&E{r}&IF(F{r}<>""," src:"&F{r},""))')
ws.freeze_panes = 'D8'

# =========================================================
# ③ SvcObj_登録変更削除
# =========================================================
ws = sheet(SI, 2)
header(ws, '③ 今回の登録・変更・削除一覧（ServiceObject）',
       '黄色(B〜K列)に入力。nameは既存一覧からプルダウン選択可（変更・削除時）。L〜AB列は関数で既存一覧と自動照合し、判定がNGの行はJSON発行対象外。変更時に既存と異なる項目は橙色で表示。'
       'portは「80」「8000-8080」「80,443,8000-8080」の形式（0〜65535、カンマ区切り、空白不可）。source_portも同じ形式で任意。tagは「,」区切り。'
       'timeout系は通常は空欄（既定値）。変更(PUT)は全項目置換のため、残したい値もすべて入力してください。', 28, 56)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(X{FIRST}:X{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(X{FIRST}:X{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(X{FIRST}:X{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行（③⑦）→ ④へ')
btn(ws, 'I3:J4', '▶ 入力クリア')
labels(ws, 6, {'B': '必須', 'C': '必須', 'D': '必須※', 'E': '必須※', 'F': '任意', 'G': '任意', 'H': '任意', 'I': '任意', 'J': '任意', 'K': '任意'})
put(ws, 'L6', '※ protocol・portは登録/変更時に必須（削除時は不要）。任意欄は空欄可（変更時の空欄は既存値が消えます）。halfclose/timewaitはtcpのみ',
    Font(name=FN, size=9, color='595959'), border=False)
heads(ws, 7, ['No', '操作', 'name', 'protocol', 'port', 'source_port', 'description', 'tag', 'timeout\n(1-604800)',
              'halfclose_timeout\n(1-604800,tcp)', 'timewait_timeout\n(1-600,tcp)',
              '既存有無', '既存id', '既存protocol', '既存port', '既存source_port', '既存description', '既存tag',
              '既存timeout', '既存halfclose', '既存timewait', '一覧内重複', '同一定義(別名)', '判定', '理由',
              'portチェック', 'source_portチェック', 'timeoutチェック'],
      [5, 8, 18, 9, 22, 13, 20, 12, 10, 13, 12, 9, 22, 10, 18, 12, 16, 12, 9, 9, 9, 10, 12, 8, 40, 30, 30, 30])
ws.row_dimensions[7].height = 44
EXC = {'M': 'B', 'N': 'D', 'O': 'E', 'P': 'F', 'Q': 'G', 'R': 'H', 'S': 'I', 'T': 'J', 'U': 'K'}
for r in range(FIRST, LAST + 1):
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    for col in 'BCDEFGHIJK':
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@' if col in 'BCDEFGH' else 'General')
    put(ws, f'L{r}', f'=IF($C{r}="","",IF(COUNTIF({ex(SE,"C")},$C{r})>0,"あり","なし"))', fl=F_FX)
    for col, src in EXC.items():
        put(ws, f'{col}{r}', f'=IF($L{r}="あり",INDEX({ex(SE,src)},MATCH($C{r},{ex(SE,"C")},0))&"","")', fl=F_FX)
    put(ws, f'V{r}', f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,"重複",""))', fl=F_FX)
    put(ws, f'W{r}', f'=IF(OR($C{r}="",$B{r}="削除",$D{r}="",$E{r}=""),"",IF(SUMPRODUCT(--({ex(SE,"M")}=$D{r}&"/"&$E{r}&IF($F{r}<>""," src:"&$F{r},"")),--({ex(SE,"C")}<>$C{r}))>0,"同定義あり",""))', fl=F_FX)
    put(ws, f'X{r}', f'=IF($Y{r}="","",IF($Y{r}="OK","OK",LEFT($Y{r},FIND(":",$Y{r})-1)))', fl=F_FX)
    grpcnt = (f'SUMPRODUCT(--ISNUMBER(FIND(","&$C{r}&",",","&{ex(GE,"D")}&",")),--({ex(GE,"D")}<>""))')
    tagcnt = f'IF($H{r}="",0,LEN($H{r})-LEN(SUBSTITUTE($H{r},",",""))+1)'
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$V{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>63', '"NG:nameが63文字超"'),
        (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
    ]
    reason = nest(conds, 'XX')
    rest = [
        (f'AND($B{r}="登録",OR($C{r}="any",$C{r}="application-default",$C{r}="service-http",$C{r}="service-https"))',
         '"NG:予約名(any / application-default / service-http / service-https)は使用できません"'),
        (f'AND($B{r}="登録",$L{r}="あり")', '"NG:既存に同名あり(登録不可)"'),
        (f'AND($B{r}="登録",COUNTIF({ex(GE,"C")},$C{r})>0)', '"NG:同名のServiceGroupが既存にあり(Service・Groupで名前は共通)"'),
        (f'AND($B{r}="登録",COUNTIFS({ex(GI,"B",LAST)},"登録",{ex(GI,"C",LAST)},$C{r})>0)', '"NG:⑦で同名のServiceGroupを登録予定"'),
        (f'AND($B{r}<>"登録",$L{r}="なし")', '"NG:既存エントリ無"'),
        (f'AND($B{r}<>"削除",OR($D{r}="",$E{r}=""))', '"NG:protocolまたはportが未入力"'),
        (f'AND($B{r}<>"削除",$D{r}<>"tcp",$D{r}<>"udp")', '"NG:protocolは tcp または udp"'),
        (f'LEFT($Z{r},2)="NG"', f'$Z{r}'),
        (f'LEFT($AA{r},2)="NG"', f'$AA{r}'),
        (f'LEFT($AB{r},2)="NG"', f'$AB{r}'),
        (f'AND($B{r}<>"削除",LENB($H{r})<>LEN($H{r}))', '"NG:tagに全角文字が含まれます"'),
        (f'AND($B{r}<>"削除",LEN($G{r})>1023)', '"NG:descriptionが1023文字超"'),
        (f'AND($B{r}<>"削除",{tagcnt}>64)', '"NG:tagが64個超"'),
        (f'AND($B{r}="変更",$D{r}&""=$N{r},$E{r}&""=$O{r},$F{r}&""=$P{r},$G{r}&""=$Q{r},$H{r}&""=$R{r},$I{r}&""=$S{r},$J{r}&""=$T{r},$K{r}&""=$U{r})',
         '"警告:既存と差分なし"'),
        (f'AND($B{r}="削除",{grpcnt}>0)',
         f'"警告:Group("&{grpcnt}&"件)のメンバー。JSON発行時にGroupから外す変更を自動で追加します"'),
        (f'$W{r}="同定義あり"', '"警告:同じprotocol/port/source_portの既存Serviceあり(別名)"'),
    ]
    reason = reason.replace('XX', NAMECHK.format(r=r) + nest(rest, '"OK"') + ')')
    put(ws, f'Y{r}', f'=IF($C{r}="","",{reason})', fl=F_FX)
    put(ws, f'Z{r}', f'=IF(OR($C{r}="",$B{r}="削除",$E{r}=""),"",{portchk(f"$E{r}", "port")})', fl=F_FX)
    put(ws, f'AA{r}', f'=IF(OR($C{r}="",$B{r}="削除",$F{r}=""),"",{portchk(f"$F{r}", "source_port")})', fl=F_FX)
    ci, cj, ck = numchk(f'$I{r}', 604800, 'timeout'), numchk(f'$J{r}', 604800, 'halfclose_timeout'), numchk(f'$K{r}', 600, 'timewait_timeout')
    put(ws, f'AB{r}', f'=IF(OR($C{r}="",$B{r}="削除"),"",IF(AND($D{r}="udp",OR($J{r}<>"",$K{r}<>"")),'
                      f'"NG:udpではhalfclose_timeout・timewait_timeoutは指定できません(tcpのみ)",'
                      f'IF({ci}<>"",{ci},IF({cj}<>"",{cj},{ck}))))', fl=F_FX)
ws.freeze_panes = 'D8'
dv = DataValidation(type='list', formula1='"登録,変更,削除"', allow_blank=True); dv.add(f'B{FIRST}:B{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='SVC_NAMES', allow_blank=True, showErrorMessage=False); dv.add(f'C{FIRST}:C{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='"tcp,udp"', allow_blank=True); dv.add(f'D{FIRST}:D{LAST}'); ws.add_data_validation(dv)
cf(ws, f'X{FIRST}:Y{LAST}', f'$X{FIRST}="NG"', 'FFC7CE')
cf(ws, f'X{FIRST}:Y{LAST}', f'$X{FIRST}="警告"', 'FFEB9C')
cf(ws, f'X{FIRST}:Y{LAST}', f'$X{FIRST}="OK"', 'C6EFCE')
for a, b in zip('DEFGHIJK', 'NOPQRSTU'):
    cf(ws, f'{a}{FIRST}:{a}{LAST}', f'AND($B{FIRST}="変更",$L{FIRST}="あり",{a}{FIRST}&""<>${b}{FIRST})', 'F8CBAD')
svc_in = [
    ('登録', 'svc-syslog', 'udp', '514', '', 'syslog転送', 'mgmt', '', '', ''),
    ('登録', 'svc-api', 'tcp', '8000-8100,9443', '', 'API（範囲とカンマ区切り）', 'prod', '', '', ''),
    ('変更', 'svc-ssh', 'tcp', '22,2222', '', 'SSH（代替ポート追加）', 'mgmt', '', '', ''),
    ('削除', 'svc-old-ftp', '', '', '', '', '', '', '', ''),
    ('登録', 'svc-dup-8443', 'tcp', '8443', '', '既存と同じポート', '', '', '', ''),
    ('登録', 'svc-ng-port', 'tcp', '70000', '', '番号が範囲外', '', '', '', ''),
    ('登録', 'svc-ng-range', 'tcp', '9000-8000', '', '開始>終了', '', '', '', ''),
    ('登録', 'svc-ng-zen', 'tcp', '８０', '', '全角数字', '', '', '', ''),
    ('登録', 'svc-ng-udp', 'udp', '161', '', 'udpにhalfclose指定', '', '', '120', ''),
]
for i, v in enumerate(svc_in):
    for j, x in enumerate(v):
        ws[f'{CL(2+j)}{FIRST+i}'] = x if x != '' else None

# =========================================================
# ⑥ SvcGroupObj_既存一覧
# =========================================================
ws = sheet(GE, 3)
header(ws, '⑥ 既存情報（SCMから取得した ServiceGroup 一覧）',
       'ボタンを押すとAPI（GET）でServiceとServiceGroupの両方を取得し、8行目以降に表を再生成します（緑=VBA書込み）。メンバーはService名またはServiceGroup名をカンマ区切りで並べたものです。', 6)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{EXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
heads(ws, 7, ['No', 'id', 'name', 'メンバー', 'tag', 'コンテナ'], [6, 38, 22, 50, 16, 16])
grp_exist = [
    ('grp-web', 'svc-web-8080,svc-https-alt', 'prod'),
    ('grp-mgmt', 'svc-ssh,svc-old-ftp', 'mgmt'),
    ('grp-all', 'grp-web,svc-dns-udp,service-https', ''),
]
for i, v in enumerate(grp_exist):
    r = FIRST + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)')
    put(ws, f'B{r}', f'6d2f3b40-000{i+1}-4d2e-8a11-00000000001{i+1}', nf='@')
    for j, x in enumerate(v):
        put(ws, f'{CL(3+j)}{r}', x if x != '' else None, nf='@')
    put(ws, f'F{r}', 'folder=Shared', nf='@')
ws.freeze_panes = 'D8'

# =========================================================
# ⑦ SvcGroupObj_登録変更削除
# =========================================================
ws = sheet(GI, 4)
header(ws, '⑦ 今回の登録・変更・削除一覧（ServiceGroup）',
       '黄色(B〜E列)に入力。メンバーはService名・ServiceGroup名を「,」区切りかセル内改行（Alt+Enter）で入力します（定義済みの service-http / service-https も指定可）。'
       'F〜S列は関数で②⑥と自動照合し、メンバーが②⑥の既存一覧か③⑦の登録予定に存在するか確認します（大文字・小文字も区別）。判定がNGの行はJSON発行対象外。'
       'ServiceGroupにはdescriptionがありません（API仕様）。変更(PUT)は全項目置換です。', 19, 56)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(L{FIRST}:L{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(L{FIRST}:L{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(L{FIRST}:L{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行（③⑦）→ ④へ')
btn(ws, 'I3:J4', '▶ 入力クリア')
labels(ws, 6, {'B': '必須', 'C': '必須', 'D': '必須※', 'E': '任意'})
put(ws, 'F6', '※ メンバーは登録/変更時に必須（削除時は不要）。tagは「,」区切り（変更時の空欄は既存値が消えます）',
    Font(name=FN, size=9, color='595959'), border=False)
heads(ws, 7, ['No', '操作', 'name', 'メンバー', 'tag', '既存有無', '既存id', '既存メンバー', '既存tag', '一覧内重複',
              'メンバー確認', '判定', '理由', '作業:メンバー(正規化)', '作業:メンバー数', '作業:一致数(延べ)',
              '作業:一致数(種類)', '作業:登録予定数', '作業:削除予定数'],
      [5, 8, 20, 40, 14, 9, 22, 36, 12, 10, 16, 8, 40, 30, 9, 9, 9, 9, 9])
PRE = '{"service-http","service-https"}'
for r in range(FIRST, LAST + 1):
    M2 = f'","&SUBSTITUTE($N{r},",",",,")&","'
    FN_ = f'","&$N{r}&","'
    def occ(rng, conds=''):   # 延べ一致数
        return f'SUMPRODUCT(--({rng}<>""){conds},(LEN({M2})-LEN(SUBSTITUTE({M2},","&{rng}&",","")))/(LEN({rng})+2))'
    def dis(rng, conds=''):   # 種類の一致数
        return f'SUMPRODUCT(--({rng}<>""){conds},--ISNUMBER(FIND(","&{rng}&",",{FN_})))'
    c3 = f',--({ex(SI,"B",LAST)}="登録"),--({ex(SI,"X",LAST)}<>"NG")'
    c4 = (f',--({ex(GI,"B",LAST)}="登録"),--(COUNTIF({ex(GE,"C")},{ex(GI,"C",LAST)})=0)'
          f',--(COUNTIF({ex(SE,"C")},{ex(GI,"C",LAST)})=0)')
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    for col in 'BCDE':
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
    put(ws, f'F{r}', f'=IF($C{r}="","",IF(COUNTIF({ex(GE,"C")},$C{r})>0,"あり","なし"))', fl=F_FX)
    for col, src in {'G': 'B', 'H': 'D', 'I': 'E'}.items():
        put(ws, f'{col}{r}', f'=IF($F{r}="あり",INDEX({ex(GE,src)},MATCH($C{r},{ex(GE,"C")},0))&"","")', fl=F_FX)
    put(ws, f'J{r}', f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,"重複",""))', fl=F_FX)
    put(ws, f'N{r}', f'=IF($D{r}="","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($D{r},CHAR(13),""),CHAR(10),",")," ",""))', fl=F_FX)
    put(ws, f'O{r}', f'=IF($N{r}="",0,LEN($N{r})-LEN(SUBSTITUTE($N{r},",",""))+1)', fl=F_FX)
    put(ws, f'P{r}', f'=IF($N{r}="",0,{occ(ex(SE,"C"))}+{occ(ex(GE,"C"))}+{occ(ex(SI,"C",LAST),c3)}+{occ(ex(GI,"C",LAST),c4)}'
                     f'+SUMPRODUCT((LEN({M2})-LEN(SUBSTITUTE({M2},","&{PRE}&",","")))/(LEN({PRE})+2)))', fl=F_FX)
    put(ws, f'Q{r}', f'=IF($N{r}="",0,{dis(ex(SE,"C"))}+{dis(ex(GE,"C"))}+{dis(ex(SI,"C",LAST),c3)}+{dis(ex(GI,"C",LAST),c4)}'
                     f'+SUMPRODUCT(--ISNUMBER(FIND(","&{PRE}&",",{FN_}))))', fl=F_FX)
    put(ws, f'R{r}', f'=IF($N{r}="",0,{dis(ex(SI,"C",LAST),c3)}+{dis(ex(GI,"C",LAST),c4)})', fl=F_FX)
    put(ws, f'S{r}', f'=IF($N{r}="",0,{dis(ex(SI,"C",LAST),chr(44)+"--("+ex(SI,"B",LAST)+"=\"削除\"),--("+ex(SI,"X",LAST)+"<>\"NG\")")}'
                     f'+{dis(ex(GI,"C",LAST),chr(44)+"--("+ex(GI,"B",LAST)+"=\"削除\")")})', fl=F_FX)
    put(ws, f'K{r}', f'=IF(OR($C{r}="",$B{r}="削除",$N{r}=""),"",'
                     f'IF(OR(LEFT($N{r},1)=",",RIGHT($N{r},1)=",",ISNUMBER(FIND(",,",$N{r}))),"空の要素あり",'
                     f'IF(ISNUMBER(FIND(","&$C{r}&",",","&$N{r}&",")),"自分自身を含む",'
                     f'IF($P{r}<$O{r},"未登録あり("&($O{r}-$P{r})&"件)",'
                     f'IF($P{r}>$Q{r},"メンバー重複",'
                     f'IF($R{r}>0,"登録予定を含む","全て存在"))))))', fl=F_FX)
    put(ws, f'L{r}', f'=IF($M{r}="","",IF($M{r}="OK","OK",LEFT($M{r},FIND(":",$M{r})-1)))', fl=F_FX)
    gcnt = (f'SUMPRODUCT(--ISNUMBER(FIND(","&$C{r}&",",","&{ex(GE,"D")}&",")),--({ex(GE,"D")}<>""))')
    tagcnt = f'IF($E{r}="",0,LEN($E{r})-LEN(SUBSTITUTE($E{r},",",""))+1)'
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$J{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>63', '"NG:nameが63文字超"'),
        (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
    ]
    rest = [
        (f'AND($B{r}="登録",OR($C{r}="any",$C{r}="application-default",$C{r}="service-http",$C{r}="service-https"))',
         '"NG:予約名(any / application-default / service-http / service-https)は使用できません"'),
        (f'AND($B{r}="登録",$F{r}="あり")', '"NG:既存に同名あり(登録不可)"'),
        (f'AND($B{r}="登録",COUNTIF({ex(SE,"C")},$C{r})>0)', '"NG:同名のServiceが既存にあり(Service・Groupで名前は共通)"'),
        (f'AND($B{r}="登録",COUNTIFS({ex(SI,"B",LAST)},"登録",{ex(SI,"C",LAST)},$C{r})>0)', '"NG:③で同名のServiceを登録予定"'),
        (f'AND($B{r}<>"登録",$F{r}="なし")', '"NG:既存エントリ無"'),
        (f'AND($B{r}<>"削除",$D{r}="")', '"NG:メンバーが未入力"'),
        (f'AND($B{r}<>"削除",LENB($D{r})<>LEN($D{r}))', '"NG:メンバーに全角文字が含まれます"'),
        (f'AND($B{r}<>"削除",LENB($E{r})<>LEN($E{r}))', '"NG:tagに全角文字が含まれます"'),
        (f'AND($B{r}<>"削除",{tagcnt}>64)', '"NG:tagが64個超"'),
        (f'$K{r}="空の要素あり"', '"NG:メンバーの先頭・末尾の区切り(カンマ/改行)や連続した区切りを削除してください"'),
        (f'$K{r}="自分自身を含む"', '"NG:メンバーに自分自身は指定できません"'),
        (f'LEFT($K{r},5)="未登録あり"', '"NG:②⑥の既存一覧にも③⑦の登録予定にも無いメンバーが含まれます(大文字・小文字も区別)"'),
        (f'$K{r}="メンバー重複"', '"NG:同じメンバーが重複しています"'),
        (f'AND($B{r}<>"削除",$S{r}>0)', '"警告:③⑦で削除予定のメンバーを含む(JSON発行時にメンバーから外します)"'),
        (f'$K{r}="登録予定を含む"', '"警告:③⑦で登録予定のメンバーを含む(④で先に登録します)"'),
        (f'AND($B{r}="変更",$N{r}=$H{r},$E{r}&""=$I{r})', '"警告:既存と差分なし"'),
        (f'AND($B{r}="削除",{gcnt}>0)',
         f'"警告:他のGroup("&{gcnt}&"件)のメンバー。JSON発行時にGroupから外す変更を自動で追加します"'),
    ]
    reason = nest(conds, NAMECHK.format(r=r) + nest(rest, '"OK"') + ')')
    put(ws, f'M{r}', f'=IF($C{r}="","",{reason})', fl=F_FX)
ws.freeze_panes = 'D8'
dv = DataValidation(type='list', formula1='"登録,変更,削除"', allow_blank=True); dv.add(f'B{FIRST}:B{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='SVCG_NAMES', allow_blank=True, showErrorMessage=False); dv.add(f'C{FIRST}:C{LAST}'); ws.add_data_validation(dv)
cf(ws, f'L{FIRST}:M{LAST}', f'$L{FIRST}="NG"', 'FFC7CE')
cf(ws, f'L{FIRST}:M{LAST}', f'$L{FIRST}="警告"', 'FFEB9C')
cf(ws, f'L{FIRST}:M{LAST}', f'$L{FIRST}="OK"', 'C6EFCE')
cf(ws, f'D{FIRST}:D{LAST}', f'AND($B{FIRST}="変更",$F{FIRST}="あり",$N{FIRST}<>$H{FIRST})', 'F8CBAD')
cf(ws, f'E{FIRST}:E{LAST}', f'AND($B{FIRST}="変更",$F{FIRST}="あり",E{FIRST}&""<>$I{FIRST})', 'F8CBAD')
for r in range(FIRST, LAST + 1):
    ws[f'D{r}'].alignment = Alignment(wrap_text=True, vertical='top')
grp_in = [
    ('登録', 'grp-api', 'svc-api\nsvc-syslog', 'prod'),
    ('変更', 'grp-all', 'grp-web,svc-dns-udp,service-https,grp-api', ''),
    ('登録', 'grp-ng-case', 'svc-Web-8080', ''),
    ('登録', 'grp-ng-self', 'grp-ng-self,svc-ssh', ''),
    ('登録', 'grp-ng-dup', 'svc-ssh,svc-ssh', ''),
    ('登録', 'grp-ng-sep', 'svc-ssh,,svc-dns-udp', ''),
]
for i, v in enumerate(grp_in):
    for j, x in enumerate(v):
        ws[f'{CL(2+j)}{FIRST+i}'] = x if x != '' else None

# =========================================================
# ⑧⑨ 差分確認
# =========================================================
def diffsheet(title, idx, heading, fields, widths, rows, btnrng):
    ws = sheet(title, idx)
    ncol = 3 + len(fields) + 5
    header(ws, heading,
           '④のAPI実行が終わると、VBAがSCMに今設定されている全件を取得して一覧表示し、既存一覧シート（実行前）と比べた変化をC列に付けます。'
           '行の背景色で変化を表します（追加=緑、変更=黄、変化なし=白、削除=グレー）。削除されたものは一覧の後ろに、既存一覧（実行前）の内容で表示します。'
           '判定が要確認のもの（判定列が赤）はPush前に理由を確認してください。ボタンでいつでも作り直せます。', ncol)
    put(ws, 'A3', '作成日時', Font(name=FN, size=10, bold=True))
    put(ws, 'B3', None, fl=F_VBA, nf='yyyy/mm/dd hh:mm:ss')
    for lab, val in [(('C3', 'D3'), '全件'), (('E3', 'F3'), '変化なし'), (('G3', 'H3'), '追加'),
                     (('C4', 'D4'), '変更'), (('E4', 'F4'), '削除'), (('G4', 'H4'), '要確認')]:
        stat(ws, lab, val, None, F_VBA)
    btn(ws, btnrng, '▶ 差分チェック（再作成）')
    put(ws, 'C5', '凡例：', Font(name=FN, size=10, bold=True), border=False)
    for a, t, c in [('D5', '追加', 'C6EFCE'), ('E5', '変更', 'FFFF99'), ('F5', '変化なし', 'FFFFFF'), ('G5', '削除', 'D9D9D9'), ('H5', '要確認(判定列)', 'FFC7CE')]:
        put(ws, a, t, fl=fill(c), al=Alignment(horizontal='center'))
    names = ['No', 'name', '変化'] + [f + '（現在）' for f in fields] + ['変更項目', '変更前（既存一覧）', '④の操作', 'HTTP', '判定']
    heads(ws, 7, names, [5, 20, 9] + widths + [16, 34, 9, 7, 26])
    jc = CL(ncol)
    cf(ws, f'{jc}{FIRST}:{jc}{EXL}', f'AND(${jc}{FIRST}<>"",${jc}{FIRST}<>"想定どおり",LEFT(${jc}{FIRST},4)<>"変化なし")', 'FFC7CE')
    cf(ws, f'A{FIRST}:{jc}{EXL}', f'$C{FIRST}="追加"', 'C6EFCE')
    cf(ws, f'A{FIRST}:{jc}{EXL}', f'$C{FIRST}="変更"', 'FFFF99')
    cf(ws, f'A{FIRST}:{jc}{EXL}', f'$C{FIRST}="削除"', 'D9D9D9')
    stats = {'': 0, '追加': 0, '変更': 0, '削除': 0, 'chk': 0}
    for i, v in enumerate(rows):
        r = FIRST + i
        vals = [i + 1] + list(v)
        for j, x in enumerate(vals):
            put(ws, f'{CL(1+j)}{r}', x if x != '' else None, al=Alignment(wrap_text=True, vertical='top'))
        ch = v[1]
        stats['' if ch == '変化なし' else ch] += 1
        if v[-1] not in ('', '想定どおり') and not v[-1].startswith('変化なし'): stats['chk'] += 1
    ws['D3'] = len(rows) - stats['削除']; ws['F3'] = stats['']; ws['H3'] = stats['追加']
    ws['D4'] = stats['変更']; ws['F4'] = stats['削除']; ws['H4'] = stats['chk']
    ws['B3'] = datetime.datetime(2026, 10, 26, 13, 6)
    ws.freeze_panes = 'D8'

diffsheet(SD, 6, '⑧ 差分確認（ServiceObject）：現在の設定一覧と既存一覧の比較',
          ['protocol', 'port', 'source_port', 'description', 'tag', 'timeout'], [9, 22, 12, 20, 12, 22], [
    ('svc-web-8080', '変化なし', 'tcp', '8080', '', 'Web管理画面', 'prod', '', '', '', '', '', ''),
    ('svc-https-alt', '変化なし', 'tcp', '8443', '', 'HTTPS代替ポート', 'prod', '', '', '', '', '', ''),
    ('svc-dns-udp', '変化なし', 'udp', '53', '', 'DNS', '', '', '', '', '', '', ''),
    ('svc-ssh', '変更', 'tcp', '22,2222', '', 'SSH（代替ポート追加）', 'mgmt', '', 'port description', 'protocol: tcp\nport: 22\ndescription: SSH\ntag: mgmt', '変更', '200', '想定どおり'),
    ('svc-app-range', '変化なし', 'tcp', '9000-9010,9100', '', 'アプリ', 'prod', 'timeout=7200', '', '', '', '', ''),
    ('svc-syslog', '追加', 'udp', '514', '', 'syslog転送', 'mgmt', '', '', '', '登録', '201', '想定どおり'),
    ('svc-api', '追加', 'tcp', '8000-8100,9443', '', 'API（範囲とカンマ区切り）', 'prod', '', '', '', '登録', '201', '想定どおり'),
    ('svc-dup-8443', '追加', 'tcp', '8443', '', '既存と同じポート', '', '', '', '', '登録', '201', '想定どおり'),
    ('svc-old-ftp', '削除', 'tcp', '21', '', '旧FTP（廃止予定）', 'mgmt', '', '', '', '削除', '200', '想定どおり'),
], 'J3:K4')
diffsheet(GD, 7, '⑨ 差分確認（ServiceGroup）：現在の設定一覧と既存一覧の比較',
          ['メンバー', 'tag'], [44, 14], [
    ('grp-web', '変化なし', 'svc-web-8080,svc-https-alt', 'prod', '', '', '', '', ''),
    ('grp-mgmt', '変更', 'svc-ssh', 'mgmt', 'メンバー', 'メンバー: svc-ssh,svc-old-ftp\ntag: mgmt', '変更', '200', '想定どおり'),
    ('grp-all', '変更', 'grp-web,svc-dns-udp,service-https,grp-api', '', 'メンバー', 'メンバー: grp-web,svc-dns-udp,service-https', '変更', '200', '想定どおり'),
    ('grp-api', '追加', 'svc-api,svc-syslog', 'prod', '', '', '登録', '201', '想定どおり'),
], 'J3:K4')


# ④ JSON_CURL サンプル（VBAのJSON発行と同じ形式）
import json as _j
BASE = 'https://api.sase.paloaltonetworks.com'
def jstr(v): return _j.dumps(v, ensure_ascii=False)
def svcbody(nm, pr, port, desc, tag):
    b = '{"name":' + jstr(nm) + ',"protocol":{"' + pr + '":{"port":' + jstr(port) + '}}'
    if desc: b += ',"description":' + jstr(desc)
    if tag: b += ',"tag":[' + ','.join(jstr(t) for t in tag.split(',')) + ']'
    return b + '}'
def grpbody(nm, mem, tag):
    b = '{"name":' + jstr(nm) + ',"members":[' + ','.join(jstr(m) for m in mem.split(',')) + ']'
    if tag: b += ',"tag":[' + ','.join(jstr(t) for t in tag.split(',')) + ']'
    return b + '}'
def curl(m, u, b):
    c = f"curl -s -X {m} '{u}' -H \"Authorization: Bearer $TOKEN\""
    if b: c += " -H 'Content-Type: application/json' -d '" + b.replace("'", "'\\''") + "'"
    return c
S, G = BASE + '/sse/config/v1/services', BASE + '/sse/config/v1/service-groups'
rows = [
    ('登録', 'svc-syslog', 'POST', S + '?folder=Shared', svcbody('svc-syslog', 'udp', '514', 'syslog転送', 'mgmt'), 201),
    ('登録', 'svc-api', 'POST', S + '?folder=Shared', svcbody('svc-api', 'tcp', '8000-8100,9443', 'API（範囲とカンマ区切り）', 'prod'), 201),
    ('変更', 'svc-ssh', 'PUT', S + '/5c1e2a30-0004-4d2e-8a11-000000000004', svcbody('svc-ssh', 'tcp', '22,2222', 'SSH（代替ポート追加）', 'mgmt'), 200),
    ('登録', 'svc-dup-8443', 'POST', S + '?folder=Shared', svcbody('svc-dup-8443', 'tcp', '8443', '既存と同じポート', ''), 201),
    ('登録(Group)', 'grp-api', 'POST', G + '?folder=Shared', grpbody('grp-api', 'svc-api,svc-syslog', 'prod'), 201),
    ('変更(Group)', 'grp-all', 'PUT', G + '/6d2f3b40-0003-4d2e-8a11-000000000013', grpbody('grp-all', 'grp-web,svc-dns-udp,service-https,grp-api', ''), 200),
    ('変更(Group)自動', 'grp-mgmt', 'PUT', G + '/6d2f3b40-0002-4d2e-8a11-000000000012', grpbody('grp-mgmt', 'svc-ssh', 'mgmt'), 200),
    ('削除', 'svc-old-ftp', 'DELETE', S + '/5c1e2a30-0006-4d2e-8a11-000000000006', '', 200),
]
js = wb[SJ]
for i, (op, nm, m, u, b, st) in enumerate(rows):
    r = 10 + i
    for j, v in enumerate([i + 1, op, nm, m, u, b or None, curl(m, u, b), st]):
        put(js, f'{CL(1+j)}{r}', v, al=Alignment(wrap_text=True, vertical='top'))

# 名前定義
wb.defined_names['EX_FETCHED_AT'] = DefinedName('EX_FETCHED_AT', attr_text=f"{Q(SE)}!$B$3")
wb.defined_names['EXG_FETCHED_AT'] = DefinedName('EXG_FETCHED_AT', attr_text=f"{Q(GE)}!$B$3")
wb.defined_names['SVC_NAMES'] = DefinedName('SVC_NAMES', attr_text=f"OFFSET({Q(SE)}!$C$8,0,0,MAX(1,COUNTA({Q(SE)}!$C$8:$C$5000)),1)")
wb.defined_names['SVCG_NAMES'] = DefinedName('SVCG_NAMES', attr_text=f"OFFSET({Q(GE)}!$C$8,0,0,MAX(1,COUNTA({Q(GE)}!$C$8:$C$5000)),1)")
wb[SE]['B3'] = datetime.datetime(2026, 10, 26, 13, 1)
wb[GE]['B3'] = datetime.datetime(2026, 10, 26, 13, 1)

order = ['基本設定', SE, SI, GE, GI, SJ, SD, GD, SL]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for w in wb: w.sheet_view.tabSelected = (w.title == '基本設定')
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('saved', OUT)

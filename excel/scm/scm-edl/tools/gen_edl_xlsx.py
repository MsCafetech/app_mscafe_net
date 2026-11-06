# -*- coding: utf-8 -*-
# SCM External Dynamic List (EDL) Excel 生成（ServiceObjectツールのブックから 基本設定/JSON_CURL/実行ログ を流用）
import sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

SRC = "/mnt/project-files/scm-service/SCM_ServiceObject.xlsx"
OUT = sys.argv[1]

SE, SI = "EDL_既存一覧", "EDL_登録変更削除"
SJ, SD, SL = "JSON_CURL", "EDL_差分確認", "実行ログ"
FIRST, LAST, EXL = 8, 507, 5000

FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
F_TITLE, F_HEAD, F_IN, F_FX, F_VBA, F_BTN = fill('1F3864'), fill('2F5597'), fill('FFF2CC'), fill('F2F2F2'), fill('E2EFDA'), fill('C55A11')
F_REQ, F_OPT = fill('C00000'), fill('7F7F7F')


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

wb = openpyxl.load_workbook(SRC)
for n in ['SvcObj_既存一覧', 'SvcObj_登録変更削除', 'SvcGroupObj_既存一覧', 'SvcGroupObj_登録変更削除', 'SvcObj_差分確認', 'SvcGroupObj_差分確認']:
    del wb[n]
for n in ['EX_FETCHED_AT', 'EXG_FETCHED_AT', 'SVC_NAMES', 'SVCG_NAMES']:
    del wb.defined_names[n]

def sheet(title, idx):
    ws = wb.create_sheet(title, idx)
    ws.sheet_view.showGridLines = False
    return ws

def nest(conds, last):
    f = last
    for cond, val in reversed(conds):
        f = f'IF({cond},{val},{f})'
    return f

NAMECHK = ('IF(SUMPRODUCT(--ISERROR(FIND(MID($C{r}&REPT(" ",63),ROW($1:$63),1),'
           '" 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")))>0,'
           '"NG:nameに使用不可の文字(英数字 空白 . _ - のみ)",')

TYPES = ['ip', 'domain', 'url', 'imsi', 'imei', 'predefined_ip', 'predefined_url']
RECS = ['five_minute', 'hourly', 'daily', 'weekly', 'monthly']
DOWS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday']
HOURS = [f'{h:02d}' for h in range(24)]
PREDEF = ['panw-bulletproof-ip-list', 'panw-highrisk-ip-list', 'panw-known-ip-list', 'panw-torexit-ip-list',
          'panw-auth-portal-exclude-list']
arr = lambda xs: '{' + ','.join(f'"{x}"' for x in xs) + '}'

# =========================================================
# 基本設定 / JSON_CURL / 実行ログ（ServiceObjectツールから流用して文言を変更）
# =========================================================
cfg = wb['基本設定']
cfg['B7'] = '/sse/config/v1/external-dynamic-lists'
cfg['D7'] = 'External Dynamic List（証明書プロファイルの候補取得に /sse/config/v1/certificate-profiles をVBA内で使用）'
cfg['B19'] = None; cfg['B20'] = None; cfg['B21'] = None

js = wb[SJ]
js['A2'] = ('③のJSON発行ボタンでVBAが10行目以降を生成します（緑=VBA書込み）。curlはbash形式。先に5行目でトークン取得→6行目でTOKEN変数を設定してから各行を実行。'
            '「API実行」ボタンでExcelから直接実行も可能（結果は⑤実行ログへ）。登録・変更（③の行順）→削除（③の行順）の順に並べます。'
            'auth（username/password）を指定した行はJSONとcurlにpasswordがそのまま入るため、共有時は注意してください。')
for row in js.iter_rows(min_row=10, max_row=js.max_row):
    for c in row: c.value = None

lg = wb[SL]
for row in lg.iter_rows(min_row=6, max_row=lg.max_row):
    for c in row:
        c.value = None; c.border = Border(); c.number_format = 'General'

# =========================================================
# ② EDL_既存一覧
# =========================================================
ws = sheet(SE, 1)
header(ws, '② 既存情報（SCMから取得した External Dynamic List 一覧）',
       'ボタンを押すとAPI（GET）でEDLを取得し、8行目以降に表を再生成します（B〜O列=VBA書込み、A列=関数）。直接編集はしないでください。定義済みのEDLは含めません。'
       'certificate_profile の None は空欄、expand_domain は有効なときだけ true と表示します。passwordはSCMから取得できないため表示しません。'
       'Q列は証明書プロファイル（③のプルダウン候補、VBA書込み）、S列は定義済みリストの名前（③で type=predefined_ip / predefined_url のときのurl候補）です。', 19, 56)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{EXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
heads(ws, 7, ['No', 'id', 'name', 'type', 'url（取得元）', 'recurring\n(更新間隔)', 'at\n(時)', 'day_of_week', 'day_of_month',
              'description', 'exception_list', 'certificate_profile', 'expand_domain', 'username', 'コンテナ', '',
              '（候補）\n証明書プロファイル', '', '（候補）\n定義済みリスト'],
      [6, 38, 22, 14, 42, 12, 6, 12, 8, 24, 30, 18, 9, 14, 16, 2, 22, 2, 30])
ws['P7'].fill = PatternFill(); ws['P7'].border = Border(); ws['R7'].fill = PatternFill(); ws['R7'].border = Border()
edl_exist = [
    ('edl-block-ip', 'ip', 'https://feeds.example.com/block-ip.txt', 'hourly', '', '', '', '社内ブロックIP', '', '', '', ''),
    ('edl-block-domain', 'domain', 'https://feeds.example.com/block-domain.txt', 'daily', '03', '', '', 'ブロックドメイン', 'test.example.com', '', 'true', ''),
    ('edl-allow-url', 'url', 'https://feeds.example.com/allow-url.txt', 'weekly', '02', 'sunday', '', '許可URL', '', 'edl-cert-profile', '', 'feeduser'),
    ('edl-highrisk', 'predefined_ip', 'panw-highrisk-ip-list', '', '', '', '', '高リスクIP（例外あり）', '203.0.113.10', '', '', ''),
    ('edl-old-list', 'ip', 'http://old.example.com/list.txt', 'five_minute', '', '', '', '旧リスト（廃止予定）', '', '', '', ''),
]
for i, v in enumerate(edl_exist):
    r = FIRST + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)')
    put(ws, f'B{r}', f'7e3a4c50-000{i+1}-4d2e-8a11-00000000002{i+1}', nf='@')
    for j, x in enumerate(v):
        put(ws, f'{CL(3+j)}{r}', x if x != '' else None, nf='@', al=Alignment(vertical='top', wrap_text=(j == 8)))
    put(ws, f'O{r}', 'folder=Shared', nf='@')
for i, x in enumerate(['edl-cert-profile', 'client-cert-profile']):
    put(ws, f'Q{FIRST+i}', x, nf='@')
for i, x in enumerate(PREDEF):
    put(ws, f'S{FIRST+i}', x, nf='@')
ws.freeze_panes = 'D8'

# =========================================================
# ③ EDL_登録変更削除
# =========================================================
ws = sheet(SI, 2)
header(ws, '③ 今回の登録・変更・削除一覧（External Dynamic List）',
       '黄色(B〜O列)に入力。type・recurring・at・曜日・日・certificate_profile・expand_domain はプルダウンで選択、nameは既存一覧から選択可（変更・削除時）。'
       'P〜AH列は関数で既存一覧と自動照合し、判定がNGの行はJSON発行対象外。変更時に既存と異なる項目は橙色で表示。'
       'urlは http:// または https:// で始まる取得元URL（type=predefined_ip / predefined_url のときは定義済みリスト名をプルダウンから選択）。'
       'exception_list は「,」区切りかセル内改行（Alt+Enter）。変更(PUT)は全項目置換のため、残したい値もすべて入力してください（usernameを使う場合はpasswordも毎回必要）。', 34, 70)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(AE{FIRST}:AE{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(AE{FIRST}:AE{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(AE{FIRST}:AE{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行 → ④へ')
btn(ws, 'I3:J4', '▶ 入力クリア')
labels(ws, 6, {'B': '必須', 'C': '必須', 'D': '必須※', 'E': '必須※', 'F': '必須※', 'G': '条件', 'H': '条件', 'I': '条件',
               'J': '任意', 'K': '任意', 'L': '任意', 'M': '任意', 'N': '任意', 'O': '条件'})
put(ws, 'P6', '※ type・url は登録/変更時に必須、recurring は predefined 以外で必須。at=daily/weekly/monthly、曜日=weekly、日=monthly のとき必須。'
              'password は username を入力したとき必須。空欄の任意項目は送信しません（certificate_profile 空欄=None、expand_domain 空欄=false）',
    Font(name=FN, size=9, color='595959'), border=False)
heads(ws, 7, ['No', '操作', 'name', 'type', 'url（取得元）', 'recurring\n(更新間隔)', 'at\n(時 00-23)', 'day_of_week\n(weekly)',
              'day_of_month\n(monthly)', 'description', 'exception_list\n(「,」か改行区切り)', 'certificate_profile\n(空欄=None)',
              'expand_domain\n(domainのみ)', 'username', 'password',
              '既存有無', '既存id', '既存type', '既存url', '既存recurring', '既存at', '既存day_of_week', '既存day_of_month',
              '既存description', '既存exception_list', '既存certificate_profile', '既存expand_domain', '既存username',
              '一覧内重複', '同一URL(別名)', '判定', '理由', '作業:exception(正規化)', '作業:exception件数'],
      [5, 8, 20, 14, 40, 12, 9, 12, 10, 20, 30, 18, 11, 12, 12,
       9, 22, 12, 30, 11, 6, 10, 8, 16, 24, 16, 9, 12, 10, 12, 8, 46, 26, 9])
ws.row_dimensions[7].height = 44
EXC = {'Q': 'B', 'R': 'D', 'S': 'E', 'T': 'F', 'U': 'G', 'V': 'H', 'W': 'I', 'X': 'J', 'Y': 'K', 'Z': 'L', 'AA': 'M', 'AB': 'N'}
for r in range(FIRST, LAST + 1):
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    for col in 'BCDEFGHIJKLMNO':
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@',
            al=Alignment(vertical='top', wrap_text=(col == 'K')))
    put(ws, f'P{r}', f'=IF($C{r}="","",IF(COUNTIF({ex(SE,"C")},$C{r})>0,"あり","なし"))', fl=F_FX)
    for col, src in EXC.items():
        put(ws, f'{col}{r}', f'=IF($P{r}="あり",INDEX({ex(SE,src)},MATCH($C{r},{ex(SE,"C")},0))&"","")', fl=F_FX)
    put(ws, f'AC{r}', f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,"重複",""))', fl=F_FX)
    put(ws, f'AD{r}', f'=IF(OR($C{r}="",$B{r}="削除",$E{r}=""),"",IF(SUMPRODUCT(--({ex(SE,"E")}=$E{r}),--({ex(SE,"D")}=$D{r}),--({ex(SE,"C")}<>$C{r}))>0,"同URLあり",""))', fl=F_FX)
    put(ws, f'AE{r}', f'=IF($AF{r}="","",IF($AF{r}="OK","OK",LEFT($AF{r},FIND(":",$AF{r})-1)))', fl=F_FX)
    put(ws, f'AG{r}', f'=IF($K{r}="","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($K{r},CHAR(13),""),CHAR(10),",")," ",""))', fl=F_FX)
    put(ws, f'AH{r}', f'=IF($AG{r}="",0,LEN($AG{r})-LEN(SUBSTITUTE($AG{r},",",""))+1)', fl=F_FX)
    act = f'$B{r}<>"削除"'
    pre = f'LEFT($D{r},11)="predefined_"'
    npre = f'LEFT($D{r},11)<>"predefined_"'
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$AC{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>63', '"NG:nameが63文字超"'),
        (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
    ]
    rest = [
        (f'AND($B{r}="登録",$P{r}="あり")', '"NG:既存に同名あり(登録不可)"'),
        (f'AND($B{r}<>"登録",$P{r}="なし")', '"NG:既存エントリ無"'),
        (f'AND({act},$D{r}="")', '"NG:typeが未選択"'),
        (f'AND({act},ISNA(MATCH($D{r},{arr(TYPES)},0)))', '"NG:typeは一覧から選択(ip / domain / url / imsi / imei / predefined_ip / predefined_url)"'),
        (f'AND({act},$E{r}="")', '"NG:urlが未入力"'),
        (f'AND({act},LENB($E{r})<>LEN($E{r}))', '"NG:urlに全角文字が含まれます"'),
        (f'AND({act},LEN($E{r})>255)', '"NG:urlが255文字超"'),
        (f'AND({act},{npre},LEFT(LOWER($E{r}),7)<>"http://",LEFT(LOWER($E{r}),8)<>"https://")',
         '"NG:urlは http:// または https:// で始まる取得元URLを入力"'),
        (f'AND({act},{pre},($F{r}&$G{r}&$H{r}&$I{r}&$L{r}&$M{r}&$N{r}&$O{r})<>"")',
         '"NG:predefinedではrecurring・at・曜日・日・certificate_profile・expand_domain・username・passwordは指定できません(空欄にしてください)"'),
        (f'AND({act},{npre},$F{r}="")', '"NG:recurring(更新間隔)が未選択"'),
        (f'AND({act},{npre},ISNA(MATCH($F{r},{arr(RECS)},0)))', '"NG:recurringは five_minute / hourly / daily / weekly / monthly から選択"'),
        (f'AND({act},OR($F{r}="daily",$F{r}="weekly",$F{r}="monthly"),$G{r}="")', '"NG:daily・weekly・monthlyではat(時)が必須"'),
        (f'AND({act},$G{r}<>"",ISNA(MATCH($G{r},{arr(HOURS)},0)))', '"NG:atは00〜23の2桁(プルダウンから選択)"'),
        (f'AND({act},$F{r}="weekly",$H{r}="")', '"NG:weeklyではday_of_week(曜日)が必須"'),
        (f'AND({act},$H{r}<>"",ISNA(MATCH($H{r},{arr(DOWS)},0)))', '"NG:day_of_weekはプルダウンから選択(sunday〜saturday)"'),
        (f'AND({act},$F{r}="monthly",$I{r}="")', '"NG:monthlyではday_of_month(日)が必須"'),
        (f'AND({act},$I{r}<>"",ISNA(MATCH($I{r},{arr([str(d) for d in range(1, 32)])},0)))', '"NG:day_of_monthは1〜31(プルダウンから選択)"'),
        (f'AND({act},{npre},OR(AND($G{r}<>"",OR($F{r}="five_minute",$F{r}="hourly")),AND($H{r}<>"",$F{r}<>"weekly"),AND($I{r}<>"",$F{r}<>"monthly")))',
         '"NG:選択したrecurringでは不要なat・曜日・日が入力されています(空欄にしてください)"'),
        (f'AND({act},$M{r}<>"",$M{r}<>"true")', '"NG:expand_domainは true か空欄"'),
        (f'AND({act},$M{r}<>"",$D{r}<>"domain")', '"NG:expand_domainは type=domain のときのみ指定可"'),
        (f'AND({act},$N{r}="",$O{r}<>"")', '"NG:passwordを使う場合はusernameも入力"'),
        (f'AND({act},$N{r}<>"",$O{r}="")', '"NG:usernameを使う場合はpasswordも入力(変更は全項目置換。既存のpasswordはSCMから取得できません)"'),
        (f'AND({act},OR(LEN($N{r})>255,LEN($O{r})>255))', '"NG:username・passwordは255文字以内"'),
        (f'AND({act},LEN($J{r})>255)', '"NG:descriptionが255文字超"'),
        (f'AND({act},LENB($K{r})<>LEN($K{r}))', '"NG:exception_listに全角文字が含まれます"'),
        (f'AND({act},$AG{r}<>"",OR(LEFT($AG{r},1)=",",RIGHT($AG{r},1)=",",ISNUMBER(FIND(",,",$AG{r}))))',
         '"NG:exception_listの先頭・末尾の区切り(カンマ/改行)や連続した区切りを削除してください"'),
        (f'AND({act},$AH{r}>100)', '"警告:exception_listが100件超(SCMで受け付けられない場合があります)"'),
        (f'AND({act},{pre},COUNTIF({ex(SE,"S",FIRST+19)},$E{r})=0)', '"警告:urlが定義済みリストの候補にありません(SCMの名前を確認してください)"'),
        (f'AND($B{r}="変更",$P{r}="あり",$D{r}<>$R{r})', '"警告:typeを変更します(SCMで受け付けられない場合は削除→登録してください)"'),
        (f'AND($B{r}="変更",$D{r}&""=$R{r},$E{r}&""=$S{r},$F{r}&""=$T{r},$G{r}&""=$U{r},$H{r}&""=$V{r},$I{r}&""=$W{r},'
         f'$J{r}&""=$X{r},$AG{r}&""=$Y{r},$L{r}&""=$Z{r},$M{r}&""=$AA{r},$N{r}&""=$AB{r},$O{r}="")',
         '"警告:既存と差分なし"'),
        (f'$AD{r}="同URLあり"', '"警告:同じtype・urlの既存EDLあり(別名)"'),
    ]
    reason = nest(conds, NAMECHK.format(r=r) + nest(rest, '"OK"') + ')')
    put(ws, f'AF{r}', f'=IF($C{r}="","",{reason})', fl=F_FX)
ws.freeze_panes = 'D8'
def dvl(f, rng, err=True):
    dv = DataValidation(type='list', formula1=f, allow_blank=True, showErrorMessage=err)
    dv.add(rng); ws.add_data_validation(dv)
dvl('"登録,変更,削除"', f'B{FIRST}:B{LAST}')
dvl('EDL_NAMES', f'C{FIRST}:C{LAST}', False)
dvl('"' + ','.join(TYPES) + '"', f'D{FIRST}:D{LAST}')
dvl('PRE_NAMES', f'E{FIRST}:E{LAST}', False)
dvl('"' + ','.join(RECS) + '"', f'F{FIRST}:F{LAST}')
dvl('"' + ','.join(HOURS) + '"', f'G{FIRST}:G{LAST}')
dvl('"' + ','.join(DOWS) + '"', f'H{FIRST}:H{LAST}')
dvl('"' + ','.join(str(d) for d in range(1, 32)) + '"', f'I{FIRST}:I{LAST}')
dvl('CERT_NAMES', f'L{FIRST}:L{LAST}', False)
dvl('"true"', f'M{FIRST}:M{LAST}')
cf(ws, f'AE{FIRST}:AF{LAST}', f'$AE{FIRST}="NG"', 'FFC7CE')
cf(ws, f'AE{FIRST}:AF{LAST}', f'$AE{FIRST}="警告"', 'FFEB9C')
cf(ws, f'AE{FIRST}:AF{LAST}', f'$AE{FIRST}="OK"', 'C6EFCE')
for a, b in zip('DEFGHIJLMN', ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Z', 'AA', 'AB']):
    cf(ws, f'{a}{FIRST}:{a}{LAST}', f'AND($B{FIRST}="変更",$P{FIRST}="あり",{a}{FIRST}&""<>${b}{FIRST})', 'F8CBAD')
cf(ws, f'K{FIRST}:K{LAST}', f'AND($B{FIRST}="変更",$P{FIRST}="あり",$AG{FIRST}&""<>$Y{FIRST})', 'F8CBAD')
edl_in = [
    ('登録', 'edl-c2-ip', 'ip', 'https://feeds.example.com/c2-ip.txt', 'five_minute', '', '', '', 'C2サーバIP', '198.51.100.5', '', '', '', ''),
    ('登録', 'edl-phish-url', 'url', 'https://feeds.example.com/phish.txt', 'monthly', '01', '', '1', 'フィッシングURL', '', 'edl-cert-profile', '', '', ''),
    ('変更', 'edl-block-domain', 'domain', 'https://feeds.example.com/block-domain.txt', 'daily', '04', '', '', 'ブロックドメイン',
     'test.example.com\nqa.example.com', '', 'true', '', ''),
    ('登録', 'edl-torexit', 'predefined_ip', 'panw-torexit-ip-list', '', '', '', '', 'Tor出口ノード', '', '', '', '', ''),
    ('削除', 'edl-old-list', '', '', '', '', '', '', '', '', '', '', '', ''),
    ('変更', 'edl-allow-url', 'url', 'https://feeds.example.com/allow-url.txt', 'weekly', '02', 'sunday', '', '許可URL', '', 'edl-cert-profile', '', 'feeduser', ''),
    ('登録', 'edl-ng-hourly', 'ip', 'https://feeds.example.com/x.txt', 'hourly', '03', '', '', 'hourlyに時刻指定', '', '', '', '', ''),
    ('登録', 'edl-ng-expand', 'ip', 'https://feeds.example.com/y.txt', 'daily', '05', '', '', 'ipにexpand_domain', '', '', 'true', '', ''),
    ('登録', 'edl-ng-scheme', 'domain', 'feeds.example.com/z.txt', 'hourly', '', '', '', 'http://なし', '', '', '', '', ''),
]
for i, v in enumerate(edl_in):
    for j, x in enumerate(v):
        ws[f'{CL(2+j)}{FIRST+i}'] = x if x != '' else None

# =========================================================
# ⑥ 差分確認
# =========================================================
DF = ['type', 'url', '更新間隔', 'description', 'exception_list', 'certificate_profile', 'expand_domain', 'username']
def diffsheet(title, idx, heading, fields, widths, rows, btnrng):
    ws = sheet(title, idx)
    ncol = 3 + len(fields) + 5
    header(ws, heading,
           '④のAPI実行が終わると、VBAがSCMに今設定されている全件を取得して一覧表示し、既存一覧シート（実行前）と比べた変化をC列に付けます。'
           '行の背景色で変化を表します（追加=緑、変更=黄、変化なし=白、削除=グレー）。変更された項目のセルはオレンジ・太字です。削除されたものは一覧の後ろに、既存一覧（実行前）の内容で表示します。'
           '判定が要確認のもの（判定列が赤）はPush前に理由を確認してください。passwordは比較できません。ボタンでいつでも作り直せます。', ncol)
    put(ws, 'A3', '作成日時', Font(name=FN, size=10, bold=True))
    put(ws, 'B3', None, fl=F_VBA, nf='yyyy/mm/dd hh:mm:ss')
    for lab, val in [(('C3', 'D3'), '全件'), (('E3', 'F3'), '変化なし'), (('G3', 'H3'), '追加'),
                     (('C4', 'D4'), '変更'), (('E4', 'F4'), '削除'), (('G4', 'H4'), '要確認')]:
        stat(ws, lab, val, None, F_VBA)
    btn(ws, btnrng, '▶ 差分チェック（再作成）')
    put(ws, 'C5', '凡例：', Font(name=FN, size=10, bold=True), border=False)
    for a, t, c in [('D5', '追加', 'C6EFCE'), ('E5', '変更', 'FFFF99'), ('F5', '変化なし', 'FFFFFF'), ('G5', '削除', 'D9D9D9'), ('H5', '要確認(判定列)', 'FFC7CE')]:
        put(ws, a, t, fl=fill(c), al=Alignment(horizontal='center'))
    put(ws, 'I5', '変更セル', Font(name=FN, size=10, bold=True), fill('FFC000'), Alignment(horizontal='center'))
    names = ['No', 'name', '変化'] + [f + '（現在）' for f in fields] + ['変更項目', '変更前（既存一覧）', '④の操作', 'HTTP', '判定']
    heads(ws, 7, names, [5, 20, 9] + widths + [16, 34, 9, 7, 26])
    stats = {'': 0, '追加': 0, '変更': 0, '削除': 0, 'chk': 0}
    for i, v in enumerate(rows):
        r = FIRST + i
        vals = [i + 1] + list(v)
        for j, x in enumerate(vals):
            put(ws, f'{CL(1+j)}{r}', x if x != '' else None, al=Alignment(wrap_text=True, vertical='top'))
        ch = v[1]
        col = {'追加': 'C6EFCE', '変更': 'FFFF99', '削除': 'D9D9D9'}.get(ch)
        for j in range(len(vals)):
            c = ws[f'{CL(1+j)}{r}']
            if col: c.fill = fill(col)
        # 変更セル（オレンジ・太字）
        if ch == '変更':
            for fname in v[2 + len(fields)].split():
                c = ws[f'{CL(4 + fields.index(fname))}{r}']
                c.fill = fill('FFC000'); c.font = Font(name=FN, size=10, bold=True)
        stats['' if ch == '変化なし' else ch] += 1
        if v[-1] not in ('', '想定どおり') and not v[-1].startswith('変化なし'):
            stats['chk'] += 1
            ws[f'{CL(len(vals))}{r}'].fill = fill('FFC7CE')
    ws['D3'] = len(rows) - stats['削除']; ws['F3'] = stats['']; ws['H3'] = stats['追加']
    ws['D4'] = stats['変更']; ws['F4'] = stats['削除']; ws['H4'] = stats['chk']
    ws['B3'] = datetime.datetime(2026, 10, 28, 15, 40)
    ws.freeze_panes = 'D8'

E = ''
diffsheet(SD, 4, '⑥ 差分確認（External Dynamic List）：現在の設定一覧と既存一覧の比較', DF, [12, 34, 16, 18, 24, 16, 9, 11], [
    ('edl-block-ip', '変化なし', 'ip', 'https://feeds.example.com/block-ip.txt', 'hourly', '社内ブロックIP', E, E, E, E, E, E, E, E, E),
    ('edl-block-domain', '変更', 'domain', 'https://feeds.example.com/block-domain.txt', 'daily 04時', 'ブロックドメイン',
     'test.example.com,qa.example.com', E, 'true', E, '更新間隔 exception_list',
     'type: domain\nurl: https://feeds.example.com/block-domain.txt\n更新間隔: daily 03時\ndescription: ブロックドメイン\nexception_list: test.example.com\nexpand_domain: true',
     '変更', '200', '想定どおり'),
    ('edl-allow-url', '変化なし', 'url', 'https://feeds.example.com/allow-url.txt', 'weekly sunday 02時', '許可URL', E, 'edl-cert-profile', E, 'feeduser', E, E, E, E, E),
    ('edl-highrisk', '変化なし', 'predefined_ip', 'panw-highrisk-ip-list', E, '高リスクIP（例外あり）', '203.0.113.10', E, E, E, E, E, E, E, E),
    ('edl-c2-ip', '追加', 'ip', 'https://feeds.example.com/c2-ip.txt', 'five_minute', 'C2サーバIP', '198.51.100.5', E, E, E, E, E, '登録', '201', '想定どおり'),
    ('edl-phish-url', '追加', 'url', 'https://feeds.example.com/phish.txt', 'monthly 1日 01時', 'フィッシングURL', E, 'edl-cert-profile', E, E, E, E, '登録', '201', '想定どおり'),
    ('edl-torexit', '追加', 'predefined_ip', 'panw-torexit-ip-list', E, 'Tor出口ノード', E, E, E, E, E, E, '登録', '201', '想定どおり'),
    ('edl-old-list', '削除', 'ip', 'http://old.example.com/list.txt', 'five_minute', '旧リスト（廃止予定）', E, E, E, E, E, E, '削除', '200', '想定どおり'),
], 'J3:K4')

# ④ JSON_CURL サンプル（VBAのJSON発行と同じ形式）
import json as _j
BASE = 'https://api.sase.paloaltonetworks.com'
U = BASE + '/sse/config/v1/external-dynamic-lists'
def jstr(v): return _j.dumps(v, ensure_ascii=False)
def body(nm, ty, url, rec='', at='', dow='', dom='', desc='', exc='', cert='', expand='', user='', pw=''):
    pre = ty.startswith('predefined_')
    b = '"url":' + jstr(url)
    if not pre:
        rj = {'daily': '{"at":' + jstr(at) + '}', 'weekly': '{"day_of_week":' + jstr(dow) + ',"at":' + jstr(at) + '}',
              'monthly': '{"day_of_month":' + str(dom) + ',"at":' + jstr(at) + '}'}.get(rec, '{}')
        b += ',"recurring":{"' + rec + '":' + rj + '}'
    if desc: b += ',"description":' + jstr(desc)
    if exc: b += ',"exception_list":[' + ','.join(jstr(x) for x in exc.replace('\n', ',').split(',')) + ']'
    if not pre:
        if cert: b += ',"certificate_profile":' + jstr(cert)
        if ty == 'domain' and expand == 'true': b += ',"expand_domain":true'
        if user: b += ',"auth":{"username":' + jstr(user) + ',"password":' + jstr(pw) + '}'
    return '{"name":' + jstr(nm) + ',"type":{"' + ty + '":{' + b + '}}}'
def curl(m, u, b):
    c = f"curl -s -X {m} '{u}' -H \"Authorization: Bearer $TOKEN\""
    if b: c += " -H 'Content-Type: application/json' -d '" + b.replace("'", "'\\''") + "'"
    return c
rows = [
    ('登録', 'edl-c2-ip', 'POST', U + '?folder=Shared',
     body('edl-c2-ip', 'ip', 'https://feeds.example.com/c2-ip.txt', 'five_minute', desc='C2サーバIP', exc='198.51.100.5'), 201),
    ('登録', 'edl-phish-url', 'POST', U + '?folder=Shared',
     body('edl-phish-url', 'url', 'https://feeds.example.com/phish.txt', 'monthly', '01', dom=1, desc='フィッシングURL', cert='edl-cert-profile'), 201),
    ('変更', 'edl-block-domain', 'PUT', U + '/7e3a4c50-0002-4d2e-8a11-000000000022',
     body('edl-block-domain', 'domain', 'https://feeds.example.com/block-domain.txt', 'daily', '04', desc='ブロックドメイン',
          exc='test.example.com\nqa.example.com', expand='true'), 200),
    ('登録', 'edl-torexit', 'POST', U + '?folder=Shared',
     body('edl-torexit', 'predefined_ip', 'panw-torexit-ip-list', desc='Tor出口ノード'), 201),
    ('削除', 'edl-old-list', 'DELETE', U + '/7e3a4c50-0005-4d2e-8a11-000000000025', '', 200),
]
js = wb[SJ]
for i, (op, nm, m, u, b, st) in enumerate(rows):
    r = 10 + i
    for j, v in enumerate([i + 1, op, nm, m, u, b or None, curl(m, u, b), st]):
        put(js, f'{CL(1+j)}{r}', v, al=Alignment(wrap_text=True, vertical='top'))

# 名前定義
wb.defined_names['EX_FETCHED_AT'] = DefinedName('EX_FETCHED_AT', attr_text=f"{Q(SE)}!$B$3")
wb.defined_names['EDL_NAMES'] = DefinedName('EDL_NAMES', attr_text=f"OFFSET({Q(SE)}!$C$8,0,0,MAX(1,COUNTA({Q(SE)}!$C$8:$C$5000)),1)")
wb.defined_names['CERT_NAMES'] = DefinedName('CERT_NAMES', attr_text=f"OFFSET({Q(SE)}!$Q$8,0,0,MAX(1,COUNTA({Q(SE)}!$Q$8:$Q$5000)),1)")
wb.defined_names['PRE_NAMES'] = DefinedName('PRE_NAMES', attr_text=f"OFFSET({Q(SE)}!$S$8,0,0,MAX(1,COUNTA({Q(SE)}!$S$8:$S$27)),1)")
wb[SE]['B3'] = datetime.datetime(2026, 10, 28, 15, 30)

order = ['基本設定', SE, SI, SJ, SD, SL]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for w in wb: w.sheet_view.tabSelected = (w.title == '基本設定')
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('saved', OUT)

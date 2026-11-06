# -*- coding: utf-8 -*-
# SCM Application / ApplicationGroup Excel 生成
import sys, datetime
import json as _j
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

SRC = '/mnt/project-files/scm-service/SCM_ServiceObject.xlsx'
OUT = sys.argv[1]

AE, AI, GE, GI = 'AppObj_既存一覧', 'AppObj_登録変更削除', 'AppGroupObj_既存一覧', 'AppGroupObj_登録変更削除'
RF, CH = 'AppRef_定義済み・照会', 'App_選択肢'
SJ, AD, GD, SL = 'JSON_CURL', 'AppObj_差分確認', 'AppGroupObj_差分確認', '実行ログ'
FIRST, LAST, EXL, REFL = 8, 507, 5000, 10007

FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
F_TITLE, F_HEAD, F_IN, F_FX, F_VBA, F_BTN = fill('1F3864'), fill('2F5597'), fill('FFF2CC'), fill('F2F2F2'), fill('E2EFDA'), fill('C55A11')
F_REQ, F_OPT, F_HEAD2 = fill('C00000'), fill('7F7F7F'), fill('548235')

wb = openpyxl.load_workbook(SRC)
for n in ['SvcObj_既存一覧', 'SvcObj_登録変更削除', 'SvcGroupObj_既存一覧', 'SvcGroupObj_登録変更削除', 'SvcObj_差分確認', 'SvcGroupObj_差分確認']:
    del wb[n]
for n in ['EX_FETCHED_AT', 'EXG_FETCHED_AT', 'SVC_NAMES', 'SVCG_NAMES']:
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

def heads(ws, row, names, widths, start=1, fl=F_HEAD):
    for i, (n, w) in enumerate(zip(names, widths), start):
        put(ws, f'{CL(i)}{row}', n, Font(name=FN, size=10, bold=True, color='FFFFFF'), fl,
            Alignment(horizontal='center', vertical='center', wrap_text=True))
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[row].height = 30

def labels(ws, row, items):
    for col, t in items.items():
        put(ws, f'{col}{row}', t, Font(name=FN, size=9, bold=True, color='FFFFFF'),
            F_REQ if t.startswith('必須') else F_OPT, Alignment(horizontal='center', vertical='center'), False)

def stat(ws, addr, label, formula, fl=F_FX):
    put(ws, addr[0], label, Font(name=FN, size=10, bold=True))
    put(ws, addr[1], formula, Font(name=FN, size=10, bold=True), fl)

def cf(ws, rng, formula, color):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(color), stopIfTrue=False))

Q = lambda s: "'" + s + "'"
def ex(sh, col, last=EXL): return f"{Q(sh)}!${col}${FIRST}:${col}${last}"
REF_B, REF_C = ex(RF, 'B', REFL), ex(RF, 'C', REFL)

# ---------------- 項目定義（②と③で同じ列に並べる） ----------------
FLAGS = [('evasive_behavior', '回避的な動作'), ('consume_big_bandwidth', '帯域幅を多く使用'), ('used_by_malware', 'マルウェアが使用'),
         ('able_to_transfer_file', 'ファイル転送可'), ('has_known_vulnerability', '既知の脆弱性'),
         ('tunnel_other_application', '他Appをトンネル'), ('prone_to_misuse', '悪用されやすい'), ('pervasive_use', '広く使用'),
         ('tunnel_applications', '他Appの検出継続'), ('data_ident', 'データパターン検査'), ('file_type_ident', 'ファイルタイプ検査'),
         ('virus_ident', 'ウイルス検査'), ('no_appid_caching', 'App-IDキャッシュ無効')]
FIELDS = [('category', 16), ('subcategory', 18), ('technology', 15), ('risk\n(1-5)', 6),
          ('TCPポート\n(default)', 16), ('UDPポート\n(default)', 14), ('IPプロトコル\n(0-255)', 10),
          ('ICMPタイプ', 8), ('ICMPコード', 8), ('ICMPv6タイプ', 8), ('ICMPv6コード', 8),
          ('description', 22), ('parent_app', 14),
          ('timeout\n(0-604800)', 9), ('tcp_timeout\n(0-604800)', 9), ('udp_timeout\n(0-604800)', 9),
          ('tcp_half_closed\n_timeout\n(1-604800)', 10), ('tcp_time_wait\n_timeout\n(1-600)', 10)] + \
         [(f'{k}\n({j})', 9) for k, j in FLAGS]
FC = [CL(4 + i) for i in range(len(FIELDS))]       # D..AH
assert FC[-1] == 'AH' and FC[17] == 'U' and FC[18] == 'V'
FLAG_RNG = ('V', 'AH')

# ---------------- 共通の数式部品 ----------------
NAMECHK = ('IF(SUMPRODUCT(--ISERROR(FIND(MID($C{r}&REPT(" ",63),ROW($1:$63),1),'
           '" 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")))>0,'
           '"NG:nameに使用不可の文字(英数字 空白 . _ - のみ)",')

def listchk(s, label, mx):
    """「80」「8000-8080」「80,443,8000-8080」形式（0～mx）のチェック。OKなら空文字"""
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
         f'NG:{label}は数字・カンマ・ハイフンのみ(空白不可)'),
        (f'OR(LEFT({s},1)=",",RIGHT({s},1)=",",ISNUMBER(FIND(",,",{s})))',
         f'NG:{label}の先頭・末尾のカンマや連続したカンマを削除してください'),
        (f'SUMPRODUCT(--(LEN({seg})=12),--ISERROR(FIND(",",{seg})))>0',
         f'NG:{label}に形式不正の要素があります'),
        (f'SUMPRODUCT({start},--(((LEN({E})-LEN(SUBSTITUTE({E},"-","")))>1)+(LEFT({E},1)="-")+(RIGHT({E},1)="-")>0))>0',
         f'NG:{label}の範囲は 開始-終了 の形式'),
        (f'SUMPRODUCT({start},--(((--(0&{A}))>{mx})+((--(0&{B}))>{mx})>0))>0',
         f'NG:{label}の番号は0～{mx}'),
        (f'SUMPRODUCT({start},--ISNUMBER(FIND("-",{E})),--((--(0&{A}))>(--(0&{B}))))>0',
         f'NG:{label}の範囲の開始番号が終了番号より大きい'),
    ]
    f = '""'
    for cond, msg in reversed(parts):
        f = f'IF({cond},"{msg}",{f})'
    return f

def numchk(c, mn, mx, label):
    m = f'"NG:{label}は{mn}～{mx}の整数"'
    return f'IF({c}="","",IF(ISERROR(--{c}),{m},IF(OR(--{c}<{mn},--{c}>{mx},--{c}<>INT(--{c})),{m},"")))'

def nest(conds, last):
    f = last
    for cond, val in reversed(conds):
        f = f'IF({cond},{val},{f})'
    return f

def canon(r):   # 定義（全項目を | で連結）：差分なし判定用
    return '&"|"&'.join(f'${c}{r}' for c in FC)

# =========================================================
# 基本設定 / JSON_CURL / 実行ログ（Serviceツールから流用して文言を変更）
# =========================================================
cfg = wb['基本設定']
cfg['B7'] = '/sse/config/v1/applications'
cfg['D7'] = 'Application（ApplicationGroupは /sse/config/v1/application-groups、フィルタは /sse/config/v1/application-filters をVBA内で使用）'

js = wb[SJ]
js['A2'] = ('③⑦のJSON発行ボタンでVBAが10行目以降を生成します（緑=VBA書込み）。curlはbash形式。先に5行目でトークン取得→6行目でTOKEN変数を設定してから各行を実行。'
            '「API実行」ボタンでExcelから直接実行も可能（結果は⑤実行ログへ）。JSON発行は③と⑦のどちらのボタンでも両方から一覧を作り直し、'
            'Application登録・変更→Group登録・変更→Group削除→Application削除の順に並べます。削除するApplication／GroupがほかのGroupのメンバーなら、'
            'Groupから外す変更（変更(Group)自動）を先に追加します。Groupの中にGroupを入れる場合は、中のGroupを先に登録します。')
for row in js.iter_rows(min_row=10, max_row=js.max_row):
    for c in row: c.value = None

lg = wb[SL]
for row in lg.iter_rows(min_row=6, max_row=lg.max_row):
    for c in row:
        c.value = None; c.border = Border(); c.number_format = 'General'

# =========================================================
# App_選択肢（category / subcategory / technology）
# =========================================================
CATS = {
    'business-systems': ['auth-service', 'database', 'erp-crm', 'general-business', 'management', 'office-programs',
                         'software-development', 'software-update', 'storage-backup'],
    'collaboration': ['email', 'instant-messaging', 'internet-conferencing', 'social-business', 'social-networking',
                      'voip-video', 'web-posting'],
    'general-internet': ['file-sharing', 'internet-utility'],
    'media': ['audio-streaming', 'gaming', 'photo-video'],
    'networking': ['encrypted-tunnel', 'infrastructure', 'ip-protocol', 'proxy', 'remote-access', 'routing'],
}
TECHS = ['browser-based', 'client-server', 'network-protocol', 'peer-to-peer']
ws = sheet(CH, 1)
header(ws, '選択肢（category・subcategory・technology）',
       'category・subcategory・technologyのプルダウンと判定に使います（PAN-OS／SCMのApp-ID分類）。SCMで選べる値が増えた場合はA・B列（組み合わせ）、D列、F列に行を追加してください。'
       'subcategoryは、A列のcategoryとの組み合わせでチェックします。', 6)
heads(ws, 4, ['category', 'subcategory', '', 'technology', '', 'category（一覧）'], [20, 24, 3, 20, 3, 22])
for c in 'CE':
    ws[f'{c}4'].fill = PatternFill(); ws[f'{c}4'].border = Border()
r = 5
for cat, subs in CATS.items():
    for s in subs:
        put(ws, f'A{r}', cat, nf='@'); put(ws, f'B{r}', s, nf='@'); r += 1
for i, t in enumerate(TECHS):
    put(ws, f'D{5+i}', t, nf='@')
for i, c in enumerate(CATS):
    put(ws, f'F{5+i}', c, nf='@')
for nm, col in [('APP_SUB_CAT', 'A'), ('APP_SUB_LIST', 'B'), ('APP_TECH_LIST', 'D'), ('APP_CAT_LIST', 'F')]:
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"OFFSET({Q(CH)}!${col}$5,0,0,MAX(1,COUNTA({Q(CH)}!${col}$5:${col}$300)),1)")

# =========================================================
# ② AppObj_既存一覧
# =========================================================
ws = sheet(AE, 1)
header(ws, '② 既存情報（SCMから取得した Application（カスタムアプリケーション）一覧）',
       'ボタンを押すとAPI（GET）でApplication・ApplicationGroup・アプリケーションフィルタを取得し、②⑥とAppRefシートを作り直します（B～AJ列=VBA書込み、A・AK列=関数）。直接編集はしないでください。'
       'defaultのポートは TCP・UDP 別にまとめて表示します。特性・オプションは true のものに○。AJ列はこのツールで編集しない項目（signature等）で、変更時にそのまま引き継ぎます。', 37)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{EXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
heads(ws, 7, ['No', 'id', 'name'] + [f[0] for f in FIELDS] + ['コンテナ', '引継ぎ項目(JSON)\nsignature等', '定義(関数)'],
      [6, 38, 20] + [f[1] for f in FIELDS] + [16, 30, 30])
ws.row_dimensions[7].height = 58
SIG = ('"signature":[{"name":"sig-portal","and_condition":[{"name":"and1","or_condition":[{"name":"or1","operator":'
       '{"pattern_match":{"context":"http-req-host-header","pattern":"portal\\\\.example\\\\.co\\\\.jp"}}}]}]}]')
def app(cat, sub, tech, risk, tcp='', udp='', ipp='', desc='', parent='', tmo=None, flags=(), it='', ic=''):
    v = {c: '' for c in FC}
    v.update({'D': cat, 'E': sub, 'F': tech, 'G': str(risk), 'H': tcp, 'I': udp, 'J': ipp, 'K': it, 'L': ic, 'O': desc, 'P': parent})
    for k, x in (tmo or {}).items(): v[k] = x
    for fl in flags: v[FC[18 + [f[0] for f in FLAGS].index(fl)]] = '○'
    return v
app_exist = [
    ('corp-erp', app('business-systems', 'erp-crm', 'client-server', 3, tcp='8443,9443', desc='社内ERP'), ''),
    ('corp-chat', app('collaboration', 'instant-messaging', 'client-server', 2, tcp='5222', desc='社内チャット',
                      flags=['able_to_transfer_file']), ''),
    ('corp-backup', app('business-systems', 'storage-backup', 'client-server', 2, tcp='873', desc='バックアップ転送',
                        tmo={'R': '7200'}, flags=['able_to_transfer_file', 'consume_big_bandwidth']), ''),
    ('corp-legacy-ftp', app('general-internet', 'file-sharing', 'client-server', 4, tcp='2121', desc='旧FTP（廃止予定）',
                            flags=['able_to_transfer_file', 'has_known_vulnerability']), ''),
    ('corp-gre-tun', app('networking', 'ip-protocol', 'network-protocol', 3, ipp='47', desc='拠点間GRE'), ''),
    ('corp-web-portal', app('business-systems', 'general-business', 'browser-based', 1, tcp='443', desc='社内ポータル'), SIG),
]
for i, (nm, v, extra) in enumerate(app_exist):
    r = FIRST + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)')
    put(ws, f'B{r}', f'7a3c1e50-000{i+1}-4b2e-9c11-00000000000{i+1}', nf='@')
    put(ws, f'C{r}', nm, nf='@')
    for c in FC:
        put(ws, f'{c}{r}', v[c] or None, nf='@')
    put(ws, f'AI{r}', 'folder=Shared', nf='@')
    put(ws, f'AJ{r}', extra or None, nf='@')
    put(ws, f'AK{r}', f'=IF(C{r}="","",{canon(r)})')
ws.freeze_panes = 'D8'

# =========================================================
# ③ AppObj_登録変更削除
# =========================================================
ws = sheet(AI, 2)
header(ws, '③ 今回の登録・変更・削除一覧（Application：カスタムアプリケーション）',
       '黄色(B～AH列)に入力。nameは既存一覧からプルダウン選択可（変更・削除時）。AI列以降は関数で②⑥⑦・AppRefと自動照合し、判定がNGの行はJSON発行対象外。変更時に既存と異なる項目は橙色で表示。'
       'nameは31文字以内で、既存のApplication・ApplicationGroup・アプリケーションフィルタ・定義済みApp-IDと同じ名前は使えません（大文字・小文字の違いだけでも不可）。'
       'defaultはポート(TCP/UDP)・IPプロトコル・ICMP・ICMPv6のどれか1種類。ポートは「80」「8000-8080」「80,443,8000-8080」形式か「dynamic」。特性・オプションは「○」で true。'
       '変更(PUT)は全項目置換のため、残したい値もすべて入力してください（signature等はそのまま引き継ぎます）。', 50, 70)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(AM{FIRST}:AM{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(AM{FIRST}:AM{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(AM{FIRST}:AM{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行（③⑦）→ ④へ')
btn(ws, 'I3:J4', '▶ 名前照会（API）')
btn(ws, 'L3:M4', '▶ 入力クリア')
lab = {'B': '必須', 'C': '必須', 'D': '必須※', 'E': '必須※', 'F': '必須※', 'G': '必須※'}
lab.update({c: '任意' for c in FC[4:]})
labels(ws, 6, lab)
put(ws, 'AI6', '※ category・subcategory・technology・riskは登録/変更時に必須（削除時は不要）。任意欄は空欄可（変更時の空欄は既存値が消えます）',
    Font(name=FN, size=9, color='595959'), border=False)
CHK = ['既存有無', '既存id', '既存行', '一覧内重複', '判定', '理由', 'TCPポート\nチェック', 'UDPポート\nチェック',
       'IPプロトコル\nチェック', 'ICMPタイプ\nチェック', 'ICMPコード\nチェック', 'ICMPv6タイプ\nチェック', 'ICMPv6コード\nチェック',
       'timeout\nチェック', '名前照会\n(AppRef)', '定義(入力)']
heads(ws, 7, ['No', '操作', 'name'] + [f[0] for f in FIELDS] + CHK,
      [5, 8, 18] + [f[1] for f in FIELDS] + [9, 22, 6, 9, 8, 48, 22, 22, 22, 18, 18, 18, 18, 22, 12, 20])
ws.row_dimensions[7].height = 58
for c in FC[18:]:
    ws[f'{c}7'].fill = F_HEAD2
EXT = Q(AE)
for r in range(FIRST, LAST + 1):
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    put(ws, f'B{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
    put(ws, f'C{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
    for c in FC:
        put(ws, f'{c}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@',
            al=Alignment(horizontal='center') if c >= 'V' or len(c) == 2 else None)
    put(ws, f'AI{r}', f'=IF($C{r}="","",IF(COUNTIF({ex(AE,"C")},$C{r})>0,"あり","なし"))', fl=F_FX)
    put(ws, f'AJ{r}', f'=IF($AI{r}="あり",INDEX({ex(AE,"B")},$AK{r})&"","")', fl=F_FX)
    put(ws, f'AK{r}', f'=IF($AI{r}="あり",MATCH($C{r},{ex(AE,"C")},0),"")', fl=F_FX)
    put(ws, f'AL{r}', f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,"重複",""))', fl=F_FX)
    put(ws, f'AM{r}', f'=IF($AN{r}="","",IF($AN{r}="OK","OK",LEFT($AN{r},FIND(":",$AN{r})-1)))', fl=F_FX)
    nd = f'$B{r}<>"削除"'
    for col, src, label, mx in [('AO', 'H', 'TCPポート', 65535), ('AP', 'I', 'UDPポート', 65535), ('AQ', 'J', 'IPプロトコル', 255),
                                ('AR', 'K', 'ICMPタイプ', 255), ('AS', 'L', 'ICMPコード', 255),
                                ('AT', 'M', 'ICMPv6タイプ', 255), ('AU', 'N', 'ICMPv6コード', 255)]:
        dyn = f'IF(${src}{r}="dynamic","",' if src in 'HI' else '('
        put(ws, f'{col}{r}', f'=IF(OR($C{r}="",$B{r}="削除",${src}{r}=""),"",{dyn}{listchk(f"${src}{r}", label, mx)}))', fl=F_FX)
    tm = [numchk(f'$Q{r}', 0, 604800, 'timeout'), numchk(f'$R{r}', 0, 604800, 'tcp_timeout'),
          numchk(f'$S{r}', 0, 604800, 'udp_timeout'), numchk(f'$T{r}', 1, 604800, 'tcp_half_closed_timeout'),
          numchk(f'$U{r}', 1, 600, 'tcp_time_wait_timeout')]
    f = tm[-1]
    for t in reversed(tm[:-1]):
        f = f'IF({t}<>"",{t},{f})'
    put(ws, f'AV{r}', f'=IF(OR($C{r}="",$B{r}="削除"),"",{f})', fl=F_FX)
    put(ws, f'AW{r}', f'=IF(OR($C{r}="",$B{r}<>"登録"),"",IF(COUNTIFS({REF_C},$C{r},{REF_B},"<>照会:なし")>0,"重複あり",'
                      f'IF(COUNTIF({REF_C},$C{r})>0,"照会済み(なし)","未照会")))', fl=F_FX)
    put(ws, f'AX{r}', f'=IF($C{r}="","",{canon(r)})', fl=F_FX)
    grpcnt = f'SUMPRODUCT(--ISNUMBER(FIND(","&$C{r}&",",","&{ex(GE,"D")}&",")),--({ex(GE,"D")}<>""))'
    EXV = lambda col: f'IF($AK{r}="","",INDEX({ex(AE,col)},$AK{r})&"")'
    exextra = EXV('AJ')
    deftypes = (f'(--(($H{r}<>"")+($I{r}<>"")>0)+--($J{r}<>"")+--(($K{r}<>"")+($L{r}<>"")>0)'
                f'+--(($M{r}<>"")+($N{r}<>"")>0))')
    flagbad = f'SUMPRODUCT(--(${FLAG_RNG[0]}{r}:${FLAG_RNG[1]}{r}<>""),--(${FLAG_RNG[0]}{r}:${FLAG_RNG[1]}{r}<>"○"))>0'
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$AL{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>31', '"NG:nameが31文字超"'),
        (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
    ]
    rest = [
        (f'AND($B{r}="登録",$C{r}="any")', '"NG:予約名(any)は使用できません"'),
        (f'AND($B{r}="登録",$AI{r}="あり")', '"NG:既存に同名のApplicationあり(登録不可。大文字・小文字の違いだけでも不可)"'),
        (f'AND($B{r}="登録",COUNTIF({ex(GE,"C")},$C{r})>0)', '"NG:同名のApplicationGroupが既存にあり(Application・Group・フィルタで名前は共通)"'),
        (f'AND($B{r}="登録",COUNTIFS({ex(GI,"B",LAST)},"登録",{ex(GI,"C",LAST)},$C{r})>0)', '"NG:⑦で同名のApplicationGroupを登録予定"'),
        (f'$AW{r}="重複あり"', '"NG:定義済みApp-ID・アプリケーションフィルタ等に同名あり(AppRef_定義済み・照会シート参照)"'),
        (f'AND($B{r}<>"登録",$AI{r}="なし")', '"NG:既存エントリ無"'),
        (f'AND($B{r}<>"登録",NOT(EXACT({EXV("C")},$C{r})))', '"NG:nameの大文字・小文字が既存と異なります(既存一覧の表記で入力)"'),
        (f'AND($B{r}="変更",{exextra}="#TOO_LONG")', '"NG:既存のsignature等が大きすぎるため、このツールでは変更できません(SCM画面で変更してください)"'),
        (f'AND({nd},OR($D{r}="",$E{r}="",$F{r}="",$G{r}=""))', '"NG:category・subcategory・technology・riskは必須"'),
        (f'AND({nd},COUNTIF(APP_CAT_LIST,$D{r})=0)', '"NG:categoryが選択肢にありません(App_選択肢シート参照)"'),
        (f'AND({nd},COUNTIFS(APP_SUB_CAT,$D{r},APP_SUB_LIST,$E{r})=0)', f'"NG:subcategoryがcategory("&$D{r}&")に対応していません(App_選択肢シート参照)"'),
        (f'AND({nd},COUNTIF(APP_TECH_LIST,$F{r})=0)', '"NG:technologyが選択肢にありません(App_選択肢シート参照)"'),
        (f'AND({nd},ISERROR(MATCH($G{r}&"",{{"1","2","3","4","5"}},0)))', '"NG:riskは1～5の整数"'),
        (f'AND({nd},{deftypes}>1)', '"NG:defaultはポート(TCP/UDP)・IPプロトコル・ICMP・ICMPv6のいずれか1種類だけ指定できます"'),
        (f'AND({nd},$K{r}="",$L{r}<>"")', '"NG:ICMPコードだけは指定できません(ICMPタイプが必須)"'),
        (f'AND({nd},$M{r}="",$N{r}<>"")', '"NG:ICMPv6コードだけは指定できません(ICMPv6タイプが必須)"'),
    ] + [(f'LEFT(${c}{r},2)="NG"', f'${c}{r}') for c in ['AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV']] + [
        (f'AND({nd},LEN($O{r})>1023)', '"NG:descriptionが1023文字超"'),
        (f'AND({nd},LEN($P{r})>127)', '"NG:parent_appが127文字超"'),
        (f'AND({nd},LENB($P{r})<>LEN($P{r}))', '"NG:parent_appに全角文字が含まれます"'),
        (f'AND({nd},$P{r}<>"",$P{r}=$C{r})', '"NG:parent_appに自分自身は指定できません"'),
        (f'AND({nd},{flagbad})', '"NG:特性・オプション(V～AH列)は「○」か空欄で入力してください"'),
        (f'AND($B{r}="変更",$AX{r}={EXV("AK")})', '"警告:既存と差分なし"'),
        (f'AND($B{r}="削除",{grpcnt}>0)',
         f'"警告:Group("&{grpcnt}&"件)のメンバー。JSON発行時にGroupから外す変更を自動で追加します"'),
        (f'$AW{r}="未照会"', '"警告:定義済みApp-IDとの重複が未確認です(「名前照会（API）」ボタンで確認してください)"'),
        (f'AND($B{r}="変更",{exextra}<>"")', '"警告:既存のsignature等(このツールで編集しない項目)はそのまま引き継ぎます"'),
    ]
    reason = nest(conds, NAMECHK.format(r=r) + nest(rest, '"OK"') + ')')
    put(ws, f'AN{r}', f'=IF($C{r}="","",{reason})', fl=F_FX)
ws.freeze_panes = 'D8'
dv = DataValidation(type='list', formula1='"登録,変更,削除"', allow_blank=True); dv.add(f'B{FIRST}:B{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='APP_NAMES', allow_blank=True, showErrorMessage=False); dv.add(f'C{FIRST}:C{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='APP_CAT_LIST', allow_blank=True, showErrorMessage=False); dv.add(f'D{FIRST}:D{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='APP_SUB_LIST', allow_blank=True, showErrorMessage=False); dv.add(f'E{FIRST}:E{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='APP_TECH_LIST', allow_blank=True, showErrorMessage=False); dv.add(f'F{FIRST}:F{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='"1,2,3,4,5"', allow_blank=True); dv.add(f'G{FIRST}:G{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='"○"', allow_blank=True); dv.add(f'V{FIRST}:AH{LAST}'); ws.add_data_validation(dv)
cf(ws, f'AM{FIRST}:AN{LAST}', f'$AM{FIRST}="NG"', 'FFC7CE')
cf(ws, f'AM{FIRST}:AN{LAST}', f'$AM{FIRST}="警告"', 'FFEB9C')
cf(ws, f'AM{FIRST}:AN{LAST}', f'$AM{FIRST}="OK"', 'C6EFCE')
cf(ws, f'D{FIRST}:AH{LAST}', f'AND($B{FIRST}="変更",$AI{FIRST}="あり",D{FIRST}&""<>INDEX(APP_EX_TBL,$AK{FIRST},COLUMN(D{FIRST}))&"")', 'F8CBAD')
app_in = [
    ('登録', 'corp-hr', app('business-systems', 'general-business', 'browser-based', 2, tcp='443,8443', desc='人事システム')),
    ('登録', 'corp-iot', app('networking', 'infrastructure', 'client-server', 3, udp='5683', desc='IoTセンサー（CoAP）')),
    ('変更', 'corp-chat', app('collaboration', 'instant-messaging', 'client-server', 3, tcp='5222,5223', desc='社内チャット（TLSポート追加）',
                            flags=['able_to_transfer_file'])),
    ('変更', 'corp-web-portal', app('business-systems', 'general-business', 'browser-based', 1, tcp='443', desc='社内ポータル（新）')),
    ('削除', 'corp-legacy-ftp', {c: '' for c in FC}),
    ('登録', 'web-browsing', app('general-internet', 'internet-utility', 'browser-based', 4, tcp='80', desc='定義済みApp-IDと同名')),
    ('登録', 'Corp-ERP', app('business-systems', 'erp-crm', 'client-server', 3, tcp='8443', desc='大文字・小文字違いの同名')),
    ('登録', 'corp-ng-sub', app('media', 'erp-crm', 'client-server', 2, tcp='8080', desc='categoryとsubcategoryの不一致')),
    ('登録', 'corp-ng-port', app('business-systems', 'database', 'client-server', 2, tcp='70000', desc='番号が範囲外')),
    ('登録', 'corp-ng-def', app('networking', 'routing', 'network-protocol', 2, tcp='179', ipp='89', desc='default 2種類')),
    ('登録', 'corp-ng-flag', app('business-systems', 'database', 'client-server', 2, tcp='1521', desc='特性に○以外',
                               flags=['pervasive_use'])),
]
app_in[-1][2]['AC'] = 'yes'
for i, (op, nm, v) in enumerate(app_in):
    r = FIRST + i
    ws[f'B{r}'] = op; ws[f'C{r}'] = nm
    for c in FC:
        ws[f'{c}{r}'] = v[c] or None

# =========================================================
# ⑥ AppGroupObj_既存一覧
# =========================================================
ws = sheet(GE, 3)
header(ws, '⑥ 既存情報（SCMから取得した ApplicationGroup 一覧）',
       'ボタンを押すとAPI（GET）でApplication・ApplicationGroup・フィルタを取得し、表を再生成します（緑=VBA書込み）。メンバーはApplication名（定義済みApp-IDを含む）・ApplicationGroup名・フィルタ名をカンマ区切りで並べたものです。', 5)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{EXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
heads(ws, 7, ['No', 'id', 'name', 'メンバー', 'コンテナ'], [6, 38, 22, 60, 16])
grp_exist = [
    ('grp-corp-apps', 'corp-erp,corp-chat,corp-web-portal'),
    ('grp-transfer', 'corp-backup,corp-legacy-ftp,ftp'),
    ('grp-all-corp', 'grp-corp-apps,ssl,web-browsing'),
]
for i, v in enumerate(grp_exist):
    r = FIRST + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)')
    put(ws, f'B{r}', f'8b4d2f60-000{i+1}-4b2e-9c11-00000000001{i+1}', nf='@')
    for j, x in enumerate(v):
        put(ws, f'{CL(3+j)}{r}', x, nf='@')
    put(ws, f'E{r}', 'folder=Shared', nf='@')
ws.freeze_panes = 'D8'

# =========================================================
# ⑦ AppGroupObj_登録変更削除
# =========================================================
ws = sheet(GI, 4)
header(ws, '⑦ 今回の登録・変更・削除一覧（ApplicationGroup）',
       '黄色(B～D列)に入力。メンバーはApplication名（定義済みApp-IDも可）・ApplicationGroup名・フィルタ名を「,」区切りかセル内改行（Alt+Enter）で入力します。'
       'E～S列は関数で②⑥・③⑦・AppRefと照合します（大文字・小文字も区別）。どこにも無いメンバーは「名前照会（API）」でSCMに問い合わせ、無いと分かればNG、未照会なら警告です。'
       '判定がNGの行はJSON発行対象外。ApplicationGroupにはdescription・tagがありません（API仕様）。変更(PUT)は全項目置換です。', 19, 60)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(J{FIRST}:J{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(J{FIRST}:J{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(J{FIRST}:J{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行（③⑦）→ ④へ')
btn(ws, 'I3:J4', '▶ 名前照会（API）')
btn(ws, 'L3:M4', '▶ 入力クリア')
labels(ws, 6, {'B': '必須', 'C': '必須', 'D': '必須※'})
put(ws, 'E6', '※ メンバーは登録/変更時に必須（削除時は不要）', Font(name=FN, size=9, color='595959'), border=False)
heads(ws, 7, ['No', '操作', 'name', 'メンバー', '既存有無', '既存id', '既存メンバー', '一覧内重複',
              'メンバー確認', '判定', '理由', '名前照会\n(AppRef)', '作業:メンバー(正規化)', '作業:メンバー数', '作業:一致数(延べ)',
              '作業:一致数(種類)', '作業:登録予定数', '作業:削除予定数', '作業:照会で無し'],
      [5, 8, 20, 40, 9, 22, 36, 10, 18, 8, 48, 12, 30, 9, 9, 9, 9, 9, 9])
ws.row_dimensions[7].height = 44
for r in range(FIRST, LAST + 1):
    M2 = f'","&SUBSTITUTE($M{r},",",",,")&","'
    FN_ = f'","&$M{r}&","'
    def occ(rng, conds=''):
        return f'SUMPRODUCT(--({rng}<>""){conds},(LEN({M2})-LEN(SUBSTITUTE({M2},","&{rng}&",","")))/(LEN({rng})+2))'
    def dis(rng, conds=''):
        return f'SUMPRODUCT(--({rng}<>""){conds},--ISNUMBER(FIND(","&{rng}&",",{FN_})))'
    c3 = f',--({ex(AI,"B",LAST)}="登録"),--({ex(AI,"AM",LAST)}<>"NG")'
    c7 = (f',--({ex(GI,"B",LAST)}="登録"),--(COUNTIF({ex(GE,"C")},{ex(GI,"C",LAST)})=0)'
          f',--(COUNTIF({ex(AE,"C")},{ex(GI,"C",LAST)})=0)')
    ck = f',--({REF_B}<>"照会:なし")'
    cn = f',--({REF_B}="照会:なし")'
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    for col in 'BCD':
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
    put(ws, f'E{r}', f'=IF($C{r}="","",IF(COUNTIF({ex(GE,"C")},$C{r})>0,"あり","なし"))', fl=F_FX)
    for col, src in {'F': 'B', 'G': 'D'}.items():
        put(ws, f'{col}{r}', f'=IF($E{r}="あり",INDEX({ex(GE,src)},MATCH($C{r},{ex(GE,"C")},0))&"","")', fl=F_FX)
    put(ws, f'H{r}', f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,"重複",""))', fl=F_FX)
    put(ws, f'L{r}', f'=IF(OR($C{r}="",$B{r}<>"登録"),"",IF(COUNTIFS({REF_C},$C{r},{REF_B},"<>照会:なし")>0,"重複あり",'
                     f'IF(COUNTIF({REF_C},$C{r})>0,"照会済み(なし)","未照会")))', fl=F_FX)
    put(ws, f'M{r}', f'=IF($D{r}="","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($D{r},CHAR(13),""),CHAR(10),",")," ",""))', fl=F_FX)
    put(ws, f'N{r}', f'=IF($M{r}="",0,LEN($M{r})-LEN(SUBSTITUTE($M{r},",",""))+1)', fl=F_FX)
    put(ws, f'O{r}', f'=IF($M{r}="",0,{occ(ex(AE,"C"))}+{occ(ex(GE,"C"))}+{occ(ex(AI,"C",LAST),c3)}+{occ(ex(GI,"C",LAST),c7)}'
                     f'+{occ(REF_C,ck)})', fl=F_FX)
    put(ws, f'P{r}', f'=IF($M{r}="",0,{dis(ex(AE,"C"))}+{dis(ex(GE,"C"))}+{dis(ex(AI,"C",LAST),c3)}+{dis(ex(GI,"C",LAST),c7)}'
                     f'+{dis(REF_C,ck)})', fl=F_FX)
    put(ws, f'Q{r}', f'=IF($M{r}="",0,{dis(ex(AI,"C",LAST),c3)}+{dis(ex(GI,"C",LAST),c7)})', fl=F_FX)
    put(ws, f'R{r}', f'=IF($M{r}="",0,{dis(ex(AI,"C",LAST),chr(44)+"--("+ex(AI,"B",LAST)+"=\"削除\"),--("+ex(AI,"AM",LAST)+"<>\"NG\")")}'
                     f'+{dis(ex(GI,"C",LAST),chr(44)+"--("+ex(GI,"B",LAST)+"=\"削除\")")})', fl=F_FX)
    put(ws, f'S{r}', f'=IF($M{r}="",0,{occ(REF_C,cn)}-{occ(ex(AI,"C",LAST),c3+",--("+ex(AI,"AW",LAST)+"=\"照会済み(なし)\")")}'
                     f'-{occ(ex(GI,"C",LAST),c7+",--("+ex(GI,"L",LAST)+"=\"照会済み(なし)\")")})', fl=F_FX)
    put(ws, f'I{r}', f'=IF(OR($C{r}="",$B{r}="削除",$M{r}=""),"",'
                     f'IF(OR(LEFT($M{r},1)=",",RIGHT($M{r},1)=",",ISNUMBER(FIND(",,",$M{r}))),"空の要素あり",'
                     f'IF(ISNUMBER(FIND(","&$C{r}&",",","&$M{r}&",")),"自分自身を含む",'
                     f'IF($S{r}>0,"SCMに無い("&$S{r}&"件)",'
                     f'IF($O{r}>$P{r},"メンバー重複",'
                     f'IF($O{r}<$N{r},"未確認あり("&($N{r}-$O{r})&"件)",'
                     f'IF($Q{r}>0,"登録予定を含む","全て存在")))))))', fl=F_FX)
    put(ws, f'J{r}', f'=IF($K{r}="","",IF($K{r}="OK","OK",LEFT($K{r},FIND(":",$K{r})-1)))', fl=F_FX)
    gcnt = f'SUMPRODUCT(--ISNUMBER(FIND(","&$C{r}&",",","&{ex(GE,"D")}&",")),--({ex(GE,"D")}<>""))'
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$H{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>31', '"NG:nameが31文字超"'),
        (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
    ]
    rest = [
        (f'AND($B{r}="登録",$C{r}="any")', '"NG:予約名(any)は使用できません"'),
        (f'AND($B{r}="登録",$E{r}="あり")', '"NG:既存に同名のApplicationGroupあり(登録不可。大文字・小文字の違いだけでも不可)"'),
        (f'AND($B{r}="登録",COUNTIF({ex(AE,"C")},$C{r})>0)', '"NG:同名のApplicationが既存にあり(Application・Group・フィルタで名前は共通)"'),
        (f'AND($B{r}="登録",COUNTIFS({ex(AI,"B",LAST)},"登録",{ex(AI,"C",LAST)},$C{r})>0)', '"NG:③で同名のApplicationを登録予定"'),
        (f'$L{r}="重複あり"', '"NG:定義済みApp-ID・アプリケーションフィルタ等に同名あり(AppRef_定義済み・照会シート参照)"'),
        (f'AND($B{r}<>"登録",$E{r}="なし")', '"NG:既存エントリ無"'),
        (f'AND($B{r}<>"登録",NOT(EXACT(IF($E{r}="あり",INDEX({ex(GE,"C")},MATCH($C{r},{ex(GE,"C")},0)),""),$C{r})))', '"NG:nameの大文字・小文字が既存と異なります(既存一覧の表記で入力)"'),
        (f'AND($B{r}<>"削除",$D{r}="")', '"NG:メンバーが未入力"'),
        (f'AND($B{r}<>"削除",LENB($D{r})<>LEN($D{r}))', '"NG:メンバーに全角文字が含まれます"'),
        (f'$I{r}="空の要素あり"', '"NG:メンバーの先頭・末尾の区切り(カンマ/改行)や連続した区切りを削除してください"'),
        (f'$I{r}="自分自身を含む"', '"NG:メンバーに自分自身は指定できません"'),
        (f'LEFT($I{r},6)="SCMに無い"', '"NG:SCMに存在しないメンバーがあります(名前照会で無しと確認済み。AppRefシートの「照会:なし」行を参照。大文字・小文字も区別)"'),
        (f'$I{r}="メンバー重複"', '"NG:同じメンバーが重複しています"'),
        (f'LEFT($I{r},5)="未確認あり"', f'"警告:②⑥③⑦・AppRefのどこにも無いメンバーが"&($N{r}-$O{r})&"件あります。定義済みApp-IDなら問題ありません(「名前照会（API）」で確認できます)"'),
        (f'AND($B{r}<>"削除",$R{r}>0)', '"警告:③⑦で削除予定のメンバーを含む(JSON発行時にメンバーから外します)"'),
        (f'$I{r}="登録予定を含む"', '"警告:③⑦で登録予定のメンバーを含む(④で先に登録します)"'),
        (f'AND($B{r}="変更",$M{r}=$G{r})', '"警告:既存と差分なし"'),
        (f'AND($B{r}="削除",{gcnt}>0)',
         f'"警告:他のGroup("&{gcnt}&"件)のメンバー。JSON発行時にGroupから外す変更を自動で追加します"'),
        (f'$L{r}="未照会"', '"警告:定義済みApp-IDとの重複が未確認です(「名前照会（API）」ボタンで確認してください)"'),
    ]
    reason = nest(conds, NAMECHK.format(r=r) + nest(rest, '"OK"') + ')')
    put(ws, f'K{r}', f'=IF($C{r}="","",{reason})', fl=F_FX)
ws.freeze_panes = 'D8'
dv = DataValidation(type='list', formula1='"登録,変更,削除"', allow_blank=True); dv.add(f'B{FIRST}:B{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='APPG_NAMES', allow_blank=True, showErrorMessage=False); dv.add(f'C{FIRST}:C{LAST}'); ws.add_data_validation(dv)
cf(ws, f'J{FIRST}:K{LAST}', f'$J{FIRST}="NG"', 'FFC7CE')
cf(ws, f'J{FIRST}:K{LAST}', f'$J{FIRST}="警告"', 'FFEB9C')
cf(ws, f'J{FIRST}:K{LAST}', f'$J{FIRST}="OK"', 'C6EFCE')
cf(ws, f'D{FIRST}:D{LAST}', f'AND($B{FIRST}="変更",$E{FIRST}="あり",$M{FIRST}<>$G{FIRST})', 'F8CBAD')
for r in range(FIRST, LAST + 1):
    ws[f'D{r}'].alignment = Alignment(wrap_text=True, vertical='top')
grp_in = [
    ('登録', 'grp-new-hr', 'corp-hr\nweb-browsing\nssl'),
    ('変更', 'grp-corp-apps', 'corp-erp,corp-chat,corp-web-portal,corp-hr,zoom'),
    ('登録', 'grp-saas', 'ms-teams,office-suite-x'),
    ('登録', 'grp-ng-missing', 'corp-erp,corp-crm'),
    ('登録', 'grp-ng-self', 'grp-ng-self,ssl'),
    ('登録', 'grp-ng-dup', 'ssl,ssl'),
]
for i, v in enumerate(grp_in):
    for j, x in enumerate(v):
        ws[f'{CL(2+j)}{FIRST+i}'] = x

# =========================================================
# AppRef_定義済み・照会（定義済みApp-ID・フィルタ・名前照会の結果）
# =========================================================
ws = sheet(RF, 5)
header(ws, 'AppRef 定義済みApp-ID・アプリケーションフィルタ・名前照会の結果',
       'Application・ApplicationGroup・アプリケーションフィルタ・定義済みApp-IDは名前が共通のため、③⑦の登録名がここにある名前と同じならNGにします（大文字・小文字の違いも不可）。⑦のメンバーの存在確認にも使います。'
       '「定義済み」行は手で追加できます（SCM画面のアプリケーション一覧やApplipediaから名前をC列に貼り付け、B列は「定義済み」）。'
       '「既存情報取得（API）」で「フィルタ」行（とAPIが返した定義済みApp-ID）を取り直し、「名前照会（API）」で③⑦の登録名と⑦の不明メンバーをSCMに問い合わせて「照会:あり／照会:なし」行を追加・更新します。', 4, 70)
stat(ws, ('A3', 'B3'), '定義済み', f'=COUNTIF(B{FIRST}:B{REFL},"定義済み")')
stat(ws, ('A4', 'B4'), 'フィルタ', f'=COUNTIF(B{FIRST}:B{REFL},"フィルタ")')
stat(ws, ('A5', 'B5'), '照会:あり / なし', f'=COUNTIF(B{FIRST}:B{REFL},"照会:あり")&" / "&COUNTIF(B{FIRST}:B{REFL},"照会:なし")')
btn(ws, 'D3:D4', '▶ 名前照会（API）')
heads(ws, 7, ['No', '区分', 'name', '補足（取得元・照会日時）'], [6, 12, 32, 48])
PREDEF = ['web-browsing', 'ssl', 'dns', 'ping', 'icmp', 'traceroute', 'ssh', 'telnet', 'ftp', 'smtp', 'pop3', 'imap', 'ntp',
          'snmp', 'syslog', 'ldap', 'kerberos', 'ms-rdp', 'msrpc', 'ms-ds-smb', 'netbios-ns', 'netbios-ss', 'dhcp',
          'http-proxy', 'ike', 'ipsec-esp', 'sip', 'rtsp', 'google-base', 'gmail-base', 'youtube-base', 'facebook-base',
          'linkedin-base', 'dropbox-base', 'box-base', 'salesforce-base', 'ms-teams', 'ms-office365-base',
          'sharepoint-online', 'outlook-web-online', 'zoom', 'webex-base', 'slack-base', 'github-base', 'windows-update',
          'apple-update', 'unknown-tcp', 'unknown-udp', 'unknown-p2p', 'incomplete', 'insufficient-data', 'not-applicable']
ref_rows = [('定義済み', n, '初期値（主要App-IDの一部。全件は画面等から貼り付け）') for n in PREDEF]
ref_rows += [('フィルタ', 'flt-high-risk', 'API取得（application-filters）'),
             ('照会:なし', 'corp-hr', '名前照会 2026/10/26 13:00 applications / application-groups / application-filters に無し'),
             ('照会:なし', 'grp-new-hr', '名前照会 2026/10/26 13:00 applications / application-groups / application-filters に無し'),
             ('照会:なし', 'corp-crm', '名前照会 2026/10/26 13:00 applications / application-groups / application-filters に無し'),
             ('照会:なし', 'grp-ng-missing', '名前照会 2026/10/26 13:00 applications / application-groups / application-filters に無し')]
for i, (b, c, d) in enumerate(ref_rows):
    r = FIRST + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)')
    put(ws, f'B{r}', b, nf='@'); put(ws, f'C{r}', c, nf='@'); put(ws, f'D{r}', d, nf='@')
dv = DataValidation(type='list', formula1='"定義済み,フィルタ,照会:あり,照会:なし"', allow_blank=True); dv.add(f'B{FIRST}:B{REFL}'); ws.add_data_validation(dv)
cf(ws, f'A{FIRST}:D{REFL}', f'$B{FIRST}="照会:なし"', 'EDEDED')
ws.freeze_panes = 'C8'

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

def dtext(v):   # VBA の DefaultText と同じ表示
    if v['H'] or v['I']:
        return ' '.join(p + '/' + v[c] for p, c in (('tcp', 'H'), ('udp', 'I')) if v[c])
    if v['J']: return 'ip-protocol=' + v['J']
    if v['K']: return 'icmp type=' + v['K'] + (' code=' + v['L'] if v['L'] else '')
    if v['M']: return 'icmp6 type=' + v['M'] + (' code=' + v['N'] if v['N'] else '')
    return ''
TMO = ['timeout', 'tcp_timeout', 'udp_timeout', 'tcp_half_closed_timeout', 'tcp_time_wait_timeout']
def ttext(v): return ' '.join(f'{k}={v[c]}' for k, c in zip(TMO, 'QRSTU') if v[c])
def ftext(v): return ','.join(k for (k, _), c in zip(FLAGS, FC[18:]) if v[c] == '○')
def drow(v): return [v['D'], v['E'], v['F'], v['G'], dtext(v), v['O'], v['P'], ttext(v), ftext(v)]
DF = ['category', 'subcategory', 'technology', 'risk', 'default', 'description', 'parent_app', 'timeout', '特性・オプション']
def describe(v):
    return '\n'.join(f'{n}: {x}' for n, x in zip(DF, drow(v)) if x)
E = {nm: v for nm, v, _ in app_exist}
I = {nm: v for _, nm, v in app_in}
rows = []
for nm, v, _ in app_exist:
    if nm == 'corp-legacy-ftp': continue
    if nm in ('corp-chat', 'corp-web-portal'):
        a, b = I[nm], v
        chg = ' '.join(n for n, x, y in zip(DF, drow(a), drow(b)) if x != y)
        rows.append((nm, '変更', *drow(a), chg, describe(b), '変更', '200', '想定どおり'))
    else:
        rows.append((nm, '変化なし', *drow(v), '', '', '', '', ''))
for nm in ('corp-hr', 'corp-iot'):
    rows.append((nm, '追加', *drow(I[nm]), '', '', '登録', '201', '想定どおり'))
rows.append(('corp-legacy-ftp', '削除', *drow(E['corp-legacy-ftp']), '', '', '削除', '200', '想定どおり'))
diffsheet(AD, 7, '⑧ 差分確認（Application）：現在の設定一覧と既存一覧の比較', DF,
          [16, 16, 14, 6, 18, 22, 12, 16, 24], rows, 'J3:K4')
diffsheet(GD, 8, '⑨ 差分確認（ApplicationGroup）：現在の設定一覧と既存一覧の比較', ['メンバー'], [60], [
    ('grp-corp-apps', '変更', 'corp-erp,corp-chat,corp-web-portal,corp-hr,zoom', 'メンバー',
     'メンバー: corp-erp,corp-chat,corp-web-portal', '変更', '200', '想定どおり'),
    ('grp-transfer', '変更', 'corp-backup,ftp', 'メンバー', 'メンバー: corp-backup,corp-legacy-ftp,ftp', '変更', '200', '想定どおり'),
    ('grp-all-corp', '変化なし', 'grp-corp-apps,ssl,web-browsing', '', '', '', '', ''),
    ('grp-new-hr', '追加', 'corp-hr,web-browsing,ssl', '', '', '登録', '201', '想定どおり'),
    ('grp-saas', '追加', 'ms-teams,office-suite-x', '', '', '登録', '201', '想定どおり'),
], 'J3:K4')

# ④ JSON_CURL サンプル（VBAのJSON発行と同じ形式）
BASE = 'https://api.sase.paloaltonetworks.com'
def jstr(v): return _j.dumps(v, ensure_ascii=False)
def ports(proto, lst):   # 1要素63文字以内に分割
    res, cur = [], ''
    for p in lst.split(','):
        nxt = p if cur == '' else cur + ',' + p
        if len(proto) + 1 + len(nxt) > 63 and cur:
            res.append(proto + '/' + cur); cur = p
        else:
            cur = nxt
    if cur: res.append(proto + '/' + cur)
    return res
def appbody(nm, v, extra=''):
    b = '{"name":' + jstr(nm) + ',"category":' + jstr(v['D']) + ',"subcategory":' + jstr(v['E']) + \
        ',"technology":' + jstr(v['F']) + ',"risk":' + v['G']
    if v['O']: b += ',"description":' + jstr(v['O'])
    if v['P']: b += ',"parent_app":' + jstr(v['P'])
    if v['H'] or v['I']:
        pl = (ports('tcp', v['H']) if v['H'] else []) + (ports('udp', v['I']) if v['I'] else [])
        b += ',"default":{"port":[' + ','.join(jstr(p) for p in pl) + ']}'
    elif v['J']:
        b += ',"default":{"ident_by_ip_protocol":' + jstr(v['J']) + '}'
    for k, c in zip(TMO, 'QRSTU'):
        if v[c]: b += f',"{k}":{int(v[c])}'
    for (k, _), c in zip(FLAGS, FC[18:]):
        if v[c] == '○': b += f',"{k}":true'
    if extra: b += ',' + extra
    return b + '}'
def grpbody(nm, mem):
    return '{"name":' + jstr(nm) + ',"members":[' + ','.join(jstr(m) for m in mem.split(',')) + ']}'
def curl(m, u, b):
    c = f"curl -s -X {m} '{u}' -H \"Authorization: Bearer $TOKEN\""
    if b: c += " -H 'Content-Type: application/json' -d '" + b.replace("'", "'\\''") + "'"
    return c
A_, G_ = BASE + '/sse/config/v1/applications', BASE + '/sse/config/v1/application-groups'
jrows = [
    ('登録', 'corp-hr', 'POST', A_ + '?folder=Shared', appbody('corp-hr', I['corp-hr']), 201),
    ('登録', 'corp-iot', 'POST', A_ + '?folder=Shared', appbody('corp-iot', I['corp-iot']), 201),
    ('変更', 'corp-chat', 'PUT', A_ + '/7a3c1e50-0002-4b2e-9c11-000000000002', appbody('corp-chat', I['corp-chat']), 200),
    ('変更', 'corp-web-portal', 'PUT', A_ + '/7a3c1e50-0006-4b2e-9c11-000000000006', appbody('corp-web-portal', I['corp-web-portal'], SIG), 200),
    ('登録(Group)', 'grp-new-hr', 'POST', G_ + '?folder=Shared', grpbody('grp-new-hr', 'corp-hr,web-browsing,ssl'), 201),
    ('変更(Group)', 'grp-corp-apps', 'PUT', G_ + '/8b4d2f60-0001-4b2e-9c11-000000000011', grpbody('grp-corp-apps', 'corp-erp,corp-chat,corp-web-portal,corp-hr,zoom'), 200),
    ('登録(Group)', 'grp-saas', 'POST', G_ + '?folder=Shared', grpbody('grp-saas', 'ms-teams,office-suite-x'), 201),
    ('変更(Group)自動', 'grp-transfer', 'PUT', G_ + '/8b4d2f60-0002-4b2e-9c11-000000000012', grpbody('grp-transfer', 'corp-backup,ftp'), 200),
    ('削除', 'corp-legacy-ftp', 'DELETE', A_ + '/7a3c1e50-0004-4b2e-9c11-000000000004', '', 200),
]
js = wb[SJ]
for i, (op, nm, m, u, b, st) in enumerate(jrows):
    r = 10 + i
    for j, v in enumerate([i + 1, op, nm, m, u, b or None, curl(m, u, b), st]):
        put(js, f'{CL(1+j)}{r}', v, al=Alignment(wrap_text=True, vertical='top'))

# 名前定義
wb.defined_names['EX_FETCHED_AT'] = DefinedName('EX_FETCHED_AT', attr_text=f"{Q(AE)}!$B$3")
wb.defined_names['EXG_FETCHED_AT'] = DefinedName('EXG_FETCHED_AT', attr_text=f"{Q(GE)}!$B$3")
wb.defined_names['APP_NAMES'] = DefinedName('APP_NAMES', attr_text=f"OFFSET({Q(AE)}!$C$8,0,0,MAX(1,COUNTA({Q(AE)}!$C$8:$C$5000)),1)")
wb.defined_names['APPG_NAMES'] = DefinedName('APPG_NAMES', attr_text=f"OFFSET({Q(GE)}!$C$8,0,0,MAX(1,COUNTA({Q(GE)}!$C$8:$C$5000)),1)")
wb.defined_names['APP_EX_TBL'] = DefinedName('APP_EX_TBL', attr_text=f"{Q(AE)}!$A$8:$AK$5000")
wb[AE]['B3'] = datetime.datetime(2026, 10, 26, 13, 1)
wb[GE]['B3'] = datetime.datetime(2026, 10, 26, 13, 1)

order = ['基本設定', AE, AI, GE, GI, RF, CH, SJ, AD, GD, SL]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for w in wb: w.sheet_view.tabSelected = (w.title == '基本設定')
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('saved', OUT)

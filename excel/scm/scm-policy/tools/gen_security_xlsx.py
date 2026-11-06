# -*- coding: utf-8 -*-
# SCM セキュリティポリシー（security-rules）Excel 生成（復号ポリシーツールと同じ構成）
#   Tag ツールのブックから 基本設定 / JSON_CURL(→JSON化) / 実行ログ を流用
import sys, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

SRC = '/mnt/project-files/scm-tag/SCM_Tag.xlsx'
OUT = sys.argv[1] if len(sys.argv) > 1 else '/mnt/project-files/scm-policy/SCM_SecurityPolicy.xlsx'

SG, SS, SJ, SD, SL, SP, SC = 'SecRule_取得', 'SecRule_設定', 'JSON化', 'SecRule_設定差分チェック', '実行ログ', 'ProfGroup_一覧', 'SecRule_選択肢'
FIRST, SLAST = 8, 3000

FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
F_TITLE, F_HEAD, F_IN, F_FX, F_VBA, F_BTN = fill('1F3864'), fill('2F5597'), fill('FFF2CC'), fill('F2F2F2'), fill('E2EFDA'), fill('C55A11')
F_HDR, F_ADD, F_DEL, F_CELL = fill('1F3864'), fill('C6EFCE'), fill('D9D9D9'), fill('FFC000')
F_REQ, F_OPT, F_SEC = fill('C00000'), fill('7F7F7F'), fill('D9E1F2')

# 項目（API キー, 見出し, 種類 a=配列 b=真偽 s=文字列 t=種類 c=証明書, 幅）
FIELDS = [
    ('disabled', '無効\n(disabled)', 'b', 7), ('tag', 'タグ\n(tag)', 'a', 12),
    ('from', '送信元ゾーン\n(from)', 'a', 12), ('source', '送信元アドレス\n(source)', 'a', 20),
    ('negate_source', '送信元否定\n(negate_source)', 'b', 8), ('source_user', '送信元ユーザー\n(source_user)', 'a', 20),
    ('source_hip', '送信元HIP\n(source_hip)', 'a', 12), ('to', '宛先ゾーン\n(to)', 'a', 12),
    ('destination', '宛先アドレス\n(destination)', 'a', 20), ('negate_destination', '宛先否定\n(negate_destination)', 'b', 8),
    ('destination_hip', '宛先HIP\n(destination_hip)', 'a', 12), ('application', 'アプリケーション\n(application)', 'a', 20),
    ('service', 'サービス\n(service)', 'a', 14), ('category', 'URLカテゴリ\n(category)', 'a', 22),
    ('action', 'アクション\n(action)', 's', 11), ('profile_setting', 'プロファイルグループ\n(profile_setting)', 'g', 18),
    ('log_setting', 'ログ転送\n(log_setting)', 's', 16), ('description', '説明\n(description)', 's', 28),
]
NF = len(FIELDS)
PICK = {'tag', 'from', 'source', 'source_user', 'source_hip', 'to', 'destination', 'destination_hip', 'application', 'service', 'category'}
REQ = ('from', 'source', 'source_user', 'to', 'destination', 'application', 'service', 'category', 'action')
GLAST = 8 + NF + 1; GHF = GLAST + 2      # ② 取得時JSON列 / 見出し行の folder・position（非表示）
SLASTC = 11 + NF; SHF = SLASTC + 2       # ③
DLAST = 7 + NF + 5; DHF = DLAST + 2      # ⑤

SCOPES = [('Prisma Access', 'Shared', 'Prisma Access 全体（共通）'),
          ('Remote Networks', 'Remote Networks', 'リモートネットワーク'),
          ('Service Connections', 'Service Connections', 'サービスコネクション'),
          ('Mobile Users Container', 'Mobile Users Container', 'GlobalProtect と Explicit Proxy の共通'),
          ('GlobalProtect', 'Mobile Users', 'モバイルユーザー（GlobalProtect）'),
          ('Explicit Proxy', 'Mobile Users Explicit Proxy', 'モバイルユーザー（明示型プロキシ）')]
def label(folder):
    for l, f, _ in SCOPES:
        if f == folder: return f if l == f else f'{l}（{f}）'
    return folder

PREDEF_URL = ['abortion', 'abused-drugs', 'adult', 'alcohol-and-tobacco', 'artificial-intelligence', 'auctions', 'business-and-economy',
    'command-and-control', 'compromised-website', 'computer-and-internet-info', 'content-delivery-networks', 'copyright-infringement',
    'cryptocurrency', 'dating', 'dynamic-dns', 'educational-institutions', 'encrypted-dns', 'entertainment-and-arts', 'extremism',
    'financial-services', 'gambling', 'games', 'government', 'grayware', 'hacking', 'health-and-medicine', 'high-risk',
    'home-and-garden', 'hunting-and-fishing', 'insufficient-content', 'internet-communications-and-telephony', 'internet-portals',
    'job-search', 'legal', 'low-risk', 'malware', 'medium-risk', 'military', 'motor-vehicles', 'music', 'newly-registered-domain',
    'news', 'not-resolved', 'nudity', 'online-storage-and-backup', 'parked', 'peer-to-peer', 'personal-sites-and-blogs',
    'philosophy-and-political-advocacy', 'phishing', 'private-ip-addresses', 'proxy-avoidance-and-anonymizers', 'questionable',
    'ransomware', 'real-estate', 'real-time-detection', 'recreation-and-hobbies', 'reference-and-research', 'religion',
    'remote-access', 'scanning-activity', 'search-engines', 'sex-education', 'shareware-and-freeware', 'shopping',
    'social-networking', 'society', 'sports', 'stock-advice-and-tools', 'streaming-media', 'swimsuits-and-intimate-apparel',
    'training-and-tools', 'translation', 'travel', 'unknown', 'weapons', 'web-advertisements', 'web-based-email', 'web-hosting']

wb = openpyxl.load_workbook(SRC)
for n in ['Tag_既存一覧', 'Tag_登録変更削除', 'Tag_差分確認', 'Tag_色一覧']:
    del wb[n]
for n in ['EX_FETCHED_AT', 'TAG_NAMES', 'TAG_COLORS', 'CFG_CONT_TYPE', 'CFG_CONT_NAME']:
    if n in wb.defined_names: del wb.defined_names[n]

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

def header(ws, title, desc, lastcol, h2=48):
    lc = CL(lastcol)
    ws.merge_cells(f'A1:{lc}1'); ws.merge_cells(f'A2:{lc}2')
    put(ws, 'A1', title, Font(name=FN, size=14, bold=True, color='FFFFFF'), F_TITLE, Alignment(vertical='center'), False)
    put(ws, 'A2', desc, Font(name=FN, size=9, color='595959'), None, Alignment(vertical='center', wrap_text=True), False)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = h2

def btn(ws, rng, text):
    ws.merge_cells(rng)
    put(ws, rng.split(':')[0], text, Font(name=FN, size=10, bold=True, color='FFFFFF'), F_BTN,
        Alignment(horizontal='center', vertical='center', wrap_text=True), False)

def heads(ws, row, names, widths, start=1):
    for i, (n, w) in enumerate(zip(names, widths), start):
        put(ws, f'{CL(i)}{row}', n, Font(name=FN, size=9, bold=True, color='FFFFFF'), F_HEAD,
            Alignment(horizontal='center', vertical='center', wrap_text=True))
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[row].height = 40

def stat(ws, addr, lab, formula, fl=F_FX):
    put(ws, addr[0], lab, Font(name=FN, size=10, bold=True))
    put(ws, addr[1], formula, Font(name=FN, size=10, bold=True), fl)

def note(ws, addr, text, color='595959'):
    put(ws, addr, text, Font(name=FN, size=9, color=color), border=False)

def defname(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)

# =========================================================
# ① 基本設定
# =========================================================
cfg = wb['基本設定']
cfg['B7'] = '/sse/config/v1/security-rules'
cfg['D7'] = 'セキュリティポリシー（Security Rules）'
cfg['A14'] = '取得・登録の条件'
cfg['A15'] = 'position（Pre/Post）'
cfg['B15'] = '両方'
cfg['D15'] = '両方（既定）/ pre / post。両方は Pre をすべて（親→子）並べたあとに Post（子→親）を並べる'
cfg['A16'] = 'オブジェクト取得folder'
cfg['B16'] = 'Shared'
cfg['C16'].value = '必須'
cfg['D16'] = '選択肢（アドレス・サービス・アプリ等）を取得する基本のfolder。このほか下の○のスコープからも取得し、定義folderをスコープとして記録'
for col in 'ABCD':
    cfg[f'{col}17']._style = copy.copy(cfg[f'{col}16']._style)
cfg['A17'] = 'カスタムApp取得'; cfg['B17'] = 'する'; cfg['C17'] = '任意'
cfg['D17'] = 'する=選択肢の取得時に applications を読み、定義済み（約11万件）は読み飛ばして自作Appだけを候補にする（件数が多いと時間がかかります）／しない=取得しない'
cfg['B19'] = None; cfg['B20'] = None; cfg['B21'] = None
keep = []
for dv in cfg.data_validations.dataValidation:
    if str(dv.sqref) not in ('B15', 'B16', 'B17'): keep.append(dv)
cfg.data_validations.dataValidation = keep
dv = DataValidation(type='list', formula1='"pre,post,両方"', allow_blank=False); dv.add('B15'); cfg.add_data_validation(dv)
dv = DataValidation(type='list', formula1='"Shared,Mobile Users,Remote Networks,Service Connections,Mobile Users Container,Mobile Users Explicit Proxy"', allow_blank=False); dv.add('B16'); cfg.add_data_validation(dv)
defname('CFG_POSITION', "'基本設定'!$B$15")
defname('CFG_OBJ_FOLDER', "'基本設定'!$B$16")
dv = DataValidation(type='list', formula1='"する,しない"', allow_blank=False); dv.add('B17'); cfg.add_data_validation(dv)
defname('CFG_APP_FETCH', "'基本設定'!$B$17")
# 1回の取得件数
for col in 'ABCD':
    cfg[f'{col}27']._style = copy.copy(cfg[f'{col}26']._style)
cfg['A27'] = '1回の取得件数'; cfg['B27'] = 200; cfg['C27'] = '任意'; cfg['D27'] = 'ページ送り1回の件数（既定 200）'
cfg.merge_cells('D27:E27')
defname('CFG_PAGE_LIMIT', "'基本設定'!$B$27")
# 取得・登録の対象スコープ
cfg['A36']._style = copy.copy(cfg['A28']._style)
cfg['A36'] = '取得・登録の対象スコープ：○を付けたものを上から順に取得し、スコープごとに表化（ポリシーの1行目にスコープ名を表示）'
for col in 'BCDE':
    cfg[f'{col}36']._style = copy.copy(cfg['A28']._style)
for i, (h, src) in enumerate(zip(['スコープ（表示名）', '対象', 'API folder名', '説明'], 'ABCD')):
    put(cfg, f'{src}37', h, Font(name=FN, size=9, bold=True, color='FFFFFF'), F_HEAD, Alignment(horizontal='center'))
cfg.merge_cells('D37:E37')
for i, (l, f, d) in enumerate(SCOPES):
    r = 38 + i
    put(cfg, f'A{r}', l, Font(name=FN, size=10, bold=True))
    put(cfg, f'B{r}', '○', Font(name=FN, size=10, color='0000FF'), F_IN, Alignment(horizontal='center'))
    put(cfg, f'C{r}', f, Font(name=FN, size=9), F_FX)
    put(cfg, f'D{r}', d, Font(name=FN, size=9, color='595959'), border=False)
    cfg.merge_cells(f'D{r}:E{r}')
dv = DataValidation(type='list', formula1='"○"', allow_blank=True); dv.add('B38:B43'); cfg.add_data_validation(dv)
defname('CFG_SCOPES', "'基本設定'!$A$38:$C$43")
note(cfg, 'A44', '※ API folder名は変更しないでください。グローバル（All）のルールは Prisma Access の応答に含まれる場合に「参照のみ」で表示します。')
note(cfg, 'A45', '※ 子folderの応答に含まれる上位folderのルール（継承分）は、定義元のスコープにだけ表示します。')
cfg.column_dimensions['C'].width = 26

# =========================================================
# ④ JSON化（JSON_CURL を流用）
# =========================================================
js = wb['JSON_CURL']
js.title = SJ
js['A1'] = '④ JSON化（生成JSON・CURLコマンド）'
js['A2'] = ('③設定の「JSON化」ボタンでVBAが10行目以降を生成します。並び：①変更(PUT) → ②削除(DELETE) → ③スコープごとに上から 追加(POST)＋位置へ移動(:move) / 移動。'
            '追加の位置は、③設定で直前にある（削除しない）ルールの後ろ（直前が無ければ先頭 top）です。'
            '{{id:name}} は「API実行」時に、直前の追加（POST）で作られたルールのidに置き換えます（curlを手で実行する場合は置き換えてください）。'
            '実行後は⑤設定差分チェックに結果が出ます。')
js.row_dimensions[2].height = 60
for row in js.iter_rows(min_row=10, max_row=js.max_row):
    for c in row: c.value = None
for i, (h, w) in enumerate([('スコープ\n(folder|position)', 26), ('対象id', 22), ('移動先', 22), ('追加後のid', 22)]):
    col = CL(9 + i)
    js[f'{col}9']._style = copy.copy(js['H9']._style)
    js[f'{col}9'] = h
    js.column_dimensions[col].width = w
js['B9'] = '操作'

lg = wb[SL]
lg['A1'] = '⑥ 実行ログ（API実行結果・Push・ジョブ状況）'
for row in lg.iter_rows(min_row=6, max_row=lg.max_row):
    for c in row:
        c.value = None; c.border = Border(); c.number_format = 'General'

# =========================================================
# サンプル（取得したときの見え方）
# =========================================================
def rule(name, src='any', dst='any', app='any', svc='application-default', cat='any', action='allow', prof='best-practice', desc='', tag='',
         user='any', frm='trust', to='untrust', dis='', lset='Cortex Data Lake'):
    v = {'disabled': dis, 'tag': tag, 'from': frm, 'source': src, 'negate_source': '', 'source_user': user, 'source_hip': '',
         'to': to, 'destination': dst, 'negate_destination': '', 'destination_hip': '', 'application': app, 'service': svc,
         'category': cat, 'action': action, 'profile_setting': prof if action == 'allow' else '', 'log_setting': lset, 'description': desc}
    return [name] + [v[k] for k, *_ in FIELDS]

SAMPLE = [   # Pre をすべて（親→子）→ Post をすべて（子→親）
    ('Shared', 'pre', [
        ('c1d2e3f4-1111-4000-8000-000000000001', rule('Block-Malicious-URL', cat='malware\nphishing\ncommand-and-control', action='deny', desc='危険カテゴリを遮断')),
        ('c1d2e3f4-1111-4000-8000-000000000002', rule('Block-Bad-IP', dst='EDL-Block-IP', action='drop', desc='脅威IPを遮断')),
    ]),
    ('Remote Networks', 'pre', [
        ('c1d2e3f4-2222-4000-8000-000000000001', rule('RN-Allow-Branch-Server', src='RN-Branch-Servers', app='ssl\nweb-browsing', desc='拠点サーバー')),
    ]),
    ('Service Connections', 'pre', []),
    ('Mobile Users Container', 'pre', []),
    ('Mobile Users', 'pre', [
        ('c1d2e3f4-3333-4000-8000-000000000001', rule('GP-Allow-O365', user='CN=all-staff,OU=Groups,DC=example,DC=co,DC=jp', app='ms-office365-base\nms-teams', desc='Microsoft 365')),
    ]),
    ('Mobile Users Explicit Proxy', 'pre', [
        ('c1d2e3f4-4444-4000-8000-000000000001', rule('EP-Allow-Web', app='web-browsing\nssl', desc='明示型プロキシ')),
    ]),
    ('Mobile Users Explicit Proxy', 'post', []),
    ('Mobile Users', 'post', [
        ('c1d2e3f4-3333-4000-8000-000000000002', rule('GP-Allow-Web', app='web-browsing\nssl', desc='GPユーザーのWeb通信')),
    ]),
    ('Mobile Users Container', 'post', []),
    ('Service Connections', 'post', []),
    ('Remote Networks', 'post', []),
    ('Shared', 'post', [
        ('c1d2e3f4-1111-4000-8000-000000000003', rule('Deny-All', app='any', svc='any', action='deny', desc='最後にすべて拒否')),
    ]),
]
DEPTH = {'All': 0, 'Shared': 1, 'Remote Networks': 2, 'Service Connections': 2, 'Mobile Users Container': 2, 'Mobile Users': 3, 'Mobile Users Explicit Proxy': 3}
DCOL = {0: '404040', 1: '1F3864', 2: '2F5A97', 3: '4472C4'}
def lvl(folder): return max(DEPTH[folder] - 1, 0)   # 0=第1階層 1=第2階層 2=第3階層
SCOPE_HEADS = ['スコープ\n第1階層\n(Prisma Access)', 'スコープ\n第2階層\n(RN / SC / MUC)', 'スコープ\n第3階層\n(GP / EP)']
SCOPE_W = [16, 18, 18]
def glabel(folder, pos): return f"{label(folder)} {'【Post】' if pos == 'post' else '【Pre】'}"


def write_header_row(ws, r, labcol, lastcol, hfcol, folder, pos, n):
    ws.cell(r, 1, '▼')
    ws.cell(r, labcol + lvl(folder), f'■ {glabel(folder, pos)} ／ {n} 件')
    ws.cell(r, hfcol, folder); ws.cell(r, hfcol + 1, pos)
    for c in range(1, hfcol + 2):
        cell = ws.cell(r, c)
        cell.font = Font(name=FN, size=10, bold=True, color='FFFFFF')
        if c <= lastcol: cell.fill = fill(DCOL[DEPTH[folder]])
        cell.number_format = '@'
    ws.cell(r, labcol + lvl(folder)).alignment = Alignment(horizontal='left')

def data_cell(ws, r, c, v, fl=None, strike=False, bold=False):
    cell = ws.cell(r, c, v if v != '' else None)
    cell.font = Font(name=FN, size=10, strike=strike, bold=bold)
    cell.alignment = Alignment(wrap_text=True, vertical='top')
    cell.border = BOX
    cell.number_format = '@'
    if fl: cell.fill = fl
    return cell

# =========================================================
# ② SecRule_取得
# =========================================================
ws = sheet(SG, 1)
header(ws, '② 取得及び表化（現在のセキュリティポリシー）',
       '「まとめて取得」で、基本設定の○のスコープ（Prisma Access / Remote Networks / Service Connections / Mobile Users Container / GlobalProtect / Explicit Proxy）を Pre・Post とも取得し、'
       'Pre をすべて（Prisma Access → RN → SC → MUC → GP → EP）並べたあとに、Post を評価順（EP → GP → MUC → SC → RN → Prisma Access）に並べて1つの表にします。スコープは階層ごとに3列（第1階層 Prisma Access／第2階層 RN・SC・MUC／第3階層 GP・EP）に分けて段ずらしし、見出し（▼の行）は階層が深いほど明るい青です。'
       'Remote Networks・Service Connections・GlobalProtect・Explicit Proxy は別々の枝で、それぞれの通信には自分の枝と上位（Prisma Access・グローバル）のルールだけが効きます。'
       '同時にプロファイルグループ一覧と選択肢（アドレス・アプリ・サービス・URLカテゴリ等）も取得します。このシートは編集しないでください（③設定と差分チェックの基準＝実行前の状態）。'
       '現在はサンプル行が入っています（取得すると置き換わります）。', GLAST, 60)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), 'position', '=CFG_POSITION')
btn(ws, 'D3:F4', '▶ まとめて取得（ポリシー・プロファイル・選択肢）')
btn(ws, 'H3:I4', '▶ ポリシーのみ取得')
note(ws, 'K3', '順番 = スコープ内のルール順（APIの返却順）。配列の項目はセル内改行で1行1件。')
gh = ['順番'] + SCOPE_HEADS + ['folder', 'position', 'id', 'name'] + [f[1] for f in FIELDS] + ['取得時JSON']
gw = [5] + SCOPE_W + [14, 7, 20, 26] + [f[3] for f in FIELDS] + [30]
heads(ws, 7, gh, gw)
ws.column_dimensions[CL(GHF)].hidden = True; ws.column_dimensions[CL(GHF + 1)].hidden = True
r = FIRST
get_rows = []   # (r, folder, id, vals, seq) for ③
for folder, pos, rules in SAMPLE:
    write_header_row(ws, r, 2, GLAST, GHF, folder, pos, len(rules)); get_rows.append(('H', folder, pos, len(rules))); r += 1
    for n, (rid, vals) in enumerate(rules, 1):
        data_cell(ws, r, 1, n)
        for k in range(3):
            data_cell(ws, r, 2 + k, glabel(folder, pos) if k == lvl(folder) else None, bold=(n == 1))
        data_cell(ws, r, 5, folder); data_cell(ws, r, 6, pos); data_cell(ws, r, 7, rid)
        for i, v in enumerate(vals):
            data_cell(ws, r, 8 + i, v)
        data_cell(ws, r, GLAST, '{"id":"%s","name":"%s",...}' % (rid, vals[0]))
        ws.cell(r, 1).number_format = 'General'
        get_rows.append(('R', folder, rid, vals, n, pos)); r += 1
ws.freeze_panes = 'I8'

# =========================================================
# ③ SecRule_設定
# =========================================================
ws = sheet(SS, 2)
header(ws, '③ 設定（今あるポリシーに 追加・削除・変更 の操作を付けて編集）',
       '「取得結果を展開」で②の表をこのシートに作り、黄色セルを編集します。変更=操作「変更」を選んでセルを書き換え／削除=操作「削除」／'
       '追加=追加したい位置の1つ上の行（スコープの先頭なら ▼の行）を選んで「選択行の下に追加行を挿入」→ name 等を入力。'
       '並べ替え=行を切り取り→挿入して操作「移動」（内容も変えるなら「変更+移動」）。'
       'タグ・ゾーン・アドレス・ユーザー・HIP・サービス・URLカテゴリはダブルクリック（または Ctrl+Shift+M）でチェックボックス選択。'
       '「入力チェック」で C・D列に判定と理由、取得時から変わったセルを橙色で表示します。追加は登録（POST）後に、その位置へ移動（:move）するところまで④に出力します。', SLASTC, 72)
stat(ws, ('A3', 'B3'), '操作あり', f'=COUNTA(B{FIRST}:B{SLAST})')
stat(ws, ('C3', 'D3'), 'NG', f'=COUNTIF(C{FIRST}:C{SLAST},"NG")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(C{FIRST}:C{SLAST},"警告")')
stat(ws, ('C4', 'D4'), 'OK', f'=COUNTIF(C{FIRST}:C{SLAST},"OK")')
btn(ws, 'F3:H4', '▶ 取得結果を展開（作り直し）')
btn(ws, 'J3:K4', '▶ 選択行の下に追加行を挿入')
btn(ws, 'M3:N4', '▶ 入力チェック')
btn(ws, 'P3:Q4', '▶ JSON化 → ④へ')
note(ws, 'S3', '色：黄=入力 ／ 緑=追加の行 ／ グレー＋取り消し線=削除の行 ／ 橙=取得時から変わったセル ／ 濃紺=スコープ見出し')
note(ws, 'S4', '操作：追加 / 変更 / 削除 / 移動 / 変更+移動（空欄の行は変更しません）')
for col in range(1, SLASTC + 1):
    lab = None
    if col == 2: lab = '必須'
    elif col == 11: lab = '必須'
    elif col >= 12:
        k = FIELDS[col - 12][0]
        if k in REQ: lab = '必須'
        elif k in PICK: lab = None
    if lab:
        put(ws, f'{CL(col)}6', lab, Font(name=FN, size=9, bold=True, color='FFFFFF'), F_REQ, Alignment(horizontal='center'), False)
    if col >= 12 and FIELDS[col - 12][0] in PICK:
        put(ws, f'{CL(col)}5', 'ダブルクリック', Font(name=FN, size=8, bold=True, color='FFFFFF'), fill('548235'), Alignment(horizontal='center'), False)
sh_ = ['順番', '操作', '判定', '理由'] + SCOPE_HEADS + ['folder', 'position', 'id', 'name\n(63文字以内)'] + [f[1] for f in FIELDS]
sw = [5, 10, 6, 30] + SCOPE_W + [12, 7, 12, 26] + [f[3] for f in FIELDS]
heads(ws, 7, sh_, sw)
for c in (8, 9, 10, SHF, SHF + 1):
    ws.column_dimensions[CL(c)].hidden = True

SET_SAMPLE_OPS = {  # id -> (操作, 変更内容)
    'c1d2e3f4-1111-4000-8000-000000000002': ('変更', {'description': '脅威IPを遮断（パートナー提供リストを追加）', 'destination': 'EDL-Block-IP\nEDL-Block-IP-Partner'}),
    'c1d2e3f4-2222-4000-8000-000000000001': ('削除', {}),
}
r = FIRST
fidx = {f[0]: i + 1 for i, f in enumerate(FIELDS)}
def set_row(ws, r, folder, op, rid, vals, seq, changed=(), pos='pre'):
    base = F_ADD if op == '追加' else (F_DEL if op == '削除' else None)
    data_cell(ws, r, 1, seq, F_FX); ws.cell(r, 1).number_format = 'General'
    data_cell(ws, r, 2, op, base or F_IN)
    data_cell(ws, r, 3, 'OK' if op else None, F_ADD if op else F_FX)
    data_cell(ws, r, 4, None, F_FX)
    for k in range(3):
        data_cell(ws, r, 5 + k, glabel(folder, pos) if k == lvl(folder) else None, base or F_FX, strike=(op == '削除'), bold=(seq == 1))
    data_cell(ws, r, 8, folder, base or F_FX); data_cell(ws, r, 9, pos, base or F_FX); data_cell(ws, r, 10, rid, base or F_FX)
    for i, v in enumerate(vals):
        c = data_cell(ws, r, 11 + i, v, base or F_IN, strike=(op == '削除'))
        if i in changed:
            c.fill = F_CELL; c.font = Font(name=FN, size=10, bold=True)
for row in get_rows:
    if row[0] == 'H':
        write_header_row(ws, r, 5, SLASTC, SHF, row[1], row[2], row[3]); r += 1
        continue
    _, folder, rid, vals, seq, pos = row
    op, chg = SET_SAMPLE_OPS.get(rid, ('', {}))
    vals = list(vals); changed = []
    for k, v in chg.items():
        vals[fidx[k]] = v; changed.append(fidx[k])
    set_row(ws, r, folder, op, rid, vals, seq, changed, pos); r += 1
    if rid == 'c1d2e3f4-1111-4000-8000-000000000001':   # この下に追加（2番目に入る）
        newv = rule('Block-P2P', app='bittorrent\nunknown-p2p', action='deny', desc='行挿入で追加（2番目に入る例）')
        set_row(ws, r, 'Shared', '追加', None, newv, None); r += 1
ws.freeze_panes = 'L8'
dv = DataValidation(type='list', formula1='"追加,変更,削除,移動,変更+移動"', allow_blank=True); dv.add(f'B{FIRST}:B{SLAST}'); ws.add_data_validation(dv)
bools = [CL(11 + fidx[k]) for k in ('disabled', 'negate_source', 'negate_destination')]
dv = DataValidation(type='list', formula1='"true,false"', allow_blank=True)
for c in bools: dv.add(f'{c}{FIRST}:{c}{SLAST}')
ws.add_data_validation(dv)
c = CL(11 + fidx['action']); dv = DataValidation(type='list', formula1='"allow,deny,drop,reset-client,reset-server,reset-both"', allow_blank=True); dv.add(f'{c}{FIRST}:{c}{SLAST}'); ws.add_data_validation(dv)
c = CL(11 + fidx['profile_setting']); dv = DataValidation(type='list', formula1='SEC_PROFGROUPS', allow_blank=True, showErrorMessage=False); dv.add(f'{c}{FIRST}:{c}{SLAST}'); ws.add_data_validation(dv)

# =========================================================
# ⑤ SecRule_設定差分チェック
# =========================================================
ws = sheet(SD, 4)
header(ws, '⑤ 設定差分チェック（実行前プレビュー／実行後チェック）',
       '「実行前プレビュー」= ②取得 と ③設定 の比較（APIは呼びません）。「実行後チェック」= SCMから再取得して②取得（実行前）と比較し、④の操作と③の並び順どおりかを判定します（API実行後は自動で作成）。'
       '行の色：緑=追加 ／ 黄=変更・移動 ／ グレー＋取り消し線=削除（直前にあったルールの下に表示）、変わったセルは橙色・太字、判定が想定外のときは判定セルが赤。順番は「実行前→実行後」。', DLAST, 60)
stat(ws, ('A3', 'B3'), '作成日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('C3', 'D3'), '総数', None, F_VBA); stat(ws, ('E3', 'F3'), '変化なし', None, F_VBA); stat(ws, ('G3', 'H3'), '追加', None, F_VBA)
stat(ws, ('C4', 'D4'), '変更', None, F_VBA); stat(ws, ('E4', 'F4'), '削除', None, F_VBA); stat(ws, ('G4', 'H4'), '要確認', None, F_VBA)
stat(ws, ('I3', 'J3'), '方式', None, F_VBA); stat(ws, ('I4', 'J4'), '移動', None, F_VBA)
btn(ws, 'L3:M4', '▶ 実行前プレビュー（取得⇔設定）')
btn(ws, 'O3:P4', '▶ 実行後チェック（SCM再取得）')
dh = ['No'] + SCOPE_HEADS + ['変化', '順番\n(前→後)', 'name'] + [f[1] for f in FIELDS] + ['変更項目', '変更前', '操作', 'HTTP', '判定']
dw = [5] + SCOPE_W + [10, 8, 26] + [f[3] for f in FIELDS] + [18, 30, 12, 7, 30]
heads(ws, 7, dh, dw)
ws.column_dimensions['J'].width = max(ws.column_dimensions['J'].width, 14)
for c in (DHF, DHF + 1): ws.column_dimensions[CL(c)].hidden = True
ws.freeze_panes = 'H8'

# =========================================================
# ProfGroup_一覧（プロファイルグループ）
# =========================================================
ws = sheet(SP, 99)
header(ws, '⑦ プロファイルグループ一覧（③設定の「プロファイルグループ」列のプルダウン元）',
       'ボタン（または②の「まとめて取得」）で、Shared と基本設定で○のスコープの profile-groups を取得します（同じものは1回だけ表示）。'
       'セキュリティルールに付けるのはプロファイルグループです（各列はグループに入っているセキュリティプロファイル名）。このシートは編集しないでください。', 12, 40)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '件数', f'=COUNTA(B{FIRST}:B2000)')
btn(ws, 'D3:E4', '▶ プロファイルグループ取得（API）')
heads(ws, 7, ['No', 'name', '定義folder', '種別', 'ウイルス・WildFire\n(virus_and_wildfire_analysis)', 'スパイウェア\n(spyware)',
              '脆弱性防御\n(vulnerability)', 'URLフィルタリング\n(url_filtering)', 'ファイルブロック\n(file_blocking)',
              'DNSセキュリティ\n(dns_security)', 'SaaSセキュリティ\n(saas_security)', 'id'],
      [5, 24, 14, 10, 24, 20, 20, 20, 20, 20, 18, 24])
prof = [('best-practice', 'Shared', '定義済み', 'best-practice', 'best-practice', 'best-practice', 'best-practice', 'best-practice', 'best-practice', '', ''),
        ('PG-Strict', 'Shared', 'カスタム', 'AV-Strict', 'AS-Strict', 'VP-Strict', 'URL-Strict', 'FB-Strict', 'DNS-Strict', '', 'b0b0b0b0-0000-4000-8000-000000000002')]
for i, p in enumerate(prof):
    data_cell(ws, FIRST + i, 1, i + 1); ws.cell(FIRST + i, 1).number_format = 'General'
    for j, v in enumerate(p): data_cell(ws, FIRST + i, 2 + j, v)
ws.freeze_panes = 'C8'
defname('SEC_PROFGROUPS', f"OFFSET('{SP}'!$B${FIRST},0,0,MAX(1,COUNTA('{SP}'!$B${FIRST}:$B$2000)),1)")

# =========================================================
# SecRule_選択肢
# =========================================================
ws = sheet(SC, 99)
header(ws, '⑧ 選択肢（③設定のダブルクリック選択の候補・入力チェックの基準）',
       'ボタン（または②の「まとめて取得」）で、基本設定の「オブジェクト取得folder」（通常 Shared）と○のスコープから、アドレス・アドレスグループ・EDL・サービス・サービスグループ・'
       'カスタムURLカテゴリ・タグ・HIPプロファイル・アプリグループ・アプリフィルタを取得し、定義されている folder を「スコープ」列に入れます。自作App は「オブジェクト取得folder」だけから取得します（基本設定「カスタムApp取得」=する のとき）。'
       'ゾーン・ユーザー／グループ・定義済みApp-ID は手入力の列です（青字・黄色）。ゾーンは「スコープ」列に使える folder を入れてください（複数はカンマ区切り、空欄=どこでも可）。②取得のルールで使われているゾーンは、取得時に使われている folder 付きで自動追加します。'
       '③の入力チェックでは、ルールのスコープ（folder）から見えないオブジェクト・ゾーン（自分のfolderか上位folderで定義されていないもの）をNGにします。', 30, 84)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
btn(ws, 'D3:E4', '▶ 選択肢を取得（API）')
S_ = ('スコープ', 13)
cols = [('No', 5, None), ('ゾーン\n(from / to)', 14, '手入力'), (S_[0], S_[1], '手入力'),
        ('アドレス', 22, 'API'), (S_[0], S_[1], 'API'), ('アドレスグループ', 22, 'API'), (S_[0], S_[1], 'API'),
        ('EDL', 22, 'API'), (S_[0], S_[1], 'API'), ('EDL種別', 8, 'API'), ('リージョン', 14, 'API'),
        ('サービス', 18, 'API'), (S_[0], S_[1], 'API'), ('サービスグループ', 18, 'API'), (S_[0], S_[1], 'API'),
        ('URLカテゴリ\n(カスタム)', 24, 'API'), (S_[0], S_[1], 'API'), ('URLカテゴリ\n(事前定義)', 30, 'API'),
        ('タグ', 14, 'API'), (S_[0], S_[1], 'API'), ('HIPプロファイル', 18, 'API'), (S_[0], S_[1], 'API'),
        ('ユーザー／グループ\n(source_user)', 30, '手入力'), ('定義済みApp-ID\n(application)', 22, '手入力'),
        ('自作App\n(application)', 20, 'API'), (S_[0], S_[1], 'API'), ('アプリグループ', 18, 'API'), (S_[0], S_[1], 'API'),
        ('アプリフィルタ', 18, 'API'), (S_[0], S_[1], 'API')]
heads(ws, 7, [c[0] for c in cols], [c[1] for c in cols])
for i, c in enumerate(cols, 1):
    if c[2]:
        put(ws, f'{CL(i)}6', c[2], Font(name=FN, size=8, bold=True, color='FFFFFF'), fill('548235') if c[2] == '手入力' else F_OPT, Alignment(horizontal='center'), False)
    if c[0] == S_[0]:
        ws.cell(7, i).fill = fill('8EA9DB')
PREDEF_APP = ['web-browsing', 'ssl', 'dns', 'ping', 'icmp', 'traceroute', 'ssh', 'telnet', 'ftp', 'smtp', 'pop3', 'imap', 'ntp',
          'snmp', 'syslog', 'ldap', 'kerberos', 'ms-rdp', 'msrpc', 'ms-ds-smb', 'netbios-ns', 'netbios-ss', 'dhcp',
          'http-proxy', 'ike', 'ipsec-esp', 'sip', 'rtsp', 'google-base', 'gmail-base', 'youtube-base', 'facebook-base',
          'linkedin-base', 'dropbox-base', 'box-base', 'salesforce-base', 'ms-teams', 'ms-office365-base',
          'sharepoint-online', 'outlook-web-online', 'zoom', 'webex-base', 'slack-base', 'github-base', 'windows-update',
          'apple-update', 'bittorrent', 'unknown-tcp', 'unknown-udp', 'unknown-p2p', 'incomplete', 'insufficient-data', 'not-applicable']
SH = 'Shared'
seed = {2: ['trust', 'untrust', 'RN-Trust', 'RN-Untrust', 'MU-Trust', 'MU-Untrust'],
        3: ['', '', 'Remote Networks', 'Remote Networks', 'Mobile Users Container', 'Mobile Users Container'],
        4: ['RN-Branch-Servers'], 5: ['Remote Networks'], 8: ['EDL-Block-IP', 'EDL-Block-IP-Partner'], 9: [SH, SH], 10: ['ip', 'ip'],
        12: ['tcp-8443'], 13: [SH], 16: ['URL-Custom-Allow'], 17: [SH], 18: PREDEF_URL, 19: ['prod'], 20: [SH],
        23: ['CN=all-staff,OU=Groups,DC=example,DC=co,DC=jp'], 24: PREDEF_APP,
        25: ['corp-portal'], 26: [SH], 27: ['grp-corp-apps'], 28: [SH], 29: ['filter-high-risk'], 30: [SH]}
MANUAL = (2, 3, 23, 24)
for col, vals in seed.items():
    for i, v in enumerate(vals):
        c = data_cell(ws, FIRST + i, col, v, F_IN if col in MANUAL else None)
        if col in MANUAL: c.font = Font(name=FN, size=10, color='0000FF')
for r in range(FIRST, FIRST + 300):
    for col in MANUAL:
        c = ws.cell(r, col)
        c.fill = F_IN; c.border = BOX; c.number_format = '@'
        if c.value is None: c.font = Font(name=FN, size=10, color='0000FF')
dv = DataValidation(type='list', formula1='"Shared,Remote Networks,Service Connections,Mobile Users Container,Mobile Users,Mobile Users Explicit Proxy"',
                    allow_blank=True, showErrorMessage=False)
dv.add(f'C{FIRST}:C{FIRST + 299}'); ws.add_data_validation(dv)
note(ws, 'AF7', '※ スコープ＝定義されている folder。ルールからは「自分の folder と上位の folder」（例：GlobalProtect なら Mobile Users・Mobile Users Container・Shared）の定義だけが見えます。')
note(ws, 'AF8', '※ ゾーンのスコープは手入力（空欄＝どのスコープでも使える。カンマ区切りで複数可）。プルダウン以外の値も入力できます。')
note(ws, 'AF9', '※ 事前定義URLカテゴリ・リージョン・定義済みApp-ID はどのスコープでも使えます。事前定義URLカテゴリは取得するとSCMの一覧に置き換わります（0件のときは残します）。')
note(ws, 'AF10', '※ EDL は種類が ip のものをアドレス列、url のものをURLカテゴリ列の候補にします（EDL種別列で判定）。')
note(ws, 'AF11', '※ 定義済みApp-ID（X列）は取得で上書きしません。使うApp-IDを追記してください。一覧に無い値もセルに直接入力できます（入力チェックでは警告）。')
ws.freeze_panes = 'B8'

# シート順
order = ['基本設定', SG, SS, SJ, SD, SL, SP, SC]
wb._sheets = [wb[n] for n in order]
# シート名に①②…を付ける（名前定義・数式の参照も置換）
NUM = '①②③④⑤⑥⑦⑧'
ren = {o: NUM[i] + o for i, o in enumerate(order)}
for o, n in ren.items():
    wb[o].title = n
def fix(t):
    for o, n in ren.items():
        t = t.replace(f"'{o}'!", f"'{n}'!").replace(f"{o}!", f"'{n}'!") if f"'{o}'!" in t or f"{o}!" in t else t
    return t
for nm in list(wb.defined_names.keys()):
    wb.defined_names[nm].attr_text = fix(wb.defined_names[nm].attr_text)
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                c.value = fix(c.value)
    for dv in ws.data_validations.dataValidation:
        if dv.formula1: dv.formula1 = fix(dv.formula1)
wb.active = 0
wb.save(OUT)
print('saved', OUT)

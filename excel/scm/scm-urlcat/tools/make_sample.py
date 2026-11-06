# -*- coding: utf-8 -*-
# SCM_URLCategory.xlsx にサンプルデータを入れた見本ファイルを作る（VBA実行後の見た目を再現）
#   使い方: python3 make_sample.py 入力.xlsx 出力.xlsx
import sys, datetime, json
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
NOW = datetime.datetime(2026, 10, 29, 10, 0, 0)
FN = 'Meiryo UI'
thin = Side(style='thin', color='000000')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
G, Y, GR, R, O, W, HD, SEC = 'C6EFCE', 'FFFF99', 'D9D9D9', 'FFC7CE', 'FFC000', 'FFEB9C', '2F5597', 'D9E1F2'
BASE = 'https://api.sase.paloaltonetworks.com/sse/config/v1/url-categories'

def cell(ws, r, c, v, fl=None, bold=False, color='000000', nf='@', wrap=True):
    x = ws.cell(r, c, v)
    x.font = Font(name=FN, size=10, bold=bold, color=color)
    if fl: x.fill = fill(fl)
    x.border = BOX
    x.alignment = Alignment(vertical='top', wrap_text=wrap)
    if nf and not isinstance(v, (int, float)): x.number_format = nf
    return x

# ---------------- SCM上の既存（取得結果のつもり） ----------------
WL, BL = 'WL-Allow', 'BL-Block'
existing = [
    ('a1f0c2d4-0001', 'Corp-SaaS', 'URL List', ['box.com', '*.box.com', 'slack.com', '*.slack.com', 'zoom.us', '*.zoom.us'], '社内利用SaaS'),
    ('a1f0c2d4-0002', 'Partner-Sites', 'URL List', ['partner-a.co.jp', 'partner-b.co.jp', 'portal.partner-c.com'], '取引先サイト'),
    ('a1f0c2d4-0003', 'Dev-Tools', 'URL List', ['github.com', '*.github.com', 'pypi.org', 'registry.npmjs.org'], '開発ツール'),
    ('a1f0c2d4-0004', 'Block-Social', 'Category Match', ['social-networking', 'streaming-media'], 'SNS・動画'),
    ('a1f0c2d4-0005', 'Old-Campaign', 'URL List', ['campaign2024.example.jp', 'lp.example.jp'], '2024年キャンペーン（終了）'),
    ('a1f0c2d4-0006', WL, 'URL List', ['update.microsoft.com', '*.windowsupdate.com', 'login.microsoftonline.com'], 'ホワイトリスト'),
    ('a1f0c2d4-0007', BL, 'URL List', ['malware-test.example.com', 'phish.example.net', 'bad.example.org'], 'ブラックリスト'),
    ('a1f0c2d4-0008', 'Cloud-Backup', 'URL List', ['dummy.fqdn/'], 'バックアップ（ダミーで作成済み）'),
]
DUMMY = 'dummy.fqdn/'
SLASH = lambda v: v if v.endswith('/') or '/' in v else v + '/'
NOSLASH = {'portal.partner-c.com'}                 # 既存に残っている「末尾 / なし」の例（警告になる）
existing = [(i, n, t, [x if (t != 'URL List' or x in NOSLASH) else SLASH(x) for x in l], d) for i, n, t, l, d in existing]
EX = {e[1]: e for e in existing}
other = [EX[WL], EX[BL]] + [e for e in existing if e[1] not in (WL, BL)]   # 横長ビューはWL→BL→残り

# ---------------- 基本設定 ----------------
cfg = wb['基本設定']
cfg['B16'] = 'Shared'
cfg['B26'] = WL
cfg['B28'] = BL

# ---------------- ② 既存一覧 ----------------
ws = wb['URLCat_既存一覧（リスト一覧）']
ws['B3'] = NOW; ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
for i, (id_, nm, typ, lst, desc) in enumerate(existing):
    r = 8 + i
    cell(ws, r, 1, f'=IF(C{r}="","",ROW()-7)', nf=None)
    for c, v in [(2, id_), (3, nm), (4, typ), (5, '\n'.join(lst[:5]) + (f'\n…ほか {len(lst) - 5} 件（全件はリスト内容シート）' if len(lst) > 5 else '')), (6, len(lst)), (7, desc), (8, 'folder=Shared')]:
        cell(ws, r, c, v)
    cell(ws, r, 9, f'=IF(C{r}="","",IF(AND(CFG_WL_NAME<>"",EXACT(C{r},CFG_WL_NAME)),"ホワイトリスト",'
                   f'IF(AND(CFG_BL_NAME<>"",EXACT(C{r},CFG_BL_NAME)),"ブラックリスト","一般")))', nf=None)

# ---------------- ⑤ 登録更新（リスト内容）（編集後の姿） ----------------
#   None = 消したセル（空きのまま）
wide = {
    WL: ['update.microsoft.com/', '*.windowsupdate.com/', 'login.microsoftonline.com/', '*.office.com/'],  # WLに1件追加
    BL: ['malware-test.example.com/', 'phish.example.net/', 'bad.example.org/'],
    'Corp-SaaS': ['box.com/', '*.box.com/', 'slack.com/', '*.slack.com/', None, None,      # zoom.us/ *.zoom.us/ を消した
                  'teams.microsoft.com/', 'slack.com/'],                                  # 追加・NG例（重複）
    'Partner-Sites': ['partner-a.co.jp/', 'partner-b.co.jp/', 'portal.partner-c.com',      # 警告例：末尾 / なし
                      'https://portal.partner-d.com/', 'part*.partner-e.com/', '*.*.partner-f.com/'],  # NG例
    'Dev-Tools': ['github.com/', '*.github.com/', 'pypi.org/', 'registry.npmjs.org/',
                  '*.npmjs.org/', 'gitlab.*.com/', '*.atlassian.net/jira'],               # 追加（2件は警告例）
    'Block-Social': ['social-networking', 'streaming-media'],
    'Old-Campaign': ['campaign2024.example.jp/', 'lp.example.jp/'],                        # 操作=削除
    'Cloud-Backup': [DUMMY, 'veeam.example.com/', '*.backup.example.net/'],               # ダミーの下に追記
    'Cloud-Storage': ['dropbox.com/', '*.dropbox.com/', 'drive.example.co.jp/'],          # 操作=登録（値ごと登録）
}
OPS = {'Old-Campaign': '削除', 'Cloud-Storage': '登録'}
EDIT = {'Dev-Tools': (None, '開発ツール（npm全体を許可）')}                         # description を書き換え
NEWCAT = (None, 'Cloud-Storage', 'URL List', [], 'クラウドストレージ（許可）')
def rule(e):
    p = e.find('/')
    host, path = (e[:p], e[p + 1:]) if p >= 0 else (e, '')
    if host == '': return 'NG:ホスト名がありません（先頭が / です）'
    lb = host.split('.')
    stars = [i for i, x in enumerate(lb) if '*' in x]
    if any(lb[i] != '*' for i in stars):
        return 'NG:ワイルドカード(*)はドットで区切った1区切り全体にしか使えません（aa*.aaa.com は登録不可。*.aaa.com/ の形にしてください）'
    if len(stars) > 1: return 'NG:ワイルドカード(*)は1か所だけです（*.*.com は登録不可）'
    w = []
    if '.' not in host: w.append('ドット(.)を含みません')
    if stars and stars[0] > 0: w.append('途中のワイルドカード（aaa.*.com/）は登録できますが推奨しません')
    if p < 0: w.append(f'末尾に / が無いため前方一致になり、{e}.net なども一致します（{e}/ を推奨）')
    elif path: w.append('パス指定（/ 以降）はSSL復号化が前提です')
    return ('警告:' + '／'.join(w)) if w else ''
have = {e[1]: e[3] for e in other}
LIST_ROWS = [(OPS.get(nm, ''), nm, typ, EDIT.get(nm, (None, None))[1] or desc, len(lst)) for _, nm, typ, lst, desc in other]
LIST_ROWS.append(('登録', 'Cloud-Storage', 'URL List', 'クラウドストレージ（許可）', None))
LIST_ROWS.append(('', 'Test-List', 'URL List', '', None))                      # NG例：登録の指定なし
IN, VB = 'FFF2CC', 'E2EFDA'

# ---------------- ③ 登録更新（リスト一覧） ----------------
ws = wb['URLCat_登録更新（リスト一覧）']
ws['B3'] = NOW
for i, (op, nm, typ, desc, cnt) in enumerate(LIST_ROWS):
    r = 8 + i
    for c, v in [(2, op or None), (3, nm), (4, typ), (5, desc or None), (6, cnt)]:
        x = ws.cell(r, c, v)
        x.font = Font(name=FN, size=10, strike=(op == '削除' and 3 <= c <= 5))
    if op == '登録': ws.cell(r, 2).fill = fill(G)
    if op == '削除': ws.cell(r, 2).fill = fill(GR)
    judge, msg = '変更なし', None
    if op: judge = op
    if nm in EDIT:
        judge = '変更'; ws.cell(r, 5).fill = fill(O); ws.cell(r, 5).font = Font(name=FN, size=10, bold=True)
    if nm == 'Test-List':
        judge, msg = 'NG', 'SCMに無いリストです。新しく作る場合は操作を「登録」にしてください'
    ws.cell(r, 7, judge); ws.cell(r, 8, msg)
    if judge == 'NG': ws.cell(r, 7).fill = fill(R)

# ---------------- ④ 既存一覧（リスト内容） ----------------
def wide_head(ws, c, nm, typ, L):
    cell(ws, 5, c, nm, fl=HD, bold=True, color='FFFFFF')
    cell(ws, 6, c, typ, fl=SEC)
    cell(ws, 7, c, f'=IF({L}5="","",COUNTA({L}8:{L}1048576))', fl=SEC, nf=None)
    ws.column_dimensions[L].width = 30
ws = wb['URLCat_既存一覧（リスト内容）']
ws['B3'] = NOW
for j, (_, nm, typ, lst, _d) in enumerate(other):
    c = 2 + j; L = get_column_letter(c)
    wide_head(ws, c, nm, typ, L)
    for k, e in enumerate(lst):
        ws.cell(8 + k, c, e).number_format = '@'

# ---------------- ⑤ 登録更新（リスト内容） ----------------
ws = wb['URLCat_登録更新（リスト内容）']
ws['B3'] = NOW
want = {}
for j, (id_, nm, typ, lst, desc) in enumerate(other + [NEWCAT]):
    c = 2 + j; L = get_column_letter(c)
    vals = wide[nm]
    op = OPS.get(nm, '')
    wide_head(ws, c, nm, typ, L)
    if op == '削除': ws.cell(5, c).fill = fill(GR); ws.cell(5, c).font = Font(name=FN, size=10, bold=True)
    seen = set()
    for k, e in enumerate(vals):
        if e is None: continue
        x = ws.cell(8 + k, c, e)
        x.number_format = '@'
        x.font = Font(name=FN, size=10)
        if op == '削除':
            x.fill = fill(GR); x.font = Font(name=FN, size=10, strike=True); continue
        if e == DUMMY:
            msg = '#DUMMY'
        elif e.lower().startswith(('http://', 'https://')):
            msg = 'NG:http:// / https:// は不要です'
        elif typ == 'URL List':
            msg = rule(e)
        else:
            msg = ''
        if not msg.startswith('NG') and e in seen:
            msg = 'NG:同じリストに同じ値があります（重複）'
        if not msg.startswith('NG'):
            seen.add(e); want.setdefault(nm, []).append(e)
            if msg == '' and e not in have.get(nm, []):
                x.fill = fill(G)
        if msg == '#DUMMY':
            msg = f'警告:ダミー（{DUMMY}）です。ほかの値を追記したので消してかまいません' if len([v for v in vals if v]) > 1 else ''
        if msg:
            x.fill = fill(R if msg.startswith('NG') else W)
            x.comment = Comment(msg[msg.index(':') + 1:], 'VBA')

# ---------------- ⑥ JSON_CURL ----------------
def body(nm, typ, lst, desc):
    d = {'name': nm, 'type': typ, 'list': lst}
    if desc: d['description'] = desc
    return json.dumps(d, ensure_ascii=False, separators=(',', ':'))
def curl(m, u, b):
    c = f"curl -s -X {m} '{u}' -H \"Authorization: Bearer $TOKEN\""
    if b: c += f" -H 'Content-Type: application/json' -d '{b}'"
    return c
jrows = [
    ('変更', 'Dev-Tools', 'PUT', BASE + '/a1f0c2d4-0003', body('Dev-Tools', 'URL List', want['Dev-Tools'], '開発ツール（npm全体を許可）')),
    ('登録', 'Cloud-Storage', 'POST', BASE + '?folder=Shared', body('Cloud-Storage', 'URL List', want['Cloud-Storage'], 'クラウドストレージ（許可）')),
    ('変更(URL)', 'Cloud-Backup', 'PUT', BASE + '/a1f0c2d4-0008', body('Cloud-Backup', 'URL List', want['Cloud-Backup'], 'バックアップ（ダミーで作成済み）')),
    ('変更(URL)', 'Corp-SaaS', 'PUT', BASE + '/a1f0c2d4-0001', body('Corp-SaaS', 'URL List', want['Corp-SaaS'], '社内利用SaaS')),
    ('変更(URL)', WL, 'PUT', BASE + '/a1f0c2d4-0006', body(WL, 'URL List', want[WL], 'ホワイトリスト')),
    ('削除', 'Old-Campaign', 'DELETE', BASE + '/a1f0c2d4-0005', ''),
]
ws = wb['JSON_CURL']
for i, (op, nm, m, u, b) in enumerate(jrows):
    r = 10 + i
    for c, v in enumerate([i + 1, op, nm, m, u, b, curl(m, u, b)], 1):
        cell(ws, r, c, v, fl='E2EFDA')

# ---------------- 差分（JSON発行後の姿） ----------------
summ = []   # name, kind, bc, ac, adds, dels, op, btype, ftype, bdesc, fdesc
for id_, nm, typ, lst, desc in other + [NEWCAT]:
    fin = want.get(nm, lst)
    adds = [e for e in fin if e not in lst]; dels = [e for e in lst if e not in fin]
    if nm == 'Old-Campaign':
        summ.append((nm, '削除', len(lst), 0, [], list(lst), '削除', typ, '', desc, ''))
    elif nm == 'Cloud-Storage':
        summ.append((nm, '新規', 0, len(fin), list(fin), [], '登録', '', 'URL List', '', desc))
    elif nm in EDIT:
        summ.append((nm, '変更', len(lst), len(fin), adds, dels, '変更', typ, typ, desc, EDIT[nm][1]))
    elif adds or dels:
        summ.append((nm, '変更', len(lst), len(fin), adds, dels, '変更(URL)', typ, typ, desc, desc))
    else:
        summ.append((nm, '変化なし', len(lst), len(fin), [], [], '', typ, typ, desc, desc))
kfill = {'新規': G, '変更': Y, '削除': GR}

# ⑦ 差分チェック（リスト一覧）
ws = wb['URLCat_差分チェック（リスト一覧）']
ws['B3'] = NOW; ws['B4'] = 'プレビュー（JSON作成済み・未実行）'
for i, (nm, kind, bc, ac, adds, dels, op, bt, ft, bd, fd) in enumerate(summ):
    r = 8 + i
    vals = [i + 1, nm, kind, op or None, bt or None, (ft or None) if kind != '削除' else None, bd or None,
            (fd or None) if kind != '削除' else None, bc, ac if kind != '削除' else None,
            f'+{len(adds)}' if adds else None, f'-{len(dels)}' if dels else None, None]
    for c, v in enumerate(vals, 1):
        x = cell(ws, r, c, v, fl=kfill.get(kind) if c <= 12 else None, nf=None if c in (1, 9, 10) else '@')
    if kind == '削除':
        ws.cell(r, 2).font = Font(name=FN, size=10, strike=True)
    if kind == '変更':
        hi = []
        if fd != bd: hi.append(8)
        if ac != bc: hi.append(10)
        if adds: hi.append(11)
        if dels: hi.append(12)
        for c in hi:
            ws.cell(r, c).fill = fill(O); ws.cell(r, c).font = Font(name=FN, size=10, bold=True)

# ⑧ 差分チェック（リスト内容）
ws = wb['URLCat_差分チェック（リスト内容）']
ws['B3'] = NOW; ws['B4'] = 'プレビュー（JSON作成済み・未実行）'
for k, (nm, kind, bc, ac, adds, dels, op, bt, ft, bd, fd) in enumerate(summ):
    c = 2 + k
    cnt = f'{bc} → {ac} 件' + (f'（+{len(adds)} / -{len(dels)}）' if adds or dels else '')
    cell(ws, 5, c, nm, fl=HD, bold=True, color='FFFFFF')
    cell(ws, 6, c, kind, fl=kfill.get(kind, SEC))
    cell(ws, 7, c, cnt, fl=SEC)
    if kind == '削除': ws.cell(5, c).font = Font(name=FN, size=10, bold=True, color='FFFFFF', strike=True)
    ws.column_dimensions[get_column_letter(c)].width = 30
    vals = [(e, e in dels) for e in have.get(nm, [])] + [(e, None) for e in adds]
    for i, (e, isdel) in enumerate(vals):
        x = ws.cell(8 + i, c, e); x.number_format = '@'
        x.font = Font(name=FN, size=10, strike=bool(isdel))
        if isdel: x.fill = fill(GR)
        elif isdel is None: x.fill = fill(G)

# ---------------- 実行ログ ----------------
ws = wb['実行ログ']
for i, (op, nm, m, u, st, res, msg) in enumerate([
        ('トークン取得', '', 'POST', 'https://auth.apps.paloaltonetworks.com/oauth2/access_token', 200, '成功', 'expires_in=899'),
        ('既存取得', '', 'GET', BASE + '?folder=Shared', 200, '成功', 'URLリスト 8 件、リスト内容 24 件 / 既存一覧（リスト一覧・リスト内容）と登録更新（リスト一覧・リスト内容）を作り直しました。')]):
    r = 6 + i
    for c, v in enumerate([NOW, op, nm, m, u, st, res, msg], 1):
        x = cell(ws, r, c, v, nf=None)
    ws.cell(r, 1).number_format = 'yyyy/mm/dd hh:mm:ss'

for s in wb: s.sheet_view.tabSelected = False
wb.active = wb.sheetnames.index('URLCat_差分チェック（リスト一覧）')
wb.active.sheet_view.tabSelected = True
wb.save(OUT)
print('saved', OUT)

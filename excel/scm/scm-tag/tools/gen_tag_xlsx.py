# -*- coding: utf-8 -*-
# SCM Tag Excel 生成（ServiceObject ツールのブックから 基本設定 / JSON_CURL / 実行ログ を流用）
import sys, datetime, json as _j
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

SRC = '/mnt/project-files/scm-service/SCM_ServiceObject.xlsx'
OUT = sys.argv[1]

TE, TI, SJ, TD, SL, TC = 'Tag_既存一覧', 'Tag_登録変更削除', 'JSON_CURL', 'Tag_差分確認', '実行ログ', 'Tag_色一覧'
FIRST, LAST, EXL = 8, 507, 5000

FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
F_TITLE, F_HEAD, F_IN, F_FX, F_VBA, F_BTN = fill('1F3864'), fill('2F5597'), fill('FFF2CC'), fill('F2F2F2'), fill('E2EFDA'), fill('C55A11')
F_REQ, F_OPT = fill('C00000'), fill('7F7F7F')

# API仕様（Tags.yaml）の color 一覧と、色見本の目安
COLORS = [
    ('Red', 'FF0000'), ('Green', '00A651'), ('Blue', '0070C0'), ('Yellow', 'FFFF00'), ('Copper', 'B87333'),
    ('Orange', 'FFA500'), ('Purple', '800080'), ('Gray', '808080'), ('Light Green', '90EE90'), ('Cyan', '00FFFF'),
    ('Light Gray', 'D3D3D3'), ('Blue Gray', '6699CC'), ('Lime', '00FF00'), ('Black', '000000'), ('Gold', 'FFD700'),
    ('Brown', '8B4513'), ('Olive', '808000'), ('Maroon', '800000'), ('Red-Orange', 'FF4500'), ('Yellow-Orange', 'FFAE42'),
    ('Forest Green', '228B22'), ('Turquoise Blue', '00C5CD'), ('Azure Blue', '007FFF'), ('Cerulean Blue', '2A52BE'),
    ('Midnight Blue', '191970'), ('Medium Blue', '0000CD'), ('Cobalt Blue', '0047AB'), ('Violet Blue', '324AB2'),
    ('Blue Violet', '8A2BE2'), ('Medium Violet', '9370DB'), ('Medium Rose', 'D98695'), ('Lavender', 'E6E6FA'),
    ('Orchid', 'DA70D6'), ('Thistle', 'D8BFD8'), ('Peach', 'FFDAB9'), ('Salmon', 'FA8072'), ('Magenta', 'FF00FF'),
    ('Red Violet', 'C71585'), ('Mahogany', 'C04000'), ('Burnt Sienna', 'E97451'), ('Chestnut', '954535'),
]
CMAP = dict(COLORS)

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

def heads(ws, row, names, widths):
    for i, (n, w) in enumerate(zip(names, widths), 1):
        put(ws, f'{CL(i)}{row}', n, Font(name=FN, size=10, bold=True, color='FFFFFF'), F_HEAD,
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

def nest(conds, last):
    f = last
    for cond, val in reversed(conds):
        f = f'IF({cond},{val},{f})'
    return f

# =========================================================
# 基本設定 / JSON_CURL / 実行ログ
# =========================================================
cfg = wb['基本設定']
cfg['B7'] = '/sse/config/v1/tags'
cfg['D7'] = 'Tag（タグ）'
cfg['B19'] = None; cfg['B20'] = None; cfg['B21'] = None

js = wb[SJ]
js['A2'] = ('③のJSON発行ボタンでVBAが10行目以降を生成します（緑=VBA書込み）。curlはbash形式。先に5行目でトークン取得→6行目でTOKEN変数を設定してから各行を実行。'
            '「API実行」ボタンでExcelから直接実行も可能（結果は⑤実行ログへ）。登録→変更→削除の順に並べます。'
            '実行後は⑥Tag_差分確認に結果の一覧が出ます。')
for row in js.iter_rows(min_row=10, max_row=js.max_row):
    for c in row: c.value = None

lg = wb[SL]
for row in lg.iter_rows(min_row=6, max_row=lg.max_row):
    for c in row:
        c.value = None; c.border = Border(); c.number_format = 'General'

# =========================================================
# ⑦ Tag_色一覧（color のプルダウン元・色見本）
# =========================================================
ws = sheet(TC, 99)
header(ws, '⑦ タグの色一覧（color のプルダウン元）',
       'API仕様（Tags.yaml）で指定できる color の一覧です。③のcolor列のプルダウンはこの表を参照します。色見本はSCM画面の色の目安です。'
       '名前は大文字・小文字も含めてこのとおりに指定します。この表は編集しないでください。', 4, 36)
heads(ws, 7, ['No', 'color', '色見本（目安）', '備考'], [6, 20, 16, 30])
for i, (n, h) in enumerate(COLORS):
    r = FIRST + i
    put(ws, f'A{r}', i + 1)
    put(ws, f'B{r}', n, nf='@')
    put(ws, f'C{r}', None, fl=fill(h))
    put(ws, f'D{r}', None)
put(ws, f'A{FIRST + len(COLORS)}', '※ colorを空欄にすると色なし（No Color）になります。', Font(name=FN, size=9, color='595959'), border=False)
ws.freeze_panes = 'A8'

# =========================================================
# ② Tag_既存一覧
# =========================================================
ex_tags = [
    ('prod', 'Red', '本番環境'),
    ('mgmt', 'Blue', '管理系'),
    ('dev', 'Green', '開発環境'),
    ('test', 'Yellow', ''),
    ('old-project', 'Gray', '旧プロジェクト（廃止予定）'),
    ('DMZ', 'Orange', 'DMZセグメント'),
]
ws = sheet(TE, 1)
header(ws, '② 既存情報（SCMから取得した Tag 一覧）',
       'ボタンを押すとAPI（GET）でタグを取得し、8行目以降に表を再生成します（B〜G列=VBA書込み、A列=関数）。直接編集はしないでください。'
       'E列の色見本はcolorに合わせてVBAが塗ります（⑦色一覧の目安）。', 7)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{EXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
heads(ws, 7, ['No', 'id', 'name', 'color', '色見本', 'comments', 'コンテナ'], [6, 38, 24, 16, 8, 40, 16])
for i, (n, c, cm) in enumerate(ex_tags):
    r = FIRST + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)')
    put(ws, f'B{r}', f'7a3c4d50-000{i+1}-4d2e-8a11-00000000002{i+1}', nf='@')
    put(ws, f'C{r}', n, nf='@')
    put(ws, f'D{r}', c, nf='@')
    put(ws, f'E{r}', None, fl=fill(CMAP[c]))
    put(ws, f'F{r}', cm or None, nf='@')
    put(ws, f'G{r}', 'folder=Shared', nf='@')
ws.freeze_panes = 'D8'

# =========================================================
# ③ Tag_登録変更削除
#   A No B 操作 C name D color E 色見本 F comments
#   G 既存有無 H 既存id I 既存color J 既存comments K 既存コンテナ L 一覧内重複 M 判定 N 理由 O 作業:既存行
# =========================================================
ws = sheet(TI, 2)
header(ws, '③ 今回の登録・変更・削除一覧（Tag）',
       '黄色(B〜D・F列)に入力。操作・name（変更・削除時）・colorはプルダウンから選択できます。E列はcolorの色見本です。G〜O列は関数で②既存一覧と自動照合し、判定がNGの行はJSON発行対象外。'
       'nameは大文字・小文字も区別して照合します。変更時に既存と異なる項目は橙色で表示。変更(PUT)は全項目置換のため、残したい値（color・comments）もすべて入力してください。'
       'タグ名の変更（リネーム）はできません。オブジェクトやルールで使用中のタグを削除すると、SCMがエラー（E009 Reference Not Zero）を返し、実行ログに記録されます。', 15, 60)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(M{FIRST}:M{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(M{FIRST}:M{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(M{FIRST}:M{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行 → ④へ')
btn(ws, 'I3:J4', '▶ 入力クリア')
labels(ws, 6, {'B': '必須', 'C': '必須', 'D': '任意', 'F': '任意'})
put(ws, 'G6', '※ color・commentsは空欄可（変更時の空欄は既存値が消えます。colorの空欄は色なし）', Font(name=FN, size=9, color='595959'), border=False)
heads(ws, 7, ['No', '操作', 'name\n(127文字以内)', 'color', '色見本', 'comments\n(1023文字以内)',
              '既存有無', '既存id', '既存color', '既存comments', '既存コンテナ', '一覧内重複', '判定', '理由', '作業:既存行'],
      [5, 8, 24, 16, 8, 36, 9, 22, 14, 24, 14, 10, 8, 44, 9])
ws.row_dimensions[7].height = 36
TEC = ex(TE, 'C')
for r in range(FIRST, LAST + 1):
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    for col in 'BCDF':
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
    put(ws, f'E{r}', None)
    put(ws, f'O{r}', f'=IF($C{r}="","",IFERROR(MATCH(TRUE,INDEX(EXACT({TEC},$C{r}),0),0),""))', fl=F_FX)
    put(ws, f'G{r}', f'=IF($C{r}="","",IF($O{r}<>"","あり","なし"))', fl=F_FX)
    for col, src in {'H': 'B', 'I': 'D', 'J': 'F', 'K': 'G'}.items():
        put(ws, f'{col}{r}', f'=IF($O{r}="","",INDEX({ex(TE, src)},$O{r})&"")', fl=F_FX)
    put(ws, f'L{r}', f'=IF($C{r}="","",IF(SUMPRODUCT(--EXACT($C${FIRST}:$C${LAST},$C{r}))>1,"重複",""))', fl=F_FX)
    put(ws, f'M{r}', f'=IF($N{r}="","",IF($N{r}="OK","OK",LEFT($N{r},FIND(":",$N{r})-1)))', fl=F_FX)
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$L{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>127', '"NG:nameが127文字超"'),
        (f'TRIM($C{r})<>$C{r}', '"NG:nameの前後や連続した空白を削除してください"'),
        (f'AND($B{r}="登録",$G{r}="あり")', '"NG:既存に同名あり(登録不可)"'),
        (f'AND($B{r}<>"登録",$G{r}="なし")', '"NG:既存エントリ無(大文字・小文字も区別)"'),
        (f'AND($B{r}<>"登録",ISNUMBER(SEARCH("predefined",$K{r})))', '"NG:定義済みタグは変更・削除できません"'),
        (f'AND($B{r}<>"削除",$D{r}<>"",SUMPRODUCT(--EXACT(TAG_COLORS,$D{r}))=0)', '"NG:colorはプルダウン(⑦色一覧)から選択してください"'),
        (f'AND($B{r}<>"削除",LEN($F{r})>1023)', '"NG:commentsが1023文字超"'),
        (f'AND($B{r}="登録",COUNTIF({TEC},$C{r})>0)', '"警告:大文字・小文字だけが違う既存タグあり"'),
        (f'AND($B{r}<>"登録",$K{r}<>"",$K{r}<>CFG_CONT_TYPE&"="&CFG_CONT_NAME)', f'"警告:基本設定と別のコンテナ("&$K{r}&")のタグ"'),
        (f'AND($B{r}="変更",EXACT($D{r}&"",$I{r}),EXACT($F{r}&"",$J{r}))', '"警告:既存と差分なし"'),
        (f'LENB($C{r})<>LEN($C{r})', '"警告:nameに全角文字を含む(他のオブジェクトから参照するときは同じ表記で入力)"'),
    ]
    put(ws, f'N{r}', f'=IF($C{r}="","",{nest(conds, chr(34) + "OK" + chr(34))})', fl=F_FX)
    ws[f'F{r}'].alignment = Alignment(wrap_text=True, vertical='top')
ws.freeze_panes = 'D8'
dv = DataValidation(type='list', formula1='"登録,変更,削除"', allow_blank=True); dv.add(f'B{FIRST}:B{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='TAG_NAMES', allow_blank=True, showErrorMessage=False); dv.add(f'C{FIRST}:C{LAST}'); ws.add_data_validation(dv)
dv = DataValidation(type='list', formula1='TAG_COLORS', allow_blank=True, showErrorMessage=True,
                    errorTitle='color', error='colorはプルダウン（⑦色一覧）から選択してください。'); dv.add(f'D{FIRST}:D{LAST}'); ws.add_data_validation(dv)
cf(ws, f'M{FIRST}:N{LAST}', f'$M{FIRST}="NG"', 'FFC7CE')
cf(ws, f'M{FIRST}:N{LAST}', f'$M{FIRST}="警告"', 'FFEB9C')
cf(ws, f'M{FIRST}:N{LAST}', f'$M{FIRST}="OK"', 'C6EFCE')
cf(ws, f'D{FIRST}:D{LAST}', f'AND($B{FIRST}="変更",$G{FIRST}="あり",NOT(EXACT($D{FIRST}&"",$I{FIRST})))', 'F8CBAD')
cf(ws, f'F{FIRST}:F{LAST}', f'AND($B{FIRST}="変更",$G{FIRST}="あり",NOT(EXACT($F{FIRST}&"",$J{FIRST})))', 'F8CBAD')
for n, h in COLORS:   # 色見本（E列）
    cf(ws, f'E{FIRST}:E{LAST}', f'EXACT($D{FIRST},"{n}")', h)
tag_in = [
    ('登録', 'pci', 'Purple', 'PCI DSS対象'),
    ('登録', 'vendor-access', 'Cerulean Blue', '保守ベンダー接続'),
    ('変更', 'mgmt', 'Cobalt Blue', '管理系（色を変更）'),
    ('変更', 'test', 'Yellow', 'テスト環境'),
    ('削除', 'old-project', '', ''),
    ('登録', 'Prod', 'Red', '大文字違い'),
    ('登録', 'dev', 'Green', '既存と同名'),
    ('変更', 'staging', 'Lime', '既存に無い'),
]
for i, v in enumerate(tag_in):
    for j, x in enumerate(v):
        ws[f'{"BCDF"[j]}{FIRST+i}'] = x if x != '' else None

# =========================================================
# ⑥ Tag_差分確認（行の色・変更セルの色は VBA が直接塗る。条件付き書式は使わない）
# =========================================================
ws = sheet(TD, 4)
fields = ['color', 'comments']
ncol = 3 + len(fields) + 5
header(ws, '⑥ 差分確認（Tag）：現在の設定一覧と既存一覧の比較',
       '④のAPI実行が終わると、VBAがSCMに今設定されている全件を取得して一覧表示し、②既存一覧（実行前）と比べた変化をC列に付けます。'
       '行の背景色で変化を表します（追加=緑、変更=黄、変化なし=白、削除=グレー）。変更された項目のセルは橙色・太字です。削除されたものは一覧の後ろに、既存一覧（実行前）の内容で表示します。'
       '判定が要確認のもの（判定列が赤）はPush前に理由を確認してください。ボタンでいつでも作り直せます。', ncol, 56)
put(ws, 'A3', '作成日時', Font(name=FN, size=10, bold=True))
put(ws, 'B3', None, fl=F_VBA, nf='yyyy/mm/dd hh:mm:ss')
for lab, val in [(('C3', 'D3'), '全件'), (('E3', 'F3'), '変化なし'), (('G3', 'H3'), '追加'),
                 (('C4', 'D4'), '変更'), (('E4', 'F4'), '削除'), (('G4', 'H4'), '要確認')]:
    stat(ws, lab, val, None, F_VBA)
btn(ws, 'J3:K4', '▶ 差分チェック（再作成）')
put(ws, 'C5', '凡例：', Font(name=FN, size=10, bold=True), border=False)
for a, t, c in [('D5', '追加', 'C6EFCE'), ('E5', '変更', 'FFFF99'), ('F5', '変化なし', 'FFFFFF'), ('G5', '削除', 'D9D9D9'),
                ('H5', '要確認(判定列)', 'FFC7CE'), ('I5', '変更セル', 'FFC000')]:
    put(ws, a, t, Font(name=FN, size=10, bold=(t == '変更セル')), fill(c), Alignment(horizontal='center'))
heads(ws, 7, ['No', 'name', '変化', 'color（現在）', 'comments（現在）', '変更項目', '変更前（既存一覧）', '④の操作', 'HTTP', '判定'],
      [5, 24, 9, 16, 34, 16, 34, 9, 7, 26])
ws.column_dimensions['K'].width = 12
diff_rows = [
    ('prod', '変化なし', 'Red', '本番環境', '', '', '', '', ''),
    ('mgmt', '変更', 'Cobalt Blue', '管理系（色を変更）', 'color comments', 'color: Blue\ncomments: 管理系', '変更', '200', '想定どおり'),
    ('dev', '変化なし', 'Green', '開発環境', '', '', '', '', ''),
    ('test', '変更', 'Yellow', 'テスト環境', 'comments', 'color: Yellow', '変更', '200', '想定どおり'),
    ('DMZ', '変化なし', 'Orange', 'DMZセグメント', '', '', '', '', ''),
    ('pci', '追加', 'Purple', 'PCI DSS対象', '', '', '登録', '201', '想定どおり'),
    ('vendor-access', '追加', 'Cerulean Blue', '保守ベンダー接続', '', '', '登録', '201', '想定どおり'),
    ('Prod', '追加', 'Red', '大文字違い', '', '', '登録', '201', '想定どおり'),
    ('old-project', '削除', 'Gray', '旧プロジェクト（廃止予定）', '', '', '削除', '200', '想定どおり'),
]
ROWFILL = {'追加': 'C6EFCE', '変更': 'FFFF99', '削除': 'D9D9D9'}
cnt = {'変化なし': 0, '追加': 0, '変更': 0, '削除': 0}
for i, v in enumerate(diff_rows):
    r = FIRST + i
    vals = [i + 1] + list(v)
    for j, x in enumerate(vals):
        c = put(ws, f'{CL(1+j)}{r}', x if x != '' else None, al=Alignment(wrap_text=True, vertical='top'))
        if v[1] in ROWFILL: c.fill = fill(ROWFILL[v[1]])
    if v[1] == '変更':
        for k, f in enumerate(fields):
            if f in v[4].split():
                c = ws[f'{CL(4+k)}{r}']; c.fill = fill('FFC000'); c.font = Font(name=FN, size=10, bold=True)
    cnt[v[1]] += 1
ws['D3'] = len(diff_rows) - cnt['削除']; ws['F3'] = cnt['変化なし']; ws['H3'] = cnt['追加']
ws['D4'] = cnt['変更']; ws['F4'] = cnt['削除']; ws['H4'] = 0
ws['B3'] = datetime.datetime(2026, 10, 28, 15, 40)
ws.freeze_panes = 'D8'

# ④ JSON_CURL サンプル（VBAのJSON発行と同じ形式）
BASE = 'https://api.sase.paloaltonetworks.com'
T = BASE + '/sse/config/v1/tags'
def jstr(v): return _j.dumps(v, ensure_ascii=False)
def body(nm, color, cm):
    b = '{"name":' + jstr(nm)
    if color: b += ',"color":' + jstr(color)
    if cm: b += ',"comments":' + jstr(cm)
    return b + '}'
def curl(m, u, b):
    c = f"curl -s -X {m} '{u}' -H \"Authorization: Bearer $TOKEN\""
    if b: c += " -H 'Content-Type: application/json' -d '" + b.replace("'", "'\\''") + "'"
    return c
rows = [
    ('登録', 'pci', 'POST', T + '?folder=Shared', body('pci', 'Purple', 'PCI DSS対象'), 201),
    ('登録', 'vendor-access', 'POST', T + '?folder=Shared', body('vendor-access', 'Cerulean Blue', '保守ベンダー接続'), 201),
    ('登録', 'Prod', 'POST', T + '?folder=Shared', body('Prod', 'Red', '大文字違い'), 201),
    ('変更', 'mgmt', 'PUT', T + '/7a3c4d50-0002-4d2e-8a11-000000000022', body('mgmt', 'Cobalt Blue', '管理系（色を変更）'), 200),
    ('変更', 'test', 'PUT', T + '/7a3c4d50-0004-4d2e-8a11-000000000024', body('test', 'Yellow', 'テスト環境'), 200),
    ('削除', 'old-project', 'DELETE', T + '/7a3c4d50-0005-4d2e-8a11-000000000025', '', 200),
]
js = wb[SJ]
for i, (op, nm, m, u, b, st) in enumerate(rows):
    r = 10 + i
    for j, v in enumerate([i + 1, op, nm, m, u, b or None, curl(m, u, b), st]):
        put(js, f'{CL(1+j)}{r}', v, al=Alignment(wrap_text=True, vertical='top'))

# 名前定義
wb.defined_names['EX_FETCHED_AT'] = DefinedName('EX_FETCHED_AT', attr_text=f"{Q(TE)}!$B$3")
wb.defined_names['TAG_NAMES'] = DefinedName('TAG_NAMES', attr_text=f"OFFSET({Q(TE)}!$C$8,0,0,MAX(1,COUNTA({Q(TE)}!$C$8:$C$5000)),1)")
wb.defined_names['TAG_COLORS'] = DefinedName('TAG_COLORS', attr_text=f"{Q(TC)}!$B${FIRST}:$B${FIRST + len(COLORS) - 1}")
wb[TE]['B3'] = datetime.datetime(2026, 10, 28, 15, 30)

order = ['基本設定', TE, TI, SJ, TD, SL, TC]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for w in wb: w.sheet_view.tabSelected = (w.title == '基本設定')
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('saved', OUT)

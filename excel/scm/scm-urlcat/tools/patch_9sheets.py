# -*- coding: utf-8 -*-
# 9シート構成に作り替える（2026-10-30 ユーザー指示）
#   基本設定 / 既存一覧（リスト一覧） / 登録更新（リスト一覧） / 既存一覧（リスト内容） / 登録更新（リスト内容） /
#   JSON_CURL / 差分チェック（リスト一覧） / 差分チェック（リスト内容） / 実行ログ
#   順番：… → patch_dchk.py → patch_nolist.py → このスクリプト
#   使い方: python3 patch_9sheets.py 入力.xlsx 出力.xlsx
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
HD, SEC, IN, VBA = '2F5597', 'D9E1F2', 'FFF2CC', 'E2EFDA'

EXL, LREG, CEX, CREG = 'URLCat_既存一覧（リスト一覧）', 'URLCat_登録更新（リスト一覧）', 'URLCat_既存一覧（リスト内容）', 'URLCat_登録更新（リスト内容）'
DLIST, DCONT = 'URLCat_差分チェック（リスト一覧）', 'URLCat_差分チェック（リスト内容）'

def btn(ws, rng, t):
    ws.merge_cells(rng)
    c = ws[rng.split(':')[0]]
    c.value = t
    c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.fill = fill('C55A11')
    c.alignment = Alignment(horizontal='center', vertical='center')

def base(name, idx, title, desc, buttons, a3='作成日時', a4=None, h2=60, last='L'):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.merge_cells(f'A1:{last}1'); ws.merge_cells(f'A2:{last}2')
    ws['A1'] = title
    ws['A1'].font = Font(name=FN, size=14, bold=True, color='FFFFFF'); ws['A1'].fill = fill('1F3864')
    ws.row_dimensions[1].height = 26
    ws['A2'] = desc
    ws['A2'].font = Font(name=FN, size=9, color='595959'); ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[2].height = h2
    for a, t in [('A3', a3), ('A4', a4)]:
        if t:
            ws[a] = t; ws[a].font = Font(name=FN, size=10, bold=True)
            b = ws['B' + a[1]]; b.font = Font(name=FN, size=10, bold=True); b.fill = fill(VBA)
    ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
    for rng, t in buttons:
        btn(ws, rng, t)
    return ws

def label(ws, r, t, c=1):
    x = ws.cell(r, c, t)
    x.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); x.fill = fill(HD); x.border = BOX
    x.alignment = Alignment(vertical='top', wrap_text=True)

def legend(ws, col, items):
    ws.cell(3, col - 1, '凡例：').font = Font(name=FN, size=10, bold=True)
    for k, (t, c) in enumerate(items):
        x = ws.cell(3, col + k, t)
        x.font = Font(name=FN, size=10, bold=True); x.fill = fill(c); x.border = BOX
        x.alignment = Alignment(horizontal='center')

def wide_frame(ws, labels, ncol=30, entry_row=8, editable=False):
    for r, t in labels:
        label(ws, r, t)
    ws.column_dimensions['A'].width = 14
    for col in range(2, 2 + ncol):
        L = get_column_letter(col)
        ws.column_dimensions[L].width = 30
        for r, _ in labels[:-1]:
            x = ws.cell(r, col)
            x.border = BOX; x.alignment = Alignment(wrap_text=True, vertical='top'); x.number_format = '@'
            if r == labels[0][0]:
                x.fill = fill(HD); x.font = Font(name=FN, size=10, bold=True, color='FFFFFF')
            else:
                x.fill = fill(SEC); x.font = Font(name=FN, size=10)
    ws.freeze_panes = f'B{entry_row}'

# ---- 既存一覧（リスト一覧）：名前を変えて流用 ----
ex = wb['URLCat_既存一覧']
ex.title = EXL
for n in ('EX_FETCHED_AT', 'EX_NAMES'):
    if n in wb.defined_names:
        wb.defined_names[n].attr_text = wb.defined_names[n].attr_text.replace("'URLCat_既存一覧'", f"'{EXL}'")
ex['A1'] = '② 既存一覧（リスト一覧）：SCMから取得したURLリスト（カスタムURLカテゴリ）の一覧'
ex['A2'] = ('「既存情報取得」でAPI（GET）を実行し、8行目以降を作り直します（B〜H列=VBA書込み、A・I列=関数）。リスト内容は先頭5件だけ表示します（全件は既存一覧（リスト内容））。'
            '同時に 既存一覧（リスト内容）・登録更新（リスト一覧・リスト内容）も作り直します（登録更新に未反映の編集があるときは確認します）。差分チェックの基準になるので直接編集しないでください。')
ex['E7'] = 'リスト内容（先頭5件）'

# ---- 登録更新（リスト一覧） ----
ws = base(LREG, 3, '③ 登録更新（リスト一覧）：リストの作成・削除、type・descriptionの変更（1行1リスト）',
          '「既存情報取得」でSCMの現在の一覧が入ります。新しいリストは空いている行に name・type・description を書いて操作を「登録」に、削除は操作を「削除」にします。'
          'type・descriptionを書き換えると変更になります（橙色）。行を消したリストは変更しません。リストの中身は登録更新（リスト内容）で編集します（登録時に中身が無ければダミー1件で作成）。'
          '「差分チェック」でG列に判定、H列に理由を書きます（NG=赤、警告=黄）。結果は差分チェック（リスト一覧・リスト内容）シートに出ます。',
          [('D3:E4', '▶ 差分チェック'), ('G3:H4', '▶ JSON発行 → JSON_CURLへ'), ('J3:K4', '▶ 既存一覧から作り直す')], a3='取得日時', h2=75)
for c, (t, w) in enumerate([('No', 6), ('操作', 10), ('name', 28), ('type', 16), ('description', 40), ('件数(既存)', 10), ('判定', 14), ('理由', 50)], 1):
    label(ws, 7, t, c); ws.column_dimensions[get_column_letter(c)].width = w
for r in range(8, 208):
    ws.cell(r, 1, f'=IF(C{r}="","",ROW()-7)')
    for c in range(1, 9):
        x = ws.cell(r, c); x.border = BOX; x.font = Font(name=FN, size=10); x.alignment = Alignment(vertical='top', wrap_text=c in (5, 8))
        if 2 <= c <= 5: x.fill = fill(IN); x.number_format = '@'
        elif c >= 6: x.fill = fill(VBA)
dv1 = DataValidation(type='list', formula1='"登録,削除"', allow_blank=True); dv1.add('B8:B207')
dv2 = DataValidation(type='list', formula1='"URL List,Category Match"', allow_blank=True); dv2.add('D8:D207')
ws.add_data_validation(dv1); ws.add_data_validation(dv2)
ws.freeze_panes = 'D8'

# ---- 既存一覧（リスト内容） ----
ws = base(CEX, 4, '④ 既存一覧（リスト内容）：SCMから取得したリストの中身（横長：1列1リスト、1セル1件）',
          '「既存情報取得」でVBAが作り直します。差分チェックの基準になるので直接編集しないでください。基本設定のホワイトリスト／ブラックリストを先頭に並べます。',
          [('D3:E4', '▶ 既存情報取得（API）')], a3='取得日時', h2=40)
wide_frame(ws, [(5, 'name'), (6, 'type'), (7, '件数'), (8, '値 →')], ncol=10)
for col in range(2, 12):
    L = get_column_letter(col); ws.cell(7, col).value = f'=IF({L}5="","",COUNTA({L}8:{L}1048576))'

# ---- 登録更新（リスト内容） ----
ws = base(CREG, 5, '⑤ 登録更新（リスト内容）：リストの中身の更新（横長：1列1リスト、1セル1件）',
          '「既存情報取得」で既存一覧（リスト内容）と同じ内容が入ります。8行目以降を編集します（追加は列の下に追記、削除はセルを消す。空きセルは詰めなくてかまいません）。'
          '登録更新（リスト一覧）で「登録」したリストは、空いている列の5行目に同じ name を書くとその中身で作成します。値が1件も無い列のリストは中身を変更しません。'
          '「差分チェック」でNG=赤、警告=黄、追加する値=緑になり、理由はセルのコメントに出ます。'
          '書式：*.aaa.com/ を推奨。aa*.aaa.com・*.*.com はNG。末尾 / なし（前方一致）・途中の*・パス指定（/以降）は警告',
          [('D3:E4', '▶ 差分チェック'), ('G3:H4', '▶ JSON発行 → JSON_CURLへ'), ('J3:K4', '▶ 既存一覧から作り直す')], a3='取得日時', h2=75)
wide_frame(ws, [(5, 'name'), (6, 'type'), (7, '件数'), (8, '値 →')], ncol=30)
for col in range(2, 32):
    L = get_column_letter(col); ws.cell(7, col).value = f'=IF({L}5="","",COUNTA({L}8:{L}1048576))'

# ---- 差分チェック（リスト一覧） ----
ws = base(DLIST, 7, '⑦ 差分チェック（リスト一覧）：既存一覧と登録更新の比較（1行1リスト）',
          '「差分チェック」「JSON発行」でVBAが作り直します（直接編集しないでください）。行の色：新規=緑、変更=黄、削除=グレー（name に取り消し線）、変化なし=色なし。'
          '変わる項目は橙色。「API実行」のあとはSCMから取得し直し、M列「実行結果」に OK=緑、要確認=赤 を書きます。「変化なしを隠す」で変化の無い行を非表示にできます。',
          [('D3:E4', '▶ 変化なしを隠す'), ('G3:H4', '▶ 全て表示'), ('J3:K4', '▶ 実行結果の再確認')], a4='状態', h2=50, last='M')
legend(ws, 15, [('新規', 'C6EFCE'), ('変更', 'FFFF99'), ('削除', 'D9D9D9'), ('変わる項目', 'FFC000'), ('要確認', 'FFC7CE')])
for c, (t, w) in enumerate([('No', 6), ('name', 26), ('判定', 10), ('操作(JSON)', 12), ('type(現在)', 14), ('type(反映後)', 14),
                            ('description(現在)', 28), ('description(反映後)', 28), ('件数(現在)', 10), ('件数(反映後)', 10),
                            ('追加', 8), ('削除', 8), ('実行結果', 36)], 1):
    label(ws, 7, t, c); ws.column_dimensions[get_column_letter(c)].width = w
for c in range(15, 20):
    ws.column_dimensions[get_column_letter(c)].width = 12
ws.freeze_panes = 'C8'

# ---- 差分チェック（リスト内容） ----
ws = base(DCONT, 8, '⑧ 差分チェック（リスト内容）：既存一覧と登録更新の比較（横長：1列1リスト）',
          '「差分チェック」「JSON発行」でVBAが作り直します（直接編集しないでください）。SCMにある値をSCMの順に並べ、削除される値はグレーと取り消し線、その下に追加される値を緑で表示します。'
          '6行目の判定は 新規=緑、変更=黄、削除=グレー、変化なし=色なし。「変化なしを隠す」で変化の無いリスト（列）を非表示にできます。',
          [('D3:E4', '▶ 変化なしを隠す'), ('G3:H4', '▶ 全て表示')], a4='状態', h2=50)
legend(ws, 14, [('追加', 'C6EFCE'), ('削除', 'D9D9D9'), ('変更', 'FFFF99'), ('NG', 'FFC7CE')])
wide_frame(ws, [(5, 'name'), (6, '判定'), (7, '件数'), (8, '値 →')], ncol=12)

# ---- 古いシートを削除・並べ替え・番号 ----
for n in ('URLCat_横長ビュー', 'URLCat_差分チェック'):
    if n in wb.sheetnames:
        del wb[n]
order = ['基本設定', EXL, LREG, CEX, CREG, 'JSON_CURL', DLIST, DCONT, '実行ログ']
wb._sheets = [wb[n] for n in order]
for nm, old, new in [('JSON_CURL', '⑤', '⑥'), ('実行ログ', '⑥', '⑨')]:
    a = wb[nm]['A1']
    assert isinstance(a.value, str) and a.value.startswith(old), (nm, a.value)
    a.value = new + a.value[1:]
j = wb['JSON_CURL']
j['A2'] = ('登録更新シートの「JSON発行」ボタンでVBAが10行目以降を作り直します（緑=VBA書込み）。順番は リストの登録・変更 → 中身だけの変更 → リストの削除。'
           'curlはbash形式で、先に5行目でトークン取得→6行目でTOKEN変数を設定してから実行します。JSONが30,000文字を超える行はJ列以降に分割して保存し、'
           'curlは「JSONファイル出力」で作る body_No.json を読み込む形にしています。「API実行」ボタンでExcelから直接実行もできます'
           '（結果は実行ログへ。実行後はSCMから再取得し、差分チェック（リスト一覧）のM列「実行結果」に反映結果を出します）。')
cfg = wb['基本設定']
cfg['A25'] = 'ホワイトリスト／ブラックリスト（登録更新・差分チェックの先頭に並べるリスト名）'
cfg['D26'] = '登録更新・差分チェックの先頭に並べます（空欄でもかまいません）'
cfg['D28'] = '登録更新・差分チェックの2番目に並べます（空欄でもかまいません）'
cfg['D31'] = '1リストあたりの上限件数'
cfg['D35'] = '登録更新（リスト内容）に中身を書かずに登録したリストは、この1件で作成します（URLリストは空にできないため）'

wb.active = 0
for s in wb: s.sheet_view.tabSelected = False
wb.active.sheet_view.tabSelected = True
wb.save(OUT)
print('saved', OUT, wb.sheetnames)

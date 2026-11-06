# -*- coding: utf-8 -*-
# URL既存一覧・差分確認（WL/BL以外）シートを削除し、差分を横長ビューに表示する形にする（2026-10-29 ユーザー指示）
#   差分の基準はシートに持たず、入力チェック・JSON発行のたびにSCMから取得する
#   順番：update_urlcat_longformat.py → patch_url_rules.py → patch_wide_edit.py → patch_op_wide.py → このスクリプト
#   使い方: python3 patch_nouex.py 入力.xlsx 出力.xlsx
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
SPARE = 30  # VBA の VW_SPARE と同じ

for nm in ['URLCat_URL既存一覧', 'URLCat_差分確認']:
    if nm in wb.sheetnames:
        del wb[nm]
for nm in ['UL_FETCHED_AT']:
    if nm in wb.defined_names:
        del wb.defined_names[nm]

# ---- 番号の振り直し ----
for nm, old, new in [('URLCat_横長ビュー', '④', '③'),
                     ('URLCat_WL_既存一覧', '⑤', '④'), ('URLCat_WL_記載', '⑥', '⑤'), ('URLCat_BL_既存一覧', '⑦', '⑥'),
                     ('URLCat_BL_記載', '⑧', '⑦'), ('JSON_CURL', '⑨', '⑧'),
                     ('URLCat_WL_差分確認', '⑪', '⑨'), ('URLCat_BL_差分確認', '⑫', '⑩'), ('実行ログ', '⑬', '⑪')]:
    a = wb[nm]['A1']
    assert isinstance(a.value, str) and a.value.startswith(old), (nm, a.value)
    a.value = new + a.value[1:]

# ---- 説明文 ----
e = wb['URLCat_既存一覧']
e['A2'] = e['A2'].value.replace('同時に URL既存一覧・横長ビュー・ホワイトリスト', '同時に 横長ビュー・ホワイトリスト')
j = wb['JSON_CURL']
j['A2'] = j['A2'].value.replace('（結果は実行ログへ、実行後は差分確認シートで反映結果を確認）',
                                '（結果は実行ログへ。実行後はSCMから再取得し、横長ビューの12行目「実行結果」とWL/BLの差分確認シートに反映結果を出します）')

# ---- 横長ビュー：10～12行目に差分・削除される値・実行結果 ----
ws = wb['URLCat_横長ビュー']
ws['A2'] = ('「既存情報取得」でSCMの現在の内容に作り直されます。13行目以降の値を編集します（追加は列の下に追記、削除はセルを消す。空きセルは詰めなくてかまいません）。'
            'カテゴリの登録は空いている列に name（6行目）を書いて操作（5行目）を「登録」に、削除は操作を「削除」にします。type（7行目）・description（9行目）を書き換えると変更になります。'
            '「入力チェック」「JSON発行」はその時点のSCMの内容と比較し、10行目に差分（新規=緑・変更=黄・削除=グレー、+追加/-削除の件数）、11行目に削除される値を出します。'
            'セルの色：NG=赤、警告=黄、追加する値=緑、変わる type/description=橙。理由はセルのコメント。API実行後は12行目に反映結果（OK=緑、要確認=赤）を出します。'
            '書式：*.aaa.com/ を推奨。aa*.aaa.com・*.*.com はNG。末尾 / なし（前方一致）・途中の*・パス指定（/以降）は警告')
ws.row_dimensions[2].height = 100
ws['J3'] = '▶ SCMから取得し直す'
ws.merge_cells('M3:N4')
c = ws['M3']
c.value = '▶ 実行結果の再確認'
c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.fill = fill('C55A11')
c.alignment = Alignment(horizontal='center', vertical='center')
for r, t in [(10, '差分'), (11, '削除される値'), (12, '実行結果'), (13, 'エントリ →')]:
    c = ws.cell(r, 1, t)
    c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.fill = fill('2F5597'); c.border = BOX
    c.alignment = Alignment(vertical='top', wrap_text=True)
last = 1 + SPARE
for col in range(2, last + 1):
    L = get_column_letter(col)
    ws.cell(8, col).value = f'=IF({L}6="","",COUNTA({L}13:{L}1048576))'
    for r in (10, 11, 12):
        c = ws.cell(r, col)
        c.fill = fill('E2EFDA'); c.border = BOX; c.font = Font(name=FN, size=10)
        c.alignment = Alignment(wrap_text=True, vertical='top')
ws.row_dimensions[11].height = 48
ws.freeze_panes = 'B13'

wb.active = 0
for s in wb: s.sheet_view.tabSelected = False
wb.active.sheet_view.tabSelected = True
wb.save(OUT)
print('saved', OUT, wb.sheetnames)

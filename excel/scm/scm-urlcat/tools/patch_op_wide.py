# -*- coding: utf-8 -*-
# 横長ビューでカテゴリの登録・削除・type/description変更も行う形にする（2026-10-29 ユーザー指示）
#   5行目に「操作」（登録／削除）を追加し、③カテゴリ登録変更削除シートを削除する
#   順番：update_urlcat_longformat.py → patch_url_rules.py → patch_wide_edit.py → このスクリプト
#   使い方: python3 patch_op_wide.py 入力.xlsx 出力.xlsx
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
SPARE = 30  # VBA の VW_SPARE と同じ

# ---- ③ を削除 ----
if 'URLCat_カテゴリ登録変更削除' in wb.sheetnames:
    del wb['URLCat_カテゴリ登録変更削除']

# ---- 番号の振り直し ----
for nm, old, new in [('URLCat_URL既存一覧', '④', '③'), ('URLCat_横長ビュー', '⑤', '④'),
                     ('URLCat_WL_既存一覧', '⑥', '⑤'), ('URLCat_WL_記載', '⑦', '⑥'), ('URLCat_BL_既存一覧', '⑧', '⑦'),
                     ('URLCat_BL_記載', '⑨', '⑧'), ('JSON_CURL', '⑩', '⑨'), ('URLCat_差分確認', '⑪', '⑩'),
                     ('URLCat_WL_差分確認', '⑫', '⑪'), ('URLCat_BL_差分確認', '⑬', '⑫'), ('実行ログ', '⑭', '⑬')]:
    a = wb[nm]['A1']
    assert isinstance(a.value, str) and a.value.startswith(old), (nm, a.value)
    a.value = new + a.value[1:]

# ---- 説明文 ----
wb['基本設定']['D35'] = '横長ビューで値を書かずに登録したカテゴリは、この1件で作成します（URLリストは空にできないため）'
wb['URLCat_URL既存一覧']['A2'] = '「既存情報取得」でVBAが作り直します（直接編集しないでください）。④横長ビューとの差分比較の基準（実行前の状態）になります。'
j = wb['JSON_CURL']
j['A2'] = j['A2'].value.replace('（カテゴリ登録変更削除・横長ビュー・WL/BL記載のどのシートからでも同じ）', '（横長ビュー・WL/BL記載のどのシートからでも同じ）')
d = wb['URLCat_差分確認']
d['A2'] = d['A2'].value.replace('（既存＝④URL既存一覧、反映後＝⑤横長ビュー・③）', '（既存＝③URL既存一覧、反映後＝④横長ビュー）')

# ---- 横長ビュー：5行目に操作を追加 ----
ws = wb['URLCat_横長ビュー']
ws['A2'] = ('「既存情報取得」でSCMの現在の内容に作り直されます。10行目以降の値を編集します（追加は列の下に追記、削除はセルを消す。空きセルは詰めなくてかまいません）。'
            'カテゴリの登録は空いている列に name（6行目）を書いて操作（5行目）を「登録」に、削除は操作を「削除」にします。'
            'type（7行目）・description（9行目）を書き換えると変更になります。登録時に値が無ければダミー1件（基本設定）で作成します。'
            '「入力チェック」「JSON発行」でNGは赤、警告は黄、追加する値は緑、変わる type/description は橙色になり、理由はセルのコメントに出ます。値が1件も無い列のカテゴリは変更しません。'
            '書式：*.aaa.com/ を推奨。aa*.aaa.com・*.*.com はNG。末尾 / なし（前方一致）・途中の*・パス指定（/以降）は警告')
ws.row_dimensions[2].height = 85
# 5行目以降（B列～）を消して作り直す
for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
    for c in row:
        if c.column >= 2:
            c.value = None; c.fill = PatternFill(); c.border = Border(); c.font = Font(name=FN, size=10)
for r, t in [(5, '操作'), (6, 'name'), (7, 'type'), (8, '件数'), (9, 'description'), (10, 'エントリ →')]:
    c = ws.cell(r, 1, t)
    c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.fill = fill('2F5597'); c.border = BOX
    c.alignment = Alignment(vertical='top')
last = 1 + SPARE
for col in range(2, last + 1):
    L = get_column_letter(col)
    ws.column_dimensions[L].width = 30
    for r in range(5, 10):
        c = ws.cell(r, col)
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical='top')
        if r == 6:
            c.fill = fill('2F5597'); c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.number_format = '@'
        elif r == 8:
            c.fill = fill('D9E1F2'); c.font = Font(name=FN, size=10)
            c.value = f'=IF({L}6="","",COUNTA({L}10:{L}1048576))'
        else:
            c.fill = fill('FFF2CC'); c.font = Font(name=FN, size=10); c.number_format = '@'
ws.data_validations.dataValidation = []
dv1 = DataValidation(type='list', formula1='"登録,削除"', allow_blank=True)
dv1.add(f'B5:{get_column_letter(last)}5')
dv2 = DataValidation(type='list', formula1='"URL List,Category Match"', allow_blank=True)
dv2.add(f'B7:{get_column_letter(last)}7')
ws.add_data_validation(dv1); ws.add_data_validation(dv2)
ws.freeze_panes = 'B10'

wb.active = 0
for s in wb: s.sheet_view.tabSelected = False
wb.active.sheet_view.tabSelected = True
wb.save(OUT)
print('saved', OUT, wb.sheetnames)

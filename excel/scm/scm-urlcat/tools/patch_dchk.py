# -*- coding: utf-8 -*-
# 差分チェックシート（横長：追加=緑、削除=グレー）を追加し、横長ビューは編集専用に戻す（2026-10-29 ユーザー指示）
#   順番：… → patch_op_wide.py → patch_nouex.py → このスクリプト
#   使い方: python3 patch_dchk.py 入力.xlsx 出力.xlsx
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
SPARE = 30
def btn(ws, rng, t):
    ws.merge_cells(rng)
    c = ws[rng.split(':')[0]]
    c.value = t
    c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.fill = fill('C55A11')
    c.alignment = Alignment(horizontal='center', vertical='center')
def label(ws, r, t):
    c = ws.cell(r, 1, t)
    c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.fill = fill('2F5597'); c.border = BOX
    c.alignment = Alignment(vertical='top', wrap_text=True)

# ---- 横長ビュー：10～12行目の差分行をやめ、10行目からエントリに戻す ----
ws = wb['URLCat_横長ビュー']
ws['A2'] = ('「既存情報取得」でSCMの現在の内容に作り直されます。10行目以降の値を編集します（追加は列の下に追記、削除はセルを消す。空きセルは詰めなくてかまいません）。'
            'カテゴリの登録は空いている列に name（6行目）を書いて操作（5行目）を「登録」に、削除は操作を「削除」にします。type（7行目）・description（9行目）を書き換えると変更になります。'
            '「差分チェック」でその時点のSCMの内容と比べ、結果を差分チェックシートに出します（追加=緑、削除=グレー）。NG=赤・警告=黄はこのシートのセルに出し、理由はセルのコメントに書きます。'
            '書式：*.aaa.com/ を推奨。aa*.aaa.com・*.*.com はNG。末尾 / なし（前方一致）・途中の*・パス指定（/以降）は警告')
ws.row_dimensions[2].height = 85
ws['G3'] = '▶ 差分チェック'
if 'M3:N4' in [str(r) for r in ws.merged_cells.ranges]:
    ws.unmerge_cells('M3:N4')
for a in ('M3', 'N3', 'M4', 'N4'):
    ws[a].value = None; ws[a].fill = PatternFill(); ws[a].font = Font(name=FN, size=10)
label(ws, 10, 'エントリ →')
for r in (11, 12, 13):
    c = ws.cell(r, 1); c.value = None; c.fill = PatternFill(); c.border = Border(); c.font = Font(name=FN, size=10)
for col in range(2, 2 + SPARE):
    L = get_column_letter(col)
    ws.cell(8, col).value = f'=IF({L}6="","",COUNTA({L}10:{L}1048576))'
    for r in (10, 11, 12):
        c = ws.cell(r, col); c.fill = PatternFill(); c.border = Border(); c.font = Font(name=FN, size=10)
        c.alignment = Alignment(vertical='top')
ws.row_dimensions[11].height = None
ws.freeze_panes = 'B10'

# ---- 差分チェックシートを横長ビューの次に追加 ----
NAME = 'URLCat_差分チェック'
if NAME in wb.sheetnames:
    del wb[NAME]
d = wb.create_sheet(NAME, wb.sheetnames.index('URLCat_横長ビュー') + 1)
d.sheet_view.showGridLines = False
d.merge_cells('A1:L1'); d.merge_cells('A2:L2')
d['A1'] = '④ 差分チェック（横長：1列1カテゴリ。SCMの現在の内容と横長ビューの比較）'
d['A1'].font = Font(name=FN, size=14, bold=True, color='FFFFFF'); d['A1'].fill = fill('1F3864')
d.row_dimensions[1].height = 26
d['A2'] = ('横長ビューの「差分チェック」「JSON発行」でVBAが作り直します（直接編集しないでください）。SCMから取得した現在の値をSCMの順に並べ、'
           '削除される値はグレー、その下に追加される値を緑で表示します。6行目の判定は 新規=緑、変更=黄、削除=グレー、変化なし=色なし。'
           '8行目の type/description の変更は橙色。「API実行」のあとはSCMから取得し直し、9行目「実行結果」に OK=緑、要確認=赤 を書きます。'
           '「変化なしを隠す」で変化の無いカテゴリ（列）を非表示にできます。')
d['A2'].font = Font(name=FN, size=9, color='595959'); d['A2'].alignment = Alignment(wrap_text=True, vertical='top')
d.row_dimensions[2].height = 60
for a, t in [('A3', '作成日時'), ('A4', '状態')]:
    d[a] = t; d[a].font = Font(name=FN, size=10, bold=True)
for a in ('B3', 'B4'):
    d[a].font = Font(name=FN, size=10, bold=True); d[a].fill = fill('E2EFDA')
d['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
btn(d, 'D3:E4', '▶ 変化なしを隠す')
btn(d, 'G3:H4', '▶ 全て表示')
btn(d, 'J3:K4', '▶ 実行結果の再確認')
d['M3'] = '凡例：'; d['M3'].font = Font(name=FN, size=10, bold=True)
for k, (t, c) in enumerate([('追加', 'C6EFCE'), ('削除', 'D9D9D9'), ('変更', 'FFFF99'), ('type/desc変更', 'FFC000'), ('要確認・NG', 'FFC7CE')]):
    x = d.cell(3, 14 + k, t)
    x.font = Font(name=FN, size=10, bold=True); x.fill = fill(c); x.border = BOX
    x.alignment = Alignment(horizontal='center')
for r, t in [(5, 'name'), (6, '判定'), (7, '件数'), (8, 'type・description'), (9, '実行結果'), (10, '値 →')]:
    label(d, r, t)
d.column_dimensions['A'].width = 14
for col in range(2, 2 + SPARE):
    d.column_dimensions[get_column_letter(col)].width = 30
d.freeze_panes = 'B10'

# ---- 番号の振り直し・説明文 ----
for nm, old, new in [('URLCat_WL_既存一覧', '④', '⑤'), ('URLCat_WL_記載', '⑤', '⑥'), ('URLCat_BL_既存一覧', '⑥', '⑦'),
                     ('URLCat_BL_記載', '⑦', '⑧'), ('JSON_CURL', '⑧', '⑨'),
                     ('URLCat_WL_差分確認', '⑨', '⑩'), ('URLCat_BL_差分確認', '⑩', '⑪'), ('実行ログ', '⑪', '⑫')]:
    a = wb[nm]['A1']
    assert isinstance(a.value, str) and a.value.startswith(old), (nm, a.value)
    a.value = new + a.value[1:]
j = wb['JSON_CURL']
j['A2'] = j['A2'].value.replace('横長ビューの12行目「実行結果」', '差分チェックシートの9行目「実行結果」')

wb.active = 0
for s in wb: s.sheet_view.tabSelected = False
wb.active.sheet_view.tabSelected = True
wb.save(OUT)
print('saved', OUT, wb.sheetnames)

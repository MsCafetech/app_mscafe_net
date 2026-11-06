# -*- coding: utf-8 -*-
# WL/BL専用シートを削除し、WL/BLも横長ビューの列として扱う（2026-10-29 ユーザー指示）
#   順番：… → patch_nouex.py → patch_dchk.py → このスクリプト
#   使い方: python3 patch_nolist.py 入力.xlsx 出力.xlsx
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
FN = 'Meiryo UI'

for lk in ('WL', 'BL'):
    for part in ('既存一覧', '記載', '差分確認'):
        nm = f'URLCat_{lk}_{part}'
        if nm in wb.sheetnames:
            del wb[nm]
for nm in list(wb.defined_names.keys()):
    if nm.startswith(('WLEX_', 'BLEX_')) or nm in ('CFG_WL_DESC', 'CFG_BL_DESC'):
        del wb.defined_names[nm]

# ---- 基本設定 ----
cfg = wb['基本設定']
cfg['A25'] = 'ホワイトリスト／ブラックリスト（横長ビューの先頭列に並べるカテゴリ名）'
cfg['D26'] = '横長ビュー・差分チェックの先頭列に並べます（空欄でもかまいません）'
cfg['D28'] = '横長ビュー・差分チェックの2列目に並べます（空欄でもかまいません）'
for r in (27, 29):
    for c in range(1, 5):
        x = cfg.cell(r, c)
        x.value = None; x.fill = PatternFill(); x.border = Border(); x.font = Font(name=FN, size=10)
cfg['D31'] = '横長ビューの1カテゴリあたりの上限件数'

# ---- 番号・説明文 ----
for nm, old, new in [('JSON_CURL', '⑨', '⑤'), ('実行ログ', '⑫', '⑥')]:
    a = wb[nm]['A1']
    assert isinstance(a.value, str) and a.value.startswith(old), (nm, a.value)
    a.value = new + a.value[1:]
wb['URLCat_横長ビュー']['A1'] = '③ 横長ビュー（編集用：1列1カテゴリ。基本設定のホワイトリスト／ブラックリストを先頭に並べます）'
e = wb['URLCat_既存一覧']
e['A2'] = e['A2'].value.replace('同時に 横長ビュー・ホワイトリスト／ブラックリストの既存一覧シートも更新します', '同時に 横長ビューも作り直します')
j = wb['JSON_CURL']
j['A2'] = ('横長ビューの「JSON発行」ボタンでVBAが10行目以降を作り直します（緑=VBA書込み）。順番は カテゴリ登録・変更 → URLだけの変更 → カテゴリ削除。'
           'curlはbash形式で、先に5行目でトークン取得→6行目でTOKEN変数を設定してから実行します。JSONが30,000文字を超える行はJ列以降に分割して保存し、'
           'curlは「JSONファイル出力」で作る body_No.json を読み込む形にしています。「API実行」ボタンでExcelから直接実行もできます'
           '（結果は実行ログへ。実行後はSCMから再取得し、差分チェックシートの9行目「実行結果」に反映結果を出します）。')

wb.active = 0
for s in wb: s.sheet_view.tabSelected = False
wb.active.sheet_view.tabSelected = True
wb.save(OUT)
print('saved', OUT, wb.sheetnames)

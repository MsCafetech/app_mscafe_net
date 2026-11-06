# -*- coding: utf-8 -*-
# 横長ビューを編集用にし、URL記載シートを削除する（2026-10-29 ユーザー指示）
#   ③の登録はURLが無い前提でダミー1件（基本設定 CFG_DUMMY）で作成
#   順番：update_urlcat_longformat.py → patch_url_rules.py → このスクリプト
#   使い方: python3 patch_wide_edit.py 入力.xlsx 出力.xlsx
import sys, re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
DUMMY = 'dummy.fqdn/'

# ---- URL記載シートを削除 ----
if 'URLCat_URL記載' in wb.sheetnames:
    del wb['URLCat_URL記載']
if 'CAT_NAMES' in wb.defined_names:
    del wb.defined_names['CAT_NAMES']

# ---- 基本設定：ダミー ----
cfg = wb['基本設定']
cfg['A35'] = '空カテゴリ用ダミー'
cfg['A35'].font = Font(name=FN, size=10)
cfg['B35'] = DUMMY
cfg['B35'].font = Font(name=FN, size=10, color='0000FF'); cfg['B35'].fill = fill('FFF2CC'); cfg['B35'].border = BOX
cfg['C35'] = '任意'
cfg['C35'].font = Font(name=FN, size=9, bold=True, color='FFFFFF'); cfg['C35'].fill = fill('7F7F7F')
cfg['C35'].alignment = Alignment(horizontal='center')
cfg['D35'] = '③で登録するカテゴリはURLが無い前提のため、この1件で作成します（URLリストは空にできないため）'
cfg['D35'].font = Font(name=FN, size=10)
wb.defined_names['CFG_DUMMY'] = DefinedName('CFG_DUMMY', attr_text="'基本設定'!$B$35")

# ---- ③ カテゴリ登録変更削除 ----
ws = wb['URLCat_カテゴリ登録変更削除']
ws['A2'] = ('黄色(B〜E列)に入力。カテゴリそのものの新規作成・削除と、type・descriptionの変更に使います。'
            '登録はURLが無い前提で、基本設定の「空カテゴリ用ダミー」1件で作成します。作成後に「既存情報取得」すると⑤横長ビューに列ができるので、そこへURLを追記します。'
            'URLの追加・削除だけならこのシートへの記載は不要です。G〜O列は関数で既存一覧と自動照合し、判定がNGの行はJSON発行の対象外。'
            '変更時に既存と異なる項目は橙色で表示します（descriptionの空欄は既存値が消えます）。')
ws['F6'] = '※ typeは登録/変更時に必須（削除時は不要）。登録時のエントリはF列のダミー1件です'
ws['F7'] = '作成時のエントリ'
for r in range(8, 508):
    ws[f'F{r}'] = f'=IF(AND($B{r}="登録",$C{r}<>""),CFG_DUMMY,"")'
    f = ws[f'N{r}'].value
    for seg in [f'IF(AND($B{r}="登録",N($F{r})=0),"NG:④URL記載シートにこのカテゴリの行がありません（1件以上必要）",',
                f'IF(AND($B{r}="削除",N($F{r})>0),"警告:④URL記載シートにある "&$F{r}&" 行は無視されます（行も消しておくと分かりやすいです）",']:
        if seg in f:
            assert f.endswith(')')
            f = f.replace(seg, '')[:-1]
    f = f.replace('（FQDNだけの変更ならこの行は不要）', '（URLだけの変更ならこの行は不要）')
    ws[f'N{r}'] = f

# ---- ④ URL既存一覧 ----
ws = wb['URLCat_URL既存一覧']
ws['A1'] = '④ URL既存一覧（SCMから取得したWL/BL以外の全エントリ：カテゴリ名 | FQDN、1行1件）'
ws['A2'] = '「既存情報取得」でVBAが作り直します（直接編集しないでください）。⑤横長ビューとの差分比較の基準（実行前の状態）になります。'
ws['E7'] = None
ws['E7'].fill = PatternFill(); ws['E7'].border = Border()

# ---- ⑤ 横長ビュー（編集用） ----
ws = wb['URLCat_横長ビュー']
ws['A1'] = '⑤ 横長ビュー（編集用：1列1カテゴリ、WL/BLは含みません）'
ws['A2'] = ('「既存情報取得」でSCMの現在の内容に作り直されます。9行目以降の値を編集します（追加は列の下に追記、削除はセルを消す。空きセルは詰めなくてかまいません）。'
            '5〜8行目（name・type・件数・description）は表示のみです。新しいカテゴリは③で登録（ダミー1件で作成）→ API実行 → 既存情報取得で列ができます。'
            '「入力チェック」「JSON発行」でNGは赤、警告は黄色になり、理由はセルのコメントに出ます。値が1件も無い列のカテゴリは変更しません。'
            '書式：*.aaa.com/ を推奨。aa*.aaa.com・*.*.com はNG。末尾 / なし（前方一致）・途中の*・パス指定（/以降）は警告')
ws.row_dimensions[2].height = 70
for rng, t in [('D3:E4', '▶ JSON発行 → JSON_CURLへ'), ('G3:H4', '▶ 入力チェック'), ('J3:K4', '▶ 既存から作り直す')]:
    ws.merge_cells(rng)
    c = ws[rng.split(':')[0]]
    c.value = t
    c.font = Font(name=FN, size=10, bold=True, color='FFFFFF'); c.fill = fill('C55A11')
    c.alignment = Alignment(horizontal='center', vertical='center')

# ---- 番号の振り直し・説明文 ----
for nm, old, new in [('URLCat_WL_既存一覧', '⑦', '⑥'), ('URLCat_WL_記載', '⑧', '⑦'), ('URLCat_BL_既存一覧', '⑨', '⑧'),
                     ('URLCat_BL_記載', '⑩', '⑨'), ('JSON_CURL', '⑪', '⑩'), ('URLCat_差分確認', '⑫', '⑪'),
                     ('URLCat_WL_差分確認', '⑬', '⑫'), ('URLCat_BL_差分確認', '⑭', '⑬'), ('実行ログ', '⑮', '⑭')]:
    a = wb[nm]['A1']
    if isinstance(a.value, str) and a.value.startswith(old):
        a.value = new + a.value[1:]
d = wb['URLCat_差分確認']
d['A2'] = d['A2'].value.replace('（既存＝⑤URL既存一覧、反映後＝④URL記載・③）', '（既存＝④URL既存一覧、反映後＝⑤横長ビュー・③）')
j = wb['JSON_CURL']
j['A2'] = j['A2'].value.replace('カテゴリ登録変更削除・URL記載・WL/BL記載', 'カテゴリ登録変更削除・横長ビュー・WL/BL記載').replace('FQDNの変更', 'URLの変更')
e = wb['URLCat_既存一覧']
e['A2'] = ('ボタンを押すとAPI（GET）でJSONを取得し、8行目以降に表を再生成します（B〜H列=VBA書込み、A・I列=関数）。listは1行1件（セル内改行）で表示します。'
           '同時に URL既存一覧・横長ビュー・ホワイトリスト／ブラックリストの既存一覧シートも更新します（横長ビューに未反映の編集があるときは確認します）。直接編集はしないでください。')

wb.active = 0
for s in wb: s.sheet_view.tabSelected = False
wb.active.sheet_view.tabSelected = True
wb.save(OUT)
print('saved', OUT, wb.sheetnames)

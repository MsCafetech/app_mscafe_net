# -*- coding: utf-8 -*-
# SCM_URLCategory.xlsx を縦長形式（カテゴリ名 | FQDN）に変更する
#   WL/BL の専用シートはそのまま残し、それ以外のカスタムURLカテゴリを
#   URLCat_URL記載（1行1件）で管理する形に置き換える
#   使い方: python3 update_urlcat_longformat.py 入力.xlsx 出力.xlsx
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

SRC, OUT = sys.argv[1], sys.argv[2]

SE = 'URLCat_既存一覧'
SC = 'URLCat_カテゴリ登録変更削除'
SU = 'URLCat_URL記載'
SX = 'URLCat_URL既存一覧'
SV = 'URLCat_横長ビュー'
SD = 'URLCat_差分確認'
FIRST, LAST, EXL = 8, 507, 5000
UL = 100007          # URL記載：プルダウン・書式を付ける最終行（これより下も読み込みます）

FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
F_TITLE, F_HEAD, F_IN, F_FX, F_VBA, F_BTN = fill('1F3864'), fill('2F5597'), fill('FFF2CC'), fill('F2F2F2'), fill('E2EFDA'), fill('C55A11')
F_REQ, F_OPT, F_SEC = fill('C00000'), fill('7F7F7F'), fill('D9E1F2')

wb = openpyxl.load_workbook(SRC)
for n in ['URLCat_登録変更削除', SD]:
    if n in wb.sheetnames:
        del wb[n]

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

def heads(ws, row, names, widths, start=1):
    for i, (n, w) in enumerate(zip(names, widths), start):
        put(ws, f'{CL(i)}{row}', n, Font(name=FN, size=10, bold=True, color='FFFFFF'), F_HEAD,
            Alignment(horizontal='center', vertical='center', wrap_text=True))
        if w: ws.column_dimensions[CL(i)].width = w
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

def colstyle(ws, col, fl=None, nf=None, font=None):
    d = ws.column_dimensions[col]
    if fl: d.fill = fl
    if nf: d.number_format = nf
    if font: d.font = font

Q = lambda s: "'" + s + "'"
def nest(conds, last):
    f = last
    for cond, val in reversed(conds):
        f = f'IF({cond},{val},{f})'
    return f

# =========================================================
# 基本設定・既存一覧・JSON_CURL・実行ログ（文言だけ変更）
# =========================================================
cfg = wb['基本設定']
cfg['D31'] = 'WL/BL記載シートの入力行数（8〜2507行目）、URL記載シートの1カテゴリあたりの上限件数'

ex = wb[SE]
ex['A1'] = '② 既存情報（SCMから取得したカスタムURLカテゴリ一覧）'
ex['A2'] = ('ボタンを押すとAPI（GET）でJSONを取得し、8行目以降に表を再生成します（B〜H列=VBA書込み、A・I列=関数）。listは1行1件（セル内改行）で表示します。'
            '同時に URL既存一覧・横長ビュー・ホワイトリスト／ブラックリストの既存一覧シートも更新し、URL記載シートが空なら既存の内容を書き出します。直接編集はしないでください。')

js = wb['JSON_CURL']
js['A1'] = '⑪ 生成JSON・CURLコマンド'
js['A2'] = ('「JSON発行」ボタン（カテゴリ登録変更削除・URL記載・WL/BL記載のどのシートからでも同じ）でVBAが10行目以降を作り直します（緑=VBA書込み）。'
            '順番は カテゴリ登録・変更とFQDNの変更 → ホワイトリスト → ブラックリスト → カテゴリ削除。curlはbash形式で、先に5行目でトークン取得→6行目でTOKEN変数を設定してから実行します。'
            'JSONが30,000文字を超える行はJ列以降に分割して保存し、curlは「JSONファイル出力」で作る body_No.json を読み込む形にしています。'
            '「API実行」ボタンでExcelから直接実行もできます（結果は実行ログへ、実行後は差分確認シートで反映結果を確認）。')
wb['実行ログ']['A1'] = '⑮ 実行ログ（API実行結果・Push・ジョブ状況・差分確認）'
for nm, t in [('URLCat_WL_既存一覧', '⑦ ホワイトリスト 既存情報（SCMから取得したエントリ一覧：1行1件）'),
              ('URLCat_WL_記載', '⑧ ホワイトリスト 記載（反映後の全エントリを1行1件で記載 → 既存と差分比較）'),
              ('URLCat_BL_既存一覧', '⑨ ブラックリスト 既存情報（SCMから取得したエントリ一覧：1行1件）'),
              ('URLCat_BL_記載', '⑩ ブラックリスト 記載（反映後の全エントリを1行1件で記載 → 既存と差分比較）'),
              ('URLCat_WL_差分確認', '⑬ 差分確認（ホワイトリスト）：現在のエントリ一覧と既存一覧の比較'),
              ('URLCat_BL_差分確認', '⑭ 差分確認（ブラックリスト）：現在のエントリ一覧と既存一覧の比較')]:
    wb[nm]['A1'] = t
for lk in ['WL', 'BL']:
    wb[f'URLCat_{lk}_記載']['I3'] = '▶ JSON発行 → JSON_CURLへ'
    d = wb[f'URLCat_{lk}_差分確認']
    d['A2'] = d['A2'].value.replace('④のAPI実行', 'JSON_CURLのAPI実行')

# =========================================================
# ③ カテゴリ登録変更削除（name / type / description。FQDNはURL記載シート）
# =========================================================
ws = sheet(SC, 2)
header(ws, '③ カテゴリの登録・変更・削除（ホワイトリスト/ブラックリスト以外）',
       '黄色(B〜E列)に入力。カテゴリそのものの新規作成・削除と、type・descriptionの変更に使います。FQDN（list）は④URL記載シートに「カテゴリ名 | FQDN」で書きます。'
       '登録は④URL記載シートに同じカテゴリ名の行が1件以上必要です。FQDNだけを追加・削除するときは、このシートへの記載は不要です。'
       'F〜O列は関数で既存一覧と自動照合し、判定がNGの行はJSON発行の対象外。変更時に既存と異なる項目は橙色で表示します（descriptionの空欄は既存値が消えます）。', 15, 56)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(M{FIRST}:M{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(M{FIRST}:M{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(M{FIRST}:M{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行 → JSON_CURLへ')
btn(ws, 'I3:J4', '▶ 入力クリア')
labels(ws, 6, {'B': '必須', 'C': '必須', 'D': '必須※', 'E': '任意'})
put(ws, 'F6', '※ typeは登録/変更時に必須（削除時は不要）。Category Matchのときは④URL記載に定義済みカテゴリ名を書きます', Font(name=FN, size=9), border=False)
heads(ws, 7, ['No', '操作', 'name', 'type', 'description', 'URL記載の件数', '既存有無', '既存id', '既存type', '既存件数',
              '既存description', '一覧内重複', '判定', '理由', '作業:既存行'],
      [5, 8, 24, 14, 30, 10, 9, 22, 13, 9, 22, 10, 8, 48, 10])
UC = f"{Q(SU)}!$A${FIRST}:$A${UL}"
EXC = f"{Q(SE)}!$C${FIRST}:$C${EXL}"
NAMECHK = ('SUMPRODUCT(--ISERROR(FIND(MID($C{r}&REPT(" ",63),ROW($1:$63),1),'
           '" 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")))>0')
for r in range(FIRST, LAST + 1):
    for col in 'BCDE':
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
    fx = {
        'A': f'=IF(C{r}="","",ROW()-7)',
        'F': f'=IF($C{r}="","",COUNTIF({UC},$C{r}))',
        'G': f'=IF($C{r}="","",IF($O{r}>0,"あり","なし"))',
        'H': f'=IF($O{r}>0,INDEX({Q(SE)}!$B:$B,$O{r})&"","")',
        'I': f'=IF($O{r}>0,INDEX({Q(SE)}!$D:$D,$O{r})&"","")',
        'J': f'=IF($O{r}>0,INDEX({Q(SE)}!$F:$F,$O{r}),"")',
        'K': f'=IF($O{r}>0,INDEX({Q(SE)}!$G:$G,$O{r})&"","")',
        'L': f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,IF(SUMPRODUCT(--EXACT($C${FIRST}:$C${LAST},$C{r}))>1,"重複",""),""))',
        'M': f'=IF($N{r}="","",IF($N{r}="OK","OK",LEFT($N{r},FIND(":",$N{r})-1)))',
        'N': '=' + nest([
            (f'$C{r}=""', '""'),
            (f'$B{r}=""', '"NG:操作未選択"'),
            (f'$L{r}="重複"', '"NG:一覧内でnameが重複"'),
            (f'LEN($C{r})>63', '"NG:nameが63文字超"'),
            (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
            (NAMECHK.format(r=r), '"NG:nameに使用不可の文字(英数字 空白 . _ - のみ)"'),
            (f'OR(AND(CFG_WL_NAME<>"",EXACT($C{r},CFG_WL_NAME)),AND(CFG_BL_NAME<>"",EXACT($C{r},CFG_BL_NAME)))',
             '"NG:ホワイトリスト/ブラックリストは専用の記載シートで操作してください"'),
            (f'AND($B{r}="登録",$G{r}="あり")', '"NG:既存に同名あり(登録不可)"'),
            (f'AND($B{r}<>"登録",$G{r}="なし")', '"NG:既存エントリ無"'),
            (f'AND($B{r}<>"削除",$D{r}="")', '"NG:typeが未選択"'),
            (f'AND($B{r}="登録",N($F{r})=0)', '"NG:④URL記載シートにこのカテゴリの行がありません（1件以上必要）"'),
            (f'AND($B{r}<>"削除",LEN($E{r})>255)', '"NG:descriptionが255文字超"'),
            (f'AND($B{r}="削除",N($F{r})>0)', '"警告:④URL記載シートにある "&$F' + str(r) + '&" 行は無視されます（行も消しておくと分かりやすいです）"'),
            (f'AND($B{r}="変更",$D{r}<>$I{r})', '"警告:typeを変更します（"&$I' + str(r) + '&"→"&$D' + str(r) + '&"）"'),
            (f'AND($B{r}="変更",EXACT($D{r},$I{r}),EXACT($E{r}&"",$K{r}))',
             '"警告:type・descriptionが既存と同じです（FQDNだけの変更ならこの行は不要）"'),
        ], '"OK"'),
        'O': f'=IF($C{r}="",0,IF(COUNTIF({EXC},$C{r})=0,0,SUMPRODUCT(MAX(ROW({EXC})*EXACT({EXC},$C{r})))))',
    }
    for col, f in fx.items():
        put(ws, f'{col}{r}', f, Font(name=FN, size=10), F_FX)
ws.freeze_panes = 'D8'
dv = DataValidation(type='list', formula1='"登録,変更,削除"', allow_blank=True); ws.add_data_validation(dv); dv.add(f'B{FIRST}:B{LAST}')
dv = DataValidation(type='list', formula1='EX_NAMES', allow_blank=True, showErrorMessage=False); ws.add_data_validation(dv); dv.add(f'C{FIRST}:C{LAST}')
dv = DataValidation(type='list', formula1='"URL List,Category Match"', allow_blank=True); ws.add_data_validation(dv); dv.add(f'D{FIRST}:D{LAST}')
cf(ws, f'M{FIRST}:N{LAST}', f'$M{FIRST}="NG"', 'FFC7CE')
cf(ws, f'M{FIRST}:N{LAST}', f'$M{FIRST}="警告"', 'FFEB9C')
cf(ws, f'M{FIRST}:N{LAST}', f'$M{FIRST}="OK"', 'C6EFCE')
cf(ws, f'D{FIRST}:D{LAST}', f'AND($B{FIRST}="変更",$G{FIRST}="あり",NOT(EXACT(D{FIRST}&"",$I{FIRST})))', 'F8CBAD')
cf(ws, f'E{FIRST}:E{LAST}', f'AND($B{FIRST}="変更",$G{FIRST}="あり",NOT(EXACT(E{FIRST}&"",$K{FIRST})))', 'F8CBAD')

# =========================================================
# ④ URL記載（縦長：カテゴリ名 | FQDN、1行1件 = 反映後の全件）
# =========================================================
ws = sheet(SU, 3)
header(ws, '④ URL記載（WL/BL以外の全カスタムURLカテゴリ：カテゴリ名 | FQDN を1行1件。反映後の状態を記載）',
       'このシートが「反映後のあるべき姿」です。追加はどの行に追記してもOK（カテゴリ名はプルダウン選択）、削除は行ごと削除します。フィルタで絞り込んだまま行削除・追記して構いません（並び順は関係ありません）。'
       '「JSON発行」でカテゴリごとに既存と比べ、FQDNが変わったカテゴリだけをPUTします。シートに1行も無いカテゴリは変更しません（カテゴリの削除は③で行います）。'
       'C〜E列は「入力チェック」または「JSON発行」でVBAが書き込みます（NG行は反映対象外）。既存情報取得時にこのシートが空なら、既存の全件を書き出します。', 9, 60)
stat(ws, ('A3', 'B3'), '記載件数', f'=COUNTA(B{FIRST}:B{UL})')
stat(ws, ('A4', 'B4'), 'カテゴリ数', None, F_VBA)
stat(ws, ('C3', 'D3'), 'NG', f'=COUNTIF(E{FIRST}:E{UL},"NG")')
stat(ws, ('C4', 'D4'), '警告', f'=COUNTIF(E{FIRST}:E{UL},"警告")')
btn(ws, 'F3:F4', '▶ 入力チェック')
btn(ws, 'G3:G4', '▶ JSON発行 → JSON_CURLへ')
btn(ws, 'H3:H4', '▶ 既存から作り直す')
btn(ws, 'I3:I4', '▶ 入力クリア')
labels(ws, 6, {'A': '必須', 'B': '必須'})
put(ws, 'C6', '1カテゴリの上限は基本設定「1リストの上限件数」。カテゴリ名の候補に③の登録予定を足すには「入力チェック」を押してください', Font(name=FN, size=9), border=False)
heads(ws, 7, ['カテゴリ名', 'FQDN / URL（Category Matchはカテゴリ名）', '既存との照合', '判定', '理由'], [28, 50, 16, 8, 50])
for col in 'AB':
    colstyle(ws, col, F_IN, '@', Font(name=FN, size=10, color='0000FF'))
for col in 'CDE':
    colstyle(ws, col, F_VBA, '@', Font(name=FN, size=10))
ws.freeze_panes = 'A8'
ws.auto_filter.ref = f'A7:E{UL}'
dv = DataValidation(type='list', formula1='CAT_NAMES', allow_blank=True, showErrorMessage=True,
                    errorTitle='カテゴリ名', error='一覧に無いカテゴリ名です。新しいカテゴリは③で登録してから「入力チェック」を押すと候補に出ます。')
ws.add_data_validation(dv); dv.add(f'A{FIRST}:A{UL}')
cf(ws, f'D{FIRST}:E{UL}', f'$D{FIRST}="NG"', 'FFC7CE')
cf(ws, f'D{FIRST}:E{UL}', f'$D{FIRST}="警告"', 'FFEB9C')

# =========================================================
# ⑤ URL既存一覧（縦長：VBA書込み。差分の基準）
# =========================================================
ws = sheet(SX, 4)
header(ws, '⑤ URL既存一覧（SCMから取得したWL/BL以外の全エントリ：カテゴリ名 | FQDN、1行1件）',
       '「既存情報取得」でVBAが作り直します（直接編集しないでください）。④URL記載との差分比較の基準（実行前の状態）になります。'
       'E列は④のカテゴリ名プルダウンの候補（既存カテゴリ＋③で登録予定のもの）です。', 6, 40)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '件数', f'=COUNTA(B{FIRST}:B1048576)')
heads(ws, 7, ['カテゴリ名', 'エントリ', 'type'], [28, 50, 14])
heads(ws, 7, ['カテゴリ候補（④のプルダウン用）'], [30], start=5)
for col in 'ABCE':
    colstyle(ws, col, None, '@')
ws.freeze_panes = 'A8'
ws.auto_filter.ref = 'A7:C7'

# =========================================================
# ⑥ 横長ビュー（閲覧用：1列1カテゴリ、WL/BLは含まない）
# =========================================================
ws = sheet(SV, 5)
header(ws, '⑥ 横長ビュー（閲覧専用：1列1カテゴリ、WL/BLは含みません）',
       '「既存情報取得」のたびにVBAが作り直します（SCMの現在の内容）。編集は④URL記載シートで行ってください。', 12, 30)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
for r, t in [(5, 'name'), (6, 'type'), (7, '件数'), (8, 'description'), (9, 'エントリ →')]:
    put(ws, f'A{r}', t, Font(name=FN, size=10, bold=True, color='FFFFFF'), F_HEAD, Alignment(horizontal='center', vertical='center'))
ws.column_dimensions['A'].width = 12
ws.freeze_panes = 'B10'

# =========================================================
# ⑫ 差分確認（カテゴリ別サマリ + 追加・削除FQDNの詳細）
# =========================================================
ws = sheet(SD, len(wb.sheetnames) - 3)   # WL/BL差分確認・実行ログの前
header(ws, '⑫ 差分確認（WL/BL以外のカスタムURLカテゴリ）：上=カテゴリ別サマリ、下=追加・削除されるFQDNの詳細',
       '「JSON発行」で実行前のプレビューを作ります（既存＝⑤URL既存一覧、反映後＝④URL記載・③）。行の背景色：新規=緑、変更=黄、削除=グレー、変化なし=白。'
       'カテゴリ名をクリックすると下の詳細へ移動します。「API実行」後はVBAがSCMから再取得し、K列「確認」と詳細のE列「実行後」に結果を書き込みます（想定と違うものは赤）。'
       '「変化なしを隠す」で変化の無いカテゴリを非表示にできます。', 13, 56)
stat(ws, ('A3', 'B3'), '作成日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '状態', None, F_VBA)
for addr, lab in [(('C3', 'D3'), '新規'), (('E3', 'F3'), '変更'), (('G3', 'H3'), '削除'),
                  (('C4', 'D4'), '変化なし'), (('E4', 'F4'), '追加FQDN'), (('G4', 'H4'), '削除FQDN')]:
    stat(ws, addr, lab, None, F_VBA)
stat(ws, ('I4', 'J4'), '要確認', None, F_VBA)
btn(ws, 'K3:L3', '▶ 差分チェック（実行後の再確認）')
btn(ws, 'K4', '▶ 変化なしを隠す')
btn(ws, 'L4', '▶ 全て表示')
put(ws, 'C5', '凡例：', Font(name=FN, size=9), border=False)
for col, t, c in [('D', '新規', 'C6EFCE'), ('E', '変更', 'FFFF99'), ('F', '削除', 'D9D9D9'), ('G', '変化なし', 'FFFFFF'), ('H', '要確認', 'FFC7CE')]:
    put(ws, f'{col}5', t, Font(name=FN, size=9), fill(c), Alignment(horizontal='center'))
heads(ws, 7, ['No', 'カテゴリ名', '判定', '変更前件数', '変更後件数', '追加', '削除', 'type・descriptionの変更',
              'JSON_CURLの操作', 'HTTP', '確認', 'type（反映後）', 'description（反映後）'],
      [6, 30, 10, 10, 10, 8, 8, 30, 12, 7, 36, 13, 26])
ws.freeze_panes = 'C8'

# =========================================================
# 名前定義・シート順
# =========================================================
wb.defined_names['CAT_NAMES'] = DefinedName('CAT_NAMES', attr_text=f"OFFSET({Q(SX)}!$E${FIRST},0,0,MAX(1,COUNTA({Q(SX)}!$E${FIRST}:$E${EXL})),1)")
for n in ['UL_FETCHED_AT', 'DIFF_CREATED_AT']:
    if n in wb.defined_names:
        del wb.defined_names[n]
wb.defined_names['UL_FETCHED_AT'] = DefinedName('UL_FETCHED_AT', attr_text=f"{Q(SX)}!$B$3")

order = ['基本設定', SE, SC, SU, SX, SV, 'URLCat_WL_既存一覧', 'URLCat_WL_記載', 'URLCat_BL_既存一覧', 'URLCat_BL_記載',
         'JSON_CURL', SD, 'URLCat_WL_差分確認', 'URLCat_BL_差分確認', '実行ログ']
assert sorted(order) == sorted(wb.sheetnames), wb.sheetnames
wb._sheets = [wb[n] for n in order]
wb.active = 0
wb.save(OUT)
print('saved', OUT, wb.sheetnames)

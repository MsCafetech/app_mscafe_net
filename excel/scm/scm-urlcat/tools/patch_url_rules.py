# -*- coding: utf-8 -*-
# WL/BL記載シートの判定式に Palo Alto URLリストの書式ルールを追加し、注記を更新する（2026-10-29）
#   aa*.aaa.com / *.*.com = NG、末尾 / なし（前方一致）・途中の * ・パス指定 = 警告
#   使い方: python3 patch_url_rules.py 入力.xlsx 出力.xlsx
import sys
import openpyxl

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
RULE = '書式：*.aaa.com/ を推奨。aa*.aaa.com・*.*.com はNG。末尾 / なし（前方一致）・途中の*（aaa.*.com/）・パス指定（/以降）は警告'

HTTP = 'IF(OR(LEFT(LOWER($B{r}),7)="http://",LEFT(LOWER($B{r}),8)="https://"),"NG:http:// / https:// は不要です",'
OLDW = 'IF(ISERROR(FIND(".",$B{r})),"警告:ドット(.)を含みません","OK")'

def rule_parts(r):
    H = f'LEFT($B{r},FIND("/",$B{r}&"/")-1)'
    X = f'SUBSTITUTE(SUBSTITUTE("."&{H}&".",".*.",".#."),".*.",".#.")'
    NS = f'(LEN({X})-LEN(SUBSTITUTE({X},"#","")))'
    ng = (f'IF({H}="","NG:ホスト名がありません（先頭が / です）",'
          f'IF(ISNUMBER(FIND("*",{X})),"NG:ワイルドカード(*)はドットで区切った1区切り全体にしか使えません（aa*.aaa.com は登録不可。*.aaa.com/ の形にしてください）",'
          f'IF({NS}>1,"NG:ワイルドカード(*)は1か所だけです（*.*.com は登録不可）",')
    w = (f'IF(ISERROR(FIND(".",{H})),"ドット(.)を含みません／","")&'
         f'IF(AND({NS}=1,LEFT({H},2)<>"*."),"途中のワイルドカード（aaa.*.com/）は登録できますが推奨しません／","")&'
         f'IF(ISERROR(FIND("/",$B{r})),"末尾に / が無いため前方一致になり、"&$B{r}&".net なども一致します（"&$B{r}&"/ を推奨）／",'
         f'IF(LEN($B{r})>LEN({H})+1,"パス指定（/ 以降）はSSL復号化が前提です／",""))')
    warn = f'IF({w}="","OK","警告:"&LEFT({w},LEN({w})-1))'
    return ng, warn

for lk in ['WL', 'BL']:
    ws = wb[f'URLCat_{lk}_記載']
    for r in range(8, 2508):
        f = ws[f'F{r}'].value
        if not isinstance(f, str) or 'ワイルドカード' in f:
            continue
        h, o = HTTP.format(r=r), OLDW.format(r=r)
        assert h in f and o in f, (lk, r)
        ng, warn = rule_parts(r)
        f = f.replace(h, h + ng).replace(o, warn) + ')))'
        assert len(f) < 8000
        ws[f'F{r}'] = f
    ws['C6'] = '1行1件。2508行目以降は対象外です。' + RULE
wb['URLCat_URL記載']['C6'] = ('1カテゴリの上限は基本設定「1リストの上限件数」。カテゴリ名の候補に③の登録予定を足すには「入力チェック」を押してください。' + RULE)
wb.save(OUT)
print('saved', OUT)

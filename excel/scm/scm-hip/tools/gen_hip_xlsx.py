# -*- coding: utf-8 -*-
# SCM HIPオブジェクト / HIPプロファイル Excel 生成
#   python3 gen_hip_xlsx.py 出力.xlsx
#   SCM_ServiceObject.xlsx を読み込み、基本設定 / JSON_CURL / 実行ログ を流用する
import sys, os, datetime, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hipcore as H

SRC = '/mnt/project-files/scm-service/SCM_ServiceObject.xlsx'
OUT = sys.argv[1]

SE, SI, PE, PI = 'HipObj_既存一覧', 'HipObj_登録変更削除', 'HipProfile_既存一覧', 'HipProfile_登録変更削除'
SV = 'HIP_ベンダー一覧'
SC, SJ, SD, PD, SL = 'HIP_選択肢', 'JSON_CURL', 'HipObj_差分確認', 'HipProfile_差分確認', '実行ログ'
FIRST, LAST, EXL, PEXL = 8, 507, 5000, 2000
MODE_AND, MODE_OR, MODE_RAW = 'すべて満たす(and)', 'いずれか(or)', '直接記述'

FN = 'Meiryo UI'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', fgColor=c)
F_TITLE, F_HEAD, F_IN, F_FX, F_VBA, F_BTN = fill('1F3864'), fill('2F5597'), fill('FFF2CC'), fill('F2F2F2'), fill('E2EFDA'), fill('C55A11')
F_REQ, F_OPT = fill('C00000'), fill('7F7F7F')

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

def heads(ws, row, names, widths, start=1):
    for i, (n, w) in enumerate(zip(names, widths), start):
        put(ws, f'{CL(i)}{row}', n, Font(name=FN, size=10, bold=True, color='FFFFFF'), F_HEAD,
            Alignment(horizontal='center', vertical='center', wrap_text=True))
        ws.column_dimensions[CL(i)].width = w

def labels(ws, row, items):
    for col, t in items.items():
        put(ws, f'{col}{row}', t, Font(name=FN, size=9, bold=True, color='FFFFFF'),
            F_REQ if t.startswith('必須') else F_OPT, Alignment(horizontal='center', vertical='center'), False)

def stat(ws, addr, label, formula, fl=F_FX):
    put(ws, addr[0], label, Font(name=FN, size=10, bold=True))
    put(ws, addr[1], formula, Font(name=FN, size=10, bold=True), fl)

def cf(ws, rng, formula, color):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(color), stopIfTrue=False))

def dvlist(ws, rng, f1, err=True):
    dv = DataValidation(type='list', formula1=f1, allow_blank=True, showErrorMessage=err)
    dv.add(rng); ws.add_data_validation(dv)

Q = lambda s: "'" + s + "'"   # シート名の引用
def ex(sh, col, last=EXL, first=FIRST): return f"{Q(sh)}!${col}${first}:${col}${last}"

def nest(conds, last):
    f = last
    for cond, val in reversed(conds):
        f = f'IF({cond},{val},{f})'
    return f

NAMECHK = ('IF(SUMPRODUCT(--ISERROR(FIND(MID($C{r}&REPT(" ",31),ROW($1:$31),1),'
           '" 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")))>0,'
           '"NG:nameに使用不可の文字(英数字 空白 . _ - のみ)",')

# =========================================================
# 項目定義（② ③ の F列以降。並びは hipcore.FIELD_KEYS / modHip.FieldKeys と同じ）
# =========================================================
YN, YNA, STROP = '"yes,no"', '"yes,no,not-available"', '"is,is_not,contains"'
OSV, WITHIN, TOP = '"Microsoft,Apple,Google,Linux,Other"', '"within,not_within"', '"not_available,within,not_within"'
PVOP = '"greater_equal,greater_than,is,is_not,less_equal,less_than,contains,within,not_within"'
SEVOP = '"greater_equal,greater_than,is,is_not,less_equal,less_than"'
FIELDS = [  # key, 見出し, 幅, プルダウン, カテゴリ
    ('description', 'description\n(255文字以内)', 22, None, '基本'),
    ('hi_os_vendor', 'OS\nベンダー', 11, OSV, 'host'), ('hi_os_value', 'OS 値\n(空欄=All)', 16, 'OS_VALUES', 'host'),
    ('hi_managed', '管理対象\n(managed)', 9, YN, 'host'),
    ('hi_domain_op', 'ドメイン\n条件', 8, STROP, 'host'), ('hi_domain_val', 'ドメイン\n値', 16, None, 'host'),
    ('hi_hostname_op', 'ホスト名\n条件', 8, STROP, 'host'), ('hi_hostname_val', 'ホスト名\n値', 14, None, 'host'),
    ('hi_clientver_op', 'GPクライアント\nバージョン 条件', 10, STROP, 'host'), ('hi_clientver_val', 'GPクライアント\nバージョン 値', 12, None, 'host'),
    ('hi_serial_op', 'シリアル番号\n条件', 9, STROP, 'host'), ('hi_serial_val', 'シリアル番号\n値', 12, None, 'host'),
    ('hi_hostid_op', 'ホストID\n条件', 8, STROP, 'host'), ('hi_hostid_val', 'ホストID\n値', 12, None, 'host'),
    ('am_installed', 'インストール\n(is_installed)', 10, YN, 'am'), ('am_rtp', 'リアルタイム\n保護', 11, YNA, 'am'),
    ('am_virdef_op', '定義ファイル\n条件', 10, WITHIN, 'am'), ('am_virdef_unit', '定義ファイル\n単位', 9, '"days,versions"', 'am'),
    ('am_virdef_n', '定義ファイル\n数(1-65535)', 9, None, 'am'),
    ('am_scan_op', '最終スキャン\n条件', 12, TOP, 'am'), ('am_scan_unit', '最終スキャン\n単位', 8, '"days,hours"', 'am'),
    ('am_scan_n', '最終スキャン\n数(1-65535)', 9, None, 'am'),
    ('am_pv_op', '製品バージョン\n条件', 12, PVOP, 'am'), ('am_pv_val', '製品バージョン\n値', 12, None, 'am'),
    ('am_vendor', 'ベンダー | 製品\n(ダブルクリックで選択)', 30, None, 'am'), ('am_excl', 'ベンダー\n除外', 8, YN, 'am'),
    ('db_installed', 'インストール\n(is_installed)', 10, YN, 'db'),
    ('db_time_op', '最終バックアップ\n条件', 12, TOP, 'db'), ('db_time_unit', '最終バックアップ\n単位', 9, '"days,hours"', 'db'),
    ('db_time_n', '最終バックアップ\n数(1-65535)', 9, None, 'db'),
    ('db_vendor', 'ベンダー | 製品\n(ダブルクリックで選択)', 26, None, 'db'), ('db_excl', 'ベンダー\n除外', 8, YN, 'db'),
    ('de_installed', 'インストール\n(is_installed)', 10, YN, 'de'),
    ('de_locations', '暗号化場所 | 状態\n(1行1件、!=is_not)', 24, None, 'de'),
    ('de_vendor', 'ベンダー | 製品\n(ダブルクリックで選択)', 26, None, 'de'), ('de_excl', 'ベンダー\n除外', 8, YN, 'de'),
    ('fw_installed', 'インストール\n(is_installed)', 10, YN, 'fw'), ('fw_enabled', '有効\n(is_enabled)', 10, YNA, 'fw'),
    ('fw_vendor', 'ベンダー | 製品\n(ダブルクリックで選択)', 26, None, 'fw'), ('fw_excl', 'ベンダー\n除外', 8, YN, 'fw'),
    ('pm_installed', 'インストール\n(is_installed)', 10, YN, 'pm'), ('pm_enabled', '有効\n(is_enabled)', 10, YNA, 'pm'),
    ('pm_check', '不足パッチ\n判定(check)', 11, '"has-any,has-none,has-all"', 'pm'),
    ('pm_sev_op', '重要度\n条件', 12, SEVOP, 'pm'), ('pm_sev_val', '重要度\n値(0-100000)', 10, None, 'pm'),
    ('pm_patches', 'パッチ\n(1行1件)', 18, None, 'pm'),
    ('pm_vendor', 'ベンダー | 製品\n(ダブルクリックで選択)', 26, None, 'pm'), ('pm_excl', 'ベンダー\n除外', 8, YN, 'pm'),
    ('dlp_installed', 'インストール\n(is_installed)', 10, YN, 'dlp'), ('dlp_enabled', '有効\n(is_enabled)', 10, YNA, 'dlp'),
    ('dlp_vendor', 'ベンダー | 製品\n(ダブルクリックで選択)', 26, None, 'dlp'), ('dlp_excl', 'ベンダー\n除外', 8, YN, 'dlp'),
    ('ce_profile', '証明書\nプロファイル', 18, 'CERT_NAMES', 'ce'), ('ce_attrs', '証明書属性 | 値\n(1行1件)', 22, None, 'ce'),
    ('nw_op', 'ネットワーク\n条件', 9, '"is,is_not"', 'nw'), ('nw_type', 'ネットワーク\n種別', 10, '"wifi,mobile,ethernet,unknown"', 'nw'),
    ('nw_val', 'SSID／キャリア\n(wifi・mobile)', 16, None, 'nw'),
    ('cc_process', 'プロセス\n(1行1件、!=停止中)', 22, None, 'cc'),
    ('cc_registry', 'レジストリ\n(1行1件)', 32, None, 'cc'), ('cc_plist', 'plist\n(1行1件)', 28, None, 'cc'),
]
assert [f[0] for f in FIELDS] == H.FIELD_KEYS
F0 = 6                                   # F列
COL = {f[0]: CL(F0 + i) for i, f in enumerate(FIELDS)}
FL = CL(F0 + len(FIELDS) - 1)            # BM
assert FL == 'BM'
CATS = [('基本', '基本', 'D9E1F2'), ('host', 'ホスト情報 host_info', 'DDEBF7'), ('am', 'マルウェア対策 anti_malware', 'FCE4D6'),
        ('db', 'ディスクバックアップ disk_backup', 'E2EFDA'), ('de', 'ディスク暗号化 disk_encryption', 'FFF2CC'),
        ('fw', 'ファイアウォール firewall', 'EDE2F6'), ('pm', 'パッチ管理 patch_management', 'DDEBF7'),
        ('dlp', 'データ損失防止 data_loss_prevention', 'FCE4D6'), ('ce', '証明書 certificate', 'E2EFDA'),
        ('nw', 'ネットワーク network_info', 'FFF2CC'), ('cc', 'カスタムチェック custom_checks', 'EDE2F6')]

def bands(ws, row):
    for key, label, color in CATS:
        cols = [F0 + i for i, f in enumerate(FIELDS) if f[4] == key]
        a, b = CL(cols[0]), CL(cols[-1])
        if a != b: ws.merge_cells(f'{a}{row}:{b}{row}')
        put(ws, f'{a}{row}', label, Font(name=FN, size=9, bold=True), fill(color), Alignment(horizontal='center', vertical='center'))
        for c in cols: ws[f'{CL(c)}{row}'].fill = fill(color); ws[f'{CL(c)}{row}'].border = BOX
    ws.row_dimensions[row].height = 18

# =========================================================
# 基本設定 / JSON_CURL / 実行ログ（流用して文言を変更）
# =========================================================
cfg = wb['基本設定']
cfg['B7'] = '/sse/config/v1/hip-objects'
cfg['D7'] = 'HIPオブジェクト（HIPプロファイルは /sse/config/v1/hip-profiles、証明書プロファイル名の取得は /sse/config/v1/certificate-profiles をVBA内で使用）'
cfg['D16'] = '必須。運用ルールにより Shared（Prisma Access）に登録します'
cfg['B19'] = None; cfg['B20'] = None; cfg['B21'] = None

js = wb[SJ]
js['A2'] = ('③⑦のJSON発行ボタンでVBAが10行目以降を生成します（緑=VBA書込み）。curlはbash形式。先に5行目でトークン取得→6行目でTOKEN変数を設定してから各行を実行。'
            '「API実行」ボタンでExcelから直接実行も可能（結果は⑤実行ログへ）。JSON発行は③と⑦のどちらのボタンでも両方から一覧を作り直し、'
            'HIPオブジェクト登録・変更→HIPプロファイル登録・変更→HIPプロファイル削除→HIPオブジェクト削除の順に並べます（プロファイルが参照するオブジェクトを先に作り、参照が外れてから削除するため）。')
for row in js.iter_rows(min_row=10, max_row=js.max_row):
    for c in row: c.value = None

lg = wb[SL]
for row in lg.iter_rows(min_row=6, max_row=lg.max_row):
    for c in row:
        c.value = None; c.border = Border(); c.number_format = 'General'

# =========================================================
# サンプル（既存一覧・入力例）
# =========================================================
ex_objs = [
    {'name': 'hip-win-managed', 'description': '社用Windows（ドメイン参加）',
     'host_info': {'criteria': {'os': {'contains': {'Microsoft': 'All'}}, 'domain': {'is': 'corp.example.local'}, 'managed': True}}},
    {'name': 'hip-av-defender', 'description': 'Defender 有効・定義3日以内',
     'anti_malware': {'criteria': {'is_installed': True, 'real_time_protection': 'yes', 'virdef_version': {'within': {'days': 3}}},
                      'vendor': [{'name': 'Microsoft Corporation', 'product': ['Microsoft Defender Antivirus']}]}},
    {'name': 'hip-disk-enc', 'description': 'Cドライブ暗号化',
     'disk_encryption': {'criteria': {'is_installed': True, 'encrypted_locations': [{'name': 'C:\\', 'encryption_state': {'is': 'encrypted'}}]}}},
    {'name': 'hip-fw-on', 'description': 'OSファイアウォール有効',
     'firewall': {'criteria': {'is_installed': True, 'is_enabled': 'yes'}}},
    {'name': 'hip-patch', 'description': '重要度3以上の未適用パッチなし',
     'patch_management': {'criteria': {'is_installed': True, 'is_enabled': 'yes',
                                       'missing_patches': {'check': 'has-none', 'severity': {'greater_equal': 3}}}}},
    {'name': 'hip-mac', 'description': '社用Mac', 'host_info': {'criteria': {'os': {'contains': {'Apple': 'All'}}}}},
    {'name': 'hip-old-check', 'description': '旧エージェント（廃止予定）',
     'custom_checks': {'criteria': {'process_list': [{'name': 'legacy-agent.exe', 'running': True}],
                                    'registry_key': [{'name': 'HKLM\\SOFTWARE\\Legacy\\Agent'}]}}},
    {'name': 'hip-android', 'description': 'Android（脱獄なし）',
     'host_info': {'criteria': {'os': {'contains': {'Google': 'All'}}}},
     'mobile_device': {'criteria': {'jailbroken': False, 'passcode_set': True}}},
]
ex_profs = [
    ('prof-corp-win', '社用Windowsの基準', '"hip-win-managed" and "hip-av-defender" and "hip-disk-enc"'),
    ('prof-mac', '社用Mac', '"hip-mac"'),
    ('prof-legacy', '旧エージェント端末', '"hip-old-check"'),
    ('prof-complex', '複合条件の例', '"hip-win-managed" and ("hip-av-defender" or "hip-fw-on")'),
]
# ③ 入力例（項目キー → 値）
def R(**kw): return kw
obj_in = [
    ('登録', 'hip-win11', R(description='Windows 11 社用端末', hi_os_vendor='Microsoft', hi_os_value='Microsoft Windows 11', hi_managed='yes')),
    ('登録', 'hip-crowdstrike', R(description='CrowdStrike 稼働中', am_installed='yes', am_rtp='yes', am_scan_op='within', am_scan_unit='days', am_scan_n='7',
                                  am_vendor='CrowdStrike, Inc. | CrowdStrike Falcon')),
    ('変更', 'hip-av-defender', R(description='Defender 有効・定義1日以内', am_installed='yes', am_rtp='yes', am_virdef_op='within', am_virdef_unit='days',
                                  am_virdef_n='1', am_vendor='Microsoft Corporation | Microsoft Defender Antivirus')),
    ('変更', 'hip-fw-on', R(description='OSファイアウォール有効', fw_installed='yes', fw_enabled='yes')),
    ('削除', 'hip-old-check', R()),
    ('登録', 'hip-ng-days', R(description='日数が範囲外', am_installed='yes', am_virdef_op='within', am_virdef_unit='days', am_virdef_n='0')),
    ('登録', 'hip-ng-empty', R(description='条件なし')),
    ('削除', 'hip-mac', R()),
]
prof_in = [
    ('登録', 'prof-win11', 'Windows 11 基準', MODE_AND, 'hip-win11,hip-crowdstrike', '', ''),
    ('変更', 'prof-corp-win', '社用Windowsの基準（FW追加）', MODE_AND, 'hip-win-managed,hip-av-defender,hip-disk-enc,hip-fw-on', '', ''),
    ('削除', 'prof-legacy', '', '', '', '', ''),
    ('登録', 'prof-not-legacy', 'WindowsかMac（旧エージェント除く）', MODE_OR, 'hip-win-managed,hip-mac', 'hip-fw-on', ''),
    ('登録', 'prof-ng-unknown', '存在しない名前', MODE_AND, 'hip-unknown', '', ''),
]

# =========================================================
# ② HipObj_既存一覧
# =========================================================
NF = len(FIELDS)
ws = sheet(SE, 1)
header(ws, '② 既存情報（SCMから取得した HIPオブジェクト 一覧）',
       'ボタンを押すとAPI（GET）でHIPオブジェクトとHIPプロファイルの両方を取得し、8行目以降に表を再生成します（B〜BN列=VBA書込み）。直接編集はしないでください。'
       'F列以降の並びは③と同じです。E列はこのツールでは編集しない条件（モバイルデバイス mobile_device）で、③で変更してもそのまま引き継ぎます。BN列は取得したJSON（変更時の引き継ぎと差分確認に使用）。', 14)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{EXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
bands(ws, 6)
heads(ws, 7, ['No', 'id', 'name', 'コンテナ', '対象外の条件\n(mobile_device)'] + [f[1] for f in FIELDS] + ['JSON（取得結果）'],
      [6, 30, 18, 14, 22] + [f[2] for f in FIELDS] + [40])
ws.row_dimensions[7].height = 44
for i, it in enumerate(ex_objs):
    r = FIRST + i
    row = H.obj_to_row(it)
    al = Alignment(wrap_text=True, vertical='top')
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', al=al)
    put(ws, f'B{r}', f'7a3c{i+1:04d}-1b2c-4d5e-8f90-0000000000{i+1:02d}', nf='@', al=al)
    put(ws, f'C{r}', it['name'], nf='@', al=al)
    put(ws, f'D{r}', 'folder=Shared', nf='@', al=al)
    put(ws, f'E{r}', H.flat_text(it.get('mobile_device')) or None, nf='@', al=al)
    for k, c in COL.items():
        put(ws, f'{c}{r}', row.get(k) or None, nf='@', al=al)
    o = dict(it); o['id'] = f'7a3c{i+1:04d}-1b2c-4d5e-8f90-0000000000{i+1:02d}'; o['folder'] = 'Shared'
    put(ws, f'BN{r}', json.dumps(o, ensure_ascii=False, separators=(',', ':')), nf='@', al=Alignment(vertical='top'))
ws.freeze_panes = 'D8'

# =========================================================
# ③ HipObj_登録変更削除
#   B=操作 C=name D=判定 E=理由 F〜BM=項目 BN〜=関数
# =========================================================
ws = sheet(SI, 2)
W = ['BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY']
header(ws, '③ 今回の登録・変更・削除一覧（HIPオブジェクト）',
       '黄色(B・C列とF〜BM列)に入力。nameは既存一覧からプルダウン選択可（変更・削除時）。多くの欄はプルダウンで選びます。D・E列とBN列以降は関数で②と自動照合し、判定がNGの行はJSON発行対象外。'
       '変更時に既存と異なる項目は橙色。変更(PUT)は全項目置換のため、先に「変更行に既存値を読込」で現在の値を入れてから直してください。'
       'ベンダー欄はダブルクリック（またはCtrl+Shift+M）で「HIP_ベンダー一覧」から選べます。使わないカテゴリは空欄のまま（1つでも入力したカテゴリだけJSONに入ります）。1行1件の欄はセル内改行（Alt+Enter）で複数行。書式は「HIP_選択肢」シートの書式ガイド参照。'
       'モバイルデバイスの条件はこのツールでは編集せず、変更時は既存の内容を引き継ぎます。', 20, 70)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(D{FIRST}:D{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(D{FIRST}:D{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(D{FIRST}:D{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行（③⑦）→ ④へ')
btn(ws, 'I3:J4', '▶ 変更行に既存値を読込')
btn(ws, 'L3:M4', '▶ 入力クリア')
labels(ws, 5, {'B': '必須', 'C': '必須'})
put(ws, 'D5', '← 判定・理由は関数', Font(name=FN, size=9, color='595959'), border=False)
bands(ws, 6)
heads(ws, 7, ['No', '操作', 'name\n(31文字以内)', '判定', '理由'] + [f[1] for f in FIELDS] +
      ['既存有無', '既存id', '作業:既存行', '一覧内重複', 'チェック:\nホスト情報', 'チェック:\nマルウェア対策', 'チェック:\nバックアップ',
       'チェック:\nパッチ管理', 'チェック:\nネットワーク', '作業:参照する\nプロファイル数', '作業:入力\n項目数', '作業:既存の\n対象外条件'],
      [5, 8, 18, 7, 40] + [f[2] for f in FIELDS] + [8, 22, 8, 9, 26, 26, 26, 26, 26, 12, 8, 14])
ws.row_dimensions[7].height = 44
C = COL
def pair(op, val, label):
    return [(f'AND(${op}{{r}}<>"",${val}{{r}}="")', f'"NG:{label}の値が未入力"'),
            (f'AND(${op}{{r}}="",${val}{{r}}<>"")', f'"NG:{label}の条件が未選択"'),
            (f'LEN(${val}{{r}})>255', f'"NG:{label}の値が255文字超"')]
def numck(op_cond, n, mn, mx, label):
    m = f'"NG:{label}は{mn}〜{mx}の整数"'
    return [(f'IF({op_cond},ISERROR(--${n}{{r}}),FALSE)', m),
            (f'IF({op_cond},OR(--${n}{{r}}<{mn},--${n}{{r}}>{mx},--${n}{{r}}<>INT(--${n}{{r}})),FALSE)', m)]
def timeck(k, label, na):
    op, u, n = C[k + '_op'], C[k + '_unit'], C[k + '_n']
    c = [(f'AND(${op}{{r}}="",OR(${u}{{r}}<>"",${n}{{r}}<>""))', f'"NG:{label}の条件が未選択"')]
    if na:
        c.append((f'AND(${op}{{r}}="not_available",OR(${u}{{r}}<>"",${n}{{r}}<>""))', f'"NG:{label}が not_available のときは単位・数は空欄"'))
    live = f'AND(${op}{{r}}<>"",${op}{{r}}<>"not_available")'
    c.append((f'AND({live},OR(${u}{{r}}="",${n}{{r}}=""))', f'"NG:{label}の単位・数が未入力"'))
    return c + numck(live, n, 1, 65535, f'{label}の数')
CHK = {
    'BR': [(f'AND(${C["hi_os_vendor"]}{{r}}="",${C["hi_os_value"]}{{r}}<>"")', '"NG:OSの値を入れるときはベンダーも選択"'),
           (f'AND(${C["hi_os_vendor"]}{{r}}="Other",${C["hi_os_value"]}{{r}}="")', '"NG:OSのベンダーがOtherのときは値を入力"'),
           (f'LEN(${C["hi_os_value"]}{{r}})>255', '"NG:OSの値が255文字超"')]
          + pair(C['hi_domain_op'], C['hi_domain_val'], 'ドメイン') + pair(C['hi_hostname_op'], C['hi_hostname_val'], 'ホスト名')
          + pair(C['hi_clientver_op'], C['hi_clientver_val'], 'GPクライアントバージョン')
          + pair(C['hi_serial_op'], C['hi_serial_val'], 'シリアル番号') + pair(C['hi_hostid_op'], C['hi_hostid_val'], 'ホストID'),
    'BS': timeck('am_virdef', '定義ファイル', False) + timeck('am_scan', '最終スキャン', True) + [
           (f'AND(${C["am_pv_op"]}{{r}}="",${C["am_pv_val"]}{{r}}<>"")', '"NG:製品バージョンの条件が未選択"'),
           (f'AND(${C["am_pv_op"]}{{r}}<>"",${C["am_pv_val"]}{{r}}="")', '"NG:製品バージョンの値が未入力"')]
          + numck(f'OR(${C["am_pv_op"]}{{r}}="within",${C["am_pv_op"]}{{r}}="not_within")', C['am_pv_val'], 1, 65535,
                  '製品バージョン(within/not_within)の値(バージョン数)')
          + [(f'LEN(${C["am_pv_val"]}{{r}})>255', '"NG:製品バージョンの値が255文字超"')],
    'BT': timeck('db_time', '最終バックアップ', True),
    'BU': [(f'AND(${C["pm_sev_op"]}{{r}}="",${C["pm_sev_val"]}{{r}}<>"")', '"NG:重要度の条件が未選択"'),
           (f'AND(${C["pm_sev_op"]}{{r}}<>"",${C["pm_sev_val"]}{{r}}="")', '"NG:重要度の値が未入力"')]
          + numck(f'${C["pm_sev_op"]}{{r}}<>""', C['pm_sev_val'], 0, 100000, '重要度の値')
          + [(f'AND(${C["pm_check"]}{{r}}="",OR(${C["pm_sev_op"]}{{r}}<>"",${C["pm_patches"]}{{r}}<>""))',
              '"NG:重要度・パッチを指定するときは不足パッチの判定(check)も選択"')],
    'BV': [(f'AND(${C["nw_op"]}{{r}}="",OR(${C["nw_type"]}{{r}}<>"",${C["nw_val"]}{{r}}<>""))', '"NG:ネットワークの条件(is / is_not)が未選択"'),
           (f'AND(${C["nw_op"]}{{r}}<>"",${C["nw_type"]}{{r}}="")', '"NG:ネットワークの種別が未選択"'),
           (f'AND(${C["nw_op"]}{{r}}="is",${C["nw_type"]}{{r}}="ethernet")', '"NG:ethernet は is_not のときだけ指定できます"'),
           (f'AND(OR(${C["nw_type"]}{{r}}="ethernet",${C["nw_type"]}{{r}}="unknown"),${C["nw_val"]}{{r}}<>"")', '"NG:ethernet・unknown のときはSSID／キャリアは空欄"'),
           (f'LEN(${C["nw_val"]}{{r}})>1023', '"NG:SSID／キャリアが1023文字超"')],
}
G0 = CL(F0 + 1)   # G列（descriptionの次＝条件の先頭）
EXF = f"{Q(SE)}!$F${FIRST}:${FL}${EXL}"
for r in range(FIRST, LAST + 1):
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    for col in ['B', 'C'] + list(COL.values()):
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@',
            al=Alignment(wrap_text=True, vertical='top'))
    put(ws, f'BN{r}', f'=IF($C{r}="","",IF(COUNTIF({ex(SE,"C")},$C{r})>0,"あり","なし"))', fl=F_FX)
    put(ws, f'BO{r}', f'=IF($BN{r}="あり",INDEX({ex(SE,"B")},MATCH($C{r},{ex(SE,"C")},0))&"","")', fl=F_FX)
    put(ws, f'BP{r}', f'=IF($BN{r}="あり",MATCH($C{r},{ex(SE,"C")},0),"")', fl=F_FX)
    put(ws, f'BQ{r}', f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,"重複",""))', fl=F_FX)
    for col, conds in CHK.items():
        cc = [(a.format(r=r), b) for a, b in conds]
        put(ws, f'{col}{r}', f'=IF(OR($C{r}="",$B{r}="削除"),"",{nest(cc, chr(34)*2)})', fl=F_FX)
    q = f'""""&$C{r}&""""'
    put(ws, f'BW{r}', f'=IF($C{r}="","",SUMPRODUCT(--ISNUMBER(FIND({q},{ex(PE,"E",PEXL)})),--({ex(PE,"H",PEXL)}=0))'
                      f'+SUMPRODUCT(--({ex(PI,"B",LAST)}<>"削除"),--({ex(PI,"B",LAST)}<>""),--ISNUMBER(FIND({q},{ex(PI,"I",LAST)}))))', fl=F_FX)
    put(ws, f'BX{r}', f'=SUMPRODUCT(--(${G0}{r}:${FL}{r}<>""))', fl=F_FX)
    put(ws, f'BY{r}', f'=IF($BN{r}="あり",INDEX({ex(SE,"E")},$BP{r})&"","")', fl=F_FX)
    put(ws, f'D{r}', f'=IF($E{r}="","",IF($E{r}="OK","OK",LEFT($E{r},FIND(":",$E{r})-1)))', fl=F_FX)
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$BQ{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>31', '"NG:nameが31文字超"'),
        (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
    ]
    rest = [
        (f'AND($B{r}="登録",$BN{r}="あり")', '"NG:既存に同名あり(登録不可)"'),
        (f'AND($B{r}<>"登録",$BN{r}="なし")', '"NG:既存エントリ無"'),
        (f'AND($B{r}<>"削除",LEN($F{r})>255)', '"NG:descriptionが255文字超"'),
        (f'AND($B{r}<>"削除",$BX{r}=0,$BY{r}="")', '"NG:条件が1つも入力されていません"'),
        (f'LEFT($BR{r},2)="NG"', f'$BR{r}'), (f'LEFT($BS{r},2)="NG"', f'$BS{r}'), (f'LEFT($BT{r},2)="NG"', f'$BT{r}'),
        (f'LEFT($BU{r},2)="NG"', f'$BU{r}'), (f'LEFT($BV{r},2)="NG"', f'$BV{r}'),
        (f'AND($B{r}="削除",$BW{r}>0)', f'"NG:HIPプロファイル("&$BW{r}&"件)のmatchで使われています(⑦でmatchから外すかプロファイルを削除してください)"'),
        (f'AND($B{r}="登録",COUNTIF({ex(PE,"C",PEXL)},$C{r})>0)', '"警告:同名のHIPプロファイルが既存にあります"'),
        (f'IF($B{r}="変更",SUMPRODUCT(--($F{r}:${FL}{r}&""<>INDEX({EXF},$BP{r},0)&""))=0,FALSE)', '"警告:既存と差分なし"'),
        (f'AND($B{r}="変更",$BY{r}<>"")', '"警告:モバイルデバイスの条件(このツールでは編集しない)は既存の内容を引き継ぎます"'),
    ]
    reason = nest(conds, NAMECHK.format(r=r) + nest(rest, '"OK"') + ')')
    put(ws, f'E{r}', f'=IF($C{r}="","",{reason})', fl=F_FX, al=Alignment(wrap_text=True, vertical='top'))
ws.freeze_panes = 'F8'
dvlist(ws, f'B{FIRST}:B{LAST}', '"登録,変更,削除"')
dvlist(ws, f'C{FIRST}:C{LAST}', 'HIPOBJ_NAMES', err=False)
for k, h, w, dv, cat in FIELDS:
    if dv: dvlist(ws, f'{COL[k]}{FIRST}:{COL[k]}{LAST}', dv, err=dv not in ('OS_VALUES', 'CERT_NAMES'))
cf(ws, f'D{FIRST}:E{LAST}', f'$D{FIRST}="NG"', 'FFC7CE')
cf(ws, f'D{FIRST}:E{LAST}', f'$D{FIRST}="警告"', 'FFEB9C')
cf(ws, f'D{FIRST}:E{LAST}', f'$D{FIRST}="OK"', 'C6EFCE')
cf(ws, f'F{FIRST}:{FL}{LAST}', f'AND($B{FIRST}="変更",$BN{FIRST}="あり",F{FIRST}&""<>INDEX({Q(SE)}!F${FIRST}:F${EXL},$BP{FIRST})&"")', 'F8CBAD')
for i, (op, nm, vals) in enumerate(obj_in):
    r = FIRST + i
    ws[f'B{r}'] = op; ws[f'C{r}'] = nm
    for k, v in vals.items(): ws[f'{COL[k]}{r}'] = v

# =========================================================
# ⑥ HipProfile_既存一覧
# =========================================================
ws = sheet(PE, 3)
header(ws, '⑥ 既存情報（SCMから取得した HIPプロファイル 一覧）',
       'ボタンを押すとAPI（GET）でHIPオブジェクトとHIPプロファイルの両方を取得し、8行目以降に表を再生成します（B〜G列=VBA書込み、H列=関数）。'
       'matchはHIPオブジェクト名を "名前" で囲み and / or / not で組み合わせた式です。F列はmatchで使われているHIPオブジェクト名。', 8)
stat(ws, ('A3', 'B3'), '取得日時', None, F_VBA); ws['B3'].number_format = 'yyyy/mm/dd hh:mm:ss'
stat(ws, ('A4', 'B4'), '取得件数', f'=COUNTA(C{FIRST}:C{PEXL})')
stat(ws, ('A5', 'B5'), '取得元', '=CFG_CONT_TYPE&"="&CFG_CONT_NAME')
btn(ws, 'D3:E4', '▶ 既存情報取得（API）')
heads(ws, 7, ['No', 'id', 'name', 'description', 'match', '参照HIPオブジェクト', 'コンテナ', '作業:⑦にあり'], [6, 30, 20, 24, 50, 36, 14, 9])
for r in range(FIRST, PEXL + 1):
    ws[f'H{r}'] = f'=IF(C{r}="","",IF(COUNTIF({ex(PI,"C",LAST)},C{r})>0,1,0))'
    ws[f'H{r}'].font = Font(name=FN, size=9, color='808080')
for i, (nm, desc, m) in enumerate(ex_profs):
    r = FIRST + i
    al = Alignment(wrap_text=True, vertical='top')
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', al=al)
    put(ws, f'B{r}', f'8b4d{i+1:04d}-2c3d-4e5f-9a01-0000000000{i+1:02d}', nf='@', al=al)
    for col, v in zip('CDEF', [nm, desc, m, ','.join(dict.fromkeys(m.split('"')[1::2]))]):
        put(ws, f'{col}{r}', v, nf='@', al=al)
    put(ws, f'G{r}', 'folder=Shared', nf='@', al=al)
ws.freeze_panes = 'D8'

# =========================================================
# ⑦ HipProfile_登録変更削除
# =========================================================
ws = sheet(PI, 4)
header(ws, '⑦ 今回の登録・変更・削除一覧（HIPプロファイル）',
       '黄色(B〜H列)に入力。組み方を選び、F列（対象）・G列（除外＝not）はセルをダブルクリックするとHIPオブジェクトのチェックリストが開きます（Ctrl+Shift+Mでも可）。'
       'I列に関数でmatch式を作ります（and：「"A" and "B"」、or：「"A" or "B"」、除外は「and not "C"」）。括弧を使う複雑な式は組み方「直接記述」にしてH列に式を書いてください。'
       'J〜AA列は関数で②⑥と照合し、HIPオブジェクトが②の既存一覧か③の登録予定にあるか確認します（大文字・小文字も区別）。判定がNGの行はJSON発行対象外。変更(PUT)は全項目置換です。', 17, 70)
stat(ws, ('A3', 'B3'), '対象件数', f'=COUNTA(C{FIRST}:C{LAST})')
stat(ws, ('C3', 'D3'), 'OK', f'=COUNTIF(P{FIRST}:P{LAST},"OK")')
stat(ws, ('A4', 'B4'), '警告', f'=COUNTIF(P{FIRST}:P{LAST},"警告")')
stat(ws, ('C4', 'D4'), 'NG', f'=COUNTIF(P{FIRST}:P{LAST},"NG")')
btn(ws, 'F3:G4', '▶ JSON発行（③⑦）→ ④へ')
btn(ws, 'I3:J4', '▶ 変更行に既存値を読込')
btn(ws, 'L3:M4', '▶ 入力クリア')
labels(ws, 6, {'B': '必須', 'C': '必須', 'D': '任意', 'E': '必須※', 'F': '必須※', 'G': '任意', 'H': '必須※'})
put(ws, 'I6', '※ 登録/変更時。F列かG列のどちらか1つ以上（直接記述のときはH列）。F・G列はダブルクリックで選択',
    Font(name=FN, size=9, color='595959'), border=False)
heads(ws, 7, ['No', '操作', 'name\n(31文字以内)', 'description\n(255文字以内)', '組み方', '対象HIPオブジェクト\n(ダブルクリック)',
              '除外HIPオブジェクト(not)\n(ダブルクリック)', 'match 直接記述\n(組み方=直接記述)', '生成される match\n(関数)',
              '既存有無', '既存id', '既存description', '既存match', '一覧内重複', 'メンバー確認', '判定', '理由',
              '作業:対象(正規化)', '作業:除外(正規化)', '作業:肯定部', '作業:否定部', '作業:全メンバー', '作業:メンバー数',
              '作業:一致数(延べ)', '作業:一致数(種類)', '作業:登録予定数', '作業:削除予定数'],
      [5, 8, 18, 20, 15, 30, 24, 30, 40, 8, 22, 18, 30, 9, 14, 7, 40, 20, 20, 20, 20, 20, 8, 8, 8, 8, 8])
ws.row_dimensions[7].height = 44
for r in range(FIRST, LAST + 1):
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-7)', fl=F_FX)
    for col in 'BCDEFGH':
        put(ws, f'{col}{r}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@', al=Alignment(wrap_text=True, vertical='top'))
    put(ws, f'J{r}', f'=IF($C{r}="","",IF(COUNTIF({ex(PE,"C",PEXL)},$C{r})>0,"あり","なし"))', fl=F_FX)
    for col, src in {'K': 'B', 'L': 'D', 'M': 'E'}.items():
        put(ws, f'{col}{r}', f'=IF($J{r}="あり",INDEX({ex(PE,src,PEXL)},MATCH($C{r},{ex(PE,"C",PEXL)},0))&"","")', fl=F_FX,
            al=Alignment(wrap_text=True, vertical='top'))
    put(ws, f'N{r}', f'=IF($C{r}="","",IF(COUNTIF($C${FIRST}:$C${LAST},$C{r})>1,"重複",""))', fl=F_FX)
    for col, src in (('R', 'F'), ('S', 'G')):
        put(ws, f'{col}{r}', f'=IF(${src}{r}="","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(${src}{r},CHAR(13),""),CHAR(10),","),", ",",")," ,",","))', fl=F_FX)
    isor = f'$E{r}="{MODE_OR}"'
    cntR = f'(LEN($R{r})-LEN(SUBSTITUTE($R{r},",",""))+1)'
    paren = f'AND({cntR}>1,{isor},$S{r}<>"")'
    put(ws, f'T{r}', f'=IF($R{r}="","",IF({paren},"(","")&""""&SUBSTITUTE($R{r},",",""" "&IF({isor},"or","and")&" """)&""""&IF({paren},")",""))', fl=F_FX)
    put(ws, f'U{r}', f'=IF($S{r}="","","not """&SUBSTITUTE($S{r},",",""" and not """)&"""")', fl=F_FX)
    put(ws, f'I{r}', f'=IF(OR($C{r}="",$B{r}="削除"),"",IF($E{r}="{MODE_RAW}",$H{r},$T{r}&IF(AND($T{r}<>"",$U{r}<>"")," and ","")&$U{r}))', fl=F_FX,
        al=Alignment(wrap_text=True, vertical='top'))
    put(ws, f'V{r}', f'=IF(AND($R{r}="",$S{r}=""),"",IF($R{r}="",$S{r},IF($S{r}="",$R{r},$R{r}&","&$S{r})))', fl=F_FX)
    put(ws, f'W{r}', f'=IF($V{r}="",0,LEN($V{r})-LEN(SUBSTITUTE($V{r},",",""))+1)', fl=F_FX)
    M2 = f'","&SUBSTITUTE($V{r},",",",,")&","'
    FN_ = f'","&$V{r}&","'
    def occ(rng, conds=''):
        return f'SUMPRODUCT(--({rng}<>""){conds},(LEN({M2})-LEN(SUBSTITUTE({M2},","&{rng}&",","")))/(LEN({rng})+2))'
    def dis(rng, conds=''):
        return f'SUMPRODUCT(--({rng}<>""){conds},--ISNUMBER(FIND(","&{rng}&",",{FN_})))'
    c3 = f',--({ex(SI,"B",LAST)}="登録"),--({ex(SI,"D",LAST)}<>"NG")'
    c3d = f',--({ex(SI,"B",LAST)}="削除"),--({ex(SI,"D",LAST)}<>"NG")'
    put(ws, f'X{r}', f'=IF($V{r}="",0,{occ(ex(SE,"C"))}+{occ(ex(SI,"C",LAST),c3)})', fl=F_FX)
    put(ws, f'Y{r}', f'=IF($V{r}="",0,{dis(ex(SE,"C"))}+{dis(ex(SI,"C",LAST),c3)})', fl=F_FX)
    put(ws, f'Z{r}', f'=IF($V{r}="",0,{dis(ex(SI,"C",LAST),c3)})', fl=F_FX)
    put(ws, f'AA{r}', f'=IF($V{r}="",0,{dis(ex(SI,"C",LAST),c3d)})', fl=F_FX)
    put(ws, f'O{r}', f'=IF(OR($C{r}="",$B{r}="削除",$E{r}="{MODE_RAW}",$V{r}=""),"",'
                     f'IF(OR(LEFT($V{r},1)=",",RIGHT($V{r},1)=",",ISNUMBER(FIND(",,",$V{r}))),"空の要素あり",'
                     f'IF($X{r}<$W{r},"未登録あり("&($W{r}-$X{r})&"件)",'
                     f'IF($X{r}>$Y{r},"メンバー重複",'
                     f'IF($AA{r}>0,"削除予定を含む",'
                     f'IF($Z{r}>0,"登録予定を含む","全て存在"))))))', fl=F_FX)
    put(ws, f'P{r}', f'=IF($Q{r}="","",IF($Q{r}="OK","OK",LEFT($Q{r},FIND(":",$Q{r})-1)))', fl=F_FX)
    conds = [
        (f'$B{r}=""', '"NG:操作未選択"'),
        (f'$N{r}="重複"', '"NG:一覧内でnameが重複"'),
        (f'LEN($C{r})>31', '"NG:nameが31文字超"'),
        (f'LENB($C{r})<>LEN($C{r})', '"NG:nameに全角文字が含まれます"'),
    ]
    rest = [
        (f'AND($B{r}="登録",$J{r}="あり")', '"NG:既存に同名あり(登録不可)"'),
        (f'AND($B{r}<>"登録",$J{r}="なし")', '"NG:既存エントリ無"'),
        (f'AND($B{r}<>"削除",$E{r}="")', '"NG:組み方が未選択"'),
        (f'AND($B{r}<>"削除",$E{r}<>"{MODE_RAW}",$R{r}="",$S{r}="")', '"NG:HIPオブジェクトが未選択(F列かG列)"'),
        (f'AND($B{r}<>"削除",$E{r}="{MODE_RAW}",$H{r}="")', '"NG:組み方が直接記述のときはH列にmatchを入力"'),
        (f'AND($B{r}<>"削除",LEN($I{r})>2048)', '"NG:matchが2048文字超"'),
        (f'AND($B{r}<>"削除",LEN($D{r})>255)', '"NG:descriptionが255文字超"'),
        (f'$O{r}="空の要素あり"', '"NG:F・G列の先頭・末尾の区切り(カンマ/改行)や連続した区切りを削除してください"'),
        (f'LEFT($O{r},5)="未登録あり"', '"NG:②の既存一覧にも③の登録予定にも無いHIPオブジェクトが含まれます(大文字・小文字も区別)"'),
        (f'$O{r}="メンバー重複"', '"NG:同じHIPオブジェクトが重複しています(対象と除外の両方に入れていないか確認)"'),
        (f'$O{r}="削除予定を含む"', '"NG:③で削除予定のHIPオブジェクトを含みます"'),
        (f'AND($B{r}<>"削除",$E{r}="{MODE_RAW}")', '"警告:直接記述のmatchはJSON発行時にVBAが参照先を確認します"'),
        (f'$O{r}="登録予定を含む"', '"警告:③で登録予定のHIPオブジェクトを含む(④で先に登録します)"'),
        (f'AND($B{r}="変更",$D{r}&""=$L{r},$I{r}=$M{r})', '"警告:既存と差分なし"'),
        (f'AND($B{r}="登録",COUNTIF({ex(SE,"C")},$C{r})>0)', '"警告:同名のHIPオブジェクトが既存にあります"'),
    ]
    reason = nest(conds, NAMECHK.format(r=r) + nest(rest, '"OK"') + ')')
    put(ws, f'Q{r}', f'=IF($C{r}="","",{reason})', fl=F_FX, al=Alignment(wrap_text=True, vertical='top'))
ws.freeze_panes = 'D8'
dvlist(ws, f'B{FIRST}:B{LAST}', '"登録,変更,削除"')
dvlist(ws, f'C{FIRST}:C{LAST}', 'HIPPROF_NAMES', err=False)
dvlist(ws, f'E{FIRST}:E{LAST}', f'"{MODE_AND},{MODE_OR},{MODE_RAW}"')
cf(ws, f'P{FIRST}:Q{LAST}', f'$P{FIRST}="NG"', 'FFC7CE')
cf(ws, f'P{FIRST}:Q{LAST}', f'$P{FIRST}="警告"', 'FFEB9C')
cf(ws, f'P{FIRST}:Q{LAST}', f'$P{FIRST}="OK"', 'C6EFCE')
cf(ws, f'D{FIRST}:D{LAST}', f'AND($B{FIRST}="変更",$J{FIRST}="あり",$D{FIRST}&""<>$L{FIRST})', 'F8CBAD')
cf(ws, f'I{FIRST}:I{LAST}', f'AND($B{FIRST}="変更",$J{FIRST}="あり",$I{FIRST}<>$M{FIRST})', 'F8CBAD')
cf(ws, f'F{FIRST}:H{LAST}', f'AND($E{FIRST}="{MODE_RAW}",COLUMN()<8)', 'D9D9D9')
cf(ws, f'H{FIRST}:H{LAST}', f'AND($E{FIRST}<>"{MODE_RAW}",$B{FIRST}<>"")', 'D9D9D9')
for i, v in enumerate(prof_in):
    for j, x in enumerate(v):
        ws[f'{CL(2+j)}{FIRST+i}'] = x if x != '' else None

# =========================================================
# HIP_選択肢（プルダウン候補・書式ガイド）
# =========================================================
ws = sheet(SC, 5)
header(ws, 'HIP_選択肢：プルダウン候補と、1行1件の欄の書式', 'B列のOS値は参考の候補です（SCM画面や既存一覧の値に合わせて自由に追加・修正できます。候補に無い値も入力可）。'
       'D列の証明書プロファイル名は「既存情報取得」でVBAが書き込みます。F列以降は③の複数行の欄の書き方です（「!」は否定）。', 9, 40)
heads(ws, 4, ['', 'OS 値の候補（編集可）', '', '証明書プロファイル名\n（VBA書込み）', '', '欄', '書式', '例', '対応するAPIの項目'],
      [2, 30, 2, 28, 2, 22, 44, 44, 40])
os_vals = ['All', 'Microsoft Windows 11', 'Microsoft Windows 10', 'Microsoft Windows Server 2022', 'Microsoft Windows Server 2019',
           'Mac OS X', 'iOS', 'Android', 'Chrome OS', 'Linux']
for i, v in enumerate(os_vals):
    put(ws, f'B{5+i}', v, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
for i in range(5 + len(os_vals), 5 + 40):
    put(ws, f'B{i}', None, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
put(ws, 'D5', None, fl=F_VBA, nf='@')
guide = [
    ('ベンダー | 製品\n（各カテゴリ共通）', 'セルをダブルクリックして HIP_ベンダー一覧 から選ぶ（手入力も可）\nベンダー名 | 製品1 ; 製品2\n製品を指定しないときはベンダー名だけ', 'Microsoft Corporation | Microsoft Defender Antivirus\nCrowdStrike, Inc.', 'vendor[].name / vendor[].product[]'),
    ('ベンダー除外', 'yes = 指定したベンダー以外が対象', 'yes', 'exclude_vendor'),
    ('暗号化場所 | 状態', '場所 | 状態（状態を省略すると encrypted）\n先頭に ! を付けると「状態ではない」(is_not)\n状態 = encrypted / unencrypted / partial / unknown', 'C:\\ | encrypted\n!D:\\ | unencrypted', 'encrypted_locations[].name / encryption_state.is / is_not'),
    ('パッチ', '1行に1つ', 'KB5031356', 'missing_patches.patches[]'),
    ('証明書属性 | 値', '属性名 | 値（値は省略可）', 'Common Name | pc-001.corp.example.local', 'certificate_attributes[].name / value'),
    ('プロセス', 'プロセス名（実行中であること）\n先頭に ! を付けると「実行していないこと」', 'CSFalconService.exe\n!torrent.exe', 'process_list[].name / running'),
    ('レジストリ', '[!]キー … キーがあること（! で無いこと）\nキー | [!]値名 | データ … 値の条件（データは省略可）\nキー | (既定) | データ … 既定値のデータ\n同じキーの行は1つのキーにまとめます',
     'HKLM\\SOFTWARE\\Corp\\Agent\nHKLM\\SOFTWARE\\Corp\\Agent | Version | 5.2\n!HKLM\\SOFTWARE\\BadApp', 'registry_key[].name / negate / default_value_data / registry_value[]'),
    ('plist', '[!]plist名 … plistがあること（! で無いこと）\nplist名 | [!]キー | 値 … キーの条件（値は省略可）', 'com.corp.agent\ncom.corp.agent | Enabled | true', 'plist[].name / negate / key[]'),
    ('OS', 'ベンダーを選び、値は候補から選ぶか入力（空欄=All）。Other は値が必須', 'Microsoft / Microsoft Windows 11', 'host_info.criteria.os.contains.<ベンダー>'),
    ('ドメイン等（条件+値）', '条件 is / is_not / contains と値の組み合わせ', 'is / corp.example.local', 'host_info.criteria.domain など'),
    ('定義ファイル・最終スキャン・最終バックアップ', '条件 within / not_within（スキャン・バックアップは not_available も）+ 単位 + 数', 'within / days / 3', 'virdef_version / last_scan_time / last_backup_time'),
    ('製品バージョン', '条件が within / not_within のときは値にバージョン数(1〜65535)、それ以外は文字列', 'greater_equal / 4.18', 'anti_malware.criteria.product_version'),
    ('ネットワーク', '条件 is / is_not + 種別。wifi は SSID、mobile はキャリアを入力可。ethernet は is_not のみ', 'is_not / mobile / docomo', 'network_info.criteria.network'),
    ('モバイルデバイス', 'このツールでは編集しません（②E列に表示、変更時は既存の内容を引き継ぎ）', '', 'mobile_device'),
]
for i, g in enumerate(guide):
    r = 5 + i
    for col, v in zip('FGHI', g):
        put(ws, f'{col}{r}', v or None, al=Alignment(wrap_text=True, vertical='top'))
    ws.row_dimensions[r].height = max(30, 15 * max(str(x).count('\n') + 1 for x in g))


# =========================================================
# HIP_ベンダー一覧（ベンダー欄のダブルクリック選択の候補）
# =========================================================
ws = sheet(SV, 6)
header(ws, 'HIP_ベンダー一覧：ベンダー欄（ダブルクリック選択）の候補',
       '③のベンダー欄をダブルクリックすると、同じカテゴリの行がチェックリストで出ます。「ベンダーだけ」を選ぶと全製品、「ベンダー | 製品」を選ぶとその製品に限定します。'
       '初期候補は代表的なものの参考値です。SCM／GlobalProtectが使うベンダー名・製品名と完全に一致する必要があるため、「既存情報取得」を実行すると既存のHIPオブジェクトで使われている名前をVBAが自動で追加します（E列=既存から取得）。'
       '足りないものは行を追加してください（B列はプルダウン、製品は空欄可）。', 6, 56)
heads(ws, 5, ['No', 'カテゴリ', 'ベンダー', '製品（空欄=ベンダーのみ）', '出所'], [5, 18, 30, 40, 16])
VSEED = [
    ('マルウェア対策', 'Microsoft Corporation', 'Microsoft Defender Antivirus'),
    ('マルウェア対策', 'CrowdStrike, Inc.', 'CrowdStrike Falcon'),
    ('マルウェア対策', 'Palo Alto Networks', 'Cortex XDR'),
    ('マルウェア対策', 'Trend Micro, Inc.', 'Trend Micro Apex One'),
    ('マルウェア対策', 'Symantec Corporation', 'Symantec Endpoint Protection'),
    ('マルウェア対策', 'McAfee, Inc.', 'McAfee Endpoint Security'),
    ('マルウェア対策', 'Sophos Limited', 'Sophos Endpoint Agent'),
    ('マルウェア対策', 'ESET', 'ESET Endpoint Security'),
    ('マルウェア対策', 'SentinelOne', 'SentinelOne Agent'),
    ('マルウェア対策', 'Apple Inc.', 'XProtect'),
    ('ディスク暗号化', 'Microsoft Corporation', 'BitLocker Drive Encryption'),
    ('ディスク暗号化', 'Apple Inc.', 'FileVault'),
    ('ディスク暗号化', 'Symantec Corporation', 'Symantec Endpoint Encryption'),
    ('ディスク暗号化', 'Sophos Limited', 'Sophos SafeGuard'),
    ('ファイアウォール', 'Microsoft Corporation', 'Windows Firewall'),
    ('ファイアウォール', 'Apple Inc.', 'Mac OS X Builtin Firewall'),
    ('ファイアウォール', 'Symantec Corporation', 'Symantec Endpoint Protection'),
    ('ファイアウォール', 'ESET', 'ESET Endpoint Security'),
    ('パッチ管理', 'Microsoft Corporation', 'Windows Update Agent'),
    ('パッチ管理', 'Apple Inc.', 'Software Update'),
    ('パッチ管理', 'Ivanti', 'Ivanti Patch for Windows'),
    ('ディスクバックアップ', 'Microsoft Corporation', 'Windows Backup and Restore'),
    ('ディスクバックアップ', 'Apple Inc.', 'Time Machine'),
    ('データ損失防止', 'Symantec Corporation', 'Symantec DLP Agent'),
    ('データ損失防止', 'Forcepoint', 'Forcepoint DLP Endpoint'),
    ('データ損失防止', 'Digital Guardian', 'Digital Guardian Agent'),
]
for i, (c, v, p) in enumerate(VSEED):
    r = 6 + i
    put(ws, f'A{r}', f'=IF(C{r}="","",ROW()-5)')
    for col, x in zip('BCDE', [c, v, p, '初期候補（参考）']):
        put(ws, f'{col}{r}', x, Font(name=FN, size=10, color='0000FF'), F_IN, nf='@')
dvlist(ws, 'B6:B2000', '"マルウェア対策,ディスクバックアップ,ディスク暗号化,ファイアウォール,パッチ管理,データ損失防止"')
ws.freeze_panes = 'C6'
ws.auto_filter.ref = 'A5:E2000'
# =========================================================
# ⑧⑨ 差分確認
# =========================================================
def diffsheet(title, idx, heading, fields, widths, rows):
    ws = sheet(title, idx)
    ncol = 3 + len(fields) + 5
    header(ws, heading,
           '④のAPI実行が終わると、VBAがSCMに今設定されている全件を取得して一覧表示し、既存一覧シート（実行前）と比べた変化をC列に付けます。'
           '行の背景色で変化を表します（追加=緑、変更=黄、変化なし=白、削除=グレー）。変更があった項目のセルは橙色・太字です。削除されたものは一覧の後ろに、既存一覧（実行前）の内容で表示します。'
           '判定が要確認のもの（判定列が赤）はPush前に理由を確認してください。ボタンでいつでも作り直せます。', ncol)
    put(ws, 'A3', '作成日時', Font(name=FN, size=10, bold=True))
    put(ws, 'B3', None, fl=F_VBA, nf='yyyy/mm/dd hh:mm:ss')
    for lab, val in [(('C3', 'D3'), '全件'), (('E3', 'F3'), '変化なし'), (('G3', 'H3'), '追加'),
                     (('C4', 'D4'), '変更'), (('E4', 'F4'), '削除'), (('G4', 'H4'), '要確認')]:
        stat(ws, lab, val, None, F_VBA)
    btn(ws, 'J3:K4', '▶ 差分チェック（再作成）')
    put(ws, 'C5', '凡例：', Font(name=FN, size=10, bold=True), border=False)
    for a, t, c in [('D5', '追加', 'C6EFCE'), ('E5', '変更', 'FFFF99'), ('F5', '変化なし', 'FFFFFF'), ('G5', '削除', 'D9D9D9'),
                    ('H5', '変更セル', 'FFC000'), ('I5', '要確認(判定列)', 'FFC7CE')]:
        put(ws, a, t, fl=fill(c), al=Alignment(horizontal='center'), font=Font(name=FN, size=10, bold=(t == '変更セル')))
    names = ['No', 'name', '変化'] + [f + '（現在）' for f in fields] + ['変更項目', '変更前（既存一覧）', '④の操作', 'HTTP', '判定']
    heads(ws, 7, names, [5, 20, 9] + widths + [18, 40, 9, 7, 26])
    stats = {'': 0, '追加': 0, '変更': 0, '削除': 0, 'chk': 0}
    color = {'追加': 'C6EFCE', '変更': 'FFFF99', '削除': 'D9D9D9'}
    for i, (nm, ch, vals, changed, before, op, st, judge) in enumerate(rows):
        r = FIRST + i
        vv = [i + 1, nm, ch] + list(vals) + [' '.join(changed), before, op, st, judge]
        for j, x in enumerate(vv):
            c = put(ws, f'{CL(1+j)}{r}', x if x != '' else None, al=Alignment(wrap_text=True, vertical='top'))
            if ch in color: c.fill = fill(color[ch])
        for f in changed:
            c = ws[f'{CL(4 + fields.index(f))}{r}']; c.fill = fill('FFC000'); c.font = Font(name=FN, size=10, bold=True)
        stats['' if ch == '変化なし' else ch] += 1
        if judge not in ('', '想定どおり') and not judge.startswith('変化なし'):
            stats['chk'] += 1; ws[f'{CL(ncol)}{r}'].fill = fill('FFC7CE')
    ws['D3'] = len(rows) - stats['削除']; ws['F3'] = stats['']; ws['H3'] = stats['追加']
    ws['D4'] = stats['変更']; ws['F4'] = stats['削除']; ws['H4'] = stats['chk']
    ws['B3'] = datetime.datetime(2026, 10, 28, 16, 0)
    ws.freeze_panes = 'D8'

DC = ['description', 'host_info', 'anti_malware', 'disk_backup', 'disk_encryption', 'firewall', 'patch_management',
      'data_loss_prevention', 'certificate', 'network_info', 'custom_checks', 'mobile_device']
def dvals(o): return [o.get('description', '')] + [H.flat_text(o.get(c)) if c in o else '' for c in DC[1:]]
exmap = {o['name']: o for o in ex_objs}
inmap = {nm: H.row_to_obj(nm, v, exmap.get(nm)) for op, nm, v in obj_in[:4]}
def chrow(nm, op, st):
    b, a = dvals(exmap[nm]), dvals(inmap[nm])
    ch = [DC[i] for i in range(len(DC)) if a[i] != b[i]]
    desc = '\n'.join(f'[{DC[i]}]\n{b[i]}' if b[i] else f'[{DC[i]}] （なし）' for i in range(len(DC)) if a[i] != b[i])
    return (nm, '変更' if ch else '変化なし', a, ch, desc, op, st, '想定どおり' if ch else '変化なし（既存と同じ内容）')
obj_rows = [
    ('hip-win-managed', '変化なし', dvals(exmap['hip-win-managed']), [], '', '', '', ''),
    chrow('hip-av-defender', '変更', '200'),
    ('hip-disk-enc', '変化なし', dvals(exmap['hip-disk-enc']), [], '', '', '', ''),
    chrow('hip-fw-on', '変更', '200'),
    ('hip-patch', '変化なし', dvals(exmap['hip-patch']), [], '', '', '', ''),
    ('hip-mac', '変化なし', dvals(exmap['hip-mac']), [], '', '', '', ''),
    ('hip-android', '変化なし', dvals(exmap['hip-android']), [], '', '', '', ''),
    ('hip-win11', '追加', dvals(inmap['hip-win11']), [], '', '登録', '201', '想定どおり'),
    ('hip-crowdstrike', '追加', dvals(inmap['hip-crowdstrike']), [], '', '登録', '201', '想定どおり'),
    ('hip-old-check', '削除', dvals(exmap['hip-old-check']), [], '', '削除', '200', '想定どおり'),
]
diffsheet(SD, 7, '⑧ 差分確認（HIPオブジェクト）：現在の設定一覧と既存一覧の比較', DC,
          [18, 26, 30, 24, 26, 22, 28, 22, 22, 22, 30, 22], obj_rows)
corp_new = '"hip-win-managed" and "hip-av-defender" and "hip-disk-enc" and "hip-fw-on"'
prof_rows = [
    ('prof-corp-win', '変更', ['社用Windowsの基準（FW追加）', corp_new], ['description', 'match'],
     '[description]\n社用Windowsの基準\n[match]\n' + ex_profs[0][2], '変更', '200', '想定どおり'),
    ('prof-mac', '変化なし', ['社用Mac', '"hip-mac"'], [], '', '', '', ''),
    ('prof-complex', '変化なし', ['複合条件の例', ex_profs[3][2]], [], '', '', '', ''),
    ('prof-win11', '追加', ['Windows 11 基準', '"hip-win11" and "hip-crowdstrike"'], [], '', '登録', '201', '想定どおり'),
    ('prof-not-legacy', '追加', ['WindowsかMac（旧エージェント除く）', '("hip-win-managed" or "hip-mac") and not "hip-fw-on"'], [], '', '登録', '201', '想定どおり'),
    ('prof-legacy', '削除', ['旧エージェント端末', '"hip-old-check"'], [], '', '削除', '200', '想定どおり'),
]
diffsheet(PD, 8, '⑨ 差分確認（HIPプロファイル）：現在の設定一覧と既存一覧の比較', ['description', 'match'], [26, 60], prof_rows)

# =========================================================
# ④ JSON_CURL サンプル（VBAのJSON発行と同じ形式）
# =========================================================
BASE = 'https://api.sase.paloaltonetworks.com'
def jd(o): return json.dumps(o, ensure_ascii=False, separators=(',', ':'))
def curl(m, u, b):
    c = f"curl -s -X {m} '{u}' -H \"Authorization: Bearer $TOKEN\""
    if b: c += " -H 'Content-Type: application/json' -d '" + b.replace("'", "'\\''") + "'"
    return c
O, P = BASE + '/sse/config/v1/hip-objects', BASE + '/sse/config/v1/hip-profiles'
oid = {o['name']: f'7a3c{i+1:04d}-1b2c-4d5e-8f90-0000000000{i+1:02d}' for i, o in enumerate(ex_objs)}
pid = {p[0]: f'8b4d{i+1:04d}-2c3d-4e5f-9a01-0000000000{i+1:02d}' for i, p in enumerate(ex_profs)}
def pbody(nm, d, m):
    b = {'name': nm}
    if d: b['description'] = d
    b['match'] = m
    return jd(b)
rows = [
    ('登録', 'hip-win11', 'POST', O + '?folder=Shared', jd(inmap['hip-win11']), 201),
    ('登録', 'hip-crowdstrike', 'POST', O + '?folder=Shared', jd(inmap['hip-crowdstrike']), 201),
    ('変更', 'hip-av-defender', 'PUT', O + '/' + oid['hip-av-defender'], jd(inmap['hip-av-defender']), 200),
    ('変更', 'hip-fw-on', 'PUT', O + '/' + oid['hip-fw-on'], jd(inmap['hip-fw-on']), 200),
    ('登録(Profile)', 'prof-win11', 'POST', P + '?folder=Shared', pbody('prof-win11', 'Windows 11 基準', '"hip-win11" and "hip-crowdstrike"'), 201),
    ('変更(Profile)', 'prof-corp-win', 'PUT', P + '/' + pid['prof-corp-win'], pbody('prof-corp-win', '社用Windowsの基準（FW追加）', corp_new), 200),
    ('登録(Profile)', 'prof-not-legacy', 'POST', P + '?folder=Shared',
     pbody('prof-not-legacy', 'WindowsかMac（旧エージェント除く）', '("hip-win-managed" or "hip-mac") and not "hip-fw-on"'), 201),
    ('削除(Profile)', 'prof-legacy', 'DELETE', P + '/' + pid['prof-legacy'], '', 200),
    ('削除', 'hip-old-check', 'DELETE', O + '/' + oid['hip-old-check'], '', 200),
]
js = wb[SJ]
for i, (op, nm, m, u, b, st) in enumerate(rows):
    r = 10 + i
    for j, v in enumerate([i + 1, op, nm, m, u, b or None, curl(m, u, b), st]):
        put(js, f'{CL(1+j)}{r}', v, al=Alignment(wrap_text=True, vertical='top'))

# 名前定義
wb.defined_names['EX_FETCHED_AT'] = DefinedName('EX_FETCHED_AT', attr_text=f"{Q(SE)}!$B$3")
wb.defined_names['EXP_FETCHED_AT'] = DefinedName('EXP_FETCHED_AT', attr_text=f"{Q(PE)}!$B$3")
wb.defined_names['HIPOBJ_NAMES'] = DefinedName('HIPOBJ_NAMES', attr_text=f"OFFSET({Q(SE)}!$C$8,0,0,MAX(1,COUNTA({Q(SE)}!$C$8:$C${EXL})),1)")
wb.defined_names['HIPPROF_NAMES'] = DefinedName('HIPPROF_NAMES', attr_text=f"OFFSET({Q(PE)}!$C$8,0,0,MAX(1,COUNTA({Q(PE)}!$C$8:$C${PEXL})),1)")
wb.defined_names['OS_VALUES'] = DefinedName('OS_VALUES', attr_text=f"OFFSET({Q(SC)}!$B$5,0,0,MAX(1,COUNTA({Q(SC)}!$B$5:$B$44)),1)")
wb.defined_names['CERT_TOP'] = DefinedName('CERT_TOP', attr_text=f"{Q(SC)}!$D$5")
wb.defined_names['CERT_NAMES'] = DefinedName('CERT_NAMES', attr_text=f"OFFSET({Q(SC)}!$D$5,0,0,MAX(1,COUNTA({Q(SC)}!$D$5:$D$1005)),1)")
wb[SE]['B3'] = datetime.datetime(2026, 10, 28, 15, 50)
wb[PE]['B3'] = datetime.datetime(2026, 10, 28, 15, 50)

order = ['基本設定', SE, SI, PE, PI, SC, SV, SJ, SD, PD, SL]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for w in wb: w.sheet_view.tabSelected = (w.title == '基本設定')
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('saved', OUT)

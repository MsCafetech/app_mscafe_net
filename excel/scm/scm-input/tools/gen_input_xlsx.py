# -*- coding: utf-8 -*-
"""
SCM 統合入力シート（お客様記入用）ジェネレータ
  - 既存ツール（scm-excel / scm-service / scm-hip / scm-edl）の「登録変更削除」シートを
    数式・入力規則・条件付き書式ごと取り込み、参照先を新しいシートに付け替える
  - ② 既存データ抽出 は各ツールの「既存一覧」と同じ列並びのブロックを横に並べた1枚のシート
  - ① 基本設定・実行ログ は既存ツール（SCM_SecurityPolicy.xlsx）と同じ（トークン取得・API取得（GETのみ）用）
  - ⑦ カスタムURLカテゴリ、⑩ アプリケーション、⑪ ユーザ/グループ、記入状況 は新規に作成
使い方: python3 gen_input_xlsx.py            -> SCM_入力シート.xlsx（空欄）
        python3 gen_input_xlsx.py --sample   -> SCM_入力シート_サンプル入り.xlsx
"""
import re
import sys
from copy import copy

import openpyxl
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.utils import column_index_from_string as CI
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.styles.differential import DifferentialStyle

PF = "/mnt/project-files"
SRC = {
    "adr": f"{PF}/scm-excel/SCM_AddressObject.xlsx",
    "svc": f"{PF}/scm-service/SCM_ServiceObject.xlsx",
    "hip": f"{PF}/scm-hip/SCM_HIP.xlsx",
    "edl": f"{PF}/scm-edl/SCM_EDL.xlsx",
    "pol": f"{PF}/scm-policy/SCM_SecurityPolicy.xlsx",
    "dec": f"{PF}/scm-decryption/SCM_DecryptionPolicy.xlsx",
}
SAMPLE = "--sample" in sys.argv
OUT = f"{PF}/scm-input/" + ("SCM_入力シート_サンプル入り.xlsx" if SAMPLE else "SCM_入力シート.xlsx")
for a in sys.argv[1:]:
    if a.startswith("--out="):
        OUT = a[6:]

S1, S2 = "①基本設定", "②既存データ抽出"
SS, SLOG = "記入状況", "実行ログ"
S3, S4, S5, S6 = "③アドレス", "④アドレスグループ", "⑤サービス", "⑥サービスグループ"
S7, S8, S8B, S9 = "⑦カスタムURLカテゴリ", "⑧HIP一覧", "⑧-2HIPプロファイル", "⑨EDL"
S10, S11 = "⑩アプリケーション", "⑪ユーザ・グループ"
R1, R2 = "参照_HIP選択肢", "参照_HIPベンダー一覧"
S2B, S2C, S12, S13, R3 = "②-2既存セキュリティポリシー", "②-3既存復号ポリシー", "⑫セキュリティポリシー", "⑬復号ポリシー", "参照_ポリシー選択肢"

# ---------------------------------------------------------------- styles
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_IN = PatternFill("solid", fgColor="FFF2CC")      # 入力欄（黄）
F_HEAD = PatternFill("solid", fgColor="1F4E78")    # 見出し
F_SUB = PatternFill("solid", fgColor="DDEBF7")     # 補助見出し
F_CALC = PatternFill("solid", fgColor="F2F2F2")    # 関数
F_BTN = PatternFill("solid", fgColor="F4B084")     # ボタン位置
F_BLK = PatternFill("solid", fgColor="E2EFDA")     # ②ブロック見出し
FT_HEAD = Font(bold=True, color="FFFFFF")
FT_TTL = Font(bold=True, size=14, color="1F4E78")
FT_B = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
NG_FILL, WARN_FILL, OK_FILL, CHG_FILL = "FFC7CE", "FFEB9C", "C6EFCE", "F8CBAD"


def dxf(rgb, bold=False):
    return DifferentialStyle(fill=PatternFill("solid", bgColor=rgb, fgColor=rgb), font=Font(bold=bold) if bold else None)


def judge_cf(ws, rng, col_letter, first=8):
    for val, rgb in (("NG", NG_FILL), ("警告", WARN_FILL), ("OK", OK_FILL)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${col_letter}{first}="{val}"'], fill=PatternFill("solid", bgColor=rgb, fgColor=rgb)))


# ---------------------------------------------------------------- ② blocks
# key, title, (src, sheet) or None, ncols, last_row
BLOCKS = [
    ("adr", "アドレス", ("adr", "AdrObj_既存一覧"), 13, 5000),
    ("adrg", "アドレスグループ", ("adr", "AdrGroupObj_既存一覧"), 11, 5000),
    ("svc", "サービス", ("svc", "SvcObj_既存一覧"), 13, 5000),
    ("svcg", "サービスグループ", ("svc", "SvcGroupObj_既存一覧"), 6, 5000),
    ("url", "カスタムURLカテゴリ（一覧）", None, 5, 5000),
    ("urlv", "カスタムURLカテゴリ（URL・1行1件）", None, 3, 30007),
    ("hip", "HIPオブジェクト", ("hip", "HipObj_既存一覧"), 66, 5000),
    ("hipp", "HIPプロファイル", ("hip", "HipProfile_既存一覧"), 8, 2000),
    ("edl", "EDL", ("edl", "EDL_既存一覧"), 19, 5000),
    ("pg", "セキュリティプロファイルグループ", ("pol", "⑦ProfGroup_一覧"), 12, 2000),
    ("prof", "セキュリティプロファイル", None, 4, 5000),
]
OFF = {}
col = 1
for k, _t, _s, n, _l in BLOCKS:
    OFF[k] = col - 1
    col += n + 1
B_LAST = {k: l for k, _t, _s, _n, l in BLOCKS}

SHEET_MAP = {
    "AdrObj_既存一覧": (S2, OFF["adr"]), "AdrGroupObj_既存一覧": (S2, OFF["adrg"]),
    "AdrObj_登録変更削除": (S3, 0), "AdrGroupObj_登録変更削除": (S4, 0),
    "SvcObj_既存一覧": (S2, OFF["svc"]), "SvcGroupObj_既存一覧": (S2, OFF["svcg"]),
    "SvcObj_登録変更削除": (S5, 0), "SvcGroupObj_登録変更削除": (S6, 0),
    "HipObj_既存一覧": (S2, OFF["hip"]), "HipProfile_既存一覧": (S2, OFF["hipp"]),
    "HipObj_登録変更削除": (S8, 0), "HipProfile_登録変更削除": (S8B, 0),
    "HIP_選択肢": (R1, 0), "HIP_ベンダー一覧": (R2, 0),
    "EDL_既存一覧": (S2, OFF["edl"]), "EDL_登録変更削除": (S9, 0),
}
NAME_MAP = {
    "adr": {"EX_NAMES": "ADR_NAMES", "EXG_NAMES": "ADRG_NAMES"},
    "svc": {},
    "hip": {"CERT_NAMES": "HIP_CERT_NAMES", "CERT_TOP": "HIP_CERT_TOP"},
    "edl": {"CERT_NAMES": "EDL_CERT_NAMES", "PRE_NAMES": "EDL_PRE_NAMES"},
    "pol": {"SEC_PROFGROUPS": "PROFGROUP_NAMES"},
    "dec": {"DEC_PROFILES": "DEC_PROF_NAMES", "DEC_CERTS": "DEC_CERT_NAMES"},
}

CELL_RE = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")
COLR_RE = re.compile(r"^(\$?)([A-Z]{1,3})$")


def shift_part(p, off):
    m = CELL_RE.match(p)
    if m:
        return m.group(1) + CL(CI(m.group(2)) + off) + m.group(3) + m.group(4)
    m = COLR_RE.match(p)
    if m:
        return m.group(1) + CL(CI(m.group(2)) + off)
    return p


def shift_ref(ref, off):
    if off == 0:
        return ref
    parts = ref.split(":")
    if len(parts) == 1:
        return shift_part(ref, off) if CELL_RE.match(ref) else ref
    if all(CELL_RE.match(x) for x in parts) or all(COLR_RE.match(x) for x in parts):
        return ":".join(shift_part(x, off) for x in parts)
    return ref


def q(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def remap(f, own_off=0, names=None, texts=None):
    """数式（先頭 = あり/なし）の参照を付け替える"""
    if not isinstance(f, str):
        return f
    eq = f.startswith("=")
    tok = Tokenizer(f if eq else "=" + f)
    for t in tok.items:
        if t.type == "OPERAND" and t.subtype == "RANGE":
            v = t.value
            if "!" in v:
                sh, ref = v.rsplit("!", 1)
                shn = sh[1:-1].replace("''", "'") if sh.startswith("'") else sh
                if shn in SHEET_MAP:
                    ns, off = SHEET_MAP[shn]
                    t.value = q(ns) + "!" + shift_ref(ref, off)
            elif names and v in names:
                t.value = names[v]
            else:
                t.value = shift_ref(v, own_off)
        elif t.type == "OPERAND" and t.subtype == "TEXT" and texts:
            for a, b in texts:
                if a in t.value:
                    t.value = t.value.replace(a, b)
    out = tok.render()
    return out if eq else out[1:]


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def copy_dv(src_ws, dst_ws, own_off, names, min_row=1):
    for dv in src_ws.data_validations.dataValidation:
        rngs = []
        for r in str(dv.sqref).split():
            if ":" in r:
                a, b = r.split(":")
            else:
                a = b = r
            if int(re.sub(r"[A-Z$]", "", a)) < min_row:
                continue
            rngs.append(shift_ref(a.replace("$", ""), own_off) + ":" + shift_ref(b.replace("$", ""), own_off))
        if not rngs:
            continue
        nd = DataValidation(type=dv.type, formula1=remap(dv.formula1, own_off, names) if dv.formula1 else None,
                            formula2=dv.formula2, allow_blank=dv.allow_blank, showDropDown=dv.showDropDown,
                            showInputMessage=dv.showInputMessage, showErrorMessage=dv.showErrorMessage,
                            prompt=dv.prompt, promptTitle=dv.promptTitle, error=dv.error, errorTitle=dv.errorTitle,
                            errorStyle=dv.errorStyle, operator=dv.operator)
        nd.sqref = openpyxl.worksheet.cell_range.MultiCellRange(" ".join(rngs))
        dst_ws.add_data_validation(nd)


def copy_cf(src_ws, dst_ws, own_off, names, texts=None):
    for cf in src_ws.conditional_formatting:
        rng = " ".join(shift_ref(str(r), own_off) for r in cf.sqref.ranges)
        for r in cf.rules:
            nr = copy(r)
            if r.formula:
                nr.formula = [remap(x, own_off, names, texts) for x in r.formula]
            dst_ws.conditional_formatting.add(rng, nr)


# ---------------------------------------------------------------- load sources
src_wb = {k: openpyxl.load_workbook(p) for k, p in SRC.items()}
wb = openpyxl.Workbook()
wb.remove(wb.active)
for nm in (S1, SS, S2, S2B, S2C, S3, S4, S5, S6, S7, S8, S8B, S9, S10, S11, S12, S13, R1, R2, R3, "チェック結果", SLOG):
    wb.create_sheet(nm)

# ======================================================================
# 登録変更削除シートの取り込み
# ======================================================================
JSON_MEMBER = "件)のメンバー。JSON発行時にGroupから外す変更を自動で追加します"
INPUT_SHEETS = {
    # new: (srcKey, srcSheet, input cols, judge col, reason col, helper col ranges, last col, texts, A1, A2, freeze)
    S3: dict(src="adr", sheet="AdrObj_登録変更削除", inputs="B:G", judge="P", reason="Q",
             helpers=["H:O", "R:X"], texts=[(JSON_MEMBER, "件)のメンバーです。④でそのグループのメンバーから外してください")],
             a1="③ アドレスオブジェクト（登録・変更・削除）",
             a2="黄色(B〜G列)に記入してください。操作は「登録」「変更」「削除」から選択。変更・削除のnameは既存のnameをプルダウンから選べます。"
                "P列（判定）・Q列（理由）は関数で②既存データと自動照合します（NG=赤：修正が必要、警告=黄：確認してください）。"
                "変更の行で既存と異なる項目は橙色になります。変更は全項目の置き換えになるため、残したい値もすべて記入してください"
                "（「変更行に既存値を読込」で現在の値を入れてから直すと簡単です）。tagは「,」区切り。"),
    S4: dict(src="adr", sheet="AdrGroupObj_登録変更削除", inputs="B:G", judge="P", reason="Q", helpers=["H:O"],
             texts=[("NG:②既存一覧にも③にも無いメンバーが含まれます(大文字・小文字も区別)",
                     "NG:②既存データにも③の登録予定にも無いメンバーが含まれます(大文字・小文字も区別)"),
                    ("警告:③で削除予定のメンバーを含む(JSON発行時にメンバーから外します)",
                     "NG:③で削除予定のアドレスがメンバーに含まれます(メンバーから外してください)"),
                    ("警告:③で登録予定のメンバーを含む(④でAddressの登録後に実行します)", "OK")],
             a1="④ アドレスグループ（登録・変更・削除）",
             a2="黄色(B〜G列)に記入してください。種別が static のときは、値（メンバー）のセルをダブルクリックするとアドレスの一覧からチェックボックスで選べます"
                "（Ctrl+Shift+M でも可。手入力するときは「,」区切りかセル内改行）。dynamic のときはタグ条件式（例：'prod' and 'db'）を記入します。"
                "メンバーは②既存データか③で登録するアドレスである必要があります（大文字・小文字も区別）。P列（判定）・Q列（理由）は関数で自動チェックします。"),
    S5: dict(src="svc", sheet="SvcObj_登録変更削除", inputs="B:K", judge="X", reason="Y", helpers=["L:W", "Z:AB"],
             texts=[("NG:⑦で同名のServiceGroupを登録予定", "NG:⑥で同名のサービスグループを登録予定"),
                    (JSON_MEMBER, "件)のメンバーです。⑥でそのグループのメンバーから外してください")],
             a1="⑤ サービス（登録・変更・削除）",
             a2="黄色(B〜K列)に記入してください。portは「80」「8000-8080」「80,443,8000-8080」の形式（0〜65535、カンマ区切り、空白不可）。source_portも同じ形式で任意。"
                "timeout系は通常は空欄（既定値）。X列（判定）・Y列（理由）は関数で②既存データと自動照合します。変更は全項目の置き換えになるため、残したい値もすべて記入してください。"),
    S6: dict(src="svc", sheet="SvcGroupObj_登録変更削除", inputs="B:E", judge="L", reason="M", helpers=["F:K", "N:S"],
             texts=[("NG:②⑥の既存一覧にも③⑦の登録予定にも無いメンバーが含まれます(大文字・小文字も区別)",
                     "NG:②既存データにも⑤⑥の登録予定にも無いメンバーが含まれます(大文字・小文字も区別)"),
                    ("NG:③で同名のServiceを登録予定", "NG:⑤で同名のサービスを登録予定"),
                    (JSON_MEMBER, "件)のメンバーです。そのグループのメンバーから外してください"),
                    ("警告:③⑦で削除予定のメンバーを含む(JSON発行時にメンバーから外します)",
                     "NG:⑤⑥で削除予定のメンバーが含まれます(メンバーから外してください)"),
                    ("警告:③⑦で登録予定のメンバーを含む(④で先に登録します)", "OK")],
             a1="⑥ サービスグループ（登録・変更・削除）",
             a2="黄色(B〜E列)に記入してください。メンバー欄（D列）をダブルクリックすると、サービス・サービスグループの一覧からチェックボックスで選べます"
                "（Ctrl+Shift+M でも可。定義済みの service-http / service-https も選べます。手入力は「,」区切りかセル内改行）。"
                "メンバーは②既存データか⑤⑥で登録するものである必要があります（大文字・小文字も区別）。サービスグループには description がありません。"),
    S8: dict(src="hip", sheet="HipObj_登録変更削除", inputs="F:BM", judge="D", reason="E", helpers=["BN:BY"],
             texts=[("(⑦でmatchから外すかプロファイルを削除してください)", "(⑧-2でmatchから外すかプロファイルを削除してください)"),
                    ("(このツールでは編集しない)", "(このシートでは記入不要)")],
             a1="⑧ HIP一覧：HIPオブジェクト（登録・変更・削除）",
             a2="黄色(B・C列とF〜BM列)に記入してください。多くの欄はプルダウンで選びます。ベンダー欄はダブルクリック（またはCtrl+Shift+M）で「参照_HIPベンダー一覧」から選べます。"
                "使わないカテゴリは空欄のままにしてください。1行1件の欄はセル内改行（Alt+Enter）で複数行。書式は「参照_HIP選択肢」シートを参照。"
                "D列（判定）・E列（理由）は関数で②既存データと自動照合します。HIPプロファイルは「⑧-2HIPプロファイル」シートに記入します。"),
    S8B: dict(src="hip", sheet="HipProfile_登録変更削除", inputs="B:H", judge="P", reason="Q", helpers=["J:O", "R:AA"],
              texts=[("NG:②の既存一覧にも③の登録予定にも無いHIPオブジェクトが含まれます(大文字・小文字も区別)",
                      "NG:②既存データにも⑧の登録予定にも無いHIPオブジェクトが含まれます(大文字・小文字も区別)"),
                     ("NG:③で削除予定のHIPオブジェクトを含みます", "NG:⑧で削除予定のHIPオブジェクトを含みます"),
                     ("警告:③で登録予定のHIPオブジェクトを含む(④で先に登録します)", "OK"),
                     ("警告:直接記述のmatchはJSON発行時にVBAが参照先を確認します", "警告:直接記述のmatchは「一括チェック」で参照先を確認します")],
              a1="⑧-2 HIP一覧：HIPプロファイル（登録・変更・削除）",
              a2="黄色(B〜H列)に記入してください。組み方を選び、F列（対象）・G列（除外＝not）はセルをダブルクリックするとHIPオブジェクトの一覧からチェックボックスで選べます（Ctrl+Shift+Mでも可）。"
                 "I列に関数でmatch式を作ります。括弧を使う複雑な式は組み方「直接記述」にしてH列に式を書いてください。"
                 "HIPオブジェクトは②既存データか⑧で登録するものである必要があります（大文字・小文字も区別）。P列（判定）・Q列（理由）は関数で自動チェックします。"),
    S9: dict(src="edl", sheet="EDL_登録変更削除", inputs="B:O", judge="AE", reason="AF", helpers=["P:AD", "AG:AH"],
             texts=[("NG:usernameを使う場合はpasswordも入力(変更は全項目置換。既存のpasswordはSCMから取得できません)",
                     "警告:usernameを使う場合、passwordはこのシートに書かず別途お知らせください")],
             a1="⑨ EDL：外部動的リスト（登録・変更・削除）",
             a2="黄色(B〜N列)に記入してください。type・recurring・at・曜日・日・certificate_profile・expand_domain はプルダウンで選択。urlは http:// または https:// で始まる取得元URL"
                "（type=predefined_ip / predefined_url のときは定義済みリスト名をプルダウンから選択）。exception_list は「,」区切りかセル内改行（Alt+Enter）。"
                "パスワードはこのシートに記入せず、別途お知らせください。AE列（判定）・AF列（理由）は関数で②既存データと自動照合します。"),
}

for new, cfg in INPUT_SHEETS.items():
    sw = src_wb[cfg["src"]][cfg["sheet"]]
    ws = wb[new]
    names = NAME_MAP[cfg["src"]]
    texts = cfg["texts"]
    in_c1, in_c2 = (CI(x) for x in cfg["inputs"].split(":"))
    for row in sw.iter_rows(min_row=1, max_row=507):
        for c in row:
            v = c.value
            d = ws.cell(row=c.row, column=c.column)
            copy_style(c, d)
            if isinstance(v, str) and v.startswith("="):
                d.value = remap(v, 0, names, texts)
            else:
                if c.row >= 8 and not SAMPLE:
                    v = None
                d.value = v
    for m in sw.merged_cells.ranges:
        ws.merge_cells(str(m))
    for k, dim in sw.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
    for r, dim in sw.row_dimensions.items():
        if dim.height:
            ws.row_dimensions[r].height = dim.height
    ws.freeze_panes = sw.freeze_panes
    copy_dv(sw, ws, 0, names)
    copy_cf(sw, ws, 0, names, texts)
    ws["A1"] = cfg["a1"]
    ws["A2"] = cfg["a2"]
    ws.row_dimensions[2].height = 75
    # ボタン位置（旧：JSON発行 / 入力クリア 等）
    for addr in ("F3", "I3", "L3"):
        if ws[addr].value:
            ws[addr].value = None
            ws[addr].fill = PatternFill()
    ws["F3"] = "▶ 変更行に既存値を読込"
    ws["F3"].fill = F_BTN
    ws["F3"].alignment = CENTER
    ws["F3"].font = FT_B
    for h in cfg["helpers"]:
        a, b = h.split(":")
        ws.column_dimensions.group(a, b, hidden=True, outline_level=1)
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.sheet_properties.tabColor = "FFC000"

# 個別の手直し
ws = wb[S9]
ws["O7"] = "password\n（記入しないでください）"
ws["O6"] = "記入不要"
wb[S8B]["I7"].value = wb[S8B]["I7"].value  # そのまま

# ======================================================================
# 参照シート（HIP_選択肢 / HIP_ベンダー一覧）
# ======================================================================
for new, sheet, a2 in (
    (R1, "HIP_選択肢", "⑧HIP一覧のプルダウン候補と、1行1件の欄の書き方です。B列のOS値は参考の候補です（追加・修正できます）。D列の証明書プロファイル名は、分かる範囲で記入するか担当者が記入します。"),
    (R2, "HIP_ベンダー一覧", "⑧HIP一覧のベンダー欄をダブルクリックすると、同じカテゴリの行がチェックリストで出ます。「ベンダーだけ」を選ぶと全製品、「ベンダー | 製品」を選ぶとその製品に限定します。"
                            "初期候補は代表的なものの参考値です。足りないものは行を追加してください（B列はプルダウン、製品は空欄可）。"),
):
    sw = src_wb["hip"][sheet]
    ws = wb[new]
    for row in sw.iter_rows():
        for c in row:
            d = ws.cell(row=c.row, column=c.column)
            copy_style(c, d)
            d.value = remap(c.value, 0, NAME_MAP["hip"]) if isinstance(c.value, str) and c.value.startswith("=") else c.value
    for m in sw.merged_cells.ranges:
        ws.merge_cells(str(m))
    for k, dim in sw.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
    for r, dim in sw.row_dimensions.items():
        if dim.height:
            ws.row_dimensions[r].height = dim.height
    copy_dv(sw, ws, 0, NAME_MAP["hip"])
    ws["A2"] = a2
    ws.sheet_properties.tabColor = "A6A6A6"
if not SAMPLE:
    ws = wb[R1]
    for r in range(5, 1006):
        ws.cell(r, 4).value = None

# ======================================================================
# ② 既存データ抽出
# ======================================================================
ws2 = wb[S2]
ws2["A1"] = "② 既存データ抽出（現在SCMに登録されている設定）"
ws2["A1"].font = FT_TTL
ws2["A2"] = ("担当者が①基本設定でトークンを取得し「SCMから取得（API）」で取得します（GET のみ）。既存の各ツール（SCM_AddressObject.xlsx、SCM_SecurityPolicy.xlsx 等）で取得した既存一覧の取込（「既存ツールから取込」）や、各ツールの既存一覧の8行目以降の貼り付けでも入れられます。"
             "③〜⑨の判定はこのシートと照合します。お客様は参照のみで、編集は不要です。ブロックは横に並んでいます（4行目のリンクで移動できます）。")
ws2["A2"].alignment = WRAP
ws2.merge_cells("A2:M2")
ws2.row_dimensions[2].height = 48
ws2["A3"] = "取得日時"
ws2["A3"].font = FT_B
ws2["B3"].number_format = "yyyy/mm/dd hh:mm"
ws2["B3"].fill = F_CALC
ws2["D3"] = "▶ 既存ツールから取込"
ws2.merge_cells("D3:F3")
ws2["H3"] = "▶ 既存データをクリア"
ws2.merge_cells("H3:J3")
ws2["L3"] = "▶ SCMから取得（API）"
ws2.merge_cells("L3:N3")
for a in ("D3", "H3", "L3"):
    ws2[a].fill = F_BTN
    ws2[a].alignment = CENTER
    ws2[a].font = FT_B
ws2["A4"] = "移動："
ws2["A4"].font = FT_B
ws2.freeze_panes = "A8"


def block_header(k, title, ncols, headers=None, widths=None):
    off = OFF[k]
    c1, c2 = off + 1, off + ncols
    ws2.cell(6, c1).value = f"■ {title}"
    ws2.cell(6, c1).font = FT_B
    ws2.cell(6, c1).fill = F_BLK
    ws2.merge_cells(start_row=6, start_column=c1, end_row=6, end_column=min(c2, c1 + 7))
    ws2.cell(5, c1).value = "件数"
    nc_ = off + (2 if k == "pg" else 3)
    ws2.cell(5, c1 + 1).value = f"=COUNTA({CL(nc_)}8:{CL(nc_)}{B_LAST[k]})"
    ws2.cell(5, c1 + 1).fill = F_CALC
    if headers:
        for i, h in enumerate(headers):
            c = ws2.cell(7, c1 + i)
            c.value = h
            c.font = FT_HEAD
            c.fill = F_HEAD
            c.alignment = CENTER
            c.border = BOX
    if widths:
        for i, w in enumerate(widths):
            ws2.column_dimensions[CL(c1 + i)].width = w


link_col = 2
for k, title, src, ncols, last in BLOCKS:
    off = OFF[k]
    # 移動リンク
    lc = ws2.cell(4, link_col)
    lc.value = title.replace("カスタムURLカテゴリ", "URLカテゴリ")
    lc.hyperlink = f"#{q(S2)}!{CL(off + 1)}6"
    lc.font = Font(color="0563C1", underline="single")
    link_col += 1
    if src:
        sw = src_wb[src[0]][src[1]]
        names = NAME_MAP[src[0]]
        block_header(k, title, ncols)
        # 見出し(7行目)
        for c in range(1, ncols + 1):
            s = sw.cell(7, c)
            d = ws2.cell(7, off + c)
            copy_style(s, d)
            d.value = s.value
            w = sw.column_dimensions[CL(c)].width
            if w:
                ws2.column_dimensions[CL(off + c)].width = w
        # 関数の列は最終行まで、値の列はサンプルのみ
        fcols = {c for c in range(1, ncols + 1) if isinstance(sw.cell(8, c).value, str) and str(sw.cell(8, c).value).startswith("=")}
        for c in fcols:
            base = sw.cell(8, c).value
            from openpyxl.formula.translate import Translator
            tr = Translator(remap(base, off, names), origin=f"{CL(off + c)}8")
            for r in range(8, last + 1):
                d = ws2.cell(r, off + c)
                d.value = tr.translate_formula(f"{CL(off + c)}{r}")
        sample_rows = [r for r in range(8, sw.max_row + 1) if sw.cell(r, 2 if k == "pg" else 3).value]
        for r in range(8, (max(sample_rows) if sample_rows else 7) + 1):
            for c in range(1, ncols + 1):
                s = sw.cell(r, c)
                d = ws2.cell(r, off + c)
                copy_style(s, d)
                if c not in fcols and not (k == "pg" and c == 1):
                    d.value = s.value if SAMPLE else None
        if k == "pg":
            for r in range(8, last + 1):
                ws2.cell(r, off + 1).value = f'=IF({CL(off + 2)}{r}="","",ROW()-7)'
        if k == "hipp":
            ws2.cell(7, off + 8).value = "作業:⑧-2にあり"
        if k == "edl":
            # 候補列（証明書プロファイル / 定義済みリスト）は空欄版でも残す（定義済みリストの名前）
            for r in range(8, 28):
                if sw.cell(r, 19).value:
                    ws2.cell(r, off + 19).value = sw.cell(r, 19).value
                if SAMPLE and sw.cell(r, 17).value:
                    ws2.cell(r, off + 17).value = sw.cell(r, 17).value
    else:
        if k == "url":
            block_header(k, title, ncols, ["No", "name", "type", "description", "URL件数(関数)"], [6, 26, 14, 30, 10])
            nc, vc = CL(OFF["urlv"] + 2), CL(OFF["urlv"] + 3)
            for r in range(8, last + 1):
                ws2.cell(r, off + 1).value = f'=IF({CL(off + 2)}{r}="","",ROW()-7)'
                ws2.cell(r, off + 5).value = f'=IF({CL(off + 2)}{r}="","",COUNTIFS(${nc}$8:${nc}${B_LAST["urlv"]},{CL(off + 2)}{r},${vc}$8:${vc}${B_LAST["urlv"]},"<>"))'
        elif k == "prof":
            block_header(k, title, ncols, ["No", "プロファイルの種類", "name", "備考"], [6, 22, 26, 30])
            for r in range(8, last + 1):
                ws2.cell(r, off + 1).value = f'=IF({CL(off + 3)}{r}="","",ROW()-7)'
            if SAMPLE:
                pr = [("ウイルス・WildFire", "AV-Strict"), ("スパイウェア", "AS-Strict"), ("脆弱性防御", "VP-Strict"),
                      ("URLフィルタリング", "URL-Strict"), ("ファイルブロック", "FB-Strict"), ("DNSセキュリティ", "DNS-Strict")]
                for i, (t, n) in enumerate(pr):
                    ws2.cell(8 + i, off + 2).value, ws2.cell(8 + i, off + 3).value = t, n
                    ws2.cell(8 + i, off + 4).value = "プロファイルグループから抽出"
        else:
            block_header(k, title, ncols, ["No", "name（カテゴリ）", "URL（1行1件）"], [6, 26, 40])
            for r in range(8, last + 1):
                ws2.cell(r, off + 1).value = f'=IF({CL(off + 2)}{r}="","",ROW()-7)'

# ② の URL サンプル
if SAMPLE:
    ou, ov = OFF["url"], OFF["urlv"]
    cats = [("WL_corp-allow", "URL List", "ホワイトリスト", ["*.example.co.jp/", "portal.example.com/", "intra.example.net/"]),
            ("BL_block", "URL List", "ブラックリスト", ["*.badsite.example/", "game.example.org/"]),
            ("cat-saas", "URL List", "業務SaaS", ["app.saas-a.example/", "*.saas-b.example/"])]
    r = 8
    rv = 8
    for nm, ty, de, urls in cats:
        ws2.cell(r, ou + 2).value, ws2.cell(r, ou + 3).value, ws2.cell(r, ou + 4).value = nm, ty, de
        r += 1
        for u in urls:
            ws2.cell(rv, ov + 2).value, ws2.cell(rv, ov + 3).value = nm, u
            rv += 1
    ws2["B3"] = "=DATE(2026,11,16)+TIME(9,0,0)"

# ② 名前定義
def add_name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def blk_names(k, col=3):
    c = CL(OFF[k] + col)
    return f"OFFSET({q(S2)}!${c}$8,0,0,MAX(1,COUNTA({q(S2)}!${c}$8:${c}${B_LAST[k]})),1)"


add_name("ADR_NAMES", blk_names("adr"))
add_name("ADRG_NAMES", blk_names("adrg"))
add_name("SVC_NAMES", blk_names("svc"))
add_name("SVCG_NAMES", blk_names("svcg"))
add_name("URL_NAMES", blk_names("url", 2))
add_name("HIPOBJ_NAMES", blk_names("hip"))
add_name("HIPPROF_NAMES", blk_names("hipp"))
add_name("EDL_NAMES", blk_names("edl"))
ce, pe = CL(OFF["edl"] + 17), CL(OFF["edl"] + 19)
add_name("EDL_CERT_NAMES", f"OFFSET({q(S2)}!${ce}$8,0,0,MAX(1,COUNTA({q(S2)}!${ce}$8:${ce}$5000)),1)")
add_name("EDL_PRE_NAMES", f"OFFSET({q(S2)}!${pe}$8,0,0,MAX(1,COUNTA({q(S2)}!${pe}$8:${pe}$27)),1)")
add_name("OS_VALUES", f"OFFSET({q(R1)}!$B$5,0,0,MAX(1,COUNTA({q(R1)}!$B$5:$B$44)),1)")
add_name("HIP_CERT_TOP", f"{q(R1)}!$D$5")
add_name("HIP_CERT_NAMES", f"OFFSET({q(R1)}!$D$5,0,0,MAX(1,COUNTA({q(R1)}!$D$5:$D$1005)),1)")
add_name("PROFGROUP_NAMES", blk_names("pg", 2))
add_name("PROF_NAMES", blk_names("prof"))

# ======================================================================
# ⑦ カスタムURLカテゴリ（横長：1列1カテゴリ）
# ======================================================================
ws = wb[S7]
NCAT = 50
C1, C2 = 2, 1 + NCAT
U_OP, U_NAME, U_TYPE, U_DESC, U_EX, U_CNT, U_JUDGE, U_REASON, U_VBA, U_LBL, U_FIRST, U_LAST = 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 3015
ws["A1"] = "⑦ カスタムURLカテゴリ（登録・変更・削除）：1列1カテゴリ、URLは1セル1件"
ws["A1"].font = FT_TTL
ws["A2"] = ("B列から右へ1列に1カテゴリずつ記入してください。6行目=操作（登録・変更・削除）、7行目=カテゴリ名（変更・削除は既存の名前をプルダウンから選択）、8行目=type、9行目=description、"
            f"{U_FIRST}行目から下=URL（1セル1件、最大2,500件）。変更は「変更後のURLの全件」を記入します（「既存内容を読込」で現在のURLを入れてから追加・削除すると簡単です）。"
            "URLの書き方：http:// は不要、ホスト名のあとに / を付ける（例 www.example.com/）、ワイルドカードは先頭のみ（例 *.example.com/）。"
            f"同じ列の重複は赤になります。{U_JUDGE}・{U_REASON}行目は関数の判定、{U_VBA}行目は「一括チェック」（VBA）のURL書式チェック結果です。")
ws["A2"].alignment = WRAP
ws.merge_cells("A2:P2")
ws.row_dimensions[2].height = 80
ws["A3"] = "カテゴリ数"
ws["B3"] = f'=COUNTA(B{U_NAME}:{CL(C2)}{U_NAME})'
ws["C3"] = "OK"
ws["D3"] = f'=COUNTIF(B{U_JUDGE}:{CL(C2)}{U_JUDGE},"OK")'
ws["A4"] = "警告"
ws["B4"] = f'=COUNTIF(B{U_JUDGE}:{CL(C2)}{U_JUDGE},"警告")'
ws["C4"] = "NG"
ws["D4"] = f'=COUNTIF(B{U_JUDGE}:{CL(C2)}{U_JUDGE},"NG")'
for a in ("A3", "C3", "A4", "C4"):
    ws[a].font = FT_B
for a in ("B3", "D3", "B4", "D4"):
    ws[a].fill = F_CALC
ws["F3"] = "▶ 既存内容を読込（変更の列）"
ws.merge_cells("F3:G4")
ws["I3"] = "▶ 一括チェック"
ws.merge_cells("I3:J4")
for a in ("F3", "I3"):
    ws[a].fill = F_BTN
    ws[a].alignment = CENTER
    ws[a].font = FT_B
labels = {U_OP: "操作", U_NAME: "カテゴリ名", U_TYPE: "type", U_DESC: "description", U_EX: "既存有無(関数)",
          U_CNT: "URL件数(関数)", U_JUDGE: "判定(関数)", U_REASON: "理由(関数)", U_VBA: "URLチェック(VBA)", U_LBL: "URL（1セル1件）↓"}
for r, t in labels.items():
    c = ws.cell(r, 1)
    c.value = t
    c.font = FT_HEAD
    c.fill = F_HEAD
    c.alignment = CENTER
    c.border = BOX
ws.column_dimensions["A"].width = 20
un, uu = CL(OFF["url"] + 2), CL(OFF["urlv"] + 2)
for ci in range(C1, C2 + 1):
    x = CL(ci)
    ws.column_dimensions[x].width = 30
    for r in (U_OP, U_NAME, U_TYPE, U_DESC):
        ws.cell(r, ci).fill = F_IN
        ws.cell(r, ci).border = BOX
    for r in (U_EX, U_CNT, U_JUDGE, U_REASON, U_VBA):
        ws.cell(r, ci).fill = F_CALC
        ws.cell(r, ci).border = BOX
        ws.cell(r, ci).alignment = WRAP
    ws.cell(U_LBL, ci).fill = F_SUB
    ws.cell(U_EX, ci).value = f'=IF({x}{U_NAME}="","",IF(COUNTIF({q(S2)}!${un}$8:${un}$5000,{x}{U_NAME})>0,"あり","なし"))'
    ws.cell(U_CNT, ci).value = f'=COUNTA({x}{U_FIRST}:{x}{U_LAST})'
    ws.cell(U_JUDGE, ci).value = f'=IF({x}{U_REASON}="","",IF({x}{U_REASON}="OK","OK",LEFT({x}{U_REASON},FIND(":",{x}{U_REASON})-1)))'
    N, OP, TY, CNT, EX, VB = f"{x}{U_NAME}", f"{x}{U_OP}", f"{x}{U_TYPE}", f"{x}{U_CNT}", f"{x}{U_EX}", f"{x}{U_VBA}"
    conds = [
        (f'{N}=""', '"NG:カテゴリ名（7行目）が未入力"'),
        (f'{OP}=""', '"NG:操作未選択"'),
        (f'COUNTIF($B${U_NAME}:${CL(C2)}${U_NAME},{N})>1', '"NG:同じカテゴリ名の列が複数あります"'),
        (f'LENB({N})<>LEN({N})', '"NG:カテゴリ名に全角文字が含まれます"'),
        (f'SUMPRODUCT(--ISERROR(FIND(MID({N}&REPT(" ",63),ROW($1:$63),1)," 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")))>0', '"NG:カテゴリ名に使用不可の文字(英数字 空白 . _ - のみ)"'),
        (f'LEN({N})>63', '"NG:カテゴリ名が63文字超"'),
        (f'AND({OP}="登録",{EX}="あり")', '"NG:既存に同名あり(登録不可)"'),
        (f'AND({OP}<>"登録",{EX}="なし")', '"NG:既存エントリ無"'),
        (f'AND({OP}<>"削除",{TY}="")', '"NG:typeが未選択"'),
        (f'{CNT}>2500', '"NG:URLが2,500件超("&' + CNT + '&"件)"'),
        (f'LEFT({VB},2)="NG"', VB),
        (f'AND({OP}="削除",{CNT}>0)', '"警告:削除するカテゴリにURLが記入されています(削除されます)"'),
        (f'AND({OP}="登録",{CNT}=0)', '"警告:URLが未記入です"'),
        (f'AND({OP}="変更",{CNT}=0)', '"警告:URLが0件です(変更後のURLの全件を記入してください。「既存内容を読込」も使えます)"'),
        (f'LEFT({VB},2)="警告"', VB),
    ]
    f = '"OK"'
    for cnd, msg in reversed(conds):
        f = f"IF({cnd},{msg},{f})"
    anyin = f'COUNTA({x}{U_OP}:{x}{U_DESC},{x}{U_FIRST}:{x}{U_LAST})=0'
    ws.cell(U_REASON, ci).value = f'=IF({anyin},"",{f})'
    # 変更列：既存に無いURL（追加）は緑
    ws.conditional_formatting.add(f"{x}{U_FIRST}:{x}{U_LAST}", FormulaRule(
        formula=[f'AND({x}${U_OP}="変更",{x}{U_FIRST}<>"",COUNTIFS({q(S2)}!${uu}$8:${uu}${B_LAST["urlv"]},{x}${U_NAME},{q(S2)}!${CL(OFF["urlv"] + 3)}$8:${CL(OFF["urlv"] + 3)}${B_LAST["urlv"]},{x}{U_FIRST})=0)'],
        fill=PatternFill("solid", bgColor=OK_FILL, fgColor=OK_FILL)))
    ws.conditional_formatting.add(f"{x}{U_FIRST}:{x}{U_LAST}", Rule(type="duplicateValues", dxf=dxf(NG_FILL, True)))
for val, rgb in (("NG", NG_FILL), ("警告", WARN_FILL), ("OK", OK_FILL)):
    ws.conditional_formatting.add(f"B{U_JUDGE}:{CL(C2)}{U_REASON}", FormulaRule(formula=[f'B${U_JUDGE}="{val}"'], fill=PatternFill("solid", bgColor=rgb, fgColor=rgb)))
dv = DataValidation(type="list", formula1='"登録,変更,削除"', allow_blank=True)
dv.add(f"B{U_OP}:{CL(C2)}{U_OP}")
ws.add_data_validation(dv)
dv = DataValidation(type="list", formula1="URL_NAMES", allow_blank=True, showErrorMessage=False, showInputMessage=True,
                    prompt="変更・削除は既存のカテゴリ名を一覧から選択（登録は新しい名前を直接入力）")
dv.add(f"B{U_NAME}:{CL(C2)}{U_NAME}")
ws.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"URL List,Category Match"', allow_blank=True)
dv.add(f"B{U_TYPE}:{CL(C2)}{U_TYPE}")
ws.add_data_validation(dv)
ws.freeze_panes = f"B{U_FIRST}"
for r in range(U_FIRST, U_FIRST + 200):
    for ci in range(C1, C2 + 1):
        ws.cell(r, ci).number_format = "@"
ws.row_dimensions[U_REASON].height = 45
ws.row_dimensions[U_VBA].height = 30
ws.sheet_properties.tabColor = "FFC000"
if SAMPLE:
    smp = [("変更", "WL_corp-allow", "URL List", "ホワイトリスト", ["*.example.co.jp/", "portal.example.com/", "newapp.example.com/"]),
           ("登録", "cat-partner", "URL List", "取引先", ["partner-a.example/", "*.partner-b.example/", "aa*.bad.example/", "partner-c.example", "partner-a.example/"]),
           ("削除", "cat-saas", "", "", []),
           ("登録", "cat-empty", "URL List", "", [])]
    for i, (op, nm, ty, de, urls) in enumerate(smp):
        ci = C1 + i
        ws.cell(U_OP, ci).value, ws.cell(U_NAME, ci).value, ws.cell(U_TYPE, ci).value, ws.cell(U_DESC, ci).value = op, nm, ty or None, de or None
        for j, u in enumerate(urls):
            ws.cell(U_FIRST + j, ci).value = u


# ======================================================================
# 共通：シンプルな入力シートの枠（⑩⑪）
# ======================================================================
def frame(ws, a1, a2, req, heads, widths, inputs, judge, reason, last_in):
    ws["A1"] = a1
    ws["A1"].font = FT_TTL
    ws.merge_cells(f"A1:{reason}1")
    ws["A2"] = a2
    ws["A2"].alignment = WRAP
    ws.merge_cells(f"A2:{reason}2")
    ws.row_dimensions[2].height = 75
    ws["A3"] = "対象件数"
    ws["B3"] = f'=COUNTA({inputs[0]}8:{inputs[0]}507)'
    ws["C3"] = "OK"
    ws["D3"] = f'=COUNTIF({judge}8:{judge}507,"OK")'
    ws["A4"] = "警告"
    ws["B4"] = f'=COUNTIF({judge}8:{judge}507,"警告")'
    ws["C4"] = "NG"
    ws["D4"] = f'=COUNTIF({judge}8:{judge}507,"NG")'
    for a in ("A3", "C3", "A4", "C4"):
        ws[a].font = FT_B
    for a in ("B3", "D3", "B4", "D4"):
        ws[a].fill = F_CALC
    for i, (h, w) in enumerate(zip(heads, widths)):
        c = ws.cell(7, i + 1)
        c.value = h
        c.font = FT_HEAD
        c.fill = F_HEAD
        c.alignment = CENTER
        c.border = BOX
        ws.column_dimensions[CL(i + 1)].width = w
    for i, t in enumerate(req):
        if t:
            ws.cell(6, i + 2).value = t
            ws.cell(6, i + 2).alignment = CENTER
    ws.row_dimensions[7].height = 36
    for r in range(8, 508):
        ws.cell(r, 1).value = f'=IF({inputs[0]}{r}="","",ROW()-7)'
        for ci in range(2, CI(last_in) + 1):
            ws.cell(r, ci).fill = F_IN
            ws.cell(r, ci).border = BOX
            ws.cell(r, ci).alignment = WRAP
        for ci in range(CI(last_in) + 1, len(heads) + 1):
            ws.cell(r, ci).fill = F_CALC
            ws.cell(r, ci).border = BOX
            ws.cell(r, ci).alignment = WRAP
    judge_cf(ws, f"{judge}8:{reason}507", judge)
    ws.freeze_panes = "D8"
    ws.sheet_properties.tabColor = "FFC000"


def chain(conds, default='"OK"'):
    f = default
    for cnd, msg in reversed(conds):
        f = f"IF({cnd},{msg},{f})"
    return f


NAMECH = '" 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-"'

# ======================================================================
# ⑩ アプリケーション（手動更新）
# ======================================================================
ws = wb[S10]
frame(ws, "⑩ アプリケーション（手動更新）：ポリシーで使うアプリケーションの一覧",
      "SCMから自動では取り出さないため、お客様に一覧を記入・更新していただくシートです。1行に1アプリケーション。操作は「現状のまま」（今使っていて変更なし）「登録」（新しく使う）「変更」「削除」（使わなくなる）から選択。"
      "アプリケーション名はSCM（App-ID）の名前（例 ms-office365、web-browsing）を英数字で記入してください。自社で作るアプリ（カスタム）は区分を「カスタム」にしてください（31文字以内）。"
      "H列（判定）・I列（理由）は関数で自動チェックします（重複・文字種・必須項目）。",
      ["必須", "必須", "必須※", "任意", "任意", "任意"],
      ["No", "操作", "アプリケーション名\n(App-ID)", "区分", "用途・説明", "利用対象\n(部署・ユーザ等)", "備考", "判定", "理由", "一覧内重複"],
      [6, 12, 30, 12, 36, 24, 24, 8, 50, 10], ["C"], "H", "I", "G")
for r in range(8, 508):
    C, B, D, E = f"$C{r}", f"$B{r}", f"$D{r}", f"$E{r}"
    ws.cell(r, 10).value = f'=IF({C}="","",IF(COUNTIF($C$8:$C$507,{C})>1,"重複",""))'
    conds = [
        (f'{C}=""', '"NG:アプリケーション名が未入力"'),
        (f'{B}=""', '"NG:操作未選択"'),
        (f'$J{r}="重複"', '"NG:一覧内でアプリケーション名が重複(大文字・小文字は区別しません)"'),
        (f'LENB({C})<>LEN({C})', '"NG:アプリケーション名に全角文字が含まれます"'),
        (f'ISNUMBER(FIND(" ",{C}))', '"NG:アプリケーション名に空白が含まれます"'),
        (f'SUMPRODUCT(--ISERROR(FIND(MID({C}&REPT("-",63),ROW($1:$63),1),"0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")))>0', '"NG:アプリケーション名に使用不可の文字(英数字 . _ - のみ)"'),
        (f'LEN({C})>63', '"NG:アプリケーション名が63文字超"'),
        (f'AND({B}<>"削除",{D}="")', '"NG:区分(定義済み/カスタム)が未選択"'),
        (f'AND({D}="カスタム",LEN({C})>31)', '"NG:カスタムアプリケーションの名前は31文字以内"'),
        (f'AND(OR({B}="登録",{B}="変更"),{E}="")', '"警告:用途・説明の記入をお願いします"'),
    ]
    ws.cell(r, 9).value = f'=IF(COUNTA($B{r}:$G{r})=0,"",{chain(conds)})'
    ws.cell(r, 8).value = f'=IF($I{r}="","",IF($I{r}="OK","OK",LEFT($I{r},FIND(":",$I{r})-1)))'
dv = DataValidation(type="list", formula1='"現状のまま,登録,変更,削除"', allow_blank=True)
dv.add("B8:B507")
ws.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"定義済み,カスタム"', allow_blank=True)
dv.add("D8:D507")
ws.add_data_validation(dv)
ws.column_dimensions.group("J", "J", hidden=True, outline_level=1)
if SAMPLE:
    for i, row in enumerate([
        ("現状のまま", "ms-office365", "定義済み", "Microsoft 365 全般", "全社", None),
        ("登録", "zoom", "定義済み", "Web会議", "営業部", None),
        ("登録", "Zoom", "定義済み", "", "", "大文字違いの重複"),
        ("削除", "bittorrent", "定義済み", None, None, None),
        ("登録", "custom-erp-application-name-toolong", "カスタム", "基幹系", "経理部", None),
        ("変更", "web browsing", "定義済み", "閲覧", None, None),
    ]):
        for j, v in enumerate(row):
            ws.cell(8 + i, 2 + j).value = v

# ======================================================================
# ⑪ ユーザ/グループ（手動更新）
# ======================================================================
ws = wb[S11]
frame(ws, "⑪ ユーザ・グループ（手動更新）：ポリシーで使うユーザ・グループの一覧",
      "SCMから自動では取り出さないため、お客様に一覧を記入・更新していただくシートです。1行に1ユーザまたは1グループ。操作は「現状のまま」「登録」「変更」「削除」から選択。"
      "ユーザの名前(ID)は user@domain（UPN／メールアドレス）形式で記入してください。グループはディレクトリ（Entra ID／AD 等）のグループ名を記入します。"
      "ユーザ行の所属グループ（G列）はダブルクリックすると、このシートのグループ行からチェックボックスで選べます（Ctrl+Shift+Mでも可。手入力は「,」区切りかセル内改行）。"
      "J列（判定）・K列（理由）は関数で自動チェックします。",
      ["必須", "必須", "必須", "任意", "任意", "任意", "任意", "任意"],
      ["No", "操作", "種別", "名前(ID)\n(ユーザ=user@domain)", "表示名", "ドメイン", "所属グループ\n(ユーザのみ・ダブルクリックで選択)",
       "用途・説明", "備考", "判定", "理由", "一覧内重複", "作業:所属(正規化)", "作業:所属数", "作業:一致数", "作業:削除予定一致", "作業:参照ユーザ数"],
      [6, 12, 10, 32, 20, 18, 32, 30, 20, 8, 50, 10, 20, 8, 8, 8, 8], ["D"], "J", "K", "I")
for r in range(8, 508):
    B, C, D, G, M, N, O, P, Qc = (f"$B{r}", f"$C{r}", f"$D{r}", f"$G{r}", f"$M{r}", f"$N{r}", f"$O{r}", f"$P{r}", f"$Q{r}")
    ws.cell(r, 12).value = f'=IF({D}="","",IF(COUNTIFS($C$8:$C$507,{C},$D$8:$D$507,{D})>1,"重複",""))'
    ws.cell(r, 13).value = f'=IF({G}="","",UPPER(SUBSTITUTE(SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE({G},CHAR(13),""),CHAR(10),","))," ,",","),", ",",")))'
    ws.cell(r, 14).value = f'=IF({M}="",0,LEN({M})-LEN(SUBSTITUTE({M},",",""))+1)'
    ws.cell(r, 15).value = (f'=IF({M}="",0,SUMPRODUCT(--($C$8:$C$507="グループ"),--($B$8:$B$507<>"削除"),--($D$8:$D$507<>""),'
                            f'--ISNUMBER(FIND(","&UPPER($D$8:$D$507)&",",","&{M}&","))))')
    ws.cell(r, 16).value = (f'=IF({M}="",0,SUMPRODUCT(--($C$8:$C$507="グループ"),--($B$8:$B$507="削除"),--($D$8:$D$507<>""),'
                            f'--ISNUMBER(FIND(","&UPPER($D$8:$D$507)&",",","&{M}&","))))')
    ws.cell(r, 17).value = (f'=IF(OR({C}<>"グループ",{D}=""),0,SUMPRODUCT(--($C$8:$C$507="ユーザ"),--($B$8:$B$507<>"削除"),'
                            f'--ISNUMBER(FIND(","&UPPER({D})&",",","&$M$8:$M$507&","))))')
    conds = [
        (f'{D}=""', '"NG:名前(ID)が未入力"'),
        (f'{B}=""', '"NG:操作未選択"'),
        (f'{C}=""', '"NG:種別(ユーザ/グループ)が未選択"'),
        (f'$L{r}="重複"', '"NG:同じ種別・名前の行が重複(大文字・小文字は区別しません)"'),
        (f'AND({C}="ユーザ",ISNUMBER(FIND(" ",TRIM({D}))))', '"NG:ユーザの名前(ID)に空白が含まれます"'),
        (f'AND({C}="ユーザ",LENB({D})<>LEN({D}))', '"NG:ユーザの名前(ID)に全角文字が含まれます"'),
        (f'AND({B}<>"削除",{P}>0)', '"NG:削除するグループを所属グループに指定しています"'),
        (f'AND({C}="ユーザ",ISERROR(FIND("@",{D})))', '"警告:ユーザの名前(ID)は user@domain（UPN／メール）形式で記入してください"'),
        (f'AND({C}="グループ",{G}<>"")', '"警告:グループ行の所属グループは使いません(ユーザ行に記入してください)"'),
        (f'AND({B}<>"削除",{C}="ユーザ",{N}>{O})', '"警告:このシートに無いグループが所属グループに含まれます("&(' + N + '-' + O + ')&"件)。グループの行も追加してください"'),
        (f'AND({B}="削除",{C}="グループ",{Qc}>0)', '"警告:このグループを所属グループにしているユーザ("&' + Qc + '&"件)があります"'),
    ]
    ws.cell(r, 11).value = f'=IF(COUNTA($B{r}:$I{r})=0,"",{chain(conds)})'
    ws.cell(r, 10).value = f'=IF($K{r}="","",IF($K{r}="OK","OK",LEFT($K{r},FIND(":",$K{r})-1)))'
dv = DataValidation(type="list", formula1='"現状のまま,登録,変更,削除"', allow_blank=True)
dv.add("B8:B507")
ws.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"ユーザ,グループ"', allow_blank=True)
dv.add("C8:C507")
ws.add_data_validation(dv)
ws.column_dimensions.group("L", "Q", hidden=True, outline_level=1)
if SAMPLE:
    for i, row in enumerate([
        ("現状のまま", "グループ", "GG-Sales", "営業部", "example.co.jp", None, "営業部全員"),
        ("登録", "グループ", "GG-Dev", "開発部", "example.co.jp", None, "開発部"),
        ("削除", "グループ", "GG-Old", "旧組織", "example.co.jp", None, None),
        ("登録", "ユーザ", "taro.yamada@example.co.jp", "山田 太郎", "example.co.jp", "GG-Sales,GG-Dev", "管理者"),
        ("登録", "ユーザ", "hanako", "佐藤 花子", None, "GG-Sales\nGG-Marketing", None),
        ("変更", "ユーザ", "jiro@example.co.jp", "鈴木 次郎", None, "GG-Old", None),
    ]):
        for j, v in enumerate(row):
            ws.cell(8 + i, 2 + j).value = v

# ======================================================================
# ②-2 / ②-3 既存ポリシー、⑫ / ⑬ 入力、参照_ポリシー選択肢
#   既存ツール（scm-policy / scm-decryption）の ②取得・③設定・⑧選択肢 と同じ列並び
# ======================================================================
P_DEPTH = {"All": 0, "Shared": 1, "Remote Networks": 2, "Service Connections": 2, "Mobile Users Container": 2,
           "Mobile Users": 3, "Mobile Users Explicit Proxy": 3}
P_DCOL = {0: "404040", 1: "1F3864", 2: "2F5A97", 3: "4472C4"}
P_FLOW = [("Shared", "pre"), ("Mobile Users Container", "pre"), ("Mobile Users", "pre"), ("Mobile Users Explicit Proxy", "pre"),
          ("Mobile Users Container", "post"), ("Remote Networks", "pre"), ("Shared", "post")]
P_LBL = {"Shared": "Prisma Access（Shared）", "Mobile Users": "GlobalProtect（Mobile Users）",
         "Mobile Users Explicit Proxy": "Explicit Proxy（Mobile Users Explicit Proxy）"}


# 復号ポリシーもセキュリティポリシーと同じ評価順。Service Connections（Pre/Post なし）は復号ポリシーだけ
P_FLOW_DEC = P_FLOW[:6] + [("Service Connections", "pre")] + P_FLOW[6:]


def p_glabel(f, p, key="pol"):
    lab = P_LBL.get(f, f)
    if f in ("Mobile Users", "Mobile Users Explicit Proxy", "Remote Networks", "Service Connections"):
        return lab
    return f"{lab} {'【Post】' if p == 'post' else '【Pre】'}"


def copy_whole(sw, ws, names, max_row, cols=None):
    for row in sw.iter_rows(min_row=1, max_row=max_row):
        for c in row:
            if cols and c.row >= 8 and c.column not in cols:
                continue
            d = ws.cell(c.row, c.column)
            copy_style(c, d)
            d.value = remap(c.value, 0, names) if isinstance(c.value, str) and c.value.startswith("=") else c.value
    for m in sw.merged_cells.ranges:
        if m.min_row <= 7:
            ws.merge_cells(str(m))
    for k, dim in sw.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
        if dim.hidden:
            ws.column_dimensions[k].hidden = True
    for r, dim in sw.row_dimensions.items():
        if dim.height and r <= max_row:
            ws.row_dimensions[r].height = dim.height
    ws.freeze_panes = sw.freeze_panes
    copy_dv(sw, ws, 0, names)


SINGLE_RB = ("Mobile Users", "Mobile Users Explicit Proxy", "Remote Networks", "Service Connections")


def reorder_policy(ws, key, is_set):
    """サンプルのポリシー表を評価順（P_FLOW / P_FLOW_DEC）に並べ替える。
    Pre/Post なしの folder の post の行は pre（1つのルールベース）にまとめ、並びに無い空の見出しは消す"""
    nf = 18 if key == "pol" else 21
    hf = (13 + nf) if is_set else (11 + nf)
    scope_c = 5 if is_set else 2
    pos_c = 9 if is_set else 6
    lastc = hf + 1
    order = P_FLOW if key == "pol" else P_FLOW_DEC
    cur, groups = None, {}
    max_r = ws.max_row
    for r in range(8, max_r + 1):
        snap = [(ws.cell(r, c).value, copy(ws.cell(r, c)._style)) for c in range(1, lastc + 1)]
        if ws.cell(r, 1).value == "▼":
            f, pp = ws.cell(r, hf).value, ws.cell(r, hf + 1).value
            if f in SINGLE_RB:
                pp = "pre"
            cur = (f, pp)
            groups.setdefault(cur, {"head": snap, "rules": []})
        elif cur and any(v is not None for v, _ in snap):
            groups[cur]["rules"].append(snap)
    for k in groups:
        if k not in order and groups[k]["rules"]:
            raise SystemExit(f"sample rules in unsupported scope {k}")
    blank = copy(ws.cell(max_r + 50, 1)._style)
    for r in range(8, max_r + 1):
        for c in range(1, lastc + 1):
            ws.cell(r, c).value = None
            ws.cell(r, c)._style = copy(blank)
    r = 8
    for k in order:
        f, pp = k
        g = groups.get(k)
        if g is None:
            continue
        for c, (v, st) in enumerate(g["head"], 1):
            ws.cell(r, c).value = v
            ws.cell(r, c)._style = st
        for c in range(scope_c, scope_c + 3):
            ws.cell(r, c).value = None
        lv = max(P_DEPTH[f] - 1, 0)
        ws.cell(r, scope_c + lv).value = f"■ {p_glabel(f, pp, key)} ／ {len(g['rules'])} 件"
        ws.cell(r, scope_c + lv)._style = copy(ws.cell(r, 1)._style)
        ws.cell(r, hf + 1).value = pp
        r += 1
        for i, rule in enumerate(g["rules"], 1):
            for c, (v, st) in enumerate(rule, 1):
                ws.cell(r, c).value = v
                ws.cell(r, c)._style = st
            ws.cell(r, 1).value = i
            ws.cell(r, pos_c).value = pp
            r += 1
    return r - 1


POLS = [
    # srcKey, 取得シート, 設定シート, 新 取得, 新 入力, NF, 表示名, ツール名
    ("pol", "②SecRule_取得", "③SecRule_設定", S2B, S12, 18, "セキュリティポリシー", "SCM_SecurityPolicy.xlsx"),
    ("dec", "②DecRule_取得", "③DecRule_設定", S2C, S13, 21, "復号ポリシー", "SCM_DecryptionPolicy.xlsx"),
]
for key, gsh, ssh, ng, ns, nf, lbl, tool in POLS:
    names = NAME_MAP[key]
    s_last, s_hf = 11 + nf, 13 + nf
    # ---- 既存（取得）
    sw = src_wb[key][gsh]
    ws = wb[ng]
    copy_whole(sw, ws, names, sw.max_row if SAMPLE else 7)
    if SAMPLE:
        reorder_policy(ws, key, False)
    ws["A1"] = f"{ng[:3]} 既存の{lbl}（現在SCMに設定されている内容）"
    ws["A2"] = (f"担当者が①基本設定の「既存情報をまとめて取得（API）」で取得します（既存ツール（{tool}）で取得した{lbl}を「②既存データ抽出」の「既存ツールから取込」で取り込むこともできます）。"
                f"スコープ（Prisma Access → Mobile Users Container → GlobalProtect・Explicit Proxy → Remote Networks）ごとに▼の見出し行があり、その下が評価順のルールです。"
                f"「{ns}」の展開と入力チェックの基準です。お客様は参照のみで、編集は不要です。")
    ws["D3"] = "（①の「既存情報をまとめて取得（API）」または②の「既存ツールから取込」で入ります）"
    ws["H3"] = None
    ws["H3"].fill = PatternFill()
    ws["A4"] = None
    ws["B4"] = None
    ws.sheet_properties.tabColor = "70AD47"
    # ---- 入力（設定）
    sw = src_wb[key][ssh]
    ws = wb[ns]
    copy_whole(sw, ws, names, sw.max_row if SAMPLE else 7)
    if SAMPLE:
        reorder_policy(ws, key, True)
    ws["A1"] = f"{ns[0]} {lbl}（追加・変更・削除・移動）"
    ws["A2"] = (f"「既存ポリシーを展開」で②-{'2' if key == 'pol' else '3'}の表をこのシートに作り、黄色セルに記入します。変更＝操作「変更」を選んでセルを書き換え／削除＝操作「削除」／"
                "追加＝追加したい位置の1つ上の行（スコープの先頭なら ▼の行）を選んで「選択行の下に追加行を挿入」→ name 等を記入／並べ替え＝行を切り取り→挿入して操作「移動」（内容も変えるなら「変更+移動」）。"
                "ゾーン・アドレス・ユーザー・HIP・" + ("アプリケーション・" if key == "pol" else "") + "サービス・URLカテゴリ・タグはダブルクリック（または Ctrl+Shift+M）で、"
                "そのスコープから見えるものだけをチェックボックスで選べます。「入力チェック」（記入状況の「一括チェック」でも可）で C・D列に判定と理由を出します："
                "③〜⑪で削除予定のものを使っている／ルールのスコープから見えないフォルダで定義されたものを使っている＝NG、②既存データにも③〜⑪の登録にも無い＝警告。"
                "既存から変わったセルは橙色です。グローバル（All）は参照のみです。" + ("Service Connections には設定できません。" if key == "pol" else "GlobalProtect・Explicit Proxy・Remote Networks・Service Connections は Pre/Post なしです（並び順はセキュリティポリシーと同じ）。"))
    for r in range(8, ws.max_row + 1):
        if ws.cell(r, 1).value != "▼":
            ws.cell(r, 3).value = None
            ws.cell(r, 4).value = None
            ws.cell(r, 3).fill = PatternFill("solid", fgColor="F2F2F2")
    ws["F3"] = "▶ 既存ポリシーを展開（作り直し）"
    ws["P3"] = None
    ws["P3"].fill = PatternFill()
    ws.row_dimensions[2].height = 96
    if not SAMPLE:
        for i, (f, p) in enumerate(P_FLOW if key == "pol" else P_FLOW_DEC):
            r = 8 + i
            lv = max(P_DEPTH[f] - 1, 0)
            ws.cell(r, 1).value = "▼"
            ws.cell(r, 5 + lv).value = f"■ {p_glabel(f, p, key)} ／ 0 件"
            ws.cell(r, s_hf).value = f
            ws.cell(r, s_hf + 1).value = p
            for c in range(1, s_hf + 2):
                cell = ws.cell(r, c)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.number_format = "@"
                if c <= s_last:
                    cell.fill = PatternFill("solid", fgColor=P_DCOL[P_DEPTH[f]])
    ws.sheet_properties.tabColor = "FFC000"

# ---- 参照_ポリシー選択肢
sw = src_wb["pol"]["⑧SecRule_選択肢"]
ws = wb[R3]
keep_cols = {1, 2, 3, 11, 18, 19, 20, 23, 24, 25, 26, 27, 28, 29, 30} if SAMPLE else {1, 18, 23, 24}
copy_whole(sw, ws, {}, sw.max_row, keep_cols)
if not SAMPLE:
    for r in range(8, sw.max_row + 1):
        if ws.cell(r, 23).value:
            ws.cell(r, 23).value = None
for r in range(7, 12):
    ws.cell(r, 32).value = None
for r in range(8, 308):
    for c in (4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 21, 22):
        ws.cell(r, c).value = None
ws.unmerge_cells("A1:AD1")
ws.unmerge_cells("A2:AD2")
ws.merge_cells("A1:AI1")
ws.merge_cells("A2:AI2")
ws["A1"] = "参照_ポリシー選択肢（⑫⑬のダブルクリック選択の候補・入力チェックの基準）"
ws["A2"] = ("アドレス・アドレスグループ・EDL・サービス・サービスグループ・カスタムURLカテゴリ・HIPプロファイル・プロファイルグループの列は、"
            "「選択肢を作り直す」・⑫⑬の入力チェック・ダブルクリックのたびに、②既存データ（スコープ＝コンテナの folder）と③〜⑨の「登録」行（スコープ＝①基本設定のオブジェクト取得folder）から自動で作り直します。"
            "ゾーン・タグ・リージョン・事前定義URLカテゴリ・アプリグループ・アプリフィルタ・自作App・証明書・復号プロファイルは、①基本設定の「既存情報をまとめて取得（API）」か、既存ツール（SCM_SecurityPolicy.xlsx／SCM_DecryptionPolicy.xlsx）からの取込で入ります（ゾーンは取得したルールで使われているものを追加。ゾーンとスコープは手で追記できます）。"
            "ユーザー（⑪）とアプリケーション（⑩）は記入内容を追記します。スコープ列＝定義されている folder（空欄＝どこでも使える）。")
ws.row_dimensions[2].height = 84
ws["D3"] = "▶ 選択肢を作り直す（②と③～⑪から）"
hs, hsc = sw.cell(7, 4), sw.cell(7, 5)
for c, t, src in ((31, "証明書\n(復号 ssl_inbound_inspection)", hs), (32, "復号プロファイル", hs), (34, "プロファイルグループ", hs), (35, "スコープ", hsc)):
    d = ws.cell(7, c)
    copy_style(src, d)
    d.value = t
    ws.column_dimensions[CL(c)].width = 13 if t == "スコープ" else 24
    lab = ws.cell(6, c)
    copy_style(sw.cell(6, 4), lab)
    lab.value = "自動" if c >= 34 else "取込"
ws.column_dimensions["AG"].width = 3
for c in (4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 21, 22):
    copy_style(sw.cell(6, 4), ws.cell(6, c))
    ws.cell(6, c).value = "自動"
for c in (11, 18, 19, 20, 25, 26, 27, 28, 29, 30):
    copy_style(sw.cell(6, 4), ws.cell(6, c))
    ws.cell(6, c).value = "取込"
notes = ["※ スコープ＝定義されている folder。ルールからは「自分の folder と上位の folder」（GlobalProtect・Explicit Proxy なら Mobile Users Container・Shared、Remote Networks・Mobile Users Container なら Shared）の定義だけが見えます。",
         "※ ゾーンのスコープは手入力（空欄＝どのスコープでも使える。カンマ区切りで複数可）。",
         "※ 事前定義URLカテゴリ・リージョン・定義済みApp-ID・ユーザーはどのスコープでも使えます。",
         "※ EDL は種類が ip のものをアドレス、url のものをURLカテゴリの候補にします。",
         "※ ③〜⑪で「削除」の行にしたものは候補に出さず、ポリシーで使っていると NG にします。"]
for i, t in enumerate(notes):
    ws.cell(7 + i, 37).value = t
    ws.cell(7 + i, 37).font = Font(size=9, color="595959")
if SAMPLE:
    dsw = src_wb["dec"]["⑧DecRule_選択肢"]
    for r in range(8, dsw.max_row + 1):
        if dsw.cell(r, 14).value:
            ws.cell(r, 31).value = dsw.cell(r, 14).value
    psw = src_wb["dec"]["⑦DecProfile_一覧"]
    for r in range(8, psw.max_row + 1):
        if psw.cell(r, 2).value:
            ws.cell(r, 32).value = psw.cell(r, 2).value
for r in range(8, 308):
    for c in (2, 3):
        ws.cell(r, c).fill = F_IN
ws.sheet_properties.tabColor = "A6A6A6"
add_name("DEC_CERT_NAMES", f"OFFSET({q(R3)}!$AE$8,0,0,MAX(1,COUNTA({q(R3)}!$AE$8:$AE$10007)),1)")
add_name("DEC_PROF_NAMES", f"OFFSET({q(R3)}!$AF$8,0,0,MAX(1,COUNTA({q(R3)}!$AF$8:$AF$10007)),1)")

# ======================================================================
# ① 基本設定（既存ツール SCM_SecurityPolicy.xlsx の ①基本設定 と同じ。担当者がトークン取得・API取得に使う）
# ======================================================================
sw = src_wb["pol"]["①基本設定"]
ws = wb[S1]
for row in sw.iter_rows(min_row=1, max_row=sw.max_row):
    for c in row:
        d = ws.cell(c.row, c.column)
        copy_style(c, d)
        d.value = c.value
for m in sw.merged_cells.ranges:
    ws.merge_cells(str(m))
for k, dim in sw.column_dimensions.items():
    ws.column_dimensions[k].width = dim.width
for r, dim in sw.row_dimensions.items():
    if dim.height:
        ws.row_dimensions[r].height = dim.height
copy_dv(sw, ws, 0, {})
for n, dn in src_wb["pol"].defined_names.items():
    if n.startswith("CFG_"):
        add_name(n, dn.attr_text)
ws["A1"] = "① 基本設定（API接続先・認証情報・アクセストークン）※担当者用"
ws["A2"] = ("担当者がトークンを取得し、「既存情報をまとめて取得（API）」で②・②-2・②-3・参照_ポリシー選択肢を作り直します（GET のみ。SCM の設定は変えません）。"
            "client_secret を含むため、お客様にお渡しする前に「接続情報を消去」を押してください。お客様の記入は不要です。")
ws.row_dimensions[2].height = 45
ws["B7"] = "/sse/config/v1"
ws["D7"] = "このブックでは種類ごとのエンドポイントを自動で使います（変更不要）"
for a in ("B10", "B11", "B12"):
    ws[a] = None
ws["D10"] = "担当者が入力"
ws["D11"] = "担当者が入力（お客様に渡す前に「接続情報を消去」）"
ws["D12"] = "テナントサービスグループID（担当者が入力）"
ws["A14"] = "取得の条件"
ws["D16"] = "アドレス・サービス・HIP・EDL・プロファイル等を取得する基本の folder。このほか下の○のスコープからも取得し、定義folderをスコープとして記録。③～⑪で登録するものはこの folder に作る前提でチェック"
ws["A28"] = "※ このブックは取得（GET）だけを行います。登録・変更・削除・設定反映（Push）は、SCM の画面または各ツールで実施してください。"
for r, lab, btn, txt in ((30, "既存情報の取得", "▶ 既存情報をまとめて取得（API）",
                          "②既存データ抽出（アドレス・グループ・サービス・URLカテゴリ・HIP・EDL・プロファイルグループ）、②-2・②-3 既存ポリシー、参照_ポリシー選択肢 を SCM の現在の内容で作り直します。"),
                         (32, "接続情報の消去", "▶ 接続情報を消去（お客様に渡す前に）",
                          "client_secret とアクセストークンを消します（client_id・TSG ID も消すか選べます）。消してから保存し、お客様にお渡しください。")):
    ws.cell(r, 1).value = lab
    ws.cell(r, 1).font = FT_B
    ws.cell(r, 2).value = txt
    ws.cell(r, 2).alignment = WRAP
    ws.cell(r, 4).value = btn
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    copy_style(sw["D18"], ws.cell(r, 4))
    ws.row_dimensions[r].height = 45
ws["B40"] = "○"
ws["D40"] = "オブジェクト・復号ポリシーは取得します（セキュリティポリシーは Service Connections に設定できないため取得しません）"
ws["A36"] = "取得の対象スコープ：○を付けたものを上から順に取得し、スコープごとに表化（ポリシーの1行目にスコープ名を表示）"
ws.freeze_panes = sw.freeze_panes
ws.sheet_properties.tabColor = "C00000"
if SAMPLE:
    ws["B10"] = "dummy-sa@1234567890.iam.panserviceaccount.com"
    ws["B12"] = "1234567890"

# ---------------------------------------------------------------- 実行ログ（既存ツールの ⑥実行ログ と同じ）
sw = src_wb["pol"]["⑥実行ログ"]
ws = wb[SLOG]
for row in sw.iter_rows(min_row=1, max_row=5):
    for c in row:
        d = ws.cell(c.row, c.column)
        copy_style(c, d)
        d.value = c.value
for m in sw.merged_cells.ranges:
    if m.min_row <= 5 and str(m) != "F3":
        ws.merge_cells(str(m))
for k, dim in sw.column_dimensions.items():
    ws.column_dimensions[k].width = dim.width
for r, dim in sw.row_dimensions.items():
    if dim.height and r <= 5:
        ws.row_dimensions[r].height = dim.height
ws["A1"] = "実行ログ（API取得の結果）"
ws["A2"] = "トークン取得・既存情報の取得（GET）の結果をVBAが追記します。HTTPステータス 200=成功、それ以外=失敗（レスポンス本文を記録）。"
ws.freeze_panes = sw.freeze_panes
ws.sheet_properties.tabColor = "A6A6A6"

# ======================================================================
# 記入状況（各シートの記入件数・判定の集計と一括チェック）
# ======================================================================
ws = wb[SS]
ws["A1"] = "SCM（Prisma Access）設定 入力シート（お客様記入用）— 記入状況"
ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
ws.merge_cells("A1:H1")
ws["A2"] = ("このブックは、Strata Cloud Manager（Prisma Access）へ登録・変更・削除する設定内容をご記入いただくためのものです。"
            "黄色のセルにご記入ください。各シートの「判定」列が NG（赤）の行は修正をお願いします。警告（黄）は内容のご確認をお願いします。")
ws["A2"].alignment = WRAP
ws.merge_cells("A2:H2")
ws.row_dimensions[2].height = 45
ws.column_dimensions["A"].width = 6
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 44
for c, w in zip("DEFGH", (12, 10, 10, 10, 30)):
    ws.column_dimensions[c].width = w

ws["A4"] = "1. 記入するシートと記入状況（判定は関数で自動集計）"
ws["A4"].font = FT_B
hdr = ["No", "シート", "記入内容", "記入件数", "OK", "警告", "NG", "状態"]
for i, h in enumerate(hdr):
    c = ws.cell(5, 1 + i)
    c.value = h
    c.font = FT_HEAD
    c.fill = F_HEAD
    c.alignment = CENTER
    c.border = BOX
SUMS = [
    (S2, "現在SCMに登録されている設定（セキュリティプロファイルを含む。ポリシーは②-2・②-3。担当者がAPIで取得。参照のみ）", None, None, None),
    (S3, "アドレスオブジェクトの登録・変更・削除", "C8:C507", "P8:P507", None),
    (S4, "アドレスグループの登録・変更・削除（メンバーはダブルクリックで選択）", "C8:C507", "P8:P507", None),
    (S5, "サービスの登録・変更・削除", "C8:C507", "X8:X507", None),
    (S6, "サービスグループの登録・変更・削除（メンバーはダブルクリックで選択）", "C8:C507", "L8:L507", None),
    (S7, "カスタムURLカテゴリの登録・変更・削除（1列1カテゴリ）", f"B{U_NAME}:{CL(C2)}{U_NAME}", f"B{U_JUDGE}:{CL(C2)}{U_JUDGE}", None),
    (S8, "HIPオブジェクトの登録・変更・削除", "C8:C507", "D8:D507", None),
    (S8B, "HIPプロファイルの登録・変更・削除（HIPオブジェクトはダブルクリックで選択）", "C8:C507", "P8:P507", None),
    (S9, "EDL（外部動的リスト）の登録・変更・削除", "C8:C507", "AE8:AE507", None),
    (S10, "ポリシーで使うアプリケーションの一覧（手動更新）", "C8:C507", "H8:H507", None),
    (S11, "ポリシーで使うユーザ・グループの一覧（手動更新）", "D8:D507", "J8:J507", None),
    (S12, "セキュリティポリシーの追加・変更・削除・移動（スコープごと。判定は「入力チェック」か「一括チェック」）", "B8:B3000", "C8:C3000", None),
    (S13, "復号ポリシーの追加・変更・削除・移動（スコープごと。判定は「入力チェック」か「一括チェック」）", "B8:B3000", "C8:C3000", None),
]
for i, (sh, desc, cnt, jd, _x) in enumerate(SUMS):
    r = 6 + i
    ws.cell(r, 1).value = "⑧-2" if sh == S8B else sh[0]
    c = ws.cell(r, 2)
    c.value = sh
    c.hyperlink = f"#{q(sh)}!A1"
    c.font = Font(color="0563C1", underline="single")
    ws.cell(r, 3).value = desc
    ws.cell(r, 3).alignment = WRAP
    if cnt:
        ws.cell(r, 4).value = f"=COUNTA({q(sh)}!{cnt})"
        ws.cell(r, 5).value = f'=COUNTIF({q(sh)}!{jd},"OK")'
        ws.cell(r, 6).value = f'=COUNTIF({q(sh)}!{jd},"警告")'
        ws.cell(r, 7).value = f'=COUNTIF({q(sh)}!{jd},"NG")'
        if sh in (S12, S13):
            ws.cell(r, 8).value = f'=IF(D{r}=0,"記入なし",IF(E{r}+F{r}+G{r}<D{r},"未チェック（入力チェックを押してください）",IF(G{r}>0,"要修正（NGあり）",IF(F{r}>0,"要確認（警告あり）","OK"))))'
        else:
            ws.cell(r, 8).value = f'=IF(D{r}=0,"記入なし",IF(G{r}>0,"要修正（NGあり）",IF(F{r}>0,"要確認（警告あり）","OK")))'
    else:
        ws.cell(r, 4).value = f'=IF({q(S2)}!B3="","",{q(S2)}!B3)'
        ws.cell(r, 4).number_format = "yyyy/mm/dd"
        ws.cell(r, 8).value = "取得日時（D列）"
    for ci in range(1, 9):
        ws.cell(r, ci).border = BOX
        if ci >= 4:
            ws.cell(r, ci).fill = F_CALC
    ws.cell(r, 1).alignment = CENTER
rT = 6 + len(SUMS)
ws.cell(rT, 3).value = "合計"
ws.cell(rT, 3).font = FT_B
for ci, L in ((4, "D"), (5, "E"), (6, "F"), (7, "G")):
    ws.cell(rT, ci).value = f"=SUM({L}7:{L}{rT - 1})"
    ws.cell(rT, ci).font = FT_B
    ws.cell(rT, ci).border = BOX
ws.conditional_formatting.add(f"H6:H{rT - 1}", FormulaRule(formula=['LEFT($H6,3)="要修正"'], fill=PatternFill("solid", bgColor=NG_FILL, fgColor=NG_FILL)))
ws.conditional_formatting.add(f"H6:H{rT - 1}", FormulaRule(formula=['LEFT($H6,3)="要確認"'], fill=PatternFill("solid", bgColor=WARN_FILL, fgColor=WARN_FILL)))
ws.conditional_formatting.add(f"H6:H{rT - 1}", FormulaRule(formula=['LEFT($H6,5)="未チェック"'], fill=PatternFill("solid", bgColor=WARN_FILL, fgColor=WARN_FILL)))
ws.conditional_formatting.add(f"H6:H{rT - 1}", FormulaRule(formula=['$H6="OK"'], fill=PatternFill("solid", bgColor=OK_FILL, fgColor=OK_FILL)))
assert rT + 1 == 20, rT
ws.cell(rT + 1, 2).value = "▶ 一括チェック"
ws.cell(rT + 1, 2).fill = F_BTN
ws.cell(rT + 1, 2).font = FT_B
ws.cell(rT + 1, 2).alignment = CENTER
ws.cell(rT + 1, 3).value = "（マクロ）関数では見られない項目（URLの書式、メンバーの重複、EDL例外リスト、HIP直接記述の参照先、⑫⑬ポリシーのスコープと参照先）も確認し、結果を「チェック結果」シートに一覧にします。"
ws.cell(rT + 1, 3).alignment = WRAP
ws.merge_cells(start_row=rT + 1, start_column=3, end_row=rT + 1, end_column=8)
ws.row_dimensions[rT + 1].height = 32

r0 = rT + 3
ws.cell(r0, 1).value = "2. 記入のしかた"
ws.cell(r0, 1).font = FT_B
howto = [
    "黄色のセルに記入します。灰色のセルは関数（自動）です。操作・種別などはプルダウンから選びます。",
    "各シートの「操作」は 登録（新規）／変更（既存の内容を変える）／削除 から選びます。変更・削除の名前は、②既存データにある名前をプルダウンから選べます。",
    "変更は「変更後の全項目」を記入します（空欄にした項目は消えます）。「変更行に既存値を読込」ボタンで現在の値を入れてから直すと簡単です。",
    "グループのメンバー欄などは、セルをダブルクリックするとチェックボックスの一覧から選べます（Ctrl+Shift+M でも開きます）。",
    "判定：OK（緑）＝問題なし、警告（黄）＝内容をご確認ください、NG（赤）＝修正が必要です。理由の列に内容が出ます。変更の行で既存と異なる項目は橙色になります。",
    "名前（name）は英数字・空白・ . _ - のみ使えます。大文字・小文字は区別されます。",
    "セキュリティポリシー（⑫）・復号ポリシー（⑬）はスコープ（Prisma Access／Mobile Users Container／GlobalProtect／Explicit Proxy／Remote Networks）ごとの表です。使えるオブジェクトは、そのスコープと上位（GlobalProtect・Explicit Proxy なら Mobile Users Container と Shared、Remote Networks・Mobile Users Container なら Shared）で定義されたものだけです。",
    "①基本設定・②既存データ・②-2・②-3・実行ログは担当者用です（SCM から取得した現在の設定。お客様は記入不要です）。",
    "パスワードなどの秘密情報はこのブックに書かず、別途お知らせください。",
]
for i, t in enumerate(howto):
    c = ws.cell(r0 + 1 + i, 2)
    c.value = f"・{t}"
    c.alignment = WRAP
    ws.merge_cells(start_row=r0 + 1 + i, start_column=2, end_row=r0 + 1 + i, end_column=8)
    ws.row_dimensions[r0 + 1 + i].height = 30
r1 = r0 + 2 + len(howto)
ws.cell(r1, 1).value = "3. 色の見方"
ws.cell(r1, 1).font = FT_B
legend = [("記入欄", "FFF2CC"), ("関数（自動）", "F2F2F2"), ("OK", OK_FILL), ("警告", WARN_FILL), ("NG", NG_FILL), ("変更で既存と異なる項目", CHG_FILL), ("ボタンの位置", "F4B084")]
for i, (t, rgb) in enumerate(legend):
    c = ws.cell(r1 + 1 + i, 2)
    c.value = t
    c.fill = PatternFill("solid", fgColor=rgb)
    c.border = BOX
ws.sheet_properties.tabColor = "1F4E78"

# ---------------------------------------------------------------- チェック結果（VBAが作り直す）
ws = wb["チェック結果"]
ws["A1"] = "チェック結果（記入状況の「一括チェック」で作り直します）"
ws["A1"].font = FT_TTL
for i, h in enumerate(["No", "シート", "セル（クリックで移動）", "名前", "判定", "内容"]):
    c = ws.cell(4, 1 + i)
    c.value = h
    c.font = FT_HEAD
    c.fill = F_HEAD
for c, w in zip("ABCDEF", (6, 22, 14, 30, 8, 90)):
    ws.column_dimensions[c].width = w
ws["A5"] = "まだ実行していません（マクロ有効のブックで「一括チェック」を押してください）。"
ws.sheet_properties.tabColor = "A6A6A6"

# ---------------------------------------------------------------- 仕上げ
for nm in wb.sheetnames:
    w = wb[nm]
    w.sheet_view.zoomScale = 90
wb.active = 0
wb.save(OUT)
print("saved", OUT)

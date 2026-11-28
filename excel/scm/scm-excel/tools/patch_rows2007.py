# 登録変更削除シートの枠を 507 行目 → 2007 行目（2000行）に広げる（2026-11-19、1000件単位の追加対応）
import sys, re, openpyxl
from copy import copy
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter as CL
src, out = sys.argv[1], sys.argv[2]
OLD_LAST, ROW_LAST = 507, 2007
RNG_RE = re.compile(r"(\$?[A-Z]{1,3}\$?)([0-9]{1,2}):(\$?[A-Z]{1,3}\$?)507(?![0-9])")
def widen(f):
    if not isinstance(f, str) or "507" not in f: return f
    return RNG_RE.sub(lambda m: f"{m.group(1)}{m.group(2)}:{m.group(3)}{ROW_LAST}" if 6 <= int(m.group(2)) <= 16 else m.group(0), f)
def widen_sqref(sq):
    o = []
    for r in str(sq).split():
        cr = CellRange(r)
        if cr.max_row == OLD_LAST and 6 <= cr.min_row <= 16:
            cr = CellRange(min_col=cr.min_col, min_row=cr.min_row, max_col=cr.max_col, max_row=ROW_LAST)
        o.append(cr.coord)
    return " ".join(o)
wb = openpyxl.load_workbook(src)
row_sheets = [w.title for w in wb.worksheets if w.max_row == OLD_LAST]
print("row sheets:", row_sheets)
nf = 0
for w in wb.worksheets:
    for row in w.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                v = widen(c.value)
                if v != c.value: c.value = v; nf += 1
    for dv in w.data_validations.dataValidation:
        dv.formula1 = widen(dv.formula1)
        dv.sqref = MultiCellRange(widen_sqref(dv.sqref))
    cfs = []
    for cf in w.conditional_formatting:
        rules = []
        for r in cf.rules:
            if r.formula: r.formula = [widen(x) for x in r.formula]
            rules.append(r)
        cfs.append((widen_sqref(cf.sqref), rules))
    w.conditional_formatting = ConditionalFormattingList()
    for sq, rules in cfs:
        for r in rules: w.conditional_formatting.add(sq, r)
for nm_, dn in list(wb.defined_names.items()):
    dn.attr_text = widen(dn.attr_text)
for nm_ in row_sheets:
    w = wb[nm_]
    mc = w.max_column
    tmpl = [w.cell(OLD_LAST, ci) for ci in range(1, mc + 1)]
    for r in range(OLD_LAST + 1, ROW_LAST + 1):
        for ci, s_ in enumerate(tmpl, start=1):
            d = w.cell(r, ci)
            if s_.has_style: d._style = copy(s_._style)
            v = s_.value
            if isinstance(v, str) and v.startswith("="):
                d.value = Translator(v, origin=s_.coordinate).translate_formula(f"{CL(ci)}{r}")
    h = w.row_dimensions[OLD_LAST].height
    if h:
        for r in range(OLD_LAST + 1, ROW_LAST + 1): w.row_dimensions[r].height = h
wb.calculation.fullCalcOnLoad = True
wb.save(out)
print("widened formulas:", nf, "saved", out)

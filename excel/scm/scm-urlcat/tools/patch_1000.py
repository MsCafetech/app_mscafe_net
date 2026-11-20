# URLカテゴリ：③ 登録更新（リスト一覧）を 2007 行目まで、⑤ 登録更新（リスト内容）を 1500 列（B～）まで用意（2026-11-19）
import sys, openpyxl
from copy import copy
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.utils import get_column_letter as CL
src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
w = wb['URLCat_登録更新（リスト一覧）']
OLD, NEW = w.max_row, 2007
tmpl = [w.cell(OLD, c) for c in range(1, w.max_column + 1)]
for r in range(OLD + 1, NEW + 1):
    for ci, s in enumerate(tmpl, start=1):
        d = w.cell(r, ci)
        if s.has_style: d._style = copy(s._style)
    w.cell(r, 1).value = f'=IF(C{r}="","",ROW()-7)'
for dv in w.data_validations.dataValidation:
    out_ = []
    for rg in str(dv.sqref).split():
        cr = CellRange(rg)
        if cr.max_row == OLD: cr = CellRange(min_col=cr.min_col, min_row=cr.min_row, max_col=cr.max_col, max_row=NEW)
        out_.append(cr.coord)
    dv.sqref = MultiCellRange(" ".join(out_))
print('list rows', OLD, '->', NEW)
w = wb['URLCat_登録更新（リスト内容）']
OLDC, NEWC = w.max_column, 1 + 1500
tm = {r: w.cell(r, OLDC) for r in range(1, 9)}
wd = w.column_dimensions[CL(OLDC)].width
for c in range(OLDC + 1, NEWC + 1):
    L = CL(c)
    for r in (5, 6, 7, 8):
        s = tm[r]; d = w.cell(r, c)
        if s.has_style: d._style = copy(s._style)
    w.cell(7, c).value = f'=IF({L}5="","",COUNTA({L}8:{L}1048576))'
    w.column_dimensions[L].width = wd
print('wide cols', OLDC, '->', NEWC)
wb.calculation.fullCalcOnLoad = True
wb.save(out)

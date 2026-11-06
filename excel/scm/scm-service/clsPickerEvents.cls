VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPickerEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================
' clsPickerEvents : このブックのダブルクリックを受けて modSvc の TryPick を呼ぶ
'   （シートモジュールへの貼り付けを不要にするためのクラス）
'==========================================================
Option Explicit

Private WithEvents wb As Workbook
Attribute wb.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set wb = ThisWorkbook
End Sub

Private Sub wb_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If modSvc.TryPick(Target) Then Cancel = True
End Sub

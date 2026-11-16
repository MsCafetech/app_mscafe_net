Attribute VB_Name = "modPickerSetup"
'==========================================================
' modPickerSetup : frmMemberPick（メンバー選択フォーム）をマクロで作成する
'   使い方 : CreatePickerForm を1回実行（frmMemberPick_コード.txt を選択）
'   前提   : ファイル → オプション → トラストセンター → トラストセンターの設定
'            → マクロの設定 →「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック
'            （作成が終わったらチェックを外してかまいません。このモジュールも削除可）
'==========================================================
Option Explicit

Private Const FORM_NAME As String = "frmMemberPick"
Private Const CODE_FILE As String = "frmMemberPick_コード.txt"

Public Sub CreatePickerForm()
    Dim proj As Object, comp As Object, path As String, code As String, n As Long

    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    n = proj.VBComponents.Count     ' 信頼設定が無いとここでエラー
    If Err.Number <> 0 Or proj Is Nothing Then
        On Error GoTo 0
        MsgBox "VBAプロジェクトにアクセスできません。" & vbLf & _
               "トラストセンターの「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを入れてから再実行してください。", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    path = ThisWorkbook.path & Application.PathSeparator & CODE_FILE
    If ThisWorkbook.path = "" Or Dir(path) = "" Then
        path = CStr(Application.GetOpenFilename("テキスト (*.txt),*.txt", , CODE_FILE & " を選択"))
        If path = "False" Then Exit Sub
    End If
    code = ReadAnsi(path)
    If InStr(code, "Private Sub UserForm_Initialize") = 0 Then
        MsgBox "選んだファイルは " & CODE_FILE & " ではないようです。", vbExclamation
        Exit Sub
    End If

    ' 既存（壊れたものを含む）の frmMemberPick を削除してから作り直す
    On Error Resume Next
    Set comp = proj.VBComponents(FORM_NAME)
    On Error GoTo 0
    If Not comp Is Nothing Then
        If MsgBox("既存の " & FORM_NAME & " を削除して作り直します。よろしいですか？", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
        proj.VBComponents.Remove comp
        Set comp = Nothing
    End If

    Set comp = proj.VBComponents.Add(3)          ' 3 = vbext_ct_MSForm（ユーザーフォーム）
    On Error Resume Next
    comp.Name = FORM_NAME
    If Err.Number <> 0 Then
        On Error GoTo 0
        proj.VBComponents.Remove comp            ' 名前を付けられなかったフォームは残さない
        MsgBox "フォーム名「" & FORM_NAME & "」を付けられませんでした。" & vbLf & _
               "ブックを保存して Excel を一度終了し、開き直してから再実行してください。", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString code
    End With
    MsgBox FORM_NAME & " を作成しました。ブックを .xlsm で保存してください。", vbInformation
End Sub

' Shift-JIS（ANSI）のテキストを読み込む
Private Function ReadAnsi(ByVal path As String) As String
    Dim f As Integer, b() As Byte
    f = FreeFile
    Open path For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim b(0 To LOF(f) - 1)
        Get #f, , b
    End If
    Close #f
    If LOF_OK(b) Then ReadAnsi = StrConv(b, vbUnicode)
End Function

Private Function LOF_OK(ByRef b() As Byte) As Boolean
    On Error Resume Next
    LOF_OK = (UBound(b) >= 0)
End Function

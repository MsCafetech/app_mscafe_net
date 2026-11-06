Attribute VB_Name = "modUrlCat"
'==========================================================
' modUrlCat : SCM (Prisma Access Configuration API) カスタムURLカテゴリ（URLリスト）連携
'   前提 : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'   シート :
'     URLCat_既存一覧（リスト一覧）／（リスト内容） … 既存情報取得でSCMの現在を表示（差分の基準）
'     URLCat_登録更新（リスト一覧） … リストの作成・削除、type/descriptionの変更（1行1リスト）
'     URLCat_登録更新（リスト内容） … リストの中身の更新（横長：1列1リスト、1セル1件）
'     URLCat_差分チェック（リスト一覧）／（リスト内容） … 差分チェック・JSON発行で反映予定、API実行後は結果
'   基本設定にWL/BLの名前を書くと、その2つを先頭に並べる
'==========================================================
' ※ 共通処理（認証・HTTP・ログ 等）は modScmCommon にあります。modScmCommon（JSONパーサを含む）も併せてインポートしてください。
Option Explicit

Private Const SH_CFG As String = "基本設定"
Private Const SH_EXIST As String = "URLCat_既存一覧（リスト一覧）"
Private Const SH_LREG As String = "URLCat_登録更新（リスト一覧）"
Private Const SH_CEX As String = "URLCat_既存一覧（リスト内容）"
Private Const SH_CREG As String = "URLCat_登録更新（リスト内容）"
Private Const SH_JSON As String = "JSON_CURL"
Private Const SH_DLIST As String = "URLCat_差分チェック（リスト一覧）"
Private Const SH_DCONT As String = "URLCat_差分チェック（リスト内容）"
Private Const SH_LOG As String = "実行ログ"

Private Const EX_FIRST As Long = 8      ' 既存一覧（リスト一覧） データ開始行
Private Const EX_SHOW As Long = 5       ' 既存一覧（リスト一覧）のリスト内容に表示する件数（残りは「…ほか n 件」）
Private Const LR_FIRST As Long = 8      ' 登録更新（リスト一覧） データ開始行
Private Const LR_OP As Long = 2         '   B 操作（登録／削除）
Private Const LR_NAME As Long = 3       '   C name
Private Const LR_TYPE As Long = 4       '   D type
Private Const LR_DESC As Long = 5       '   E description
Private Const LR_CNT As Long = 6        '   F 件数（既存）
Private Const LR_JUDGE As Long = 7      '   G 判定（VBA）
Private Const LR_MSG As Long = 8        '   H 理由（VBA）
Private Const LR_SPARE As Long = 200    '   新しいリスト用に書式を用意しておく空き行数
Private Const WV_NAME As Long = 5       ' 横長シート（既存一覧・登録更新のリスト内容） name行
Private Const WV_TYPE As Long = 6       '   type行
Private Const WV_CNT As Long = 7        '   件数行（関数）
Private Const WV_FIRST As Long = 8      '   値の開始行
Private Const WV_SPARE As Long = 30     '   登録更新（リスト内容）で新しいリスト用に書式を用意しておく空き列数
Private Const DL_FIRST As Long = 8      ' 差分チェック（リスト一覧） データ開始行
Private Const DL_COLS As Long = 13      '   A～M
Private Const DL_RES As Long = 13       '   M 実行結果
Private Const DC_NAME As Long = 5       ' 差分チェック（リスト内容） name行
Private Const DC_KIND As Long = 6       '   判定行
Private Const DC_CNT As Long = 7        '   件数行
Private Const DC_FIRST As Long = 8      '   値の開始行
Private Const JS_FIRST As Long = 10     ' JSON_CURL データ開始行
Private Const JS_PARTS As Long = 9      ' JSON_CURL JSON分割数列(I)。J列以降に続き
Private Const JS_LASTCOL As Long = 40   ' JSON_CURL クリア対象の最終列
Private Const LOG_FIRST As Long = 6     ' 実行ログ データ開始行
Private Const PAGE_LIMIT As Long = 200
Private Const CELL_MAX As Long = 30000  ' 1セルに入れる最大文字数（Excelの上限32767より小さく）

Private Const CL_NEW As Long = &HCEEFC6     ' 緑  C6EFCE（RGBをBGRで指定）
Private Const CL_MOD As Long = &H99FFFF     ' 黄  FFFF99
Private Const CL_DEL As Long = &HD9D9D9     ' グレー D9D9D9
Private Const CL_CHK As Long = &HCEC7FF     ' 赤  FFC7CE
Private Const CL_WARN As Long = &H9CEBFF    ' 薄黄 FFEB9C
Private Const CL_HI As Long = &HC0FF&       ' オレンジ FFC000（変わる項目のセル）
Private Const CL_IN As Long = &HCCF2FF      ' 入力欄 FFF2CC
Private Const CL_VBA As Long = &HDAEFE2     ' VBA書込み欄 E2EFDA
Private Const CL_SEC As Long = &HF2E1D9     ' 見出し D9E1F2
Private Const CL_HEAD As Long = &H97552F    ' 見出し 2F5597
Private Const FX_STRIKE As Long = -1        ' 色ではなく取り消し線（削除の表示）


'----------------------------------------------------------
' ボタン配置（初回のみ実行）
'----------------------------------------------------------
Public Sub SetupButtons()
    AddBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    AddBtn SH_EXIST, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_CEX, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_LREG, "D3:E4", "差分チェック", "RunDiffCheck"
    AddBtn SH_LREG, "G3:H4", "JSON発行 → JSON_CURLへ", "GenerateJson"
    AddBtn SH_LREG, "J3:K4", "既存一覧から作り直す", "CopyExistingLists"
    AddBtn SH_CREG, "D3:E4", "差分チェック", "RunDiffCheck"
    AddBtn SH_CREG, "G3:H4", "JSON発行 → JSON_CURLへ", "GenerateJson"
    AddBtn SH_CREG, "J3:K4", "既存一覧から作り直す", "CopyExistingContents"
    AddBtn SH_JSON, "G4:H4", "API実行（全行）", "ExecuteAll"
    AddBtn SH_JSON, "G5:H5", "JSONファイル出力", "ExportJsonFiles"
    AddBtn SH_JSON, "G6:H6", "クリア", "ClearJson"
    RemoveBtn SH_JSON, "PushConfig"     ' Push機能は削除（PushはSCM画面から実施）。旧ボタンを消す
    AddBtn SH_DLIST, "D3:E4", "変化なしを隠す", "HideUnchanged"
    AddBtn SH_DLIST, "G3:H4", "全て表示", "ShowAllDiff"
    AddBtn SH_DLIST, "J3:K4", "実行結果の再確認", "BuildDiffCheck"
    AddBtn SH_DCONT, "D3:E4", "変化なしを隠す", "HideUnchangedCont"
    AddBtn SH_DCONT, "G3:H4", "全て表示", "ShowAllCont"
    AddBtn SH_LOG, "F3", "ログクリア", "ClearLog"
    RemoveBtn SH_LOG, "CheckJobs"      ' ジョブ状況確認も削除。旧ボタンを消す
    MsgBox "ボタンを配置しました。", vbInformation
End Sub

'----------------------------------------------------------
' 既存情報取得（GET → 既存一覧2枚と登録更新2枚を作り直す）
'----------------------------------------------------------
Public Sub FetchExisting()
    Dim items As Collection, item As Object, ws As Worksheet, r As Long, meta As Object, have As Object
    Dim msg As String, nUrl As Long, k As Variant
    If Not CheckContainer() Then Exit Sub
    If Not CheckListNames() Then Exit Sub
    ' 登録更新シートに未反映の編集があれば確認（前回取得分と比べる）
    If ListsEdited(ReadExistMeta()) Or ContentsEdited(ReadWideSheet(SH_CEX)) Then
        If MsgBox("登録更新シートに、まだSCMへ反映していない編集があります。" & vbLf & _
                  "既存情報取得をすると、登録更新シートはSCMの現在の内容で作り直され、編集は消えます。続けますか？", _
                  vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    End If
    If Not EnsureToken() Then Exit Sub
    Application.ScreenUpdating = False
    Set items = FetchAllItems("既存取得")
    If items Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Set meta = MetaFromItems(items)
    Set have = HaveFromItems(items)

    ' 既存一覧（リスト一覧）：リスト内容は先頭 EX_SHOW 件だけ表示
    Set ws = Worksheets(SH_EXIST)
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 9)).Clear
    r = EX_FIRST
    For Each item In items
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        ws.Cells(r, 4).Value = Nz(item, "type")
        ws.Cells(r, 5).Value = ListPreview(item)
        ws.Cells(r, 6).NumberFormat = "General"
        ws.Cells(r, 6).Value = ListCount(item)
        ws.Cells(r, 7).Value = Nz(item, "description")
        ws.Cells(r, 8).Value = ContainerOf(item)
        ws.Cells(r, 9).Formula = "=IF(C" & r & "="""","""",IF(AND(CFG_WL_NAME<>"""",EXACT(C" & r & ",CFG_WL_NAME)),""ホワイトリスト""," & _
                                 "IF(AND(CFG_BL_NAME<>"""",EXACT(C" & r & ",CFG_BL_NAME)),""ブラックリスト"",""一般"")))"
        r = r + 1
    Next
    If r > EX_FIRST Then
        With ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 9))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    Range("EX_FETCHED_AT").Value = Now

    WriteListReg meta
    WriteWide SH_CEX, meta, have, False
    WriteWide SH_CREG, meta, have, True
    For Each k In have.Keys
        nUrl = nUrl + have(k).Count
    Next
    msg = "URLリスト " & items.Count & " 件、リスト内容 " & nUrl & " 件" & vbLf & _
          "既存一覧（リスト一覧・リスト内容）と登録更新（リスト一覧・リスト内容）を作り直しました。"
    Application.ScreenUpdating = True
    WriteLog "既存取得", "", "GET", WithQuery(BaseUrl(), ContainerQuery()), 200, Replace(msg, vbLf, " / ")
    MsgBox "取得しました。" & vbLf & msg, vbInformation
End Sub

' 既存一覧（リスト一覧）に出すリスト内容：先頭 EX_SHOW 件と「…ほか n 件」
Private Function ListPreview(ByVal item As Object) As String
    Dim v As Variant, n As Long, s As String
    If ListCount(item) = 0 Then Exit Function
    For Each v In item("list")
        n = n + 1
        If n > EX_SHOW Then Exit For
        If s <> "" Then s = s & vbLf
        s = s & CStr(v)
    Next
    If ListCount(item) > EX_SHOW Then s = s & vbLf & "…ほか " & (ListCount(item) - EX_SHOW) & " 件（全件はリスト内容シート）"
    ListPreview = s
End Function

'----------------------------------------------------------
' 登録更新（リスト一覧）：1行1リスト。B 操作、C name、D type、E description、F 件数（既存）、G 判定、H 理由
'----------------------------------------------------------
Private Sub WriteListReg(ByVal meta As Object)
    Dim ws As Worksheet, names As Collection, a() As Variant, k As Variant, i As Long, n As Long
    Set ws = Worksheets(SH_LREG)
    Set names = WideOrder(meta)
    n = names.Count
    ws.Range(ws.Cells(LR_FIRST, 1), ws.Cells(ws.Rows.Count, LR_MSG)).Clear
    FormatListReg ws, LR_FIRST + n + LR_SPARE - 1
    ws.Range("B3").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Range("B3").Value = Now
    If n = 0 Then Exit Sub
    ReDim a(1 To n, 1 To LR_CNT - LR_OP + 1)
    For Each k In names
        i = i + 1
        a(i, LR_NAME - LR_OP + 1) = k
        a(i, LR_TYPE - LR_OP + 1) = meta(k)(1)
        a(i, LR_DESC - LR_OP + 1) = meta(k)(2)
        a(i, LR_CNT - LR_OP + 1) = meta(k)(3)
    Next
    ws.Cells(LR_FIRST, LR_OP).Resize(n, LR_CNT - LR_OP + 1).Value = a
End Sub

Private Sub FormatListReg(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range(ws.Cells(LR_FIRST, 1), ws.Cells(lastRow, LR_MSG))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlTop
        .Font.Bold = False
        .Font.Color = vbBlack
    End With
    ws.Range(ws.Cells(LR_FIRST, 1), ws.Cells(lastRow, 1)).Formula = "=IF(C" & LR_FIRST & "="""","""",ROW()-" & (LR_FIRST - 1) & ")"
    ws.Range(ws.Cells(LR_FIRST, LR_OP), ws.Cells(lastRow, LR_DESC)).NumberFormat = "@"
    ws.Range(ws.Cells(LR_FIRST, LR_OP), ws.Cells(lastRow, LR_DESC)).Interior.Color = CL_IN
    ws.Range(ws.Cells(LR_FIRST, LR_CNT), ws.Cells(lastRow, LR_MSG)).Interior.Color = CL_VBA
    ws.Range(ws.Cells(LR_FIRST, LR_DESC), ws.Cells(lastRow, LR_DESC)).WrapText = True
    ws.Range(ws.Cells(LR_FIRST, LR_MSG), ws.Cells(lastRow, LR_MSG)).WrapText = True
    With ws.Range(ws.Cells(LR_FIRST, LR_OP), ws.Cells(lastRow, LR_OP)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="登録,削除"
    End With
    With ws.Range(ws.Cells(LR_FIRST, LR_TYPE), ws.Cells(lastRow, LR_TYPE)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="URL List,Category Match"
    End With
End Sub

' 登録更新（リスト一覧）の B～E 列（少なくとも2行の2次元配列）
Private Function ListRegValues() As Variant
    Dim ws As Worksheet, lastR As Long
    Set ws = Worksheets(SH_LREG)
    lastR = LastRow(ws, LR_NAME)
    If LastRow(ws, LR_OP) > lastR Then lastR = LastRow(ws, LR_OP)
    If LastRow(ws, LR_TYPE) > lastR Then lastR = LastRow(ws, LR_TYPE)
    If LastRow(ws, LR_DESC) > lastR Then lastR = LastRow(ws, LR_DESC)
    If lastR < LR_FIRST + 1 Then lastR = LR_FIRST + 1
    ListRegValues = ws.Range(ws.Cells(LR_FIRST, LR_OP), ws.Cells(lastR, LR_DESC)).Value
End Function

' 登録更新（リスト一覧）に反映すると変わる編集があるか
Private Function ListsEdited(ByVal meta As Object) As Boolean
    Dim v As Variant, i As Long, nm As String, typ As String
    v = ListRegValues()
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, LR_NAME - LR_OP + 1)))
        If Trim$(SafeStr(v(i, 1))) <> "" Then GoTo Edited
        If nm <> "" Then
            If Not meta.Exists(nm) Then GoTo Edited
            typ = Trim$(SafeStr(v(i, LR_TYPE - LR_OP + 1)))
            If typ <> "" And typ <> meta(nm)(1) Then GoTo Edited
            If SafeStr(v(i, LR_DESC - LR_OP + 1)) <> meta(nm)(2) Then GoTo Edited
        End If
    Next
    Exit Function
Edited:
    ListsEdited = True
End Function

' 登録更新（リスト内容）に反映すると変わる編集があるか（値が1件も無い列は「変更しない」扱い）
Private Function ContentsEdited(ByVal have As Object) As Boolean
    Dim d As Object, k As Variant
    Set d = ReadWideSheet(SH_CREG)
    For Each k In d.Keys
        If d(k).Count > 0 Then
            If Not have.Exists(k) Then GoTo Edited
            If Not SameSet(have(k), d(k)) Then GoTo Edited
        End If
    Next
    Exit Function
Edited:
    ContentsEdited = True
End Function

'----------------------------------------------------------
' 横長シート（既存一覧・登録更新のリスト内容）：1列1リスト。WL/BL（基本設定の名前）を先頭に、残りはSCMの順
'   5行目 name、6行目 type、7行目 件数、8行目以降 値
'----------------------------------------------------------
Private Sub WriteWide(ByVal shName As String, ByVal meta As Object, ByVal have As Object, ByVal editable As Boolean)
    Dim ws As Worksheet, names As Collection, k As Variant, c As Long, n As Long, mx As Long
    Dim arr() As Variant, i As Long, e As Variant, calc As Long
    Set ws = Worksheets(shName)
    Set names = WideOrder(meta)
    calc = Application.Calculation
    Application.Calculation = xlCalculationManual
    ws.Range(ws.Cells(WV_NAME, 2), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear
    n = names.Count
    For Each k In names
        If have.Exists(k) Then
            If have(k).Count > mx Then mx = have(k).Count
        End If
    Next
    FormatWide ws, n + 1, editable
    ws.Range("B3").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Range("B3").Value = Now
    If n > 0 Then
        ReDim arr(1 To WV_FIRST - WV_NAME + mx, 1 To n)
        For Each k In names
            c = c + 1
            arr(1, c) = k
            arr(WV_TYPE - WV_NAME + 1, c) = meta(k)(1)
            If have.Exists(k) Then
                i = WV_FIRST - WV_NAME
                For Each e In have(k).Keys
                    i = i + 1
                    arr(i, c) = e
                Next
            End If
        Next
        ' 件数行（関数）は上書きしないよう、name・type と値に分けて書き込む
        ws.Cells(WV_NAME, 2).Resize(WV_CNT - WV_NAME, n).Value = SubRows(arr, 1, WV_CNT - WV_NAME, n)
        If mx > 0 Then ws.Cells(WV_FIRST, 2).Resize(mx, n).Value = SubRows(arr, WV_FIRST - WV_NAME + 1, UBound(arr, 1), n)
    End If
    Application.Calculation = calc
End Sub

' 書式・件数の関数。lastCol は最後のリスト列（登録更新は新しいリスト用に WV_SPARE 列多めに用意）
Private Sub FormatWide(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal editable As Boolean)
    Dim upto As Long
    upto = lastCol
    If editable Then upto = lastCol + WV_SPARE
    If upto < 2 Then upto = 2
    ws.Range(ws.Columns(2), ws.Columns(upto)).NumberFormat = "@"       ' 列ごと文字列（UsedRangeを広げない）
    ws.Range(ws.Columns(2), ws.Columns(upto)).ColumnWidth = 30
    With ws.Range(ws.Cells(WV_CNT, 2), ws.Cells(WV_CNT, upto))
        .NumberFormat = "General"
        .FormulaR1C1 = "=IF(R" & WV_NAME & "C="""",""""," & "COUNTA(R" & WV_FIRST & "C:R" & ws.Rows.Count & "C))"
    End With
    PaintWideHead ws, upto
End Sub

' 見出し（5～7行目）の色を元に戻す：name=紺、type・件数=薄青
Private Sub PaintWideHead(ByVal ws As Worksheet, ByVal upto As Long)
    If upto < 2 Then Exit Sub
    With ws.Range(ws.Cells(WV_NAME, 2), ws.Cells(WV_CNT, upto))
        .Interior.Color = CL_SEC
        .Font.Bold = False
        .Font.Color = vbBlack
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    With ws.Range(ws.Cells(WV_NAME, 2), ws.Cells(WV_NAME, upto))
        .Interior.Color = CL_HEAD
        .Font.Bold = True
        .Font.Color = vbWhite
    End With
End Sub

' 横長シートの最後のリスト列（name か、値のある最後の列）
Private Function WideLastCol(ByVal ws As Worksheet) As Long
    Dim c As Long
    c = ws.Cells(WV_NAME, ws.Columns.Count).End(xlToLeft).Column
    With ws.UsedRange
        If .Column + .Columns.Count - 1 > c Then c = .Column + .Columns.Count - 1
    End With
    Do While c >= 2
        If SafeStr(ws.Cells(WV_NAME, c).Value) <> "" Then Exit Do
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(WV_FIRST, c), ws.Cells(ws.Rows.Count, c))) > 0 Then Exit Do
        c = c - 1
    Loop
    WideLastCol = c
End Function

' 横長シートの最終行（値が無くても WV_FIRST）
Private Function WideMaxRow(ByVal ws As Worksheet) As Long
    With ws.UsedRange
        WideMaxRow = .Row + .Rows.Count - 1
    End With
    If WideMaxRow < WV_FIRST Then WideMaxRow = WV_FIRST
End Function

' 横長シートを読む：リスト名 -> Dictionary(値 -> True)
Private Function ReadWideSheet(ByVal shName As String) As Object
    Dim ws As Worksheet, d As Object, v As Variant, lastC As Long, c As Long, r As Long, nm As String, e As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets(shName)
    lastC = WideLastCol(ws)
    If lastC >= 2 Then
        v = ws.Range(ws.Cells(WV_NAME, 2), ws.Cells(WideMaxRow(ws), lastC)).Value
        For c = 1 To lastC - 1
            nm = Trim$(SafeStr(v(1, c)))
            If nm <> "" And Not d.Exists(nm) Then
                d.Add nm, CreateObject("Scripting.Dictionary")
                For r = WV_FIRST - WV_NAME + 1 To UBound(v, 1)
                    e = Trim$(SafeStr(v(r, c)))
                    If e <> "" Then AddEntry d, nm, e
                Next
            End If
        Next
    End If
    Set ReadWideSheet = d
End Function

' 既存一覧（リスト一覧）：name -> Array(id, type, description, 件数)
Private Function ReadExistMeta() As Object
    Dim d As Object, ws As Worksheet, r As Long, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets(SH_EXIST)
    For r = EX_FIRST To LastRow(ws, 3)
        nm = SafeStr(ws.Cells(r, 3).Value)
        If nm <> "" And Not d.Exists(nm) Then
            d.Add nm, Array(SafeStr(ws.Cells(r, 2).Value), SafeStr(ws.Cells(r, 4).Value), SafeStr(ws.Cells(r, 7).Value), Val(SafeStr(ws.Cells(r, 6).Value)))
        End If
    Next
    Set ReadExistMeta = d
End Function

'----------------------------------------------------------
' 既存一覧から作り直す（APIは使わない）
'----------------------------------------------------------
Public Sub CopyExistingLists()
    Dim meta As Object
    If Not ExistLoaded() Then Exit Sub
    Set meta = ReadExistMeta()
    If ListsEdited(meta) Then
        If MsgBox("登録更新（リスト一覧）を既存一覧の内容で作り直します（今の編集は消えます）。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    End If
    Application.ScreenUpdating = False
    WriteListReg meta
    Application.ScreenUpdating = True
    MsgBox "登録更新（リスト一覧）を作り直しました。", vbInformation
End Sub

Public Sub CopyExistingContents()
    Dim meta As Object, have As Object
    If Not ExistLoaded() Then Exit Sub
    Set meta = ReadExistMeta()
    Set have = ReadWideSheet(SH_CEX)
    If ContentsEdited(have) Then
        If MsgBox("登録更新（リスト内容）を既存一覧の内容で作り直します（今の編集は消えます）。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    End If
    Application.ScreenUpdating = False
    WriteWide SH_CREG, meta, have, True
    Application.ScreenUpdating = True
    MsgBox "登録更新（リスト内容）を作り直しました。追加は列の下に追記、削除はセルを消してください（詰めなくてもかまいません）。", vbInformation
End Sub

' 既存情報取得が済んでいるか
Private Function ExistLoaded() As Boolean
    If IsEmpty(Range("EX_FETCHED_AT").Value) Or SafeStr(Range("EX_FETCHED_AT").Value) = "" Then
        MsgBox "先に「既存情報取得」を実行してください（差分は前回の既存情報取得の内容と比べます）。", vbExclamation
        Exit Function
    End If
    ExistLoaded = True
End Function

'----------------------------------------------------------
' 差分チェック（両方の登録更新シートをチェックし、差分チェックシート2枚に反映予定を出す）
'   差分の基準は既存一覧シート（前回の既存情報取得の時点）
'----------------------------------------------------------
Public Sub RunDiffCheck()
    Dim r As Object
    If Not CheckListNames() Then Exit Sub
    If Not ExistLoaded() Then Exit Sub
    Application.ScreenUpdating = False
    Set r = CheckAll("差分チェック（JSON未作成）")
    Application.ScreenUpdating = True
    If r("ng") > 0 Then
        If r("lc")("ng") > 0 Then Worksheets(SH_LREG).Activate Else Worksheets(SH_CREG).Activate
    Else
        Worksheets(SH_DLIST).Activate
    End If
    MsgBox "差分チェックが終わりました（前回の既存情報取得の内容と比較）。" & vbLf & _
           DiffSummary(r("diffs")) & vbLf & _
           "NG " & r("ng") & " 件（赤）、警告 " & r("warn") & " 件（黄）" & _
           IIf(r("chk")("over") <> "", vbLf & "上限超のリスト：" & r("chk")("over"), "") & vbLf & _
           "NGと警告の理由は、リスト一覧はH列、リスト内容はセルのコメントにあります。" & vbLf & _
           "差分は差分チェック（リスト一覧・リスト内容）シートに出しています（追加=緑、削除=グレーと取り消し線）。", _
           IIf(r("ng") > 0, vbExclamation, vbInformation)
End Sub

' 両方の登録更新シートをチェックして計画を作り、差分チェックシート2枚を作り直す
'   戻り値 Dictionary：meta, have, lc（CheckListRows）, chk（CheckUrlRows）, diffs, ng, warn
Private Function CheckAll(ByVal status As String) As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    Set r("meta") = ReadExistMeta()
    Set r("have") = ReadWideSheet(SH_CEX)
    Set r("lc") = CheckListRows(r("meta"))
    Set r("chk") = CheckUrlRows(r("meta"), r("have"), r("lc")("cats"))
    Set r("diffs") = PlanAll(r("meta"), r("have"), r("lc")("cats"), r("chk")("want"))
    RenderListDiff r("diffs"), r("meta"), r("lc")("rows"), status
    RenderContentDiff r("diffs"), r("chk"), r("have"), status
    r("ng") = r("lc")("ng") + r("chk")("ng")
    r("warn") = r("lc")("warn") + r("chk")("warn")
    Set CheckAll = r
End Function

'----------------------------------------------------------
' 登録更新（リスト一覧）のチェック。G列に判定、H列に理由を書き、色を付ける
'   戻り値 Dictionary：cats（name -> Array(操作 登録/変更/削除, type, description, 既存id, 行)）,
'     rows（name -> 行。このシートにある名前）, ng, warn
'----------------------------------------------------------
Private Function CheckListRows(ByVal meta As Object) As Object
    Dim ws As Worksheet, res As Object, cats As Object, rows As Object, paint As Object, v As Variant, out() As Variant
    Dim i As Long, n As Long, r As Long, nm As String, uop As String, typ As String, desc As String, msg As String
    Dim op As String, id As String, exists As Boolean, nNg As Long, nWarn As Long
    Set res = CreateObject("Scripting.Dictionary")
    Set cats = CreateObject("Scripting.Dictionary")
    Set rows = CreateObject("Scripting.Dictionary")
    Set paint = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets(SH_LREG)
    v = ListRegValues()
    n = UBound(v, 1)
    ReDim out(1 To n, 1 To 2)
    With ws.Range(ws.Cells(LR_FIRST, LR_OP), ws.Cells(LR_FIRST + n - 1, LR_DESC))
        .Interior.Color = CL_IN
        .Font.Bold = False
    End With
    ws.Range(ws.Cells(LR_FIRST, LR_JUDGE), ws.Cells(LR_FIRST + n - 1, LR_MSG)).Interior.Color = CL_VBA
    For i = 1 To n
        r = LR_FIRST + i - 1
        uop = Trim$(SafeStr(v(i, 1)))
        nm = Trim$(SafeStr(v(i, LR_NAME - LR_OP + 1)))
        typ = Trim$(SafeStr(v(i, LR_TYPE - LR_OP + 1)))
        desc = SafeStr(v(i, LR_DESC - LR_OP + 1))
        msg = "": op = ""
        exists = False
        If nm = "" Then
            If uop <> "" Or typ <> "" Or desc <> "" Then msg = "NG:nameが空欄です"
        ElseIf rows.Exists(nm) Then
            msg = "NG:同じnameの行が " & rows(nm) & " 行目にもあります。この行は反映しません"
        Else
            rows.Add nm, r
            exists = meta.Exists(nm)
            Select Case uop
                Case "登録"
                    If exists Then
                        msg = "NG:SCMに同じ名前のリストがあります（登録できません）"
                    Else
                        op = "登録"
                        If typ = "" Then
                            typ = "URL List"
                            msg = "警告:typeが空欄のため URL List で登録します"
                        End If
                    End If
                Case "削除"
                    If Not exists Then
                        msg = "NG:SCMに無いリストは削除できません"
                    Else
                        op = "削除"
                        If IsListName(nm) Then msg = "警告:ホワイトリスト/ブラックリスト（基本設定の名前）を削除します。ポリシーで使っていないか確認してください"
                    End If
                Case ""
                    If Not exists Then
                        msg = "NG:SCMに無いリストです。新しく作る場合は操作を「登録」にしてください"
                    Else
                        If typ = "" Then typ = meta(nm)(1)
                        If typ <> meta(nm)(1) Then
                            op = "変更"
                            PaintAdd paint, CL_HI, ws, r, LR_TYPE
                        End If
                        If desc <> meta(nm)(2) Then
                            op = "変更"
                            PaintAdd paint, CL_HI, ws, r, LR_DESC
                            If desc = "" Then msg = "警告:descriptionを空欄にすると、既存のdescriptionは消えます"
                        End If
                    End If
                Case Else
                    msg = "NG:操作は「登録」「削除」か空欄にしてください"
            End Select
            If op = "登録" Or op = "変更" Then
                If typ <> "URL List" And typ <> "Category Match" Then
                    msg = "NG:typeは URL List か Category Match です"
                    op = ""
                End If
            End If
        End If
        If op <> "" Then
            id = ""
            If exists Then id = meta(nm)(0)
            cats.Add nm, Array(op, typ, desc, id, r)
            If op = "登録" Then PaintAdd paint, CL_NEW, ws, r, LR_OP
            If op = "削除" Then
                PaintAdd paint, CL_DEL, ws, r, LR_OP
                PaintAddRange paint, FX_STRIKE, ws.Range(ws.Cells(r, LR_NAME), ws.Cells(r, LR_DESC))
            End If
        End If
        If Left$(msg, 2) = "NG" Then
            out(i, 1) = "NG"
            nNg = nNg + 1
            PaintAdd paint, CL_CHK, ws, r, LR_JUDGE
        ElseIf nm <> "" Then
            out(i, 1) = IIf(op <> "", op, "変更なし") & IIf(msg <> "", "・警告", "")
            If msg <> "" Then
                nWarn = nWarn + 1
                PaintAdd paint, CL_WARN, ws, r, LR_JUDGE
            End If
        End If
        If msg <> "" Then out(i, 2) = Mid$(msg, InStr(msg, ":") + 1)
    Next
    ws.Range(ws.Cells(LR_FIRST, LR_JUDGE), ws.Cells(LR_FIRST + n - 1, LR_MSG)).Value = out
    ws.Range(ws.Cells(LR_FIRST, LR_OP), ws.Cells(LR_FIRST + n - 1, LR_DESC)).Font.Strikethrough = False
    ApplyPaint ws, paint
    Set res("cats") = cats
    Set res("rows") = rows
    res("ng") = nNg
    res("warn") = nWarn
    Set CheckListRows = res
End Function

'----------------------------------------------------------
' 登録更新（リスト内容）のチェック（NG=赤、警告=黄のセル色と、理由をセルのコメントに書く）
'   追加する値=緑、削除するリストの値=グレー
'   シートは配列で一度に読み、色は同じ色のセルをまとめて塗る（件数が多くても遅くならないように）
'   戻り値 Dictionary：want（リスト名 -> Dictionary(値)）, hdr（リスト名 -> 列）, ngcols（反映しない列 -> True）,
'     rows, ng, warn, over
'----------------------------------------------------------
Private Function CheckUrlRows(ByVal meta As Object, ByVal have As Object, ByVal cats As Object) As Object
    Dim ws As Worksheet, res As Object, want As Object, seen As Object, seenL As Object
    Dim overs As Object, hdr As Object, paint As Object, oldE As Object, ngcols As Object
    Dim marks As Collection, dummies As Collection, x As Variant, v As Variant
    Dim lastC As Long, maxR As Long, c As Long, r As Long, rr As Long, nE As Long
    Dim nm As String, e As String, msg As String, op As String, typ As String
    Dim colNg As String, isCm As Boolean, exists As Boolean, isNew As Boolean
    Dim mx As Long, nRows As Long, nNg As Long, nWarn As Long, nCmt As Long, k As Variant, ov As String, dmy As String
    Set res = CreateObject("Scripting.Dictionary")
    Set want = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set seenL = CreateObject("Scripting.Dictionary")
    Set overs = CreateObject("Scripting.Dictionary")
    Set hdr = CreateObject("Scripting.Dictionary")
    Set ngcols = CreateObject("Scripting.Dictionary")
    Set paint = CreateObject("Scripting.Dictionary")   ' 色 -> Collection(セル番地)
    Set marks = New Collection                          ' Array(行, 列, 判定文字列)
    Set dummies = New Collection                        ' Array(行, 列, リスト名)
    mx = Val(Cfg("CFG_LIST_MAX"))
    If mx <= 0 Then mx = 2500
    dmy = DummyEntry()
    Set ws = Worksheets(SH_CREG)
    lastC = WideLastCol(ws)
    maxR = WideMaxRow(ws)
    If lastC >= 2 Then
        ' 色とコメントをいったん消す
        ws.Range(ws.Cells(WV_NAME, 2), ws.Cells(maxR, lastC)).ClearComments
        PaintWideHead ws, lastC
        With ws.Range(ws.Cells(WV_FIRST, 2), ws.Cells(maxR, lastC))
            .Interior.Pattern = xlNone
            .Font.Bold = False
            .Font.Color = vbBlack
            .Font.Strikethrough = False
        End With
        v = ws.Range(ws.Cells(WV_NAME, 2), ws.Cells(maxR, lastC)).Value
    End If

    For c = 2 To lastC
        nm = Trim$(SafeStr(v(1, c - 1)))
        exists = False
        If nm <> "" Then exists = meta.Exists(nm)
        op = ""
        If nm <> "" Then
            If cats.Exists(nm) Then op = cats(nm)(0)
        End If
        colNg = ""
        If nm = "" Then
            colNg = "NG:この列のリスト名（" & WV_NAME & "行目）が空欄です"
        ElseIf hdr.Exists(nm) Then
            colNg = "NG:同じリスト名の列が " & ws.Cells(WV_NAME, hdr(nm)).Address(False, False) & " にもあります。この列は反映しません"
            marks.Add Array(WV_NAME, c, colNg)
        Else
            hdr.Add nm, c
            If Not exists And op <> "登録" Then
                colNg = "NG:SCMに無いリストです。新しく作る場合は登録更新（リスト一覧）で操作を「登録」にしてください"
                marks.Add Array(WV_NAME, c, colNg)
            End If
        End If
        typ = ""
        If op = "登録" Or op = "変更" Then
            typ = cats(nm)(1)
        ElseIf exists Then
            typ = meta(nm)(1)
        End If
        isCm = (typ = "Category Match")
        Set oldE = Nothing
        If exists Then
            If have.Exists(nm) Then Set oldE = have(nm)
        End If
        If op = "削除" Then PaintAdd paint, CL_DEL, ws, WV_NAME, c

        nE = 0
        For r = WV_FIRST - WV_NAME + 1 To UBound(v, 1)
            e = Trim$(SafeStr(v(r, c - 1)))
            If e <> "" Then
                nRows = nRows + 1
                nE = nE + 1
                rr = WV_NAME + r - 1
                If colNg <> "" Then
                    msg = colNg
                ElseIf op = "削除" Then
                    msg = ""
                    PaintAdd paint, CL_DEL, ws, rr, c
                    PaintAdd paint, FX_STRIKE, ws, rr, c
                ElseIf e = dmy Then
                    msg = ""
                Else
                    msg = EntryCheck(e, isCm)
                End If
                If colNg = "" And op <> "削除" And Left$(msg, 2) <> "NG" Then
                    If seen.Exists(nm & vbTab & e) Then
                        msg = "NG:同じリストに同じ値があります（重複）"
                    ElseIf seenL.Exists(nm & vbTab & LCase$(e)) And msg = "" Then
                        msg = "警告:同じリストに大文字・小文字だけ異なる値があります"
                    End If
                    If Left$(msg, 2) <> "NG" Then
                        seen(nm & vbTab & e) = True
                        seenL(nm & vbTab & LCase$(e)) = True
                        AddEntry want, nm, e
                        If msg = "" And e <> dmy Then
                            isNew = True
                            If Not oldE Is Nothing Then isNew = Not oldE.Exists(e)
                            If isNew Then PaintAdd paint, CL_NEW, ws, rr, c
                        End If
                    End If
                End If
                If msg <> "" Then
                    marks.Add Array(rr, c, msg)
                ElseIf e = dmy And colNg = "" And op <> "削除" Then
                    dummies.Add Array(rr, c, nm)
                End If
            End If
        Next
        If colNg <> "" And (nm <> "" Or nE > 0) Then ngcols(c) = True
    Next

    ' 1リストの上限超はリストごと NG（登録・変更もしない）
    For Each k In want.Keys
        If want(k).Count > mx Then
            overs(k) = want(k).Count
            If ov <> "" Then ov = ov & "、"
            ov = ov & k & "（" & want(k).Count & " 件）"
        End If
    Next
    For Each k In overs.Keys
        want.Remove k
        If cats.Exists(k) Then cats.Remove k
        ngcols(hdr(k)) = True
        marks.Add Array(WV_NAME, hdr(k), "NG:このリストの件数が上限(" & mx & ")超（" & overs(k) & " 件）。このリストは反映しません")
    Next
    ' ダミーは、ほかに値があるときだけ「消してよい」を出す
    For Each x In dummies
        If want.Exists(x(2)) Then
            If want(x(2)).Count > 1 Then marks.Add Array(x(0), x(1), "警告:ダミー（" & dmy & "）です。ほかの値を追記したので消してかまいません")
        End If
    Next

    ' NG・警告の色（緑・グレー・橙より後に塗るので優先される）とコメント
    For Each x In marks
        msg = x(2)
        If Left$(msg, 2) = "NG" Then
            PaintAdd paint, CL_CHK, ws, x(0), x(1)
            nNg = nNg + 1
        Else
            PaintAdd paint, CL_WARN, ws, x(0), x(1)
            nWarn = nWarn + 1
        End If
    Next
    ApplyPaint ws, paint
    For Each x In marks
        If nCmt >= 3000 Then Exit For                  ' コメントが多すぎると重くなるため上限
        msg = x(2)
        If ws.Cells(x(0), x(1)).Comment Is Nothing Then
            ws.Cells(x(0), x(1)).AddComment Mid$(msg, InStr(msg, ":") + 1)
            nCmt = nCmt + 1
        End If
    Next
    Set res("want") = want
    Set res("hdr") = hdr
    Set res("ngcols") = ngcols
    res("rows") = nRows
    res("ng") = nNg
    res("warn") = nWarn
    res("over") = ov
    Set CheckUrlRows = res
End Function

' 塗る予定のセルを色ごとに貯める（FX_STRIKE は取り消し線）
Private Sub PaintAdd(ByVal paint As Object, ByVal clr As Long, ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long)
    PaintAddRange paint, clr, ws.Cells(r, c)
End Sub

Private Sub PaintAddRange(ByVal paint As Object, ByVal clr As Long, ByVal rg As Range)
    If Not paint.Exists(clr) Then paint.Add clr, New Collection
    paint(clr).Add rg.Address(False, False)
End Sub

' 色ごとにセル番地をまとめて塗る（Range("A1,B2,…") は255文字までなので区切って塗る）
Private Sub ApplyPaint(ByVal ws As Worksheet, ByVal paint As Object)
    Dim clr As Variant, a As Variant, s As String
    For Each clr In Array(CL_NEW, CL_MOD, CL_DEL, CL_HI, CL_WARN, CL_CHK, FX_STRIKE)
        If paint.Exists(clr) Then
            s = ""
            For Each a In paint(clr)
                If Len(s) + Len(a) > 240 Then
                    PaintRange ws.Range(s), CLng(clr)
                    s = ""
                End If
                If s <> "" Then s = s & ","
                s = s & a
            Next
            If s <> "" Then PaintRange ws.Range(s), CLng(clr)
        End If
    Next
End Sub

Private Sub PaintRange(ByVal rg As Range, ByVal clr As Long)
    If clr = FX_STRIKE Then
        rg.Font.Strikethrough = True
        Exit Sub
    End If
    rg.Interior.Color = clr
    rg.Font.Color = vbBlack
    If clr = CL_HI Then rg.Font.Bold = True
End Sub

' 2次元配列の行 r1～r2 を取り出す
Private Function SubRows(ByRef a() As Variant, ByVal r1 As Long, ByVal r2 As Long, ByVal nCol As Long) As Variant
    Dim b() As Variant, r As Long, c As Long
    ReDim b(1 To r2 - r1 + 1, 1 To nCol)
    For r = r1 To r2
        For c = 1 To nCol
            b(r - r1 + 1, c) = a(r, c)
        Next
    Next
    SubRows = b
End Function

' 横長ビューの列の順：ホワイトリスト → ブラックリスト → 残り（SCMの順）
Private Function WideOrder(ByVal meta As Object) As Collection
    Dim c As New Collection, lk As Variant, nm As String, k As Variant
    For Each lk In Array("WL", "BL")
        nm = Cfg("CFG_" & lk & "_NAME")
        If nm <> "" Then
            If meta.Exists(nm) Then c.Add nm
        End If
    Next
    For Each k In meta.Keys
        If Not IsListName(CStr(k)) Then c.Add CStr(k)
    Next
    Set WideOrder = c
End Function

'----------------------------------------------------------
' 取得結果の整理
'----------------------------------------------------------
' 取得結果：name -> Array(id, type, description, 件数)
Private Function MetaFromItems(ByVal items As Collection) As Object
    Dim d As Object, item As Object, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    For Each item In items
        nm = Nz(item, "name")
        If nm <> "" And Not d.Exists(nm) Then
            d.Add nm, Array(Nz(item, "id"), Nz(item, "type"), Nz(item, "description"), ListCount(item))
        End If
    Next
    Set MetaFromItems = d
End Function

' 取得結果：カテゴリ名 -> Dictionary(エントリ -> True)
Private Function HaveFromItems(ByVal items As Collection) As Object
    Dim d As Object, item As Object, nm As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each item In items
        nm = Nz(item, "name")
        If nm <> "" And Not d.Exists(nm) Then
            d.Add nm, CreateObject("Scripting.Dictionary")
            If ListCount(item) > 0 Then
                For Each v In item("list")
                    AddEntry d, nm, CStr(v)
                Next
            End If
        End If
    Next
    Set HaveFromItems = d
End Function

Private Sub AddEntry(ByVal d As Object, ByVal nm As String, ByVal e As String)
    Dim s As Object
    If Not d.Exists(nm) Then d.Add nm, CreateObject("Scripting.Dictionary")
    Set s = d(nm)
    If e <> "" And Not s.Exists(e) Then s.Add e, True
End Sub

'----------------------------------------------------------
' 値のチェック
'----------------------------------------------------------
' 空カテゴリ用のダミー（基本設定 CFG_DUMMY、未設定なら dummy.fqdn/）
Private Function DummyEntry() As String
    DummyEntry = CfgOpt("CFG_DUMMY")
    If DummyEntry = "" Then DummyEntry = "dummy.fqdn/"
End Function

' 1エントリのチェック。問題なしは ""、それ以外は "NG:..." / "警告:..."
Private Function EntryCheck(ByVal e As String, ByVal isCm As Boolean) As String
    Dim i As Long, cd As Long
    For i = 1 To Len(e)
        cd = AscW(Mid$(e, i, 1))
        If cd < 33 Or cd > 126 Then
            If cd = 32 Or cd = 9 Or cd = 10 Or cd = 13 Or cd = 12288 Then
                EntryCheck = "NG:空白・改行が含まれます（1セル1件）"
            Else
                EntryCheck = "NG:全角文字・制御文字が含まれます"
            End If
            Exit Function
        End If
    Next
    If Len(e) > 255 Then
        EntryCheck = "NG:255文字超"
    ElseIf isCm Then
        EntryCheck = ""
    ElseIf LCase$(Left$(e, 7)) = "http://" Or LCase$(Left$(e, 8)) = "https://" Then
        EntryCheck = "NG:http:// / https:// は不要です"
    Else
        EntryCheck = UrlRuleCheck(e)
    End If
End Function

' Palo Alto の URL リストの書式ルール（ホスト部 = 最初の / より前）
'   aa*.aaa.com / *.*.com は登録不可（NG）
'   aaa.aaa.com（末尾 / なし）は前方一致で aaa.aaa.com.net 等にも一致 → aaa.aaa.com/ を推奨（警告）
'   aaa.*.com/ は登録可能だが非推奨、*.aaa.com/aaa はSSL復号化が前提（警告）
Private Function UrlRuleCheck(ByVal e As String) As String
    Dim host As String, path As String, p As Long, lb() As String, i As Long, nStar As Long, starPos As Long, w As String
    p = InStr(e, "/")
    If p > 0 Then
        host = Left$(e, p - 1)
        path = Mid$(e, p + 1)
    Else
        host = e
    End If
    If host = "" Then
        UrlRuleCheck = "NG:ホスト名がありません（先頭が / です）"
        Exit Function
    End If
    lb = Split(host, ".")
    For i = 0 To UBound(lb)
        If InStr(lb(i), "*") > 0 Then
            If lb(i) <> "*" Then
                UrlRuleCheck = "NG:ワイルドカード(*)はドットで区切った1区切り全体にしか使えません（aa*.aaa.com は登録不可。*.aaa.com/ の形にしてください）"
                Exit Function
            End If
            nStar = nStar + 1
            If starPos = 0 Then starPos = i + 1
        End If
    Next
    If nStar > 1 Then
        UrlRuleCheck = "NG:ワイルドカード(*)は1か所だけです（*.*.com は登録不可）"
        Exit Function
    End If
    If InStr(host, ".") = 0 Then w = w & "ドット(.)を含みません／"
    If nStar = 1 And starPos > 1 Then w = w & "途中のワイルドカード（aaa.*.com/）は登録できますが推奨しません／"
    If p = 0 Then
        w = w & "末尾に / が無いため前方一致になり、" & e & ".net なども一致します（" & e & "/ を推奨）／"
    ElseIf path <> "" Then
        w = w & "パス指定（/ 以降）はSSL復号化が前提です／"
    End If
    If w <> "" Then UrlRuleCheck = "警告:" & Left$(w, Len(w) - 1)
End Function

'----------------------------------------------------------
' JSON発行（登録更新シート2枚 → JSON_CURL）
'   差分チェックと同じチェックをして差分チェックシート2枚を作り直し、JSON_CURL を作り直す
'   実行順：リストの登録・変更（リスト一覧の行順） → 内容だけの変更（名前順） → リストの削除（リスト一覧の行順）
'----------------------------------------------------------
Public Sub GenerateJson()
    Dim wsOut As Worksheet, o As Long, r As Object, c As Object, diffs As Collection, warnDel As Boolean
    Dim nNew As Long, nMod As Long, nModUrl As Long, nDel As Long, nAddE As Long, nDelE As Long, noRow As String, nNoRow As Long
    Dim ordered As Collection, oc As Variant
    If Not CheckContainer() Then Exit Sub
    If Not CheckListNames() Then Exit Sub
    If Not ExistLoaded() Then Exit Sub
    Set wsOut = Worksheets(SH_JSON)
    Application.ScreenUpdating = False
    Set r = CheckAll("プレビュー（JSON未作成）")
    Application.ScreenUpdating = True
    If r("ng") > 0 Then
        If MsgBox("判定NGが " & r("ng") & " 件あります（リスト一覧：" & r("lc")("ng") & " 件、リスト内容：" & r("chk")("ng") & " 件）。" & vbLf & _
                  IIf(r("chk")("over") <> "", "上限超のリスト：" & r("chk")("over") & vbLf, "") & _
                  "NGの行・セル・列を除外してJSONを発行しますか？（理由はリスト一覧のH列、リスト内容のセルのコメント）", vbYesNo + vbQuestion) = vbNo Then
            If r("lc")("ng") > 0 Then Worksheets(SH_LREG).Activate Else Worksheets(SH_CREG).Activate
            Exit Sub
        End If
    End If

    Set diffs = r("diffs")
    For Each c In diffs
        Select Case c("kind")
            Case "新規": nNew = nNew + 1
            Case "変更": nMod = nMod + 1: If c("op") = "変更(URL)" Then nModUrl = nModUrl + 1
            Case "削除": nDel = nDel + 1
        End Select
        nAddE = nAddE + c("nAdd")
        nDelE = nDelE + c("nDel")
        If c("noRow") Then
            nNoRow = nNoRow + 1
            If nNoRow <= 10 Then noRow = noRow & vbLf & "　・" & c("name")
        End If
    Next
    If nNoRow > 10 Then noRow = noRow & vbLf & "　…ほか " & (nNoRow - 10) & " 件"
    If nDel > 0 Or nDelE > 0 Then warnDel = True
    If MsgBox("リスト：新規 " & nNew & " 件、変更 " & nMod & " 件（うち内容だけ " & nModUrl & " 件）、削除 " & nDel & " 件" & vbLf & _
              "　値：追加 " & nAddE & " 件／削除 " & nDelE & " 件（詳細は差分チェックシート）" & _
              IIf(nNoRow > 0, vbLf & "リスト内容に値が無いため中身を変更しないリスト（" & nNoRow & " 件）：" & noRow, "") & vbLf & vbLf & _
              IIf(warnDel, "※削除予定があります。差分チェックシートのグレー（取り消し線）を確認してください。" & vbLf, "") & _
              "JSON_CURLにJSONを作成します。続けますか？", vbYesNo + IIf(warnDel, vbExclamation, vbQuestion)) = vbNo Then
        Worksheets(SH_DLIST).Activate
        Exit Sub
    End If

    Application.ScreenUpdating = False
    ClearJsonRows wsOut
    o = JS_FIRST
    ' 1. リストの登録・変更（リスト一覧の行順）→ 内容だけの変更（名前順）
    Set ordered = New Collection
    For Each c In diffs
        If c("op") = "登録" Or c("op") = "変更" Then ordered.Add c
    Next
    For Each c In diffs
        If c("op") = "変更(URL)" Then ordered.Add c
    Next
    For Each oc In SortByRow(ordered)
        WriteJsonRow wsOut, o, oc("op"), oc("name"), oc("method"), oc("url"), oc("body")
    Next
    ' 2. リストの削除（リスト一覧の行順）
    Set ordered = New Collection
    For Each c In diffs
        If c("op") = "削除" Then ordered.Add c
    Next
    For Each oc In SortByRow(ordered)
        WriteJsonRow wsOut, o, oc("op"), oc("name"), oc("method"), oc("url"), ""
    Next
    If o > JS_FIRST Then
        With wsOut.Range(wsOut.Cells(JS_FIRST, 1), wsOut.Cells(o - 1, 9))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    Worksheets(SH_DLIST).Range("B4").Value = "プレビュー（JSON作成済み・未実行）"
    Worksheets(SH_DCONT).Range("B4").Value = "プレビュー（JSON作成済み・未実行）"
    Application.ScreenUpdating = True
    wsOut.Activate
    MsgBox (o - JS_FIRST) & " 件のJSONをJSON_CURLに生成しました。" & vbLf & _
           "反映予定は差分チェック（リスト一覧・リスト内容）シートに出しています。", vbInformation
End Sub

'----------------------------------------------------------
' 差分チェック（リスト一覧）：1行1リスト
'   A No、B name、C 判定、D 操作、E/F type（現在/反映後）、G/H description（現在/反映後）、
'   I/J 件数（現在/反映後）、K 追加、L 削除、M 実行結果
'   行の色：新規=緑、変更=黄、削除=グレー（取り消し線）。変わる項目は橙
'----------------------------------------------------------
Private Sub RenderListDiff(ByVal diffs As Collection, ByVal meta As Object, ByVal rows As Object, ByVal status As String)
    Dim ws As Worksheet, order As Collection, byName As Object, c As Object, k As Variant
    Dim a() As Variant, paint As Object, i As Long, r As Long, n As Long, nm As String, bt As String, bd As String
    Set ws = Worksheets(SH_DLIST)
    ws.Rows(DL_FIRST & ":" & ws.Rows.Count).Hidden = False
    ws.Range(ws.Cells(DL_FIRST, 1), ws.Cells(ws.Rows.Count, DL_COLS)).Clear
    ws.Range("B3").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Range("B3").Value = Now
    ws.Range("B4").Value = status
    ' 並び：登録更新（リスト一覧）の行順 → その他（名前順）
    Set byName = CreateObject("Scripting.Dictionary")
    For Each c In diffs
        Set byName(c("name")) = c
    Next
    Set order = New Collection
    For Each k In rows.Keys
        If byName.Exists(k) Then order.Add byName(k)
    Next
    For Each c In diffs
        If Not rows.Exists(c("name")) Then order.Add c
    Next
    n = order.Count
    If n = 0 Then Exit Sub
    ReDim a(1 To n, 1 To DL_COLS)
    Set paint = CreateObject("Scripting.Dictionary")
    For Each c In order
        i = i + 1
        r = DL_FIRST + i - 1
        nm = c("name")
        bt = "": bd = ""
        If meta.Exists(nm) Then
            bt = meta(nm)(1)
            bd = meta(nm)(2)
        End If
        a(i, 1) = i
        a(i, 2) = nm
        a(i, 3) = c("kind")
        a(i, 4) = c("op")
        a(i, 5) = bt
        a(i, 7) = bd
        a(i, 9) = c("bc")
        If c("kind") <> "削除" Then
            a(i, 6) = c("ftype")
            a(i, 8) = c("fdesc")
            a(i, 10) = c("ac")
        End If
        If c("nAdd") > 0 Then a(i, 11) = "+" & c("nAdd")
        If c("nDel") > 0 Then a(i, 12) = "-" & c("nDel")
        Select Case c("kind")
            Case "新規": PaintAddRange paint, CL_NEW, ws.Range(ws.Cells(r, 1), ws.Cells(r, DL_COLS - 1))
            Case "変更"
                PaintAddRange paint, CL_MOD, ws.Range(ws.Cells(r, 1), ws.Cells(r, DL_COLS - 1))
                If c("ftype") <> bt Then PaintAdd paint, CL_HI, ws, r, 6
                If c("fdesc") <> bd Then PaintAdd paint, CL_HI, ws, r, 8
                If c("ac") <> c("bc") Then PaintAdd paint, CL_HI, ws, r, 10
                If c("nAdd") > 0 Then PaintAdd paint, CL_HI, ws, r, 11
                If c("nDel") > 0 Then PaintAdd paint, CL_HI, ws, r, 12
            Case "削除"
                PaintAddRange paint, CL_DEL, ws.Range(ws.Cells(r, 1), ws.Cells(r, DL_COLS - 1))
                PaintAddRange paint, FX_STRIKE, ws.Range(ws.Cells(r, 2), ws.Cells(r, 2))
        End Select
    Next
    ws.Range(ws.Cells(DL_FIRST, 1), ws.Cells(DL_FIRST + n - 1, DL_COLS)).NumberFormat = "@"
    ws.Range(ws.Cells(DL_FIRST, 1), ws.Cells(DL_FIRST + n - 1, 1)).NumberFormat = "General"
    ws.Range(ws.Cells(DL_FIRST, 9), ws.Cells(DL_FIRST + n - 1, 10)).NumberFormat = "General"
    With ws.Range(ws.Cells(DL_FIRST, 1), ws.Cells(DL_FIRST + n - 1, DL_COLS))
        .Value = a
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Color = vbBlack
    End With
    ApplyPaint ws, paint
End Sub

'----------------------------------------------------------
' 差分チェック（リスト内容）：横長（1列1リスト）
'   5行目 name、6行目 判定、7行目 件数（前 → 後、+追加/-削除）
'   8行目以降：既存の値（SCMの順。削除されるものはグレーと取り消し線）→ 追加される値（緑）
'----------------------------------------------------------
Private Sub RenderContentDiff(ByVal diffs As Collection, ByVal chk As Object, ByVal have As Object, ByVal status As String)
    Dim ws As Worksheet, hdr As Object, ngc As Object, byName As Object, order As Collection, c As Object
    Dim paint As Object, blocks As Collection, a() As Variant, k As Variant, x As Variant
    Dim nC As Long, mx As Long, n As Long, j As Long, i As Long, col As Long, i0 As Long, nm As String, s As String, clr As Long
    Set ws = Worksheets(SH_DCONT)
    Set hdr = chk("hdr")
    Set ngc = chk("ngcols")
    ' 並び：登録更新（リスト内容）の列順 → その列が無いリスト（名前順）
    Set byName = CreateObject("Scripting.Dictionary")
    For Each c In diffs
        Set byName(c("name")) = c
    Next
    Set order = New Collection
    For Each k In hdr.Keys
        If byName.Exists(k) Then order.Add byName(k)
    Next
    For Each c In diffs
        If Not hdr.Exists(c("name")) Then order.Add c
    Next
    ws.Range(ws.Columns(2), ws.Columns(ws.Columns.Count)).EntireColumn.Hidden = False
    ws.Range(ws.Cells(DC_NAME, 2), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear
    ws.Range("B3").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Range("B3").Value = Now
    ws.Range("B4").Value = status
    nC = order.Count
    If nC = 0 Then Exit Sub
    For Each c In order
        n = c("nAdd")
        If have.Exists(c("name")) Then n = n + have(c("name")).Count
        If n > mx Then mx = n
    Next
    ReDim a(1 To DC_FIRST - DC_NAME + mx, 1 To nC)
    Set paint = CreateObject("Scripting.Dictionary")
    Set blocks = New Collection                      ' 追加される値（列の下にまとまる）：Array(開始行, 終了行, 列)
    For Each c In order
        j = j + 1
        col = j + 1
        nm = c("name")
        a(1, j) = nm
        If hdr.Exists(nm) And ngc.Exists(hdr(nm)) Then
            a(DC_KIND - DC_NAME + 1, j) = "NG：反映しません（理由は登録更新（リスト内容）のコメント）"
            PaintAdd paint, CL_CHK, ws, DC_KIND, col
        Else
            a(DC_KIND - DC_NAME + 1, j) = c("kind")
            clr = 0
            Select Case c("kind")
                Case "新規": clr = CL_NEW
                Case "変更": clr = CL_MOD
                Case "削除": clr = CL_DEL
            End Select
            If clr <> 0 Then PaintAdd paint, clr, ws, DC_KIND, col
            If c("kind") = "削除" Then PaintAdd paint, FX_STRIKE, ws, DC_NAME, col
        End If
        s = c("bc") & " → " & c("ac") & " 件"
        If c("nAdd") > 0 Or c("nDel") > 0 Then s = s & "（+" & c("nAdd") & " / -" & c("nDel") & "）"
        a(DC_CNT - DC_NAME + 1, j) = s
        i = DC_FIRST - DC_NAME
        If have.Exists(nm) Then
            For Each k In have(nm).Keys
                i = i + 1
                a(i, j) = k
                If c("dels").Exists(k) Then
                    PaintAdd paint, CL_DEL, ws, DC_NAME + i - 1, col
                    PaintAdd paint, FX_STRIKE, ws, DC_NAME + i - 1, col
                End If
            Next
        End If
        If c("nAdd") > 0 Then
            i0 = i + 1
            For Each k In c("adds").Keys
                i = i + 1
                a(i, j) = k
            Next
            blocks.Add Array(DC_NAME + i0 - 1, DC_NAME + i - 1, col)
        End If
    Next
    With ws.Range(ws.Cells(DC_NAME, 2), ws.Cells(DC_NAME + UBound(a, 1) - 1, nC + 1))
        .NumberFormat = "@"
        .Value = a
        .Font.Color = vbBlack
        .VerticalAlignment = xlTop
    End With
    With ws.Range(ws.Cells(DC_NAME, 2), ws.Cells(DC_CNT, nC + 1))
        .Interior.Color = CL_SEC
        .Borders.LineStyle = xlContinuous
        .WrapText = True
    End With
    With ws.Range(ws.Cells(DC_NAME, 2), ws.Cells(DC_NAME, nC + 1))
        .Interior.Color = CL_HEAD
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    ws.Range(ws.Columns(2), ws.Columns(nC + 1)).ColumnWidth = 30
    For Each x In blocks
        PaintRange ws.Range(ws.Cells(x(0), x(2)), ws.Cells(x(1), x(2))), CL_NEW
    Next
    ApplyPaint ws, paint
End Sub

' 差分チェック（リスト一覧）：変化なしの行を隠す（実行結果が要確認の行は隠さない）
Public Sub HideUnchanged()
    Dim ws As Worksheet, lastR As Long, r As Long, v As Variant
    Set ws = Worksheets(SH_DLIST)
    lastR = LastRow(ws, 2)
    If lastR < DL_FIRST Then Exit Sub
    v = ws.Range(ws.Cells(DL_FIRST, 1), ws.Cells(lastR + 1, DL_COLS)).Value
    Application.ScreenUpdating = False
    For r = 1 To lastR - DL_FIRST + 1
        If SafeStr(v(r, 3)) = "変化なし" And Left$(SafeStr(v(r, DL_RES)), 3) <> "要確認" Then ws.Rows(DL_FIRST + r - 1).Hidden = True
    Next
    Application.ScreenUpdating = True
End Sub

Public Sub ShowAllDiff()
    With Worksheets(SH_DLIST)
        .Rows(DL_FIRST & ":" & .Rows.Count).Hidden = False
    End With
End Sub

' 差分チェック（リスト内容）：変化なしの列を隠す
Public Sub HideUnchangedCont()
    Dim ws As Worksheet, lastC As Long, c As Long, v As Variant
    Set ws = Worksheets(SH_DCONT)
    lastC = ws.Cells(DC_NAME, ws.Columns.Count).End(xlToLeft).Column
    If lastC < 2 Then Exit Sub
    v = ws.Range(ws.Cells(DC_KIND, 2), ws.Cells(DC_KIND + 1, lastC)).Value
    Application.ScreenUpdating = False
    For c = 2 To lastC
        If SafeStr(v(1, c - 1)) = "変化なし" Then ws.Columns(c).Hidden = True
    Next
    Application.ScreenUpdating = True
End Sub

Public Sub ShowAllCont()
    With Worksheets(SH_DCONT)
        .Range(.Columns(2), .Columns(.Columns.Count)).EntireColumn.Hidden = False
    End With
End Sub

'----------------------------------------------------------
' リストごとの計画
'----------------------------------------------------------
' すべてのカテゴリの計画（PlanCategory）を名前順に
Private Function PlanAll(ByVal meta As Object, ByVal have As Object, ByVal cats As Object, ByVal want As Object) As Collection
    Dim diffs As New Collection, names As Variant, i As Long, c As Object
    names = SortedKeys(UnionNames(meta, cats, Nothing))
    For i = LBound(names) To UBound(names)
        Set c = PlanCategory(CStr(names(i)), meta, have, cats, want)
        If Not c Is Nothing Then diffs.Add c
    Next
    Set PlanAll = diffs
End Function

' 計画の件数（メッセージ用）
Private Function DiffSummary(ByVal diffs As Collection) As String
    Dim c As Object, nNew As Long, nMod As Long, nDel As Long, nAddE As Long, nDelE As Long
    For Each c In diffs
        Select Case c("kind")
            Case "新規": nNew = nNew + 1
            Case "変更": nMod = nMod + 1
            Case "削除": nDel = nDel + 1
        End Select
        nAddE = nAddE + c("nAdd")
        nDelE = nDelE + c("nDel")
    Next
    DiffSummary = "リスト：新規 " & nNew & " 件、変更 " & nMod & " 件、削除 " & nDel & " 件" & vbLf & _
                  "　値：追加 " & nAddE & " 件／削除 " & nDelE & " 件"
End Function

' 1カテゴリの計画。Dictionary を返す（対象外は Nothing）
'   name, kind(新規/変更/削除/変化なし), op, method, url, body, bc, ac, nAdd, nDel, meta, ftype, fdesc,
'   adds/dels（エントリ -> 実行後の結果）, http, chk, noRow, row（横長ビューの列。並べ替え用）
Private Function PlanCategory(ByVal nm As String, ByVal meta As Object, ByVal have As Object, _
                              ByVal cats As Object, ByVal want As Object) As Object
    Dim c As Object, cr As Variant, before As Object, final As Object, k As Variant, op As String
    Dim typ As String, desc As String, exists As Boolean, mc As String
    exists = meta.Exists(nm)
    If exists Then
        If have.Exists(nm) Then Set before = have(nm) Else Set before = CreateObject("Scripting.Dictionary")
    Else
        Set before = CreateObject("Scripting.Dictionary")
    End If
    Set c = NewDiffItem(nm)
    c("bc") = before.Count
    If cats.Exists(nm) Then
        cr = cats(nm)
        op = cr(0)
        c("row") = cr(4)
    End If
    If op = "削除" Then
        If Not exists Then Exit Function
        c("kind") = "削除": c("op") = "削除": c("method") = "DELETE"
        c("url") = BaseUrl() & "/" & meta(nm)(0)
        c("ftype") = meta(nm)(1): c("fdesc") = meta(nm)(2)
        For Each k In before.Keys
            c("dels").Add k, ""
        Next
        c("ac") = 0
    ElseIf op = "登録" Then
        ' 登録：列に値があればその値で、無ければダミー1件で作成（URLリストは空にできないため）
        If want.Exists(nm) Then
            If want(nm).Count > 0 Then Set final = want(nm)
        End If
        If final Is Nothing Then
            Set final = CreateObject("Scripting.Dictionary")
            final.Add DummyEntry(), True
        End If
        c("kind") = "新規": c("op") = "登録": c("method") = "POST"
        c("url") = WithQuery(BaseUrl(), ContainerQuery())
        typ = cr(1): desc = cr(2)
    ElseIf Not exists Then
        Exit Function
    Else
        If want.Exists(nm) Then
            Set final = want(nm)
        Else
            Set final = before
            c("noRow") = (op <> "変更" And before.Count > 0)
        End If
        If op = "変更" Then
            typ = cr(1): desc = cr(2)
            c("kind") = "変更": c("op") = "変更"
        Else
            typ = meta(nm)(1): desc = meta(nm)(2)
            If SameSet(before, final) Then
                c("kind") = "変化なし"
            Else
                c("kind") = "変更": c("op") = "変更(URL)"
            End If
        End If
        If c("op") <> "" Then
            c("method") = "PUT"
            c("url") = BaseUrl() & "/" & meta(nm)(0)
        End If
        If op = "変更" Then
            If typ <> meta(nm)(1) Then mc = "type：" & meta(nm)(1) & " → " & typ
            If desc <> meta(nm)(2) Then
                If mc <> "" Then mc = mc & vbLf
                If desc = "" Then mc = mc & "description：削除" Else mc = mc & "description：変更"
            End If
            If mc = "" Then mc = "なし"
        End If
        If c("noRow") Then mc = "リスト内容に値が無いため中身は変更しません"
    End If
    If c("kind") <> "削除" Then
        For Each k In final.Keys
            If Not before.Exists(k) Then c("adds").Add k, ""
        Next
        For Each k In before.Keys
            If Not final.Exists(k) Then c("dels").Add k, ""
        Next
        c("ac") = final.Count
        c("ftype") = typ: c("fdesc") = desc
        If c("method") <> "" Then c("body") = BuildBodyParts(nm, typ, ListJsonDict(final), desc)
    End If
    c("nAdd") = c("adds").Count
    c("nDel") = c("dels").Count
    c("meta") = mc
    Set PlanCategory = c
End Function

Private Function NewDiffItem(ByVal nm As String) As Object
    Dim c As Object
    Set c = CreateObject("Scripting.Dictionary")
    c("name") = nm: c("kind") = "変化なし": c("op") = "": c("method") = "": c("url") = "": c("body") = ""
    c("bc") = 0: c("ac") = 0: c("nAdd") = 0: c("nDel") = 0: c("meta") = "": c("ftype") = "": c("fdesc") = ""
    c("http") = "": c("chk") = "": c("noRow") = False: c("row") = 100000
    Set c("adds") = CreateObject("Scripting.Dictionary")
    Set c("dels") = CreateObject("Scripting.Dictionary")
    Set NewDiffItem = c
End Function

' JSON_CURLに書く順（横長ビューの列順。URLだけの変更は row が無いので最後に名前順）
Private Function SortByRow(ByVal c As Collection) As Collection
    Dim res As New Collection, a() As Object, i As Long, j As Long, t As Object
    If c.Count = 0 Then
        Set SortByRow = res
        Exit Function
    End If
    ReDim a(1 To c.Count)
    For i = 1 To c.Count
        Set a(i) = c(i)
    Next
    For i = 2 To c.Count                ' 安定な挿入ソート（件数は少ない）
        Set t = a(i)
        j = i - 1
        Do While j >= 1
            If a(j)("row") <= t("row") Then Exit Do
            Set a(j + 1) = a(j)
            j = j - 1
        Loop
        Set a(j + 1) = t
    Next
    For i = 1 To c.Count
        res.Add a(i)
    Next
    Set SortByRow = res
End Function

Private Function ListJsonDict(ByVal d As Object) As String
    Dim arr() As String, k As Variant, i As Long
    If d.Count = 0 Then
        ListJsonDict = "[]"
        Exit Function
    End If
    ReDim arr(1 To d.Count)
    For Each k In d.Keys
        i = i + 1
        arr(i) = JsonStr(CStr(k))
    Next
    ListJsonDict = "[" & Join(arr, ",") & "]"
End Function

Private Function SameSet(ByVal a As Object, ByVal b As Object) As Boolean
    Dim k As Variant
    If a.Count <> b.Count Then Exit Function
    For Each k In a.Keys
        If Not b.Exists(k) Then Exit Function
    Next
    SameSet = True
End Function

' 複数の Dictionary のキーを1つに
Private Function UnionNames(ByVal a As Object, ByVal b As Object, ByVal c As Object) As Object
    Dim d As Object, x As Variant, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each x In Array(a, b, c)
        If Not x Is Nothing Then
            For Each k In x.Keys
                If Not d.Exists(k) Then d.Add k, True
            Next
        End If
    Next
    Set UnionNames = d
End Function

' キーを並べ替えた配列（大文字・小文字を区別しない順）。空なら要素0の配列
Private Function SortedKeys(ByVal d As Object) As Variant
    Dim a() As String, k As Variant, i As Long
    If d.Count = 0 Then
        SortedKeys = Array()
        Exit Function
    End If
    ReDim a(0 To d.Count - 1)
    For Each k In d.Keys
        a(i) = CStr(k)
        i = i + 1
    Next
    QSort a, 0, d.Count - 1
    SortedKeys = a
End Function

Private Sub QSort(ByRef a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pv As String, t As String
    i = lo: j = hi
    pv = a((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(a(i), pv, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(a(j), pv, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QSort a, lo, j
    If i < hi Then QSort a, i, hi
End Sub

'----------------------------------------------------------
' JSON_CURL
'----------------------------------------------------------
Private Function BuildBodyParts(ByVal nm As String, ByVal typ As String, ByVal listJs As String, ByVal desc As String) As String
    Dim js As String
    js = "{""name"":" & JsonStr(nm) & ",""type"":" & JsonStr(typ) & ",""list"":" & listJs
    If desc <> "" Then js = js & ",""description"":" & JsonStr(desc)
    BuildBodyParts = js & "}"
End Function

' JSON_CURLに1行書き込み、o を次の行へ進める。長いJSONはJ列以降に分割して保存
Private Sub WriteJsonRow(ByVal wsOut As Worksheet, ByRef o As Long, ByVal op As String, ByVal nm As String, _
                         ByVal method As String, ByVal url As String, ByVal body As String)
    Dim rowNo As Long, nParts As Long, i As Long
    rowNo = o - JS_FIRST + 1
    wsOut.Range(wsOut.Cells(o, 2), wsOut.Cells(o, JS_LASTCOL)).NumberFormat = "@"
    wsOut.Cells(o, 1).Value = rowNo
    wsOut.Cells(o, 2).Value = op
    wsOut.Cells(o, 3).Value = nm
    wsOut.Cells(o, 4).Value = method
    wsOut.Cells(o, 5).Value = url
    wsOut.Cells(o, 6).Value = Left$(body, CELL_MAX)
    If Len(body) > CELL_MAX Then
        nParts = (Len(body) - 1) \ CELL_MAX
        For i = 1 To nParts
            wsOut.Cells(o, JS_PARTS + i).Value = Mid$(body, i * CELL_MAX + 1, CELL_MAX)
        Next
        wsOut.Cells(o, JS_PARTS).Value = nParts
        wsOut.Cells(o, 7).Value = BuildCurlFile(method, url, "body_" & rowNo & ".json")
    Else
        wsOut.Cells(o, 7).Value = BuildCurl(method, url, body)
    End If
    o = o + 1
End Sub

' JSON_CURLの1行のJSON全文（分割分をつなげる）
Private Function GetBody(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim nParts As Long, i As Long, s As String
    s = CStr(ws.Cells(r, 6).Value)
    nParts = Val(CStr(ws.Cells(r, JS_PARTS).Value))
    For i = 1 To nParts
        s = s & CStr(ws.Cells(r, JS_PARTS + i).Value)
    Next
    GetBody = s
End Function

Private Function BuildCurlFile(ByVal method As String, ByVal url As String, ByVal fileName As String) As String
    BuildCurlFile = "# JSONが長いため「JSONファイル出力」で作成した " & fileName & " を使います" & vbLf & _
                    "curl -s -X " & method & " '" & url & "' -H ""Authorization: Bearer $TOKEN""" & _
                    " -H 'Content-Type: application/json' --data-binary @" & fileName
End Function

Public Sub ExecuteAll()
    Dim ws As Worksheet, r As Long, n As Long, st As Long, resText As String, ok As Long, ng As Long, nChk As Long
    Set ws = Worksheets(SH_JSON)
    n = LastRow(ws, 4) - JS_FIRST + 1
    If n <= 0 Then
        MsgBox "実行対象がありません。先に「JSON発行」を実行してください。", vbExclamation
        Exit Sub
    End If
    If MsgBox(n & " 件をSCMへ実行します（元に戻せません）。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    For r = JS_FIRST To JS_FIRST + n - 1
        st = HttpCall(CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), GetBody(ws, r), resText)
        ws.Cells(r, 8).Value = st
        If st >= 200 And st < 300 Then ok = ok + 1 Else ng = ng + 1
        WriteLog CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), st, ErrSummary(resText)
    Next
    ' 実行後のSCMの状態を取得し、JSON_CURLのとおりになったかを確認
    nChk = BuildDiffCore()
    Worksheets(SH_DLIST).Activate
    MsgBox "成功 " & ok & " 件／失敗 " & ng & " 件" & vbLf & _
           "反映結果を差分チェック（リスト一覧）のM列「実行結果」に出力しました。" & _
           IIf(nChk > 0, vbLf & "要確認（赤）が " & nChk & " 件あります。", "") & vbLf & _
           "要確認が無いことを確認してから、SCMの画面でPushしてください。" & vbLf & _
           "次の作業の前に「既存情報取得」で既存一覧・登録更新シートを最新にしてください。", IIf(nChk > 0, vbExclamation, vbInformation)
End Sub

' JSON_CURLの各行のJSONを UTF-8 のファイル（body_No.json）に書き出す
Public Sub ExportJsonFiles()
    Dim ws As Worksheet, r As Long, n As Long, outDir As String, body As String
    Set ws = Worksheets(SH_JSON)
    If ThisWorkbook.Path = "" Then
        MsgBox "先にこのブックを保存してください（ブックと同じフォルダに出力します）。", vbExclamation
        Exit Sub
    End If
    If LastRow(ws, 4) < JS_FIRST Then
        MsgBox "出力対象がありません。", vbExclamation
        Exit Sub
    End If
    outDir = ThisWorkbook.Path & Application.PathSeparator & "json_" & Format$(Now, "yyyymmdd_hhnnss")
    MkDir outDir
    For r = JS_FIRST To LastRow(ws, 4)
        body = GetBody(ws, r)
        If body <> "" Then
            WriteUtf8 outDir & Application.PathSeparator & "body_" & ws.Cells(r, 1).Value & ".json", body
            n = n + 1
        End If
    Next
    MsgBox n & " 件のJSONファイルを出力しました。" & vbLf & outDir & vbLf & _
           "curlはこのフォルダに移動してから実行してください。", vbInformation
End Sub

Private Sub WriteUtf8(ByVal filePath As String, ByVal txt As String)
    Dim st As Object, bin As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.Position = 3                     ' BOMを除く
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    st.CopyTo bin
    bin.SaveToFile filePath, 2
    bin.Close
    st.Close
End Sub

'----------------------------------------------------------
' クリア
'----------------------------------------------------------
Public Sub ClearJson()
    ClearJsonRows Worksheets(SH_JSON)
End Sub

Private Sub ClearJsonRows(ByVal ws As Worksheet)
    ws.Range(ws.Cells(JS_FIRST, 1), ws.Cells(ws.Rows.Count, JS_LASTCOL)).Clear
End Sub

'==========================================================
' 実行後の確認：JSON_CURLの内容を SCM の最新状態と比較
'==========================================================
Public Sub BuildDiffCheck()
    Dim n As Long
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    n = BuildDiffCore()
    If n >= 0 Then
        MsgBox "差分を確認しました。要確認 " & n & " 件（リスト）。" & IIf(n > 0, vbLf & "差分チェック（リスト一覧）のM列の赤と実行ログを確認してください。", ""), _
               IIf(n > 0, vbExclamation, vbInformation)
    End If
End Sub

' 戻り値：要確認のカテゴリ数（取得失敗時は -1）
Private Function BuildDiffCore() As Long
    Dim items As Collection
    BuildDiffCore = -1
    Set items = FetchAllItems("差分確認")
    If items Is Nothing Then Exit Function
    Application.ScreenUpdating = False
    BuildDiffCore = CheckAfterExec(items)
    Application.ScreenUpdating = True
    WriteLog "差分確認", "", "GET", WithQuery(BaseUrl(), ContainerQuery()), 200, "要確認 " & BuildDiffCore & " 件"
End Function

' API実行後の確認：JSON_CURLの各行が SCM の現在に反映されたかを、
'   差分チェック（リスト一覧）のM列「実行結果」（その行が無いリストは実行ログのみ）に書く。戻り値は要確認の件数
Private Function CheckAfterExec(ByVal items As Collection) As Long
    Dim wsJ As Worksheet, ws As Worksheet, r As Long, op As String, nm As String, method As String, st As String
    Dim body As Object, item As Object, aSet As Object, eSet As Object, v As Variant, k As Variant
    Dim hdr As Object, paint As Object, lastR As Long, c As Long, i As Long, msg As String, nMiss As Long, nExtra As Long, cnt As Long
    Dim res() As Variant
    Set wsJ = Worksheets(SH_JSON)
    Set ws = Worksheets(SH_DLIST)
    Set hdr = CreateObject("Scripting.Dictionary")              ' リスト名 -> 行
    Set paint = CreateObject("Scripting.Dictionary")
    lastR = LastRow(ws, 2)
    If lastR >= DL_FIRST Then
        v = ws.Range(ws.Cells(DL_FIRST, 2), ws.Cells(lastR + 1, 2)).Value     ' 1行多く読んで必ず2次元配列にする
        For i = 1 To lastR - DL_FIRST + 1
            nm = Trim$(SafeStr(v(i, 1)))
            If nm <> "" And Not hdr.Exists(nm) Then hdr.Add nm, DL_FIRST + i - 1
        Next
        ReDim res(1 To lastR - DL_FIRST + 1, 1 To 1)
    End If
    For r = JS_FIRST To LastRow(wsJ, 4)
        op = SafeStr(wsJ.Cells(r, 2).Value)
        nm = SafeStr(wsJ.Cells(r, 3).Value)
        If op <> "" Then
            method = SafeStr(wsJ.Cells(r, 4).Value)
            st = SafeStr(wsJ.Cells(r, 8).Value)
            Set item = FindItem(items, nm)
            msg = ""
            If method = "DELETE" Then
                If item Is Nothing Then msg = "OK：削除済" Else msg = "要確認：削除されていません"
            ElseIf item Is Nothing Then
                msg = "要確認：SCMにありません"
            Else
                Set body = ParseJson(GetBody(wsJ, r))
                Set eSet = CreateObject("Scripting.Dictionary")
                Set aSet = CreateObject("Scripting.Dictionary")
                If body.Exists("list") Then
                    For Each v In body("list")
                        eSet(CStr(v)) = True
                    Next
                End If
                If ListCount(item) > 0 Then
                    For Each v In item("list")
                        aSet(CStr(v)) = True
                    Next
                End If
                nMiss = 0: nExtra = 0
                For Each k In eSet.Keys
                    If Not aSet.Exists(k) Then nMiss = nMiss + 1
                Next
                For Each k In aSet.Keys
                    If Not eSet.Exists(k) Then nExtra = nExtra + 1
                Next
                If nMiss > 0 Then msg = msg & "未反映 " & nMiss & " 件／"
                If nExtra > 0 Then msg = msg & "想定外の値 " & nExtra & " 件／"
                If Nz(item, "type") <> Nz(body, "type") Then msg = msg & "typeが想定と異なる（" & Nz(item, "type") & "）／"
                If Nz(item, "description") <> Nz(body, "description") Then msg = msg & "descriptionが想定と異なる／"
                If msg = "" Then
                    msg = "OK：反映済（" & aSet.Count & " 件）"
                Else
                    msg = "要確認：" & Left$(msg, Len(msg) - 1)
                End If
            End If
            If Left$(msg, 3) = "要確認" Then
                If st <> "" Then msg = msg & "（HTTP " & st & "）"
                cnt = cnt + 1
                WriteLog "差分確認", nm, "", "URLカテゴリ", 0, msg
            End If
            If hdr.Exists(nm) Then
                c = hdr(nm)
                res(c - DL_FIRST + 1, 1) = msg
                PaintAdd paint, IIf(Left$(msg, 2) = "OK", CL_NEW, CL_CHK), ws, c, DL_RES
            End If
        End If
    Next
    If lastR >= DL_FIRST Then
        With ws.Range(ws.Cells(DL_FIRST, DL_RES), ws.Cells(lastR, DL_RES))
            .Interior.Pattern = xlNone
            .Value = res
        End With
        ApplyPaint ws, paint
    End If
    Worksheets(SH_DCONT).Range("B4").Value = "実行後の確認（結果は差分チェック（リスト一覧）のM列）"
    ws.Range("B4").Value = "実行後の確認（SCMから再取得：" & Format$(Now, "yyyy/mm/dd hh:nn:ss") & "）"
    CheckAfterExec = cnt
End Function

'----------------------------------------------------------
' API・共通の小物
'----------------------------------------------------------
' 対象フォルダのカスタムURLカテゴリを全件取得（ページ送り）。失敗時は Nothing
Private Function FetchAllItems(ByVal opLabel As String) As Collection
    Dim c As New Collection, url As String, resText As String, st As Long, res As Object, item As Object
    Dim offset As Long, cnt As Long, total As Long
    offset = 0
    Do
        url = WithQuery(BaseUrl(), ContainerQuery() & "&limit=" & PAGE_LIMIT & "&offset=" & offset)
        st = HttpCall("GET", url, "", resText)
        If st <> 200 Then
            WriteLog opLabel, "", "GET", url, st, ErrSummary(resText)
            Application.ScreenUpdating = True
            MsgBox "カスタムURLカテゴリの取得に失敗しました（HTTP " & st & "）。実行ログを確認してください。", vbExclamation
            Exit Function
        End If
        Set res = ParseJson(resText)
        cnt = 0
        For Each item In res("data")
            c.Add item
            cnt = cnt + 1
        Next
        offset = offset + PAGE_LIMIT
        If res.Exists("total") Then total = res("total") Else total = 0
        If cnt < PAGE_LIMIT Or offset >= total Then Exit Do
    Loop
    Set FetchAllItems = c
End Function

' 名前が完全一致（大文字・小文字を区別）するカテゴリ
Private Function FindItem(ByVal items As Collection, ByVal nm As String) As Object
    Dim item As Object
    For Each item In items
        If StrComp(Nz(item, "name"), nm, vbBinaryCompare) = 0 Then
            Set FindItem = item
            Exit Function
        End If
    Next
End Function

Private Function JoinList(ByVal item As Object, ByVal sep As String) As String
    Dim arr() As String, v As Variant, i As Long
    If ListCount(item) = 0 Then Exit Function
    ReDim arr(1 To item("list").Count)
    For Each v In item("list")
        i = i + 1
        arr(i) = CStr(v)
    Next
    JoinList = Join(arr, sep)
End Function

Private Function ListCount(ByVal item As Object) As Long
    If Not item.Exists("list") Then Exit Function
    If Not IsObject(item("list")) Then Exit Function
    ListCount = item("list").Count
End Function

' セルに入りきらない長さは省略して表示
Private Function CellText(ByVal s As String) As String
    If Len(s) > CELL_MAX Then
        CellText = Left$(s, CELL_MAX) & vbLf & "…（長いため以下省略）"
    Else
        CellText = s
    End If
End Function

' ホワイトリスト名とブラックリスト名のチェック
Private Function CheckListNames() As Boolean
    If Left$(Cfg("CFG_LIST_CHECK"), 2) = "NG" Then
        MsgBox "基本設定のホワイトリスト／ブラックリスト名を確認してください。" & vbLf & Cfg("CFG_LIST_CHECK"), vbExclamation
        Exit Function
    End If
    CheckListNames = True
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

' 基本設定のホワイトリスト／ブラックリスト名か（大文字・小文字を区別）
Private Function IsListName(ByVal nm As String) As Boolean
    Dim lk As Variant, s As String
    If nm = "" Then Exit Function
    For Each lk In Array("WL", "BL")
        s = Cfg("CFG_" & lk & "_NAME")
        If s <> "" And StrComp(s, nm, vbBinaryCompare) = 0 Then
            IsListName = True
            Exit Function
        End If
    Next
End Function

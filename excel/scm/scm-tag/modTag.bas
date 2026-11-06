Attribute VB_Name = "modTag"
'==========================================================
' modTag : SCM (Prisma Access Configuration API) Tag 連携
'   前提 : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'==========================================================
' ※ 共通処理（認証・HTTP・ログ・差分色付け等）は modScmCommon にあります。modScmCommon（JSONパーサを含む）も併せてインポートしてください。
Option Explicit

Private Const SH_CFG As String = "基本設定"
Private Const SH_EXIST As String = "Tag_既存一覧"
Private Const SH_INPUT As String = "Tag_登録変更削除"
Private Const SH_JSON As String = "JSON_CURL"
Private Const SH_DIFF As String = "Tag_差分確認"
Private Const SH_LOG As String = "実行ログ"
Private Const SH_COLOR As String = "Tag_色一覧"

Private Const EX_FIRST As Long = 8      ' 既存一覧 データ開始行
Private Const IN_FIRST As Long = 8      ' 登録変更削除 データ開始行
Private Const IN_LAST As Long = 507     ' 登録変更削除 データ最終行
Private Const IN_JUDGE As Long = 13     ' ③ 判定列(M)
Private Const JS_FIRST As Long = 10     ' JSON_CURL データ開始行
Private Const DIFF_FIRST As Long = 8    ' 差分確認 データ開始行
Private Const COLOR_FIRST As Long = 8   ' 色一覧 データ開始行
Private Const PAGE_LIMIT As Long = 200

' ③ Tag_登録変更削除 の列
'   B=操作 C=name D=color E=色見本 F=comments H=既存id
' ② Tag_既存一覧 の列
'   B=id C=name D=color E=色見本 F=comments G=コンテナ
' ⑦ Tag_色一覧 の列
'   B=color C=色見本（セルの塗りつぶし色を②の色見本に使う）

'----------------------------------------------------------
' ボタン配置（初回のみ実行）
'----------------------------------------------------------
Public Sub SetupButtons()
    AddBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    AddBtn SH_EXIST, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUT, "F3:G4", "JSON発行 → ④へ", "GenerateJson"
    AddBtn SH_INPUT, "I3:J4", "入力クリア", "ClearInput"
    AddBtn SH_JSON, "G4:H4", "API実行（全行）", "ExecuteAll"
    AddBtn SH_JSON, "G6:H6", "クリア", "ClearJson"
    RemoveBtn SH_JSON, "PushConfig"     ' Push機能は削除（PushはSCM画面から実施）。旧ボタンを消す
    AddBtn SH_DIFF, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_LOG, "F3", "ログクリア", "ClearLog"
    RemoveBtn SH_LOG, "CheckJobs"      ' ジョブ状況確認も削除。旧ボタンを消す
    MsgBox "ボタンを配置しました。", vbInformation
End Sub

'----------------------------------------------------------
' ② 既存情報取得（Tag を GET → 表生成）
'----------------------------------------------------------
Public Sub FetchExisting()
    Dim tags As Collection
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Set tags = FetchAllItems(BaseUrl(), "既存取得", True)
    If tags Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    WriteExistTag tags
    Application.ScreenUpdating = True
    MsgBox "Tag " & tags.Count & " 件を取得しました。", vbInformation
End Sub

Private Sub WriteExistTag(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long, colors As Object, c As String
    Set ws = Worksheets(SH_EXIST)
    Set colors = ColorSamples()
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 7)).Clear
    r = EX_FIRST
    For Each item In items
        c = Nz(item, "color")
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        ws.Cells(r, 4).Value = c
        If colors.Exists(c) Then ws.Cells(r, 5).Interior.Color = colors(c)
        ws.Cells(r, 6).Value = Nz(item, "comments")
        ws.Cells(r, 7).Value = ContainerOf(item)
        r = r + 1
    Next
    If r > EX_FIRST Then
        With ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 7))
            .Borders.LineStyle = xlContinuous
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With
    End If
    Range("EX_FETCHED_AT").Value = Now
End Sub

' ⑦色一覧から color名 → 色見本の塗りつぶし色 の対応表（大文字・小文字を区別）
Private Function ColorSamples() As Object
    Dim ws As Worksheet, r As Long, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets(SH_COLOR)
    For r = COLOR_FIRST To LastRow(ws, 2)
        If CStr(ws.Cells(r, 2).Value) <> "" Then d(CStr(ws.Cells(r, 2).Value)) = ws.Cells(r, 3).Interior.Color
    Next
    Set ColorSamples = d
End Function

'----------------------------------------------------------
' ③ JSON発行（Tag_登録変更削除 → JSON_CURL）
'   実行順：登録 → 変更 → 削除
'----------------------------------------------------------
Public Sub GenerateJson()
    Dim wsIn As Worksheet, wsOut As Worksheet, r As Long, o As Long, ng As Long, op As Variant
    If Not CheckContainer() Then Exit Sub
    Application.Calculate
    Set wsIn = Worksheets(SH_INPUT)
    Set wsOut = Worksheets(SH_JSON)
    ng = Application.WorksheetFunction.CountIf(wsIn.Range(wsIn.Cells(IN_FIRST, IN_JUDGE), wsIn.Cells(IN_LAST, IN_JUDGE)), "NG")
    If ng > 0 Then
        If MsgBox("③に判定NGが " & ng & " 件あります。NG行を除外してJSONを発行しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    ClearJsonRows wsOut
    o = JS_FIRST
    For Each op In Array("登録", "変更", "削除")
        For r = IN_FIRST To IN_LAST
            If IsTarget(wsIn, r) And CStr(wsIn.Cells(r, 2).Value) = op Then WriteJsonRowTag wsIn, r, wsOut, o
        Next
    Next
    If o > JS_FIRST Then
        With wsOut.Range(wsOut.Cells(JS_FIRST, 1), wsOut.Cells(o - 1, 8))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    wsOut.Activate
    MsgBox "③から " & (o - JS_FIRST) & " 件のJSONを④に生成しました。", vbInformation
End Sub

' JSON発行の対象行か（nameあり、判定がOKか警告）
Private Function IsTarget(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim judge As String
    judge = CStr(ws.Cells(r, IN_JUDGE).Value)
    IsTarget = (CStr(ws.Cells(r, 3).Value) <> "" And (judge = "OK" Or judge = "警告"))
End Function

Private Sub WriteJsonRowTag(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long)
    Dim op As String, nm As String
    op = CStr(wsIn.Cells(r, 2).Value)
    nm = CStr(wsIn.Cells(r, 3).Value)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op, nm, "POST", WithQuery(BaseUrl(), ContainerQuery()), BuildBodyTag(wsIn, r)
        Case "変更"
            WriteJsonRow wsOut, o, op, nm, "PUT", BaseUrl() & "/" & wsIn.Cells(r, 8).Value, BuildBodyTag(wsIn, r)
        Case "削除"
            WriteJsonRow wsOut, o, op, nm, "DELETE", BaseUrl() & "/" & wsIn.Cells(r, 8).Value, ""
    End Select
End Sub

' ④に1行書き込み、o を次の行へ進める
Private Sub WriteJsonRow(ByVal wsOut As Worksheet, ByRef o As Long, ByVal op As String, ByVal nm As String, _
                         ByVal method As String, ByVal url As String, ByVal body As String)
    wsOut.Cells(o, 1).Value = o - JS_FIRST + 1
    wsOut.Cells(o, 2).Value = op
    wsOut.Cells(o, 3).Value = nm
    wsOut.Cells(o, 4).Value = method
    wsOut.Cells(o, 5).Value = url
    wsOut.Cells(o, 6).Value = body
    wsOut.Cells(o, 7).Value = BuildCurl(method, url, body)
    o = o + 1
End Sub

' Tag のボディ {"name":..,"color":..,"comments":..}（空欄の項目は省略）
Private Function BuildBodyTag(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim js As String
    js = "{""name"":" & JsonStr(CStr(ws.Cells(r, 3).Value))
    If CStr(ws.Cells(r, 4).Value) <> "" Then js = js & ",""color"":" & JsonStr(CStr(ws.Cells(r, 4).Value))
    If CStr(ws.Cells(r, 6).Value) <> "" Then js = js & ",""comments"":" & JsonStr(CStr(ws.Cells(r, 6).Value))
    BuildBodyTag = js & "}"
End Function

'----------------------------------------------------------
' ④ API実行（JSON_CURL の全行）
'----------------------------------------------------------
Public Sub ExecuteAll()
    Dim ws As Worksheet, r As Long, n As Long, st As Long, resText As String, ok As Long, ng As Long
    Set ws = Worksheets(SH_JSON)
    n = LastRow(ws, 4) - JS_FIRST + 1
    If n <= 0 Then
        MsgBox "実行対象がありません。③でJSON発行してください。", vbExclamation
        Exit Sub
    End If
    If MsgBox(n & " 件をSCMへ実行します（元に戻せません）。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    For r = JS_FIRST To JS_FIRST + n - 1
        st = HttpCall(CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), CStr(ws.Cells(r, 6).Value), resText)
        ws.Cells(r, 8).Value = st
        If st >= 200 And st < 300 Then ok = ok + 1 Else ng = ng + 1
        WriteLog CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), st, ErrSummary(resText)
    Next
    ' 実行後のSCMの状態を取得し、②既存一覧と比較して⑥差分確認に出力
    BuildDiffCore
    Worksheets(SH_DIFF).Activate
    MsgBox "成功 " & ok & " 件／失敗 " & ng & " 件" & vbLf & _
           "Tag_差分確認 に、現在の設定一覧と既存一覧（実行前）からの変化を出力しました。内容を確認してから、SCMの画面でPushしてください。", vbInformation
End Sub

'----------------------------------------------------------
' クリア系
'----------------------------------------------------------
Public Sub ClearInput()
    Dim ws As Worksheet
    If MsgBox("入力（B～D列・F列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Set ws = Worksheets(SH_INPUT)
    ws.Range("B" & IN_FIRST & ":D" & IN_LAST).ClearContents
    ws.Range("F" & IN_FIRST & ":F" & IN_LAST).ClearContents
End Sub

Public Sub ClearJson()
    ClearJsonRows Worksheets(SH_JSON)
End Sub

Private Sub ClearJsonRows(ByVal ws As Worksheet)
    ws.Range(ws.Cells(JS_FIRST, 1), ws.Cells(ws.Rows.Count, 8)).Clear
End Sub

'==========================================================
' ⑥ 差分確認（②既存一覧＝実行前 と SCMの最新状態＝実行後 を比較）
'==========================================================
Public Sub BuildDiffCheck()
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    If BuildDiffCore() Then
        Worksheets(SH_DIFF).Activate
        MsgBox "差分を作成しました。", vbInformation
    End If
End Sub

Private Function BuildDiffCore() As Boolean
    Dim before As Object, after As Object, ops As Object, items As Collection, item As Object
    Dim ws As Worksheet, r As Long, n As Long, opName As String
    Set before = CreateObject("Scripting.Dictionary")
    Set after = CreateObject("Scripting.Dictionary")
    Set ops = CreateObject("Scripting.Dictionary")
    Application.Calculate
    ' 実行前：②既存一覧シート
    Set ws = Worksheets(SH_EXIST)
    For r = EX_FIRST To LastRow(ws, 3)
        If CStr(ws.Cells(r, 3).Value) <> "" Then
            before("Tag|" & ws.Cells(r, 3).Value) = Array(CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 6).Value))
        End If
    Next
    ' 実行後：SCMから取得（シートは書き換えない）
    Set items = FetchAllItems(BaseUrl(), "差分確認", False)
    If items Is Nothing Then Exit Function
    For Each item In items
        after("Tag|" & Nz(item, "name")) = Array(Nz(item, "color"), Nz(item, "comments"))
    Next
    ' ④の操作
    Set ws = Worksheets(SH_JSON)
    For r = JS_FIRST To LastRow(ws, 4)
        opName = CStr(ws.Cells(r, 2).Value)
        If opName <> "" Then ops("Tag|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
    Next
    Application.ScreenUpdating = False
    n = WriteDiffSheet(Worksheets(SH_DIFF), "Tag", Array("color", "comments"), before, after, ops)
    Application.ScreenUpdating = True
    WriteLog "差分確認", "", "GET", BaseUrl(), 200, n & "件の変化"
    BuildDiffCore = True
End Function

' SCMの現在の全件と削除されたものを一覧表示し、既存一覧との変化を付ける。戻り値は変化のあった件数
'   列：A=No B=name C=変化 D～=各項目（現在） 続けて 変更項目 / 変更前 / ④の操作 / HTTP / 判定
Private Function WriteDiffSheet(ByVal ws As Worksheet, ByVal objKind As String, ByVal fieldNames As Variant, _
                                ByVal before As Object, ByVal after As Object, ByVal ops As Object) As Long
    Dim allKeys As Object, k As Variant, r As Long, i As Long, b As Variant, a As Variant, op As Variant
    Dim nf As Long, cF As Long, lastCol As Long
    Dim chg As String, fields As String, expected As String, judge As String, st As String, changed As Long
    Dim cntAll As Long, cntSame As Long, cntAdd As Long, cntMod As Long, cntDel As Long, cntChk As Long
    nf = UBound(fieldNames) - LBound(fieldNames) + 1
    cF = 4 + nf
    lastCol = cF + 4
    Set allKeys = CreateObject("Scripting.Dictionary")
    AddKeys allKeys, after, objKind      ' 現在の一覧（SCMの順）
    AddKeys allKeys, before, objKind     ' 削除されたもの（後ろに並ぶ）
    AddKeys allKeys, ops, objKind        ' ④にあるが両方に無いもの
    ' 行の色はVBAでセルに直接付ける（条件付き書式は優先されて上書きされるため、ここで削除する）
    With ws.Range(ws.Cells(DIFF_FIRST, 1), ws.Cells(ws.Rows.Count, lastCol))
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Font.Color = vbBlack
        .Font.Bold = False
        .Font.Strikethrough = False
        .FormatConditions.Delete
    End With
    r = DIFF_FIRST
    For Each k In allKeys.Keys
        chg = "": fields = ""
        If before.Exists(k) And after.Exists(k) Then
            b = before(k): a = after(k)
            For i = 0 To nf - 1
                If b(i) <> a(i) Then fields = fields & fieldNames(i) & " "
            Next
            If fields <> "" Then chg = "変更"
        ElseIf after.Exists(k) Then
            chg = "追加"
        ElseIf before.Exists(k) Then
            chg = "削除"
        End If
        expected = "": st = "": op = Empty
        If ops.Exists(k) Then
            op = ops(k)
            Select Case op(0)
                Case "登録": expected = "追加"
                Case "変更": expected = "変更"
                Case "削除": expected = "削除"
            End Select
            st = op(1)
        End If
        If expected = "" Then
            If chg = "" Then judge = "" Else judge = "想定外（④に無い変更）"
        ElseIf chg = expected Then
            judge = "想定どおり"
        ElseIf chg = "" Then
            If Left$(st, 1) = "2" And expected = "変更" Then
                judge = "変化なし（既存と同じ内容）"
            Else
                judge = "未反映（HTTP " & st & "）"
            End If
        Else
            judge = "想定と異なる"
        End If
        If chg <> "" Then changed = changed + 1
        If judge <> "" And judge <> "想定どおり" And Left$(judge, 4) <> "変化なし" Then cntChk = cntChk + 1
        ' 未反映の登録（SCMにも既存一覧にも無い）は一覧に出さずログに残す
        If Not after.Exists(k) And Not before.Exists(k) Then
            WriteLog "差分確認", Mid$(k, InStr(k, "|") + 1), "", objKind, 0, judge
            GoTo NextKey
        End If
        ' 削除されたものは既存一覧（実行前）の内容を表示する（想定外の削除は実行ログにも残す）
        If chg = "削除" Then
            cntDel = cntDel + 1
            If judge <> "想定どおり" Then WriteLog "差分確認", Mid$(k, InStr(k, "|") + 1), "", objKind, 0, "削除されています：" & judge
            a = before(k)
        Else
            cntAll = cntAll + 1
            Select Case chg
                Case "": cntSame = cntSame + 1
                Case "追加": cntAdd = cntAdd + 1
                Case "変更": cntMod = cntMod + 1
            End Select
            a = after(k)
        End If
        ws.Range(ws.Cells(r, 2), ws.Cells(r, lastCol)).NumberFormat = "@"
        ws.Cells(r, 1).Value = r - DIFF_FIRST + 1
        ws.Cells(r, 2).Value = Mid$(k, InStr(k, "|") + 1)
        ws.Cells(r, 3).Value = IIf(chg = "", "変化なし", chg)
        For i = 0 To nf - 1
            ws.Cells(r, 4 + i).Value = a(i)
        Next
        ws.Cells(r, cF).Value = Trim$(fields)
        If chg = "変更" Then ws.Cells(r, cF + 1).Value = Describe(fieldNames, before(k))
        If Not IsEmpty(op) Then ws.Cells(r, cF + 2).Value = op(0)
        ws.Cells(r, cF + 3).Value = st
        ws.Cells(r, cF + 4).Value = judge
        ' 行の色付けと、変更された項目セルの強調
        PaintDiffRow ws, r, lastCol, chg, judge
        If chg = "変更" Then
            For i = 0 To nf - 1
                If b(i) <> a(i) Then MarkChangedCell ws.Cells(r, 4 + i)
            Next
        End If
        r = r + 1
NextKey:
    Next
    If r > DIFF_FIRST Then
        With ws.Range(ws.Cells(DIFF_FIRST, 1), ws.Cells(r - 1, lastCol))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    ws.Range("B3").Value = Now
    ws.Range("D3").Value = cntAll
    ws.Range("F3").Value = cntSame
    ws.Range("H3").Value = cntAdd
    ws.Range("D4").Value = cntMod
    ws.Range("F4").Value = cntDel
    ws.Range("H4").Value = cntChk
    WriteDiffSheet = changed
End Function

' 変更前の内容（項目名: 値 を改行区切り、空の項目は省略）
Private Function Describe(ByVal fieldNames As Variant, ByVal v As Variant) As String
    Dim i As Long, s As String
    For i = 0 To UBound(v)
        If CStr(v(i)) <> "" Then
            If s <> "" Then s = s & vbLf
            s = s & fieldNames(i) & ": " & v(i)
        End If
    Next
    Describe = s
End Function

'----------------------------------------------------------
' 共通処理
'----------------------------------------------------------
' GET（ページ送りしながら全件取得）。失敗時は Nothing
Private Function FetchAllItems(ByVal baseU As String, ByVal opLabel As String, ByVal logOk As Boolean) As Collection
    Dim url As String, resText As String, st As Long, res As Object, item As Object
    Dim offset As Long, cnt As Long, total As Long, items As New Collection
    offset = 0
    Do
        url = WithQuery(baseU, ContainerQuery() & "&limit=" & PAGE_LIMIT & "&offset=" & offset)
        st = HttpCall("GET", url, "", resText)
        If st <> 200 Then
            WriteLog opLabel, "", "GET", url, st, ErrSummary(resText)
            Application.ScreenUpdating = True
            MsgBox opLabel & "でSCMからの取得に失敗しました（HTTP " & st & "）。実行ログを確認してください。", vbExclamation
            Exit Function
        End If
        Set res = ParseJson(resText)
        cnt = 0
        For Each item In res("data")
            items.Add item
            cnt = cnt + 1
        Next
        offset = offset + PAGE_LIMIT
        If res.Exists("total") Then total = res("total") Else total = 0
        If cnt < PAGE_LIMIT Or offset >= total Then Exit Do
    Loop
    If logOk Then WriteLog opLabel, "", "GET", WithQuery(baseU, ContainerQuery()), 200, items.Count & "件取得"
    Set FetchAllItems = items
End Function

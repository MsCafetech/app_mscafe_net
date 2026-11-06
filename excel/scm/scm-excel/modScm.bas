Attribute VB_Name = "modScm"
'==========================================================
' modScm : SCM (Prisma Access Configuration API) AddressObject 連携
'   前提 : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'   メンバー選択画面（ダブルクリック）: clsPickerEvents をインポートし、frmMemberPick は空のユーザーフォームを挿入して frmMemberPick_コード.txt を貼り付け
'==========================================================
' ※ 共通処理（認証・HTTP・ログ・差分色付け等）は modScmCommon にあります。modScmCommon（JSONパーサを含む）も併せてインポートしてください。
Option Explicit

Private Const SH_CFG As String = "基本設定"
Private Const SH_EXIST As String = "AdrObj_既存一覧"
Private Const SH_INPUT As String = "AdrObj_登録変更削除"
Private Const SH_EXISTG As String = "AdrGroupObj_既存一覧"
Private Const SH_INPUTG As String = "AdrGroupObj_登録変更削除"
Private Const SH_JSON As String = "JSON_CURL"
Private Const SH_DIFF As String = "AdrObj_差分確認"
Private Const SH_DIFFG As String = "AdrGroupObj_差分確認"
Private Const DIFF_FIRST As Long = 8    ' 差分確認 データ開始行
Private Const SH_LOG As String = "実行ログ"

Private Const EX_FIRST As Long = 8      ' 既存一覧 データ開始行
Private Const IN_FIRST As Long = 8      ' 登録変更削除 データ開始行
Private Const IN_LAST As Long = 507     ' 登録変更削除 データ最終行
Private Const JS_FIRST As Long = 10     ' JSON_CURL データ開始行
Private Const LOG_FIRST As Long = 6     ' 実行ログ データ開始行
Private Const PAGE_LIMIT As Long = 200
Private Const EP_GROUP As String = "/sse/config/v1/address-groups"


' ---- メンバー選択画面（旧 modPicker）用 ----
Private Const PICK_KEY As String = "^+m"            ' Ctrl+Shift+M
Private Const CELL_MAX As Long = 32767               ' Excelのセル文字数上限
Private Const SH_OVERFLOW As String = "ピッカー_選択結果"

Private mEvents As clsPickerEvents
Private mNames As Object     ' 候補 name -> 種別ラベル（追加順を保持）
Private mExclude As String   ' 候補から外す name（自分自身など）

'----------------------------------------------------------
' ボタン配置（初回のみ実行）
'----------------------------------------------------------
Public Sub SetupButtons()
    AddBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    AddBtn SH_EXIST, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUT, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJson"
    AddBtn SH_INPUT, "I3:J4", "入力クリア", "ClearInput"
    AddBtn SH_JSON, "G4:H4", "API実行（全行）", "ExecuteAll"
    AddBtn SH_JSON, "G6:H6", "クリア", "ClearJson"
    AddBtn SH_LOG, "F3", "ログクリア", "ClearLog"
    RemoveBtn SH_JSON, "PushConfig"     ' Push機能は削除（PushはSCM画面から実施）。旧ボタンを消す
    RemoveBtn SH_LOG, "CheckJobs"      ' ジョブ状況確認も削除。旧ボタンを消す
    AddBtn SH_EXISTG, "D3:E4", "既存Group取得（API）", "FetchExistingGroup"
    AddBtn SH_INPUTG, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJsonGroup"
    AddBtn SH_INPUTG, "I3:J4", "入力クリア", "ClearInputGroup"
    AddBtn SH_DIFF, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_DIFFG, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    MsgBox "ボタンを配置しました。", vbInformation
End Sub

'----------------------------------------------------------
' ② 既存情報取得（GET → 表生成）
'----------------------------------------------------------
Public Sub FetchExisting()
    Dim ws As Worksheet, url As String, resText As String, st As Long
    Dim res As Object, item As Object, offset As Long, r As Long, total As Long, cnt As Long
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Set ws = Worksheets(SH_EXIST)
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 13)).Clear
    r = EX_FIRST
    offset = 0
    Do
        url = WithQuery(BaseUrl(), ContainerQuery() & "&limit=" & PAGE_LIMIT & "&offset=" & offset)
        st = HttpCall("GET", url, "", resText)
        If st <> 200 Then
            WriteLog "既存取得", "", "GET", url, st, ErrSummary(resText)
            Application.ScreenUpdating = True
            MsgBox "既存情報の取得に失敗しました（HTTP " & st & "）。", vbExclamation
            Exit Sub
        End If
        Set res = ParseJson(resText)
        cnt = 0
        For Each item In res("data")
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 10)).NumberFormat = "@"
            ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
            ws.Cells(r, 2).Value = Nz(item, "id")
            ws.Cells(r, 3).Value = Nz(item, "name")
            ws.Cells(r, 4).Value = Nz(item, "ip_netmask")
            ws.Cells(r, 5).Value = Nz(item, "ip_range")
            ws.Cells(r, 6).Value = Nz(item, "ip_wildcard")
            ws.Cells(r, 7).Value = Nz(item, "fqdn")
            ws.Cells(r, 8).Value = Nz(item, "description")
            ws.Cells(r, 9).Value = JoinTags(item)
            ws.Cells(r, 10).Value = ContainerOf(item)
            ws.Cells(r, 11).Formula = "=IF(C" & r & "="""","""",IF(D" & r & "<>"""",""ip_netmask"",IF(E" & r & "<>"""",""ip_range"",IF(F" & r & "<>"""",""ip_wildcard"",IF(G" & r & "<>"""",""fqdn"","""")))))"
            ws.Cells(r, 12).Formula = "=IF(C" & r & "="""","""",D" & r & "&E" & r & "&F" & r & "&G" & r & ")"
            ws.Cells(r, 13).Value = Nz(item, "type")
            r = r + 1
            cnt = cnt + 1
        Next
        offset = offset + PAGE_LIMIT
        If res.Exists("total") Then total = res("total") Else total = 0
        If cnt < PAGE_LIMIT Or offset >= total Then Exit Do
    Loop
    If r > EX_FIRST Then
        ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 13)).Borders.LineStyle = xlContinuous
    End If
    Range("EX_FETCHED_AT").Value = Now
    Application.ScreenUpdating = True
    WriteLog "既存取得", "", "GET", WithQuery(BaseUrl(), ContainerQuery()), 200, (r - EX_FIRST) & "件取得"
    MsgBox (r - EX_FIRST) & " 件取得しました。", vbInformation
End Sub

'----------------------------------------------------------
' ③⑦ JSON発行（登録変更削除 + 登録変更削除_Group → JSON_CURL）
'   どちらのボタンでも③と⑦の両方から④を作り直す。
'   実行順：Address登録・変更 → Group登録・変更 → Group削除 → Address削除
'   Groupのメンバーに入っているAddressを削除する場合は、先にGroupから外す変更を追加する。
'----------------------------------------------------------
Public Sub GenerateJson()
    GenerateAll
End Sub

Public Sub GenerateJsonGroup()
    GenerateAll
End Sub

Private Sub GenerateAll()
    Dim wsA As Worksheet, wsG As Worksheet, wsEG As Worksheet, wsOut As Worksheet
    Dim r As Long, o As Long, ng As Long, nm As String, op As String
    Dim delA As Object, grpIn As Object, autoG As Object, removed As String, rest As String
    Dim note As String, emptyG As String, k As Variant
    If Not CheckContainer() Then Exit Sub
    Application.Calculate
    Set wsA = Worksheets(SH_INPUT)
    Set wsG = Worksheets(SH_INPUTG)
    Set wsEG = Worksheets(SH_EXISTG)
    Set wsOut = Worksheets(SH_JSON)
    ng = Application.WorksheetFunction.CountIf(wsA.Range("P" & IN_FIRST & ":P" & IN_LAST), "NG") + _
         Application.WorksheetFunction.CountIf(wsG.Range("P" & IN_FIRST & ":P" & IN_LAST), "NG")
    If ng > 0 Then
        If MsgBox("③⑦に判定NGが " & ng & " 件あります。NG行を除外してJSONを発行しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    ' 削除予定のAddress、⑦の対象Group
    Set delA = CreateObject("Scripting.Dictionary")
    Set grpIn = CreateObject("Scripting.Dictionary")
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r) Then
            If CStr(wsA.Cells(r, 2).Value) = "削除" Then delA(CStr(wsA.Cells(r, 3).Value)) = True
        End If
        If IsTarget(wsG, r) Then grpIn(CStr(wsG.Cells(r, 3).Value)) = r
    Next

    ' 削除予定のAddressをメンバーに持つGroup（key = Group名, value = ⑥の行）
    Set autoG = CreateObject("Scripting.Dictionary")
    If delA.Count > 0 Then
        For r = EX_FIRST To LastRow(wsEG, 3)
            nm = CStr(wsEG.Cells(r, 3).Value)
            If nm <> "" And CStr(wsEG.Cells(r, 4).Value) <> "" Then
                rest = StripMembers(CStr(wsEG.Cells(r, 4).Value), delA, removed)
                If removed <> "" Then
                    If grpIn.Exists(nm) Then
                        op = CStr(wsG.Cells(grpIn(nm), 2).Value)
                        If op = "削除" Then
                            note = note & vbLf & "・" & nm & "：⑦でGroupごと削除（" & removed & "）"
                        End If
                    Else
                        autoG(nm) = r
                        note = note & vbLf & "・" & nm & "：メンバーから外す変更を自動追加（" & removed & "）"
                        If rest = "" Then emptyG = emptyG & vbLf & "・" & nm
                    End If
                End If
            End If
        Next
        ' ⑦で登録・変更するGroupのメンバーに削除予定のAddressがあれば外す（処理のバッティング回避）
        For Each k In grpIn.Keys
            r = grpIn(k)
            If CStr(wsG.Cells(r, 2).Value) <> "削除" And CStr(wsG.Cells(r, 4).Value) = "static" Then
                rest = StripMembers(CStr(wsG.Cells(r, 5).Value), delA, removed)
                If removed <> "" Then
                    note = note & vbLf & "・" & k & "：⑦の" & wsG.Cells(r, 2).Value & "からメンバーを外す（" & removed & "）"
                    If rest = "" Then emptyG = emptyG & vbLf & "・" & k
                End If
            End If
        Next
    End If
    If emptyG <> "" Then
        MsgBox "Addressを削除すると、次のGroupのメンバーが0件になります。" & vbLf & _
               "⑦でGroupを削除するか、残すメンバーを追加してから発行してください。" & emptyG, vbExclamation
        Exit Sub
    End If
    If note <> "" Then
        If MsgBox("③で削除するAddressが、次のGroupのメンバーに入っています。" & note & vbLf & vbLf & _
                  "Groupを先に変更（または削除）し、その後にAddressを削除する順で④に出力します。続けますか？", _
                  vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    ClearJsonRows wsOut
    o = JS_FIRST
    ' 1. Address 登録・変更
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r) And CStr(wsA.Cells(r, 2).Value) <> "削除" Then WriteJsonRowAddr wsA, r, wsOut, o
    Next
    ' 2. Group 登録・変更（削除予定のメンバーは外す）
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsG, r) And CStr(wsG.Cells(r, 2).Value) <> "削除" Then WriteJsonRowGroup wsG, r, wsOut, o, delA
    Next
    ' 2'. 削除予定のAddressをGroupから外す自動変更
    For Each k In autoG.Keys
        r = autoG(k)
        WriteJsonRow wsOut, o, "変更(Group)自動", CStr(k), "PUT", GroupUrl() & "/" & wsEG.Cells(r, 2).Value, _
                     BuildBodyGroupParts(CStr(k), "static", StripMembers(CStr(wsEG.Cells(r, 4).Value), delA, removed), _
                                         CStr(wsEG.Cells(r, 6).Value), CStr(wsEG.Cells(r, 7).Value))
    Next
    ' 3. Group 削除
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsG, r) And CStr(wsG.Cells(r, 2).Value) = "削除" Then WriteJsonRowGroup wsG, r, wsOut, o, delA
    Next
    ' 4. Address 削除
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r) And CStr(wsA.Cells(r, 2).Value) = "削除" Then WriteJsonRowAddr wsA, r, wsOut, o
    Next
    If o > JS_FIRST Then
        With wsOut.Range(wsOut.Cells(JS_FIRST, 1), wsOut.Cells(o - 1, 8))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    wsOut.Activate
    MsgBox "③⑦から " & (o - JS_FIRST) & " 件のJSONを④に生成しました（うちGroup自動変更 " & autoG.Count & " 件）。", vbInformation
End Sub

' JSON発行の対象行か（nameあり、判定がOKか警告）
Private Function IsTarget(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim judge As String
    judge = CStr(ws.Cells(r, 16).Value)
    IsTarget = (CStr(ws.Cells(r, 3).Value) <> "" And (judge = "OK" Or judge = "警告"))
End Function

Private Sub WriteJsonRowAddr(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long)
    Dim op As String
    op = CStr(wsIn.Cells(r, 2).Value)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op, CStr(wsIn.Cells(r, 3).Value), "POST", WithQuery(BaseUrl(), ContainerQuery()), BuildBody(wsIn, r)
        Case "変更"
            WriteJsonRow wsOut, o, op, CStr(wsIn.Cells(r, 3).Value), "PUT", BaseUrl() & "/" & wsIn.Cells(r, 9).Value, BuildBody(wsIn, r)
        Case "削除"
            WriteJsonRow wsOut, o, op, CStr(wsIn.Cells(r, 3).Value), "DELETE", BaseUrl() & "/" & wsIn.Cells(r, 9).Value, ""
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

' メンバー文字列（カンマ・改行区切り）から削除予定のものを除く。戻り値は残りのカンマ区切り、removed に外したもの
Private Function StripMembers(ByVal members As String, ByVal delA As Object, ByRef removed As String) As String
    Dim parts() As String, i As Long, m As String, rest As String
    removed = ""
    parts = Split(Replace(Replace(members, vbCr, ""), vbLf, ","), ",")
    For i = LBound(parts) To UBound(parts)
        m = Trim$(parts(i))
        If m <> "" Then
            If delA.Exists(m) Then
                If removed <> "" Then removed = removed & ","
                removed = removed & m
            Else
                If rest <> "" Then rest = rest & ","
                rest = rest & m
            End If
        End If
    Next
    StripMembers = rest
End Function

Private Function BuildBody(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim js As String, tags() As String, t As String, i As Long
    js = "{""name"":" & JsonStr(CStr(ws.Cells(r, 3).Value)) & _
         ",""" & ws.Cells(r, 4).Value & """:" & JsonStr(CStr(ws.Cells(r, 5).Value))
    If CStr(ws.Cells(r, 6).Value) <> "" Then js = js & ",""description"":" & JsonStr(CStr(ws.Cells(r, 6).Value))
    If CStr(ws.Cells(r, 7).Value) <> "" Then
        tags = Split(CStr(ws.Cells(r, 7).Value), ",")
        For i = LBound(tags) To UBound(tags)
            If Trim$(tags(i)) <> "" Then
                If t <> "" Then t = t & ","
                t = t & JsonStr(Trim$(tags(i)))
            End If
        Next
        js = js & ",""tag"":[" & t & "]"
    End If
    BuildBody = js & "}"
End Function

'----------------------------------------------------------
' ④ API実行（JSON_CURL の全行）
'----------------------------------------------------------
Public Sub ExecuteAll()
    Dim ws As Worksheet, r As Long, n As Long, st As Long, resText As String, ok As Long, ng As Long
    Set ws = Worksheets(SH_JSON)
    n = LastRow(ws, 4) - JS_FIRST + 1
    If n <= 0 Then
        MsgBox "実行対象がありません。③⑦でJSON発行してください。", vbExclamation
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
    ' 実行後のSCMの状態を取得し、②⑥既存一覧と比較して⑧差分確認に出力
    BuildDiffCore
    Worksheets(SH_DIFF).Activate
    MsgBox "成功 " & ok & " 件／失敗 " & ng & " 件" & vbLf & _
           "AdrObj_差分確認 と AdrGroupObj_差分確認 に、現在の設定一覧と既存一覧（実行前）からの変化を出力しました。内容を確認してから、SCMの画面でPushしてください。", vbInformation
End Sub

'----------------------------------------------------------
' クリア系
'----------------------------------------------------------
Public Sub ClearInput()
    If MsgBox("入力（B～G列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUT).Range("B" & IN_FIRST & ":G" & IN_LAST).ClearContents
End Sub

Public Sub ClearJson()
    ClearJsonRows Worksheets(SH_JSON)
End Sub

Private Sub ClearJsonRows(ByVal ws As Worksheet)
    ws.Range(ws.Cells(JS_FIRST, 1), ws.Cells(ws.Rows.Count, 8)).Clear
End Sub

'==========================================================
' AddressGroup 用（⑥既存一覧_Group / ⑦登録変更削除_Group）
'==========================================================

' ⑥ 既存Group取得
Public Sub FetchExistingGroup()
    Dim ws As Worksheet, url As String, resText As String, st As Long
    Dim res As Object, item As Object, offset As Long, r As Long, total As Long, cnt As Long
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Set ws = Worksheets(SH_EXISTG)
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 11)).Clear
    r = EX_FIRST
    offset = 0
    Do
        url = WithQuery(GroupUrl(), ContainerQuery() & "&limit=" & PAGE_LIMIT & "&offset=" & offset)
        st = HttpCall("GET", url, "", resText)
        If st <> 200 Then
            WriteLog "既存Group取得", "", "GET", url, st, ErrSummary(resText)
            Application.ScreenUpdating = True
            MsgBox "既存Groupの取得に失敗しました（HTTP " & st & "）。", vbExclamation
            Exit Sub
        End If
        Set res = ParseJson(resText)
        cnt = 0
        For Each item In res("data")
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).NumberFormat = "@"
            ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
            ws.Cells(r, 2).Value = Nz(item, "id")
            ws.Cells(r, 3).Value = Nz(item, "name")
            ws.Cells(r, 4).Value = JoinStatic(item)
            ws.Cells(r, 5).Value = DynamicFilter(item)
            ws.Cells(r, 6).Value = Nz(item, "description")
            ws.Cells(r, 7).Value = JoinTags(item)
            ws.Cells(r, 8).Value = ContainerOf(item)
            ws.Cells(r, 9).Formula = "=IF(C" & r & "="""","""",IF(D" & r & "<>"""",""static"",IF(E" & r & "<>"""",""dynamic"","""")))"
            ws.Cells(r, 10).Formula = "=IF(C" & r & "="""","""",D" & r & "&E" & r & ")"
            ws.Cells(r, 11).Value = Nz(item, "type")
            r = r + 1
            cnt = cnt + 1
        Next
        offset = offset + PAGE_LIMIT
        If res.Exists("total") Then total = res("total") Else total = 0
        If cnt < PAGE_LIMIT Or offset >= total Then Exit Do
    Loop
    If r > EX_FIRST Then
        ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 11)).Borders.LineStyle = xlContinuous
    End If
    Range("EXG_FETCHED_AT").Value = Now
    Application.ScreenUpdating = True
    WriteLog "既存Group取得", "", "GET", WithQuery(GroupUrl(), ContainerQuery()), 200, (r - EX_FIRST) & "件取得"
    MsgBox (r - EX_FIRST) & " 件取得しました。", vbInformation
End Sub

Private Sub WriteJsonRowGroup(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long, ByVal delA As Object)
    Dim op As String, nm As String, v As String, removed As String
    op = CStr(wsIn.Cells(r, 2).Value)
    nm = CStr(wsIn.Cells(r, 3).Value)
    v = CStr(wsIn.Cells(r, 5).Value)
    If CStr(wsIn.Cells(r, 4).Value) = "static" Then v = StripMembers(v, delA, removed)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "POST", WithQuery(GroupUrl(), ContainerQuery()), _
                         BuildBodyGroupParts(nm, CStr(wsIn.Cells(r, 4).Value), v, CStr(wsIn.Cells(r, 6).Value), CStr(wsIn.Cells(r, 7).Value))
        Case "変更"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "PUT", GroupUrl() & "/" & wsIn.Cells(r, 9).Value, _
                         BuildBodyGroupParts(nm, CStr(wsIn.Cells(r, 4).Value), v, CStr(wsIn.Cells(r, 6).Value), CStr(wsIn.Cells(r, 7).Value))
        Case "削除"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "DELETE", GroupUrl() & "/" & wsIn.Cells(r, 9).Value, ""
    End Select
End Sub

Private Function BuildBodyGroupParts(ByVal nm As String, ByVal typ As String, ByVal v As String, _
                                     ByVal desc As String, ByVal tag As String) As String
    Dim js As String, parts() As String, t As String, i As Long
    js = "{""name"":" & JsonStr(nm)
    If typ = "static" Then
        ' カンマ区切り・セル内改行のどちらにも対応
        parts = Split(Replace(Replace(v, vbCr, ""), vbLf, ","), ",")
        For i = LBound(parts) To UBound(parts)
            If Trim$(parts(i)) <> "" Then
                If t <> "" Then t = t & ","
                t = t & JsonStr(Trim$(parts(i)))
            End If
        Next
        js = js & ",""static"":[" & t & "]"
    Else
        js = js & ",""dynamic"":{""filter"":" & JsonStr(v) & "}"
    End If
    If desc <> "" Then js = js & ",""description"":" & JsonStr(desc)
    t = ""
    If tag <> "" Then
        parts = Split(tag, ",")
        For i = LBound(parts) To UBound(parts)
            If Trim$(parts(i)) <> "" Then
                If t <> "" Then t = t & ","
                t = t & JsonStr(Trim$(parts(i)))
            End If
        Next
        js = js & ",""tag"":[" & t & "]"
    End If
    BuildBodyGroupParts = js & "}"
End Function

Public Sub ClearInputGroup()
    If MsgBox("入力（B～G列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUTG).Range("B" & IN_FIRST & ":G" & IN_LAST).ClearContents
End Sub

Private Function GroupUrl() As String
    Dim b As String
    b = Cfg("CFG_BASE_URL")
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    GroupUrl = b & EP_GROUP
End Function

Private Function JoinStatic(ByVal item As Object) As String
    Dim m As Variant, r As String
    If Not item.Exists("static") Then Exit Function
    For Each m In item("static")
        If r <> "" Then r = r & ","
        r = r & m
    Next
    JoinStatic = r
End Function

Private Function DynamicFilter(ByVal item As Object) As String
    If Not item.Exists("dynamic") Then Exit Function
    If Not IsObject(item("dynamic")) Then Exit Function
    DynamicFilter = Nz(item("dynamic"), "filter")
End Function

'==========================================================
' ⑧ 差分確認（②⑥既存一覧＝実行前 と SCMの最新状態＝実行後 を比較）
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
    Dim before As Object, after As Object, ops As Object
    Dim ws As Worksheet, r As Long, n As Long, opName As String
    Set before = CreateObject("Scripting.Dictionary")
    Set after = CreateObject("Scripting.Dictionary")
    Set ops = CreateObject("Scripting.Dictionary")
    Application.Calculate
    ' 実行前：②⑥既存一覧シート
    SheetToDict Worksheets(SH_EXIST), "Address", 3, 11, 12, 8, 9, before
    SheetToDict Worksheets(SH_EXISTG), "Group", 3, 9, 10, 6, 7, before
    ' 実行後：SCMから取得（シートは書き換えない）
    If Not ApiToDict(BaseUrl(), "Address", after) Then Exit Function
    If Not ApiToDict(GroupUrl(), "Group", after) Then Exit Function
    ' ④の操作
    Set ws = Worksheets(SH_JSON)
    For r = JS_FIRST To LastRow(ws, 4)
        opName = CStr(ws.Cells(r, 2).Value)
        If opName <> "" Then
            If InStr(opName, "(Group)") > 0 Then
                ops("Group|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
            Else
                ops("Address|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
            End If
        End If
    Next
    Application.ScreenUpdating = False
    n = WriteDiffSheet(Worksheets(SH_DIFF), "Address", before, after, ops)
    n = n + WriteDiffSheet(Worksheets(SH_DIFFG), "Group", before, after, ops)
    Application.ScreenUpdating = True
    WriteLog "差分確認", "", "GET", BaseUrl() & " / " & GroupUrl(), 200, n & "件の差分"
    BuildDiffCore = True
End Function

' 1種類分（Address / Group）：SCMの現在の全件を一覧表示し、既存一覧との変化を付ける。戻り値は変化のあった件数
Private Function WriteDiffSheet(ByVal ws As Worksheet, ByVal objKind As String, ByVal before As Object, _
                                ByVal after As Object, ByVal ops As Object) As Long
    Dim allKeys As Object, k As Variant, r As Long, b As Variant, a As Variant, op As Variant
    Dim kind As String, fields As String, expected As String, judge As String, st As String, changed As Long
    Dim cntAll As Long, cntSame As Long, cntAdd As Long, cntMod As Long, cntDel As Long, cntChk As Long
    Set allKeys = CreateObject("Scripting.Dictionary")
    AddKeys allKeys, after, objKind      ' 現在の一覧（SCMの順）
    AddKeys allKeys, before, objKind     ' 削除されたもの
    AddKeys allKeys, ops, objKind        ' ④にあるが両方に無いもの
    ' 行の色はVBAでセルに直接付ける（条件付き書式は優先されて上書きされるため、ここで削除する）
    With ws.Range(ws.Cells(DIFF_FIRST, 1), ws.Cells(ws.Rows.Count, 12))
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
        kind = "": fields = ""
        If before.Exists(k) And after.Exists(k) Then
            b = before(k): a = after(k)
            If b(0) <> a(0) Then fields = fields & "種別 "
            If b(1) <> a(1) Then fields = fields & "値 "
            If b(2) <> a(2) Then fields = fields & "description "
            If b(3) <> a(3) Then fields = fields & "tag "
            If fields <> "" Then kind = "変更"
        ElseIf after.Exists(k) Then
            kind = "追加"
        ElseIf before.Exists(k) Then
            kind = "削除"
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
            If kind = "" Then judge = "" Else judge = "想定外（④に無い変更）"
        ElseIf kind = expected Then
            judge = "想定どおり"
        ElseIf kind = "" Then
            If Left$(st, 1) = "2" And expected = "変更" Then
                judge = "変化なし（既存と同じ内容）"
            Else
                judge = "未反映（HTTP " & st & "）"
            End If
        Else
            judge = "想定と異なる"
        End If
        If kind <> "" Then changed = changed + 1
        If judge <> "" And judge <> "想定どおり" And Left$(judge, 4) <> "変化なし" Then cntChk = cntChk + 1
        ' 削除されたものは一覧に出さず件数だけ数える（想定外の削除は実行ログに残す）
        If kind = "削除" Then
            cntDel = cntDel + 1
            If judge <> "想定どおり" Then WriteLog "差分確認", Mid$(k, InStr(k, "|") + 1), "", objKind, 0, "削除されています：" & judge
            GoTo NextKey
        End If
        cntAll = cntAll + 1
        Select Case kind
            Case "": cntSame = cntSame + 1
            Case "追加": cntAdd = cntAdd + 1
            Case "変更": cntMod = cntMod + 1
        End Select
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 11)).NumberFormat = "@"
        ws.Cells(r, 1).Value = r - DIFF_FIRST + 1
        ws.Cells(r, 2).Value = Mid$(k, InStr(k, "|") + 1)
        ws.Cells(r, 3).Value = IIf(kind = "", "変化なし", kind)
        If after.Exists(k) Then
            a = after(k)
            ws.Cells(r, 4).Value = a(0)
            ws.Cells(r, 5).Value = a(1)
            ws.Cells(r, 6).Value = a(2)
            ws.Cells(r, 7).Value = a(3)
        End If
        ws.Cells(r, 8).Value = Trim$(fields)
        If kind = "変更" Then ws.Cells(r, 9).Value = Describe(before(k))
        If Not IsEmpty(op) Then ws.Cells(r, 10).Value = op(0)
        ws.Cells(r, 11).Value = st
        ws.Cells(r, 12).Value = judge
        ' 行の色付けと、変更された項目セルの強調
        PaintDiffRow ws, r, 12, kind, judge
        If kind = "変更" Then
            If b(0) <> a(0) Then MarkChangedCell ws.Cells(r, 4)
            If b(1) <> a(1) Then MarkChangedCell ws.Cells(r, 5)
            If b(2) <> a(2) Then MarkChangedCell ws.Cells(r, 6)
            If b(3) <> a(3) Then MarkChangedCell ws.Cells(r, 7)
        End If
        r = r + 1
NextKey:
    Next
    If r > DIFF_FIRST Then
        With ws.Range(ws.Cells(DIFF_FIRST, 1), ws.Cells(r - 1, 12))
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

' 既存一覧シート → Dictionary（key = 種類|name, value = Array(種別, 値, description, tag)）
Private Sub SheetToDict(ByVal ws As Worksheet, ByVal objKind As String, ByVal cName As Long, ByVal cType As Long, _
                        ByVal cVal As Long, ByVal cDesc As Long, ByVal cTag As Long, ByVal d As Object)
    Dim r As Long
    For r = EX_FIRST To LastRow(ws, cName)
        If CStr(ws.Cells(r, cName).Value) <> "" Then
            d(objKind & "|" & ws.Cells(r, cName).Value) = Array(CStr(ws.Cells(r, cType).Value), CStr(ws.Cells(r, cVal).Value), _
                                                                CStr(ws.Cells(r, cDesc).Value), CStr(ws.Cells(r, cTag).Value))
        End If
    Next
End Sub

' API → Dictionary（ページ送りしながら全件取得）
Private Function ApiToDict(ByVal baseU As String, ByVal objKind As String, ByVal d As Object) As Boolean
    Dim url As String, resText As String, st As Long, res As Object, item As Object
    Dim offset As Long, cnt As Long, total As Long, typ As String, v As String
    offset = 0
    Do
        url = WithQuery(baseU, ContainerQuery() & "&limit=" & PAGE_LIMIT & "&offset=" & offset)
        st = HttpCall("GET", url, "", resText)
        If st <> 200 Then
            WriteLog "差分確認", "", "GET", url, st, ErrSummary(resText)
            MsgBox "差分確認のための取得に失敗しました（HTTP " & st & "）。", vbExclamation
            Exit Function
        End If
        Set res = ParseJson(resText)
        cnt = 0
        For Each item In res("data")
            If objKind = "Address" Then
                typ = "": v = ""
                If item.Exists("ip_netmask") Then typ = "ip_netmask": v = Nz(item, "ip_netmask")
                If item.Exists("ip_range") Then typ = "ip_range": v = Nz(item, "ip_range")
                If item.Exists("ip_wildcard") Then typ = "ip_wildcard": v = Nz(item, "ip_wildcard")
                If item.Exists("fqdn") Then typ = "fqdn": v = Nz(item, "fqdn")
            Else
                If item.Exists("static") Then typ = "static": v = JoinStatic(item) Else typ = "dynamic": v = DynamicFilter(item)
            End If
            d(objKind & "|" & Nz(item, "name")) = Array(typ, v, Nz(item, "description"), JoinTags(item))
            cnt = cnt + 1
        Next
        offset = offset + PAGE_LIMIT
        If res.Exists("total") Then total = res("total") Else total = 0
        If cnt < PAGE_LIMIT Or offset >= total Then Exit Do
    Loop
    ApiToDict = True
End Function

Private Function Describe(ByVal v As Variant) As String
    Describe = v(0) & ": " & v(1)
    If v(2) <> "" Then Describe = Describe & vbLf & "description: " & v(2)
    If v(3) <> "" Then Describe = Describe & vbLf & "tag: " & v(3)
End Function

Private Function JoinTags(ByVal item As Object) As String
    Dim t As Variant, r As String
    If Not item.Exists("tag") Then Exit Function
    For Each t In item("tag")
        If r <> "" Then r = r & ","
        r = r & t
    Next
    JoinTags = r
End Function

'==========================================================
' メンバー選択画面（旧 modPicker）
'   必要モジュール : modScm（このモジュール） / clsPickerEvents / frmMemberPick
'   ブックを開くと Auto_Open で自動有効化されます。
'   ダブルクリックが効かない時は PickerStart を実行（または Ctrl+Shift+M）
'   対象列は PickerRule に追加します（他ツールのブックでも同じモジュールを共用可）
'==========================================================
'----------------------------------------------------------
' 有効化
'----------------------------------------------------------
Public Sub Auto_Open()
    PickerStart True
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.OnKey PICK_KEY
End Sub

Public Sub PickerStart(Optional ByVal silent As Boolean = False)
    Set mEvents = New clsPickerEvents
    Application.OnKey PICK_KEY, "'" & ThisWorkbook.Name & "'!PickerFromSelection"
    If Not silent Then MsgBox "メンバー選択（ダブルクリック）を有効にしました。", vbInformation
End Sub

' Ctrl+Shift+M : 選択中のセルでピッカーを開く（ダブルクリックが止まっていても再有効化）
Public Sub PickerFromSelection()
    If mEvents Is Nothing Then Set mEvents = New clsPickerEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Not TryPick(ActiveCell) Then
        MsgBox "このセルはチェックリスト選択の対象外です。", vbInformation
    End If
End Sub

'----------------------------------------------------------
' 対象セルならピッカーを表示して True を返す
'----------------------------------------------------------
Public Function TryPick(ByVal Target As Range) As Boolean
    Dim ttl As String
    If Target.CountLarge <> 1 Then Exit Function
    If Not Target.Worksheet.Parent Is ThisWorkbook Then Exit Function
    Set mNames = CreateObject("Scripting.Dictionary")   ' 既定でバイナリ比較（大文字・小文字を区別）
    mExclude = ""
    If Not PickerRule(Target.Worksheet, Target, ttl) Then Exit Function
    TryPick = True
    ShowPicker Target, ttl
End Function

'----------------------------------------------------------
' 対象列の定義（ここに追加すると他の列・ツールにも広げられます）
'   AddSrc シート名, name列, 開始行, 種別ラベル [, 条件列, 条件値]
'   AddMark シート名, name列, 開始行, 条件列, 条件値, 付記
'----------------------------------------------------------
Private Function PickerRule(ByVal ws As Worksheet, ByVal c As Range, ByRef ttl As String) As Boolean
    Select Case ws.Name
        Case "AdrGroupObj_登録変更削除"
            If c.Column <> 5 Or c.Row < 8 Or c.Row > 507 Then Exit Function
            If LCase$(Trim$(SafeStr(ws.Cells(c.Row, 4).Value))) = "dynamic" Then Exit Function
            ' 候補 = ②既存一覧 + ③で登録予定（O列のメンバー確認と同じ範囲）
            mExclude = SafeStr(ws.Cells(c.Row, 3).Value)
            AddSrc "AdrObj_既存一覧", 3, 8, "アドレス"
            AddSrc "AdrObj_登録変更削除", 3, 8, "アドレス(③で登録)", 2, "登録"
            AddMark "AdrObj_登録変更削除", 3, 8, 2, "削除", " ※③で削除予定"
            ttl = "AddressGroup メンバー選択：" & mExclude
            PickerRule = True
    End Select
End Function

Private Sub AddSrc(ByVal sh As String, ByVal col As Long, ByVal firstRow As Long, ByVal kind As String, _
                   Optional ByVal condCol As Long = 0, Optional ByVal condVal As String = "")
    Dim ws As Worksheet, v As Variant, cv As Variant, i As Long, nm As String
    If Not ReadCols(sh, col, firstRow, condCol, ws, v, cv) Then Exit Sub
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" And nm <> mExclude And Not mNames.Exists(nm) Then
            If condCol = 0 Then
                mNames.Add nm, kind
            ElseIf SafeStr(cv(i, 1)) = condVal Then
                mNames.Add nm, kind
            End If
        End If
    Next
End Sub

Private Sub AddMark(ByVal sh As String, ByVal col As Long, ByVal firstRow As Long, _
                    ByVal condCol As Long, ByVal condVal As String, ByVal note As String)
    Dim ws As Worksheet, v As Variant, cv As Variant, i As Long, nm As String
    If Not ReadCols(sh, col, firstRow, condCol, ws, v, cv) Then Exit Sub
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" And SafeStr(cv(i, 1)) = condVal Then
            If mNames.Exists(nm) Then mNames(nm) = mNames(nm) & note
        End If
    Next
End Sub

' name列（と条件列）を2次元配列で一括読込（高速化）
Private Function ReadCols(ByVal sh As String, ByVal col As Long, ByVal firstRow As Long, ByVal condCol As Long, _
                          ByRef ws As Worksheet, ByRef v As Variant, ByRef cv As Variant) As Boolean
    Dim lastR As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sh)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastR < firstRow Then Exit Function
    v = ToArray(ws.Range(ws.Cells(firstRow, col), ws.Cells(lastR, col)).Value)
    If condCol > 0 Then cv = ToArray(ws.Range(ws.Cells(firstRow, condCol), ws.Cells(lastR, condCol)).Value)
    ReadCols = True
End Function

Private Function ToArray(ByVal v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    If IsArray(v) Then
        ToArray = v
    Else
        a(1, 1) = v
        ToArray = a
    End If
End Function

'----------------------------------------------------------
' フォーム表示と書き戻し
'----------------------------------------------------------
Private Sub ShowPicker(ByVal Target As Range, ByVal ttl As String)
    Dim f As frmMemberPick, cur As String, sep As String, res As String
    cur = SafeStr(Target.Value)
    If InStr(cur, vbLf) > 0 Then sep = vbLf Else sep = ","   ' 元がセル内改行なら改行で返す
    Set f = New frmMemberPick
    f.Setup ttl, mNames, SplitItems(cur), sep
    f.Show
    If f.Confirmed Then
        res = f.Result
        If Len(res) > CELL_MAX Then
            If MsgBox("選択結果が " & Format$(Len(res), "#,##0") & " 文字になり、Excelのセル上限（32,767文字）を超えます。" & vbLf & _
                      "セルには書き込まず、シート「" & SH_OVERFLOW & "」に1行1件で出力しますか？", _
                      vbExclamation + vbOKCancel) = vbOK Then
                WriteOverflow Target, f.ResultItems
            End If
        Else
            Target.Value = res
        End If
    End If
    Unload f
End Sub

Private Sub WriteOverflow(ByVal Target As Range, ByVal items As Collection)
    Dim ws As Worksheet, a() As Variant, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_OVERFLOW)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_OVERFLOW
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "対象セル"
    ws.Range("B1").Value = "'" & Target.Worksheet.Name & "'!" & Target.Address(False, False)
    ws.Range("A2").Value = "件数"
    ws.Range("B2").Value = items.Count
    ws.Range("A4").Value = "選択結果（1行1件）"
    ReDim a(1 To items.Count, 1 To 1)
    For i = 1 To items.Count
        a(i, 1) = items(i)
    Next
    ws.Range("A5").Resize(items.Count, 1).NumberFormat = "@"
    ws.Range("A5").Resize(items.Count, 1).Value = a
    ws.Columns(1).ColumnWidth = 40
    ws.Activate
    MsgBox items.Count & " 件をシート「" & SH_OVERFLOW & "」に出力しました（元のセルは変更していません）。", vbInformation
End Sub

' カンマ・セル内改行のどちらの区切りにも対応
Public Function SplitItems(ByVal s As String) As Collection
    Dim c As New Collection, parts() As String, i As Long, t As String
    parts = Split(Replace(Replace(s, vbCr, ""), vbLf, ","), ",")
    For i = LBound(parts) To UBound(parts)
        t = Trim$(parts(i))
        If t <> "" Then c.Add t
    Next
    Set SplitItems = c
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

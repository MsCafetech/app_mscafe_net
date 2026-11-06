Attribute VB_Name = "modEdl"
'==========================================================
' modEdl : SCM (Prisma Access Configuration API) External Dynamic List（EDL）連携
'   前提 : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'==========================================================
' ※ 共通処理（認証・HTTP・ログ・差分色付け等）は modScmCommon にあります。modScmCommon（JSONパーサを含む）も併せてインポートしてください。
Option Explicit

Private Const SH_CFG As String = "基本設定"
Private Const SH_EXIST As String = "EDL_既存一覧"
Private Const SH_INPUT As String = "EDL_登録変更削除"
Private Const SH_JSON As String = "JSON_CURL"
Private Const SH_DIFF As String = "EDL_差分確認"
Private Const SH_LOG As String = "実行ログ"

Private Const EX_FIRST As Long = 8      ' 既存一覧 データ開始行
Private Const EX_CERT As Long = 17      ' 既存一覧 証明書プロファイル候補の列(Q)
Private Const IN_FIRST As Long = 8      ' 登録変更削除 データ開始行
Private Const IN_LAST As Long = 507     ' 登録変更削除 データ最終行
Private Const IN_JUDGE As Long = 31     ' ③ 判定列(AE)
Private Const IN_ID As Long = 17        ' ③ 既存id列(Q)
Private Const JS_FIRST As Long = 10     ' JSON_CURL データ開始行
Private Const DIFF_FIRST As Long = 8    ' 差分確認 データ開始行
Private Const PAGE_LIMIT As Long = 200
Private Const EP_CERT As String = "/sse/config/v1/certificate-profiles"

' ③ EDL_登録変更削除 の列
'   B=操作 C=name D=type E=url(取得元) F=recurring(更新間隔) G=at(時) H=day_of_week I=day_of_month
'   J=description K=exception_list L=certificate_profile M=expand_domain N=username O=password Q=既存id
' ② EDL_既存一覧 の列
'   B=id C=name D=type E=url F=recurring G=at H=day_of_week I=day_of_month
'   J=description K=exception_list L=certificate_profile M=expand_domain N=username O=コンテナ
'   Q=証明書プロファイル候補（VBA） S=定義済みリスト候補（固定）

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
' ② 既存情報取得（EDL を GET → 表生成。証明書プロファイル名もプルダウン候補として取得）
'----------------------------------------------------------
Public Sub FetchExisting()
    Dim items As Collection, certs As Collection, skipped As Long, msg As String
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Prog "EDLを取得しています..."
    Set items = FetchAllItems(BaseUrl(), "既存取得", True, True)
    If items Is Nothing Then
        Application.StatusBar = False
        Exit Sub
    End If
    Set items = UserItems(items, skipped)
    ' 証明書プロファイルは候補用のため、取得できなくても続行する
    Prog "証明書プロファイルを取得しています..."
    Set certs = FetchAllItems(CertUrl(), "証明書プロファイル取得", True, False)
    Application.ScreenUpdating = False
    WriteExist items
    WriteCertList certs
    Application.ScreenUpdating = True
    Application.StatusBar = False
    msg = "EDL " & items.Count & " 件を取得しました。"
    If skipped > 0 Then msg = msg & vbLf & "（定義済みのEDL " & skipped & " 件は一覧に含めていません）"
    If certs Is Nothing Then
        msg = msg & vbLf & "証明書プロファイルは取得できませんでした（実行ログ参照）。certificate_profile は空欄（None）のまま使えます。"
    Else
        msg = msg & vbLf & "証明書プロファイル " & certs.Count & " 件をプルダウン候補にしました。"
    End If
    MsgBox msg, vbInformation
End Sub

Private Sub WriteExist(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long, v As Variant, i As Long
    Set ws = Worksheets(SH_EXIST)
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 15)).Clear
    r = EX_FIRST
    For Each item In items
        v = EdlValues(item)
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 15)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        For i = 0 To 10
            ws.Cells(r, 4 + i).Value = v(i)
        Next
        ws.Cells(r, 15).Value = ContainerOf(item)
        r = r + 1
    Next
    If r > EX_FIRST Then
        With ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 15))
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlTop
        End With
        ws.Range(ws.Cells(EX_FIRST, 11), ws.Cells(r - 1, 11)).WrapText = True
    End If
    Range("EX_FETCHED_AT").Value = Now
End Sub

' 証明書プロファイル名（certificate_profile のプルダウン候補）を Q 列へ
Private Sub WriteCertList(ByVal certs As Collection)
    Dim ws As Worksheet, item As Object, r As Long
    If certs Is Nothing Then Exit Sub
    Set ws = Worksheets(SH_EXIST)
    ws.Range(ws.Cells(EX_FIRST, EX_CERT), ws.Cells(ws.Rows.Count, EX_CERT)).Clear
    r = EX_FIRST
    For Each item In certs
        If Nz(item, "name") <> "" Then
            ws.Cells(r, EX_CERT).NumberFormat = "@"
            ws.Cells(r, EX_CERT).Value = Nz(item, "name")
            ws.Cells(r, EX_CERT).Borders.LineStyle = xlContinuous
            r = r + 1
        End If
    Next
End Sub

' 定義済み（id なし／folder・snippet が predefined）のEDLを除く。skipped に除いた件数
Private Function UserItems(ByVal items As Collection, ByRef skipped As Long) As Collection
    Dim res As New Collection, item As Object
    skipped = 0
    For Each item In items
        If Nz(item, "id") = "" Or LCase$(Nz(item, "snippet")) = "predefined" Or LCase$(Nz(item, "folder")) = "predefined" Then
            skipped = skipped + 1
        Else
            res.Add item
        End If
    Next
    Set UserItems = res
End Function

' 1件分の type を Array(type, url, recurring, at, day_of_week, day_of_month, description,
'   exception_list, certificate_profile, expand_domain, username) に
'   certificate_profile の "None" は空欄、expand_domain は true のときだけ "true"
Private Function EdlValues(ByVal item As Object) As Variant
    Dim v(10) As String, ty As Object, d As Object, rc As Object, sb As Object, au As Object, k As Variant, rk As Variant
    If Not item.Exists("type") Then GoTo Done
    If Not IsObject(item("type")) Then GoTo Done
    Set ty = item("type")
    For Each k In ty.Keys
        If IsObject(ty(k)) Then
            Set d = ty(k)
            v(0) = k
            v(1) = Nz(d, "url")
            If d.Exists("recurring") Then
                If IsObject(d("recurring")) Then
                    Set rc = d("recurring")
                    For Each rk In rc.Keys
                        v(2) = rk
                        If IsObject(rc(rk)) Then
                            Set sb = rc(rk)
                            v(3) = Nz(sb, "at")
                            v(4) = Nz(sb, "day_of_week")
                            v(5) = Nz(sb, "day_of_month")
                        End If
                        Exit For
                    Next
                End If
            End If
            v(6) = Nz(d, "description")
            v(7) = JoinArr(d, "exception_list")
            v(8) = Nz(d, "certificate_profile")
            If v(8) = "None" Then v(8) = ""
            If d.Exists("expand_domain") Then
                If Not IsObject(d("expand_domain")) Then
                    If LCase$(CStr(d("expand_domain"))) = "true" Then v(9) = "true"
                End If
            End If
            If d.Exists("auth") Then
                If IsObject(d("auth")) Then
                    Set au = d("auth")
                    v(10) = Nz(au, "username")
                End If
            End If
            Exit For
        End If
    Next
Done:
    EdlValues = v
End Function

'----------------------------------------------------------
' ③ JSON発行（EDL_登録変更削除 → JSON_CURL）
'   実行順：登録・変更（行順）→ 削除（行順）
'----------------------------------------------------------
Public Sub GenerateJson()
    Dim wsIn As Worksheet, wsOut As Worksheet, r As Long, o As Long, ng As Long, bad As String
    If Not CheckContainer() Then Exit Sub
    Application.Calculate
    Set wsIn = Worksheets(SH_INPUT)
    Set wsOut = Worksheets(SH_JSON)
    ng = Application.WorksheetFunction.CountIf(wsIn.Range(wsIn.Cells(IN_FIRST, IN_JUDGE), wsIn.Cells(IN_LAST, IN_JUDGE)), "NG")
    If ng > 0 Then
        If MsgBox("③に判定NGが " & ng & " 件あります。NG行を除外してJSONを発行しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    ' exception_list の各要素（文字数・重複）は関数で判定しきれないためここで確認する
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsIn, r) And CStr(wsIn.Cells(r, 2).Value) <> "削除" Then bad = bad & CheckExceptions(wsIn, r)
    Next
    If bad <> "" Then
        MsgBox "exception_list に次の問題があります。修正してから発行してください。" & bad, vbExclamation
        Exit Sub
    End If

    ClearJsonRows wsOut
    o = JS_FIRST
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsIn, r) And CStr(wsIn.Cells(r, 2).Value) <> "削除" Then WriteJsonRowEdl wsIn, r, wsOut, o
    Next
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsIn, r) And CStr(wsIn.Cells(r, 2).Value) = "削除" Then WriteJsonRowEdl wsIn, r, wsOut, o
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

' exception_list の要素の文字数（imsi=34 / imei=32 / その他=255）と重複を確認。問題があれば説明文を返す
Private Function CheckExceptions(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim parts As Variant, i As Long, mx As Long, seen As Object, s As String, p As String
    parts = SplitMembers(CStr(ws.Cells(r, 11).Value))
    Select Case CStr(ws.Cells(r, 4).Value)
        Case "imsi": mx = 34
        Case "imei": mx = 32
        Case Else: mx = 255
    End Select
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(parts)
        p = CStr(parts(i))
        If Len(p) > mx Then s = s & vbLf & "・" & r & "行目 " & ws.Cells(r, 3).Value & "：" & mx & "文字超 " & Left$(p, 40) & "..."
        If seen.Exists(p) Then
            s = s & vbLf & "・" & r & "行目 " & ws.Cells(r, 3).Value & "：重複 " & p
        Else
            seen(p) = True
        End If
    Next
    CheckExceptions = s
End Function

Private Sub WriteJsonRowEdl(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long)
    Dim op As String, nm As String
    op = CStr(wsIn.Cells(r, 2).Value)
    nm = CStr(wsIn.Cells(r, 3).Value)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op, nm, "POST", WithQuery(BaseUrl(), ContainerQuery()), BuildBodyEdl(wsIn, r)
        Case "変更"
            WriteJsonRow wsOut, o, op, nm, "PUT", BaseUrl() & "/" & wsIn.Cells(r, IN_ID).Value, BuildBodyEdl(wsIn, r)
        Case "削除"
            WriteJsonRow wsOut, o, op, nm, "DELETE", BaseUrl() & "/" & wsIn.Cells(r, IN_ID).Value, ""
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

' EDL のボディ
'   {"name":..,"type":{"<type>":{"url":..,"recurring":{"daily":{"at":"03"}},"description":..,
'    "exception_list":[..],"certificate_profile":..,"expand_domain":true,"auth":{"username":..,"password":..}}}}
'   predefined_ip / predefined_url は url・description・exception_list のみ
Private Function BuildBodyEdl(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim ty As String, rec As String, js As String, rj As String, isPre As Boolean
    ty = CStr(ws.Cells(r, 4).Value)
    isPre = (Left$(ty, 11) = "predefined_")
    js = """url"":" & JsonStr(CStr(ws.Cells(r, 5).Value))
    If Not isPre Then
        rec = CStr(ws.Cells(r, 6).Value)
        Select Case rec
            Case "daily"
                rj = "{""at"":" & JsonStr(CStr(ws.Cells(r, 7).Value)) & "}"
            Case "weekly"
                rj = "{""day_of_week"":" & JsonStr(CStr(ws.Cells(r, 8).Value)) & ",""at"":" & JsonStr(CStr(ws.Cells(r, 7).Value)) & "}"
            Case "monthly"
                rj = "{""day_of_month"":" & CLng(ws.Cells(r, 9).Value) & ",""at"":" & JsonStr(CStr(ws.Cells(r, 7).Value)) & "}"
            Case Else       ' five_minute / hourly
                rj = "{}"
        End Select
        js = js & ",""recurring"":{""" & rec & """:" & rj & "}"
    End If
    If CStr(ws.Cells(r, 10).Value) <> "" Then js = js & ",""description"":" & JsonStr(CStr(ws.Cells(r, 10).Value))
    js = js & ListJson("exception_list", CStr(ws.Cells(r, 11).Value))
    If Not isPre Then
        If CStr(ws.Cells(r, 12).Value) <> "" Then js = js & ",""certificate_profile"":" & JsonStr(CStr(ws.Cells(r, 12).Value))
        If ty = "domain" And CStr(ws.Cells(r, 13).Value) = "true" Then js = js & ",""expand_domain"":true"
        If CStr(ws.Cells(r, 14).Value) <> "" Then
            js = js & ",""auth"":{""username"":" & JsonStr(CStr(ws.Cells(r, 14).Value)) & _
                      ",""password"":" & JsonStr(CStr(ws.Cells(r, 15).Value)) & "}"
        End If
    End If
    BuildBodyEdl = "{""name"":" & JsonStr(CStr(ws.Cells(r, 3).Value)) & ",""type"":{""" & ty & """:{" & js & "}}}"
End Function

' カンマ区切り／セル内改行の一覧 → ,"key":[..]（空なら空文字）
Private Function ListJson(ByVal key As String, ByVal s As String) As String
    Dim parts As Variant, t As String, i As Long
    parts = SplitMembers(s)
    For i = 0 To UBound(parts)
        If t <> "" Then t = t & ","
        t = t & JsonStr(CStr(parts(i)))
    Next
    If t <> "" Then ListJson = ",""" & key & """:[" & t & "]"
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
        Prog "API実行中 " & (r - JS_FIRST + 1) & " / " & n
        st = HttpCall(CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), CStr(ws.Cells(r, 6).Value), resText)
        ws.Cells(r, 8).Value = st
        If st >= 200 And st < 300 Then ok = ok + 1 Else ng = ng + 1
        WriteLog CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), st, ErrSummary(resText)
    Next
    ' 実行後のSCMの状態を取得し、②既存一覧と比較して⑥差分確認に出力
    Prog "差分確認を作成しています..."
    BuildDiffCore
    Application.StatusBar = False
    Worksheets(SH_DIFF).Activate
    MsgBox "成功 " & ok & " 件／失敗 " & ng & " 件" & vbLf & _
           "EDL_差分確認 に、現在の設定一覧と既存一覧（実行前）からの変化を出力しました。内容を確認してから、SCMの画面でPushしてください。", vbInformation
End Sub

'----------------------------------------------------------
' クリア系
'----------------------------------------------------------
Public Sub ClearInput()
    If MsgBox("入力（B～O列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUT).Range("B" & IN_FIRST & ":O" & IN_LAST).ClearContents
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
    Dim ws As Worksheet, r As Long, n As Long, opName As String, v As Variant, skipped As Long, i As Long
    Dim raw(10) As String
    Set before = CreateObject("Scripting.Dictionary")
    Set after = CreateObject("Scripting.Dictionary")
    Set ops = CreateObject("Scripting.Dictionary")
    Application.Calculate
    ' 実行前：②既存一覧シート
    Set ws = Worksheets(SH_EXIST)
    For r = EX_FIRST To LastRow(ws, 3)
        If CStr(ws.Cells(r, 3).Value) <> "" Then
            For i = 0 To 10
                raw(i) = CStr(ws.Cells(r, 4 + i).Value)
            Next
            before("EDL|" & ws.Cells(r, 3).Value) = DiffValues(raw)
        End If
    Next
    ' 実行後：SCMから取得（シートは書き換えない）
    Set items = FetchAllItems(BaseUrl(), "差分確認", False, True)
    If items Is Nothing Then Exit Function
    For Each item In UserItems(items, skipped)
        v = EdlValues(item)
        after("EDL|" & Nz(item, "name")) = DiffValues(v)
    Next
    ' ④の操作
    Set ws = Worksheets(SH_JSON)
    For r = JS_FIRST To LastRow(ws, 4)
        opName = CStr(ws.Cells(r, 2).Value)
        If opName <> "" Then ops("EDL|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
    Next
    Application.ScreenUpdating = False
    n = WriteDiffSheet(Worksheets(SH_DIFF), "EDL", _
                       Array("type", "url", "更新間隔", "description", "exception_list", "certificate_profile", "expand_domain", "username"), _
                       before, after, ops)
    Application.ScreenUpdating = True
    WriteLog "差分確認", "", "GET", BaseUrl(), 200, n & "件の変化"
    BuildDiffCore = True
End Function

' EdlValues の11項目 → 差分確認の8項目（recurring・at・day_of_week・day_of_month は1つにまとめる）
Private Function DiffValues(ByVal v As Variant) As Variant
    DiffValues = Array(CStr(v(0)), CStr(v(1)), RecText(CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5))), _
                       CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)))
End Function

' 更新間隔の表示用（例 hourly / daily 03時 / weekly monday 03時 / monthly 1日 03時）
Private Function RecText(ByVal rec As String, ByVal at As String, ByVal dow As String, ByVal dom As String) As String
    Dim s As String
    s = rec
    If dow <> "" Then s = s & " " & dow
    If dom <> "" Then s = s & " " & dom & "日"
    If at <> "" Then s = s & " " & at & "時"
    RecText = s
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
                judge = "変化なし（既存と同じ内容。passwordのみの変更は表示されません）"
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
' GET（ページ送りしながら全件取得）。失敗時は Nothing（showErr=False ならメッセージを出さずログのみ）
Private Function FetchAllItems(ByVal baseU As String, ByVal opLabel As String, ByVal logOk As Boolean, ByVal showErr As Boolean) As Collection
    Dim url As String, resText As String, st As Long, res As Object, item As Object
    Dim offset As Long, cnt As Long, total As Long, items As New Collection
    offset = 0
    Do
        url = WithQuery(baseU, ContainerQuery() & "&limit=" & PAGE_LIMIT & "&offset=" & offset)
        st = HttpCall("GET", url, "", resText)
        If st <> 200 Then
            WriteLog opLabel, "", "GET", url, st, ErrSummary(resText)
            If showErr Then
                Application.ScreenUpdating = True
                MsgBox opLabel & "でSCMからの取得に失敗しました（HTTP " & st & "）。実行ログを確認してください。", vbExclamation
            End If
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

Private Function CertUrl() As String
    Dim b As String
    b = Cfg("CFG_BASE_URL")
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    CertUrl = b & EP_CERT
End Function

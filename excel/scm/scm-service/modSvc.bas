Attribute VB_Name = "modSvc"
'==========================================================
' modSvc : SCM (Prisma Access Configuration API) ServiceObject / ServiceGroup 連携
'   前提 : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'   メンバー選択画面（ダブルクリック）: clsPickerEvents をインポートし、frmMemberPick は空のユーザーフォームを挿入して frmMemberPick_コード.txt を貼り付け
'     対象 : SvcGroupObj_登録変更削除 のメンバー（D列）
'==========================================================
' ※ 共通処理（認証・HTTP・ログ・差分色付け等）は modScmCommon にあります。modScmCommon（JSONパーサを含む）も併せてインポートしてください。
Option Explicit

Private Const SH_CFG As String = "基本設定"
Private Const SH_EXIST As String = "SvcObj_既存一覧"
Private Const SH_INPUT As String = "SvcObj_登録変更削除"
Private Const SH_EXISTG As String = "SvcGroupObj_既存一覧"
Private Const SH_INPUTG As String = "SvcGroupObj_登録変更削除"
Private Const SH_JSON As String = "JSON_CURL"
Private Const SH_DIFF As String = "SvcObj_差分確認"
Private Const SH_DIFFG As String = "SvcGroupObj_差分確認"
Private Const SH_LOG As String = "実行ログ"

Private Const EX_FIRST As Long = 8      ' 既存一覧 データ開始行
Private Const IN_FIRST As Long = 8      ' 登録変更削除 データ開始行
Private Const IN_LAST As Long = 507     ' 登録変更削除 データ最終行
Private Const IN_JUDGE As Long = 24     ' ③ 判定列(X)
Private Const ING_JUDGE As Long = 12    ' ⑦ 判定列(L)
Private Const JS_FIRST As Long = 10     ' JSON_CURL データ開始行
Private Const DIFF_FIRST As Long = 8    ' 差分確認 データ開始行
Private Const LOG_FIRST As Long = 6     ' 実行ログ データ開始行
Private Const PAGE_LIMIT As Long = 200
Private Const EP_GROUP As String = "/sse/config/v1/service-groups"

' ---- メンバー選択画面（ダブルクリック）用 ----
Private Const PICK_KEY As String = "^+m"            ' Ctrl+Shift+M
Private Const CELL_MAX As Long = 32767               ' Excelのセル文字数上限
Private Const SH_OVERFLOW As String = "ピッカー_選択結果"

Private mEvents As clsPickerEvents
Private mNames As Object     ' 候補 name -> 種別ラベル（追加順を保持）
Private mExclude As String   ' 候補から外す name（自分自身など）

' ③ SvcObj_登録変更削除 の列
'   B=操作 C=name D=protocol E=port F=source_port G=description H=tag
'   I=timeout J=halfclose_timeout K=timewait_timeout M=既存id
' ⑦ SvcGroupObj_登録変更削除 の列
'   B=操作 C=name D=メンバー E=tag G=既存id
' ② SvcObj_既存一覧 の列
'   B=id C=name D=protocol E=port F=source_port G=description H=tag
'   I=timeout J=halfclose_timeout K=timewait_timeout L=コンテナ M=定義(関数)
' ⑥ SvcGroupObj_既存一覧 の列
'   B=id C=name D=メンバー E=tag F=コンテナ

'----------------------------------------------------------
' ボタン配置（初回のみ実行）
'----------------------------------------------------------
Public Sub SetupButtons()
    AddBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    AddBtn SH_EXIST, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUT, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJson"
    AddBtn SH_INPUT, "I3:J4", "入力クリア", "ClearInput"
    AddBtn SH_EXISTG, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUTG, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJsonGroup"
    AddBtn SH_INPUTG, "I3:J4", "入力クリア", "ClearInputGroup"
    AddBtn SH_JSON, "G4:H4", "API実行（全行）", "ExecuteAll"
    AddBtn SH_JSON, "G6:H6", "クリア", "ClearJson"
    RemoveBtn SH_JSON, "PushConfig"     ' Push機能は削除（PushはSCM画面から実施）。旧ボタンを消す
    AddBtn SH_DIFF, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_DIFFG, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_LOG, "F3", "ログクリア", "ClearLog"
    RemoveBtn SH_LOG, "CheckJobs"      ' ジョブ状況確認も削除。旧ボタンを消す
    MsgBox "ボタンを配置しました。", vbInformation
End Sub

'----------------------------------------------------------
' ②⑥ 既存情報取得（Service と ServiceGroup の両方を GET → 表生成）
'----------------------------------------------------------
Public Sub FetchExisting()
    Dim svcs As Collection, grps As Collection
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Set svcs = FetchAllItems(BaseUrl(), "既存取得", True)
    If svcs Is Nothing Then Exit Sub
    Set grps = FetchAllItems(GroupUrl(), "既存Group取得", True)
    If grps Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    WriteExistSvc svcs
    WriteExistGroup grps
    Application.ScreenUpdating = True
    MsgBox "Service " & svcs.Count & " 件、ServiceGroup " & grps.Count & " 件を取得しました。", vbInformation
End Sub

Private Sub WriteExistSvc(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long, v As Variant
    Set ws = Worksheets(SH_EXIST)
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 13)).Clear
    r = EX_FIRST
    For Each item In items
        v = SvcValues(item)
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 12)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        ws.Cells(r, 4).Value = v(0)
        ws.Cells(r, 5).Value = v(1)
        ws.Cells(r, 6).Value = v(2)
        ws.Cells(r, 7).Value = Nz(item, "description")
        ws.Cells(r, 8).Value = JoinArr(item, "tag")
        ws.Cells(r, 9).Value = v(3)
        ws.Cells(r, 10).Value = v(4)
        ws.Cells(r, 11).Value = v(5)
        ws.Cells(r, 12).Value = ContainerOf(item)
        ws.Cells(r, 13).Formula = "=IF(C" & r & "="""","""",D" & r & "&""/""&E" & r & "&IF(F" & r & "<>"""","" src:""&F" & r & ",""""))"
        r = r + 1
    Next
    If r > EX_FIRST Then
        ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 13)).Borders.LineStyle = xlContinuous
    End If
    Range("EX_FETCHED_AT").Value = Now
End Sub

Private Sub WriteExistGroup(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long
    Set ws = Worksheets(SH_EXISTG)
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 6)).Clear
    r = EX_FIRST
    For Each item In items
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        ws.Cells(r, 4).Value = JoinArr(item, "members")
        ws.Cells(r, 5).Value = JoinArr(item, "tag")
        ws.Cells(r, 6).Value = ContainerOf(item)
        r = r + 1
    Next
    If r > EX_FIRST Then
        ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 6)).Borders.LineStyle = xlContinuous
    End If
    Range("EXG_FETCHED_AT").Value = Now
End Sub

' 1件分の protocol を Array(protocol, port, source_port, timeout, halfclose_timeout, timewait_timeout) に
Private Function SvcValues(ByVal item As Object) As Variant
    Dim v(5) As String, pr As Object, d As Object, ov As Object, k As Variant
    If item.Exists("protocol") Then
        If IsObject(item("protocol")) Then
            Set pr = item("protocol")
            For Each k In Array("tcp", "udp")
                If pr.Exists(k) Then
                    If IsObject(pr(k)) Then
                        Set d = pr(k)
                        v(0) = k
                        v(1) = Nz(d, "port")
                        v(2) = Nz(d, "source_port")
                        If d.Exists("override") Then
                            If IsObject(d("override")) Then
                                Set ov = d("override")
                                v(3) = Nz(ov, "timeout")
                                v(4) = Nz(ov, "halfclose_timeout")
                                v(5) = Nz(ov, "timewait_timeout")
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If
    SvcValues = v
End Function

'----------------------------------------------------------
' ③⑦ JSON発行（SvcObj_登録変更削除 + SvcGroupObj_登録変更削除 → JSON_CURL）
'   どちらのボタンでも③と⑦の両方から④を作り直す。
'   実行順：Service登録・変更 → Group登録・変更 → Group自動変更 → Group削除 → Service削除
'   ・削除するService／GroupがほかのGroupのメンバーなら、先にGroupから外す変更を追加する。
'   ・Groupのメンバーに⑦で登録・変更するGroupがあれば、そのGroupを先に登録・変更する。
'   ・Group同士を削除するときは、メンバーとして含んでいる側（親）を先に削除する。
'----------------------------------------------------------
Public Sub GenerateJson()
    GenerateAll
End Sub

Public Sub GenerateJsonGroup()
    GenerateAll
End Sub

Private Sub GenerateAll()
    Dim wsA As Worksheet, wsG As Worksheet, wsEG As Worksheet, wsOut As Worksheet
    Dim r As Long, o As Long, ng As Long, nm As String
    Dim delN As Object, grpIn As Object, autoG As Object, k As Variant, rw As Variant
    Dim removed As String, rest As String, note As String, emptyG As String
    Dim addOrder As Collection, delOrder As Collection
    If Not CheckContainer() Then Exit Sub
    Application.Calculate
    Set wsA = Worksheets(SH_INPUT)
    Set wsG = Worksheets(SH_INPUTG)
    Set wsEG = Worksheets(SH_EXISTG)
    Set wsOut = Worksheets(SH_JSON)
    ng = Application.WorksheetFunction.CountIf(wsA.Range(wsA.Cells(IN_FIRST, IN_JUDGE), wsA.Cells(IN_LAST, IN_JUDGE)), "NG") + _
         Application.WorksheetFunction.CountIf(wsG.Range(wsG.Cells(IN_FIRST, ING_JUDGE), wsG.Cells(IN_LAST, ING_JUDGE)), "NG")
    If ng > 0 Then
        If MsgBox("③⑦に判定NGが " & ng & " 件あります。NG行を除外してJSONを発行しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    ' 削除予定（Service と Group。名前は共通）、⑦の対象Group
    Set delN = CreateObject("Scripting.Dictionary")
    Set grpIn = CreateObject("Scripting.Dictionary")
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) Then
            If CStr(wsA.Cells(r, 2).Value) = "削除" Then delN(CStr(wsA.Cells(r, 3).Value)) = "Service"
        End If
        If IsTarget(wsG, r, ING_JUDGE) Then
            nm = CStr(wsG.Cells(r, 3).Value)
            grpIn(nm) = r
            If CStr(wsG.Cells(r, 2).Value) = "削除" Then delN(nm) = "Group"
        End If
    Next

    ' 削除予定をメンバーに持つ既存Group（⑦に無いもの。key = Group名, value = ⑥の行）
    Set autoG = CreateObject("Scripting.Dictionary")
    If delN.Count > 0 Then
        For r = EX_FIRST To LastRow(wsEG, 3)
            nm = CStr(wsEG.Cells(r, 3).Value)
            If nm <> "" And Not grpIn.Exists(nm) Then
                rest = StripMembers(CStr(wsEG.Cells(r, 4).Value), delN, removed)
                If removed <> "" Then
                    autoG(nm) = r
                    note = note & vbLf & "・" & nm & "：メンバーから外す変更を自動追加（" & removed & "）"
                    If rest = "" Then emptyG = emptyG & vbLf & "・" & nm
                End If
            End If
        Next
        ' ⑦で登録・変更するGroupのメンバーに削除予定があれば外す（処理のバッティング回避）
        For Each k In grpIn.Keys
            r = grpIn(k)
            If CStr(wsG.Cells(r, 2).Value) <> "削除" Then
                rest = StripMembers(CStr(wsG.Cells(r, 4).Value), delN, removed)
                If removed <> "" Then
                    note = note & vbLf & "・" & k & "：⑦の" & wsG.Cells(r, 2).Value & "からメンバーを外す（" & removed & "）"
                    If rest = "" Then emptyG = emptyG & vbLf & "・" & k
                End If
            End If
        Next
    End If
    If emptyG <> "" Then
        MsgBox "削除するService／Groupを外すと、次のGroupのメンバーが0件になります（メンバーは1件以上必須）。" & vbLf & _
               "⑦でGroupを削除するか、残すメンバーを追加してから発行してください。" & emptyG, vbExclamation
        Exit Sub
    End If
    If note <> "" Then
        If MsgBox("削除するService／Groupが、次のGroupのメンバーに入っています。" & note & vbLf & vbLf & _
                  "Groupを先に変更（または削除）し、その後に削除する順で④に出力します。続けますか？", _
                  vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    Set addOrder = OrderAddGroups(wsG, grpIn)
    If addOrder Is Nothing Then Exit Sub
    Set delOrder = OrderDelGroups(wsG, wsEG, grpIn)

    ClearJsonRows wsOut
    o = JS_FIRST
    ' 1. Service 登録・変更
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) And CStr(wsA.Cells(r, 2).Value) <> "削除" Then WriteJsonRowSvc wsA, r, wsOut, o
    Next
    ' 2. Group 登録・変更（メンバーに入るGroupを先に。削除予定のメンバーは外す）
    For Each rw In addOrder
        WriteJsonRowGroup wsG, CLng(rw), wsOut, o, delN
    Next
    ' 3. 削除予定をGroupから外す自動変更
    For Each k In autoG.Keys
        r = autoG(k)
        WriteJsonRow wsOut, o, "変更(Group)自動", CStr(k), "PUT", GroupUrl() & "/" & wsEG.Cells(r, 2).Value, _
                     BuildBodyGroup(CStr(k), StripMembers(CStr(wsEG.Cells(r, 4).Value), delN, removed), CStr(wsEG.Cells(r, 5).Value))
    Next
    ' 4. Group 削除（親Group → 子Group）
    For Each rw In delOrder
        WriteJsonRowGroup wsG, CLng(rw), wsOut, o, delN
    Next
    ' 5. Service 削除
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) And CStr(wsA.Cells(r, 2).Value) = "削除" Then WriteJsonRowSvc wsA, r, wsOut, o
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
Private Function IsTarget(ByVal ws As Worksheet, ByVal r As Long, ByVal judgeCol As Long) As Boolean
    Dim judge As String
    judge = CStr(ws.Cells(r, judgeCol).Value)
    IsTarget = (CStr(ws.Cells(r, 3).Value) <> "" And (judge = "OK" Or judge = "警告"))
End Function

' ⑦の登録・変更の順番：メンバーに⑦で登録・変更する別のGroupを含むものは、そのGroupの後にする
'   循環している場合は Nothing を返す
Private Function OrderAddGroups(ByVal wsG As Worksheet, ByVal grpIn As Object) As Collection
    Dim pending As Object, res As New Collection, k As Variant, parts As Variant
    Dim i As Long, ready As Boolean, progress As Boolean, cyc As String
    Set pending = CreateObject("Scripting.Dictionary")
    For Each k In grpIn.Keys
        If CStr(wsG.Cells(grpIn(k), 2).Value) <> "削除" Then pending(k) = grpIn(k)
    Next
    Do While pending.Count > 0
        progress = False
        For Each k In pending.Keys
            ready = True
            parts = SplitMembers(CStr(wsG.Cells(pending(k), 4).Value))
            For i = 0 To UBound(parts)
                If CStr(parts(i)) <> CStr(k) And pending.Exists(CStr(parts(i))) Then ready = False
            Next
            If ready Then
                res.Add pending(k)
                pending.Remove k
                progress = True
            End If
        Next
        If Not progress Then
            For Each k In pending.Keys
                cyc = cyc & vbLf & "・" & k
            Next
            MsgBox "⑦のGroup同士がお互いをメンバーに含んでいる（循環している）ため、登録・変更の順番を決められません。" & _
                   "メンバーを見直してください。" & cyc, vbExclamation
            Exit Function
        End If
    Loop
    Set OrderAddGroups = res
End Function

' ⑦の削除の順番：削除する別のGroupのメンバーに入っているGroupは、そのGroup（親）の後にする
Private Function OrderDelGroups(ByVal wsG As Worksheet, ByVal wsEG As Worksheet, ByVal grpIn As Object) As Collection
    Dim pending As Object, memb As Object, res As New Collection, k As Variant, j As Variant
    Dim r As Long, ready As Boolean, progress As Boolean
    Set pending = CreateObject("Scripting.Dictionary")
    For Each k In grpIn.Keys
        If CStr(wsG.Cells(grpIn(k), 2).Value) = "削除" Then pending(k) = grpIn(k)
    Next
    Set memb = CreateObject("Scripting.Dictionary")
    For r = EX_FIRST To LastRow(wsEG, 3)
        If CStr(wsEG.Cells(r, 3).Value) <> "" Then memb(CStr(wsEG.Cells(r, 3).Value)) = CStr(wsEG.Cells(r, 4).Value)
    Next
    Do While pending.Count > 0
        progress = False
        For Each k In pending.Keys
            ready = True
            For Each j In pending.Keys
                If CStr(j) <> CStr(k) And memb.Exists(j) Then
                    If IsMember(memb(j), CStr(k)) Then ready = False
                End If
            Next
            If ready Then
                res.Add pending(k)
                pending.Remove k
                progress = True
            End If
        Next
        If Not progress Then
            ' 循環（通常は起きない）：残りは入力順
            For Each k In pending.Keys
                res.Add pending(k)
            Next
            Exit Do
        End If
    Loop
    Set OrderDelGroups = res
End Function

Private Sub WriteJsonRowSvc(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long)
    Dim op As String, nm As String
    op = CStr(wsIn.Cells(r, 2).Value)
    nm = CStr(wsIn.Cells(r, 3).Value)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op, nm, "POST", WithQuery(BaseUrl(), ContainerQuery()), BuildBodySvc(wsIn, r)
        Case "変更"
            WriteJsonRow wsOut, o, op, nm, "PUT", BaseUrl() & "/" & wsIn.Cells(r, 13).Value, BuildBodySvc(wsIn, r)
        Case "削除"
            WriteJsonRow wsOut, o, op, nm, "DELETE", BaseUrl() & "/" & wsIn.Cells(r, 13).Value, ""
    End Select
End Sub

Private Sub WriteJsonRowGroup(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long, ByVal delN As Object)
    Dim op As String, nm As String, members As String, removed As String
    op = CStr(wsIn.Cells(r, 2).Value)
    nm = CStr(wsIn.Cells(r, 3).Value)
    members = StripMembers(CStr(wsIn.Cells(r, 4).Value), delN, removed)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "POST", WithQuery(GroupUrl(), ContainerQuery()), _
                         BuildBodyGroup(nm, members, CStr(wsIn.Cells(r, 5).Value))
        Case "変更"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "PUT", GroupUrl() & "/" & wsIn.Cells(r, 7).Value, _
                         BuildBodyGroup(nm, members, CStr(wsIn.Cells(r, 5).Value))
        Case "削除"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "DELETE", GroupUrl() & "/" & wsIn.Cells(r, 7).Value, ""
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

' Service のボディ {"name":..,"protocol":{"tcp":{"port":..,"source_port":..,"override":{..}}},"description":..,"tag":[..]}
Private Function BuildBodySvc(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim js As String, ov As String, c As Long, keyNames As Variant
    js = """port"":" & JsonStr(CStr(ws.Cells(r, 5).Value))
    If CStr(ws.Cells(r, 6).Value) <> "" Then js = js & ",""source_port"":" & JsonStr(CStr(ws.Cells(r, 6).Value))
    keyNames = Array("timeout", "halfclose_timeout", "timewait_timeout")
    For c = 9 To 11
        If Trim$(CStr(ws.Cells(r, c).Value)) <> "" Then
            If ov <> "" Then ov = ov & ","
            ov = ov & """" & keyNames(c - 9) & """:" & CLng(ws.Cells(r, c).Value)
        End If
    Next
    If ov <> "" Then js = js & ",""override"":{" & ov & "}"
    js = "{""name"":" & JsonStr(CStr(ws.Cells(r, 3).Value)) & _
         ",""protocol"":{""" & CStr(ws.Cells(r, 4).Value) & """:{" & js & "}}"
    If CStr(ws.Cells(r, 7).Value) <> "" Then js = js & ",""description"":" & JsonStr(CStr(ws.Cells(r, 7).Value))
    BuildBodySvc = js & TagJson(CStr(ws.Cells(r, 8).Value)) & "}"
End Function

' ServiceGroup のボディ {"name":..,"members":[..],"tag":[..]}（descriptionはAPI仕様に無い）
Private Function BuildBodyGroup(ByVal nm As String, ByVal members As String, ByVal tag As String) As String
    Dim parts As Variant, t As String, i As Long
    parts = SplitMembers(members)
    For i = 0 To UBound(parts)
        If t <> "" Then t = t & ","
        t = t & JsonStr(CStr(parts(i)))
    Next
    BuildBodyGroup = "{""name"":" & JsonStr(nm) & ",""members"":[" & t & "]" & TagJson(tag) & "}"
End Function

' tag（カンマ区切り）→ ,"tag":[..]（空なら空文字）
Private Function TagJson(ByVal tag As String) As String
    Dim parts() As String, t As String, i As Long
    If Trim$(tag) = "" Then Exit Function
    parts = Split(tag, ",")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            If t <> "" Then t = t & ","
            t = t & JsonStr(Trim$(parts(i)))
        End If
    Next
    If t <> "" Then TagJson = ",""tag"":[" & t & "]"
End Function

' メンバー文字列から削除予定のものを除く。戻り値は残りのカンマ区切り、removed に外したもの
Private Function StripMembers(ByVal members As String, ByVal delN As Object, ByRef removed As String) As String
    Dim parts As Variant, i As Long, m As String, rest As String
    removed = ""
    parts = SplitMembers(members)
    For i = 0 To UBound(parts)
        m = CStr(parts(i))
        If delN.Exists(m) Then
            If removed <> "" Then removed = removed & ","
            removed = removed & m
        Else
            If rest <> "" Then rest = rest & ","
            rest = rest & m
        End If
    Next
    StripMembers = rest
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
    ' 実行後のSCMの状態を取得し、②⑥既存一覧と比較して⑧⑨差分確認に出力
    BuildDiffCore
    Worksheets(SH_DIFF).Activate
    MsgBox "成功 " & ok & " 件／失敗 " & ng & " 件" & vbLf & _
           "SvcObj_差分確認 と SvcGroupObj_差分確認 に、現在の設定一覧と既存一覧（実行前）からの変化を出力しました。内容を確認してから、SCMの画面でPushしてください。", vbInformation
End Sub

'----------------------------------------------------------
' クリア系
'----------------------------------------------------------
Public Sub ClearInput()
    If MsgBox("入力（B～K列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUT).Range("B" & IN_FIRST & ":K" & IN_LAST).ClearContents
End Sub

Public Sub ClearInputGroup()
    If MsgBox("入力（B～E列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUTG).Range("B" & IN_FIRST & ":E" & IN_LAST).ClearContents
End Sub

Public Sub ClearJson()
    ClearJsonRows Worksheets(SH_JSON)
End Sub

Private Sub ClearJsonRows(ByVal ws As Worksheet)
    ws.Range(ws.Cells(JS_FIRST, 1), ws.Cells(ws.Rows.Count, 8)).Clear
End Sub

'==========================================================
' ⑧⑨ 差分確認（②⑥既存一覧＝実行前 と SCMの最新状態＝実行後 を比較）
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
    Dim ws As Worksheet, r As Long, n As Long, opName As String, v As Variant
    Set before = CreateObject("Scripting.Dictionary")
    Set after = CreateObject("Scripting.Dictionary")
    Set ops = CreateObject("Scripting.Dictionary")
    Application.Calculate
    ' 実行前：②⑥既存一覧シート
    Set ws = Worksheets(SH_EXIST)
    For r = EX_FIRST To LastRow(ws, 3)
        If CStr(ws.Cells(r, 3).Value) <> "" Then
            before("Service|" & ws.Cells(r, 3).Value) = Array(CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), _
                CStr(ws.Cells(r, 6).Value), CStr(ws.Cells(r, 7).Value), CStr(ws.Cells(r, 8).Value), _
                OvrText(CStr(ws.Cells(r, 9).Value), CStr(ws.Cells(r, 10).Value), CStr(ws.Cells(r, 11).Value)))
        End If
    Next
    Set ws = Worksheets(SH_EXISTG)
    For r = EX_FIRST To LastRow(ws, 3)
        If CStr(ws.Cells(r, 3).Value) <> "" Then
            before("Group|" & ws.Cells(r, 3).Value) = Array(CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value))
        End If
    Next
    ' 実行後：SCMから取得（シートは書き換えない）
    Set items = FetchAllItems(BaseUrl(), "差分確認", False)
    If items Is Nothing Then Exit Function
    For Each item In items
        v = SvcValues(item)
        after("Service|" & Nz(item, "name")) = Array(v(0), v(1), v(2), Nz(item, "description"), JoinArr(item, "tag"), _
                                                     OvrText(v(3), v(4), v(5)))
    Next
    Set items = FetchAllItems(GroupUrl(), "差分確認", False)
    If items Is Nothing Then Exit Function
    For Each item In items
        after("Group|" & Nz(item, "name")) = Array(JoinArr(item, "members"), JoinArr(item, "tag"))
    Next
    ' ④の操作
    Set ws = Worksheets(SH_JSON)
    For r = JS_FIRST To LastRow(ws, 4)
        opName = CStr(ws.Cells(r, 2).Value)
        If opName <> "" Then
            If InStr(opName, "(Group)") > 0 Then
                ops("Group|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
            Else
                ops("Service|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
            End If
        End If
    Next
    Application.ScreenUpdating = False
    n = WriteDiffSheet(Worksheets(SH_DIFF), "Service", _
                       Array("protocol", "port", "source_port", "description", "tag", "timeout"), before, after, ops)
    n = n + WriteDiffSheet(Worksheets(SH_DIFFG), "Group", Array("メンバー", "tag"), before, after, ops)
    Application.ScreenUpdating = True
    WriteLog "差分確認", "", "GET", BaseUrl() & " / " & GroupUrl(), 200, n & "件の変化"
    BuildDiffCore = True
End Function

' 1種類分（Service / Group）：SCMの現在の全件と削除されたものを一覧表示し、既存一覧との変化を付ける。戻り値は変化のあった件数
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

' override の表示用（例 timeout=3600 halfclose=120 timewait=15）
Private Function OvrText(ByVal tmo As String, ByVal half As String, ByVal tw As String) As String
    Dim s As String
    If tmo <> "" Then s = "timeout=" & tmo
    If half <> "" Then s = Trim$(s & " halfclose=" & half)
    If tw <> "" Then s = Trim$(s & " timewait=" & tw)
    OvrText = s
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

Private Function GroupUrl() As String
    Dim b As String
    b = Cfg("CFG_BASE_URL")
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    GroupUrl = b & EP_GROUP
End Function

'==========================================================
' メンバー選択画面（旧 modPicker）
'   必要モジュール : modSvc（このモジュール） / clsPickerEvents / frmMemberPick
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
        Case SH_INPUTG
            If c.Column <> 4 Or c.Row < IN_FIRST Or c.Row > IN_LAST Then Exit Function
            ' 候補 = ②⑥既存一覧 + ③⑦で登録予定 + 定義済みサービス（K列のメンバー確認と同じ範囲）
            mExclude = SafeStr(ws.Cells(c.Row, 3).Value)
            AddSrc SH_EXIST, 3, EX_FIRST, "サービス"
            AddSrc SH_EXISTG, 3, EX_FIRST, "サービスグループ"
            AddSrc SH_INPUT, 3, IN_FIRST, "サービス(③で登録)", 2, "登録"
            AddSrc SH_INPUTG, 3, IN_FIRST, "サービスグループ(⑦で登録)", 2, "登録"
            AddFixed "service-http", "定義済み"
            AddFixed "service-https", "定義済み"
            AddMark SH_INPUT, 3, IN_FIRST, 2, "削除", " ※③で削除予定"
            AddMark SH_INPUTG, 3, IN_FIRST, 2, "削除", " ※⑦で削除予定"
            ttl = "ServiceGroup メンバー選択：" & mExclude
            PickerRule = True
    End Select
End Function

' 定義済み（シートに無い）候補を追加
Private Sub AddFixed(ByVal nm As String, ByVal kind As String)
    If nm <> mExclude And Not mNames.Exists(nm) Then mNames.Add nm, kind
End Sub

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

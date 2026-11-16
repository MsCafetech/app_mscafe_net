Attribute VB_Name = "modInFetch"
'==========================================================
' modInFetch : 統合入力シートの「既存情報をまとめて取得（API）」・接続情報の消去
'   ①基本設定のトークンで SCM（Prisma Access Configuration API）から GET だけを行い、
'   ②既存データ抽出・②-2／②-3 既存ポリシー・参照_ポリシー選択肢 を作り直す（登録・変更・削除は一切しません）
'   取得する folder = ①の「オブジェクト取得folder」＋「取得・登録の対象スコープ」で○のもの（同じ id は1回だけ）
'   前提 : modScmCommon（トークン・HTTP・JSON・実行ログ）、modScmInput、modInPolicy を同じブックに入れる
'==========================================================
Option Explicit

Private Const S2 As String = "②既存データ抽出"
Private Const R1 As String = "参照_HIP選択肢"
Private Const R2 As String = "参照_HIPベンダー一覧"
Private Const SH_CFG As String = "①基本設定"

Private Const EX_FIRST As Long = 8
Private Const EX_LAST As Long = 5000
Private Const EXP_LAST As Long = 2000
Private Const URLV_LAST As Long = 30007
Private Const VD_FIRST As Long = 6
Private Const CELL_MAX As Long = 32767

' ② ブロックの位置（modScmInput と同じ）
Private Const OFF_ADR As Long = 0
Private Const OFF_ADRG As Long = 14
Private Const OFF_SVC As Long = 26
Private Const OFF_SVCG As Long = 40
Private Const OFF_URL As Long = 47
Private Const OFF_URLV As Long = 53
Private Const OFF_HIP As Long = 57
Private Const OFF_HIPP As Long = 124
Private Const OFF_EDL As Long = 133
Private Const OFF_PG As Long = 153

Private Const F_FIRST As Long = 6       ' HIPオブジェクト：F列(6)～BM列(65) が description と条件の各項目

'==========================================================
' ボタン（modScmInput.SetupButtons から呼ぶ）
'==========================================================
Public Sub SetupFetchButtons()
    FBtn SH_CFG, "D30:E30", "既存情報をまとめて取得（API）", "FetchExistingAll"
    FBtn SH_CFG, "D32:E32", "接続情報を消去（お客様に渡す前に）", "ClearCredentials"
    FBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    FBtn S2, "L3:N3", "SCMから取得（API）", "FetchExistingAll"
    FBtn "実行ログ", "F3:G3", "ログクリア", "ClearLog"
End Sub

Private Sub FBtn(ByVal sh As String, ByVal addr As String, ByVal caption As String, ByVal macro As String)
    Dim ws As Worksheet, rng As Range, b As Object
    Set ws = Worksheets(sh)
    Set rng = ws.Range(addr)
    On Error Resume Next
    ws.Buttons("btn_" & macro & "_" & Replace(addr, ":", "")).Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    b.Name = "btn_" & macro & "_" & Replace(addr, ":", "")
    b.Caption = caption
    b.OnAction = macro
End Sub

'==========================================================
' 既存情報をまとめて取得（API・GET のみ）
'==========================================================
Public Sub FetchExistingAll()
    Dim msg As String, ng As Long, t0 As Single, fl As Collection, f As Variant, fs As String
    Set fl = InFolders()
    For Each f In fl
        fs = fs & IIf(fs = "", "", "、") & f
    Next
    If MsgBox("SCM から既存の設定を取得し、次のシートを作り直します（GET のみ。SCM の設定は変えません）。" & vbLf & vbLf & _
              "・②既存データ抽出（アドレス・グループ・サービス・URLカテゴリ・HIP・EDL・プロファイルグループ）" & vbLf & _
              "・②-2既存セキュリティポリシー／②-3既存復号ポリシー" & vbLf & _
              "・参照_ポリシー選択肢（タグ・リージョン・アプリ・証明書・復号プロファイル等）" & vbLf & vbLf & _
              "取得する folder：" & fs & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    t0 = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo EH
    ng = ng + FetchAdr(msg)
    ng = ng + FetchAdrGroup(msg)
    ng = ng + FetchSvc(msg)
    ng = ng + FetchSvcGroup(msg)
    ng = ng + FetchUrl(msg)
    ng = ng + FetchHip(msg)
    ng = ng + FetchHipProfile(msg)
    ng = ng + FetchEdl(msg)
    ng = ng + FetchProfGroups(msg)
    Worksheets(S2).Range("B3").Value = Now
    ng = ng + modInPolicy.FetchPolicies(msg)
    ng = ng + modInPolicy.FetchPolicyChoices(msg)
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Worksheets(S2).Activate
    msg = "SCM から取得しました（" & SecText(Timer - t0) & "）。" & msg
    If ng > 0 Then msg = msg & vbLf & vbLf & ng & " 種類は取得に失敗しました（実行ログ参照。そのブロックは前の内容のまま）。"
    msg = msg & vbLf & vbLf & "⑫⑬を既存ポリシーで作り直す場合は、⑫⑬の「既存ポリシーを展開」を押してください。" & vbLf & _
          "お客様にお渡しする前に、①の「接続情報を消去」を押してください。"
    MsgBox msg, IIf(ng > 0, vbExclamation, vbInformation)
    Exit Sub
EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "取得中にエラーが発生しました：" & Err.Description & vbLf & "（ここまでに取得できた分はシートに書かれています）", vbExclamation
End Sub

' 戻り値：失敗なら 1
Private Function FetchAdr(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long
    Set items = InGetFolders("/sse/config/v1/addresses", "既存取得（アドレス）")
    If items Is Nothing Then FetchAdr = 1: msg = msg & vbLf & "　アドレス：失敗": Exit Function
    a = NewRows(items.Count, 13)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        a(i, 2) = Nz(item, "id")
        a(i, 3) = Nz(item, "name")
        a(i, 4) = Nz(item, "ip_netmask")
        a(i, 5) = Nz(item, "ip_range")
        a(i, 6) = Nz(item, "ip_wildcard")
        a(i, 7) = Nz(item, "fqdn")
        a(i, 8) = Nz(item, "description")
        a(i, 9) = JoinArr(item, "tag")
        a(i, 10) = ContainerOf(item)
        a(i, 13) = Nz(item, "type")
    Next
    PutBlock OFF_ADR, "2-10,13", a, items.Count, EX_LAST
    msg = msg & vbLf & "　アドレス：" & items.Count & " 件"
End Function

Private Function FetchAdrGroup(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long
    Set items = InGetFolders("/sse/config/v1/address-groups", "既存取得（アドレスグループ）")
    If items Is Nothing Then FetchAdrGroup = 1: msg = msg & vbLf & "　アドレスグループ：失敗": Exit Function
    a = NewRows(items.Count, 11)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        a(i, 2) = Nz(item, "id")
        a(i, 3) = Nz(item, "name")
        a(i, 4) = JoinArr(item, "static")
        If item.Exists("dynamic") Then
            If IsObject(item("dynamic")) Then a(i, 5) = Nz(item("dynamic"), "filter")
        End If
        a(i, 6) = Nz(item, "description")
        a(i, 7) = JoinArr(item, "tag")
        a(i, 8) = ContainerOf(item)
        a(i, 11) = Nz(item, "type")
    Next
    PutBlock OFF_ADRG, "2-8,11", a, items.Count, EX_LAST
    msg = msg & vbLf & "　アドレスグループ：" & items.Count & " 件"
End Function

Private Function FetchSvc(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long, v As Variant
    Set items = InGetFolders("/sse/config/v1/services", "既存取得（サービス）")
    If items Is Nothing Then FetchSvc = 1: msg = msg & vbLf & "　サービス：失敗": Exit Function
    a = NewRows(items.Count, 12)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        v = SvcValues(item)
        a(i, 2) = Nz(item, "id")
        a(i, 3) = Nz(item, "name")
        a(i, 4) = v(0)
        a(i, 5) = v(1)
        a(i, 6) = v(2)
        a(i, 7) = Nz(item, "description")
        a(i, 8) = JoinArr(item, "tag")
        a(i, 9) = v(3)
        a(i, 10) = v(4)
        a(i, 11) = v(5)
        a(i, 12) = ContainerOf(item)
    Next
    PutBlock OFF_SVC, "2-12", a, items.Count, EX_LAST
    msg = msg & vbLf & "　サービス：" & items.Count & " 件"
End Function

' protocol を Array(protocol, port, source_port, timeout, halfclose_timeout, timewait_timeout) に
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

Private Function FetchSvcGroup(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long
    Set items = InGetFolders("/sse/config/v1/service-groups", "既存取得（サービスグループ）")
    If items Is Nothing Then FetchSvcGroup = 1: msg = msg & vbLf & "　サービスグループ：失敗": Exit Function
    a = NewRows(items.Count, 6)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        a(i, 2) = Nz(item, "id")
        a(i, 3) = Nz(item, "name")
        a(i, 4) = JoinArr(item, "members")
        a(i, 5) = JoinArr(item, "tag")
        a(i, 6) = ContainerOf(item)
    Next
    PutBlock OFF_SVCG, "2-6", a, items.Count, EX_LAST
    msg = msg & vbLf & "　サービスグループ：" & items.Count & " 件"
End Function

' カスタムURLカテゴリ：一覧（name・type・description）と URL（1行1件）
Private Function FetchUrl(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, u() As Variant, i As Long, n As Long, x As Variant
    Set items = InGetFolders("/sse/config/v1/url-categories", "既存取得（URLカテゴリ）")
    If items Is Nothing Then FetchUrl = 1: msg = msg & vbLf & "　カスタムURLカテゴリ：失敗": Exit Function
    a = NewRows(items.Count, 4)
    u = NewRows(URLV_LAST - EX_FIRST + 1, 3)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        a(i, 2) = Nz(item, "name")
        a(i, 3) = Nz(item, "type")
        a(i, 4) = Nz(item, "description")
        If item.Exists("list") Then
            If IsObject(item("list")) Then
                For Each x In item("list")
                    If Not IsObject(x) And n < UBound(u, 1) Then
                        n = n + 1
                        u(n, 2) = Nz(item, "name")
                        u(n, 3) = CStr(x)
                    End If
                Next
            End If
        End If
    Next
    PutBlock OFF_URL, "2-4", a, items.Count, EX_LAST
    PutBlock OFF_URLV, "2-3", u, n, URLV_LAST
    msg = msg & vbLf & "　カスタムURLカテゴリ：" & items.Count & " 件（URL " & n & " 件）"
End Function

' HIPオブジェクト（B=id C=name D=コンテナ E=mobile_device F～BM=項目 BN=JSON）。使われているベンダー・製品を参照_HIPベンダー一覧へ
Private Function FetchHip(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long, d As Object, keys As Variant, k As Long, js As String
    Set items = InGetFolders("/sse/config/v1/hip-objects", "既存取得（HIPオブジェクト）")
    If items Is Nothing Then FetchHip = 1: msg = msg & vbLf & "　HIPオブジェクト：失敗": Exit Function
    keys = HipKeys()
    a = NewRows(items.Count, 66)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        Set d = ObjToRow(item)
        a(i, 2) = Nz(item, "id")
        a(i, 3) = Nz(item, "name")
        a(i, 4) = ContainerOf(item)
        a(i, 5) = FlatText(Child(item, "mobile_device"))
        For k = 0 To UBound(keys)
            If d.Exists(keys(k)) Then a(i, F_FIRST + k) = d(keys(k))
        Next
        js = InJsonOf(item)
        If Len(js) > CELL_MAX Then
            WriteLog "既存取得（HIPオブジェクト）", Nz(item, "name"), "", "", 0, "JSONがセルの上限を超えたため切り詰めました"
            js = Left$(js, CELL_MAX)
        End If
        a(i, 66) = js
    Next
    PutBlock OFF_HIP, "2-66", a, items.Count, EX_LAST
    MergeVendors items
    msg = msg & vbLf & "　HIPオブジェクト：" & items.Count & " 件"
End Function

Private Function FetchHipProfile(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long
    Set items = InGetFolders("/sse/config/v1/hip-profiles", "既存取得（HIPプロファイル）")
    If items Is Nothing Then FetchHipProfile = 1: msg = msg & vbLf & "　HIPプロファイル：失敗": Exit Function
    a = NewRows(items.Count, 7)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        a(i, 2) = Nz(item, "id")
        a(i, 3) = Nz(item, "name")
        a(i, 4) = Nz(item, "description")
        a(i, 5) = Nz(item, "match")
        a(i, 6) = JoinColl(QuotedNames(Nz(item, "match")), ",")
        a(i, 7) = ContainerOf(item)
    Next
    PutBlock OFF_HIPP, "2-7", a, items.Count, EXP_LAST
    msg = msg & vbLf & "　HIPプロファイル：" & items.Count & " 件"
End Function

' EDL（定義済みは除く）と証明書プロファイル名（EDL の Q 列と 参照_HIP選択肢 D 列の候補）
Private Function FetchEdl(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long, v As Variant, k As Long, skipN As Long
    Dim users As New Collection, certs As Collection, c() As Variant, n As Long, wsR As Worksheet
    Set items = InGetFolders("/sse/config/v1/external-dynamic-lists", "既存取得（EDL）")
    If items Is Nothing Then
        FetchEdl = 1: msg = msg & vbLf & "　EDL：失敗"
    Else
        For Each item In items
            If Nz(item, "id") = "" Or LCase$(Nz(item, "snippet")) = "predefined" Or LCase$(Nz(item, "folder")) = "predefined" Then
                skipN = skipN + 1
            Else
                users.Add item
            End If
        Next
        a = NewRows(users.Count, 15)
        For Each item In users
            i = i + 1
            If i > UBound(a, 1) Then Exit For
            v = EdlValues(item)
            a(i, 2) = Nz(item, "id")
            a(i, 3) = Nz(item, "name")
            For k = 0 To 10
                a(i, 4 + k) = v(k)
            Next
            a(i, 15) = ContainerOf(item)
        Next
        PutBlock OFF_EDL, "2-15", a, users.Count, EX_LAST
        msg = msg & vbLf & "　EDL：" & users.Count & " 件" & IIf(skipN > 0, "（定義済み " & skipN & " 件は除く）", "")
    End If
    ' 証明書プロファイル（候補用。失敗しても数えない）
    Set certs = InGetFolders("/sse/config/v1/certificate-profiles", "既存取得（証明書プロファイル）")
    If Not certs Is Nothing Then
        c = NewRows(certs.Count, 17)
        For Each item In certs
            If Nz(item, "name") <> "" And n < UBound(c, 1) Then
                n = n + 1
                c(n, 17) = Nz(item, "name")
            End If
        Next
        PutBlock OFF_EDL, "17", c, n, EX_LAST
        Set wsR = Worksheets(R1)
        wsR.Range("D5:D1005").ClearContents
        For k = 1 To n
            If k > 1001 Then Exit For
            wsR.Cells(4 + k, 4).NumberFormat = "@"
            wsR.Cells(4 + k, 4).Value = c(k, 17)
        Next
        msg = msg & vbLf & "　証明書プロファイル：" & n & " 件"
    End If
End Function

' type を Array(type, url, recurring, at, day_of_week, day_of_month, description,
'   exception_list, certificate_profile, expand_domain, username) に（パスワードは読みません）
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

' セキュリティプロファイルグループ（B=name C=folder D=定義済み/カスタム E～K=各プロファイル L=id）→ セキュリティプロファイルも作り直す
Private Function FetchProfGroups(ByRef msg As String) As Long
    Dim items As Collection, item As Variant, a() As Variant, i As Long, k As Long, keys As Variant
    Set items = InGetFolders("/sse/config/v1/profile-groups", "既存取得（プロファイルグループ）")
    If items Is Nothing Then FetchProfGroups = 1: msg = msg & vbLf & "　プロファイルグループ：失敗": Exit Function
    keys = Array("virus_and_wildfire_analysis", "spyware", "vulnerability", "url_filtering", "file_blocking", "dns_security", "saas_security")
    a = NewRows(items.Count, 12)
    For Each item In items
        i = i + 1
        If i > UBound(a, 1) Then Exit For
        a(i, 2) = Nz(item, "name")
        a(i, 3) = Nz(item, "folder")
        If Nz(item, "id") = "" Or InStr(LCase$(Nz(item, "folder") & " " & Nz(item, "snippet")), "predefined") > 0 Then
            a(i, 4) = "定義済み"
        Else
            a(i, 4) = "カスタム"
        End If
        For k = 0 To 6
            a(i, 5 + k) = JoinLines(item, CStr(keys(k)))
        Next
        a(i, 12) = Nz(item, "id")
    Next
    PutBlock OFF_PG, "2-12", a, items.Count, EXP_LAST
    msg = msg & vbLf & "　プロファイルグループ：" & items.Count & " 件"
    modScmInput.ImportProfiles msg
End Function

'==========================================================
' 接続情報を消去（お客様にお渡しする前に）
'==========================================================
Public Sub ClearCredentials()
    Dim ans As VbMsgBoxResult, nm As Variant
    ans = MsgBox("client_secret とアクセストークンを消去します。" & vbLf & _
                 "client_id と TSG ID も消しますか？" & vbLf & vbLf & _
                 "はい＝すべて消す／いいえ＝client_secret とトークンだけ消す／キャンセル＝何もしない", vbYesNoCancel + vbQuestion)
    If ans = vbCancel Then Exit Sub
    For Each nm In Array("CFG_CLIENT_SECRET", "CFG_TOKEN", "CFG_TOKEN_AT", "CFG_TOKEN_TTL")
        ClearName CStr(nm)
    Next
    If ans = vbYes Then
        ClearName "CFG_CLIENT_ID"
        ClearName "CFG_TSG_ID"
    End If
    MsgBox "消去しました。ブックを保存してからお客様にお渡しください。", vbInformation
End Sub

Private Sub ClearName(ByVal nm As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.ClearContents
    On Error GoTo 0
End Sub

'==========================================================
' 共通（modInPolicy からも使う）
'==========================================================
Public Function InEnc(ByVal s As String) As String
    InEnc = Application.WorksheetFunction.EncodeURL(s)
End Function

Public Function InApiBase() As String
    Dim b As String
    b = Cfg("CFG_BASE_URL")
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    InApiBase = b
End Function

Private Function PageLimit() As Long
    PageLimit = Val(CfgOpt("CFG_PAGE_LIMIT"))
    If PageLimit <= 0 Then PageLimit = 200
End Function

' 取得対象スコープ（○のもの、上から順）: Collection of Array(表示名, folder)
Public Function InSelScopes() As Collection
    Dim c As New Collection, rng As Range, i As Long
    Set rng = ThisWorkbook.Names("CFG_SCOPES").RefersToRange
    For i = 1 To rng.Rows.Count
        If Trim$(SafeStr(rng.Cells(i, 2).Value)) = "○" And Trim$(SafeStr(rng.Cells(i, 3).Value)) <> "" Then
            c.Add Array(Trim$(SafeStr(rng.Cells(i, 1).Value)), Trim$(SafeStr(rng.Cells(i, 3).Value)))
        End If
    Next
    Set InSelScopes = c
End Function

' オブジェクトを取得する folder（オブジェクト取得folder → ○のスコープ。重複なし）
Public Function InFolders() As Collection
    Dim c As New Collection, d As Object, sc As Variant, base As String
    Set d = CreateObject("Scripting.Dictionary")
    base = CfgOpt("CFG_OBJ_FOLDER")
    If base = "" Then base = "Shared"
    d(base) = True: c.Add base
    For Each sc In InSelScopes()
        If Not d.Exists(CStr(sc(1))) Then d(CStr(sc(1))) = True: c.Add CStr(sc(1))
    Next
    Set InFolders = c
End Function

' GET（ページ送りで全件）。失敗時は Nothing（実行ログに記録）。期限前・401 ではトークンを取り直す
Public Function InGetAll(ByVal url0 As String, ByVal q As String, ByVal opLabel As String, _
                         ByRef st As Long, ByRef resMsg As String) As Collection
    Dim url As String, resText As String, res As Object, item As Variant
    Dim offset As Long, cnt As Long, total As Long, lim As Long, items As New Collection
    lim = PageLimit()
    Do
        url = WithQuery(url0, q & "&limit=" & lim & "&offset=" & offset)
        If Not EnsureToken() Then st = 401: resMsg = "トークンを取得できません": Exit Function
        st = HttpCall("GET", url, "", resText)
        If st = 401 Then
            If FetchToken() Then st = HttpCall("GET", url, "", resText)
        End If
        If st <> 200 Then
            resMsg = ErrSummary(resText)
            WriteLog opLabel, "", "GET", url, st, resMsg
            Exit Function
        End If
        Set res = ParseJson(resText)
        cnt = 0: total = 0
        If Not res Is Nothing Then
            If res.Exists("data") Then
                If IsObject(res("data")) Then
                    For Each item In res("data")
                        items.Add item
                        cnt = cnt + 1
                    Next
                End If
            End If
            If res.Exists("total") Then total = Val(SafeStr(res("total")))
        End If
        offset = offset + lim
        If cnt < lim Then Exit Do
        If total > 0 And offset >= total Then Exit Do
        Prog opLabel & " 取得中 " & items.Count & " / " & total & " 件"
    Loop
    Set InGetAll = items
End Function

' 複数 folder から GET して1つにまとめる（同じ id は1回だけ。folder が無い要素には取得した folder を入れる）
'   すべての folder で失敗したら Nothing
Public Function InGetFolders(ByVal ep As String, ByVal opLabel As String) As Collection
    Dim f As Variant, items As Collection, item As Variant, c As New Collection, seen As Object, k As String
    Dim okN As Long, st As Long, m As String, n As Long
    Set seen = CreateObject("Scripting.Dictionary")
    For Each f In InFolders()
        Prog opLabel & "：" & f & " を取得中…"
        Set items = InGetAll(InApiBase() & ep, "folder=" & InEnc(CStr(f)), opLabel, st, m)
        If Not items Is Nothing Then
            okN = okN + 1
            n = 0
            For Each item In items
                If IsObject(item) Then
                    If Nz(item, "folder") = "" And Nz(item, "snippet") = "" And Nz(item, "device") = "" Then item("folder") = CStr(f)
                    k = Nz(item, "id")
                    If k = "" Then k = Nz(item, "name") & "|" & Nz(item, "folder") & Nz(item, "snippet")
                    If Not seen.Exists(k) Then
                        seen(k) = True
                        c.Add item
                        n = n + 1
                    End If
                End If
            Next
            WriteLog opLabel, CStr(f), "GET", WithQuery(InApiBase() & ep, "folder=" & InEnc(CStr(f))), 200, items.Count & " 件（新しく加えたもの " & n & " 件）"
        End If
    Next
    If okN > 0 Then Set InGetFolders = c
End Function

Public Function InJsonOf(ByVal v As Variant) As String
    Dim k As Variant, s As String, x As Variant
    If IsObject(v) Then
        If v Is Nothing Then
            InJsonOf = "null"
        ElseIf TypeName(v) = "Dictionary" Then
            For Each k In v.Keys
                If s <> "" Then s = s & ","
                s = s & JsonStr(CStr(k)) & ":" & InJsonOf(v(k))
            Next
            InJsonOf = "{" & s & "}"
        Else
            For Each x In v
                If s <> "" Then s = s & ","
                s = s & InJsonOf(x)
            Next
            InJsonOf = "[" & s & "]"
        End If
    Else
        Select Case VarType(v)
            Case vbBoolean: InJsonOf = IIf(v, "true", "false")
            Case vbNull, vbEmpty: InJsonOf = "null"
            Case vbString: InJsonOf = JsonStr(v)
            Case Else: InJsonOf = Replace(CStr(v), ",", ".")
        End Select
    End If
End Function

'==========================================================
' ② への書き込み
'==========================================================
' 1 To n（最低1行）× 1 To ncols の配列
Private Function NewRows(ByVal n As Long, ByVal ncols As Long) As Variant()
    Dim a() As Variant
    If n < 1 Then n = 1
    If n > EX_LAST * 6 Then n = EX_LAST * 6
    ReDim a(1 To n, 1 To ncols)
    NewRows = a
End Function

' spec（例 "2-10,13"）の列だけを消して書く（関数の列はそのまま）。a の列番号 = ブロック内の列番号
Private Sub PutBlock(ByVal off As Long, ByVal spec As String, ByRef a() As Variant, ByVal n As Long, ByVal lastAllowed As Long)
    Dim ws As Worksheet, part As Variant, c1 As Long, c2 As Long, blk() As Variant, i As Long, j As Long
    Set ws = Worksheets(S2)
    If EX_FIRST + n - 1 > lastAllowed Then n = lastAllowed - EX_FIRST + 1
    If n > UBound(a, 1) Then n = UBound(a, 1)
    For Each part In Split(spec, ",")
        c1 = CLng(Split(part, "-")(0))
        If InStr(part, "-") > 0 Then c2 = CLng(Split(part, "-")(1)) Else c2 = c1
        ws.Range(ws.Cells(EX_FIRST, off + c1), ws.Cells(lastAllowed, off + c2)).ClearContents
        If n > 0 Then
            ReDim blk(1 To n, 1 To c2 - c1 + 1)
            For i = 1 To n
                For j = c1 To c2
                    blk(i, j - c1 + 1) = a(i, j)
                Next
            Next
            With ws.Range(ws.Cells(EX_FIRST, off + c1), ws.Cells(EX_FIRST + n - 1, off + c2))
                .NumberFormat = "@"
                .Value = blk
            End With
        End If
    Next
End Sub

Private Function JoinLines(ByVal item As Object, ByVal key As String) As String
    Dim t As Variant, r As String
    If Not item.Exists(key) Then Exit Function
    If Not IsObject(item(key)) Then JoinLines = SafeStr(item(key)): Exit Function
    For Each t In item(key)
        If Not IsObject(t) Then
            If r <> "" Then r = r & vbLf
            r = r & SafeStr(t)
        End If
    Next
    JoinLines = r
End Function

'==========================================================
' HIPオブジェクト JSON → 行（既存ツール SCM_HIP.xlsx と同じ表現）
'==========================================================
Private Function HipKeys() As Variant
    HipKeys = Array("description", _
        "hi_os_vendor", "hi_os_value", "hi_managed", "hi_domain_op", "hi_domain_val", "hi_hostname_op", "hi_hostname_val", _
        "hi_clientver_op", "hi_clientver_val", "hi_serial_op", "hi_serial_val", "hi_hostid_op", "hi_hostid_val", _
        "am_installed", "am_rtp", "am_virdef_op", "am_virdef_unit", "am_virdef_n", "am_scan_op", "am_scan_unit", "am_scan_n", _
        "am_pv_op", "am_pv_val", "am_vendor", "am_excl", _
        "db_installed", "db_time_op", "db_time_unit", "db_time_n", "db_vendor", "db_excl", _
        "de_installed", "de_locations", "de_vendor", "de_excl", _
        "fw_installed", "fw_enabled", "fw_vendor", "fw_excl", _
        "pm_installed", "pm_enabled", "pm_check", "pm_sev_op", "pm_sev_val", "pm_patches", "pm_vendor", "pm_excl", _
        "dlp_installed", "dlp_enabled", "dlp_vendor", "dlp_excl", _
        "ce_profile", "ce_attrs", _
        "nw_op", "nw_type", "nw_val", _
        "cc_process", "cc_registry", "cc_plist")
End Function

Private Function ObjToRow(ByVal item As Object) As Object
    Dim r As Object, c As Object, o As Object, x As Object, k As Variant, op As Variant, cat As Variant
    Dim p As String, s As String, t As String, e As Object, v As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("description") = Nz(item, "description")

    ' ホスト情報
    Set c = Child(Child(item, "host_info"), "criteria")
    If Not c Is Nothing Then
        Set x = Child(Child(c, "os"), "contains")
        If Not x Is Nothing Then
            For Each k In Array("Microsoft", "Apple", "Google", "Linux", "Other")
                If x.Exists(k) Then r("hi_os_vendor") = k: r("hi_os_value") = Nz(x, CStr(k))
            Next
        End If
        If c.Exists("managed") Then r("hi_managed") = YN(c("managed"))
        For Each k In Array("domain|hi_domain", "host_name|hi_hostname", "client_version|hi_clientver", _
                            "serial_number|hi_serial", "host_id|hi_hostid")
            Set x = Child(c, Split(k, "|")(0))
            If Not x Is Nothing Then
                p = Split(k, "|")(1)
                For Each op In Array("is", "is_not", "contains")
                    If x.Exists(op) Then r(p & "_op") = op: r(p & "_val") = Nz(x, CStr(op))
                Next
            End If
        Next
    End If

    ' ベンダー系カテゴリ
    For Each cat In Array("anti_malware|am", "disk_backup|db", "disk_encryption|de", "firewall|fw", _
                          "patch_management|pm", "data_loss_prevention|dlp")
        Set o = Child(item, Split(cat, "|")(0))
        p = Split(cat, "|")(1)
        If Not o Is Nothing Then
            Set c = Child(o, "criteria")
            If Not c Is Nothing Then
                If c.Exists("is_installed") Then r(p & "_installed") = YN(c("is_installed"))
                If c.Exists("is_enabled") Then r(p & "_enabled") = Nz(c, "is_enabled")
            End If
            If o.Exists("vendor") Then r(p & "_vendor") = VendorText(o("vendor"))
            If o.Exists("exclude_vendor") Then r(p & "_excl") = YN(o("exclude_vendor"))
            If Not c Is Nothing Then
                Select Case p
                    Case "am"
                        If c.Exists("real_time_protection") Then r("am_rtp") = Nz(c, "real_time_protection")
                        TimeToRow Child(c, "virdef_version"), r, "am_virdef", Array("days", "versions")
                        TimeToRow Child(c, "last_scan_time"), r, "am_scan", Array("days", "hours")
                        Set x = Child(c, "product_version")
                        If Not x Is Nothing Then
                            For Each op In Array("greater_equal", "greater_than", "is", "is_not", "less_equal", "less_than", "contains", "within", "not_within")
                                If x.Exists(op) Then
                                    r("am_pv_op") = op
                                    If op = "within" Or op = "not_within" Then
                                        r("am_pv_val") = Nz(Child(x, CStr(op)), "versions")
                                    Else
                                        r("am_pv_val") = Nz(x, CStr(op))
                                    End If
                                End If
                            Next
                        End If
                    Case "db"
                        TimeToRow Child(c, "last_backup_time"), r, "db_time", Array("days", "hours")
                    Case "de"
                        If c.Exists("encrypted_locations") Then
                            s = ""
                            For Each e In c("encrypted_locations")
                                Set x = Child(e, "encryption_state")
                                t = Nz(e, "name")
                                If Not x Is Nothing Then
                                    If x.Exists("is_not") Then
                                        t = "!" & t & " | " & Nz(x, "is_not")
                                    ElseIf x.Exists("is") Then
                                        t = t & " | " & Nz(x, "is")
                                    End If
                                End If
                                s = AddLine(s, t)
                            Next
                            r("de_locations") = s
                        End If
                    Case "pm"
                        Set x = Child(c, "missing_patches")
                        If Not x Is Nothing Then
                            r("pm_check") = Nz(x, "check")
                            Set v = Child(x, "severity")
                            If Not v Is Nothing Then
                                For Each op In Array("greater_equal", "greater_than", "is", "is_not", "less_equal", "less_than")
                                    If v.Exists(op) Then r("pm_sev_op") = op: r("pm_sev_val") = Nz(v, CStr(op))
                                Next
                            End If
                            If x.Exists("patches") Then r("pm_patches") = JoinColl(x("patches"), vbLf)
                        End If
                End Select
            End If
        End If
    Next

    ' 証明書
    Set c = Child(Child(item, "certificate"), "criteria")
    If Not c Is Nothing Then
        r("ce_profile") = Nz(c, "certificate_profile")
        If c.Exists("certificate_attributes") Then
            s = ""
            For Each e In c("certificate_attributes")
                t = Nz(e, "name")
                If Nz(e, "value") <> "" Then t = t & " | " & Nz(e, "value")
                s = AddLine(s, t)
            Next
            r("ce_attrs") = s
        End If
    End If

    ' ネットワーク
    Set x = Child(Child(Child(item, "network_info"), "criteria"), "network")
    If Not x Is Nothing Then
        For Each op In Array("is", "is_not")
            Set o = Child(x, CStr(op))
            If Not o Is Nothing Then
                r("nw_op") = op
                For Each k In Array("wifi", "mobile", "ethernet", "unknown")
                    If o.Exists(k) Then
                        r("nw_type") = k
                        Set v = Child(o, CStr(k))
                        If Not v Is Nothing Then r("nw_val") = Nz(v, "ssid") & Nz(v, "carrier")
                    End If
                Next
            End If
        Next
    End If

    ' カスタムチェック
    Set c = Child(Child(item, "custom_checks"), "criteria")
    If Not c Is Nothing Then
        If c.Exists("process_list") Then
            s = ""
            For Each e In c("process_list")
                t = Nz(e, "name")
                If e.Exists("running") Then
                    If Not CBool(e("running")) Then t = "!" & t
                End If
                s = AddLine(s, t)
            Next
            r("cc_process") = s
        End If
        If c.Exists("registry_key") Then
            s = ""
            For Each e In c("registry_key")
                If CollCount(e, "registry_value") = 0 And Nz(e, "default_value_data") = "" Then
                    s = AddLine(s, NegMark(e) & Nz(e, "name"))
                End If
                If Nz(e, "default_value_data") <> "" Then s = AddLine(s, Nz(e, "name") & " | (既定) | " & Nz(e, "default_value_data"))
                If CollCount(e, "registry_value") > 0 Then
                    For Each v In e("registry_value")
                        t = Nz(e, "name") & " | " & NegMark(v) & Nz(v, "name")
                        If Nz(v, "value_data") <> "" Then t = t & " | " & Nz(v, "value_data")
                        s = AddLine(s, t)
                    Next
                End If
            Next
            r("cc_registry") = s
        End If
        If c.Exists("plist") Then
            s = ""
            For Each e In c("plist")
                If CollCount(e, "key") = 0 Then
                    s = AddLine(s, NegMark(e) & Nz(e, "name"))
                Else
                    For Each v In e("key")
                        t = Nz(e, "name") & " | " & NegMark(v) & Nz(v, "name")
                        If Nz(v, "value") <> "" Then t = t & " | " & Nz(v, "value")
                        s = AddLine(s, t)
                    Next
                End If
            Next
            r("cc_plist") = s
        End If
    End If
    Set ObjToRow = r
End Function

Private Sub TimeToRow(ByVal o As Object, ByVal r As Object, ByVal k As String, ByVal units As Variant)
    Dim op As Variant, u As Variant, x As Object
    If o Is Nothing Then Exit Sub
    For Each op In Array("not_available", "within", "not_within")
        If o.Exists(op) Then
            r(k & "_op") = op
            Set x = Child(o, CStr(op))
            If Not x Is Nothing Then
                For Each u In units
                    If x.Exists(u) Then r(k & "_unit") = u: r(k & "_n") = Nz(x, CStr(u))
                Next
            End If
        End If
    Next
End Sub

Private Function VendorText(ByVal vs As Variant) As String
    Dim v As Object, s As String, t As String
    If Not IsObject(vs) Then Exit Function
    For Each v In vs
        t = Nz(v, "name")
        If CollCount(v, "product") > 0 Then t = t & " | " & JoinColl(v("product"), " ; ")
        s = AddLine(s, t)
    Next
    VendorText = s
End Function

' 既存HIPオブジェクトで使われているベンダー・製品を 参照_HIPベンダー一覧 に追加（同じものがあれば追加しない）
Private Sub MergeVendors(ByVal items As Collection)
    Dim ws As Worksheet, have As Object, item As Variant, o As Object, vd As Variant, cat As Variant
    Dim r As Long, lbl As String, pr As Variant, n As Long
    Set ws = Worksheets(R2)
    Set have = CreateObject("Scripting.Dictionary")
    For r = VD_FIRST To LastRow(ws, 3)
        have(SafeStr(ws.Cells(r, 2).Value) & "|" & SafeStr(ws.Cells(r, 3).Value) & "|" & SafeStr(ws.Cells(r, 4).Value)) = True
    Next
    r = LastRow(ws, 3) + 1
    If r < VD_FIRST Then r = VD_FIRST
    For Each item In items
        For Each cat In Array("anti_malware|am", "disk_backup|db", "disk_encryption|de", "firewall|fw", _
                              "patch_management|pm", "data_loss_prevention|dlp")
            Set o = Child(item, Split(cat, "|")(0))
            lbl = CatLabel(Split(cat, "|")(1))
            If Not o Is Nothing Then
                If CollCount(o, "vendor") > 0 Then
                    For Each vd In o("vendor")
                        If CollCount(vd, "product") = 0 Then
                            AddVendorRow ws, have, r, lbl, Nz(vd, "name"), "", n
                        Else
                            For Each pr In vd("product")
                                AddVendorRow ws, have, r, lbl, Nz(vd, "name"), CStr(pr), n
                            Next
                        End If
                    Next
                End If
            End If
        Next
    Next
    If n > 0 Then WriteLog "既存取得（HIPオブジェクト）", "", "", R2, 200, "既存のHIPオブジェクトからベンダー・製品を " & n & " 件追加"
End Sub

Private Sub AddVendorRow(ByVal ws As Worksheet, ByVal have As Object, ByRef r As Long, ByVal lbl As String, _
                         ByVal vn As String, ByVal pr As String, ByRef n As Long)
    Dim k As String
    If vn = "" Then Exit Sub
    k = lbl & "|" & vn & "|" & pr
    If have.Exists(k) Then Exit Sub
    have(k) = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).NumberFormat = "@"
    ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-" & (VD_FIRST - 1) & ")"
    ws.Cells(r, 2).Value = lbl
    ws.Cells(r, 3).Value = vn
    ws.Cells(r, 4).Value = pr
    ws.Cells(r, 5).Value = "既存から取得"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Borders.LineStyle = xlContinuous
    r = r + 1
    n = n + 1
End Sub

Private Function CatLabel(ByVal p As String) As String
    Select Case p
        Case "am": CatLabel = "マルウェア対策"
        Case "db": CatLabel = "ディスクバックアップ"
        Case "de": CatLabel = "ディスク暗号化"
        Case "fw": CatLabel = "ファイアウォール"
        Case "pm": CatLabel = "パッチ管理"
        Case "dlp": CatLabel = "データ損失防止"
    End Select
End Function

' mobile_device など：オブジェクトを「項目: 値」の行にする（キーは並べ替え、criteria は省略、name付き配列は [name]）
Private Function FlatText(ByVal v As Object) As String
    Dim out As New Collection
    If v Is Nothing Then Exit Function
    FlatAdd v, "", out
    FlatText = JoinColl(out, vbLf)
End Function

Private Sub FlatAdd(ByVal v As Variant, ByVal pre As String, ByVal out As Collection)
    Dim keys As Variant, k As Variant, x As Variant, i As Long, rest As Object, p As String
    If IsObject(v) Then
        If TypeName(v) = "Dictionary" Then
            If v.Count = 0 Then out.Add IIf(pre = "", "{}", pre & ": {}"): Exit Sub
            keys = SortedKeys(v)
            For Each k In keys
                If k = "criteria" Then
                    FlatAdd v(k), pre, out
                Else
                    FlatAdd v(k), IIf(pre = "", "", pre & ".") & k, out
                End If
            Next
        Else
            i = 0
            For Each x In v
                i = i + 1
                If IsObject(x) Then
                    If TypeName(x) = "Dictionary" Then
                        If x.Exists("name") Then
                            p = pre & "[" & Nz(x, "name") & "]"
                            Set rest = CreateObject("Scripting.Dictionary")
                            For Each k In x.Keys
                                If k <> "name" Then
                                    If IsObject(x(k)) Then Set rest(k) = x(k) Else rest(k) = x(k)
                                End If
                            Next
                            If rest.Count = 0 Then out.Add p Else FlatAdd rest, p, out
                        Else
                            FlatAdd x, pre & "[" & i & "]", out
                        End If
                    Else
                        FlatAdd x, pre & "[" & i & "]", out
                    End If
                Else
                    out.Add pre & "[" & i & "]: " & ValText(x)
                End If
            Next
        End If
    Else
        out.Add pre & ": " & ValText(v)
    End If
End Sub

Private Function ValText(ByVal x As Variant) As String
    If VarType(x) = vbBoolean Then
        ValText = IIf(x, "true", "false")
    ElseIf IsNull(x) Then
        ValText = "null"
    Else
        ValText = CStr(x)
    End If
End Function

Private Function SortedKeys(ByVal d As Object) As Variant
    Dim a As Variant, i As Long, j As Long, t As Variant
    a = d.Keys
    For i = 1 To UBound(a)
        t = a(i)
        j = i - 1
        Do While j >= 0
            If StrComp(a(j), t, vbBinaryCompare) <= 0 Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next
    SortedKeys = a
End Function

Private Function Child(ByVal d As Object, ByVal k As String) As Object
    If d Is Nothing Then Exit Function
    If TypeName(d) <> "Dictionary" Then Exit Function
    If Not d.Exists(k) Then Exit Function
    If IsObject(d(k)) Then Set Child = d(k)
End Function

Private Function CollCount(ByVal d As Object, ByVal k As String) As Long
    If Not d.Exists(k) Then Exit Function
    If IsObject(d(k)) Then CollCount = d(k).Count
End Function

Private Function YN(ByVal v As Variant) As String
    If VarType(v) = vbBoolean Then
        YN = IIf(v, "yes", "no")
    ElseIf LCase$(CStr(v)) = "yes" Or LCase$(CStr(v)) = "true" Then
        YN = "yes"
    Else
        YN = "no"
    End If
End Function

Private Function NegMark(ByVal d As Object) As String
    If d.Exists("negate") Then
        If VarType(d("negate")) = vbBoolean Then
            If d("negate") Then NegMark = "!"
        ElseIf LCase$(CStr(d("negate"))) = "yes" Or LCase$(CStr(d("negate"))) = "true" Then
            NegMark = "!"
        End If
    End If
End Function

Private Function AddLine(ByVal s As String, ByVal t As String) As String
    If s = "" Then AddLine = t Else AddLine = s & vbLf & t
End Function

Private Function JoinColl(ByVal c As Variant, ByVal sep As String) As String
    Dim x As Variant, s As String, first As Boolean
    first = True
    For Each x In c
        If Not first Then s = s & sep
        s = s & CStr(x)
        first = False
    Next
    JoinColl = s
End Function

' match 式の中の "名前" を順に取り出す（重複は1回）
Private Function QuotedNames(ByVal m As String) As Collection
    Dim c As New Collection, seen As Object, i As Long, q1 As Long, q2 As Long, nm As String
    Set seen = CreateObject("Scripting.Dictionary")
    i = 1
    Do
        q1 = InStr(i, m, """")
        If q1 = 0 Then Exit Do
        q2 = InStr(q1 + 1, m, """")
        If q2 = 0 Then Exit Do
        nm = Mid$(m, q1 + 1, q2 - q1 - 1)
        If nm <> "" And Not seen.Exists(nm) Then
            seen(nm) = True
            c.Add nm
        End If
        i = q2 + 1
    Loop
    Set QuotedNames = c
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

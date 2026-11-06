Attribute VB_Name = "modApp"
'==========================================================
' modApp : SCM (Prisma Access Configuration API) Application / ApplicationGroup 連携
'   前提 : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'==========================================================
' ※ 共通処理（認証・HTTP・ログ・差分色付け等）は modScmCommon にあります。modScmCommon（JSONパーサを含む）も併せてインポートしてください。
'   メンバー選択画面（ダブルクリック）: clsPickerEvents をインポートし、frmMemberPick は空のユーザーフォームを挿入して frmMemberPick_コード.txt を貼り付け
Option Explicit

Private Const SH_CFG As String = "基本設定"
Private Const SH_EXIST As String = "AppObj_既存一覧"
Private Const SH_INPUT As String = "AppObj_登録変更削除"
Private Const SH_EXISTG As String = "AppGroupObj_既存一覧"
Private Const SH_INPUTG As String = "AppGroupObj_登録変更削除"
Private Const SH_REF As String = "AppRef_定義済み・照会"
Private Const SH_JSON As String = "JSON_CURL"
Private Const SH_DIFF As String = "AppObj_差分確認"
Private Const SH_DIFFG As String = "AppGroupObj_差分確認"
Private Const SH_LOG As String = "実行ログ"

Private Const EX_FIRST As Long = 8      ' 既存一覧 データ開始行
Private Const IN_FIRST As Long = 8      ' 登録変更削除 データ開始行
Private Const IN_LAST As Long = 507     ' 登録変更削除 データ最終行
Private Const IN_JUDGE As Long = 39     ' ③ 判定列(AM)
Private Const ING_JUDGE As Long = 10    ' ⑦ 判定列(J)
Private Const REF_FIRST As Long = 8     ' AppRef データ開始行
Private Const REF_LAST As Long = 10007  ' AppRef データ最終行（関数の参照範囲）
Private Const JS_FIRST As Long = 10     ' JSON_CURL データ開始行
Private Const DIFF_FIRST As Long = 8    ' 差分確認 データ開始行
Private Const LOG_FIRST As Long = 6     ' 実行ログ データ開始行
Private Const PAGE_LIMIT As Long = 200
Private Const EP_GROUP As String = "/sse/config/v1/application-groups"
Private Const EP_FILTER As String = "/sse/config/v1/application-filters"
Private Const REF_NONE As String = "照会:なし"
Private Const REF_FOUND As String = "照会:あり"

' ③ AppObj_登録変更削除 / ② AppObj_既存一覧 の項目列（同じ列に並ぶ）
'   D=category E=subcategory F=technology G=risk H=TCPポート I=UDPポート J=IPプロトコル
'   K=ICMPタイプ L=ICMPコード M=ICMPv6タイプ N=ICMPv6コード O=description P=parent_app
'   Q～U=timeout系（TmoKeys の順） V～AH=特性・オプション（FlagKeys の順、○=true）
' ③ その他 : B=操作 C=name AJ=既存id AK=既存行(②の何件目か)
' ② その他 : B=id C=name AI=コンテナ AJ=引継ぎ項目(JSON) AK=定義(関数)
' ⑦ AppGroupObj_登録変更削除 : B=操作 C=name D=メンバー F=既存id
' ⑥ AppGroupObj_既存一覧     : B=id C=name D=メンバー E=コンテナ
' AppRef                      : B=区分 C=name D=補足
Private Const COL_F1 As Long = 4        ' 項目の先頭列(D)
Private Const N_FIELD As Long = 31      ' 項目数(D～AH)
Private Const COL_EXID As Long = 36     ' ③ 既存id(AJ)
Private Const COL_EXROW As Long = 37    ' ③ 既存行(AK)
Private Const COL_EXTRA As Long = 36    ' ② 引継ぎ項目(AJ)
Private Const EXTRA_MAX As Long = 32000 ' セルに入る長さの目安

' ---- メンバー選択画面 用 ----
Private Const PICK_KEY As String = "^+m"            ' Ctrl+Shift+M
Private Const CELL_MAX As Long = 32767               ' Excelのセル文字数上限
Private Const SH_OVERFLOW As String = "ピッカー_選択結果"

Private mEvents As clsPickerEvents
Private mNames As Object     ' 候補 name -> 種別ラベル（追加順を保持）
Private mExclude As String   ' 候補から外す name（自分自身）

Private Function TmoKeys() As Variant
    TmoKeys = Array("timeout", "tcp_timeout", "udp_timeout", "tcp_half_closed_timeout", "tcp_time_wait_timeout")
End Function

Private Function FlagKeys() As Variant
    FlagKeys = Array("evasive_behavior", "consume_big_bandwidth", "used_by_malware", "able_to_transfer_file", _
                     "has_known_vulnerability", "tunnel_other_application", "prone_to_misuse", "pervasive_use", _
                     "tunnel_applications", "data_ident", "file_type_ident", "virus_ident", "no_appid_caching")
End Function

' このツールで編集しない項目（変更時はSCMの値をそのまま送る）
Private Function ExtraKeys() As Variant
    ExtraKeys = Array("signature", "alg_disable_capability")
End Function

'----------------------------------------------------------
' ボタン配置（初回のみ実行）
'----------------------------------------------------------
Public Sub SetupButtons()
    AddBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    AddBtn SH_EXIST, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUT, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJson"
    AddBtn SH_INPUT, "I3:J4", "名前照会（API）", "LookupNames"
    AddBtn SH_INPUT, "L3:M4", "入力クリア", "ClearInput"
    AddBtn SH_EXISTG, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUTG, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJsonGroup"
    AddBtn SH_INPUTG, "I3:J4", "名前照会（API）", "LookupNames"
    AddBtn SH_INPUTG, "L3:M4", "入力クリア", "ClearInputGroup"
    AddBtn SH_REF, "D3:D4", "名前照会（API）", "LookupNames"
    AddBtn SH_JSON, "G4:H4", "API実行（全行）", "ExecuteAll"
    AddBtn SH_JSON, "G6:H6", "クリア", "ClearJson"
    RemoveBtn SH_JSON, "PushConfig"     ' Push機能は削除（PushはSCM画面から実施）。旧ボタンを消す
    AddBtn SH_DIFF, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_DIFFG, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_LOG, "F3", "ログクリア", "ClearLog"
    RemoveBtn SH_LOG, "CheckJobs"      ' ジョブ状況確認も削除。旧ボタンを消す
    PickerStart True
    MsgBox "ボタンを配置しました。", vbInformation
End Sub

'----------------------------------------------------------
' ②⑥ 既存情報取得（Application・ApplicationGroup・フィルタを GET → 表生成）
'   定義済みApp-ID（APIが返した場合）とフィルタ名は AppRef シートに入れる
'----------------------------------------------------------
Public Sub FetchExisting()
    Dim apps As Collection, grps As Collection, flts As Collection, custom As New Collection, pre As New Collection
    Dim item As Object
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Set apps = FetchAllItems(BaseUrl(), "既存取得", True)
    If apps Is Nothing Then Exit Sub
    Set grps = FetchAllItems(EpUrl(EP_GROUP), "既存Group取得", True)
    If grps Is Nothing Then Exit Sub
    Set flts = FetchAllItems(EpUrl(EP_FILTER), "フィルタ取得", True)
    If flts Is Nothing Then Exit Sub
    For Each item In apps
        If IsPredefined(item) Then pre.Add Nz(item, "name") Else custom.Add item
    Next
    Application.ScreenUpdating = False
    WriteExistApp custom
    WriteExistGroup grps
    RefreshRef flts, pre
    Application.ScreenUpdating = True
    MsgBox "Application " & custom.Count & " 件、ApplicationGroup " & grps.Count & " 件、フィルタ " & flts.Count & " 件を取得しました。" & _
           IIf(pre.Count > 0, vbLf & "定義済みApp-ID " & pre.Count & " 件をAppRefシートに反映しました。", ""), vbInformation
End Sub

' 定義済みApp-ID か（snippet/folder に predefined を含む、または id が無い）
Private Function IsPredefined(ByVal item As Object) As Boolean
    Dim s As String
    s = LCase$(Nz(item, "snippet") & "|" & Nz(item, "folder"))
    IsPredefined = (InStr(s, "predefined") > 0 Or Nz(item, "id") = "")
End Function

Private Sub WriteExistApp(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long, v As Variant, i As Long
    Set ws = Worksheets(SH_EXIST)
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 37)).Clear
    r = EX_FIRST
    For Each item In items
        v = AppValues(item)
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 36)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        For i = 0 To N_FIELD - 1
            ws.Cells(r, COL_F1 + i).Value = v(i)
        Next
        ws.Cells(r, 35).Value = ContainerOf(item)
        ws.Cells(r, COL_EXTRA).Value = AppExtra(item)
        ws.Cells(r, 37).Formula = "=IF(C" & r & "="""","""",D" & r & "&""|""&E" & r & "&""|""&F" & r & "&""|""&G" & r & _
            "&""|""&H" & r & "&""|""&I" & r & "&""|""&J" & r & "&""|""&K" & r & "&""|""&L" & r & "&""|""&M" & r & _
            "&""|""&N" & r & "&""|""&O" & r & "&""|""&P" & r & "&""|""&Q" & r & "&""|""&R" & r & "&""|""&S" & r & _
            "&""|""&T" & r & "&""|""&U" & r & "&""|""&V" & r & "&""|""&W" & r & "&""|""&X" & r & "&""|""&Y" & r & _
            "&""|""&Z" & r & "&""|""&AA" & r & "&""|""&AB" & r & "&""|""&AC" & r & "&""|""&AD" & r & "&""|""&AE" & r & _
            "&""|""&AF" & r & "&""|""&AG" & r & "&""|""&AH" & r & ")"
        r = r + 1
    Next
    If r > EX_FIRST Then
        ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 37)).Borders.LineStyle = xlContinuous
    End If
    Range("EX_FETCHED_AT").Value = Now
End Sub

Private Sub WriteExistGroup(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long
    Set ws = Worksheets(SH_EXISTG)
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 5)).Clear
    r = EX_FIRST
    For Each item In items
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        ws.Cells(r, 4).Value = JoinArr(item, "members")
        ws.Cells(r, 5).Value = ContainerOf(item)
        r = r + 1
    Next
    If r > EX_FIRST Then
        ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 5)).Borders.LineStyle = xlContinuous
    End If
    Range("EXG_FETCHED_AT").Value = Now
End Sub

' 1件分の項目を ③②の D～AH 列の順の配列（0～30）に
Private Function AppValues(ByVal item As Object) As Variant
    Dim v(30) As String, d As Object, p As Variant, s As String, pos As Long, ks As Variant, i As Long
    v(0) = Nz(item, "category")
    v(1) = Nz(item, "subcategory")
    v(2) = Nz(item, "technology")
    v(3) = Nz(item, "risk")
    If item.Exists("default") Then
        If IsObject(item("default")) Then
            Set d = item("default")
            If d.Exists("port") Then
                If IsObject(d("port")) Then
                    For Each p In d("port")
                        s = CStr(p)
                        pos = InStr(s, "/")
                        If pos > 0 Then
                            Select Case LCase$(Left$(s, pos - 1))
                                Case "tcp": v(4) = JoinStr(v(4), Mid$(s, pos + 1))
                                Case "udp": v(5) = JoinStr(v(5), Mid$(s, pos + 1))
                            End Select
                        End If
                    Next
                End If
            End If
            v(6) = Nz(d, "ident_by_ip_protocol")
            IcmpValues d, "ident_by_icmp_type", v(7), v(8)
            IcmpValues d, "ident_by_icmp6_type", v(9), v(10)
        End If
    End If
    v(11) = Nz(item, "description")
    v(12) = Nz(item, "parent_app")
    ks = TmoKeys()
    For i = 0 To 4
        v(13 + i) = Nz(item, CStr(ks(i)))
    Next
    ks = FlagKeys()
    For i = 0 To 12
        If IsTrue(item, CStr(ks(i))) Then v(18 + i) = "○"
    Next
    AppValues = v
End Function

Private Sub IcmpValues(ByVal d As Object, ByVal key As String, ByRef icmpType As String, ByRef icmpCode As String)
    Dim ic As Object
    If Not d.Exists(key) Then Exit Sub
    If Not IsObject(d(key)) Then Exit Sub
    Set ic = d(key)
    icmpType = Nz(ic, "type")
    icmpCode = Nz(ic, "code")
End Sub

Private Function IsTrue(ByVal item As Object, ByVal key As String) As Boolean
    Dim x As Variant
    If Not item.Exists(key) Then Exit Function
    If IsObject(item(key)) Then Exit Function
    x = item(key)
    If IsNull(x) Then Exit Function
    If VarType(x) = vbBoolean Then
        IsTrue = x
    Else
        IsTrue = (LCase$(CStr(x)) = "true" Or LCase$(CStr(x)) = "yes")
    End If
End Function

' このツールで編集しない項目をJSONの断片（"signature":[..],"alg_disable_capability":".."）に
Private Function AppExtra(ByVal item As Object) As String
    Dim k As Variant, s As String
    For Each k In ExtraKeys()
        If item.Exists(k) Then
            If s <> "" Then s = s & ","
            s = s & JsonStr(CStr(k)) & ":" & ToJson(item(k))
        End If
    Next
    If Len(s) > EXTRA_MAX Then
        WriteLog "既存取得", Nz(item, "name"), "", "", 0, "signature等が " & Len(s) & " 文字のためセルに保存できません（このツールでは変更不可）"
        s = "#TOO_LONG"
    End If
    AppExtra = s
End Function

' 解析済みの値（Dictionary / Collection / 値）をJSON文字列に戻す
Private Function ToJson(ByVal v As Variant) As String
    Dim k As Variant, x As Variant, s As String
    If IsObject(v) Then
        If TypeName(v) = "Dictionary" Then
            For Each k In v.Keys
                If s <> "" Then s = s & ","
                s = s & JsonStr(CStr(k)) & ":" & ToJson(v(k))
            Next
            ToJson = "{" & s & "}"
        Else
            For Each x In v
                If s <> "" Then s = s & ","
                s = s & ToJson(x)
            Next
            ToJson = "[" & s & "]"
        End If
    ElseIf IsNull(v) Then
        ToJson = "null"
    ElseIf VarType(v) = vbBoolean Then
        ToJson = IIf(v, "true", "false")
    ElseIf VarType(v) = vbString Then
        ToJson = JsonStr(CStr(v))
    Else
        ToJson = Trim$(Str$(v))
    End If
End Function

'----------------------------------------------------------
' AppRef シート（区分・name・補足）の読み書き
'----------------------------------------------------------
Private Function LoadRef() As Collection
    Dim ws As Worksheet, r As Long, res As New Collection
    Set ws = Worksheets(SH_REF)
    For r = REF_FIRST To LastRow(ws, 3)
        If CStr(ws.Cells(r, 3).Value) <> "" Then
            res.Add Array(CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), CStr(ws.Cells(r, 4).Value))
        End If
    Next
    Set LoadRef = res
End Function

Private Sub SaveRef(ByVal rows As Collection)
    Dim ws As Worksheet, n As Long, i As Long, arr() As Variant, rw As Variant
    Set ws = Worksheets(SH_REF)
    n = rows.Count
    If n > REF_LAST - REF_FIRST + 1 Then
        MsgBox "AppRefシートが " & n & " 行になり、判定に使う範囲（" & (REF_LAST - REF_FIRST + 1) & " 行）を超えました。" & vbLf & _
               "超えた分は書き込みません。「定義済み」の不要な行を整理してください。", vbExclamation
        n = REF_LAST - REF_FIRST + 1
    End If
    With ws.Range(ws.Cells(REF_FIRST, 1), ws.Cells(REF_LAST, 4))
        .ClearContents
        .Borders.LineStyle = xlNone
    End With
    If n = 0 Then Exit Sub
    ReDim arr(1 To n, 1 To 3)
    i = 0
    For Each rw In rows
        i = i + 1
        If i > n Then Exit For
        arr(i, 1) = rw(0): arr(i, 2) = rw(1): arr(i, 3) = rw(2)
    Next
    ws.Range(ws.Cells(REF_FIRST, 2), ws.Cells(REF_FIRST + n - 1, 4)).NumberFormat = "@"
    ws.Range(ws.Cells(REF_FIRST, 2), ws.Cells(REF_FIRST + n - 1, 4)).Value = arr
    ws.Range(ws.Cells(REF_FIRST, 1), ws.Cells(REF_FIRST + n - 1, 1)).Formula = "=IF(C" & REF_FIRST & "="""","""",ROW()-7)"
    ws.Range(ws.Cells(REF_FIRST, 1), ws.Cells(REF_FIRST + n - 1, 4)).Borders.LineStyle = xlContinuous
End Sub

' 既存情報取得後：フィルタ行・API取得の定義済み行を作り直し、②⑥に入った名前の照会結果を消す
Private Sub RefreshRef(ByVal flts As Collection, ByVal pre As Collection)
    Dim known As Object, rows As Collection, keep As New Collection, rw As Variant, item As Object, nm As Variant
    Set known = ExistingNames()
    For Each item In flts
        known(Nz(item, "name")) = True
    Next
    Set rows = LoadRef()
    For Each rw In rows
        If rw(0) = "フィルタ" Then
        ElseIf rw(0) = "定義済み" And Left$(rw(2), 5) = "API取得" Then
        ElseIf (rw(0) = REF_NONE Or rw(0) = REF_FOUND) And known.Exists(rw(1)) Then
        Else
            keep.Add rw
        End If
    Next
    For Each item In flts
        keep.Add Array("フィルタ", Nz(item, "name"), "API取得（application-filters）")
    Next
    For Each nm In pre
        keep.Add Array("定義済み", CStr(nm), "API取得（applications）")
    Next
    SaveRef keep
End Sub

' ②⑥の既存の名前（大文字・小文字を区別）
Private Function ExistingNames() As Object
    Dim d As Object, ws As Variant, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    For Each ws In Array(Worksheets(SH_EXIST), Worksheets(SH_EXISTG))
        For r = EX_FIRST To LastRow(ws, 3)
            If CStr(ws.Cells(r, 3).Value) <> "" Then d(CStr(ws.Cells(r, 3).Value)) = True
        Next
    Next
    Set ExistingNames = d
End Function

'----------------------------------------------------------
' 名前照会（API）
'   ③⑦の登録名と、⑦のメンバーのうちどこにも見つからない名前を、
'   applications / application-groups / application-filters に name 指定で問い合わせ、
'   結果を AppRef シートに「照会:あり／照会:なし」で書く（同じ名前の前回の結果は置き換え）
'----------------------------------------------------------
Public Sub LookupNames()
    Dim wsA As Worksheet, wsG As Worksheet, exNames As Object, known As Object, targets As Object, refRows As Collection
    Dim r As Long, nm As String, parts As Variant, i As Long, rw As Variant, k As Variant
    Dim res As Long, where As String, keep As New Collection, newRows As New Collection, nFound As Long, nNone As Long, nErr As Long
    If Not CheckContainer() Then Exit Sub
    Application.Calculate
    Set wsA = Worksheets(SH_INPUT)
    Set wsG = Worksheets(SH_INPUTG)
    Set exNames = ExistingNames()
    Set refRows = LoadRef()
    ' メンバーとして既に分かっている名前
    Set known = CreateObject("Scripting.Dictionary")
    For Each k In exNames.Keys
        known(k) = True
    Next
    For Each rw In refRows
        If rw(0) <> REF_NONE Then known(rw(1)) = True
    Next
    For r = IN_FIRST To IN_LAST
        If CStr(wsA.Cells(r, 2).Value) = "登録" And CStr(wsA.Cells(r, 3).Value) <> "" Then known(CStr(wsA.Cells(r, 3).Value)) = True
        If CStr(wsG.Cells(r, 2).Value) = "登録" And CStr(wsG.Cells(r, 3).Value) <> "" Then known(CStr(wsG.Cells(r, 3).Value)) = True
    Next
    ' 照会する名前
    Set targets = CreateObject("Scripting.Dictionary")
    For r = IN_FIRST To IN_LAST
        nm = CStr(wsA.Cells(r, 3).Value)
        If CStr(wsA.Cells(r, 2).Value) = "登録" And nm <> "" And Not exNames.Exists(nm) And Not IsRefDefined(refRows, nm) Then targets(nm) = True
        nm = CStr(wsG.Cells(r, 3).Value)
        If CStr(wsG.Cells(r, 2).Value) = "登録" And nm <> "" And Not exNames.Exists(nm) And Not IsRefDefined(refRows, nm) Then targets(nm) = True
        If nm <> "" And CStr(wsG.Cells(r, 2).Value) <> "削除" Then
            parts = SplitMembers(CStr(wsG.Cells(r, 4).Value))
            For i = 0 To UBound(parts)
                If Not known.Exists(CStr(parts(i))) Then targets(CStr(parts(i))) = True
            Next
        End If
    Next
    If targets.Count = 0 Then
        MsgBox "照会が必要な名前はありません（③⑦の登録名と⑦のメンバーはすべて確認済みです）。", vbInformation
        Exit Sub
    End If
    If MsgBox(targets.Count & " 件の名前をSCMに照会します（1件につきGETを最大3回）。よろしいですか？", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    For Each k In targets.Keys
        Application.StatusBar = "名前照会中：" & k
        res = QueryName(CStr(k), where)
        Select Case res
            Case 1
                ' 照会で見つかったもの（②⑥に無い＝定義済みApp-ID等）
                newRows.Add Array(REF_FOUND, CStr(k), "名前照会 " & Format$(Now, "yyyy/mm/dd hh:nn") & " " & where & " に存在")
                nFound = nFound + 1
            Case 0
                newRows.Add Array(REF_NONE, CStr(k), "名前照会 " & Format$(Now, "yyyy/mm/dd hh:nn") & " applications / application-groups / application-filters に無し")
                nNone = nNone + 1
            Case Else
                nErr = nErr + 1
        End Select
    Next
    Application.StatusBar = False
    ' 前回の照会結果のうち、今回照会した名前のものは置き換える
    For Each rw In refRows
        If (rw(0) = REF_NONE Or rw(0) = REF_FOUND) And targets.Exists(rw(1)) Then
            If Not FailedName(newRows, rw(1)) Then GoTo NextRow
        End If
        keep.Add rw
NextRow:
    Next
    For Each rw In newRows
        keep.Add rw
    Next
    Application.ScreenUpdating = False
    SaveRef keep
    Application.ScreenUpdating = True
    Application.Calculate
    MsgBox "名前照会が終わりました。" & vbLf & "SCMに存在：" & nFound & " 件／存在しない：" & nNone & " 件" & _
           IIf(nErr > 0, "／照会失敗：" & nErr & " 件（実行ログを確認してください）", "") & vbLf & _
           "結果はAppRefシートに書き込み、③⑦の判定に反映しました。", IIf(nErr > 0, vbExclamation, vbInformation)
End Sub

' 照会に失敗した名前か（今回の結果に無い＝前回の結果を残す）
Private Function FailedName(ByVal newRows As Collection, ByVal nm As String) As Boolean
    Dim rw As Variant
    For Each rw In newRows
        If rw(1) = nm Then Exit Function
    Next
    FailedName = True
End Function

' AppRef に定義済み／フィルタ／照会:ありとして載っているか
Private Function IsRefDefined(ByVal refRows As Collection, ByVal nm As String) As Boolean
    Dim rw As Variant
    For Each rw In refRows
        If rw(1) = nm And rw(0) <> REF_NONE And rw(0) <> REF_FOUND Then
            IsRefDefined = True
            Exit Function
        End If
    Next
End Function

' 1件の名前を3種類のAPIに照会。1=存在 0=無し -1=失敗。where に見つかったAPI
Private Function QueryName(ByVal nm As String, ByRef where As String) As Long
    Dim ep As Variant, url As String, st As Long, resText As String, res As Object, item As Object
    where = ""
    For Each ep In Array(BaseUrl(), EpUrl(EP_GROUP), EpUrl(EP_FILTER))
        url = WithQuery(CStr(ep), ContainerQuery() & "&name=" & Application.WorksheetFunction.EncodeURL(nm))
        st = HttpCall("GET", url, "", resText)
        If st = 200 Then
            Set res = ParseJson(resText)
            If Not res Is Nothing Then
                If res.Exists("data") Then
                    For Each item In res("data")
                        ' 大文字・小文字の違いだけの名前も重複として扱う
                        If LCase$(Nz(item, "name")) = LCase$(nm) Then
                            where = Mid$(CStr(ep), InStrRev(CStr(ep), "/") + 1)
                            WriteLog "名前照会", nm, "GET", url, st, where & " に存在（" & Nz(item, "name") & "）"
                            QueryName = 1
                            Exit Function
                        End If
                    Next
                End If
            End If
        ElseIf st <> 404 Then
            WriteLog "名前照会", nm, "GET", url, st, ErrSummary(resText)
            QueryName = -1
            Exit Function
        End If
    Next
    WriteLog "名前照会", nm, "GET", WithQuery(BaseUrl(), ContainerQuery() & "&name=" & Application.WorksheetFunction.EncodeURL(nm)), 200, "3種類とも存在しない"
    QueryName = 0
End Function

'----------------------------------------------------------
' ③⑦ JSON発行（AppObj_登録変更削除 + AppGroupObj_登録変更削除 → JSON_CURL）
'   どちらのボタンでも③と⑦の両方から④を作り直す。
'   実行順：App登録・変更 → Group登録・変更 → Group自動変更 → Group削除 → App削除
'   ・削除するApp／GroupがほかのGroupのメンバーなら、先にGroupから外す変更を追加する。
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

    ' 削除予定（App と Group。名前は共通）、⑦の対象Group
    Set delN = CreateObject("Scripting.Dictionary")
    Set grpIn = CreateObject("Scripting.Dictionary")
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) Then
            If CStr(wsA.Cells(r, 2).Value) = "削除" Then delN(CStr(wsA.Cells(r, 3).Value)) = "App"
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
        MsgBox "削除するApplication／Groupを外すと、次のGroupのメンバーが0件になります（メンバーは1件以上必須）。" & vbLf & _
               "⑦でGroupを削除するか、残すメンバーを追加してから発行してください。" & emptyG, vbExclamation
        Exit Sub
    End If
    If note <> "" Then
        If MsgBox("削除するApplication／Groupが、次のGroupのメンバーに入っています。" & note & vbLf & vbLf & _
                  "Groupを先に変更（または削除）し、その後に削除する順で④に出力します。続けますか？", _
                  vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    Set addOrder = OrderAddGroups(wsG, grpIn)
    If addOrder Is Nothing Then Exit Sub
    Set delOrder = OrderDelGroups(wsG, wsEG, grpIn)

    ClearJsonRows wsOut
    o = JS_FIRST
    ' 1. Application 登録・変更
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) And CStr(wsA.Cells(r, 2).Value) <> "削除" Then WriteJsonRowApp wsA, r, wsOut, o
    Next
    ' 2. Group 登録・変更（メンバーに入るGroupを先に。削除予定のメンバーは外す）
    For Each rw In addOrder
        WriteJsonRowGroup wsG, CLng(rw), wsOut, o, delN
    Next
    ' 3. 削除予定をGroupから外す自動変更
    For Each k In autoG.Keys
        r = autoG(k)
        WriteJsonRow wsOut, o, "変更(Group)自動", CStr(k), "PUT", EpUrl(EP_GROUP) & "/" & wsEG.Cells(r, 2).Value, _
                     BuildBodyGroup(CStr(k), StripMembers(CStr(wsEG.Cells(r, 4).Value), delN, removed))
    Next
    ' 4. Group 削除（親Group → 子Group）
    For Each rw In delOrder
        WriteJsonRowGroup wsG, CLng(rw), wsOut, o, delN
    Next
    ' 5. Application 削除
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) And CStr(wsA.Cells(r, 2).Value) = "削除" Then WriteJsonRowApp wsA, r, wsOut, o
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

Private Sub WriteJsonRowApp(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long)
    Dim op As String, nm As String, extra As String, exRow As Variant
    op = CStr(wsIn.Cells(r, 2).Value)
    nm = CStr(wsIn.Cells(r, 3).Value)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op, nm, "POST", WithQuery(BaseUrl(), ContainerQuery()), BuildBodyApp(wsIn, r, "")
        Case "変更"
            ' ②の signature 等をそのまま引き継ぐ
            exRow = wsIn.Cells(r, COL_EXROW).Value
            If IsNumeric(exRow) And CStr(exRow) <> "" Then
                extra = CStr(Worksheets(SH_EXIST).Cells(EX_FIRST + CLng(exRow) - 1, COL_EXTRA).Value)
            End If
            WriteJsonRow wsOut, o, op, nm, "PUT", BaseUrl() & "/" & wsIn.Cells(r, COL_EXID).Value, BuildBodyApp(wsIn, r, extra)
        Case "削除"
            WriteJsonRow wsOut, o, op, nm, "DELETE", BaseUrl() & "/" & wsIn.Cells(r, COL_EXID).Value, ""
    End Select
End Sub

Private Sub WriteJsonRowGroup(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long, ByVal delN As Object)
    Dim op As String, nm As String, members As String, removed As String
    op = CStr(wsIn.Cells(r, 2).Value)
    nm = CStr(wsIn.Cells(r, 3).Value)
    members = StripMembers(CStr(wsIn.Cells(r, 4).Value), delN, removed)
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "POST", WithQuery(EpUrl(EP_GROUP), ContainerQuery()), BuildBodyGroup(nm, members)
        Case "変更"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "PUT", EpUrl(EP_GROUP) & "/" & wsIn.Cells(r, 6).Value, BuildBodyGroup(nm, members)
        Case "削除"
            WriteJsonRow wsOut, o, op & "(Group)", nm, "DELETE", EpUrl(EP_GROUP) & "/" & wsIn.Cells(r, 6).Value, ""
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

' Application のボディ
'   {"name":..,"category":..,"subcategory":..,"technology":..,"risk":n,"description":..,"parent_app":..,
'    "default":{..},"timeout":n,..,"evasive_behavior":true,..（引継ぎ項目）}
Private Function BuildBodyApp(ByVal ws As Worksheet, ByVal r As Long, ByVal extra As String) As String
    Dim js As String, pl As String, x As String, ks As Variant, i As Long
    js = "{""name"":" & JsonStr(CellText(ws, r, 3)) & ",""category"":" & JsonStr(CellText(ws, r, 4)) & _
         ",""subcategory"":" & JsonStr(CellText(ws, r, 5)) & ",""technology"":" & JsonStr(CellText(ws, r, 6)) & _
         ",""risk"":" & CLng(CellText(ws, r, 7))
    If CellText(ws, r, 15) <> "" Then js = js & ",""description"":" & JsonStr(CStr(ws.Cells(r, 15).Value))
    If CellText(ws, r, 16) <> "" Then js = js & ",""parent_app"":" & JsonStr(CellText(ws, r, 16))
    pl = PortItems("tcp", CellText(ws, r, 8))
    x = PortItems("udp", CellText(ws, r, 9))
    If x <> "" Then
        If pl <> "" Then pl = pl & ","
        pl = pl & x
    End If
    If pl <> "" Then
        js = js & ",""default"":{""port"":[" & pl & "]}"
    ElseIf CellText(ws, r, 10) <> "" Then
        js = js & ",""default"":{""ident_by_ip_protocol"":" & JsonStr(CellText(ws, r, 10)) & "}"
    ElseIf CellText(ws, r, 11) <> "" Then
        js = js & ",""default"":{""ident_by_icmp_type"":" & IcmpJson(CellText(ws, r, 11), CellText(ws, r, 12)) & "}"
    ElseIf CellText(ws, r, 13) <> "" Then
        js = js & ",""default"":{""ident_by_icmp6_type"":" & IcmpJson(CellText(ws, r, 13), CellText(ws, r, 14)) & "}"
    End If
    ks = TmoKeys()
    For i = 0 To 4
        If CellText(ws, r, 17 + i) <> "" Then js = js & ",""" & ks(i) & """:" & CLng(CellText(ws, r, 17 + i))
    Next
    ks = FlagKeys()
    For i = 0 To 12
        If CellText(ws, r, 22 + i) = "○" Then js = js & ",""" & ks(i) & """:true"
    Next
    If extra <> "" And extra <> "#TOO_LONG" Then js = js & "," & extra
    BuildBodyApp = js & "}"
End Function

' ポート一覧 → "tcp/80,443",.. （1要素63文字以内になるよう分割）
Private Function PortItems(ByVal proto As String, ByVal lst As String) As String
    Dim parts() As String, i As Long, cur As String, res As String
    If lst = "" Then Exit Function
    If LCase$(lst) = "dynamic" Then
        PortItems = JsonStr(proto & "/dynamic")
        Exit Function
    End If
    parts = Split(lst, ",")
    For i = LBound(parts) To UBound(parts)
        If cur <> "" And Len(proto) + 1 + Len(cur) + 1 + Len(parts(i)) > 63 Then
            res = JoinStr(res, JsonStr(proto & "/" & cur))
            cur = parts(i)
        Else
            cur = JoinStr(cur, parts(i))
        End If
    Next
    If cur <> "" Then res = JoinStr(res, JsonStr(proto & "/" & cur))
    PortItems = res
End Function

Private Function IcmpJson(ByVal icmpType As String, ByVal icmpCode As String) As String
    IcmpJson = "{""type"":" & JsonStr(icmpType)
    If icmpCode <> "" Then IcmpJson = IcmpJson & ",""code"":" & JsonStr(icmpCode)
    IcmpJson = IcmpJson & "}"
End Function

' ApplicationGroup のボディ {"name":..,"members":[..]}（description・tagはAPI仕様に無い）
Private Function BuildBodyGroup(ByVal nm As String, ByVal members As String) As String
    Dim parts As Variant, t As String, i As Long
    parts = SplitMembers(members)
    For i = 0 To UBound(parts)
        If t <> "" Then t = t & ","
        t = t & JsonStr(CStr(parts(i)))
    Next
    BuildBodyGroup = "{""name"":" & JsonStr(nm) & ",""members"":[" & t & "]}"
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
           "AppObj_差分確認 と AppGroupObj_差分確認 に、現在の設定一覧と既存一覧（実行前）からの変化を出力しました。内容を確認してから、SCMの画面でPushしてください。", vbInformation
End Sub

'----------------------------------------------------------
' クリア系
'----------------------------------------------------------
Public Sub ClearInput()
    If MsgBox("入力（B～AH列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUT).Range("B" & IN_FIRST & ":AH" & IN_LAST).ClearContents
End Sub

Public Sub ClearInputGroup()
    If MsgBox("入力（B～D列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUTG).Range("B" & IN_FIRST & ":D" & IN_LAST).ClearContents
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
    Dim ws As Worksheet, r As Long, n As Long, opName As String, v(30) As String, i As Long
    Set before = CreateObject("Scripting.Dictionary")
    Set after = CreateObject("Scripting.Dictionary")
    Set ops = CreateObject("Scripting.Dictionary")
    Application.Calculate
    ' 実行前：②⑥既存一覧シート
    Set ws = Worksheets(SH_EXIST)
    For r = EX_FIRST To LastRow(ws, 3)
        If CStr(ws.Cells(r, 3).Value) <> "" Then
            For i = 0 To N_FIELD - 1
                v(i) = CStr(ws.Cells(r, COL_F1 + i).Value)
            Next
            before("App|" & ws.Cells(r, 3).Value) = DiffValues(v)
        End If
    Next
    Set ws = Worksheets(SH_EXISTG)
    For r = EX_FIRST To LastRow(ws, 3)
        If CStr(ws.Cells(r, 3).Value) <> "" Then
            before("Group|" & ws.Cells(r, 3).Value) = Array(CStr(ws.Cells(r, 4).Value))
        End If
    Next
    ' 実行後：SCMから取得（シートは書き換えない。定義済みApp-IDは対象外）
    Set items = FetchAllItems(BaseUrl(), "差分確認", False)
    If items Is Nothing Then Exit Function
    For Each item In items
        If Not IsPredefined(item) Then after("App|" & Nz(item, "name")) = DiffValues(AppValues(item))
    Next
    Set items = FetchAllItems(EpUrl(EP_GROUP), "差分確認", False)
    If items Is Nothing Then Exit Function
    For Each item In items
        after("Group|" & Nz(item, "name")) = Array(JoinArr(item, "members"))
    Next
    ' ④の操作
    Set ws = Worksheets(SH_JSON)
    For r = JS_FIRST To LastRow(ws, 4)
        opName = CStr(ws.Cells(r, 2).Value)
        If opName <> "" Then
            If InStr(opName, "(Group)") > 0 Then
                ops("Group|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
            Else
                ops("App|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), CStr(ws.Cells(r, 8).Value))
            End If
        End If
    Next
    Application.ScreenUpdating = False
    n = WriteDiffSheet(Worksheets(SH_DIFF), "App", _
                       Array("category", "subcategory", "technology", "risk", "default", "description", "parent_app", _
                             "timeout", "特性・オプション"), before, after, ops)
    n = n + WriteDiffSheet(Worksheets(SH_DIFFG), "Group", Array("メンバー"), before, after, ops)
    Application.ScreenUpdating = True
    WriteLog "差分確認", "", "GET", BaseUrl() & " / " & EpUrl(EP_GROUP), 200, n & "件の変化"
    BuildDiffCore = True
End Function

' D～AH列の値（AppValues と同じ並び）→ 差分確認の表示項目
Private Function DiffValues(ByVal v As Variant) As Variant
    Dim ks As Variant, i As Long, tmo As String, flg As String
    ks = TmoKeys()
    For i = 0 To 4
        If v(13 + i) <> "" Then tmo = Trim$(tmo & " " & ks(i) & "=" & v(13 + i))
    Next
    ks = FlagKeys()
    For i = 0 To 12
        If v(18 + i) = "○" Then flg = JoinStr(flg, CStr(ks(i)))
    Next
    DiffValues = Array(v(0), v(1), v(2), v(3), DefaultText(v), v(11), v(12), tmo, flg)
End Function

' default の表示用（例 tcp/80,443 udp/53、ip-protocol=47、icmp type=8 code=0）
Private Function DefaultText(ByVal v As Variant) As String
    Dim s As String
    If v(4) <> "" Or v(5) <> "" Then
        If v(4) <> "" Then s = "tcp/" & v(4)
        If v(5) <> "" Then s = Trim$(s & " udp/" & v(5))
    ElseIf v(6) <> "" Then
        s = "ip-protocol=" & v(6)
    ElseIf v(7) <> "" Then
        s = "icmp type=" & v(7) & IIf(v(8) <> "", " code=" & v(8), "")
    ElseIf v(9) <> "" Then
        s = "icmp6 type=" & v(9) & IIf(v(10) <> "", " code=" & v(10), "")
    End If
    DefaultText = s
End Function

' 1種類分（App / Group）：SCMの現在の全件と削除されたものを一覧表示し、既存一覧との変化を付ける。戻り値は変化のあった件数
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

Private Function EpUrl(ByVal ep As String) As String
    Dim b As String
    b = Cfg("CFG_BASE_URL")
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    EpUrl = b & ep
End Function

' セルの値（前後の空白を除く）
Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    CellText = Trim$(CStr(ws.Cells(r, c).Value))
End Function

' カンマ区切りで連結（a が空なら b）
Private Function JoinStr(ByVal a As String, ByVal b As String) As String
    If a = "" Then JoinStr = b Else JoinStr = a & "," & b
End Function

'==========================================================
' メンバー選択画面（ダブルクリックでチェックボックス一覧）
'   必要モジュール : modApp（このモジュール） / clsPickerEvents / frmMemberPick
'   ブックを開くと Auto_Open で自動有効化されます。
'   ダブルクリックが効かない時は PickerStart を実行（または Ctrl+Shift+M）
'   対象：AppGroupObj_登録変更削除 の D列（メンバー）。操作が「削除」の行は対象外
'==========================================================
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

' 対象セルならピッカーを表示して True を返す
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

' 対象列の定義
'   AddSrc シート名, name列, 開始行, 種別ラベル [, 条件列, 条件値]
'   AddMark シート名, name列, 開始行, 条件列, 条件値, 付記
'   候補は ⑦のI列（メンバー確認）と同じ範囲：②⑥の既存、③⑦で登録予定、AppRef の定義済み・フィルタ・照会:あり
Private Function PickerRule(ByVal ws As Worksheet, ByVal c As Range, ByRef ttl As String) As Boolean
    Select Case ws.Name
        Case SH_INPUTG
            If c.Column <> 4 Or c.Row < IN_FIRST Or c.Row > IN_LAST Then Exit Function
            If SafeStr(ws.Cells(c.Row, 2).Value) = "削除" Then Exit Function
            mExclude = Trim$(SafeStr(ws.Cells(c.Row, 3).Value))      ' 自分自身は候補に出さない
            AddSrc SH_EXIST, 3, EX_FIRST, "アプリ（カスタム）"
            AddSrc SH_INPUT, 3, IN_FIRST, "アプリ（③で登録）", 2, "登録"
            AddSrc SH_EXISTG, 3, EX_FIRST, "アプリグループ"
            AddSrc SH_INPUTG, 3, IN_FIRST, "アプリグループ（⑦で登録）", 2, "登録"
            AddSrc SH_REF, 3, REF_FIRST, "アプリフィルタ", 2, "フィルタ"
            AddSrc SH_REF, 3, REF_FIRST, "定義済みApp-ID（照会で確認）", 2, REF_FOUND
            AddSrc SH_REF, 3, REF_FIRST, "定義済みApp-ID", 2, "定義済み"
            AddMark SH_INPUT, 3, IN_FIRST, 2, "削除", " ※③で削除予定"
            AddMark SH_INPUTG, 3, IN_FIRST, 2, "削除", " ※⑦で削除予定"
            ttl = "ApplicationGroup メンバー選択：" & SafeStr(ws.Cells(c.Row, 3).Value)
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

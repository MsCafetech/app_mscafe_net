Attribute VB_Name = "modDecRule"
'==========================================================
' modDecRule : SCM（Prisma Access Configuration API）復号ポリシー（decryption-rules）
'              取得・表化 / 追加・変更・削除・移動 / JSON化 / 設定差分チェック
'   前提   : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   選択画面: clsPickerEvents をインポートし、frmMemberPick を作成（frmMemberPick_コード.txt）
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'==========================================================
Option Explicit

Private Const SH_CFG As String = "①基本設定"
Private Const SH_GET As String = "②DecRule_取得"
Private Const SH_SET As String = "③DecRule_設定"
Private Const SH_JSON As String = "④JSON化"
Private Const SH_DIFF As String = "⑤DecRule_設定差分チェック"
Private Const SH_LOG As String = "⑥実行ログ"
Private Const SH_PROF As String = "⑦DecProfile_一覧"
Private Const SH_CHO As String = "⑧DecRule_選択肢"

Private Const FIRST As Long = 8         ' 各シートのデータ開始行
Private Const JS_FIRST As Long = 10     ' JSON化 データ開始行
Private Const NF As Long = 21           ' name 以外の項目数
Private Const HDR As String = "▼"       ' スコープ見出し行の印（A列）

' ② DecRule_取得 の列（A=順番 B～D=スコープ（第1～第3階層） E=folder F=position G=id H=name I～AC=項目 AD=取得時JSON）
Private Const G_SEQ As Long = 1
Private Const G_SCOPE As Long = 2
Private Const G_FOLDER As Long = 5
Private Const G_POS As Long = 6
Private Const G_ID As Long = 7
Private Const G_NAME As Long = 8
Private Const G_F0 As Long = 9
Private Const G_JSON As Long = 30
Private Const G_LAST As Long = 30
Private Const G_HF As Long = 32         ' 見出し行の folder / position（非表示列 AF・AG）

' ③ DecRule_設定 の列（A=順番 B=操作 C=判定 D=理由 E～G=スコープ（第1～第3階層） H=folder I=position J=id K=name L～AF=項目）
Private Const S_OP As Long = 2
Private Const S_JUDGE As Long = 3
Private Const S_REASON As Long = 4
Private Const S_SCOPE As Long = 5
Private Const S_FOLDER As Long = 8
Private Const S_POS As Long = 9
Private Const S_ID As Long = 10
Private Const S_NAME As Long = 11
Private Const S_F0 As Long = 12
Private Const S_LAST As Long = 32
Private Const S_HF As Long = 34         ' 見出し行の folder / position（非表示列 AH・AI）

' ⑤ 差分チェック の列（A=No B～D=スコープ（第1～第3階層） E=変化 F=順番 G=name H～AB=項目 AC=変更項目 AD=変更前 AE=操作 AF=HTTP AG=判定）
Private Const D_SCOPE As Long = 2
Private Const D_CHG As Long = 5
Private Const D_ORDER As Long = 6
Private Const D_NAME As Long = 7
Private Const D_FIELDS As Long = 29
Private Const D_BEFORE As Long = 30
Private Const D_OP As Long = 31
Private Const D_HTTP As Long = 32
Private Const D_JUDGE As Long = 33
Private Const D_LAST As Long = 33
Private Const D_HF As Long = 35

' ④ JSON化 の列 : A=No B=操作 C=name D=メソッド E=URL F=JSON G=curl H=実行結果 I=スコープ(folder|position) J=対象id K=移動先 L=追加後のid
Private Const J_LAST As Long = 12

' 選択肢シートの列
Private Const CH_ZONE As Long = 2
Private Const CH_ADDR As Long = 3
Private Const CH_ADDRG As Long = 4
Private Const CH_EDL As Long = 5
Private Const CH_REGION As Long = 6
Private Const CH_SVC As Long = 7
Private Const CH_SVCG As Long = 8
Private Const CH_URLC As Long = 9
Private Const CH_URLP As Long = 10
Private Const CH_TAG As Long = 11
Private Const CH_HIP As Long = 12
Private Const CH_USER As Long = 13
Private Const CH_CERT As Long = 14
Private Const CH_EDLTYPE As Long = 16

' 色（RGB の BGR 表記）
Private Const C_IN As Long = &HCCF2FF       ' 入力セル FFF2CC
Private Const C_FX As Long = &HF2F2F2       ' 自動 F2F2F2
Private Const C_WARN As Long = &H9CEBFF     ' 警告 FFEB9C

Private Const EP_PROF As String = "/sse/config/v1/decryption-profiles"

' ---- メンバー選択画面 用 ----
Private Const PICK_KEY As String = "^+m"            ' Ctrl+Shift+M
Private Const CELL_MAX As Long = 32767
Private Const SH_OVERFLOW As String = "ピッカー_選択結果"

Private mEvents As clsPickerEvents
Private mNames As Object     ' 候補 name -> 種別ラベル（追加順を保持）
Private mExclude As String

' 項目定義（InitF で設定）
Private mFK As Variant       ' APIのキー
Private mFT As Variant       ' a=配列 b=真偽 s=文字列 t=種類 c=証明書(種類の付属)
Private mFL As Variant       ' 表示名

'----------------------------------------------------------
' ボタン配置（初回のみ実行）
'----------------------------------------------------------
Public Sub SetupButtons()
    AddBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    AddBtn SH_GET, "D3:F4", "まとめて取得（ポリシー・プロファイル・選択肢）", "FetchAllInfo"
    AddBtn SH_GET, "H3:I4", "ポリシーのみ取得", "FetchRules"
    AddBtn SH_SET, "F3:H4", "取得結果を展開（作り直し）", "CopyGetToSet"
    AddBtn SH_SET, "J3:K4", "選択行の下に追加行を挿入", "InsertRuleRow"
    AddBtn SH_SET, "M3:N4", "入力チェック", "CheckInput"
    AddBtn SH_SET, "P3:Q4", "JSON化 → ④へ", "GenerateJson"
    AddBtn SH_JSON, "G4:H4", "API実行（全行）", "ExecuteAll"
    AddBtn SH_JSON, "G6:H6", "クリア", "ClearJson"
    RemoveBtn SH_JSON, "PushConfig"     ' Push機能は削除（PushはSCM画面から実施）。旧ボタンを消す
    AddBtn SH_DIFF, "L3:M4", "実行前プレビュー（取得⇔設定）", "BuildDiffPreview"
    AddBtn SH_DIFF, "O3:P4", "実行後チェック（SCM再取得）", "BuildDiffCheck"
    AddBtn SH_LOG, "F3", "ログクリア", "ClearLog"
    RemoveBtn SH_LOG, "CheckJobs"      ' ジョブ状況確認も削除。旧ボタンを消す
    AddBtn SH_PROF, "D3:E4", "復号プロファイル取得（API）", "FetchProfilesBtn"
    AddBtn SH_CHO, "D3:E4", "選択肢を取得（API）", "FetchChoicesBtn"
    PickerStart True
    MsgBox "ボタンを配置しました。", vbInformation
End Sub

'==========================================================
' 項目定義
'==========================================================
Private Sub InitF()
    If IsArray(mFK) Then Exit Sub
    mFK = Array("disabled", "tag", "from", "source", "negate_source", "source_user", "source_hip", _
                "to", "destination", "negate_destination", "destination_hip", "service", "category", _
                "action", "type", "@cert", "profile", "log_success", "log_fail", "log_setting", "description")
    mFT = Array("b", "a", "a", "a", "b", "a", "a", _
                "a", "a", "b", "a", "a", "a", _
                "s", "t", "c", "s", "b", "b", "s", "s")
    mFL = Array("無効", "タグ", "送信元ゾーン", "送信元アドレス", "送信元否定", "送信元ユーザー", "送信元HIP", _
                "宛先ゾーン", "宛先アドレス", "宛先否定", "宛先HIP", "サービス", "URLカテゴリ", _
                "アクション", "種類", "証明書", "復号プロファイル", "成功ログ", "失敗ログ", "ログ転送", "説明")
End Sub

' 項目番号（0=name, 1～21=各項目）の表示名
Private Function FieldLabel(ByVal i As Long) As String
    InitF
    If i = 0 Then FieldLabel = "name" Else FieldLabel = mFL(i - 1)
End Function

' APIキー → 項目番号（1～21）
Private Function FIdx(ByVal k As String) As Long
    Dim i As Long
    InitF
    For i = 0 To NF - 1
        If mFK(i) = k Then FIdx = i + 1: Exit Function
    Next
End Function

' 比較用に値をそろえる（配列は区切りを統一、真偽は true 以外を空欄に）
Private Function Cmp(ByVal i As Long, ByVal s As String) As String
    InitF
    If i = 0 Then Cmp = Trim$(s): Exit Function
    Select Case mFT(i - 1)
        Case "a": Cmp = Join(SplitMembers(s), vbLf)
        Case "b": If LCase$(Trim$(s)) = "true" Then Cmp = "true"
        Case Else
            If mFK(i - 1) = "description" Then Cmp = s Else Cmp = Trim$(s)
    End Select
End Function

' APIの1ルール → 値の配列（0=name, 1～21=項目。配列はセル内改行区切り）
Private Function RuleVals(ByVal item As Object) As Variant
    Dim v() As String, i As Long, k As String, td As Object, tk As Variant
    InitF
    ReDim v(0 To NF)
    v(0) = Nz(item, "name")
    For i = 0 To NF - 1
        k = mFK(i)
        Select Case mFT(i)
            Case "a": v(i + 1) = JoinLf(item, k)
            Case "b": v(i + 1) = BoolTxt(item, k)
            Case "s": v(i + 1) = Nz(item, k)
            Case "t"
                If item.Exists("type") Then
                    If IsObject(item("type")) Then
                        Set td = item("type")
                        For Each tk In td.Keys
                            v(i + 1) = CStr(tk)
                            If CStr(tk) = "ssl_inbound_inspection" Then
                                If Not IsObject(td(tk)) Then v(i + 2) = SafeStr(td(tk))
                            End If
                        Next
                    End If
                End If
        End Select
    Next
    RuleVals = v
End Function

Private Function JoinLf(ByVal item As Object, ByVal key As String) As String
    Dim t As Variant, r As String
    If Not item.Exists(key) Then Exit Function
    If Not IsObject(item(key)) Then JoinLf = SafeStr(item(key)): Exit Function
    For Each t In item(key)
        If Not IsObject(t) Then
            If r <> "" Then r = r & vbLf
            r = r & SafeStr(t)
        End If
    Next
    JoinLf = r
End Function

Private Function BoolTxt(ByVal item As Object, ByVal key As String) As String
    Dim x As Variant
    If Not item.Exists(key) Then Exit Function
    If IsObject(item(key)) Then Exit Function
    x = item(key)
    If VarType(x) = vbBoolean Then
        BoolTxt = IIf(x, "true", "false")
    Else
        BoolTxt = LCase$(SafeStr(x))
    End If
End Function

'==========================================================
' 基本設定の読出し
'==========================================================
Private Function Enc(ByVal s As String) As String
    Enc = Application.WorksheetFunction.EncodeURL(s)
End Function

Private Function ApiBase() As String
    Dim b As String
    b = Cfg("CFG_BASE_URL")
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    ApiBase = b
End Function

Private Function PageLimit() As Long
    PageLimit = Val(CfgOpt("CFG_PAGE_LIMIT"))
    If PageLimit <= 0 Then PageLimit = 200
End Function

' Pre / Post（基本設定 CFG_POSITION）
Private Function Positions() As Variant
    Select Case LCase$(CfgOpt("CFG_POSITION"))
        Case "post": Positions = Array("post")
        Case "両方": Positions = Array("pre", "post")
        Case Else: Positions = Array("pre")
    End Select
End Function

' 取得対象スコープ（○のもの、上から順）: Collection of Array(表示名, folder)
Private Function SelScopes() As Collection
    Dim c As New Collection, rng As Range, i As Long
    Set rng = ThisWorkbook.Names("CFG_SCOPES").RefersToRange
    For i = 1 To rng.Rows.Count
        If Trim$(SafeStr(rng.Cells(i, 2).Value)) = "○" And Trim$(SafeStr(rng.Cells(i, 3).Value)) <> "" Then
            c.Add Array(Trim$(SafeStr(rng.Cells(i, 1).Value)), Trim$(SafeStr(rng.Cells(i, 3).Value)))
        End If
    Next
    Set SelScopes = c
End Function

' folder → 表示名（例：Prisma Access（Shared））
Private Function ScopeLabel(ByVal folder As String) As String
    Dim rng As Range, i As Long, l As String
    If LCase$(folder) = "all" Then ScopeLabel = "グローバル（All）※参照のみ": Exit Function
    Set rng = ThisWorkbook.Names("CFG_SCOPES").RefersToRange
    For i = 1 To rng.Rows.Count
        If Trim$(SafeStr(rng.Cells(i, 3).Value)) = folder Then
            l = Trim$(SafeStr(rng.Cells(i, 1).Value))
            If l = folder Or l = "" Then ScopeLabel = folder Else ScopeLabel = l & "（" & folder & "）"
            Exit Function
        End If
    Next
    ScopeLabel = folder
End Function

'==========================================================
' ② 取得（folder × position ごとに GET → スコープごとに表化）
'==========================================================
Public Sub FetchAllInfo()
    Dim n As Long, pn As Long, cf As Long, msg As String
    If Not EnsureToken() Then Exit Sub
    n = FetchRulesCore()
    If n < 0 Then Exit Sub
    pn = FetchProfiles()
    cf = FetchChoices()
    Application.StatusBar = False
    msg = "復号ポリシー " & n & " 件を取得しました。"
    If pn >= 0 Then msg = msg & vbLf & "復号プロファイル " & pn & " 件を取得しました。" Else msg = msg & vbLf & "復号プロファイルの取得に失敗しました（実行ログ参照）。"
    If cf > 0 Then msg = msg & vbLf & "選択肢のうち " & cf & " 種類は取得に失敗しました（実行ログ参照、シートの値はそのまま）。"
    msg = msg & vbLf & vbLf & "③設定シートを作り直す場合は、③の「取得結果を展開」を押してください。"
    Worksheets(SH_GET).Activate
    MsgBox msg, vbInformation
End Sub

Public Sub FetchRules()
    Dim n As Long
    If Not EnsureToken() Then Exit Sub
    n = FetchRulesCore()
    Application.StatusBar = False
    If n < 0 Then Exit Sub
    Worksheets(SH_GET).Activate
    MsgBox "復号ポリシー " & n & " 件を取得しました。" & vbLf & _
           "③設定シートを作り直す場合は、③の「取得結果を展開」を押してください。", vbInformation
End Sub

' 戻り値：取得したルール件数（中止は -1）
Private Function FetchRulesCore() As Long
    Dim groups As Collection, failN As Long, g As Variant, n As Long
    FetchRulesCore = -1
    Set groups = FetchRuleGroups("ポリシー取得", failN)
    If groups Is Nothing Then Exit Function
    If groups.Count = 0 Then
        MsgBox "どのスコープも取得できませんでした。実行ログを確認してください。", vbExclamation
        Exit Function
    End If
    Application.ScreenUpdating = False
    WriteGetSheet groups
    Application.ScreenUpdating = True
    For Each g In groups
        n = n + g(3).Count
    Next
    If failN > 0 Then MsgBox failN & " 件のスコープで取得に失敗しました（実行ログ参照）。そのスコープはシートに出ません。", vbExclamation
    FetchRulesCore = n
End Function

' 戻り値：Collection of Array(folder, position, 表示名, Collection of APIの要素)
'   並び順は Pre をすべて（親→子）→ Post をすべて（子→親）。FlowOrder 参照
'   子folderの応答に含まれる上位folderのルール（継承分）は、定義元のスコープにだけ出す
'   folder=All（グローバル）のルールは「グローバル（All）」スコープとして一番外側に出す（参照のみ）
Private Function FetchRuleGroups(ByVal opLabel As String, ByRef failN As Long) As Collection
    Dim scopes As Collection, sc As Variant, ps As Variant, items As Collection, item As Variant
    Dim own As Object, seen As Object, allD As Object, c As Collection, groups As New Collection, fo As Variant
    Dim st As Long, msg As String, f As String, k As String, inh As Long, q As String, key As String
    Set scopes = SelScopes()
    If scopes.Count = 0 Then
        MsgBox "基本設定の「取得・登録の対象スコープ」で、対象に○を付けてください。", vbExclamation
        Exit Function
    End If
    Set own = CreateObject("Scripting.Dictionary")
    Set allD = CreateObject("Scripting.Dictionary")
    For Each ps In Positions()
        Set seen = CreateObject("Scripting.Dictionary")
        Set allD(CStr(ps)) = New Collection
        For Each sc In scopes
            Prog opLabel & "：" & sc(0) & "（" & ps & "）を取得中…"
            q = "folder=" & Enc(CStr(sc(1))) & "&position=" & ps
            Set items = GetAll(BaseUrl(), q, opLabel, st, msg)
            If items Is Nothing Then
                failN = failN + 1
            Else
                Set c = New Collection
                inh = 0
                For Each item In items
                    f = Nz(item, "folder")
                    If f = "" Or f = CStr(sc(1)) Then
                        c.Add item
                    ElseIf LCase$(f) = "all" Then
                        k = Nz(item, "id")
                        If k = "" Then k = Nz(item, "name")
                        If Not seen.Exists(k) Then
                            seen(k) = True
                            allD(CStr(ps)).Add item
                        End If
                    Else
                        inh = inh + 1
                    End If
                Next
                Set own(CStr(sc(1)) & "|" & ps) = c
                WriteLog opLabel, CStr(sc(0)) & "／" & ps, "GET", WithQuery(BaseUrl(), q), 200, _
                         "定義 " & c.Count & " 件（上位folderから継承 " & inh & " 件は定義元のスコープに表示）"
            End If
        Next
    Next
    ' 評価順に並べる
    For Each fo In FlowOrder(scopes)
        key = fo(0) & "|" & fo(1)
        If fo(0) = "All" Then
            If allD.Exists(CStr(fo(1))) Then
                If allD(CStr(fo(1))).Count > 0 Then groups.Add Array("All", CStr(fo(1)), GroupLabel("All", CStr(fo(1))), allD(CStr(fo(1))))
            End If
        ElseIf own.Exists(key) Then
            groups.Add Array(CStr(fo(0)), CStr(fo(1)), GroupLabel(CStr(fo(0)), CStr(fo(1))), own(key))
        End If
    Next
    Set FetchRuleGroups = groups
End Function

' 並び順：Pre をすべて（親→子）、続けて Post をすべて（子→親＝評価順）: Collection of Array(folder, position)
'   例）All Pre → Shared Pre → RN Pre → SC Pre → MUC Pre → GP Pre → EP Pre
'       → EP Post → GP Post → MUC Post → SC Post → RN Post → Shared Post → All Post（基本設定で選んだ position だけ）
Private Function FlowOrder(ByVal scopes As Collection) As Collection
    Dim res As New Collection, fs() As String, n As Long, i As Long, sc As Variant, usePre As Boolean, usePost As Boolean, ps As Variant
    For Each ps In Positions()
        If ps = "pre" Then usePre = True Else usePost = True
    Next
    ReDim fs(0 To scopes.Count)
    fs(0) = "All"
    n = 0
    For Each sc In scopes
        n = n + 1
        fs(n) = CStr(sc(1))
    Next
    If usePre Then
        For i = 0 To n
            res.Add Array(fs(i), "pre")
        Next
    End If
    If usePost Then
        For i = n To 0 Step -1
            res.Add Array(fs(i), "post")
        Next
    End If
    Set FlowOrder = res
End Function

' スコープの階層（段のずらし幅）: All=0 / Prisma Access=1 / RN・SC・MUC=2 / GlobalProtect・Explicit Proxy=3
Private Function ScopeDepth(ByVal folder As String) As Long
    Select Case LCase$(folder)
        Case "all": ScopeDepth = 0
        Case "shared": ScopeDepth = 1
        Case "remote networks", "service connections", "mobile users container": ScopeDepth = 2
        Case "mobile users", "mobile users explicit proxy": ScopeDepth = 3
        Case Else: ScopeDepth = 1
    End Select
End Function

' 表示名＋【Pre】/【Post】
Private Function GroupLabel(ByVal folder As String, ByVal pos As String) As String
    GroupLabel = ScopeLabel(folder) & " " & IIf(LCase$(pos) = "post", "【Post】", "【Pre】")
End Function

' 見出し行の色（階層が深いほど明るい青）
Private Function DepthColor(ByVal d As Long) As Long
    Select Case d
        Case 0: DepthColor = &H404040        ' 404040
        Case 1: DepthColor = &H64381F        ' 1F3864
        Case 2: DepthColor = &H975A2F        ' 2F5A97 系
        Case Else: DepthColor = &HC47244     ' 4472C4
    End Select
End Function

' 階層の列（0=第1階層 Prisma Access・グローバル / 1=第2階層 RN・SC・MUC / 2=第3階層 GP・EP）
Private Function ScopeLvl(ByVal folder As String) As Long
    ScopeLvl = ScopeDepth(folder) - 1
    If ScopeLvl < 0 Then ScopeLvl = 0
End Function

' スコープ表示を階層の列へ置く（scopeCol～scopeCol+2 のうち入っている文字を、その階層の列だけに書く）
Private Sub IndentRow(ByVal ws As Worksheet, ByVal r As Long, ByVal scopeCol As Long, ByVal nameCol As Long, ByVal folder As String)
    Dim t As String, i As Long, b As Boolean
    For i = 0 To 2
        If t = "" Then
            t = SafeStr(ws.Cells(r, scopeCol + i).Value)
            If t <> "" Then b = ws.Cells(r, scopeCol + i).Font.Bold
        End If
    Next
    ws.Range(ws.Cells(r, scopeCol), ws.Cells(r, scopeCol + 2)).ClearContents
    ws.Range(ws.Cells(r, scopeCol), ws.Cells(r, scopeCol + 2)).Font.Bold = False
    With ws.Cells(r, scopeCol + ScopeLvl(folder))
        .NumberFormat = "@"
        .Value = t
        .Font.Bold = b
    End With
End Sub

' GET（ページ送りで全件）。失敗時は Nothing（実行ログに記録）
Private Function GetAll(ByVal url0 As String, ByVal q As String, ByVal opLabel As String, _
                        ByRef st As Long, ByRef resMsg As String) As Collection
    Dim url As String, resText As String, res As Object, item As Variant
    Dim offset As Long, cnt As Long, total As Long, lim As Long, items As New Collection
    lim = PageLimit()
    Do
        url = WithQuery(url0, q & "&limit=" & lim & "&offset=" & offset)
        ' 長い取得の途中でトークンが切れても続けられるよう、期限前に取り直し、401 なら1回だけ再取得して再実行
        If Not EnsureToken() Then st = 401: resText = "トークンを取得できません": Exit Function
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
    Set GetAll = items
End Function

Private Sub WriteGetSheet(ByVal groups As Collection)
    Dim ws As Worksheet, g As Variant, item As Variant, r As Long, n As Long, v As Variant, a() As Variant, i As Long
    Set ws = Worksheets(SH_GET)
    ClearData ws, G_HF + 1
    r = FIRST
    For Each g In groups
        WriteHeaderRow ws, r, G_SCOPE, G_LAST, G_HF, CStr(g(0)), CStr(g(1)), CStr(g(2)), g(3).Count
        r = r + 1
        n = 0
        For Each item In g(3)
            n = n + 1
            v = RuleVals(item)
            ReDim a(1 To 1, 1 To G_LAST)
            a(1, G_SEQ) = n
            a(1, G_SCOPE) = g(2)
            a(1, G_FOLDER) = g(0)
            a(1, G_POS) = g(1)
            a(1, G_ID) = Nz(item, "id")
            For i = 0 To NF
                a(1, G_NAME + i) = v(i)
            Next
            a(1, G_JSON) = Left$(JsonOf(item), 32000)
            ws.Range(ws.Cells(r, 2), ws.Cells(r, G_LAST)).NumberFormat = "@"
            ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST)).Value = a
            PaintBase ws, r, G_LAST, False
            If n = 1 Then ws.Cells(r, G_SCOPE).Font.Bold = True
            IndentRow ws, r, G_SCOPE, G_NAME, CStr(g(0))
            r = r + 1
        Next
    Next
    ws.Range("B3").Value = Now
End Sub

'==========================================================
' 表の共通処理
'==========================================================
Private Sub ClearData(ByVal ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(FIRST, 1), ws.Cells(ws.Rows.Count, lastCol))
        .ClearContents
        .ClearFormats              ' 入力規則（プルダウン）は残る
        .FormatConditions.Delete
    End With
End Sub

' スコープ見出し行（A列=▼、labelCol に「■ 表示名 ／ pre ／ n 件」、hfCol・hfCol+1 に folder・position）
Private Sub WriteHeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal labelCol As Long, ByVal lastCol As Long, _
                           ByVal hfCol As Long, ByVal folder As String, ByVal pos As String, ByVal label As String, ByVal n As Long)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, hfCol + 1)).NumberFormat = "@"
    ws.Cells(r, 1).Value = HDR
    ws.Cells(r, labelCol).Value = "■ " & label & " ／ " & n & " 件"
    ws.Cells(r, hfCol).Value = folder
    ws.Cells(r, hfCol + 1).Value = pos
    PaintHeader ws, r, lastCol, hfCol, labelCol
End Sub

Private Sub PaintHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, ByVal hfCol As Long, ByVal labelCol As Long)
    Dim d As Long
    d = ScopeDepth(SafeStr(ws.Cells(r, hfCol).Value))
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        .Interior.Color = DepthColor(d)
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = vbWhite
        .Font.Strikethrough = False
        .WrapText = False
        .Borders.LineStyle = xlNone
    End With
    ws.Range(ws.Cells(r, hfCol), ws.Cells(r, hfCol + 1)).Font.Color = vbWhite
    ' 見出しの文字を階層の列へ（右の空セルにはみ出して表示）
    IndentRow ws, r, labelCol, 0, SafeStr(ws.Cells(r, hfCol).Value)
    ws.Cells(r, labelCol + ScopeLvl(SafeStr(ws.Cells(r, hfCol).Value))).Font.Bold = True
    ws.Range(ws.Cells(r, labelCol), ws.Cells(r, labelCol + 2)).WrapText = False
End Sub

' データ行の基本の書式（設定シートは入力セル=黄、自動セル=グレー）
Private Sub PaintBase(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, ByVal isSet As Boolean)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        .Interior.Pattern = xlNone
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = vbBlack
        .Font.Strikethrough = False
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
        .Borders.Color = &HBFBFBF
    End With
    If isSet Then
        ws.Cells(r, S_OP).Interior.Color = C_IN
        ws.Range(ws.Cells(r, S_NAME), ws.Cells(r, S_LAST)).Interior.Color = C_IN
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 1)).Interior.Color = C_FX
        ws.Range(ws.Cells(r, S_JUDGE), ws.Cells(r, S_ID)).Interior.Color = C_FX
    End If
End Sub

Private Function DataLastRow(ByVal ws As Worksheet) As Long
    With ws.UsedRange
        DataLastRow = .Row + .Rows.Count - 1
    End With
End Function

Private Function HeaderAbove(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim i As Long
    For i = r To FIRST Step -1
        If SafeStr(ws.Cells(i, 1).Value) = HDR Then HeaderAbove = i: Exit Function
    Next
End Function

' 取得／設定シートを読み、スコープごとにまとめる
'   戻り値 Collection of Array(folder, position, 表示名, Collection of rec)
'   rec = Array(id, 値の配列(0=name,1～21), 行番号, 操作)
Private Function ReadGroups(ByVal ws As Worksheet, ByVal isSet As Boolean, ByVal skipDel As Boolean) As Collection
    Dim groups As New Collection, cur As Collection, arr As Variant, lastR As Long, i As Long, k As Long
    Dim cId As Long, cName As Long, cOp As Long, cHF As Long, id As String, nm As String, op As String, v() As String
    Dim folder As String
    If isSet Then
        cId = S_ID: cName = S_NAME: cOp = S_OP: cHF = S_HF
    Else
        cId = G_ID: cName = G_NAME: cOp = 0: cHF = G_HF
    End If
    Set ReadGroups = groups
    lastR = DataLastRow(ws)
    If lastR < FIRST Then Exit Function
    arr = ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, cHF + 1)).Value
    For i = 1 To UBound(arr, 1)
        If SafeStr(arr(i, 1)) = HDR Then
            Set cur = New Collection
            folder = SafeStr(arr(i, cHF))
            groups.Add Array(folder, SafeStr(arr(i, cHF + 1)), GroupLabel(folder, SafeStr(arr(i, cHF + 1))), cur)
        ElseIf Not cur Is Nothing Then
            id = Trim$(SafeStr(arr(i, cId)))
            nm = SafeStr(arr(i, cName))
            If cOp > 0 Then op = Trim$(SafeStr(arr(i, cOp))) Else op = ""
            If (id <> "" Or Trim$(nm) <> "" Or op <> "") And Not (skipDel And op = "削除") Then
                ReDim v(0 To NF)
                For k = 0 To NF
                    v(k) = SafeStr(arr(i, cName + k))
                Next
                cur.Add Array(id, v, FIRST + i - 1, op)
            End If
        End If
    Next
End Function

' 取得シート：id → Array(値, folder|position, スコープ内の順番)
Private Function BeforeMap() As Object
    Dim d As Object, g As Variant, rec As Variant, n As Long
    Set d = CreateObject("Scripting.Dictionary")
    For Each g In ReadGroups(Worksheets(SH_GET), False, False)
        n = 0
        For Each rec In g(3)
            n = n + 1
            If rec(0) <> "" Then d(rec(0)) = Array(rec(1), g(0) & "|" & g(1), n)
        Next
    Next
    Set BeforeMap = d
End Function

'==========================================================
' ③ 設定シート
'==========================================================
' ②取得の内容を③設定に展開（作り直し）
Public Sub CopyGetToSet()
    Dim wsG As Worksheet, wsS As Worksheet, lastR As Long, arr As Variant, out() As Variant, i As Long, c As Long, n As Long, opN As Long
    Set wsG = Worksheets(SH_GET)
    Set wsS = Worksheets(SH_SET)
    opN = CountOps()
    If opN > 0 Then
        If MsgBox("③設定シートに操作の入った行が " & opN & " 件あります。" & vbLf & _
                  "②取得の内容で作り直すと、入力した内容はすべて消えます。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    Else
        If MsgBox("②取得の内容で③設定シートを作り直します。よろしいですか？", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    End If
    lastR = DataLastRow(wsG)
    If lastR < FIRST Then
        MsgBox "②取得シートにデータがありません。先に取得してください。", vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    arr = wsG.Range(wsG.Cells(FIRST, 1), wsG.Cells(lastR, G_HF + 1)).Value
    n = UBound(arr, 1)
    ReDim out(1 To n, 1 To S_HF + 1)
    For i = 1 To n
        If SafeStr(arr(i, 1)) = HDR Then
            out(i, 1) = HDR
            For c = 0 To 2
                out(i, S_SCOPE + c) = SafeStr(arr(i, G_SCOPE + c))
            Next
            out(i, S_HF) = SafeStr(arr(i, G_HF))
            out(i, S_HF + 1) = SafeStr(arr(i, G_HF + 1))
        Else
            out(i, 1) = SafeStr(arr(i, 1))
            For c = G_SCOPE To G_LAST - 1
                out(i, c + 3) = SafeStr(arr(i, c))
            Next
        End If
    Next
    ClearData wsS, S_HF + 1
    wsS.Range(wsS.Cells(FIRST, 2), wsS.Cells(FIRST + n - 1, S_HF + 1)).NumberFormat = "@"
    wsS.Range(wsS.Cells(FIRST, 1), wsS.Cells(FIRST + n - 1, S_HF + 1)).Value = out
    For i = 1 To n
        If out(i, 1) = HDR Then
            PaintHeader wsS, FIRST + i - 1, S_LAST, S_HF, S_SCOPE
        ElseIf SafeStr(out(i, S_NAME)) <> "" Or SafeStr(out(i, S_ID)) <> "" Then
            PaintBase wsS, FIRST + i - 1, S_LAST, True
            IndentRow wsS, FIRST + i - 1, S_SCOPE, S_NAME, SafeStr(out(i, S_FOLDER))
        End If
    Next
    Application.ScreenUpdating = True
    wsS.Activate
    MsgBox "③設定シートを作り直しました。" & vbLf & _
           "変更・削除は「操作」列で選択、追加は追加したい位置の1つ上の行を選んで「選択行の下に追加行を挿入」を押してください。", vbInformation
End Sub

Private Function CountOps() As Long
    Dim ws As Worksheet, lastR As Long
    Set ws = Worksheets(SH_SET)
    lastR = DataLastRow(ws)
    If lastR < FIRST Then Exit Function
    CountOps = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(FIRST, S_OP), ws.Cells(lastR, S_OP)))
End Function

' 選択行の下に「追加」の行を挿入（見出し行を選べばスコープの先頭に入る）
Public Sub InsertRuleRow()
    Dim ws As Worksheet, r As Long, h As Long, folder As String, pos As String, k As Variant
    Set ws = Worksheets(SH_SET)
    If ActiveSheet.Name <> SH_SET Then
        ws.Activate
        MsgBox "③設定シートで、追加したい位置の1つ上の行（先頭に入れるときは ▼スコープの行）を選択してから押してください。", vbInformation
        Exit Sub
    End If
    r = ActiveCell.Row
    h = HeaderAbove(ws, r)
    If r < FIRST Or h = 0 Then
        MsgBox "▼スコープの行、またはその下のルールの行を選択してください。", vbExclamation
        Exit Sub
    End If
    folder = SafeStr(ws.Cells(h, S_HF).Value)
    pos = SafeStr(ws.Cells(h, S_HF + 1).Value)
    If LCase$(folder) = "all" Then
        MsgBox "グローバル（All）のスコープには追加できません（参照のみ）。", vbExclamation
        Exit Sub
    End If
    InitF
    Application.ScreenUpdating = False
    ws.Rows(r + 1).Insert Shift:=xlDown
    r = r + 1
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, S_HF + 1))
        .ClearContents
        .ClearFormats
        .NumberFormat = "@"
    End With
    ws.Cells(r, S_OP).Value = "追加"
    ws.Cells(r, S_SCOPE).Value = GroupLabel(folder, pos)
    ws.Cells(r, S_FOLDER).Value = folder
    ws.Cells(r, S_POS).Value = pos
    For Each k In Array("from", "source", "source_user", "to", "destination", "service", "category")
        ws.Cells(r, S_NAME + FIdx(CStr(k))).Value = "any"
    Next
    ws.Cells(r, S_NAME + FIdx("action")).Value = "decrypt"
    ws.Cells(r, S_NAME + FIdx("type")).Value = "ssl_forward_proxy"
    PaintBase ws, r, S_LAST, True
    IndentRow ws, r, S_SCOPE, S_NAME, folder
    ws.Range(ws.Cells(r, S_SCOPE), ws.Cells(r, S_LAST)).Interior.Color = DIFF_ADD
    ws.Cells(r, S_OP).Interior.Color = DIFF_ADD
    Application.ScreenUpdating = True
    ws.Cells(r, S_NAME).Select
    MsgBox "追加行を挿入しました。name と各項目を入力してください（ゾーン・アドレス等はダブルクリックで選択）。" & vbLf & _
           "API実行時は、登録（POST）のあと、この位置（直前のルールの後ろ）へ移動する処理も自動で行います。", vbInformation
End Sub

Public Sub CheckInput()
    Dim ng As Long, okN As Long, warnN As Long
    ng = CheckSet(okN, warnN)
    Worksheets(SH_SET).Activate
    MsgBox "入力チェック結果：反映対象 " & okN & " 件（うち警告 " & warnN & " 件）／ NG " & ng & " 件" & vbLf & _
           "判定・理由は C・D 列、取得時から変わったセルは橙色で表示しています。", IIf(ng > 0, vbExclamation, vbInformation)
End Sub

' 設定シートの全行をチェックし、判定・理由・色を付ける。戻り値=NG件数
'   okN = JSON化の対象（操作あり・OK/警告）件数、warnN = そのうち警告
Private Function CheckSet(ByRef okN As Long, ByRef warnN As Long) As Long
    Dim ws As Worksheet, before As Object, groups As Collection, g As Variant, rec As Variant, nameCnt As Object
    Dim folder As String, pos As String, prevIdx As Long, r As Long, id As String, op As String, v As Variant
    Dim b As Variant, bv As Variant, ngs As String, warns As String, i As Long, diffN As Long, judge As String, k As String
    Dim ng As Long
    InitF
    Set ws = Worksheets(SH_SET)
    Set before = BeforeMap()
    Set groups = ReadGroups(ws, True, False)
    Set nameCnt = CreateObject("Scripting.Dictionary")
    For Each g In groups
        For Each rec In g(3)
            If rec(3) <> "削除" And Trim$(rec(1)(0)) <> "" Then
                k = g(0) & "|" & g(1) & "|" & Trim$(rec(1)(0))     ' Pre と Post は別のルールベース（SCM既定ルールは両方に同名あり）
                nameCnt(k) = nameCnt(k) + 1
            End If
        Next
    Next
    Application.ScreenUpdating = False
    For Each g In groups
        folder = g(0): pos = g(1): prevIdx = 0
        For Each rec In g(3)
            id = rec(0): v = rec(1): r = rec(2): op = rec(3)
            ngs = "": warns = "": bv = Empty
            ws.Range(ws.Cells(r, S_SCOPE), ws.Cells(r, S_POS)).NumberFormat = "@"
            ws.Cells(r, S_SCOPE).Value = g(2)
            ws.Cells(r, S_FOLDER).Value = folder
            ws.Cells(r, S_POS).Value = pos
            PaintBase ws, r, S_LAST, True
            IndentRow ws, r, S_SCOPE, S_NAME, folder
            ' 操作と id
            Select Case op
                Case ""
                    If id = "" Then ngs = ngs & vbLf & "NG:新しい行は操作に「追加」を選択してください"
                Case "追加"
                    If id <> "" Then ngs = ngs & vbLf & "NG:既存ルール（idあり）の行です。追加は「選択行の下に追加行を挿入」で作った行で行ってください"
                Case "変更", "削除", "移動", "変更+移動"
                    If id = "" Then ngs = ngs & vbLf & "NG:idがありません（" & op & "は取得済みの既存ルールだけが対象です）"
                Case Else
                    ngs = ngs & vbLf & "NG:操作はプルダウンから選択してください"
            End Select
            If LCase$(folder) = "all" And op <> "" Then ngs = ngs & vbLf & "NG:グローバル（All）のルールはこのツールでは操作できません"
            ' 取得時との比較・順番
            If id <> "" Then
                If before.Exists(id) Then
                    b = before(id)
                    bv = b(0)
                    If b(1) <> folder & "|" & pos Then
                        ngs = ngs & vbLf & "NG:別のスコープ（" & Replace(b(1), "|", "／") & "）のルールです。スコープをまたぐ移動はできません（削除＋追加にしてください）"
                    ElseIf op <> "移動" And op <> "変更+移動" And op <> "削除" Then
                        If b(2) < prevIdx Then ngs = ngs & vbLf & "NG:取得時と順番が変わっています。並べ替えたルールの操作は「移動」（内容も変えるなら「変更+移動」）にしてください"
                        If b(2) > prevIdx Then prevIdx = b(2)
                    End If
                Else
                    ngs = ngs & vbLf & "NG:②取得シートに無いidです（再取得して設定シートを作り直してください）"
                End If
            End If
            ' 内容（追加・変更）
            If op = "追加" Or InStr(op, "変更") > 0 Then ngs = ngs & ContentNg(v, warns)
            ' 追加・変更する行だけ判定（取得したままの既存ルールは対象外）
            If (op = "追加" Or InStr(op, "変更") > 0) And Trim$(v(0)) <> "" Then
                If nameCnt(folder & "|" & pos & "|" & Trim$(v(0))) > 1 Then ngs = ngs & vbLf & "NG:同じスコープ・同じ position（" & pos & "）内で name が重複しています"
            End If
            ' 行の色（追加=緑、削除=グレー＋取り消し線）
            If op = "追加" Then
                ws.Range(ws.Cells(r, S_SCOPE), ws.Cells(r, S_LAST)).Interior.Color = DIFF_ADD
            ElseIf op = "削除" Then
                With ws.Range(ws.Cells(r, S_SCOPE), ws.Cells(r, S_LAST))
                    .Interior.Color = DIFF_DEL
                    .Font.Strikethrough = True
                End With
            End If
            ' 取得時から変わったセルを橙色に
            diffN = 0
            If IsArray(bv) Then
                For i = 0 To NF
                    If Cmp(i, CStr(v(i))) <> Cmp(i, CStr(bv(i))) Then
                        diffN = diffN + 1
                        If op <> "削除" Then MarkChangedCell ws.Cells(r, S_NAME + i)
                    End If
                Next
                If diffN > 0 And InStr(op, "変更") = 0 And op <> "削除" And LCase$(folder) <> "all" Then
                    warns = warns & vbLf & "警告:取得時から内容が変わっていますが、操作が「変更」ではないため反映されません"
                End If
                If diffN = 0 And InStr(op, "変更") > 0 Then warns = warns & vbLf & "警告:取得時と差分がありません"
            End If
            ' 判定
            If ngs <> "" Then
                judge = "NG"
            ElseIf warns <> "" Then
                judge = "警告"
            ElseIf op <> "" Then
                judge = "OK"
            Else
                judge = ""
            End If
            ws.Cells(r, S_JUDGE).Value = judge
            ws.Cells(r, S_REASON).Value = Mid$(ngs & warns, 2)
            Select Case judge
                Case "NG": ws.Range(ws.Cells(r, S_JUDGE), ws.Cells(r, S_REASON)).Interior.Color = DIFF_NG: ng = ng + 1
                Case "警告": ws.Range(ws.Cells(r, S_JUDGE), ws.Cells(r, S_REASON)).Interior.Color = C_WARN
                Case "OK": ws.Cells(r, S_JUDGE).Interior.Color = DIFF_ADD
            End Select
            If op <> "" And (judge = "OK" Or judge = "警告") Then
                okN = okN + 1
                If judge = "警告" Then warnN = warnN + 1
            End If
        Next
    Next
    Application.ScreenUpdating = True
    CheckSet = ng
End Function

' 追加・変更する行の内容チェック。戻り値=NG（改行区切り）、warns に警告を追記
Private Function ContentNg(ByVal v As Variant, ByRef warns As String) As String
    Dim s As String, i As Long, k As String, a As Variant, j As Long, hasAny As Boolean, profs As Object
    InitF
    If Trim$(v(0)) = "" Then
        s = s & vbLf & "NG:name が空欄です"
    ElseIf Len(v(0)) > 63 Then
        s = s & vbLf & "NG:name は63文字以内にしてください"
    ElseIf Trim$(v(0)) <> v(0) Then
        s = s & vbLf & "NG:name の前後に空白があります"
    End If
    For i = 0 To NF - 1
        k = mFK(i)
        Select Case mFT(i)
            Case "a"
                a = SplitMembers(CStr(v(i + 1)))
                Select Case k
                    Case "from", "source", "source_user", "to", "destination", "service", "category"
                        If UBound(a) < 0 Then s = s & vbLf & "NG:" & mFL(i) & " は必須です（指定しないときは any）"
                End Select
                hasAny = False
                For j = 0 To UBound(a)
                    If LCase$(a(j)) = "any" Then hasAny = True
                Next
                If hasAny And UBound(a) > 0 Then s = s & vbLf & "NG:" & mFL(i) & " に any と他の値が混在しています"
            Case "b"
                Select Case LCase$(Trim$(CStr(v(i + 1))))
                    Case "", "true", "false"
                    Case Else: s = s & vbLf & "NG:" & mFL(i) & " は true / false / 空欄 のいずれかです"
                End Select
        End Select
    Next
    Select Case Trim$(CStr(v(FIdx("action"))))
        Case "decrypt", "no-decrypt"
        Case "": s = s & vbLf & "NG:アクション（decrypt / no-decrypt）を選択してください"
        Case Else: s = s & vbLf & "NG:アクションは decrypt / no-decrypt のどちらかです"
    End Select
    Select Case Trim$(CStr(v(FIdx("type"))))
        Case "ssl_forward_proxy"
            If Trim$(CStr(v(FIdx("@cert")))) <> "" Then warns = warns & vbLf & "警告:証明書は種類が ssl_inbound_inspection のときだけ使います（送信しません）"
        Case "ssl_inbound_inspection"
            If Trim$(CStr(v(FIdx("@cert")))) = "" Then s = s & vbLf & "NG:ssl_inbound_inspection のときは証明書を指定してください"
        Case ""
            If Trim$(CStr(v(FIdx("action")))) = "decrypt" Then s = s & vbLf & "NG:種類（ssl_forward_proxy など）を選択してください"
        Case Else
            s = s & vbLf & "NG:種類は ssl_forward_proxy / ssl_inbound_inspection のどちらかです"
    End Select
    If Trim$(CStr(v(FIdx("profile")))) <> "" Then
        Set profs = ProfileNames()
        If profs.Count > 0 And Not profs.Exists(Trim$(CStr(v(FIdx("profile"))))) Then
            warns = warns & vbLf & "警告:復号プロファイルが DecProfile_一覧 にありません（大文字・小文字も区別）"
        End If
    ElseIf Trim$(CStr(v(FIdx("action")))) = "decrypt" Then
        warns = warns & vbLf & "警告:復号プロファイルが空欄です（既定のプロファイルになります）"
    End If
    If Len(CStr(v(FIdx("description")))) > 1024 Then s = s & vbLf & "NG:説明は1024文字以内にしてください"
    ContentNg = s
End Function

Private Function ProfileNames() As Object
    Dim d As Object, ws As Worksheet, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets(SH_PROF)
    For r = FIRST To LastRow(ws, 2)
        If SafeStr(ws.Cells(r, 2).Value) <> "" Then d(SafeStr(ws.Cells(r, 2).Value)) = True
    Next
    Set ProfileNames = d
End Function

'==========================================================
' ④ JSON化（設定シート → JSON化シート）
'   実行順：①変更(PUT) → ②削除(DELETE) → ③スコープごとに上から 追加(POST＋位置へ移動) / 移動
'   追加・移動の位置 = 設定シートで直前にある（削除しない）ルールの後ろ。直前が無ければ先頭（top）
'==========================================================
Public Sub GenerateJson()
    Dim ng As Long, okN As Long, warnN As Long, ws As Worksheet, wsJ As Worksheet, groups As Collection
    Dim g As Variant, rec As Variant, ph As Variant, o As Long, sk As String, op As String, judge As String
    Dim tgt As Boolean, nm As String, prevRef As String, prevName As String, newRef As String
    ng = CheckSet(okN, warnN)
    If okN = 0 Then
        Worksheets(SH_SET).Activate
        MsgBox "JSON化の対象（操作あり・判定 OK／警告）がありません。" & IIf(ng > 0, vbLf & "NG が " & ng & " 件あります。D列の理由を確認してください。", ""), vbExclamation
        Exit Sub
    End If
    If ng > 0 Then
        If MsgBox("NG が " & ng & " 件あります。NG の行を除いて JSON化しますか？", vbYesNo + vbQuestion) = vbNo Then
            Worksheets(SH_SET).Activate
            Exit Sub
        End If
    End If
    Set ws = Worksheets(SH_SET)
    Set wsJ = Worksheets(SH_JSON)
    ClearJsonRows wsJ
    Set groups = ReadGroups(ws, True, False)
    o = JS_FIRST
    For Each ph In Array(1, 2, 3)
        For Each g In groups
            sk = g(0) & "|" & g(1)
            prevRef = "": prevName = ""
            For Each rec In g(3)
                op = rec(3)
                nm = Trim$(rec(1)(0))
                judge = SafeStr(ws.Cells(rec(2), S_JUDGE).Value)
                tgt = (op <> "" And (judge = "OK" Or judge = "警告"))
                Select Case ph
                    Case 1
                        If tgt And InStr(op, "変更") > 0 Then
                            WriteJ wsJ, o, "変更", nm, "PUT", BaseUrl() & "/" & rec(0), BuildBody(rec(1)), sk, rec(0), ""
                        End If
                    Case 2
                        If tgt And op = "削除" Then
                            WriteJ wsJ, o, "削除", nm, "DELETE", BaseUrl() & "/" & rec(0), "", sk, rec(0), ""
                        End If
                    Case 3
                        If op = "削除" And tgt Then
                            ' 削除するルールは位置の基準にしない
                        ElseIf op = "追加" Then
                            If tgt Then
                                newRef = "{{id:" & nm & "}}"
                                WriteJ wsJ, o, "追加", nm, "POST", _
                                       WithQuery(BaseUrl(), "folder=" & Enc(CStr(g(0))) & "&position=" & g(1)), BuildBody(rec(1)), sk, "", ""
                                WriteJ wsJ, o, "追加(位置)", nm, "POST", MoveUrl(newRef, CStr(g(0))), MoveBody(prevRef, CStr(g(1))), _
                                       sk, "", IIf(prevName = "", "（先頭）", prevName & " の後ろ")
                                prevRef = newRef: prevName = nm
                            End If
                        ElseIf rec(0) <> "" Then
                            If tgt And InStr(op, "移動") > 0 Then
                                WriteJ wsJ, o, "移動", nm, "POST", MoveUrl(CStr(rec(0)), CStr(g(0))), MoveBody(prevRef, CStr(g(1))), _
                                       sk, rec(0), IIf(prevName = "", "（先頭）", prevName & " の後ろ")
                            End If
                            prevRef = rec(0): prevName = nm
                        End If
                End Select
            Next
        Next
    Next
    If o > JS_FIRST Then
        With wsJ.Range(wsJ.Cells(JS_FIRST, 1), wsJ.Cells(o - 1, J_LAST))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
        End With
    End If
    wsJ.Activate
    MsgBox "③から " & (o - JS_FIRST) & " 行のAPI呼び出しを④に生成しました。" & vbLf & _
           "追加は「登録（POST）」と「位置へ移動（:move）」の2行です。{{id:name}} は実行時に、登録したルールのidに置き換えます。", vbInformation
End Sub

Private Sub WriteJ(ByVal ws As Worksheet, ByRef o As Long, ByVal op As String, ByVal nm As String, ByVal method As String, _
                   ByVal url As String, ByVal body As String, ByVal sk As String, ByVal id As String, ByVal anchor As String)
    ws.Range(ws.Cells(o, 2), ws.Cells(o, J_LAST)).NumberFormat = "@"
    ws.Cells(o, 1).Value = o - JS_FIRST + 1
    ws.Cells(o, 2).Value = op
    ws.Cells(o, 3).Value = nm
    ws.Cells(o, 4).Value = method
    ws.Cells(o, 5).Value = url
    ws.Cells(o, 6).Value = body
    ws.Cells(o, 7).Value = BuildCurl(method, url, body)
    ws.Cells(o, 9).Value = sk
    ws.Cells(o, 10).Value = id
    ws.Cells(o, 11).Value = anchor
    o = o + 1
End Sub

Private Function MoveUrl(ByVal idRef As String, ByVal folder As String) As String
    MoveUrl = BaseUrl() & "/" & idRef & ":move?folder=" & Enc(folder)
End Function

Private Function MoveBody(ByVal prevRef As String, ByVal pos As String) As String
    If prevRef = "" Then
        MoveBody = "{""destination"":""top"",""rulebase"":" & JsonStr(pos) & "}"
    Else
        MoveBody = "{""destination"":""after"",""destination_rule"":" & JsonStr(prevRef) & ",""rulebase"":" & JsonStr(pos) & "}"
    End If
End Function

' 設定シートの1行 → リクエストボディ（空欄の任意項目は送らない）
Private Function BuildBody(ByVal v As Variant) As String
    Dim js As String, i As Long, k As String, s As String
    InitF
    js = "{""name"":" & JsonStr(Trim$(CStr(v(0))))
    For i = 0 To NF - 1
        k = mFK(i)
        s = CStr(v(i + 1))
        Select Case mFT(i)
            Case "a"
                If UBound(SplitMembers(s)) >= 0 Then js = js & ",""" & k & """:" & JsonArr(s)
            Case "b"
                Select Case LCase$(Trim$(s))
                    Case "true", "false": js = js & ",""" & k & """:" & LCase$(Trim$(s))
                End Select
            Case "s"
                If k = "description" Then
                    If s <> "" Then js = js & ",""" & k & """:" & JsonStr(s)
                ElseIf Trim$(s) <> "" Then
                    js = js & ",""" & k & """:" & JsonStr(Trim$(s))
                End If
            Case "t"
                Select Case Trim$(s)
                    Case "ssl_forward_proxy": js = js & ",""type"":{""ssl_forward_proxy"":{}}"
                    Case "ssl_inbound_inspection": js = js & ",""type"":{""ssl_inbound_inspection"":" & JsonStr(Trim$(CStr(v(i + 2)))) & "}"
                End Select
        End Select
    Next
    BuildBody = js & "}"
End Function

Private Function JsonArr(ByVal s As String) As String
    Dim a As Variant, i As Long, r As String
    a = SplitMembers(s)
    For i = 0 To UBound(a)
        If r <> "" Then r = r & ","
        r = r & JsonStr(CStr(a(i)))
    Next
    JsonArr = "[" & r & "]"
End Function

' 値 → JSON文字列（取得時JSON列の表示用）
Private Function JsonOf(ByVal v As Variant) As String
    Dim k As Variant, s As String, x As Variant
    If IsObject(v) Then
        If TypeName(v) = "Collection" Then
            For Each x In v
                If s <> "" Then s = s & ","
                s = s & JsonOf(x)
            Next
            JsonOf = "[" & s & "]"
        Else
            For Each k In v.Keys
                If s <> "" Then s = s & ","
                s = s & JsonStr(CStr(k)) & ":" & JsonOf(v(k))
            Next
            JsonOf = "{" & s & "}"
        End If
    ElseIf IsNull(v) Then
        JsonOf = "null"
    ElseIf VarType(v) = vbBoolean Then
        JsonOf = IIf(v, "true", "false")
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
        JsonOf = Trim$(Str$(v))
    Else
        JsonOf = JsonStr(CStr(v))
    End If
End Function

'----------------------------------------------------------
' ④ API実行（JSON化シートの全行を上から順に）
'----------------------------------------------------------
Public Sub ExecuteAll()
    Dim ws As Worksheet, r As Long, lastR As Long, n As Long, st As Long, resText As String
    Dim ok As Long, ng As Long, skipN As Long, idMap As Object, url As String, body As String, method As String
    Dim sk As String, nm As String, op As String, newId As String, p() As String
    Set ws = Worksheets(SH_JSON)
    lastR = LastRow(ws, 4)
    n = lastR - JS_FIRST + 1
    If n <= 0 Then
        MsgBox "実行対象がありません。③でJSON化してください。", vbExclamation
        Exit Sub
    End If
    If MsgBox(n & " 行をSCMへ実行します（候補設定への反映。元に戻せません）。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Set idMap = CreateObject("Scripting.Dictionary")
    For r = JS_FIRST To lastR
        Prog "API実行中 " & (r - JS_FIRST + 1) & " / " & n
        op = SafeStr(ws.Cells(r, 2).Value)
        nm = SafeStr(ws.Cells(r, 3).Value)
        method = SafeStr(ws.Cells(r, 4).Value)
        url = SafeStr(ws.Cells(r, 5).Value)
        body = SafeStr(ws.Cells(r, 6).Value)
        sk = SafeStr(ws.Cells(r, 9).Value)
        If Not ResolveIds(url, sk, idMap) Or Not ResolveIds(body, sk, idMap) Then
            ws.Cells(r, 8).Value = "スキップ"
            WriteLog op, nm, method, url, 0, "スキップ：追加したルールのidが取得できていないため実行しません"
            skipN = skipN + 1
        Else
            st = HttpCall(method, url, body, resText)
            ws.Cells(r, 8).Value = st
            If st >= 200 And st < 300 Then
                ok = ok + 1
                If op = "追加" Then
                    newId = IdFromResponse(resText)
                    p = Split(sk, "|")
                    If newId = "" And UBound(p) >= 1 Then newId = FindIdByName(p(0), p(1), nm)
                    If newId <> "" Then
                        idMap(sk & "|" & nm) = newId
                        ws.Cells(r, 12).Value = newId
                    End If
                End If
            Else
                ng = ng + 1
            End If
            WriteLog op, nm, method, url, st, ErrSummary(resText)
        End If
    Next
    Application.StatusBar = False
    ' 実行後のSCMを取得し、②取得（実行前）と比較して⑤差分チェックへ
    BuildDiffCore True
    Application.StatusBar = False
    Worksheets(SH_DIFF).Activate
    MsgBox "成功 " & ok & " 件／失敗 " & ng & " 件／スキップ " & skipN & " 件" & vbLf & _
           "⑤設定差分チェックに、実行後のSCMと②取得（実行前）の差分を出力しました。確認してから、SCMの画面でPushしてください。", vbInformation
End Sub

' 文字列中の {{id:name}} を idMap（folder|position|name → id）で置換。未解決があれば False
Private Function ResolveIds(ByRef s As String, ByVal sk As String, ByVal idMap As Object) As Boolean
    Dim p As Long, q As Long, nm As String
    Do
        p = InStr(s, "{{id:")
        If p = 0 Then Exit Do
        q = InStr(p, s, "}}")
        If q = 0 Then Exit Do
        nm = Mid$(s, p + 5, q - p - 5)
        If Not idMap.Exists(sk & "|" & nm) Then Exit Function
        s = Left$(s, p - 1) & idMap(sk & "|" & nm) & Mid$(s, q + 2)
    Loop
    ResolveIds = True
End Function

Private Function IdFromResponse(ByVal resText As String) As String
    Dim res As Object, d As Variant
    On Error GoTo Done
    If Left$(Trim$(resText), 1) <> "{" Then Exit Function
    Set res = ParseJson(resText)
    If res Is Nothing Then Exit Function
    If res.Exists("id") Then IdFromResponse = Nz(res, "id"): Exit Function
    If res.Exists("data") Then
        If IsObject(res("data")) Then
            For Each d In res("data")
                If IsObject(d) Then IdFromResponse = Nz(d, "id"): Exit Function
            Next
        End If
    End If
Done:
End Function

Private Function FindIdByName(ByVal folder As String, ByVal pos As String, ByVal nm As String) As String
    Dim items As Collection, item As Variant, st As Long, msg As String
    Set items = GetAll(BaseUrl(), "folder=" & Enc(folder) & "&position=" & pos & "&name=" & Enc(nm), "追加後のid確認", st, msg)
    If items Is Nothing Then Exit Function
    For Each item In items
        If Nz(item, "name") = nm Then FindIdByName = Nz(item, "id"): Exit Function
    Next
End Function

Public Sub ClearJson()
    ClearJsonRows Worksheets(SH_JSON)
End Sub

Private Sub ClearJsonRows(ByVal ws As Worksheet)
    ws.Range(ws.Cells(JS_FIRST, 1), ws.Cells(ws.Rows.Count, J_LAST)).Clear
End Sub

'==========================================================
' ⑤ 設定差分チェック
'   実行前プレビュー : ②取得 ⇔ ③設定（削除予定の行を除く）
'   実行後チェック   : ②取得（実行前） ⇔ SCMから再取得（実行後）。④の操作と③の並び順で判定
'==========================================================
Public Sub BuildDiffPreview()
    If BuildDiffCore(False) Then
        Worksheets(SH_DIFF).Activate
        MsgBox "実行前プレビューを作成しました（②取得 と ③設定 の比較）。", vbInformation
    End If
End Sub

Public Sub BuildDiffCheck()
    If Not EnsureToken() Then Exit Sub
    If BuildDiffCore(True) Then
        Application.StatusBar = False
        Worksheets(SH_DIFF).Activate
        MsgBox "SCMから再取得して、②取得（実行前）との差分を作成しました。", vbInformation
    End If
End Sub

Private Function BuildDiffCore(ByVal live As Boolean) As Boolean
    Dim bGroups As Collection, aGroups As Collection, raw As Collection, failN As Long, ops As Object, expPred As Object
    Dim g As Variant, item As Variant, c As Collection, n As Long
    Set bGroups = ReadGroups(Worksheets(SH_GET), False, False)
    If live Then
        Set raw = FetchRuleGroups("差分チェック", failN)
        If raw Is Nothing Then Exit Function
        Set aGroups = New Collection
        For Each g In raw
            Set c = New Collection
            For Each item In g(3)
                c.Add Array(Nz(item, "id"), RuleVals(item), 0, "")
            Next
            aGroups.Add Array(g(0), g(1), g(2), c)
        Next
        Set ops = OpsFromJson()
    Else
        Set aGroups = ReadGroups(Worksheets(SH_SET), True, True)
        Set ops = OpsFromSet()
    End If
    Set expPred = PredMap()
    Application.ScreenUpdating = False
    n = WriteDiff(bGroups, aGroups, ops, expPred, live)
    Application.ScreenUpdating = True
    WriteLog "差分チェック", IIf(live, "実行後（SCM再取得）", "実行前プレビュー"), IIf(live, "GET", ""), BaseUrl(), 200, n & " 件の変化"
    BuildDiffCore = True
End Function

' ④JSON化の操作：キー（id、追加は new|folder|position|name）→ Array(操作, HTTP)
Private Function OpsFromJson() As Object
    Dim d As Object, ws As Worksheet, r As Long, k As String, op As String, st As String, cur As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets(SH_JSON)
    For r = JS_FIRST To LastRow(ws, 4)
        op = SafeStr(ws.Cells(r, 2).Value)
        If op <> "" Then
            If SafeStr(ws.Cells(r, 10).Value) <> "" Then
                k = SafeStr(ws.Cells(r, 10).Value)
            Else
                k = "new|" & SafeStr(ws.Cells(r, 9).Value) & "|" & SafeStr(ws.Cells(r, 3).Value)
            End If
            If op = "追加(位置)" Then op = "追加"
            st = SafeStr(ws.Cells(r, 8).Value)
            If d.Exists(k) Then
                cur = d(k)
                If InStr(cur(0), op) = 0 Then cur(0) = cur(0) & "+" & op
                If Left$(cur(1), 1) = "2" Or cur(1) = "" Then cur(1) = st      ' 失敗があれば失敗を残す
                d(k) = cur
            Else
                d(k) = Array(op, st)
            End If
        End If
    Next
    Set OpsFromJson = d
End Function

Private Function OpsFromSet() As Object
    Dim d As Object, g As Variant, rec As Variant, k As String
    Set d = CreateObject("Scripting.Dictionary")
    For Each g In ReadGroups(Worksheets(SH_SET), True, False)
        For Each rec In g(3)
            If rec(3) <> "" Then
                If rec(0) <> "" Then k = rec(0) Else k = "new|" & g(0) & "|" & g(1) & "|" & Trim$(rec(1)(0))
                d(k) = Array(rec(3), "")
            End If
        Next
    Next
    Set OpsFromSet = d
End Function

' ③設定の並び順：folder|position|name → 直前のルール名（先頭は空）
'   削除する行・NGで登録しない追加行は並びに含めない
Private Function PredMap() As Object
    Dim d As Object, ws As Worksheet, g As Variant, rec As Variant, prev As String, judge As String, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets(SH_SET)
    For Each g In ReadGroups(ws, True, False)
        prev = ""
        For Each rec In g(3)
            judge = SafeStr(ws.Cells(rec(2), S_JUDGE).Value)
            nm = Trim$(rec(1)(0))
            If rec(3) = "削除" And judge <> "NG" Then
            ElseIf rec(3) = "追加" And judge = "NG" Then
            ElseIf rec(0) = "" And rec(3) <> "追加" Then
            Else
                d(g(0) & "|" & g(1) & "|" & nm) = prev
                prev = nm
            End If
        Next
    Next
    Set PredMap = d
End Function

' 差分シートの作成。戻り値=変化のあった件数
Private Function WriteDiff(ByVal bGroups As Collection, ByVal aGroups As Collection, ByVal ops As Object, _
                           ByVal expPred As Object, ByVal live As Boolean) As Long
    Dim ws As Worksheet, order As Object, g As Variant, sk As Variant, r As Long, i As Long, j As Long, t As Long
    Dim bRecs As Collection, aRecs As Collection, bIdx As Object, aIds As Object, nA As Long, nB As Long
    Dim bArr() As Variant, aArr() As Variant, aB() As Long, moved() As Boolean, attach As Object, topDel As Collection
    Dim seen As Object, label As String, x As Variant, changed As Long, k As Variant, newKey As String
    Dim cntAll As Long, cntSame As Long, cntAdd As Long, cntMod As Long, cntDel As Long, cntChk As Long, cntMove As Long
    Set ws = Worksheets(SH_DIFF)
    Set order = CreateObject("Scripting.Dictionary")
    For Each g In aGroups
        order(g(0) & "|" & g(1)) = Array(g(0), g(1), g(2), g(3), New Collection)
    Next
    For Each g In bGroups
        If order.Exists(g(0) & "|" & g(1)) Then
            x = order(g(0) & "|" & g(1))
            Set x(4) = g(3)
            order(g(0) & "|" & g(1)) = x
        Else
            order(g(0) & "|" & g(1)) = Array(g(0), g(1), g(2), New Collection, g(3))
        End If
    Next
    ClearData ws, D_HF + 1
    Set seen = CreateObject("Scripting.Dictionary")
    r = FIRST
    For Each sk In order.Keys
        x = order(sk)
        Set aRecs = x(3): Set bRecs = x(4)
        nA = aRecs.Count: nB = bRecs.Count
        WriteHeaderRow ws, r, D_SCOPE, D_LAST, D_HF, CStr(x(0)), CStr(x(1)), CStr(x(2)), nA
        r = r + 1
        ' 実行前（②）の順番
        Set bIdx = CreateObject("Scripting.Dictionary")
        If nB > 0 Then ReDim bArr(1 To nB)
        i = 0
        For Each g In bRecs
            i = i + 1
            bArr(i) = g
            If g(0) <> "" Then bIdx(g(0)) = i
        Next
        ' 実行後（またはプレビュー）
        Set aIds = CreateObject("Scripting.Dictionary")
        If nA > 0 Then
            ReDim aArr(1 To nA): ReDim aB(1 To nA): ReDim moved(1 To nA)
        End If
        i = 0
        For Each g In aRecs
            i = i + 1
            aArr(i) = g
            If g(0) <> "" Then
                aIds(g(0)) = i
                If bIdx.Exists(g(0)) Then aB(i) = bIdx(g(0))
            End If
        Next
        MarkMoved aB, moved, nA
        ' 削除されたものは、②で直前にあった（残っている）ルールの下に表示する
        Set attach = CreateObject("Scripting.Dictionary")
        Set topDel = New Collection
        For j = 1 To nB
            If Not aIds.Exists(bArr(j)(0)) Then
                For t = j - 1 To 0 Step -1
                    If t = 0 Then
                        topDel.Add j
                    ElseIf aIds.Exists(bArr(t)(0)) Then
                        If Not attach.Exists(bArr(t)(0)) Then attach.Add bArr(t)(0), New Collection
                        attach(bArr(t)(0)).Add j
                        Exit For
                    End If
                Next
            End If
        Next
        For Each k In topDel
            WriteDelRow ws, r, CStr(x(2)), bArr(k), CLng(k), ops, live, cntDel, cntChk, CStr(x(0))
        Next
        For i = 1 To nA
            If WriteAfterRow(ws, r, CStr(sk), CStr(x(2)), aArr, i, bArr, aB(i), moved(i), ops, expPred, live, seen, _
                             cntSame, cntAdd, cntMod, cntChk, cntMove) Then changed = changed + 1
            cntAll = cntAll + 1
            If aArr(i)(0) <> "" Then
                If attach.Exists(aArr(i)(0)) Then
                    For Each k In attach(aArr(i)(0))
                        WriteDelRow ws, r, CStr(x(2)), bArr(k), CLng(k), ops, live, cntDel, cntChk, CStr(x(0))
                    Next
                End If
            End If
        Next
    Next
    changed = changed + cntDel
    ' 登録したはずが SCM に無い（未反映の追加）は実行ログへ
    If live Then
        For Each k In ops.Keys
            If Left$(k, 4) = "new|" And Not seen.Exists(k) Then
                cntChk = cntChk + 1
                WriteLog "差分チェック", Mid$(k, InStrRev(k, "|") + 1), "", "", 0, "未反映（追加したルールがSCMにありません。HTTP " & ops(k)(1) & "）"
            End If
        Next
    End If
    If r > FIRST Then ws.Range(ws.Cells(FIRST, 1), ws.Cells(r - 1, D_LAST)).VerticalAlignment = xlTop
    ws.Range("B3").Value = Now
    ws.Range("D3").Value = cntAll
    ws.Range("F3").Value = cntSame
    ws.Range("H3").Value = cntAdd
    ws.Range("J3").Value = IIf(live, "実行後チェック（SCM再取得）", "実行前プレビュー（②取得 ⇔ ③設定）")
    ws.Range("D4").Value = cntMod
    ws.Range("F4").Value = cntDel
    ws.Range("H4").Value = cntChk
    ws.Range("J4").Value = cntMove
    WriteDiff = changed
End Function

' 並び順が変わったルールの判定（②の順番の最長増加部分列に入らないものを「移動」とみなす）
Private Sub MarkMoved(ByRef aB() As Long, ByRef moved() As Boolean, ByVal nA As Long)
    Dim L() As Long, P() As Long, i As Long, j As Long, best As Long
    If nA = 0 Then Exit Sub
    ReDim L(1 To nA): ReDim P(1 To nA)
    For i = 1 To nA
        If aB(i) > 0 Then
            L(i) = 1
            For j = 1 To i - 1
                If aB(j) > 0 And aB(j) < aB(i) And L(j) + 1 > L(i) Then L(i) = L(j) + 1: P(i) = j
            Next
            If best = 0 Then
                best = i
            ElseIf L(i) > L(best) Then
                best = i
            End If
            moved(i) = True
        End If
    Next
    i = best
    Do While i > 0
        moved(i) = False
        i = P(i)
    Loop
End Sub

' 実行後（またはプレビュー）の1ルールを出力。戻り値=変化あり
Private Function WriteAfterRow(ByVal ws As Worksheet, ByRef r As Long, ByVal sk As String, ByVal label As String, _
                               ByRef aArr() As Variant, ByVal i As Long, ByRef bArr() As Variant, ByVal bi As Long, ByVal mv As Boolean, _
                               ByVal ops As Object, ByVal expPred As Object, ByVal live As Boolean, ByVal seen As Object, _
                               ByRef cntSame As Long, ByRef cntAdd As Long, ByRef cntMod As Long, ByRef cntChk As Long, ByRef cntMove As Long) As Boolean
    Dim rec As Variant, a As Variant, b As Variant, f As Long, fields As String, olds As String, chg As String
    Dim key As String, opStr As String, st As String, judge As String, nm As String, actPred As String, posOK As Boolean
    rec = aArr(i)
    a = rec(1)
    nm = Trim$(a(0))
    If bi > 0 Then
        b = bArr(bi)(1)
        For f = 0 To NF
            If Cmp(f, CStr(a(f))) <> Cmp(f, CStr(b(f))) Then
                fields = fields & FieldLabel(f) & " "
                olds = olds & vbLf & FieldLabel(f) & ": " & Replace(CStr(b(f)), vbLf, ", ")
            End If
        Next
        If fields <> "" Then chg = "変更"
        key = rec(0)
    Else
        chg = "追加"
        key = "new|" & sk & "|" & nm
    End If
    seen(key) = True
    If ops.Exists(key) Then opStr = ops(key)(0): st = ops(key)(1)
    ' 位置の確認（③設定の並びと比べる）
    If i > 1 Then actPred = Trim$(aArr(i - 1)(1)(0))
    posOK = True
    If expPred.Exists(sk & "|" & nm) Then posOK = (expPred(sk & "|" & nm) = actPred)
    ' 判定
    If live Then
        If opStr = "" Then
            If chg <> "" Then
                judge = "想定外（④に無い" & chg & "）"
            ElseIf Not posOK Then
                judge = "要確認：順番が③設定と異なります"
            End If
        Else
            If InStr(opStr, "削除") > 0 Then
                judge = "未反映（削除されていません。HTTP " & st & "）"
            ElseIf InStr(opStr, "変更") > 0 And chg = "" Then
                If Left$(st, 1) = "2" Then judge = "変化なし（既存と同じ内容）" Else judge = "未反映（HTTP " & st & "）"
            ElseIf InStr(opStr, "変更") = 0 And chg = "変更" Then
                judge = "想定外（項目が変わっています）"
            ElseIf Not posOK Then
                judge = "要確認：順番が③設定と異なります（直前のルール：" & IIf(actPred = "", "なし（先頭）", actPred) & "）"
            Else
                judge = "想定どおり"
            End If
        End If
    Else
        If chg = "変更" And InStr(opStr, "変更") = 0 Then
            judge = "要確認：内容が変わっていますが操作が「変更」ではありません"
        ElseIf chg = "" And InStr(opStr, "変更") > 0 Then
            judge = "変化なし（既存と同じ内容）"
        ElseIf chg = "追加" And opStr <> "追加" Then
            judge = "要確認：操作が「追加」ではありません"
        End If
    End If
    ' 書き込み
    ws.Range(ws.Cells(r, 2), ws.Cells(r, D_LAST)).NumberFormat = "@"
    ws.Cells(r, 1).Value = i
    ws.Cells(r, D_SCOPE).Value = label
    ws.Cells(r, D_CHG).Value = IIf(chg = "", IIf(mv, "移動", "変化なし"), chg & IIf(mv, "・移動", ""))
    ws.Cells(r, D_ORDER).Value = IIf(bi > 0, CStr(bi), "") & "→" & i
    For f = 0 To NF
        ws.Cells(r, D_NAME + f).Value = a(f)
    Next
    ws.Cells(r, D_FIELDS).Value = Trim$(fields)
    ws.Cells(r, D_BEFORE).Value = Mid$(olds, 2)
    ws.Cells(r, D_OP).Value = opStr
    ws.Cells(r, D_HTTP).Value = st
    ws.Cells(r, D_JUDGE).Value = judge
    DiffRowFormat ws, r
    IndentRow ws, r, D_SCOPE, D_NAME, Split(sk, "|")(0)
    If chg = "" And mv Then PaintDiffRow ws, r, D_LAST, "変更", judge Else PaintDiffRow ws, r, D_LAST, chg, judge
    If chg = "変更" Then
        For f = 0 To NF
            If Cmp(f, CStr(a(f))) <> Cmp(f, CStr(b(f))) Then MarkChangedCell ws.Cells(r, D_NAME + f)
        Next
    End If
    If mv Then MarkChangedCell ws.Cells(r, D_ORDER)
    If judge <> "" And judge <> "想定どおり" And Left$(judge, 4) <> "変化なし" Then cntChk = cntChk + 1
    Select Case chg
        Case "": If Not mv Then cntSame = cntSame + 1
        Case "追加": cntAdd = cntAdd + 1
        Case "変更": cntMod = cntMod + 1
    End Select
    If mv Then cntMove = cntMove + 1
    WriteAfterRow = (chg <> "" Or mv)
    r = r + 1
End Function

' 削除された（プレビューでは削除予定の）ルールを出力
Private Sub WriteDelRow(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal rec As Variant, ByVal bi As Long, _
                        ByVal ops As Object, ByVal live As Boolean, ByRef cntDel As Long, ByRef cntChk As Long, ByVal folder As String)
    Dim f As Long, opStr As String, st As String, judge As String
    If ops.Exists(rec(0)) Then opStr = ops(rec(0))(0): st = ops(rec(0))(1)
    If live Then
        If InStr(opStr, "削除") > 0 Then judge = "想定どおり" Else judge = "想定外（削除されています）"
        If judge <> "想定どおり" Then WriteLog "差分チェック", Trim$(rec(1)(0)), "", "", 0, "削除されています：" & judge
    Else
        If InStr(opStr, "削除") = 0 Then judge = "要確認：③設定に行がありません（削除する場合は操作で「削除」）"
    End If
    ws.Range(ws.Cells(r, 2), ws.Cells(r, D_LAST)).NumberFormat = "@"
    ws.Cells(r, D_SCOPE).Value = label
    ws.Cells(r, D_CHG).Value = "削除"
    ws.Cells(r, D_ORDER).Value = bi & "→"
    For f = 0 To NF
        ws.Cells(r, D_NAME + f).Value = rec(1)(f)
    Next
    ws.Cells(r, D_OP).Value = opStr
    ws.Cells(r, D_HTTP).Value = st
    ws.Cells(r, D_JUDGE).Value = judge
    DiffRowFormat ws, r
    IndentRow ws, r, D_SCOPE, D_NAME, folder
    PaintDiffRow ws, r, D_LAST, "削除", judge
    ws.Range(ws.Cells(r, D_NAME), ws.Cells(r, D_NAME + NF)).Font.Strikethrough = True
    If judge <> "" And judge <> "想定どおり" Then cntChk = cntChk + 1
    cntDel = cntDel + 1
    r = r + 1
End Sub

Private Sub DiffRowFormat(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, D_LAST))
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Color = vbBlack
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = &HBFBFBF
    End With
End Sub

'==========================================================
' 復号プロファイル一覧（設定シートの profile 列のプルダウン元）
'==========================================================
Public Sub FetchProfilesBtn()
    Dim n As Long
    If Not EnsureToken() Then Exit Sub
    n = FetchProfiles()
    Application.StatusBar = False
    If n < 0 Then
        MsgBox "復号プロファイルの取得に失敗しました。実行ログを確認してください。", vbExclamation
    Else
        Worksheets(SH_PROF).Activate
        MsgBox "復号プロファイル " & n & " 件を取得しました。", vbInformation
    End If
End Sub

' 戻り値：件数（すべて失敗なら -1）
Private Function FetchProfiles() As Long
    Dim folders As Object, sc As Variant, f As Variant, items As Collection, item As Variant, seen As Object
    Dim ws As Worksheet, outRows As New Collection, k As String, st As Long, msg As String, okN As Long, r As Long, x As Variant, c As Long
    Set folders = CreateObject("Scripting.Dictionary")
    folders("Shared") = True
    For Each sc In SelScopes()
        folders(CStr(sc(1))) = True
    Next
    Set seen = CreateObject("Scripting.Dictionary")
    For Each f In folders.Keys
        Prog "復号プロファイル：" & f & " を取得中…"
        Set items = GetAll(ApiBase() & EP_PROF, "folder=" & Enc(CStr(f)), "プロファイル取得", st, msg)
        If Not items Is Nothing Then
            okN = okN + 1
            For Each item In items
                k = Nz(item, "id")
                If k = "" Then k = "name:" & Nz(item, "name")
                If Not seen.Exists(k) Then
                    seen(k) = True
                    outRows.Add Array(Nz(item, "name"), IIf(Nz(item, "folder") = "", CStr(f), Nz(item, "folder")), _
                                   IIf(IsPredef(item), "定義済み", "カスタム"), _
                                   TrueKeys(item, "ssl_forward_proxy"), TrueKeys(item, "ssl_inbound_proxy"), _
                                   TrueKeys(item, "ssl_no_proxy"), ScalarKeys(item, "ssl_protocol_settings"), Nz(item, "id"))
                End If
            Next
        End If
    Next
    If okN = 0 Then FetchProfiles = -1: Exit Function
    Set ws = Worksheets(SH_PROF)
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(ws.Rows.Count, 9)).Clear
    r = FIRST
    For Each x In outRows
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 9)).NumberFormat = "@"
        ws.Cells(r, 1).Value = r - FIRST + 1
        For c = 0 To 7
            ws.Cells(r, 2 + c).Value = x(c)
        Next
        r = r + 1
    Next
    If r > FIRST Then
        With ws.Range(ws.Cells(FIRST, 1), ws.Cells(r - 1, 9))
            .Borders.LineStyle = xlContinuous
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    ws.Range("B3").Value = Now
    WriteLog "プロファイル取得", "", "GET", ApiBase() & EP_PROF, 200, outRows.Count & " 件"
    FetchProfiles = outRows.Count
End Function

Private Function IsPredef(ByVal item As Object) As Boolean
    If Nz(item, "id") = "" Then IsPredef = True: Exit Function
    If InStr(LCase$(Nz(item, "folder") & " " & Nz(item, "snippet")), "predefined") > 0 Then IsPredef = True
End Function

' 子オブジェクトのうち true の項目名（改行区切り）
Private Function TrueKeys(ByVal item As Object, ByVal key As String) As String
    Dim d As Object, k As Variant, s As String
    If Not item.Exists(key) Then Exit Function
    If Not IsObject(item(key)) Then Exit Function
    Set d = item(key)
    If TypeName(d) = "Collection" Then Exit Function
    For Each k In d.Keys
        If Not IsObject(d(k)) Then
            If VarType(d(k)) = vbBoolean Then
                If d(k) Then s = s & vbLf & k
            End If
        End If
    Next
    If s = "" Then TrueKeys = "（すべて false／既定）" Else TrueKeys = Mid$(s, 2)
End Function

Private Function ScalarKeys(ByVal item As Object, ByVal key As String) As String
    Dim d As Object, k As Variant, s As String
    If Not item.Exists(key) Then Exit Function
    If Not IsObject(item(key)) Then Exit Function
    Set d = item(key)
    If TypeName(d) = "Collection" Then Exit Function
    For Each k In d.Keys
        If Not IsObject(d(k)) Then s = s & vbLf & k & ": " & JsonOf(d(k))
    Next
    ScalarKeys = Mid$(s, 2)
End Function

'==========================================================
' 選択肢（ダブルクリック選択の候補。オブジェクトは基本設定のオブジェクト取得folder＝通常 Shared）
'==========================================================
Public Sub FetchChoicesBtn()
    Dim ng As Long
    If Not EnsureToken() Then Exit Sub
    ng = FetchChoices()
    Application.StatusBar = False
    Worksheets(SH_CHO).Activate
    If ng > 0 Then
        MsgBox "選択肢を取得しました。" & ng & " 種類は取得に失敗しました（実行ログ参照。その列は前の値のまま）。", vbExclamation
    Else
        MsgBox "選択肢を取得しました。", vbInformation
    End If
End Sub

' 戻り値：失敗した種類の数
Private Function FetchChoices() As Long
    Dim stU As Long, msgU As String
    Dim q As String, folder As String, ng As Long, custom As Object, items As Collection, item As Variant
    Dim vals As Collection, typs As Collection, nm As String, tk As Variant, t As String
    folder = CfgOpt("CFG_OBJ_FOLDER")
    If folder = "" Then folder = "Shared"
    q = "folder=" & Enc(folder)
    If Not ChoiceCol(CH_ADDR, "/sse/config/v1/addresses", q, Nothing) Then ng = ng + 1
    If Not ChoiceCol(CH_ADDRG, "/sse/config/v1/address-groups", q, Nothing) Then ng = ng + 1
    If Not ChoiceCol(CH_REGION, "/sse/config/v1/regions", q, Nothing) Then ng = ng + 1
    If Not ChoiceCol(CH_SVC, "/sse/config/v1/services", q, Nothing) Then ng = ng + 1
    If Not ChoiceCol(CH_SVCG, "/sse/config/v1/service-groups", q, Nothing) Then ng = ng + 1
    Set custom = CreateObject("Scripting.Dictionary")
    If Not ChoiceCol(CH_URLC, "/sse/config/v1/url-categories", q, custom) Then ng = ng + 1
    If Not ChoiceCol(CH_TAG, "/sse/config/v1/tags", q, Nothing) Then ng = ng + 1
    If Not ChoiceCol(CH_HIP, "/sse/config/v1/hip-profiles", q, Nothing) Then ng = ng + 1
    If Not ChoiceCol(CH_CERT, "/sse/config/v1/certificates", q, Nothing) Then ng = ng + 1
    ' EDL（名前と種類）
    Prog "選択肢：EDL を取得中…"
    Set items = ChoiceGet("/sse/config/v1/external-dynamic-lists", q)
    If items Is Nothing Then
        ng = ng + 1
    Else
        Set vals = New Collection: Set typs = New Collection
        For Each item In items
            t = ""
            If item.Exists("type") Then
                If IsObject(item("type")) Then
                    For Each tk In item("type").Keys
                        t = CStr(tk)
                    Next
                End If
            End If
            vals.Add Nz(item, "name"): typs.Add t
        Next
        WriteChoiceCol CH_EDL, vals
        WriteChoiceCol CH_EDLTYPE, typs
    End If
    ' 事前定義URLカテゴリ（url-filtering-categories のうちカスタム以外。0件なら既存の一覧を残す）
    Prog "選択肢：事前定義URLカテゴリ を取得中…"
    '   403（サービスアカウントのロールに参照権限が無い）のときは失敗に数えず、⑧の一覧（初期値または前回分）をそのまま使う
    Set items = GetAll(ApiBase() & "/sse/config/v1/url-filtering-categories", q, "選択肢取得", stU, msgU)
    If items Is Nothing Then
        If stU = 403 Then
            WriteLog "選択肢取得", "事前定義URLカテゴリ", "GET", WithQuery(ApiBase() & "/sse/config/v1/url-filtering-categories", q), 403, _
                     "参照権限なし（ロールを確認）。⑧の事前定義URLカテゴリ列は今の一覧のまま使います"
        Else
            ng = ng + 1
        End If
    Else
        Set vals = New Collection
        For Each item In items
            nm = Nz(item, "value")
            If nm = "" Then nm = Nz(item, "name")
            If nm <> "" And Not custom.Exists(nm) And InStr(LCase$(Nz(item, "type")), "custom") = 0 Then vals.Add nm
        Next
        If vals.Count > 0 Then WriteChoiceCol CH_URLP, vals
    End If
    Worksheets(SH_CHO).Range("B3").Value = Now
    FetchChoices = ng
End Function

Private Function ChoiceGet(ByVal ep As String, ByVal q As String) As Collection
    Dim st As Long, msg As String
    Set ChoiceGet = GetAll(ApiBase() & ep, q, "選択肢取得", st, msg)
    If Not ChoiceGet Is Nothing Then WriteLog "選択肢取得", "", "GET", WithQuery(ApiBase() & ep, q), 200, ChoiceGet.Count & " 件"
End Function

' 1種類を取得して列に書く（names を渡すと名前を記録）
Private Function ChoiceCol(ByVal col As Long, ByVal ep As String, ByVal q As String, ByVal nameSet As Object) As Boolean
    Dim items As Collection, item As Variant, vals As New Collection, nm As String
    Prog "選択肢：" & ep & " を取得中…"
    Set items = ChoiceGet(ep, q)
    If items Is Nothing Then Exit Function
    For Each item In items
        nm = Nz(item, "name")
        If nm <> "" Then
            vals.Add nm
            If Not nameSet Is Nothing Then nameSet(nm) = True
        End If
    Next
    WriteChoiceCol col, vals
    ChoiceCol = True
End Function

Private Sub WriteChoiceCol(ByVal col As Long, ByVal vals As Collection)
    Dim ws As Worksheet, a() As Variant, i As Long
    Set ws = Worksheets(SH_CHO)
    ws.Range(ws.Cells(FIRST, col), ws.Cells(ws.Rows.Count, col)).ClearContents
    If vals.Count = 0 Then Exit Sub
    ReDim a(1 To vals.Count, 1 To 1)
    For i = 1 To vals.Count
        a(i, 1) = vals(i)
    Next
    ws.Cells(FIRST, col).Resize(vals.Count, 1).NumberFormat = "@"
    ws.Cells(FIRST, col).Resize(vals.Count, 1).Value = a
End Sub

'==========================================================
' メンバー選択画面（ダブルクリック → チェックボックス一覧）
'   必要モジュール : modDecRule（このモジュール） / clsPickerEvents / frmMemberPick
'   ブックを開くと Auto_Open で自動有効化。効かない時は PickerStart を実行（または Ctrl+Shift+M）
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

Public Sub PickerFromSelection()
    If mEvents Is Nothing Then Set mEvents = New clsPickerEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Not TryPick(ActiveCell) Then
        MsgBox "このセルはチェックリスト選択の対象外です。", vbInformation
    End If
End Sub

Public Function TryPick(ByVal Target As Range) As Boolean
    Dim ttl As String
    If Target.CountLarge <> 1 Then Exit Function
    If Not Target.Worksheet.Parent Is ThisWorkbook Then Exit Function
    Set mNames = CreateObject("Scripting.Dictionary")   ' 大文字・小文字を区別
    mExclude = ""
    If Not PickerRule(Target.Worksheet, Target, ttl) Then Exit Function
    TryPick = True
    ShowPicker Target, ttl
End Function

' 対象列：③設定の タグ／ゾーン／アドレス／ユーザー／HIP／サービス／URLカテゴリ
Private Function PickerRule(ByVal ws As Worksheet, ByVal c As Range, ByRef ttl As String) As Boolean
    Dim fi As Long, k As String
    If ws.Name <> SH_SET Then Exit Function
    If c.Row < FIRST Then Exit Function
    If SafeStr(ws.Cells(c.Row, 1).Value) = HDR Then Exit Function
    InitF
    fi = c.Column - S_F0
    If fi < 0 Or fi > NF - 1 Then Exit Function
    If mFT(fi) <> "a" Then Exit Function
    k = mFK(fi)
    Select Case k
        Case "from", "to"
            AddFix "any"
            AddSrc SH_CHO, CH_ZONE, FIRST, "ゾーン"
        Case "source", "destination"
            AddFix "any"
            AddSrc SH_CHO, CH_ADDR, FIRST, "アドレス"
            AddSrc SH_CHO, CH_ADDRG, FIRST, "アドレスグループ"
            AddEdl "ip"
            AddSrc SH_CHO, CH_REGION, FIRST, "リージョン"
        Case "source_user"
            AddFix "any"
            AddFix "known-user"
            AddFix "unknown"
            AddFix "pre-logon"
            AddSrc SH_CHO, CH_USER, FIRST, "ユーザー／グループ"
        Case "source_hip", "destination_hip"
            AddFix "any"
            AddSrc SH_CHO, CH_HIP, FIRST, "HIPプロファイル"
        Case "service"
            AddFix "any"
            AddFix "service-http"
            AddFix "service-https"
            AddSrc SH_CHO, CH_SVC, FIRST, "サービス"
            AddSrc SH_CHO, CH_SVCG, FIRST, "サービスグループ"
        Case "category"
            AddFix "any"
            AddSrc SH_CHO, CH_URLC, FIRST, "URLカテゴリ（カスタム）"
            AddSrc SH_CHO, CH_URLP, FIRST, "URLカテゴリ（事前定義）"
            AddEdl "url"
        Case "tag"
            AddSrc SH_CHO, CH_TAG, FIRST, "タグ"
    End Select
    AddUsed SH_GET, G_F0 + fi, "使用中（②取得）"
    ttl = mFL(fi) & " の選択：" & SafeStr(ws.Cells(c.Row, S_NAME).Value)
    PickerRule = True
End Function

Private Sub AddFix(ByVal nm As String)
    If Not mNames.Exists(nm) Then mNames.Add nm, "既定値"
End Sub

' EDL のうち種類（ip / url）が合うもの
Private Sub AddEdl(ByVal want As String)
    Dim ws As Worksheet, v As Variant, cv As Variant, i As Long, nm As String
    If Not ReadCols(SH_CHO, CH_EDL, FIRST, CH_EDLTYPE, ws, v, cv) Then Exit Sub
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" And Not mNames.Exists(nm) Then
            If InStr(LCase$(SafeStr(cv(i, 1))), want) > 0 Then mNames.Add nm, "EDL（" & SafeStr(cv(i, 1)) & "）"
        End If
    Next
End Sub

' ②取得で使われている値（セル内の複数値を分解）
Private Sub AddUsed(ByVal sh As String, ByVal col As Long, ByVal kind As String)
    Dim ws As Worksheet, v As Variant, cv As Variant, i As Long, a As Variant, j As Long, nm As String
    If Not ReadCols(sh, col, FIRST, 0, ws, v, cv) Then Exit Sub
    For i = 1 To UBound(v, 1)
        a = SplitMembers(SafeStr(v(i, 1)))
        For j = 0 To UBound(a)
            nm = CStr(a(j))
            If nm <> "" And Not mNames.Exists(nm) Then mNames.Add nm, kind
        Next
    Next
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
    Dim f As frmMemberPick, cur As String, res As String
    cur = SafeStr(Target.Value)
    Set f = New frmMemberPick
    f.Setup ttl, mNames, SplitItems(cur), vbLf      ' このツールはセル内改行区切りで書き戻す
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
            Target.NumberFormat = "@"
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
    If IsObject(v) Then Exit Function
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

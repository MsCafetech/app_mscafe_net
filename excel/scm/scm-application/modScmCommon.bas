Attribute VB_Name = "modScmCommon"
'==========================================================
' modScmCommon : SCM Excel ツール共通処理（全ツールで同じものをインポート）
'   トークン取得 / HTTP / 基本設定の読出し / 実行ログ / エラー要約 /
'   curl生成 / ボタン配置 / 進捗表示 / 差分確認の色付け / メンバー文字列の分解 など
'   JSONパーサ（ParseJson / ParseJsonSkipPredef / JsonStr）も末尾に含む（旧 modJson を統合）
'   シート名「基本設定」「実行ログ」と名前定義 CFG_* は全ツール共通
'==========================================================
Option Explicit

' JSONパーサの読み取り位置（末尾の JSON セクションで使用）
Private mJs As String
Private mJp As Long

Private Const SH_LOG As String = "実行ログ"   ' シート名が「⑥実行ログ」のように番号付きでも可（LogWs）
Private Const LOG_FIRST As Long = 6     ' 実行ログ データ開始行

' 差分確認の色（行の背景色。文字は黒）
Public Const DIFF_ADD As Long = &HCEEFC6     ' 追加 = 緑 C6EFCE
Public Const DIFF_CHG As Long = &H99FFFF     ' 変更 = 黄 FFFF99
Public Const DIFF_DEL As Long = &HD9D9D9     ' 削除 = グレー D9D9D9
Public Const DIFF_NG As Long = &HCEC7FF      ' 要確認 = 赤 FFC7CE
Public Const DIFF_CELL As Long = &HC0FF&     ' 変更セル = オレンジ FFC000



'----------------------------------------------------------
' ボタン配置
'----------------------------------------------------------


Public Sub AddBtn(ByVal sh As String, ByVal addr As String, ByVal caption As String, ByVal macro As String)
    Dim ws As Worksheet, rng As Range, b As Object
    Set ws = Worksheets(sh)
    Set rng = ws.Range(addr)
    On Error Resume Next
    ws.Buttons("btn_" & macro).Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    b.Name = "btn_" & macro
    b.Caption = caption
    b.OnAction = macro
End Sub

' 旧版のボタンを削除（無ければ何もしない）。例：Push機能の削除後に旧ボタンを消す
Public Sub RemoveBtn(ByVal sh As String, ByVal macro As String)
    On Error Resume Next
    Worksheets(sh).Buttons("btn_" & macro).Delete
    On Error GoTo 0
End Sub


'----------------------------------------------------------
' ① アクセストークン
'----------------------------------------------------------


Public Sub GetAccessToken()
    If FetchToken() Then MsgBox "アクセストークンを取得しました。", vbInformation
End Sub

Public Function FetchToken() As Boolean
    Dim http As Object, url As String, body As String, res As Object
    url = Cfg("CFG_AUTH_URL")
    body = "grant_type=client_credentials&scope=tsg_id:" & Cfg("CFG_TSG_ID")
    On Error GoTo EH
    Set http = NewHttp()
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "Authorization", "Basic " & Base64(Cfg("CFG_CLIENT_ID") & ":" & Cfg("CFG_CLIENT_SECRET"))
    http.send body
    If http.Status <> 200 Then
        WriteLog "トークン取得", "", "POST", url, http.Status, http.responseText
        MsgBox "トークン取得に失敗しました（HTTP " & http.Status & "）。実行ログを確認してください。", vbExclamation
        Exit Function
    End If
    Set res = ParseJson(http.responseText)
    Range("CFG_TOKEN").Value = res("access_token")
    Range("CFG_TOKEN_AT").Value = Now
    Range("CFG_TOKEN_TTL").Value = res("expires_in")
    WriteLog "トークン取得", "", "POST", url, 200, "expires_in=" & res("expires_in")
    FetchToken = True
    Exit Function
EH:
    WriteLog "トークン取得", "", "POST", url, -1, Err.Description
    MsgBox "トークン取得でエラー：" & Err.Description, vbCritical
End Function

Public Function EnsureToken() As Boolean
    Dim tokenExp As Variant
    tokenExp = Range("CFG_TOKEN_EXP").Value
    If Cfg("CFG_TOKEN") = "" Or Not IsDate(tokenExp) Then
        EnsureToken = FetchToken()
    ElseIf Now > CDate(tokenExp) - TimeSerial(0, 1, 0) Then
        EnsureToken = FetchToken()
    Else
        EnsureToken = True
    End If
End Function


'----------------------------------------------------------
' HTTP
'----------------------------------------------------------


' 同期HTTP（body が空なら送信なし）。戻り値=HTTPステータス、通信エラーは -1
Public Function HttpCall(ByVal method As String, ByVal url As String, ByVal body As String, ByRef resText As String) As Long
    Dim http As Object
    On Error GoTo EH
    Set http = NewHttp()
    http.Open method, url, False
    http.setRequestHeader "Authorization", "Bearer " & Cfg("CFG_TOKEN")
    http.setRequestHeader "Accept", "application/json"
    If body <> "" Then
        http.setRequestHeader "Content-Type", "application/json"
        http.send body
    Else
        http.send
    End If
    resText = http.responseText
    HttpCall = http.Status
    Exit Function
EH:
    resText = Err.Description
    HttpCall = -1
End Function

Public Function NewHttp() As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If CfgOpt("CFG_PROXY") <> "" Then http.setProxy 2, CfgOpt("CFG_PROXY")
    http.setTimeouts 10000, 10000, 60000, 120000
    Set NewHttp = http
End Function

' エラー応答 {"_errors":[{"code":..,"message":..}]} を短い文字列に
Public Function ErrSummary(ByVal resText As String) As String
    Dim res As Object, e As Object, s As String
    ErrSummary = resText
    If Left$(Trim$(resText), 1) <> "{" Then Exit Function
    On Error GoTo Done
    Set res = ParseJson(resText)
    If res Is Nothing Then Exit Function
    If Not res.Exists("_errors") Then Exit Function
    For Each e In res("_errors")
        If s <> "" Then s = s & " / "
        s = s & Nz(e, "code") & ": " & Nz(e, "message")
        If Nz(e, "code") = "E009" Then s = s & "（セキュリティルールやGroup等から参照中のため削除不可）"
        If Nz(e, "code") = "E006" Then s = s & "（同名のオブジェクトが既に存在。定義済みのものと同名の可能性もあり）"
    Next
    ErrSummary = s
Done:
End Function

Public Function BuildCurl(ByVal method As String, ByVal url As String, ByVal body As String) As String
    Dim c As String
    c = "curl -s -X " & method & " '" & url & "' -H ""Authorization: Bearer $TOKEN"""
    If body <> "" Then
        c = c & " -H 'Content-Type: application/json' -d '" & Replace(body, "'", "'\''") & "'"
    End If
    BuildCurl = c
End Function


'----------------------------------------------------------
' 基本設定
'----------------------------------------------------------


' 名前定義の値（名前が無いとエラー）
Public Function Cfg(ByVal nm As String) As String
    Cfg = Trim$(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
End Function

' 名前定義の値（名前が無いブックでは空欄）
Public Function CfgOpt(ByVal nm As String) As String
    On Error Resume Next
    CfgOpt = Trim$(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
End Function

Public Function BaseUrl() As String
    Dim b As String
    b = Cfg("CFG_BASE_URL")
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    BaseUrl = b & Cfg("CFG_ENDPOINT")
End Function

Public Function CheckContainer() As Boolean
    If Cfg("CFG_CONT_NAME") = "" Then
        MsgBox "基本設定の「コンテナ名」(folder)を選択してください。", vbExclamation
        Exit Function
    End If
    CheckContainer = True
End Function

' コンテナ名が空欄なら条件なし（folder指定なし）
Public Function ContainerQuery() As String
    If CfgOpt("CFG_CONT_NAME") = "" Then Exit Function
    ContainerQuery = CfgOpt("CFG_CONT_TYPE") & "=" & Application.WorksheetFunction.EncodeURL(CfgOpt("CFG_CONT_NAME"))
End Function

Public Function WithQuery(ByVal url As String, ByVal q As String) As String
    If Left$(q, 1) = "&" Then q = Mid$(q, 2)
    If q = "" Then WithQuery = url Else WithQuery = url & "?" & q
End Function

Public Function ContainerOf(ByVal item As Object) As String
    Dim k As Variant
    For Each k In Array("folder", "snippet", "device")
        If item.Exists(k) Then
            ContainerOf = k & "=" & Nz(item, CStr(k))
            Exit Function
        End If
    Next
    ' レスポンスにfolder等が無い場合は取得条件を表示
    ContainerOf = ContainerQuery()
End Function


'----------------------------------------------------------
' 実行ログ
'----------------------------------------------------------


' 実行ログのシート（「実行ログ」が無ければ名前が「実行ログ」で終わるシート。例：⑥実行ログ）
Private Function LogWs() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set LogWs = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If Not LogWs Is Nothing Then Exit Function
    For Each ws In ThisWorkbook.Worksheets
        If Right$(ws.Name, Len(SH_LOG)) = SH_LOG Then
            Set LogWs = ws
            Exit Function
        End If
    Next
End Function

Public Sub WriteLog(ByVal op As String, ByVal nm As String, ByVal method As String, _
                     ByVal url As String, ByVal st As Long, ByVal msg As String)
    Dim ws As Worksheet, r As Long
    Set ws = LogWs()
    r = LastRow(ws, 1) + 1
    If r < LOG_FIRST Then r = LOG_FIRST
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).NumberFormat = "@"
    ws.Cells(r, 8).NumberFormat = "@"
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(r, 2).Value = op
    ws.Cells(r, 3).Value = nm
    ws.Cells(r, 4).Value = method
    ws.Cells(r, 5).Value = url
    ws.Cells(r, 6).Value = st
    ws.Cells(r, 7).Value = IIf(st >= 200 And st < 300, "成功", "失敗")
    ws.Cells(r, 8).Value = Left$(msg, 2000)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Borders.LineStyle = xlContinuous
End Sub

Public Sub ClearLog()
    If MsgBox("実行ログをクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    With LogWs()
        .Range(.Cells(LOG_FIRST, 1), .Cells(.Rows.Count, 8)).Clear
    End With
End Sub

'----------------------------------------------------------
' 進捗表示（ステータスバー）
'----------------------------------------------------------


' ステータスバー（Excel画面の左下）に進捗を表示し、画面の固まりを防ぐ
Public Sub Prog(ByVal msg As String)
    Application.StatusBar = msg
    DoEvents
End Sub

Public Function SecText(ByVal sec As Single) As String
    If sec >= 60 Then SecText = Int(sec / 60) & "分" & Format$(Int(sec) Mod 60, "00") & "秒" Else SecText = Int(sec) & "秒"
End Function


'----------------------------------------------------------
' 差分確認の色付け
'----------------------------------------------------------


' 差分確認の1行を色付け（追加=緑 / 変更=黄 / 削除=グレー、文字は黒）。判定が想定外なら最終列を赤
Public Sub PaintDiffRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, ByVal chg As String, ByVal judge As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        Select Case chg
            Case "追加": .Interior.Color = DIFF_ADD
            Case "変更": .Interior.Color = DIFF_CHG
            Case "削除": .Interior.Color = DIFF_DEL
        End Select
        .Font.Color = vbBlack
    End With
    If judge <> "" And judge <> "想定どおり" And Left$(judge, 4) <> "変化なし" Then
        ws.Cells(r, lastCol).Interior.Color = DIFF_NG
    End If
End Sub

' 変更があったセルを強調（オレンジ・太字）
Public Sub MarkChangedCell(ByVal c As Range)
    c.Interior.Color = DIFF_CELL
    c.Font.Bold = True
End Sub

Public Sub AddKeys(ByVal dst As Object, ByVal src As Object, ByVal objKind As String)
    Dim k As Variant
    For Each k In src.Keys
        If Left$(k, Len(objKind) + 1) = objKind & "|" Then dst(k) = True
    Next
End Sub


'----------------------------------------------------------
' 値・配列の補助
'----------------------------------------------------------


Public Function Nz(ByVal d As Object, ByVal key As String) As String
    If d.Exists(key) Then
        If Not IsObject(d(key)) Then
            If Not IsNull(d(key)) Then Nz = CStr(d(key))
        End If
    End If
End Function

Public Function JoinArr(ByVal item As Object, ByVal key As String) As String
    Dim t As Variant, r As String
    If Not item.Exists(key) Then Exit Function
    If Not IsObject(item(key)) Then Exit Function
    For Each t In item(key)
        If r <> "" Then r = r & ","
        r = r & t
    Next
    JoinArr = r
End Function

Public Function SplitMembers(ByVal members As String) As Variant
    Dim parts() As String, res() As String, i As Long, n As Long, m As String
    parts = Split(Replace(Replace(members, vbCr, ""), vbLf, ","), ",")
    n = -1
    If UBound(parts) >= 0 Then ReDim res(UBound(parts))
    For i = LBound(parts) To UBound(parts)
        m = Trim$(parts(i))
        If m <> "" Then
            n = n + 1
            res(n) = m
        End If
    Next
    If n < 0 Then
        SplitMembers = Split("")
    Else
        ReDim Preserve res(n)
        SplitMembers = res
    End If
End Function

Public Function IsMember(ByVal members As String, ByVal nm As String) As Boolean
    Dim parts As Variant, i As Long
    parts = SplitMembers(members)
    For i = 0 To UBound(parts)
        If CStr(parts(i)) = nm Then
            IsMember = True
            Exit Function
        End If
    Next
End Function

Public Function LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function Base64(ByVal txt As String) As String
    Dim bytes() As Byte, node As Object
    bytes = StrConv(txt, vbFromUnicode)
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64 = Replace(Replace(node.Text, vbLf, ""), vbCr, "")
End Function

'==========================================================
' JSON パーサ／エスケープ（旧 modJson。外部参照設定不要）
'   ParseJson(文字列) -> オブジェクト=Scripting.Dictionary
'                        配列      =Collection
'==========================================================
Public Function ParseJson(ByVal txt As String) As Object
    mJs = txt
    mJp = 1
    SkipWs
    Dim v As Variant
    ParseValue v
    If IsObject(v) Then Set ParseJson = v Else Set ParseJson = Nothing
End Function

' 文字列をJSON文字列リテラル("...")に変換
Public Function JsonStr(ByVal v As String) As String
    v = Replace(v, "\", "\\")
    v = Replace(v, """", "\""")
    v = Replace(v, vbCrLf, "\n")
    v = Replace(v, vbCr, "\n")
    v = Replace(v, vbLf, "\n")
    v = Replace(v, vbTab, "\t")
    JsonStr = """" & v & """"
End Function

' 一覧応答 {"data":[...],"total":..} を読み、data のうち定義済み（id なし／snippet・folder が predefined）の要素は
' オブジェクトを作らずに読み飛ばす（大量の定義済み App-ID を高速に処理するため）。rawN=読み飛ばしを含む要素数
Public Function ParseJsonSkipPredef(ByVal txt As String, ByRef rawN As Long) As Object
    Dim d As Object, k As String, v As Variant, c As Collection, st As Long, seg As String
    rawN = 0
    mJs = txt
    mJp = 1
    SkipWs
    If Mid$(mJs, mJp, 1) <> "{" Then
        Set ParseJsonSkipPredef = ParseJson(txt)
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    mJp = mJp + 1
    Do
        SkipWs
        If Mid$(mJs, mJp, 1) = "}" Or mJp > Len(mJs) Then Exit Do
        k = ParseString()
        SkipWs
        mJp = mJp + 1                       ' ":"
        SkipWs
        If k = "data" And Mid$(mJs, mJp, 1) = "[" Then
            Set c = New Collection
            mJp = mJp + 1
            Do
                SkipWs
                If Mid$(mJs, mJp, 1) = "]" Then
                    mJp = mJp + 1
                    Exit Do
                End If
                st = mJp
                SkipValue
                seg = Mid$(mJs, st, mJp - st)
                rawN = rawN + 1
                If Not LooksPredefined(seg) Then c.Add SubParse(seg)
                SkipWs
                If Mid$(mJs, mJp, 1) = "," Then mJp = mJp + 1
            Loop
            Set d(k) = c
        Else
            ParseValue v
            If IsObject(v) Then Set d(k) = v Else d(k) = v
        End If
        SkipWs
        If Mid$(mJs, mJp, 1) = "," Then mJp = mJp + 1
    Loop
    Set ParseJsonSkipPredef = d
End Function

Private Function LooksPredefined(ByVal seg As String) As Boolean
    Dim t As String
    t = LCase$(Replace(seg, " ", ""))
    If InStr(t, """id"":") = 0 Then LooksPredefined = True: Exit Function
    If InStr(t, """snippet"":""predefined") > 0 Then LooksPredefined = True: Exit Function
    If InStr(t, """folder"":""predefined") > 0 Then LooksPredefined = True
End Function

' 要素1つ分の文字列を解析（解析位置を保存・復元）
Private Function SubParse(ByVal seg As String) As Variant
    Dim s0 As String, p0 As Long, v As Variant
    s0 = mJs: p0 = mJp
    mJs = seg: mJp = 1
    ParseValue v
    mJs = s0: mJp = p0
    If IsObject(v) Then Set SubParse = v Else SubParse = v
End Function

' 値を1つ読み飛ばす（オブジェクトを作らない）
Private Sub SkipValue()
    Dim depth As Long, ch As String
    SkipWs
    ch = Mid$(mJs, mJp, 1)
    If ch = """" Then
        SkipString
        Exit Sub
    End If
    If ch <> "{" And ch <> "[" Then
        Do While mJp <= Len(mJs)
            If InStr(",}] " & vbTab & vbCr & vbLf, Mid$(mJs, mJp, 1)) > 0 Then Exit Do
            mJp = mJp + 1
        Loop
        Exit Sub
    End If
    Do While mJp <= Len(mJs)
        ch = Mid$(mJs, mJp, 1)
        Select Case ch
            Case """": SkipString
            Case "{", "["
                depth = depth + 1
                mJp = mJp + 1
            Case "}", "]"
                depth = depth - 1
                mJp = mJp + 1
                If depth = 0 Then Exit Do
            Case Else: mJp = mJp + 1
        End Select
    Loop
End Sub

' 文字列を読み飛ばす（\" を考慮）
Private Sub SkipString()
    Dim q As Long
    mJp = mJp + 1
    Do
        q = InStr(mJp, mJs, """")
        If q = 0 Then
            mJp = Len(mJs) + 1
            Exit Sub
        End If
        If Mid$(mJs, q - 1, 1) <> "\" Or BackslashEven(q - 1) Then
            mJp = q + 1
            Exit Sub
        End If
        mJp = q + 1
    Loop
End Sub

' 位置 i で終わる連続した \ の個数が偶数か（偶数ならその後の " はエスケープされていない）
Private Function BackslashEven(ByVal i As Long) As Boolean
    Dim n As Long
    Do While i >= 1
        If Mid$(mJs, i, 1) <> "\" Then Exit Do
        n = n + 1
        i = i - 1
    Loop
    BackslashEven = (n Mod 2 = 0)
End Function

Private Sub ParseValue(ByRef v As Variant)
    SkipWs
    Select Case Mid$(mJs, mJp, 1)
        Case "{": Set v = ParseObject()
        Case "[": Set v = ParseArray()
        Case """": v = ParseString()
        Case "t": mJp = mJp + 4: v = True
        Case "f": mJp = mJp + 5: v = False
        Case "n": mJp = mJp + 4: v = Null
        Case Else: v = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    mJp = mJp + 1
    SkipWs
    If Mid$(mJs, mJp, 1) = "}" Then
        mJp = mJp + 1
        Set ParseObject = d
        Exit Function
    End If
    Do
        SkipWs
        k = ParseString()
        SkipWs
        mJp = mJp + 1                       ' ":"
        ParseValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        SkipWs
        If Mid$(mJs, mJp, 1) = "," Then
            mJp = mJp + 1
        Else
            mJp = mJp + 1                   ' "}"
            Exit Do
        End If
    Loop
    Set ParseObject = d
End Function

Private Function ParseArray() As Object
    Dim c As New Collection, v As Variant
    mJp = mJp + 1
    SkipWs
    If Mid$(mJs, mJp, 1) = "]" Then
        mJp = mJp + 1
        Set ParseArray = c
        Exit Function
    End If
    Do
        ParseValue v
        c.Add v
        SkipWs
        If Mid$(mJs, mJp, 1) = "," Then
            mJp = mJp + 1
        Else
            mJp = mJp + 1                   ' "]"
            Exit Do
        End If
    Loop
    Set ParseArray = c
End Function

Private Function ParseString() As String
    ' 「"」か「\」までを InStr でまとめて切り出す（1文字ずつ連結するより大幅に速い）
    Dim r As String, q As Long, b As Long, ch As String
    mJp = mJp + 1                           ' 開始の "
    Do
        q = InStr(mJp, mJs, """")
        If q = 0 Then                   ' 閉じ " が無い（不正JSON）
            r = r & Mid$(mJs, mJp)
            mJp = Len(mJs) + 1
            Exit Do
        End If
        b = InStr(mJp, mJs, "\")
        If b = 0 Or b > q Then
            r = r & Mid$(mJs, mJp, q - mJp)
            mJp = q + 1
            Exit Do
        End If
        r = r & Mid$(mJs, mJp, b - mJp)
        ch = Mid$(mJs, b + 1, 1)
        mJp = b + 2
        Select Case ch
            Case "n": r = r & vbLf
            Case "r": r = r & vbCr
            Case "t": r = r & vbTab
            Case "b": r = r & Chr$(8)
            Case "f": r = r & Chr$(12)
            Case "u"
                r = r & ChrW$(CLng("&H" & Mid$(mJs, mJp, 4)))
                mJp = mJp + 4
            Case Else: r = r & ch       ' \" \ \/
        End Select
    Loop
    ParseString = r
End Function

Private Function ParseNumber() As Variant
    Dim st As Long
    st = mJp
    Do While mJp <= Len(mJs)
        If InStr("+-0123456789.eE", Mid$(mJs, mJp, 1)) = 0 Then Exit Do
        mJp = mJp + 1
    Loop
    ParseNumber = Val(Mid$(mJs, st, mJp - st))
End Function

Private Sub SkipWs()
    Do While mJp <= Len(mJs)
        Select Case Mid$(mJs, mJp, 1)
            Case " ", vbTab, vbCr, vbLf: mJp = mJp + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

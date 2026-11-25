Attribute VB_Name = "modScmInput"
'==========================================================
' modScmInput : SCM 設定 入力シート（お客様記入用）
'   入力とチェックが中心。API は既存情報の取得（GET）だけ（modInFetch）。JSON化・登録はしません。
'   必要なもの : modScmCommon / modScmInput（このモジュール） / modInPolicy / modInFetch / clsPickerEvents / frmMemberPick
'     frmMemberPick は空のユーザーフォームを挿入して frmMemberPick_コード.txt を貼り付け
'     （または modPickerSetup.bas をインポートして CreatePickerForm を実行）
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'==========================================================
Option Explicit

Private Const S1 As String = "記入状況"
Private Const S2 As String = "②既存データ抽出"
Private Const S3 As String = "③アドレス"
Private Const S4 As String = "④アドレスグループ"
Private Const S5 As String = "⑤サービス"
Private Const S6 As String = "⑥サービスグループ"
Private Const S7 As String = "⑦カスタムURLカテゴリ"
Private Const S8 As String = "⑧HIP一覧"
Private Const S8B As String = "⑧-2HIPプロファイル"
Private Const S9 As String = "⑨EDL"
Private Const S10 As String = "⑩アプリケーション"
Private Const S11 As String = "⑪ユーザ・グループ"
Private Const R1 As String = "参照_HIP選択肢"
Private Const R2 As String = "参照_HIPベンダー一覧"
Private Const SH_RES As String = "チェック結果"

Private Const IN_FIRST As Long = 8      ' ③～⑪ データ開始行
Private Const IN_LAST As Long = 2007    ' ③～⑪ データ最終行（2,000行。2026-11-19 に 500行から拡大）
Private Const EX_FIRST As Long = 8      ' ② データ開始行

' ② 既存データ抽出：ブロックの開始位置（列のずれ。ブロック内の列番号 + OFF_xxx = シートの列番号）
Private Const OFF_ADR As Long = 0
Private Const OFF_ADRG As Long = 14
Private Const OFF_SVC As Long = 26
Private Const OFF_SVCG As Long = 40
Private Const OFF_URL As Long = 47
Private Const OFF_URLV As Long = 54
Private Const OFF_HIP As Long = 58
Private Const OFF_HIPP As Long = 124
Private Const OFF_EDL As Long = 133
Private Const OFF_PG As Long = 153         ' セキュリティプロファイルグループ（B=id C=name）
Private Const OFF_PROF As Long = 166       ' セキュリティプロファイル（B=種類 C=name D=備考）
Private Const EX_LAST As Long = 5000
Private Const EXP_LAST As Long = 2000      ' HIPプロファイル
Private Const URLV_LAST As Long = 30007    ' URL（1行1件）

' ⑦ カスタムURLカテゴリ（横長：B列から右へ1列1カテゴリ、最大1,500列）
Private Const U_C1 As Long = 2
Private Const U_C2 As Long = 1501      ' 1,500カテゴリ（2026-11-19 に 50列から拡大）
Private Const U_OP As Long = 6
Private Const U_NAME As Long = 7
Private Const U_TYPE As Long = 8
Private Const U_DESC As Long = 9
Private Const U_VBA As Long = 14
Private Const U_FIRST As Long = 16
Private Const U_LAST As Long = 3015
Private Const URL_MAX As Long = 2500

' ⑧ HIPオブジェクト：F列(6)～BM列(65) が description と条件の各項目（②のHIPブロックも同じ並び）
Private Const F_FIRST As Long = 6
Private Const F_LAST As Long = 65
Private Const VD_FIRST As Long = 6      ' 参照_HIPベンダー一覧 データ開始行（B=カテゴリ C=ベンダー D=製品 E=出所）
Private Const MODE_AND As String = "すべて満たす(and)"
Private Const MODE_OR As String = "いずれか(or)"
Private Const MODE_RAW As String = "直接記述"

' 色（VBAで直接塗る）
Private Const CL_NG As Long = &HCEC7FF      ' FFC7CE
Private Const CL_WARN As Long = &H9CEBFF    ' FFEB9C

' ---- メンバー選択画面 用 ----
Private Const PICK_KEY As String = "^+m"            ' Ctrl+Shift+M
Private Const CELL_MAX As Long = 32767
Private Const SH_OVERFLOW As String = "ピッカー_選択結果"

Private mEvents As clsPickerEvents
Private mNames As Object     ' 候補 name -> 種別ラベル（追加順を保持）
Private mExclude As String   ' 候補から外す name（自分自身）
Private mVendor As Boolean   ' ベンダー欄のピッカー

' ---- 一括チェック 用 ----
Private mRes As Collection   ' Array(シート, 行, 列, 名前, 判定, 内容)

'==========================================================
' ボタン配置（初回のみ実行）
'==========================================================
Public Sub SetupButtons()
    AddBtn S1, "B20", "一括チェック", "CheckAll"
    ' 旧版のボタン（既存ツールから取込・右側の API 取得）が残っていれば消す
    On Error Resume Next
    Worksheets(S2).Buttons("btn_ImportFromTools").Delete
    Worksheets(S2).Buttons("btn_FetchExistingAll_L3N3").Delete
    On Error GoTo 0
    AddBtn S2, "H3:J3", "既存データをクリア", "ClearExisting"
    AddBtn S3, "F3:G4", "変更行に既存値を読込", "LoadExistingAdr"
    AddBtn S4, "F3:G4", "変更行に既存値を読込", "LoadExistingAdrGroup"
    AddBtn S5, "F3:G4", "変更行に既存値を読込", "LoadExistingSvc"
    AddBtn S6, "F3:G4", "変更行に既存値を読込", "LoadExistingSvcGroup"
    AddBtn S7, "F3:G4", "既存内容を読込（変更の列）", "LoadExistingUrl"
    AddBtn S7, "I3:J4", "一括チェック", "CheckAll"
    AddBtn S8, "F3:G4", "変更行に既存値を読込", "LoadExistingHip"
    AddBtn S8B, "F3:G4", "変更行に既存値を読込", "LoadExistingHipProfile"
    AddBtn S9, "F3:G4", "変更行に既存値を読込", "LoadExistingEdl"
    modInPolicy.SetupPolicyButtons
    modInFetch.SetupFetchButtons
    MsgBox "ボタンを配置しました。ブックを .xlsm で保存してください。", vbInformation
End Sub

Private Sub AddBtn(ByVal sh As String, ByVal addr As String, ByVal caption As String, ByVal macro As String)
    Dim ws As Worksheet, rng As Range, b As Object
    Set ws = Worksheets(sh)
    ' 結合セル1つを指定したときは結合範囲全体に置く（MergeArea は1セルの範囲にしか使えないため、複数セルの指定はそのまま）
    Set rng = ws.Range(addr)
    If rng.Cells.Count = 1 Then Set rng = rng.MergeArea
    On Error Resume Next
    ws.Buttons("btn_" & macro).Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    b.Name = "btn_" & macro
    b.Caption = caption
    b.OnAction = macro
End Sub

' セキュリティプロファイル：プロファイルグループの各列（E～K）に入っている名前を種類ごとに1回ずつ書き出す
'   （既存ツールに個別プロファイルの一覧は無いため。グループに入っていないプロファイルは手で追記）
Public Sub ImportProfiles(ByRef msg As String)
    Dim ws2 As Worksheet, kinds As Variant, r As Long, c As Long, o As Long, nm As String, seen As Object, lastR As Long
    kinds = Array("ウイルス・WildFire", "スパイウェア", "脆弱性防御", "URLフィルタリング", "ファイルブロック", "DNSセキュリティ", "SaaSセキュリティ")
    Set ws2 = Worksheets(S2)
    Set seen = CreateObject("Scripting.Dictionary")
    ws2.Range(ws2.Cells(EX_FIRST, OFF_PROF + 2), ws2.Cells(EX_LAST, OFF_PROF + 4)).ClearContents
    lastR = LastRow(ws2, OFF_PG + 3)
    o = EX_FIRST
    For c = 0 To 6
        For r = EX_FIRST To lastR
            nm = Trim$(SafeStr(ws2.Cells(r, OFF_PG + 6 + c).Value))
            If nm <> "" And Not seen.Exists(kinds(c) & "|" & nm) And o <= EX_LAST Then
                seen.Add kinds(c) & "|" & nm, True
                ws2.Cells(o, OFF_PROF + 2).Value = kinds(c)
                ws2.Cells(o, OFF_PROF + 3).NumberFormat = "@"
                ws2.Cells(o, OFF_PROF + 3).Value = nm
                ws2.Cells(o, OFF_PROF + 4).Value = "プロファイルグループから抽出"
                o = o + 1
            End If
        Next
    Next
    msg = msg & vbLf & "　セキュリティプロファイル（グループから抽出）：" & (o - EX_FIRST) & " 件"
End Sub

Private Function OneCell(ByVal v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    a(1, 1) = v
    OneCell = a
End Function

' ② のデータ（関数の列以外）を全部消す
Public Sub ClearExisting()
    Dim ws2 As Worksheet
    If MsgBox("②既存データ抽出のデータをすべて消します。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    Set ws2 = Worksheets(S2)
    Application.ScreenUpdating = False
    ClearCols ws2, OFF_ADR, "2-10,13", EX_LAST
    ClearCols ws2, OFF_ADRG, "2-8,11", EX_LAST
    ClearCols ws2, OFF_SVC, "2-12", EX_LAST
    ClearCols ws2, OFF_SVCG, "2-6", EX_LAST
    ClearCols ws2, OFF_URL, "2-4,6", EX_LAST
    ClearCols ws2, OFF_URLV, "2-3", URLV_LAST
    ClearCols ws2, OFF_HIP, "2-65", EX_LAST
    ClearCols ws2, OFF_HIPP, "2-7", EXP_LAST
    ClearCols ws2, OFF_EDL, "2-15,17", EX_LAST
    ClearCols ws2, OFF_PG, "2-12", EXP_LAST
    ClearCols ws2, OFF_PROF, "2-4", EX_LAST
    ws2.Range("B3").ClearContents
    Application.ScreenUpdating = True
    MsgBox "消しました。", vbInformation
End Sub

Private Sub ClearCols(ByVal ws As Worksheet, ByVal off As Long, ByVal cols As String, ByVal lastR As Long)
    Dim part As Variant, a As Long, b As Long
    For Each part In Split(cols, ",")
        a = CLng(Split(part, "-")(0))
        If InStr(part, "-") > 0 Then b = CLng(Split(part, "-")(1)) Else b = a
        ws.Range(ws.Cells(EX_FIRST, off + a), ws.Cells(lastR, off + b)).ClearContents
    Next
End Sub

'==========================================================
' 変更行に既存値を読込（操作=変更 の行に、②の現在の値を入れる）
'   map : "入力列:②ブロック内の列,..."
'==========================================================
Public Sub LoadExistingAdr()
    LoadRows S3, OFF_ADR, EX_LAST, "4:11,5:12,6:8,7:9"
End Sub

Public Sub LoadExistingAdrGroup()
    LoadRows S4, OFF_ADRG, EX_LAST, "4:9,5:10,6:6,7:7"
End Sub

Public Sub LoadExistingSvc()
    LoadRows S5, OFF_SVC, EX_LAST, "4:4,5:5,6:6,7:7,8:8,9:9,10:10,11:11"
End Sub

Public Sub LoadExistingSvcGroup()
    LoadRows S6, OFF_SVCG, EX_LAST, "4:4,5:5"
End Sub

Public Sub LoadExistingEdl()
    LoadRows S9, OFF_EDL, EX_LAST, "4:4,5:5,6:6,7:7,8:8,9:9,10:10,11:11,12:12,13:13,14:14"
End Sub

Public Sub LoadExistingHip()
    Dim s As String, c As Long
    For c = F_FIRST To F_LAST
        s = s & IIf(s = "", "", ",") & c & ":" & c
    Next
    LoadRows S8, OFF_HIP, EX_LAST, s
End Sub

Private Sub LoadRows(ByVal sh As String, ByVal off As Long, ByVal lastEx As Long, ByVal map As String)
    Dim ws As Worksheet, tgt As Collection, missing As String, filled As Long, mode As Long
    Dim it As Variant, r As Long, eRow As Long, m As Variant, ci As Long, ce As Long, n As Long, c1 As Long, c2 As Long
    Set ws = Worksheets(sh)
    c1 = 9999: c2 = 0
    For Each m In Split(map, ",")
        ci = CLng(Split(m, ":")(0))
        If ci < c1 Then c1 = ci
        If ci > c2 Then c2 = ci
    Next
    Set tgt = ChangeRows(ws, off, lastEx, c1, c2, missing, filled)
    If tgt Is Nothing Then Exit Sub
    mode = AskLoadMode(filled)
    If mode = 0 Then Exit Sub
    Application.ScreenUpdating = False
    For Each it In tgt
        r = it(0): eRow = it(1)
        For Each m In Split(map, ",")
            ci = CLng(Split(m, ":")(0))
            ce = CLng(Split(m, ":")(1)) + off
            If mode = 1 Or SafeStr(ws.Cells(r, ci).Value) = "" Then
                ws.Cells(r, ci).NumberFormat = "@"
                ws.Cells(r, ci).Value = SafeStr(Worksheets(S2).Cells(eRow, ce).Value)
            End If
        Next
        n = n + 1
    Next
    Application.ScreenUpdating = True
    LoadDoneMsg n, missing, ""
End Sub

Public Sub LoadExistingHipProfile()
    Dim ws As Worksheet, ws2 As Worksheet, tgt As Collection, missing As String, filled As Long, mode As Long
    Dim it As Variant, r As Long, eRow As Long, n As Long, m As String, pos As Collection, ng As Collection, md As String
    Set ws = Worksheets(S8B)
    Set ws2 = Worksheets(S2)
    Set tgt = ChangeRows(ws, OFF_HIPP, EXP_LAST, 4, 8, missing, filled)
    If tgt Is Nothing Then Exit Sub
    mode = AskLoadMode(filled)
    If mode = 0 Then Exit Sub
    For Each it In tgt
        r = it(0): eRow = it(1)
        If mode = 1 Then ws.Range(ws.Cells(r, 4), ws.Cells(r, 8)).ClearContents
        If SafeStr(ws.Cells(r, 4).Value) = "" Then ws.Cells(r, 4).Value = SafeStr(ws2.Cells(eRow, OFF_HIPP + 4).Value)
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 5), ws.Cells(r, 8))) = 0 Then
            m = SafeStr(ws2.Cells(eRow, OFF_HIPP + 5).Value)
            If ParseSimpleMatch(m, md, pos, ng) Then
                ws.Cells(r, 5).Value = md
                ws.Cells(r, 6).Value = JoinColl(pos, ",")
                ws.Cells(r, 7).Value = JoinColl(ng, ",")
            Else
                ws.Cells(r, 5).Value = MODE_RAW
                ws.Cells(r, 8).Value = m
            End If
        End If
        n = n + 1
    Next
    LoadDoneMsg n, missing, "（match は、単純な and / or / not の式は選択欄に、それ以外は「直接記述」に展開）"
End Sub

' 操作=変更 の行と、②での行番号の組 Array(入力行, ②の行) の一覧
Private Function ChangeRows(ByVal ws As Worksheet, ByVal off As Long, ByVal lastEx As Long, ByVal c1 As Long, ByVal c2 As Long, _
                            ByRef missing As String, ByRef filled As Long) As Collection
    Dim res As New Collection, idx As Object, r As Long, nm As String, total As Long, ws2 As Worksheet, v As Variant, lastR As Long
    Set ws2 = Worksheets(S2)
    Set idx = CreateObject("Scripting.Dictionary")
    lastR = LastRow(ws2, off + 3)
    If lastR > lastEx Then lastR = lastEx
    If lastR >= EX_FIRST Then
        v = ws2.Range(ws2.Cells(EX_FIRST, off + 3), ws2.Cells(lastR, off + 3)).Value
        If Not IsArray(v) Then v = OneCell(v)
        For r = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(r, 1)))
            If nm <> "" And Not idx.Exists(nm) Then idx.Add nm, r + EX_FIRST - 1
        Next
    End If
    For r = IN_FIRST To IN_LAST
        nm = Trim$(SafeStr(ws.Cells(r, 3).Value))
        If Trim$(SafeStr(ws.Cells(r, 2).Value)) = "変更" And nm <> "" Then
            total = total + 1
            If idx.Exists(nm) Then
                res.Add Array(r, idx(nm))
                If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))) > 0 Then filled = filled + 1
            Else
                missing = missing & vbLf & "・" & r & "行目 " & nm
            End If
        End If
    Next
    If total = 0 Then
        MsgBox "操作（B列）が「変更」で name（C列）が入っている行がありません。", vbExclamation
        Exit Function
    End If
    If res.Count = 0 Then
        LoadDoneMsg 0, missing, ""
        Exit Function
    End If
    Set ChangeRows = res
End Function

' 入力済みの行があるときの読み込み方。1=既存値で上書き 2=空欄の欄だけ埋める 0=中止
Private Function AskLoadMode(ByVal filled As Long) As Long
    If filled = 0 Then AskLoadMode = 1: Exit Function
    Select Case MsgBox("変更の行のうち " & filled & " 行は、すでに入力があります。" & vbLf & vbLf & _
                       "はい　　：既存の値で上書きする（入力した値は消えます）" & vbLf & _
                       "いいえ　：空欄の欄だけ既存の値で埋める（入力した値は残します）" & vbLf & _
                       "キャンセル：読み込まない", vbYesNoCancel + vbQuestion, "変更行に既存値を読込")
        Case vbYes: AskLoadMode = 1
        Case vbNo: AskLoadMode = 2
        Case Else: AskLoadMode = 0
    End Select
End Function

Private Sub LoadDoneMsg(ByVal n As Long, ByVal missing As String, ByVal note As String)
    Dim s As String
    s = n & " 行に既存値を読み込みました。" & note
    If missing <> "" Then
        s = s & vbLf & vbLf & "次の name は②既存データに無いため読み込めませんでした。" & missing
    End If
    MsgBox s, IIf(missing <> "", vbExclamation, vbInformation)
End Sub

' match が ⑧-2 の組み方（and / or と not）で表せるなら分解する
Private Function ParseSimpleMatch(ByVal m As String, ByRef mode As String, ByRef pos As Collection, ByRef ng As Collection) As Boolean
    Dim i As Long, q1 As Long, q2 As Long, before As String, nm As String
    Set pos = New Collection
    Set ng = New Collection
    m = Trim$(m)
    i = 1
    Do
        q1 = InStr(i, m, """")
        If q1 = 0 Then Exit Do
        q2 = InStr(q1 + 1, m, """")
        If q2 = 0 Then Exit Function
        nm = Mid$(m, q1 + 1, q2 - q1 - 1)
        before = LCase$(Trim$(Mid$(m, i, q1 - i)))
        If Right$(before, 3) = "not" Then ng.Add nm Else pos.Add nm
        i = q2 + 1
    Loop
    If pos.Count + ng.Count = 0 Then Exit Function
    mode = IIf(InStr(LCase$(m), " or ") > 0, MODE_OR, MODE_AND)
    ParseSimpleMatch = (GenMatch(mode, pos, ng) = m)
End Function

' ⑧-2 I列の関数と同じ規則で match を作る
Private Function GenMatch(ByVal mode As String, ByVal pos As Collection, ByVal ng As Collection) As String
    Dim op As String, p As String, n As String, v As Variant
    op = IIf(mode = MODE_OR, "or", "and")
    For Each v In pos
        If p <> "" Then p = p & " " & op & " "
        p = p & """" & v & """"
    Next
    If pos.Count > 1 And op = "or" And ng.Count > 0 Then p = "(" & p & ")"
    For Each v In ng
        If n <> "" Then n = n & " and "
        n = n & "not """ & v & """"
    Next
    If p <> "" And n <> "" Then GenMatch = p & " and " & n Else GenMatch = p & n
End Function

'==========================================================
' ⑦ 既存内容を読込（操作=変更 の列に、②の現在のURLを入れる）
'==========================================================
Public Sub LoadExistingUrl()
    Dim ws As Worksheet, ws2 As Worksheet, c As Long, nm As String, have As Object, lastR As Long, v As Variant, r As Long
    Dim meta As Object, n As Long, filled As Long, mode As Long, col As Collection, outA() As Variant, k As Long, cols As Collection
    Dim missing As String, it As Variant
    Set ws = Worksheets(S7)
    Set ws2 = Worksheets(S2)
    Set have = CreateObject("Scripting.Dictionary")
    Set meta = CreateObject("Scripting.Dictionary")
    lastR = LastRow(ws2, OFF_URLV + 2)
    If lastR >= EX_FIRST Then
        v = ws2.Range(ws2.Cells(EX_FIRST, OFF_URLV + 2), ws2.Cells(lastR, OFF_URLV + 3)).Value
        For r = 1 To UBound(v, 1)
            nm = SafeStr(v(r, 1))
            If nm <> "" Then
                If Not have.Exists(nm) Then have.Add nm, New Collection
                If SafeStr(v(r, 2)) <> "" Then have(nm).Add SafeStr(v(r, 2))
            End If
        Next
    End If
    lastR = LastRow(ws2, OFF_URL + 2)
    For r = EX_FIRST To lastR
        nm = SafeStr(ws2.Cells(r, OFF_URL + 2).Value)
        If nm <> "" And Not meta.Exists(nm) Then meta.Add nm, Array(SafeStr(ws2.Cells(r, OFF_URL + 3).Value), SafeStr(ws2.Cells(r, OFF_URL + 4).Value))
    Next
    Set cols = New Collection
    For c = U_C1 To U_C2
        nm = Trim$(SafeStr(ws.Cells(U_NAME, c).Value))
        If SafeStr(ws.Cells(U_OP, c).Value) = "変更" And nm <> "" Then
            If meta.Exists(nm) Or have.Exists(nm) Then
                cols.Add c
                If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(U_FIRST, c), ws.Cells(U_LAST, c))) > 0 Then filled = filled + 1
            Else
                missing = missing & vbLf & "・" & ws.Cells(U_NAME, c).Address(False, False) & " " & nm
            End If
        End If
    Next
    If cols.Count = 0 Then
        MsgBox "操作（6行目）が「変更」で、②既存データにあるカテゴリ名（7行目）の列がありません。" & missing, vbExclamation
        Exit Sub
    End If
    mode = AskLoadMode(filled)
    If mode = 0 Then Exit Sub
    Application.ScreenUpdating = False
    For Each it In cols
        c = it
        nm = Trim$(SafeStr(ws.Cells(U_NAME, c).Value))
        If meta.Exists(nm) Then
            If SafeStr(ws.Cells(U_TYPE, c).Value) = "" Then ws.Cells(U_TYPE, c).Value = meta(nm)(0)
            If SafeStr(ws.Cells(U_DESC, c).Value) = "" Then ws.Cells(U_DESC, c).Value = meta(nm)(1)
        End If
        If mode = 1 Or Application.WorksheetFunction.CountA(ws.Range(ws.Cells(U_FIRST, c), ws.Cells(U_LAST, c))) = 0 Then
            ws.Range(ws.Cells(U_FIRST, c), ws.Cells(U_LAST, c)).ClearContents
            If have.Exists(nm) Then
                Set col = have(nm)
                If col.Count > 0 Then
                    ReDim outA(1 To col.Count, 1 To 1)
                    For k = 1 To col.Count
                        outA(k, 1) = col(k)
                    Next
                    ws.Range(ws.Cells(U_FIRST, c), ws.Cells(U_FIRST + col.Count - 1, c)).NumberFormat = "@"
                    ws.Range(ws.Cells(U_FIRST, c), ws.Cells(U_FIRST + col.Count - 1, c)).Value = outA
                End If
            End If
        End If
        n = n + 1
    Next
    Application.ScreenUpdating = True
    MsgBox n & " 列に既存のURLを読み込みました。" & IIf(missing <> "", vbLf & vbLf & "②既存データに無いカテゴリ：" & missing, ""), vbInformation
End Sub

'==========================================================
' 一括チェック
'   1) 関数の判定（NG・警告）を全シートから集める
'   2) 関数では見られない項目を VBA でチェックする
'      ⑦ URLの書式・重複（セルを赤/黄に塗り、14行目に結果）、④ メンバーの重複、
'      ⑧-2 直接記述の参照先、⑨ exception_list の長さ・重複、⑪ 所属グループの重複
'   3) 結果を「チェック結果」シートに一覧にする
'==========================================================
Public Sub CheckAll()
    Dim ng As Long, wn As Long, it As Variant
    Set mRes = New Collection
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    CheckUrlSheet
    Application.Calculate
    CheckMemberDup S4, 5, 4, "static", "メンバー"
    CheckMemberDup S11, 7, 3, "ユーザ", "所属グループ"
    CheckHipRaw
    CheckEdlException
    CollectJudge S3, 16, 17
    CollectJudge S4, 16, 17
    CollectJudge S5, 24, 25
    CollectJudge S6, 12, 13
    CollectJudgeUrl
    CollectJudge S8, 4, 5
    CollectJudge S8B, 16, 17
    CollectJudge S9, 31, 32
    CollectJudge S10, 8, 9
    CollectJudge S11, 10, 11
    modInPolicy.CheckPolicies
    For Each it In modInPolicy.PolicyResults()
        mRes.Add it
    Next
    For Each it In mRes
        If it(4) = "NG" Then ng = ng + 1 Else wn = wn + 1
    Next
    WriteResult
    Application.ScreenUpdating = True
    If ng + wn = 0 Then
        MsgBox "チェックが終わりました。NG・警告はありません。", vbInformation
    Else
        MsgBox "チェックが終わりました。" & vbLf & "NG：" & ng & " 件　警告：" & wn & " 件" & vbLf & _
               "内容は「" & SH_RES & "」シートを見てください（セルのリンクで該当の行へ移動できます）。", _
               IIf(ng > 0, vbExclamation, vbInformation)
        Worksheets(SH_RES).Activate
    End If
End Sub

Private Sub AddRes(ByVal sh As String, ByVal r As Long, ByVal c As Long, ByVal nm As String, ByVal lvl As String, ByVal msg As String)
    mRes.Add Array(sh, r, c, nm, lvl, msg)
End Sub

' 関数の判定列（NG / 警告）を集める
Private Sub CollectJudge(ByVal sh As String, ByVal jc As Long, ByVal rc As Long)
    Dim ws As Worksheet, j As Variant, rs As Variant, nmv As Variant, r As Long, nc As Long, s As String
    Set ws = Worksheets(sh)
    nc = IIf(sh = S11, 4, 3)
    j = ws.Range(ws.Cells(IN_FIRST, jc), ws.Cells(IN_LAST, jc)).Value
    rs = ws.Range(ws.Cells(IN_FIRST, rc), ws.Cells(IN_LAST, rc)).Value
    nmv = ws.Range(ws.Cells(IN_FIRST, nc), ws.Cells(IN_LAST, nc)).Value
    For r = 1 To UBound(j, 1)
        s = SafeStr(j(r, 1))
        If s = "NG" Or s = "警告" Then
            AddRes sh, r + IN_FIRST - 1, jc, SafeStr(nmv(r, 1)), s, Mid$(SafeStr(rs(r, 1)), InStr(SafeStr(rs(r, 1)), ":") + 1)
        End If
    Next
End Sub

Private Sub CollectJudgeUrl()
    Dim ws As Worksheet, c As Long, s As String, rs As String
    Set ws = Worksheets(S7)
    For c = U_C1 To U_C2
        s = SafeStr(ws.Cells(12, c).Value)
        If s = "NG" Or s = "警告" Then
            rs = SafeStr(ws.Cells(13, c).Value)
            AddRes S7, 12, c, SafeStr(ws.Cells(U_NAME, c).Value), s, Mid$(rs, InStr(rs, ":") + 1)
        End If
    Next
End Sub

'----------------------------------------------------------
' ⑦ URL の書式・重複チェック（Palo Alto の URL リストの書き方）
'----------------------------------------------------------
Private Sub CheckUrlSheet()
    Dim ws As Worksheet, ws2 As Worksheet, c As Long, r As Long, v As Variant, nm As String, op As String, isCm As Boolean
    Dim seen As Object, seenL As Object, e As String, msg As String, nNg As Long, nWn As Long, oldE As Object
    Dim have As Object, lastR As Long, ev As Variant, k As String, nAdd As Long, nDel As Long, kk As Variant, cnt As Long
    Dim ngCells As String, wnCells As String, firstNg As String
    Set ws = Worksheets(S7)
    Set ws2 = Worksheets(S2)
    ' ② の既存URL（カテゴリ名 → Dictionary(URL)）
    Set have = CreateObject("Scripting.Dictionary")
    lastR = LastRow(ws2, OFF_URLV + 2)
    If lastR >= EX_FIRST Then
        ev = ws2.Range(ws2.Cells(EX_FIRST, OFF_URLV + 2), ws2.Cells(lastR, OFF_URLV + 3)).Value
        For r = 1 To UBound(ev, 1)
            k = SafeStr(ev(r, 1))
            If k <> "" Then
                If Not have.Exists(k) Then have.Add k, CreateObject("Scripting.Dictionary")
                If SafeStr(ev(r, 2)) <> "" Then have(k)(SafeStr(ev(r, 2))) = True
            End If
        Next
    End If
    ws.Range(ws.Cells(U_FIRST, U_C1), ws.Cells(U_LAST, U_C2)).Interior.Pattern = xlNone
    ws.Range(ws.Cells(U_VBA, U_C1), ws.Cells(U_VBA, U_C2)).ClearContents
    ' 記入のある最後の列までを1列ずつ読む（1,500列×3,000行を一度に読むとメモリが足りなくなるため）
    For c = U_C1 To UrlLastCol(ws)
        v = ws.Range(ws.Cells(U_FIRST, c), ws.Cells(U_LAST, c)).Value
        nm = Trim$(SafeStr(ws.Cells(U_NAME, c).Value))
        op = SafeStr(ws.Cells(U_OP, c).Value)
        isCm = (SafeStr(ws.Cells(U_TYPE, c).Value) = "Category Match")
        Set seen = CreateObject("Scripting.Dictionary")
        Set seenL = CreateObject("Scripting.Dictionary")
        nNg = 0: nWn = 0: nAdd = 0: nDel = 0: cnt = 0
        ngCells = "": wnCells = "": firstNg = ""
        If op <> "削除" Then
            For r = 1 To UBound(v, 1)
                e = Trim$(SafeStr(v(r, 1)))
                If e <> "" Then
                    cnt = cnt + 1
                    msg = EntryCheck(e, isCm)
                    If Left$(msg, 2) <> "NG" Then
                        If seen.Exists(e) Then
                            msg = "NG:同じカテゴリに同じURLがあります（重複）"
                        ElseIf seenL.Exists(LCase$(e)) And msg = "" Then
                            msg = "警告:大文字・小文字だけ異なるURLがあります"
                        End If
                        seen(e) = True
                        seenL(LCase$(e)) = True
                    End If
                    If msg <> "" Then
                        If Left$(msg, 2) = "NG" Then
                            nNg = nNg + 1
                            ws.Cells(U_FIRST + r - 1, c).Interior.Color = CL_NG
                            If firstNg = "" Then firstNg = Mid$(msg, 4)
                        Else
                            nWn = nWn + 1
                            ws.Cells(U_FIRST + r - 1, c).Interior.Color = CL_WARN
                        End If
                        AddRes S7, U_FIRST + r - 1, c, nm & " : " & e, Left$(msg, InStr(msg, ":") - 1), Mid$(msg, InStr(msg, ":") + 1)
                    End If
                    If op = "変更" And have.Exists(nm) Then
                        If Not have(nm).Exists(e) Then nAdd = nAdd + 1
                    End If
                End If
            Next
            If op = "変更" And have.Exists(nm) Then
                For Each kk In have(nm).Keys
                    If Not seen.Exists(CStr(kk)) Then nDel = nDel + 1
                Next
            End If
            If nm <> "" Or cnt > 0 Then
                If nNg > 0 Then
                    ws.Cells(U_VBA, c).Value = "NG:URLの書式NG " & nNg & " 件（赤のセル）。例：" & firstNg
                ElseIf nWn > 0 Then
                    ws.Cells(U_VBA, c).Value = "警告:確認が必要なURL " & nWn & " 件（黄のセル）"
                ElseIf op = "変更" Then
                    ws.Cells(U_VBA, c).Value = "OK（追加 " & nAdd & " 件・削除 " & nDel & " 件）"
                Else
                    ws.Cells(U_VBA, c).Value = "OK"
                End If
            End If
        ElseIf nm <> "" Then
            ws.Cells(U_VBA, c).Value = "OK（カテゴリを削除）"
        End If
    Next
End Sub

' ⑦で記入のある最後の列（6～9行目・URL欄のどこかに値がある列）。無ければ U_C1 - 1
Private Function UrlLastCol(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long, f As Range
    UrlLastCol = U_C1 - 1
    For r = U_OP To U_DESC
        c = ws.Cells(r, U_C2).End(xlToLeft).Column
        If Not IsEmpty(ws.Cells(r, U_C2).Value) Then c = U_C2
        If c > UrlLastCol Then UrlLastCol = c
    Next
    Set f = ws.Range(ws.Cells(U_FIRST, U_C1), ws.Cells(U_LAST, U_C2)).Find(What:="*", LookIn:=xlValues, _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then
        If f.Column > UrlLastCol Then UrlLastCol = f.Column
    End If
    If UrlLastCol > U_C2 Then UrlLastCol = U_C2
    If UrlLastCol < U_C1 - 1 Then UrlLastCol = U_C1 - 1
End Function

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
' メンバー（カンマ・改行区切り）の重複
'   col=メンバー列、condCol/condVal=対象行の条件（例 D列=static）
'----------------------------------------------------------
Private Sub CheckMemberDup(ByVal sh As String, ByVal col As Long, ByVal condCol As Long, ByVal condVal As String, ByVal label As String)
    Dim ws As Worksheet, r As Long, items As Collection, seen As Object, it As Variant, dup As String, nmc As Long
    Set ws = Worksheets(sh)
    nmc = IIf(sh = S11, 4, 3)
    For r = IN_FIRST To IN_LAST
        If SafeStr(ws.Cells(r, col).Value) <> "" And SafeStr(ws.Cells(r, condCol).Value) = condVal _
           And SafeStr(ws.Cells(r, 2).Value) <> "削除" Then
            Set items = SplitItems(SafeStr(ws.Cells(r, col).Value))
            Set seen = CreateObject("Scripting.Dictionary")
            seen.CompareMode = IIf(sh = S11, vbTextCompare, vbBinaryCompare)
            dup = ""
            For Each it In items
                If seen.Exists(it) Then
                    If InStr(dup & ",", "," & it & ",") = 0 Then dup = dup & "," & it
                Else
                    seen.Add it, True
                End If
            Next
            If dup <> "" Then
                AddRes sh, r, col, SafeStr(ws.Cells(r, nmc).Value), IIf(sh = S11, "警告", "NG"), _
                       label & "に同じ名前が重複しています（" & Mid$(dup, 2) & "）"
            End If
        End If
    Next
End Sub

'----------------------------------------------------------
' ⑧-2 直接記述の match で使っている HIPオブジェクトが ② か ⑧(登録) にあるか
'----------------------------------------------------------
Private Sub CheckHipRaw()
    Dim ws As Worksheet, ok As Object, del As Object, r As Long, nm As Variant, ng As String, dl As String, m As String
    Set ws = Worksheets(S8B)
    Set ok = NameSet(S2, OFF_HIP + 3, 0, "", EX_LAST)
    AddNames ok, S8, 3, 2, "登録"
    Set del = CreateObject("Scripting.Dictionary")
    AddNames del, S8, 3, 2, "削除"
    For r = IN_FIRST To IN_LAST
        If SafeStr(ws.Cells(r, 5).Value) = MODE_RAW And SafeStr(ws.Cells(r, 2).Value) <> "削除" Then
            m = SafeStr(ws.Cells(r, 8).Value)
            ng = "": dl = ""
            For Each nm In QuotedNames(m)
                If del.Exists(nm) Then
                    dl = dl & "," & nm
                ElseIf Not ok.Exists(nm) Then
                    ng = ng & "," & nm
                End If
            Next
            If m <> "" And QuotedNames(m).Count = 0 Then
                AddRes S8B, r, 8, SafeStr(ws.Cells(r, 3).Value), "NG", "matchにHIPオブジェクト名（""名前""）がありません"
            End If
            If ng <> "" Then AddRes S8B, r, 8, SafeStr(ws.Cells(r, 3).Value), "NG", "②既存データにも⑧の登録予定にも無いHIPオブジェクト：" & Mid$(ng, 2)
            If dl <> "" Then AddRes S8B, r, 8, SafeStr(ws.Cells(r, 3).Value), "NG", "⑧で削除予定のHIPオブジェクトを使っています：" & Mid$(dl, 2)
        End If
    Next
End Sub

'----------------------------------------------------------
' ⑨ exception_list（K列）の各値：長さ（255、imsi=34、imei=32）と重複
'----------------------------------------------------------
Private Sub CheckEdlException()
    Dim ws As Worksheet, r As Long, items As Collection, it As Variant, seen As Object, mx As Long, typ As String
    Dim lng As String, dup As String
    Set ws = Worksheets(S9)
    For r = IN_FIRST To IN_LAST
        If SafeStr(ws.Cells(r, 11).Value) <> "" And SafeStr(ws.Cells(r, 2).Value) <> "削除" Then
            typ = SafeStr(ws.Cells(r, 4).Value)
            mx = 255
            If typ = "imsi" Then mx = 34
            If typ = "imei" Then mx = 32
            Set items = SplitItems(SafeStr(ws.Cells(r, 11).Value))
            Set seen = CreateObject("Scripting.Dictionary")
            lng = "": dup = ""
            For Each it In items
                If Len(it) > mx Then lng = lng & "," & Left$(it, 30)
                If seen.Exists(it) Then dup = dup & "," & it Else seen.Add it, True
            Next
            If lng <> "" Then AddRes S9, r, 11, SafeStr(ws.Cells(r, 3).Value), "NG", "exception_list に " & mx & " 文字を超える値があります：" & Mid$(lng, 2)
            If dup <> "" Then AddRes S9, r, 11, SafeStr(ws.Cells(r, 3).Value), "NG", "exception_list に同じ値が重複しています：" & Mid$(dup, 2)
        End If
    Next
End Sub

' 名前の集合（Dictionary）を作る
Private Function NameSet(ByVal sh As String, ByVal col As Long, ByVal condCol As Long, ByVal condVal As String, ByVal lastR As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    AddNames d, sh, col, condCol, condVal, lastR
    Set NameSet = d
End Function

Private Sub AddNames(ByVal d As Object, ByVal sh As String, ByVal col As Long, ByVal condCol As Long, ByVal condVal As String, _
                     Optional ByVal lastR As Long = 0)
    Dim ws As Worksheet, v As Variant, cv As Variant, i As Long, nm As String, firstR As Long
    Set ws = Worksheets(sh)
    firstR = IIf(sh = S2, EX_FIRST, IN_FIRST)
    If lastR = 0 Then
        lastR = LastRow(ws, col)
    ElseIf LastRow(ws, col) < lastR Then
        lastR = LastRow(ws, col)
    End If
    If lastR < firstR Then Exit Sub
    v = ws.Range(ws.Cells(firstR, col), ws.Cells(lastR, col)).Value
    If Not IsArray(v) Then v = OneCell(v)
    If condCol > 0 Then
        cv = ws.Range(ws.Cells(firstR, condCol), ws.Cells(lastR, condCol)).Value
        If Not IsArray(cv) Then cv = OneCell(cv)
    End If
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" Then
            If condCol = 0 Then
                d(nm) = True
            ElseIf SafeStr(cv(i, 1)) = condVal Then
                d(nm) = True
            End If
        End If
    Next
End Sub

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

' 結果シートを作り直す
Private Sub WriteResult()
    Dim ws As Worksheet, r As Long, it As Variant, a() As Variant, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_RES)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_RES
    End If
    ws.Cells.Clear
    ws.Cells.Font.Name = "Meiryo UI"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "チェック結果（記入状況の「一括チェック」で作り直します）"
    With ws.Range("A1:F1")
        .Interior.Color = RGB(31, 56, 100)
    End With
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = vbWhite
    End With
    ws.Rows(1).RowHeight = 26
    ws.Range("A2").Value = "実行日時"
    ws.Range("B2").Value = Now
    ws.Range("B2").NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Range("A4:F4").Value = Array("No", "シート", "セル（クリックで移動）", "名前", "判定", "内容")
    With ws.Range("A4:F4")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(47, 85, 151)
    End With
    r = 5
    If mRes.Count > 0 Then
        ReDim a(1 To mRes.Count, 1 To 6)
        i = 0
        For Each it In mRes
            i = i + 1
            a(i, 1) = i
            a(i, 2) = it(0)
            a(i, 3) = Worksheets(it(0)).Cells(it(1), it(2)).Address(False, False)
            a(i, 4) = "'" & it(3)
            a(i, 5) = it(4)
            a(i, 6) = it(5)
        Next
        ws.Range("A5").Resize(mRes.Count, 6).Value = a
        i = 0
        For Each it In mRes
            ws.Hyperlinks.Add Anchor:=ws.Cells(r + i, 3), Address:="", _
                SubAddress:="'" & it(0) & "'!" & Worksheets(it(0)).Cells(it(1), it(2)).Address(False, False), _
                TextToDisplay:=CStr(a(i + 1, 3))
            If it(4) = "NG" Then ws.Cells(r + i, 5).Interior.Color = CL_NG Else ws.Cells(r + i, 5).Interior.Color = CL_WARN
            i = i + 1
        Next
    Else
        ws.Range("A5").Value = "NG・警告はありません。"
    End If
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 90
    ws.Columns("F").WrapText = True
End Sub

'==========================================================
' メンバー選択画面（ダブルクリックでチェックボックス一覧）
'   ブックを開くと Auto_Open で自動有効化されます。
'   ダブルクリックが効かない時は PickerStart を実行（または Ctrl+Shift+M）
'   対象：④ E列（static）、⑥ D列、⑧ ベンダー欄、⑧-2 F・G列、⑪ G列（ユーザ行）
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
    Set mNames = CreateObject("Scripting.Dictionary")   ' バイナリ比較（大文字・小文字を区別）
    mExclude = ""
    mVendor = False
    If modInPolicy.PolicyTryPick(Target) Then TryPick = True: Exit Function
    If Not PickerRule(Target.Worksheet, Target, ttl) Then Exit Function
    TryPick = True
    ShowPicker Target, ttl
End Function

' 対象列の定義
'   AddSrc シート名, name列, 種別ラベル [, 条件列, 条件値]
'   AddMark シート名, name列, 条件列, 条件値, 付記
Private Function PickerRule(ByVal ws As Worksheet, ByVal c As Range, ByRef ttl As String) As Boolean
    Dim key As String, cat As String, nm As String
    If c.Row < IN_FIRST Or c.Row > IN_LAST Then Exit Function
    nm = SafeStr(ws.Cells(c.Row, 3).Value)
    Select Case ws.Name
        Case S4
            If c.Column <> 5 Or SafeStr(ws.Cells(c.Row, 4).Value) <> "static" Then Exit Function
            If SafeStr(ws.Cells(c.Row, 2).Value) = "削除" Then Exit Function
            AddSrc S2, OFF_ADR + 3, "アドレス（既存）"
            AddSrc S3, 3, "アドレス（③で登録）", 2, "登録"
            AddMark S3, 3, 2, "削除", " ※③で削除予定"
            ttl = "メンバー選択（アドレス）：" & nm
            PickerRule = True
        Case S6
            If c.Column <> 4 Or SafeStr(ws.Cells(c.Row, 2).Value) = "削除" Then Exit Function
            mExclude = nm
            AddSrc S2, OFF_SVC + 3, "サービス（既存）"
            AddSrc S5, 3, "サービス（⑤で登録）", 2, "登録"
            AddSrc S2, OFF_SVCG + 3, "サービスグループ（既存）"
            AddSrc S6, 3, "サービスグループ（⑥で登録）", 2, "登録"
            If Not mNames.Exists("service-http") Then mNames.Add "service-http", "定義済み"
            If Not mNames.Exists("service-https") Then mNames.Add "service-https", "定義済み"
            AddMark S5, 3, 2, "削除", " ※⑤で削除予定"
            AddMark S6, 3, 2, "削除", " ※⑥で削除予定"
            ttl = "メンバー選択（サービス／サービスグループ）：" & nm
            PickerRule = True
        Case S8
            If c.Column < F_FIRST Or c.Column > F_LAST Then Exit Function
            key = FieldKeys()(c.Column - F_FIRST)
            If Right$(key, 7) <> "_vendor" Then Exit Function
            cat = CatLabel(Left$(key, InStr(key, "_") - 1))
            AddVendorSrc cat
            mVendor = True
            ttl = cat & " のベンダー・製品選択：" & nm
            PickerRule = True
        Case S8B
            If c.Column < 6 Or c.Column > 7 Then Exit Function
            If SafeStr(ws.Cells(c.Row, 5).Value) = MODE_RAW Or SafeStr(ws.Cells(c.Row, 2).Value) = "削除" Then Exit Function
            AddSrc S2, OFF_HIP + 3, "HIPオブジェクト（既存）"
            AddSrc S8, 3, "HIPオブジェクト（⑧で登録）", 2, "登録"
            AddMark S8, 3, 2, "削除", " ※⑧で削除予定"
            ttl = IIf(c.Column = 6, "対象HIPオブジェクト選択：", "除外HIPオブジェクト(not)選択：") & nm
            PickerRule = True
        Case S11
            If c.Column <> 7 Or SafeStr(ws.Cells(c.Row, 3).Value) <> "ユーザ" Then Exit Function
            AddSrc S11, 4, "グループ", 3, "グループ"
            AddMark S11, 4, 2, "削除", " ※削除予定"
            ttl = "所属グループ選択：" & SafeStr(ws.Cells(c.Row, 4).Value)
            PickerRule = True
    End Select
End Function

Private Sub AddSrc(ByVal sh As String, ByVal col As Long, ByVal kind As String, _
                   Optional ByVal condCol As Long = 0, Optional ByVal condVal As String = "")
    Dim v As Variant, cv As Variant, i As Long, nm As String
    If Not ReadCols(sh, col, condCol, v, cv) Then Exit Sub
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

Private Sub AddMark(ByVal sh As String, ByVal col As Long, ByVal condCol As Long, ByVal condVal As String, ByVal note As String)
    Dim v As Variant, cv As Variant, i As Long, nm As String
    If Not ReadCols(sh, col, condCol, v, cv) Then Exit Sub
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" And SafeStr(cv(i, 1)) = condVal Then
            If mNames.Exists(nm) Then mNames(nm) = mNames(nm) & note
        End If
    Next
End Sub

' name列（と条件列）を2次元配列で一括読込（②は8行目～、③～⑪も8行目～）
Private Function ReadCols(ByVal sh As String, ByVal col As Long, ByVal condCol As Long, _
                          ByRef v As Variant, ByRef cv As Variant) As Boolean
    Dim ws As Worksheet, lastR As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sh)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastR = LastRow(ws, col)
    If sh <> S2 And lastR > IN_LAST Then lastR = IN_LAST
    If lastR < IN_FIRST Then Exit Function
    v = ws.Range(ws.Cells(IN_FIRST, col), ws.Cells(lastR, col)).Value
    If Not IsArray(v) Then v = OneCell(v)
    If condCol > 0 Then
        cv = ws.Range(ws.Cells(IN_FIRST, condCol), ws.Cells(lastR, condCol)).Value
        If Not IsArray(cv) Then cv = OneCell(cv)
    End If
    ReadCols = True
End Function

Private Sub ShowPicker(ByVal Target As Range, ByVal ttl As String)
    Dim f As frmMemberPick, cur As String, sep As String, res As String
    cur = SafeStr(Target.Value)
    If InStr(cur, vbLf) > 0 Then sep = vbLf Else sep = ","
    Set f = New frmMemberPick
    If mVendor Then
        f.Setup ttl, mNames, VendorItems(cur), vbLf
    Else
        f.Setup ttl, mNames, SplitItems(cur), sep
    End If
    f.Show
    If f.Confirmed Then
        If mVendor Then res = VendorJoin(f.ResultItems) Else res = f.Result
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
    Dim c As New Collection, a As Variant, i As Long
    a = SplitVals(s)      ' DN（CN=…,OU=…）は1件のまま
    For i = 0 To UBound(a)
        c.Add CStr(a(i))
    Next
    Set SplitItems = c
End Function

' 複数値の分割（カンマ・セル内改行）。DN（CN=…,OU=…,DC=…）の途中のカンマでは分けない
'   同じ行でカンマの後が OU= / DC= / O= / L= / ST= / C= / STREET= で始まり、直前の値が「=」を含むときは、1つの値として続ける
'   例：CN=xxxx,OU=xxxx,DC=xxx,DC=local → 1件 ／ CN=a,OU=x,CN=b,OU=y → 2件
Private Function SplitVals(ByVal s As String) As Variant
    Dim lines() As String, parts() As String, res() As String, i As Long, j As Long, n As Long, t As String, lineTop As Long
    SplitVals = Split("")
    If Len(s) = 0 Then Exit Function
    ReDim res(0 To Len(s))
    n = -1
    lines = Split(Replace(s, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        lineTop = n + 1
        parts = Split(lines(i), ",")
        For j = 0 To UBound(parts)
            t = Trim$(parts(j))
            If t <> "" Then
                If n >= lineTop And IsDnPart(t) And InStr(res(n), "=") > 0 Then
                    res(n) = res(n) & "," & t
                Else
                    n = n + 1
                    res(n) = t
                End If
            End If
        Next
    Next
    If n < 0 Then Exit Function
    ReDim Preserve res(0 To n)
    SplitVals = res
End Function

' DN の2番目以降の要素（OU= など）か
Private Function IsDnPart(ByVal t As String) As Boolean
    Dim p As Long
    p = InStr(t, "=")
    If p < 2 Then Exit Function
    Select Case UCase$(Trim$(Left$(t, p - 1)))
        Case "OU", "DC", "O", "L", "ST", "C", "STREET": IsDnPart = True
    End Select
End Function

'----------------------------------------------------------
' ⑧ ベンダー欄のピッカー（参照_HIPベンダー一覧 の同じカテゴリの行）
'----------------------------------------------------------
Private Function FieldKeys() As Variant
    FieldKeys = Array("description", _
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

Private Sub AddVendorSrc(ByVal cat As String)
    Dim ws As Worksheet, v As Variant, r As Long, lastR As Long, vn As String, pr As String
    Set ws = Worksheets(R2)
    lastR = LastRow(ws, 3)
    If lastR < VD_FIRST Then Exit Sub
    v = ws.Range(ws.Cells(VD_FIRST, 2), ws.Cells(lastR, 4)).Value
    For r = 1 To UBound(v, 1)
        If Trim$(SafeStr(v(r, 1))) = cat Then
            vn = Trim$(SafeStr(v(r, 2)))
            pr = Trim$(SafeStr(v(r, 3)))
            If vn <> "" Then
                If Not mNames.Exists(vn) Then mNames.Add vn, "製品指定なし（このベンダーの全製品）"
                If pr <> "" Then
                    If Not mNames.Exists(vn & " | " & pr) Then mNames.Add vn & " | " & pr, "製品"
                End If
            End If
        End If
    Next
End Sub

' セルの値「ベンダー | 製品1 ; 製品2」の行 → 「ベンダー | 製品」単位の一覧
Private Function VendorItems(ByVal s As String) As Collection
    Dim c As New Collection, ln As Variant, a As Variant, q As Variant, n As Long
    For Each ln In Lines(s)
        a = Parts(ln)
        n = 0
        If UBound(a) >= 1 Then
            For Each q In Split(a(1), ";")
                If Trim$(q) <> "" Then
                    c.Add a(0) & " | " & Trim$(q)
                    n = n + 1
                End If
            Next
        End If
        If n = 0 Then c.Add a(0)
    Next
    Set VendorItems = c
End Function

' 選ばれた「ベンダー」「ベンダー | 製品」→ ベンダーごとに1行
Private Function VendorJoin(ByVal items As Collection) As String
    Dim order As New Collection, prods As Object, anyP As Object, it As Variant, a As Variant, vn As Variant, s As String
    Set prods = CreateObject("Scripting.Dictionary")
    Set anyP = CreateObject("Scripting.Dictionary")
    For Each it In items
        a = Parts(CStr(it))
        If Not prods.Exists(a(0)) Then
            prods.Add a(0), ""
            order.Add a(0)
        End If
        If UBound(a) >= 1 Then
            If prods(a(0)) = "" Then prods(a(0)) = a(1) Else prods(a(0)) = prods(a(0)) & " ; " & a(1)
        Else
            anyP(a(0)) = True
        End If
    Next
    For Each vn In order
        If anyP.Exists(vn) Or prods(vn) = "" Then
            s = AddLine(s, CStr(vn))
        Else
            s = AddLine(s, vn & " | " & prods(vn))
        End If
    Next
    VendorJoin = s
End Function

'----------------------------------------------------------
' 共通の小さな関数
'----------------------------------------------------------
Private Function Lines(ByVal s As String) As Collection
    Dim c As New Collection, a() As String, i As Long
    a = Split(Replace(s, vbCr, ""), vbLf)
    For i = LBound(a) To UBound(a)
        If Trim$(a(i)) <> "" Then c.Add Trim$(a(i))
    Next
    Set Lines = c
End Function

Private Function Parts(ByVal s As String) As Variant
    Dim a() As String, i As Long
    a = Split(s, "|")
    For i = LBound(a) To UBound(a)
        a(i) = Trim$(a(i))
    Next
    Parts = a
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

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

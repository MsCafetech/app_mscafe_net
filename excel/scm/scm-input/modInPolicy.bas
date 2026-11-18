Attribute VB_Name = "modInPolicy"
'==========================================================
' modInPolicy : 統合入力シートのセキュリティポリシー／復号ポリシー
'   ②-2・②-3 = 既存ポリシー（①の「既存情報をまとめて取得（API）」で GET した結果。列は既存ツールの「②SecRule_取得」「②DecRule_取得」と同じ）
'   ⑫・⑬     = 入力（既存ツールの「③…_設定」と同じ列。操作：追加・変更・削除・移動・変更+移動）
'   参照_ポリシー選択肢 = ダブルクリック選択の候補・入力チェックの基準（②既存データ＋③～⑪の入力から作る）
'   スコープ  : ルールの folder から見えるのは自分＋上位（RN・SC・MUC→Shared、GP・EP→MUC・Shared）。All は参照のみ
'   前提      : modScmInput・modInFetch・modScmCommon と同じブックに入れる（API は GET だけ。JSON化・登録はしません）
'   並び順はセキュリティ・復号とも同じ（Global Pre → Prisma Access Pre → MUC Pre → GP → EP → MUC Post → RN → SC → Prisma Access Post → Global Post）
'   GP・EP・RN・SC は Pre/Post なし
'==========================================================
Option Explicit

Private Const SH_GSEC As String = "②-2既存セキュリティポリシー"
Private Const SH_GDEC As String = "②-3既存復号ポリシー"
Private Const SH_SSEC As String = "⑫セキュリティポリシー"
Private Const SH_SDEC As String = "⑬復号ポリシー"
Private Const SH_CHO As String = "参照_ポリシー選択肢"
Private Const S2 As String = "②既存データ抽出"
Private Const S3 As String = "③アドレス"
Private Const S4 As String = "④アドレスグループ"
Private Const S5 As String = "⑤サービス"
Private Const S6 As String = "⑥サービスグループ"
Private Const S7 As String = "⑦カスタムURLカテゴリ"
Private Const S8B As String = "⑧-2HIPプロファイル"
Private Const S9 As String = "⑨EDL"
Private Const S10 As String = "⑩アプリケーション"
Private Const S11 As String = "⑪ユーザ・グループ"

Private Const FIRST As Long = 8
Private Const IN_FIRST As Long = 8
Private Const IN_LAST As Long = 507
Private Const CHO_LAST As Long = 10007
Private Const HDR As String = "▼"

' ② ブロックの位置（modScmInput と同じ）
Private Const OFF_ADR As Long = 0
Private Const OFF_ADRG As Long = 14
Private Const OFF_SVC As Long = 26
Private Const OFF_SVCG As Long = 40
Private Const OFF_URL As Long = 47
Private Const OFF_HIPP As Long = 124
Private Const OFF_EDL As Long = 133
Private Const OFF_PG As Long = 153

' ⑦ カスタムURLカテゴリ（横長）
Private Const U_C1 As Long = 2
Private Const U_C2 As Long = 51
Private Const U_OP As Long = 6
Private Const U_NAME As Long = 7

' 取得・設定シートの列（項目数 NF で変わるものは SetMode で設定）
Private Const G_SCOPE As Long = 2
Private Const G_FOLDER As Long = 5
Private Const G_ID As Long = 7
Private Const G_NAME As Long = 8
Private Const G_F0 As Long = 9
Private Const S_OP As Long = 2
Private Const S_JUDGE As Long = 3
Private Const S_REASON As Long = 4
Private Const S_SCOPE As Long = 5
Private Const S_FOLDER As Long = 8
Private Const S_POS As Long = 9
Private Const S_ID As Long = 10
Private Const S_NAME As Long = 11
Private Const S_F0 As Long = 12

' 参照_ポリシー選択肢 の列（名前の右隣 = スコープ（定義folder）。空欄＝どこでも）
Private Const CH_ZONE As Long = 2
Private Const CH_ADDR As Long = 4
Private Const CH_ADDRG As Long = 6
Private Const CH_EDL As Long = 8
Private Const CH_EDLTYPE As Long = 10
Private Const CH_REGION As Long = 11
Private Const CH_SVC As Long = 12
Private Const CH_SVCG As Long = 14
Private Const CH_URLC As Long = 16
Private Const CH_URLP As Long = 18
Private Const CH_TAG As Long = 19
Private Const CH_HIP As Long = 21
Private Const CH_USER As Long = 23
Private Const CH_APPP As Long = 24
Private Const CH_APPC As Long = 25
Private Const CH_APPG As Long = 27
Private Const CH_APPF As Long = 29
Private Const CH_CERT As Long = 31
Private Const CH_DECP As Long = 32
Private Const CH_PG As Long = 34

' 色（BGR）
Private Const C_IN As Long = &HCCF2FF
Private Const C_FX As Long = &HF2F2F2
Private Const C_WARN As Long = &H9CEBFF
Private Const C_ADD As Long = &HCEEFC6
Private Const C_DEL As Long = &HD9D9D9
Private Const C_NG As Long = &HCEC7FF
Private Const C_CHG As Long = &HADCBF8

Private Const CELL_MAX As Long = 32767

' モード（sec / dec）
Private mMode As String
Private NF As Long
Private G_LAST As Long
Private G_HF As Long
Private S_LAST As Long
Private S_HF As Long
Private mGet As String
Private mSet As String
Private mFK As Variant
Private mFT As Variant
Private mFL As Variant

Private mRef As Object       ' 種類|name → スコープ（* = どこでも）
Private mDel As Object       ' 種類|name → 削除予定のシート
Private mNames As Object     ' ピッカー候補

Private Sub SetMode(ByVal m As String)
    mMode = m
    If m = "dec" Then
        mFK = Array("disabled", "tag", "from", "source", "negate_source", "source_user", "source_hip", _
                    "to", "destination", "negate_destination", "destination_hip", "service", "category", _
                    "action", "type", "@cert", "profile", "log_success", "log_fail", "log_setting", "description")
        mFT = Array("b", "a", "a", "a", "b", "a", "a", _
                    "a", "a", "b", "a", "a", "a", _
                    "s", "t", "c", "s", "b", "b", "s", "s")
        mFL = Array("無効", "タグ", "送信元ゾーン", "送信元アドレス", "送信元否定", "送信元ユーザー", "送信元HIP", _
                    "宛先ゾーン", "宛先アドレス", "宛先否定", "宛先HIP", "サービス", "URLカテゴリ", _
                    "アクション", "種類", "証明書", "復号プロファイル", "成功ログ", "失敗ログ", "ログ転送", "説明")
        mGet = SH_GDEC: mSet = SH_SDEC
    Else
        mFK = Array("disabled", "tag", "from", "source", "negate_source", "source_user", "source_hip", _
                    "to", "destination", "negate_destination", "destination_hip", "application", "service", "category", _
                    "action", "profile_setting", "log_setting", "description")
        mFT = Array("b", "a", "a", "a", "b", "a", "a", _
                    "a", "a", "b", "a", "a", "a", "a", _
                    "s", "g", "s", "s")
        mFL = Array("無効", "タグ", "送信元ゾーン", "送信元アドレス", "送信元否定", "送信元ユーザー", "送信元HIP", _
                    "宛先ゾーン", "宛先アドレス", "宛先否定", "宛先HIP", "アプリケーション", "サービス", "URLカテゴリ", _
                    "アクション", "プロファイルグループ", "ログ転送", "説明")
        mGet = SH_GSEC: mSet = SH_SSEC
    End If
    NF = UBound(mFK) + 1
    G_LAST = G_NAME + NF + 1
    G_HF = G_LAST + 2
    S_LAST = S_NAME + NF
    S_HF = S_LAST + 2
End Sub

Private Function FIdx(ByVal k As String) As Long
    Dim i As Long
    For i = 0 To NF - 1
        If mFK(i) = k Then FIdx = i + 1: Exit Function
    Next
End Function

Private Function Cmp(ByVal i As Long, ByVal s As String) As String
    If i = 0 Then Cmp = Trim$(s): Exit Function
    Select Case mFT(i - 1)
        Case "a", "g": Cmp = Join(PSplit(s), vbLf)
        Case "b": If LCase$(Trim$(s)) = "true" Then Cmp = "true"
        Case Else
            If mFK(i - 1) = "description" Then Cmp = s Else Cmp = Trim$(s)
    End Select
End Function

'==========================================================
' ボタン（modScmInput.SetupButtons から呼ぶ）
'==========================================================
Public Sub SetupPolicyButtons()
    PBtn SH_SSEC, "F3:H4", "既存ポリシーを展開（作り直し）", "ExpandSec"
    PBtn SH_SSEC, "J3:K4", "選択行の下に追加行を挿入", "InsertSec"
    PBtn SH_SSEC, "M3:N4", "入力チェック", "CheckSecBtn"
    PBtn SH_SDEC, "F3:H4", "既存ポリシーを展開（作り直し）", "ExpandDec"
    PBtn SH_SDEC, "J3:K4", "選択行の下に追加行を挿入", "InsertDec"
    PBtn SH_SDEC, "M3:N4", "入力チェック", "CheckDecBtn"
    PBtn SH_CHO, "D3:E4", "選択肢を作り直す（②と③～⑪から）", "BuildChoicesBtn"
End Sub

Private Sub PBtn(ByVal sh As String, ByVal addr As String, ByVal caption As String, ByVal macro As String)
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

Public Sub ExpandSec()
    SetMode "sec"
    CopyGetToSet
End Sub

Public Sub ExpandDec()
    SetMode "dec"
    CopyGetToSet
End Sub

Public Sub InsertSec()
    SetMode "sec"
    InsertRuleRow
End Sub

Public Sub InsertDec()
    SetMode "dec"
    InsertRuleRow
End Sub

Public Sub CheckSecBtn()
    CheckOne "sec"
End Sub

Public Sub CheckDecBtn()
    CheckOne "dec"
End Sub

Private Sub CheckOne(ByVal m As String)
    Dim ng As Long, okN As Long, warnN As Long
    BuildChoices
    SetMode m
    ng = CheckSet(okN, warnN)
    Worksheets(mSet).Activate
    MsgBox "入力チェック結果：操作のある行 " & okN & " 件（うち警告 " & warnN & " 件）／ NG " & ng & " 件" & vbLf & _
           "判定・理由は C・D 列、既存から変わったセルは橙色で表示しています。", IIf(ng > 0, vbExclamation, vbInformation)
End Sub

' 一括チェック（modScmInput.CheckAll から呼ぶ）。判定は各シートの C・D 列に書く
Public Sub CheckPolicies()
    Dim okN As Long, warnN As Long
    BuildChoices
    SetMode "sec"
    CheckSet okN, warnN
    SetMode "dec"
    CheckSet okN, warnN
End Sub

'==========================================================
' 参照_ポリシー選択肢 を作り直す
'   作り直す列：アドレス・アドレスグループ・EDL・サービス・サービスグループ・カスタムURL・HIPプロファイル・プロファイルグループ
'     = ②既存データ（スコープ=コンテナ列の folder） ＋ ③～⑨の「登録」行（スコープ=①の設定先フォルダ）
'   追記する列：ユーザー（⑪）、定義済みApp・自作App（⑩）。取得したゾーン・タグ・リージョン等はそのまま
'   ③～⑪の「削除」行は mDel に記録（ポリシーで使っているとNG）
'==========================================================
Public Sub BuildChoicesBtn()
    BuildChoices
    Worksheets(SH_CHO).Activate
    MsgBox "選択肢を作り直しました。", vbInformation
End Sub

Public Sub BuildChoices()
    Dim wsC As Worksheet, cf As String, c As Variant
    Set wsC = Worksheets(SH_CHO)
    Set mDel = CreateObject("Scripting.Dictionary")
    cf = CustFolder()
    For Each c In Array(CH_ADDR, CH_ADDR + 1, CH_ADDRG, CH_ADDRG + 1, CH_EDL, CH_EDL + 1, CH_EDLTYPE, _
                        CH_SVC, CH_SVC + 1, CH_SVCG, CH_SVCG + 1, CH_URLC, CH_URLC + 1, CH_HIP, CH_HIP + 1, CH_PG, CH_PG + 1)
        wsC.Range(wsC.Cells(FIRST, c), wsC.Cells(CHO_LAST, c)).ClearContents
    Next
    FillBlock CH_ADDR, OFF_ADR, 13, 3, 0, 0, S3, cf, "addr"
    FillBlock CH_ADDRG, OFF_ADRG, 11, 3, 0, 0, S4, cf, "addr"
    FillBlock CH_EDL, OFF_EDL, 19, 3, 4, CH_EDLTYPE, S9, cf, "edl"
    FillBlock CH_SVC, OFF_SVC, 13, 3, 0, 0, S5, cf, "svc"
    FillBlock CH_SVCG, OFF_SVCG, 6, 3, 0, 0, S6, cf, "svc"
    FillBlock CH_URLC, OFF_URL, 6, 2, 0, 0, "", cf, "url"
    FillUrlInput cf
    FillBlock CH_HIP, OFF_HIPP, 8, 3, 0, 0, S8B, cf, "hip"
    FillBlock CH_PG, OFF_PG, 12, 3, 0, 0, "", cf, "pg"
    AppendList CH_USER, S11, 4, "user"
    AppendApps cf
    Set mRef = BuildRefMap()
End Sub

' ③～⑪で登録するオブジェクトの folder（①基本設定の「オブジェクト取得folder」。通常 Shared）
Private Function CustFolder() As String
    On Error Resume Next
    CustFolder = Trim$(SafeStr(ThisWorkbook.Names("CFG_OBJ_FOLDER").RefersToRange.Value))
    On Error GoTo 0
    If CustFolder = "" Then CustFolder = "Shared"
    CustFolder = NormFolder(CustFolder)
End Function

' ②のブロック（name列・コンテナ列）＋入力シートの登録行（C列=name、typeCol=種類）を書く
Private Sub FillBlock(ByVal ch As Long, ByVal off As Long, ByVal ncols As Long, ByVal nameCol As Long, ByVal typeCol As Long, _
                      ByVal chType As Long, ByVal inSh As String, ByVal cf As String, ByVal grp As String)
    Dim ws2 As Worksheet, wsC As Worksheet, wsI As Worksheet, contCol As Long, r As Long, o As Long, nm As String
    Dim lastR As Long, op As String, seen As Object, i As Long, h As String
    Set ws2 = Worksheets(S2)
    Set wsC = Worksheets(SH_CHO)
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To ncols
        h = SafeStr(ws2.Cells(7, off + i).Value)
        If InStr(h, "コンテナ") > 0 Or InStr(h, "定義folder") > 0 Or LCase$(h) = "folder" Then contCol = off + i: Exit For
    Next
    o = FIRST
    lastR = LastRow(ws2, off + nameCol)
    For r = FIRST To lastR
        nm = Trim$(SafeStr(ws2.Cells(r, off + nameCol).Value))
        If nm <> "" And o <= CHO_LAST Then
            wsC.Cells(o, ch).NumberFormat = "@"
            wsC.Cells(o, ch).Value = nm
            If contCol > 0 Then wsC.Cells(o, ch + 1).Value = FolderDisp(NormFolder(ContFolder(SafeStr(ws2.Cells(r, contCol).Value))))
            If chType > 0 Then wsC.Cells(o, chType).Value = SafeStr(ws2.Cells(r, off + typeCol).Value)
            seen(nm) = True
            o = o + 1
        End If
    Next
    If inSh = "" Then Exit Sub
    Set wsI = Worksheets(inSh)
    For r = IN_FIRST To IN_LAST
        nm = Trim$(SafeStr(wsI.Cells(r, 3).Value))
        op = Trim$(SafeStr(wsI.Cells(r, 2).Value))
        If nm <> "" Then
            If op = "登録" And Not seen.Exists(nm) And o <= CHO_LAST Then
                wsC.Cells(o, ch).NumberFormat = "@"
                wsC.Cells(o, ch).Value = nm
                wsC.Cells(o, ch + 1).Value = FolderDisp(cf)
                If chType > 0 Then wsC.Cells(o, chType).Value = SafeStr(wsI.Cells(r, typeCol).Value)
                seen(nm) = True
                o = o + 1
            ElseIf op = "削除" Then
                mDel(grp & "|" & nm) = inSh
            End If
        End If
    Next
End Sub

' 「folder=Shared」「Shared」などの表記から folder を取り出す
Private Function ContFolder(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, "=")
    If p > 0 Then s = Mid$(s, p + 1)
    ContFolder = Trim$(s)
End Function

' ⑦ の登録列（名前を追加）・削除列
Private Sub FillUrlInput(ByVal cf As String)
    Dim ws7 As Worksheet, wsC As Worksheet, c As Long, nm As String, op As String, o As Long
    Set ws7 = Worksheets(S7)
    Set wsC = Worksheets(SH_CHO)
    o = NextRow(wsC, CH_URLC)
    For c = U_C1 To U_C2
        nm = Trim$(SafeStr(ws7.Cells(U_NAME, c).Value))
        op = Trim$(SafeStr(ws7.Cells(U_OP, c).Value))
        If nm <> "" Then
            If op = "登録" Then
                wsC.Cells(o, CH_URLC).NumberFormat = "@"
                wsC.Cells(o, CH_URLC).Value = nm
                wsC.Cells(o, CH_URLC + 1).Value = cf
                o = o + 1
            ElseIf op = "削除" Then
                mDel("url|" & nm) = S7
            End If
        End If
    Next
End Sub

' ⑪ のユーザ・グループを追記（削除は mDel）
Private Sub AppendList(ByVal ch As Long, ByVal sh As String, ByVal nameCol As Long, ByVal grp As String)
    Dim ws As Worksheet, wsC As Worksheet, have As Object, r As Long, o As Long, nm As String, op As String
    Set ws = Worksheets(sh)
    Set wsC = Worksheets(SH_CHO)
    Set have = ColSet(wsC, ch)
    o = NextRow(wsC, ch)
    For r = IN_FIRST To IN_LAST
        nm = Trim$(SafeStr(ws.Cells(r, nameCol).Value))
        op = Trim$(SafeStr(ws.Cells(r, 2).Value))
        If nm <> "" Then
            If op = "削除" Then
                mDel(grp & "|" & nm) = sh
            ElseIf Not have.Exists(nm) Then
                wsC.Cells(o, ch).NumberFormat = "@"
                wsC.Cells(o, ch).Value = nm
                have(nm) = True
                o = o + 1
            End If
        End If
    Next
End Sub

' ⑩ のアプリ：定義済み → 定義済みApp列、カスタム → 自作App列（スコープ=設定先フォルダ）
Private Sub AppendApps(ByVal cf As String)
    Dim ws As Worksheet, wsC As Worksheet, hp As Object, hc As Object, r As Long, rowP As Long, rowC As Long
    Dim nm As String, op As String, kind As String
    Set ws = Worksheets(S10)
    Set wsC = Worksheets(SH_CHO)
    Set hp = ColSet(wsC, CH_APPP)
    Set hc = ColSet(wsC, CH_APPC)
    rowP = NextRow(wsC, CH_APPP)
    rowC = NextRow(wsC, CH_APPC)
    For r = IN_FIRST To IN_LAST
        nm = Trim$(SafeStr(ws.Cells(r, 3).Value))
        op = Trim$(SafeStr(ws.Cells(r, 2).Value))
        kind = Trim$(SafeStr(ws.Cells(r, 4).Value))
        If nm <> "" Then
            If op = "削除" Then
                mDel("app|" & nm) = S10
            ElseIf kind = "カスタム" Then
                If Not hc.Exists(nm) Then
                    wsC.Cells(rowC, CH_APPC).NumberFormat = "@"
                    wsC.Cells(rowC, CH_APPC).Value = nm
                    wsC.Cells(rowC, CH_APPC + 1).Value = cf
                    hc(nm) = True
                    rowC = rowC + 1
                End If
            ElseIf Not hp.Exists(nm) Then
                wsC.Cells(rowP, CH_APPP).NumberFormat = "@"
                wsC.Cells(rowP, CH_APPP).Value = nm
                hp(nm) = True
                rowP = rowP + 1
            End If
        End If
    Next
End Sub

Private Function ColSet(ByVal ws As Worksheet, ByVal col As Long) As Object
    Dim d As Object, r As Long, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    For r = FIRST To LastRow(ws, col)
        nm = Trim$(SafeStr(ws.Cells(r, col).Value))
        If nm <> "" Then d(nm) = True
    Next
    Set ColSet = d
End Function

Private Function NextRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    NextRow = LastRow(ws, col) + 1
    If NextRow < FIRST Then NextRow = FIRST
End Function

'==========================================================
' 取得シート・入力シートの読み取り
'   戻り値 Collection of Array(folder, position, 表示名, Collection of rec)   rec = Array(id, 値(0=name..NF), 行, 操作)
'==========================================================
Private Function ReadGroups(ByVal ws As Worksheet, ByVal isSet As Boolean) As Collection
    Dim groups As New Collection, cur As Collection, arr As Variant, lastR As Long, i As Long, k As Long
    Dim cId As Long, cName As Long, cOp As Long, cHF As Long, id As String, nm As String, op As String, v() As String
    Dim folder As String
    If isSet Then
        cId = S_ID: cName = S_NAME: cOp = S_OP: cHF = S_HF
    Else
        cId = G_ID: cName = G_NAME: cOp = 0: cHF = G_HF
    End If
    Set ReadGroups = groups
    lastR = UsedLast(ws)
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
            If id <> "" Or Trim$(nm) <> "" Or op <> "" Then
                ReDim v(0 To NF)
                For k = 0 To NF
                    v(k) = SafeStr(arr(i, cName + k))
                Next
                cur.Add Array(id, v, FIRST + i - 1, op)
            End If
        End If
    Next
End Function

Private Function BeforeMap() As Object
    Dim d As Object, g As Variant, rec As Variant, n As Long
    Set d = CreateObject("Scripting.Dictionary")
    For Each g In ReadGroups(Worksheets(mGet), False)
        n = 0
        For Each rec In g(3)
            n = n + 1
            If rec(0) <> "" Then d(rec(0)) = Array(rec(1), g(0) & "|" & g(1), n)
        Next
    Next
    Set BeforeMap = d
End Function

'==========================================================
' ⑫・⑬ 入力シート
'==========================================================
' 既存ポリシー（②-2・②-3）を入力シートに展開。既存が無いときは、設定できるスコープの見出しだけ作る
Private Sub CopyGetToSet()
    Dim wsG As Worksheet, wsS As Worksheet, lastR As Long, arr As Variant, out() As Variant, i As Long, c As Long, n As Long
    Dim opN As Long, x As Variant
    Set wsG = Worksheets(mGet)
    Set wsS = Worksheets(mSet)
    opN = CountOps()
    If opN > 0 Then
        If MsgBox(mSet & " に操作の入った行が " & opN & " 件あります。" & vbLf & _
                  "既存ポリシーの内容で作り直すと、入力した内容はすべて消えます。よろしいですか？", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    Else
        If MsgBox(mGet & " の内容で " & mSet & " を作り直します。よろしいですか？", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    End If
    Application.ScreenUpdating = False
    lastR = UsedLast(wsG)
    If lastR >= FIRST Then
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
    Else
        ' 既存なし：設定できるスコープの見出しだけ（評価順）
        ReDim out(1 To FlowOrder().Count, 1 To S_HF + 1)
        For Each x In FlowOrder()
            n = n + 1
            out(n, 1) = HDR
            out(n, S_SCOPE) = "■ " & GroupLabel(CStr(x(0)), CStr(x(1))) & " ／ 0 件"
            out(n, S_HF) = x(0)
            out(n, S_HF + 1) = x(1)
        Next
    End If
    With wsS.Range(wsS.Cells(FIRST, 1), wsS.Cells(wsS.Rows.Count, S_HF + 1))
        .ClearContents
        .ClearFormats
        .FormatConditions.Delete
    End With
    wsS.Range(wsS.Cells(FIRST, 2), wsS.Cells(FIRST + n - 1, S_HF + 1)).NumberFormat = "@"
    wsS.Range(wsS.Cells(FIRST, 1), wsS.Cells(FIRST + n - 1, S_HF + 1)).Value = out
    For i = 1 To n
        If out(i, 1) = HDR Then
            PaintHeader wsS, FIRST + i - 1
        ElseIf SafeStr(out(i, S_NAME)) <> "" Or SafeStr(out(i, S_ID)) <> "" Then
            PaintBase wsS, FIRST + i - 1
            IndentRow wsS, FIRST + i - 1, S_SCOPE, SafeStr(out(i, S_FOLDER))
        End If
    Next
    Application.ScreenUpdating = True
    wsS.Activate
    MsgBox mSet & " を作り直しました。" & vbLf & _
           "変更・削除は「操作」列で選択、追加は追加したい位置の1つ上の行（先頭なら ▼の行）を選んで「選択行の下に追加行を挿入」を押してください。", vbInformation
End Sub

' 設定できる folder・position（評価順。セキュリティ・復号とも同じ）。All は参照のみなので作らない
'   GlobalProtect・Explicit Proxy・Remote Networks・Service Connections は1つのルールベース（Pre/Post なし）
Private Function FlowOrder() As Collection
    Dim c As New Collection
    c.Add Array("Shared", "pre")
    c.Add Array("Mobile Users Container", "pre")
    c.Add Array("Mobile Users", "pre")
    c.Add Array("Mobile Users Explicit Proxy", "pre")
    c.Add Array("Mobile Users Container", "post")
    c.Add Array("Remote Networks", "pre")
    c.Add Array("Service Connections", "pre")
    c.Add Array("Shared", "post")
    Set FlowOrder = c
End Function

Private Function CountOps() As Long
    Dim ws As Worksheet, lastR As Long
    Set ws = Worksheets(mSet)
    lastR = UsedLast(ws)
    If lastR < FIRST Then Exit Function
    CountOps = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(FIRST, S_OP), ws.Cells(lastR, S_OP)))
End Function

Private Function HeaderAbove(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim i As Long
    For i = r To FIRST Step -1
        If SafeStr(ws.Cells(i, 1).Value) = HDR Then HeaderAbove = i: Exit Function
    Next
End Function

Private Sub InsertRuleRow()
    Dim ws As Worksheet, r As Long, h As Long, folder As String, pos As String, k As Variant
    Set ws = Worksheets(mSet)
    If ActiveSheet.Name <> mSet Then
        ws.Activate
        MsgBox mSet & " で、追加したい位置の1つ上の行（先頭に入れるときは ▼スコープの行）を選択してから押してください。", vbInformation
        Exit Sub
    End If
    r = ActiveCell.Row
    h = HeaderAbove(ws, r)
    If r < FIRST Or h = 0 Then
        MsgBox "▼スコープの行、またはその下のルールの行を選択してください（▼の行が無いときは先に「既存ポリシーを展開」を押してください）。", vbExclamation
        Exit Sub
    End If
    folder = SafeStr(ws.Cells(h, S_HF).Value)
    pos = SafeStr(ws.Cells(h, S_HF + 1).Value)
    If LCase$(folder) = "all" Then
        MsgBox "グローバル（All）のスコープには追加できません（参照のみ）。", vbExclamation
        Exit Sub
    End If
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
    ws.Cells(r, S_FOLDER).Value = FolderDisp(folder)
    ws.Cells(r, S_POS).Value = pos
    For Each k In Array("from", "source", "source_user", "to", "destination", "application", "category")
        If FIdx(CStr(k)) > 0 Then ws.Cells(r, S_NAME + FIdx(CStr(k))).Value = "any"
    Next
    If mMode = "dec" Then
        ws.Cells(r, S_NAME + FIdx("service")).Value = "any"
        ws.Cells(r, S_NAME + FIdx("action")).Value = "decrypt"
        ws.Cells(r, S_NAME + FIdx("type")).Value = "ssl_forward_proxy"
    Else
        ws.Cells(r, S_NAME + FIdx("service")).Value = "application-default"
        ws.Cells(r, S_NAME + FIdx("action")).Value = "allow"
    End If
    PaintBase ws, r
    IndentRow ws, r, S_SCOPE, folder
    ws.Range(ws.Cells(r, S_OP), ws.Cells(r, S_LAST)).Interior.Color = C_ADD
    Application.ScreenUpdating = True
    ws.Cells(r, S_NAME).Select
    MsgBox "追加行を挿入しました。name と各項目を入力してください（ゾーン・アドレス等はダブルクリックで選択）。", vbInformation
End Sub

'==========================================================
' 入力チェック（判定・理由・色）。戻り値=NG件数
'==========================================================
Private Function CheckSet(ByRef okN As Long, ByRef warnN As Long) As Long
    Dim ws As Worksheet, before As Object, groups As Collection, g As Variant, rec As Variant, nameCnt As Object
    Dim folder As String, pos As String, prevIdx As Long, r As Long, id As String, op As String, v As Variant
    Dim b As Variant, bv As Variant, ngs As String, warns As String, i As Long, diffN As Long, judge As String, k As String
    Dim ng As Long
    Set ws = Worksheets(mSet)
    If mRef Is Nothing Then BuildChoices
    Set before = BeforeMap()
    Set groups = ReadGroups(ws, True)
    Set nameCnt = CreateObject("Scripting.Dictionary")
    For Each g In groups
        For Each rec In g(3)
            If rec(3) <> "削除" And Trim$(rec(1)(0)) <> "" Then
                k = g(0) & "|" & g(1) & "|" & Trim$(rec(1)(0))     ' Pre と Post は別のルールベース
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
            ws.Cells(r, S_FOLDER).Value = FolderDisp(folder)
            ws.Cells(r, S_POS).Value = pos
            PaintBase ws, r
            IndentRow ws, r, S_SCOPE, folder
            Select Case op
                Case ""
                    If id = "" Then ngs = ngs & vbLf & "NG:新しい行は操作に「追加」を選択してください"
                Case "追加"
                    If id <> "" Then ngs = ngs & vbLf & "NG:既存ルール（idあり）の行です。追加は「選択行の下に追加行を挿入」で作った行で行ってください"
                Case "変更", "削除", "移動", "変更+移動"
                    If id = "" Then ngs = ngs & vbLf & "NG:idがありません（" & op & "は既存ルールだけが対象です）"
                Case Else
                    ngs = ngs & vbLf & "NG:操作はプルダウンから選択してください"
            End Select
            If LCase$(folder) = "all" And op <> "" Then ngs = ngs & vbLf & "NG:グローバル（All）のルールは操作できません（参照のみ）"
            If id <> "" Then
                If before.Exists(id) Then
                    b = before(id)
                    bv = b(0)
                    If b(1) <> folder & "|" & pos Then
                        ngs = ngs & vbLf & "NG:別のスコープ（" & Replace(b(1), "|", "／") & "）のルールです。スコープをまたぐ移動はできません（削除＋追加にしてください）"
                    ElseIf op <> "移動" And op <> "変更+移動" And op <> "削除" Then
                        If b(2) < prevIdx Then ngs = ngs & vbLf & "NG:既存と順番が変わっています。並べ替えたルールの操作は「移動」（内容も変えるなら「変更+移動」）にしてください"
                        If b(2) > prevIdx Then prevIdx = b(2)
                    End If
                Else
                    ngs = ngs & vbLf & "NG:" & mGet & " に無いidです（既存ポリシーを取り込み直して展開し直してください）"
                End If
            End If
            If op = "追加" Or InStr(op, "変更") > 0 Then ngs = ngs & ContentNg(v, warns) & RefNg(v, folder, warns)
            If (op = "追加" Or InStr(op, "変更") > 0) And Trim$(v(0)) <> "" Then
                If nameCnt(folder & "|" & pos & "|" & Trim$(v(0))) > 1 Then ngs = ngs & vbLf & "NG:同じスコープ・同じ position（" & pos & "）内で name が重複しています"
            End If
            If op = "追加" Then
                ws.Range(ws.Cells(r, S_SCOPE), ws.Cells(r, S_LAST)).Interior.Color = C_ADD
            ElseIf op = "削除" Then
                With ws.Range(ws.Cells(r, S_SCOPE), ws.Cells(r, S_LAST))
                    .Interior.Color = C_DEL
                    .Font.Strikethrough = True
                End With
            End If
            diffN = 0
            If IsArray(bv) Then
                For i = 0 To NF
                    If Cmp(i, CStr(v(i))) <> Cmp(i, CStr(bv(i))) Then
                        diffN = diffN + 1
                        If op <> "削除" Then
                            ws.Cells(r, S_NAME + i).Interior.Color = C_CHG
                            ws.Cells(r, S_NAME + i).Font.Bold = True
                        End If
                    End If
                Next
                If diffN > 0 And InStr(op, "変更") = 0 And op <> "削除" And LCase$(folder) <> "all" Then
                    warns = warns & vbLf & "警告:既存から内容が変わっていますが、操作が「変更」ではないため反映されません"
                End If
                If diffN = 0 And InStr(op, "変更") > 0 Then warns = warns & vbLf & "警告:既存と差分がありません"
            End If
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
                Case "NG"
                    ws.Range(ws.Cells(r, S_JUDGE), ws.Cells(r, S_REASON)).Interior.Color = C_NG
                    ng = ng + 1
                Case "警告"
                    ws.Range(ws.Cells(r, S_JUDGE), ws.Cells(r, S_REASON)).Interior.Color = C_WARN
                Case "OK"
                    ws.Cells(r, S_JUDGE).Interior.Color = C_ADD
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

Private Function ContentNg(ByVal v As Variant, ByRef warns As String) As String
    Dim s As String, i As Long, k As String, a As Variant, j As Long, hasAny As Boolean, act As String, x As String
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
                a = PSplit(CStr(v(i + 1)))
                Select Case k
                    Case "from", "source", "source_user", "to", "destination", "application", "service", "category"
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
    act = Trim$(CStr(v(FIdx("action"))))
    If mMode = "dec" Then
        Select Case act
            Case "decrypt", "no-decrypt"
            Case "": s = s & vbLf & "NG:アクション（decrypt / no-decrypt）を選択してください"
            Case Else: s = s & vbLf & "NG:アクションは decrypt / no-decrypt のどちらかです"
        End Select
        x = Trim$(CStr(v(FIdx("@cert"))))
        Select Case Trim$(CStr(v(FIdx("type"))))
            Case "ssl_forward_proxy"
                If x <> "" Then warns = warns & vbLf & "警告:証明書は種類が ssl_inbound_inspection のときだけ使います"
            Case "ssl_inbound_inspection"
                If x = "" Then
                    s = s & vbLf & "NG:ssl_inbound_inspection のときは証明書を指定してください"
                ElseIf ColCount(CH_CERT) > 0 And Not InCol(CH_CERT, x) Then
                    warns = warns & vbLf & "警告:証明書「" & x & "」が参照_ポリシー選択肢にありません（大文字・小文字も区別）"
                End If
            Case ""
                If act = "decrypt" Then s = s & vbLf & "NG:種類（ssl_forward_proxy など）を選択してください"
            Case Else
                s = s & vbLf & "NG:種類は ssl_forward_proxy / ssl_inbound_inspection のどちらかです"
        End Select
        x = Trim$(CStr(v(FIdx("profile"))))
        If x <> "" Then
            If ColCount(CH_DECP) > 0 And Not InCol(CH_DECP, x) Then warns = warns & vbLf & "警告:復号プロファイル「" & x & "」が参照_ポリシー選択肢にありません（大文字・小文字も区別）"
        ElseIf act = "decrypt" Then
            warns = warns & vbLf & "警告:復号プロファイルが空欄です（既定のプロファイルになります）"
        End If
    Else
        Select Case act
            Case "allow", "deny", "drop", "reset-client", "reset-server", "reset-both"
            Case "": s = s & vbLf & "NG:アクション（allow / deny / drop 等）を選択してください"
            Case Else: s = s & vbLf & "NG:アクションは allow / deny / drop / reset-client / reset-server / reset-both のいずれかです"
        End Select
        a = PSplit(CStr(v(FIdx("profile_setting"))))
        If UBound(a) >= 0 Then
            If UBound(a) > 0 Then s = s & vbLf & "NG:プロファイルグループは1つだけ指定してください"
            If ColCount(CH_PG) > 0 Then
                If Not mRef.Exists("pg|" & CStr(a(0))) Then
                    warns = warns & vbLf & "警告:プロファイルグループ「" & a(0) & "」が②既存データにありません（大文字・小文字も区別）"
                End If
            End If
            If act <> "allow" Then warns = warns & vbLf & "警告:プロファイルグループは allow のルールでのみ効きます"
        ElseIf act = "allow" Then
            warns = warns & vbLf & "警告:プロファイルグループが空欄です（脅威検査なしで許可されます）"
        End If
    End If
    If Len(CStr(v(FIdx("description")))) > 1024 Then s = s & vbLf & "NG:説明は1024文字以内にしてください"
    ContentNg = s
End Function

Private Function NormFolder(ByVal f As String) As String
    Select Case LCase$(Trim$(f))
        Case "shared", "prisma access": NormFolder = "Shared"
        Case "all", "global", "グローバル": NormFolder = "All"
        Case "remote networks", "rn": NormFolder = "Remote Networks"
        Case "service connections", "sc": NormFolder = "Service Connections"
        Case "mobile users container", "muc": NormFolder = "Mobile Users Container"
        Case "mobile users", "globalprotect", "gp": NormFolder = "Mobile Users"
        Case "mobile users explicit proxy", "explicit proxy", "ep": NormFolder = "Mobile Users Explicit Proxy"
        Case Else: NormFolder = Trim$(f)
    End Select
End Function

' ルールの folder から見える folder（自分＋上位）か。スコープはカンマ区切り可、空欄はどこでも可
Private Function ScopeAllowed(ByVal scopes As String, ByVal ruleFolder As String) As Boolean
    Dim p As Variant, t As String, rf As String
    If Trim$(scopes) = "" Then ScopeAllowed = True: Exit Function
    rf = NormFolder(ruleFolder)
    If rf = "" Then ScopeAllowed = True: Exit Function
    For Each p In Split(scopes, ",")
        t = NormFolder(CStr(p))
        If t = "" Or t = "All" Or t = rf Then ScopeAllowed = True: Exit Function
        Select Case rf
            Case "Remote Networks", "Service Connections", "Mobile Users Container"
                If t = "Shared" Then ScopeAllowed = True: Exit Function
            Case "Mobile Users", "Mobile Users Explicit Proxy"
                If t = "Shared" Or t = "Mobile Users Container" Then ScopeAllowed = True: Exit Function
        End Select
    Next
End Function

Private Function BuildRefMap() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    RefAdd d, "zone", CH_ZONE, True
    RefAdd d, "addr", CH_ADDR, True
    RefAdd d, "addr", CH_ADDRG, True
    RefAdd d, "addr", CH_EDL, True
    RefAdd d, "addr", CH_REGION, False
    RefAdd d, "svc", CH_SVC, True
    RefAdd d, "svc", CH_SVCG, True
    RefAdd d, "url", CH_URLC, True
    RefAdd d, "url", CH_URLP, False
    RefAdd d, "url", CH_EDL, True
    RefAdd d, "tag", CH_TAG, True
    RefAdd d, "hip", CH_HIP, True
    RefAdd d, "user", CH_USER, False
    RefAdd d, "app", CH_APPP, False
    RefAdd d, "app", CH_APPC, True
    RefAdd d, "app", CH_APPG, True
    RefAdd d, "app", CH_APPF, True
    RefAdd d, "pg", CH_PG, True
    Set BuildRefMap = d
End Function

Private Sub RefAdd(ByVal d As Object, ByVal grp As String, ByVal col As Long, ByVal scoped As Boolean)
    Dim ws As Worksheet, r As Long, k As String, sc As String, nm As String
    Set ws = Worksheets(SH_CHO)
    For r = FIRST To LastRow(ws, col)
        nm = Trim$(SafeStr(ws.Cells(r, col).Value))
        If nm <> "" Then
            k = grp & "|" & nm
            If scoped Then sc = Trim$(SafeStr(ws.Cells(r, col + 1).Value)) Else sc = ""
            If sc = "" Then
                d(k) = "*"
            ElseIf Not d.Exists(k) Then
                d(k) = sc
            ElseIf d(k) <> "*" Then
                d(k) = d(k) & "," & sc
            End If
        End If
    Next
End Sub

' 参照チェック：③～⑪で削除予定でないか、選択肢にあるか、ルールの folder から見えるか
Private Function RefNg(ByVal v As Variant, ByVal folder As String, ByRef warns As String) As String
    Dim i As Long, grp As String, a As Variant, j As Long, nm As String, s As String, unk As String, sc As String, dsh As String
    If mRef Is Nothing Then BuildChoices
    For i = 0 To NF - 1
        Select Case mFK(i)
            Case "from", "to": grp = "zone"
            Case "source", "destination": grp = "addr"
            Case "service": grp = "svc"
            Case "category": grp = "url"
            Case "tag": grp = "tag"
            Case "source_hip", "destination_hip": grp = "hip"
            Case "application": grp = "app"
            Case "source_user": grp = "user"
            Case Else: grp = ""
        End Select
        If grp <> "" Then
            a = PSplit(CStr(v(i + 1)))
            unk = ""
            For j = 0 To UBound(a)
                nm = CStr(a(j))
                dsh = DelSheet(grp, nm)
                If LCase$(nm) = "any" Then
                ElseIf grp = "svc" And (nm = "application-default" Or nm = "service-http" Or nm = "service-https") Then
                ElseIf grp = "user" And (nm = "known-user" Or nm = "unknown" Or nm = "pre-logon") Then
                ElseIf grp = "addr" And IsIpLike(nm) Then
                ElseIf dsh <> "" Then
                    s = s & vbLf & "NG:" & mFL(i) & "「" & nm & "」は " & dsh & " で削除予定です"
                ElseIf Not mRef.Exists(grp & "|" & nm) Then
                    unk = unk & "、" & nm
                Else
                    sc = mRef(grp & "|" & nm)
                    If sc <> "*" And Not ScopeAllowed(sc, folder) Then
                        s = s & vbLf & "NG:" & mFL(i) & "「" & nm & "」は " & sc & " の定義のため、" & ScopeLabel(folder) & " のルールでは使えません"
                    End If
                End If
            Next
            If unk <> "" Then
                warns = warns & vbLf & "警告:" & mFL(i) & " の「" & Mid$(unk, 2) & "」が②既存データにも③～⑪の登録にもありません" & UnkHint(grp)
            End If
        End If
    Next
    RefNg = s
End Function

Private Function UnkHint(ByVal grp As String) As String
    Select Case grp
        Case "app": UnkHint = "（⑩に記入してください）"
        Case "user": UnkHint = "（⑪に記入してください）"
        Case "zone": UnkHint = "（参照_ポリシー選択肢にゾーンとスコープを追記してください）"
        Case "tag": UnkHint = "（参照_ポリシー選択肢にタグを追記してください）"
    End Select
End Function

' 削除予定のシート名（EDL はアドレス・URLカテゴリの欄で使うので両方見る）
Private Function DelSheet(ByVal grp As String, ByVal nm As String) As String
    If mDel Is Nothing Then Exit Function
    If mDel.Exists(grp & "|" & nm) Then DelSheet = mDel(grp & "|" & nm): Exit Function
    If grp = "addr" Or grp = "url" Then
        If mDel.Exists("edl|" & nm) Then DelSheet = mDel("edl|" & nm)
    End If
End Function

Private Function IsIpLike(ByVal s As String) As Boolean
    Dim c As String
    c = Left$(s, 1)
    If c Like "#" Then
        IsIpLike = (InStr(s, ".") > 0 Or InStr(s, "/") > 0)
    ElseIf InStr(s, "::") > 0 Or (InStr(s, ":") > 0 And Not s Like "*[g-zG-Z]*") Then
        IsIpLike = True
    End If
End Function

Private Function InCol(ByVal col As Long, ByVal nm As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = Worksheets(SH_CHO)
    For r = FIRST To LastRow(ws, col)
        If SafeStr(ws.Cells(r, col).Value) = nm Then InCol = True: Exit Function
    Next
End Function

Private Function ColCount(ByVal col As Long) As Long
    Dim ws As Worksheet, lr As Long
    Set ws = Worksheets(SH_CHO)
    lr = LastRow(ws, col)
    If lr >= FIRST Then ColCount = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(FIRST, col), ws.Cells(lr, col)))
End Function

'==========================================================
' 表示（スコープ名・見出し・段ずらし）
'==========================================================
Private Function ScopeLabel(ByVal folder As String) As String
    Select Case folder
        Case "All": ScopeLabel = "グローバル（All）※参照のみ"
        Case "Shared", "Prisma Access": ScopeLabel = FolderDisp("Shared")
        Case "Mobile Users": ScopeLabel = "GlobalProtect（Mobile Users）"
        Case "Mobile Users Explicit Proxy": ScopeLabel = "Explicit Proxy（Mobile Users Explicit Proxy）"
        Case Else: ScopeLabel = folder
    End Select
End Function

Private Function SingleRB(ByVal folder As String) As Boolean
    Select Case LCase$(folder)
        Case "mobile users", "mobile users explicit proxy", "remote networks", "service connections": SingleRB = True
    End Select
End Function

Private Function GroupLabel(ByVal folder As String, ByVal pos As String) As String
    If SingleRB(folder) Then
        GroupLabel = ScopeLabel(folder)
        Exit Function
    End If
    GroupLabel = ScopeLabel(folder) & " " & IIf(LCase$(pos) = "post", "【Post】", "【Pre】")
End Function

Private Function ScopeDepth(ByVal folder As String) As Long
    Select Case LCase$(folder)
        Case "all": ScopeDepth = 0
        Case "shared": ScopeDepth = 1
        Case "remote networks", "service connections", "mobile users container": ScopeDepth = 2
        Case "mobile users", "mobile users explicit proxy": ScopeDepth = 3
        Case Else: ScopeDepth = 1
    End Select
End Function

Private Function ScopeLvl(ByVal folder As String) As Long
    ScopeLvl = ScopeDepth(folder) - 1
    If ScopeLvl < 0 Then ScopeLvl = 0
End Function

Private Function DepthColor(ByVal d As Long) As Long
    Select Case d
        Case 0: DepthColor = &H404040
        Case 1: DepthColor = &H64381F
        Case 2: DepthColor = &H975A2F
        Case Else: DepthColor = &HC47244
    End Select
End Function

Private Sub IndentRow(ByVal ws As Worksheet, ByVal r As Long, ByVal scopeCol As Long, ByVal folder As String)
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

Private Sub PaintHeader(ByVal ws As Worksheet, ByVal r As Long)
    Dim f As String
    f = SafeStr(ws.Cells(r, S_HF).Value)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, S_LAST))
        .Interior.Color = DepthColor(ScopeDepth(f))
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = vbWhite
        .Font.Strikethrough = False
        .WrapText = False
        .Borders.LineStyle = xlNone
    End With
    ws.Range(ws.Cells(r, S_HF), ws.Cells(r, S_HF + 1)).Font.Color = vbWhite
    IndentRow ws, r, S_SCOPE, f
    ws.Cells(r, S_SCOPE + ScopeLvl(f)).Font.Bold = True
End Sub

Private Sub PaintBase(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, S_LAST))
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
    ws.Cells(r, S_OP).Interior.Color = C_IN
    ws.Range(ws.Cells(r, S_NAME), ws.Cells(r, S_LAST)).Interior.Color = C_IN
    ws.Cells(r, 1).Interior.Color = C_FX
    ws.Range(ws.Cells(r, S_JUDGE), ws.Cells(r, S_ID)).Interior.Color = C_FX
End Sub

'==========================================================
' API 取得（GET のみ。modInFetch.FetchExistingAll から呼ぶ）
'   ②-2／②-3：既存ツールと同じく folder × position ごとに GET → スコープごとに表化（継承分は定義元のスコープだけに出す）
'   戻り値：失敗した種類の数
'==========================================================
Public Function FetchPolicies(ByRef msg As String) As Long
    Dim m As Variant, groups As Collection, failN As Long, g As Variant, n As Long, lbl As String
    For Each m In Array("sec", "dec")
        SetMode CStr(m)
        lbl = IIf(m = "dec", "復号ポリシー", "セキュリティポリシー")
        failN = 0
        Set groups = FetchRuleGroups("既存取得（" & lbl & "）", failN)
        If groups Is Nothing Then
            FetchPolicies = FetchPolicies + 1
            msg = msg & vbLf & "　" & lbl & "：失敗（①の対象スコープに○がありません）"
        ElseIf groups.Count = 0 Then
            FetchPolicies = FetchPolicies + 1
            msg = msg & vbLf & "　" & lbl & "：失敗"
        Else
            WriteGetSheet groups
            n = 0
            For Each g In groups
                n = n + g(3).Count
            Next
            msg = msg & vbLf & "　" & lbl & "：" & n & " 件" & IIf(failN > 0, "（" & failN & " 件のスコープは取得失敗）", "")
        End If
    Next
End Function

Private Function Positions() As Variant
    Select Case LCase$(CfgOpt("CFG_POSITION"))
        Case "post": Positions = Array("post")
        Case "pre": Positions = Array("pre")
        Case Else: Positions = Array("pre", "post")
    End Select
End Function

' その folder・position を取得するか
Private Function NeedFetch(ByVal folder As String, ByVal pos As String) As Boolean
    Dim ps As Variant
    ' Pre/Post なしのスコープ（GlobalProtect・Explicit Proxy・Remote Networks・Service Connections）は pre だけ取得する
    '   （post で取得すると同じルールが返り、別スコープ扱いの NG になるため。Post は使わない）
    If SingleRB(folder) Then NeedFetch = (LCase$(pos) = "pre"): Exit Function
    For Each ps In Positions()
        If ps = pos Then NeedFetch = True
    Next
End Function

Private Function FetchRuleGroups(ByVal opLabel As String, ByRef failN As Long) As Collection
    Dim scopes As Collection, sc As Variant, ps As Variant, items As Collection, item As Variant, url As String
    Dim own As Object, seen As Object, allD As Object, c As Collection, groups As New Collection, fo As Variant
    Dim st As Long, m As String, f As String, k As String, inh As Long, q As String, key As String
    Set scopes = modInFetch.InSelScopes()
    If scopes.Count = 0 Then Exit Function
    url = modInFetch.InApiBase() & IIf(mMode = "dec", "/sse/config/v1/decryption-rules", "/sse/config/v1/security-rules")
    Set own = CreateObject("Scripting.Dictionary")
    Set allD = CreateObject("Scripting.Dictionary")
    ' グローバル（All）のルールは pre・post のどちらで取得しても同じものが返るため、id で重複を除く（先に見つかった pre に出す）
    Set seen = CreateObject("Scripting.Dictionary")
    For Each ps In Array("pre", "post")
        Set allD(CStr(ps)) = New Collection
        For Each sc In scopes
            If NeedFetch(CStr(sc(1)), CStr(ps)) Then
                Prog opLabel & "：" & sc(0) & "（" & ps & "）を取得中…"
                q = "folder=" & modInFetch.InEnc(CStr(sc(1))) & "&position=" & ps
                Set items = modInFetch.InGetAll(url, q, opLabel, st, m)
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
                    WriteLog opLabel, CStr(sc(0)) & "／" & ps, "GET", WithQuery(url, q), 200, _
                             "定義 " & c.Count & " 件（上位folderから継承 " & inh & " 件は定義元のスコープに表示）"
                End If
            End If
        Next
    Next
    For Each fo In FetchOrder(scopes)
        key = fo(0) & "|" & fo(1)
        If fo(0) = "All" Then
            If allD.Exists(CStr(fo(1))) Then
                If allD(CStr(fo(1))).Count > 0 Then groups.Add Array("All", CStr(fo(1)), GroupLabel("All", CStr(fo(1))), allD(CStr(fo(1))))
            End If
        ElseIf own.Exists(key) Then
            If Not (SingleRB(CStr(fo(0))) And fo(1) = "post" And own(key).Count = 0) Then
                groups.Add Array(CStr(fo(0)), CStr(fo(1)), GroupLabel(CStr(fo(0)), CStr(fo(1))), own(key))
            End If
        End If
    Next
    Set FetchRuleGroups = groups
End Function

' 取得結果の並び（セキュリティ・復号とも同じ評価順）。All は応答に含まれるときだけ
Private Function FetchOrder(ByVal scopes As Collection) As Collection
    Dim res As New Collection, sel As Object, sc As Variant, usePre As Boolean, usePost As Boolean, ps As Variant, x As Variant
    For Each ps In Positions()
        If ps = "pre" Then usePre = True Else usePost = True
    Next
    Set sel = CreateObject("Scripting.Dictionary")
    sel("All") = True
    For Each sc In scopes
        sel(CStr(sc(1))) = True
    Next
    For Each x In Array(Array("All", "pre"), Array("Shared", "pre"), Array("Mobile Users Container", "pre"), _
                        Array("Mobile Users", "pre"), Array("Mobile Users Explicit Proxy", "pre"), _
                        Array("Mobile Users Container", "post"), _
                        Array("Remote Networks", "pre"), Array("Service Connections", "pre"), _
                        Array("Shared", "post"), Array("All", "post"))
        If sel.Exists(CStr(x(0))) Then
            If SingleRB(CStr(x(0))) Then
                res.Add x
            ElseIf (x(1) = "pre" And usePre) Or (x(1) = "post" And usePost) Then
                res.Add x
            End If
        End If
    Next
    Set FetchOrder = res
End Function

' APIの1ルール → 値の配列（0=name, 1～NF=項目。配列はセル内改行区切り）
Private Function RuleVals(ByVal item As Object) As Variant
    Dim v() As String, i As Long, k As String, td As Object, tk As Variant
    ReDim v(0 To NF)
    v(0) = Nz(item, "name")
    For i = 0 To NF - 1
        k = mFK(i)
        Select Case mFT(i)
            Case "a": v(i + 1) = JoinLf(item, k)
            Case "b": v(i + 1) = BoolTxt(item, k)
            Case "s": v(i + 1) = Nz(item, k)
            Case "g"
                If item.Exists(k) Then
                    If IsObject(item(k)) Then
                        If TypeName(item(k)) <> "Collection" Then v(i + 1) = JoinLf(item(k), "group")
                    End If
                End If
            Case "t"
                If item.Exists("type") Then
                    If IsObject(item("type")) Then
                        Set td = item("type")
                        If TypeName(td) = "Dictionary" Then
                            For Each tk In td.Keys
                                v(i + 1) = CStr(tk)
                                If CStr(tk) = "ssl_inbound_inspection" Then
                                    If Not IsObject(td(tk)) Then v(i + 2) = SafeStr(td(tk))
                                End If
                            Next
                        End If
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

' ②-2／②-3 に書く（A=順番 B～D=スコープ E=folder F=position G=id H=name I～=項目 最後=未使用（非表示）、見出し行の folder/position は非表示列）
Private Sub WriteGetSheet(ByVal groups As Collection)
    Dim ws As Worksheet, g As Variant, item As Variant, r As Long, n As Long, v As Variant, a() As Variant, i As Long
    Set ws = Worksheets(mGet)
    With ws.Range(ws.Cells(FIRST, 1), ws.Cells(ws.Rows.Count, G_HF + 1))
        .ClearContents
        .ClearFormats
        .FormatConditions.Delete
    End With
    r = FIRST
    For Each g In groups
        ws.Range(ws.Cells(r, 1), ws.Cells(r, G_HF + 1)).NumberFormat = "@"
        ws.Cells(r, 1).Value = HDR
        ws.Cells(r, G_SCOPE).Value = "■ " & g(2) & " ／ " & g(3).Count & " 件"
        ws.Cells(r, G_HF).Value = g(0)
        ws.Cells(r, G_HF + 1).Value = g(1)
        PaintHeaderG ws, r
        r = r + 1
        n = 0
        For Each item In g(3)
            n = n + 1
            v = RuleVals(item)
            ReDim a(1 To 1, 1 To G_LAST)
            a(1, 1) = n
            a(1, G_SCOPE) = g(2)
            a(1, G_FOLDER) = FolderDisp(CStr(g(0)))
            a(1, G_FOLDER + 1) = g(1)
            a(1, G_ID) = Nz(item, "id")
            For i = 0 To NF
                a(1, G_NAME + i) = v(i)
            Next
            ws.Range(ws.Cells(r, 2), ws.Cells(r, G_LAST)).NumberFormat = "@"
            ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST)).Value = a
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST))
                .Font.Name = "Meiryo UI"
                .Font.Size = 10
                .WrapText = True
                .VerticalAlignment = xlTop
                .Borders.LineStyle = xlContinuous
                .Borders.Color = &HBFBFBF
            End With
            If n = 1 Then ws.Cells(r, G_SCOPE).Font.Bold = True
            IndentRow ws, r, G_SCOPE, CStr(g(0))
            r = r + 1
        Next
    Next
    ws.Range("B3").Value = Now
End Sub

Private Sub PaintHeaderG(ByVal ws As Worksheet, ByVal r As Long)
    Dim f As String
    f = SafeStr(ws.Cells(r, G_HF).Value)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST))
        .Interior.Color = DepthColor(ScopeDepth(f))
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = vbWhite
        .WrapText = False
        .Borders.LineStyle = xlNone
    End With
    ws.Range(ws.Cells(r, G_HF), ws.Cells(r, G_HF + 1)).Font.Color = vbWhite
    IndentRow ws, r, G_SCOPE, f
    ws.Cells(r, G_SCOPE + ScopeLvl(f)).Font.Bold = True
End Sub

'==========================================================
' 参照_ポリシー選択肢 の API 取得（タグ・アプリグループ・アプリフィルタ・リージョン・事前定義URL・自作App・証明書・復号プロファイル）
'   アドレス等のオブジェクトの列は ② から作る（最後に BuildChoices）。ゾーンは取得したルールで使われているものを追加
'   戻り値：失敗した種類の数
'==========================================================
Public Function FetchPolicyChoices(ByRef msg As String) As Long
    Dim ng As Long, items As Collection, item As Variant, vals As Collection, scs As Collection, nm As String
    Dim base As String, q As String, st As Long, m As String, custom As Object, ws2 As Worksheet, r As Long, lastR As Long
    base = FolderApi(CfgOpt("CFG_OBJ_FOLDER"))
    If base = "" Then base = "Shared"
    q = "folder=" & modInFetch.InEnc(base)
    If Not ChoiceCol(CH_TAG, "/sse/config/v1/tags") Then ng = ng + 1
    If Not ChoiceCol(CH_APPG, "/sse/config/v1/application-groups") Then ng = ng + 1
    If Not ChoiceCol(CH_APPF, "/sse/config/v1/application-filters") Then ng = ng + 1
    If Not ChoiceCol(CH_DECP, "/sse/config/v1/decryption-profiles", False) Then ng = ng + 1
    If Not ChoiceCol(CH_CERT, "/sse/config/v1/certificates", False) Then ng = ng + 1
    ' リージョン（スコープなし）
    Prog "選択肢：リージョンを取得中…"
    Set items = modInFetch.InGetAll(modInFetch.InApiBase() & "/sse/config/v1/regions", q, "選択肢取得（リージョン）", st, m)
    If items Is Nothing Then
        ng = ng + 1
    Else
        Set vals = New Collection
        For Each item In items
            If Nz(item, "name") <> "" Then vals.Add Nz(item, "name")
        Next
        WriteChoiceCol CH_REGION, vals
    End If
    ' 事前定義URLカテゴリ（カスタムURLカテゴリ＝②の一覧 は除く。403 は権限なしとして今の一覧のまま）
    Set custom = CreateObject("Scripting.Dictionary")
    Set ws2 = Worksheets(S2)
    lastR = LastRow(ws2, OFF_URL + 2)
    For r = FIRST To lastR
        If SafeStr(ws2.Cells(r, OFF_URL + 2).Value) <> "" Then custom(SafeStr(ws2.Cells(r, OFF_URL + 2).Value)) = True
    Next
    Prog "選択肢：事前定義URLカテゴリを取得中…"
    Set items = modInFetch.InGetAll(modInFetch.InApiBase() & "/sse/config/v1/url-filtering-categories", q, "選択肢取得（事前定義URL）", st, m)
    If items Is Nothing Then
        If st <> 403 Then ng = ng + 1
    Else
        Set vals = New Collection
        For Each item In items
            nm = Nz(item, "value")
            If nm = "" Then nm = Nz(item, "name")
            If nm <> "" And Not custom.Exists(nm) And InStr(LCase$(Nz(item, "type")), "custom") = 0 Then vals.Add nm
        Next
        If vals.Count > 0 Then WriteChoiceCol CH_URLP, vals
    End If
    ' 自作App（定義済み約11万件は読み飛ばす。オブジェクト取得folderのみ）
    If CfgOpt("CFG_APP_FETCH") <> "しない" Then
        If CustomApps(q, base, vals, scs) Then
            WriteChoiceCol CH_APPC, vals
            WriteChoiceCol CH_APPC + 1, scs
            msg = msg & vbLf & "　自作アプリケーション：" & vals.Count & " 件"
        Else
            ng = ng + 1
        End If
    End If
    AddUsedZones
    BuildChoices
    Worksheets(SH_CHO).Range("B3").Value = Now
    msg = msg & vbLf & "　ポリシーの選択肢（タグ・リージョン・アプリ・証明書・復号プロファイル・ゾーン）" & IIf(ng > 0, "：" & ng & " 種類は失敗", "")
    FetchPolicyChoices = ng
End Function

' 1種類を ○のスコープ等から取得して 名前列 と スコープ列（右隣）に書く。withScope=False なら名前列だけ
Private Function ChoiceCol(ByVal col As Long, ByVal ep As String, Optional ByVal withScope As Boolean = True) As Boolean
    Dim items As Collection, item As Variant, vals As New Collection, scs As New Collection, fd As String
    Set items = modInFetch.InGetFolders(ep, "選択肢取得（" & Mid$(ep, InStrRev(ep, "/") + 1) & "）")
    If items Is Nothing Then Exit Function
    For Each item In items
        If Nz(item, "name") <> "" Then
            fd = Nz(item, "folder")
            If fd = "" Then fd = Nz(item, "snippet")
            vals.Add Nz(item, "name"): scs.Add FolderDisp(fd)
        End If
    Next
    WriteChoiceCol col, vals
    If withScope Then WriteChoiceCol col + 1, scs
    ChoiceCol = True
End Function

Private Sub WriteChoiceCol(ByVal col As Long, ByVal vals As Collection)
    Dim ws As Worksheet, a() As Variant, i As Long, n As Long
    Set ws = Worksheets(SH_CHO)
    ws.Range(ws.Cells(FIRST, col), ws.Cells(CHO_LAST, col)).ClearContents
    n = vals.Count
    If n = 0 Then Exit Sub
    If n > CHO_LAST - FIRST + 1 Then n = CHO_LAST - FIRST + 1
    ReDim a(1 To n, 1 To 1)
    For i = 1 To n
        a(i, 1) = vals(i)
    Next
    ws.Cells(FIRST, col).Resize(n, 1).NumberFormat = "@"
    ws.Cells(FIRST, col).Resize(n, 1).Value = a
End Sub

Private Function CustomApps(ByVal q As String, ByVal folder As String, ByRef c As Collection, ByRef sc As Collection) As Boolean
    Dim url As String, resText As String, res As Object, item As Variant, st As Long
    Dim offset As Long, rawN As Long, lim As Long, total As Long, t0 As Single, fd As String
    Set c = New Collection: Set sc = New Collection
    lim = Val(CfgOpt("CFG_PAGE_LIMIT"))
    If lim <= 0 Then lim = 200
    t0 = Timer
    Do
        url = WithQuery(modInFetch.InApiBase() & "/sse/config/v1/applications", q & "&limit=" & lim & "&offset=" & offset)
        If Not EnsureToken() Then Exit Function
        st = HttpCall("GET", url, "", resText)
        If st = 401 Then
            If FetchToken() Then st = HttpCall("GET", url, "", resText)
        End If
        If st <> 200 Then
            WriteLog "選択肢取得（自作App）", "", "GET", url, st, ErrSummary(resText)
            Exit Function
        End If
        rawN = 0
        Set res = ParseJsonSkipPredef(resText, rawN)
        resText = ""
        If Not res Is Nothing Then
            If res.Exists("total") Then total = Val(SafeStr(res("total")))
            If res.Exists("data") Then
                If IsObject(res("data")) Then
                    For Each item In res("data")
                        If IsObject(item) Then
                            If Nz(item, "id") <> "" And InStr(LCase$(Nz(item, "folder") & " " & Nz(item, "snippet")), "predefined") = 0 And Nz(item, "name") <> "" Then
                                fd = Nz(item, "folder")
                                If fd = "" Then fd = folder
                                c.Add Nz(item, "name"): sc.Add FolderDisp(fd)
                            End If
                        End If
                    Next
                End If
            End If
        End If
        offset = offset + lim
        Prog "選択肢：applications を確認中 " & Format$(IIf(total > 0 And offset > total, total, offset), "#,##0") & IIf(total > 0, " / " & Format$(total, "#,##0"), "") & _
             " 件（自作App " & c.Count & " 件・経過 " & Format$(Timer - t0, "0") & " 秒）"
        DoEvents
        If rawN < lim Then Exit Do
        If total > 0 And offset >= total Then Exit Do
    Loop
    WriteLog "選択肢取得（自作App）", "", "GET", WithQuery(modInFetch.InApiBase() & "/sse/config/v1/applications", q), 200, _
             c.Count & " 件（定義済みを除く。確認 " & offset & " 件）"
    CustomApps = True
End Function

' ②-2・②-3 のルールで使われているゾーンのうち選択肢に無いものを、使われている folder 付きで追加
Private Sub AddUsedZones()
    Dim wsG As Worksheet, wsC As Worksheet, known As Object, used As Object, lastR As Long, arr As Variant, i As Long
    Dim fi As Variant, a As Variant, j As Long, z As String, f As String, r As Long, k As Variant, m As Variant
    Set wsC = Worksheets(SH_CHO)
    Set known = CreateObject("Scripting.Dictionary")
    For r = FIRST To LastRow(wsC, CH_ZONE)
        If Trim$(SafeStr(wsC.Cells(r, CH_ZONE).Value)) <> "" Then known(Trim$(SafeStr(wsC.Cells(r, CH_ZONE).Value))) = True
    Next
    Set used = CreateObject("Scripting.Dictionary")
    For Each m In Array("sec", "dec")
        SetMode CStr(m)
        Set wsG = Worksheets(mGet)
        lastR = UsedLast(wsG)
        If lastR >= FIRST Then
            arr = wsG.Range(wsG.Cells(FIRST, 1), wsG.Cells(lastR, G_LAST)).Value
            For i = 1 To UBound(arr, 1)
                f = SafeStr(arr(i, G_FOLDER))
                If SafeStr(arr(i, 1)) <> HDR And f <> "" Then
                    For Each fi In Array(FIdx("from"), FIdx("to"))
                        a = SplitMembers(SafeStr(arr(i, G_NAME + fi)))
                        For j = 0 To UBound(a)
                            z = CStr(a(j))
                            If LCase$(z) <> "any" And Not known.Exists(z) Then
                                If Not used.Exists(z) Then used(z) = ""
                                If InStr("," & used(z) & ",", "," & f & ",") = 0 Then used(z) = used(z) & IIf(used(z) = "", "", ",") & f
                            End If
                        Next
                    Next
                End If
            Next
        End If
    Next
    If used.Count = 0 Then Exit Sub
    r = LastRow(wsC, CH_ZONE) + 1
    If r < FIRST Then r = FIRST
    For Each k In used.Keys
        wsC.Range(wsC.Cells(r, CH_ZONE), wsC.Cells(r, CH_ZONE + 1)).NumberFormat = "@"
        wsC.Cells(r, CH_ZONE).Value = k
        wsC.Cells(r, CH_ZONE + 1).Value = used(k)
        r = r + 1
    Next
    WriteLog "選択肢取得", "ゾーン", "", "", 0, used.Count & " 件を②-2・②-3のルールから追加（スコープ=使われているfolder。必要なら手で直してください）"
End Sub

'==========================================================
' 一括チェック用：判定を集める（modScmInput.CheckAll から呼ぶ）
'   戻り値 Collection of Array(シート, 行, 列, 名前, 判定, 内容)
'==========================================================
Public Function PolicyResults() As Collection
    Dim c As New Collection, sh As Variant, ws As Worksheet, r As Long, lastR As Long, j As String, rs As String
    For Each sh In Array(SH_SSEC, SH_SDEC)
        Set ws = Worksheets(CStr(sh))
        lastR = UsedLast(ws)
        For r = FIRST To lastR
            j = SafeStr(ws.Cells(r, S_JUDGE).Value)
            If (j = "NG" Or j = "警告") And SafeStr(ws.Cells(r, 1).Value) <> HDR Then
                rs = SafeStr(ws.Cells(r, S_REASON).Value)
                rs = Replace(Replace(Replace(rs, "NG:", ""), "警告:", ""), vbLf, " ／ ")
                c.Add Array(CStr(sh), r, S_JUDGE, SafeStr(ws.Cells(r, S_NAME).Value), j, rs)
            End If
        Next
    Next
    Set PolicyResults = c
End Function

'==========================================================
' ダブルクリック選択（modScmInput.TryPick から呼ぶ）
'   ⑫・⑬ のタグ／ゾーン／アドレス／ユーザー／HIP／アプリケーション／サービス／URLカテゴリ
'   候補はそのルールの folder から見えるもの（自分と上位の folder）だけ
'==========================================================
Public Function PolicyTryPick(ByVal Target As Range) As Boolean
    Dim ttl As String, ws As Worksheet
    Set ws = Target.Worksheet
    If ws.Name = SH_SSEC Then
        SetMode "sec"
    ElseIf ws.Name = SH_SDEC Then
        SetMode "dec"
    Else
        Exit Function
    End If
    Set mNames = CreateObject("Scripting.Dictionary")
    If Not PickerRule(ws, Target, ttl) Then Exit Function
    PolicyTryPick = True
    ShowPicker Target, ttl
End Function

Private Function PickerRule(ByVal ws As Worksheet, ByVal c As Range, ByRef ttl As String) As Boolean
    Dim fi As Long, k As String, f As String
    If c.Row < FIRST Then Exit Function
    If SafeStr(ws.Cells(c.Row, 1).Value) = HDR Then Exit Function
    fi = c.Column - S_F0
    If fi < 0 Or fi > NF - 1 Then Exit Function
    If mFT(fi) <> "a" Then Exit Function
    k = mFK(fi)
    f = SafeStr(ws.Cells(c.Row, S_FOLDER).Value)
    If f = "" Then
        If HeaderAbove(ws, c.Row) > 0 Then f = SafeStr(ws.Cells(HeaderAbove(ws, c.Row), S_HF).Value)
    End If
    Application.ScreenUpdating = False
    BuildChoices
    Application.ScreenUpdating = True
    Select Case k
        Case "from", "to"
            AddFix "any"
            AddSrcS CH_ZONE, "ゾーン", f
        Case "source", "destination"
            AddFix "any"
            AddSrcS CH_ADDR, "アドレス", f
            AddSrcS CH_ADDRG, "アドレスグループ", f
            AddEdl "ip", f
            AddSrc CH_REGION, "リージョン"
        Case "source_user"
            AddFix "any"
            AddFix "known-user"
            AddFix "unknown"
            AddFix "pre-logon"
            AddSrc CH_USER, "ユーザー／グループ"
        Case "source_hip", "destination_hip"
            AddFix "any"
            AddSrcS CH_HIP, "HIPプロファイル", f
        Case "service"
            AddFix "any"
            If mMode <> "dec" Then AddFix "application-default"
            AddFix "service-http"
            AddFix "service-https"
            AddSrcS CH_SVC, "サービス", f
            AddSrcS CH_SVCG, "サービスグループ", f
        Case "category"
            AddFix "any"
            AddSrcS CH_URLC, "URLカテゴリ（カスタム）", f
            AddSrc CH_URLP, "URLカテゴリ（事前定義）"
            AddEdl "url", f
        Case "application"
            AddFix "any"
            AddSrcS CH_APPC, "自作App", f
            AddSrcS CH_APPG, "アプリグループ", f
            AddSrcS CH_APPF, "アプリフィルタ", f
            AddSrc CH_APPP, "定義済みApp-ID"
        Case "tag"
            AddSrcS CH_TAG, "タグ", f
    End Select
    AddUsed G_F0 + fi, f
    ttl = mFL(fi) & " の選択：" & SafeStr(ws.Cells(c.Row, S_NAME).Value) & "（" & ScopeLabel(f) & " から見えるもの）"
    PickerRule = True
End Function

Private Sub AddFix(ByVal nm As String)
    If Not mNames.Exists(nm) Then mNames.Add nm, "既定値"
End Sub

' 名前列＋右隣のスコープ列から、ルールの folder で見えるものだけ（③～⑪で削除予定のものは出さない）
Private Sub AddSrcS(ByVal col As Long, ByVal kind As String, ByVal folder As String)
    Dim ws As Worksheet, r As Long, nm As String, sc As String
    Set ws = Worksheets(SH_CHO)
    For r = FIRST To LastRow(ws, col)
        nm = Trim$(SafeStr(ws.Cells(r, col).Value))
        sc = Trim$(SafeStr(ws.Cells(r, col + 1).Value))
        If nm <> "" And Not mNames.Exists(nm) And Not IsDeleted(nm) Then
            If ScopeAllowed(sc, folder) Then mNames.Add nm, kind & "（" & IIf(sc = "", "どこでも", sc) & "）"
        End If
    Next
End Sub

Private Function IsDeleted(ByVal nm As String) As Boolean
    Dim g As Variant
    If mDel Is Nothing Then Exit Function
    For Each g In Array("addr", "edl", "svc", "url", "hip", "app", "user")
        If mDel.Exists(g & "|" & nm) Then IsDeleted = True: Exit Function
    Next
End Function

Private Sub AddEdl(ByVal want As String, ByVal folder As String)
    Dim ws As Worksheet, r As Long, nm As String, sc As String, ty As String
    Set ws = Worksheets(SH_CHO)
    For r = FIRST To LastRow(ws, CH_EDL)
        nm = Trim$(SafeStr(ws.Cells(r, CH_EDL).Value))
        sc = Trim$(SafeStr(ws.Cells(r, CH_EDL + 1).Value))
        ty = LCase$(SafeStr(ws.Cells(r, CH_EDLTYPE).Value))
        If nm <> "" And Not mNames.Exists(nm) And Not IsDeleted(nm) Then
            If InStr(ty, want) > 0 And ScopeAllowed(sc, folder) Then mNames.Add nm, "EDL（" & ty & "・" & IIf(sc = "", "どこでも", sc) & "）"
        End If
    Next
End Sub

Private Sub AddSrc(ByVal col As Long, ByVal kind As String)
    Dim ws As Worksheet, r As Long, nm As String
    Set ws = Worksheets(SH_CHO)
    For r = FIRST To LastRow(ws, col)
        nm = Trim$(SafeStr(ws.Cells(r, col).Value))
        If nm <> "" And Not mNames.Exists(nm) And Not IsDeleted(nm) Then mNames.Add nm, kind
    Next
End Sub

' 既存ポリシーで使われている値（ルールの folder から見える folder のルールのもの）
Private Sub AddUsed(ByVal col As Long, ByVal folder As String)
    Dim ws As Worksheet, r As Long, a As Variant, j As Long, nm As String
    Set ws = Worksheets(mGet)
    For r = FIRST To UsedLast(ws)
        If SafeStr(ws.Cells(r, 1).Value) <> HDR Then
            If ScopeAllowed(SafeStr(ws.Cells(r, G_FOLDER).Value), folder) Then
                a = PSplit(SafeStr(ws.Cells(r, col).Value))
                For j = 0 To UBound(a)
                    nm = CStr(a(j))
                    If nm <> "" And Not mNames.Exists(nm) And Not IsDeleted(nm) Then mNames.Add nm, "使用中（既存ポリシー）"
                Next
            End If
        End If
    Next
End Sub

Private Sub ShowPicker(ByVal Target As Range, ByVal ttl As String)
    Dim f As frmMemberPick, res As String
    Set f = New frmMemberPick
    f.Setup ttl, mNames, modScmInput.SplitItems(SafeStr(Target.Value)), vbLf     ' セル内改行区切りで書き戻す
    f.Show
    If f.Confirmed Then
        res = f.Result
        If Len(res) > CELL_MAX Then
            MsgBox "選択結果が Excel のセル上限（32,767文字）を超えるため、書き込みませんでした。件数を減らしてください。", vbExclamation
        Else
            Target.NumberFormat = "@"
            Target.Value = res
        End If
    End If
    Unload f
End Sub

'==========================================================
' 小さな関数
'==========================================================
' セル内改行・カンマ区切り → 配列（空は UBound=-1）
Private Function PSplit(ByVal s As String) As Variant
    Dim parts() As String, i As Long, n As Long, out() As String, t As String
    PSplit = Split("")
    parts = Split(Replace(Replace(s, vbCr, ""), vbLf, ","), ",")
    If UBound(parts) < 0 Then Exit Function
    ReDim out(0 To UBound(parts))
    n = -1
    For i = 0 To UBound(parts)
        t = Trim$(parts(i))
        If t <> "" Then
            n = n + 1
            out(n) = t
        End If
    Next
    If n < 0 Then Exit Function
    ReDim Preserve out(0 To n)
    PSplit = out
End Function

' 1～12列のうち一番下のデータ行（見出し行・ルール行のどちらかが入る列）
Private Function UsedLast(ByVal ws As Worksheet) As Long
    Dim c As Long, x As Long
    UsedLast = FIRST - 1
    For c = 1 To 12
        x = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If x > UsedLast Then UsedLast = x
    Next
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsObject(v) Then Exit Function
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

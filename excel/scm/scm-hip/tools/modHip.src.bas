Attribute VB_Name = "modHip"
'==========================================================
' modHip : SCM (Prisma Access Configuration API) HIPオブジェクト / HIPプロファイル 連携
'   前提 : modScmCommon（JSONパーサを含む）を同じブックにインポートしていること
'   メンバー選択画面（ダブルクリック）: clsPickerEvents をインポートし、frmMemberPick は空のユーザーフォームを挿入して frmMemberPick_コード.txt を貼り付け
'   ボタン : SetupButtons を1回実行すると各シートにボタンを配置
'==========================================================
' ※ 共通処理（認証・HTTP・ログ・差分色付け等）は modScmCommon にあります。
Option Explicit

Private Const SH_CFG As String = "基本設定"
Private Const SH_EXIST As String = "HipObj_既存一覧"
Private Const SH_INPUT As String = "HipObj_登録変更削除"
Private Const SH_EXISTP As String = "HipProfile_既存一覧"
Private Const SH_INPUTP As String = "HipProfile_登録変更削除"
Private Const SH_CHOICE As String = "HIP_選択肢"
Private Const SH_VENDOR As String = "HIP_ベンダー一覧"
Private Const VD_FIRST As Long = 6      ' HIP_ベンダー一覧 データ開始行（B=カテゴリ C=ベンダー D=製品 E=出所）
Private Const SH_JSON As String = "JSON_CURL"
Private Const SH_DIFF As String = "HipObj_差分確認"
Private Const SH_DIFFP As String = "HipProfile_差分確認"
Private Const SH_LOG As String = "実行ログ"

Private Const EX_FIRST As Long = 8      ' 既存一覧 データ開始行
Private Const IN_FIRST As Long = 8      ' 登録変更削除 データ開始行
Private Const IN_LAST As Long = 507     ' 登録変更削除 データ最終行
Private Const JS_FIRST As Long = 10     ' JSON_CURL データ開始行
Private Const DIFF_FIRST As Long = 8    ' 差分確認 データ開始行
Private Const PAGE_LIMIT As Long = 200
Private Const EP_PROFILE As String = "/sse/config/v1/hip-profiles"
Private Const EP_CERT As String = "/sse/config/v1/certificate-profiles"

' ②③ 共通：F列(6)～BM列(65) が description と条件の各項目（並びは FieldKeys と同じ）
Private Const F_FIRST As Long = 6
Private Const F_LAST As Long = 65
' ② HipObj_既存一覧 : B=id C=name D=コンテナ E=対象外の条件(mobile_device) F～BM=項目 BN=JSON
Private Const EX_CONT As Long = 4
Private Const EX_MOBILE As Long = 5
Private Const EX_JSON As Long = 66
' ③ HipObj_登録変更削除 : B=操作 C=name D=判定 E=理由 F～BM=項目 BN=既存有無 BO=既存id（D・E と BN 以降は関数）
Private Const IN_JUDGE As Long = 4
Private Const IN_ID As Long = 67
' ⑥ HipProfile_既存一覧 : B=id C=name D=description E=match F=参照HIPオブジェクト G=コンテナ
' ⑦ HipProfile_登録変更削除 : B=操作 C=name D=description E=組み方 F=対象HIPオブジェクト G=除外HIPオブジェクト
'                            H=match直接記述 I=生成match(関数) K=既存id P=判定
Private Const ING_ID As Long = 11
Private Const ING_MATCH As Long = 9
Private Const ING_JUDGE As Long = 16
Private Const MODE_AND As String = "すべて満たす(and)"
Private Const MODE_OR As String = "いずれか(or)"
Private Const MODE_RAW As String = "直接記述"

' ---- メンバー選択画面 用 ----
Private Const PICK_KEY As String = "^+m"            ' Ctrl+Shift+M
Private Const CELL_MAX As Long = 32767               ' Excelのセル文字数上限
Private Const SH_OVERFLOW As String = "ピッカー_選択結果"

Private mEvents As clsPickerEvents
Private mNames As Object     ' 候補 name -> 種別ラベル（追加順を保持）
Private mExclude As String   ' 候補から外す name
Private mVendor As Boolean   ' ベンダー欄のピッカー（「ベンダー | 製品」で選び、ベンダーごとにまとめて書き戻す）

'----------------------------------------------------------
' ボタン配置（初回のみ実行）
'----------------------------------------------------------
Public Sub SetupButtons()
    AddBtn SH_CFG, "D18:E18", "トークン取得", "GetAccessToken"
    AddBtn SH_EXIST, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUT, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJson"
    AddBtn SH_INPUT, "I3:J4", "変更行に既存値を読込", "LoadExistingObj"
    AddBtn SH_INPUT, "L3:M4", "入力クリア", "ClearInput"
    AddBtn SH_EXISTP, "D3:E4", "既存情報取得（API）", "FetchExisting"
    AddBtn SH_INPUTP, "F3:G4", "JSON発行（③⑦）→ ④へ", "GenerateJsonProfile"
    AddBtn SH_INPUTP, "I3:J4", "変更行に既存値を読込", "LoadExistingProfile"
    AddBtn SH_INPUTP, "L3:M4", "入力クリア", "ClearInputProfile"
    AddBtn SH_JSON, "G4:H4", "API実行（全行）", "ExecuteAll"
    AddBtn SH_JSON, "G6:H6", "クリア", "ClearJson"
    RemoveBtn SH_JSON, "PushConfig"     ' Push機能は削除（PushはSCM画面から実施）。旧ボタンを消す
    AddBtn SH_DIFF, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_DIFFP, "J3:K4", "差分チェック（再作成）", "BuildDiffCheck"
    AddBtn SH_LOG, "F3", "ログクリア", "ClearLog"
    RemoveBtn SH_LOG, "CheckJobs"      ' ジョブ状況確認も削除。旧ボタンを消す
    PickerStart True
    MsgBox "ボタンを配置しました。", vbInformation
End Sub

' ③②の F～BM 列の項目キー（並び順はシートの列と同じ）
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

' 差分確認で比べる単位（description と条件のカテゴリ）
Private Function DiffCats() As Variant
    DiffCats = Array("description", "host_info", "anti_malware", "disk_backup", "disk_encryption", "firewall", _
                     "patch_management", "data_loss_prevention", "certificate", "network_info", "custom_checks", "mobile_device")
End Function

'----------------------------------------------------------
' ②⑥ 既存情報取得（HIPオブジェクトとHIPプロファイルの両方を GET → 表生成）
'   証明書プロファイル名も取得して HIP_選択肢 のプルダウン候補にする（失敗しても続行）
'----------------------------------------------------------
Public Sub FetchExisting()
    Dim objs As Collection, profs As Collection, certs As Collection
    If Not CheckContainer() Then Exit Sub
    If Not EnsureToken() Then Exit Sub
    Prog "HIPオブジェクト取得中..."
    Set objs = FetchAllItems(BaseUrl(), "既存取得", True, False)
    If objs Is Nothing Then GoTo Done
    Prog "HIPプロファイル取得中..."
    Set profs = FetchAllItems(EpUrl(EP_PROFILE), "既存Profile取得", True, False)
    If profs Is Nothing Then GoTo Done
    Prog "証明書プロファイル取得中..."
    Set certs = FetchAllItems(EpUrl(EP_CERT), "証明書プロファイル取得", True, True)
    Application.ScreenUpdating = False
    WriteExistObj objs
    MergeVendors objs
    WriteExistProfile profs
    If Not certs Is Nothing Then WriteCertList certs
    Application.ScreenUpdating = True
    MsgBox "HIPオブジェクト " & objs.Count & " 件、HIPプロファイル " & profs.Count & " 件を取得しました。" & _
           IIf(certs Is Nothing, vbLf & "（証明書プロファイル名は取得できませんでした。実行ログを確認してください）", ""), vbInformation
Done:
    Application.StatusBar = False
End Sub

Private Sub WriteExistObj(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long, d As Object, keys As Variant, i As Long, js As String
    Set ws = Worksheets(SH_EXIST)
    keys = FieldKeys()
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, EX_JSON)).Clear
    r = EX_FIRST
    For Each item In items
        Set d = ObjToRow(item)
        ws.Range(ws.Cells(r, 2), ws.Cells(r, EX_JSON)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        For i = 0 To UBound(keys)
            If d.Exists(keys(i)) Then ws.Cells(r, F_FIRST + i).Value = d(keys(i))
        Next
        ws.Cells(r, EX_MOBILE).Value = FlatText(Child(item, "mobile_device"))
        ws.Cells(r, EX_CONT).Value = ContainerOf(item)
        js = JsonOf(item)
        If Len(js) > CELL_MAX Then
            WriteLog "既存取得", Nz(item, "name"), "", "", 0, "JSONがセルの上限を超えたため切り詰めました（変更時の引き継ぎ・差分確認が正しくできません）"
            js = Left$(js, CELL_MAX)
        End If
        ws.Cells(r, EX_JSON).Value = js
        r = r + 1
    Next
    If r > EX_FIRST Then
        With ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, EX_JSON))
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlTop
        End With
        ws.Range(ws.Cells(EX_FIRST, EX_MOBILE), ws.Cells(r - 1, F_LAST)).WrapText = True
    End If
    Range("EX_FETCHED_AT").Value = Now
End Sub

Private Sub WriteExistProfile(ByVal items As Collection)
    Dim ws As Worksheet, item As Object, r As Long
    Set ws = Worksheets(SH_EXISTP)
    ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(ws.Rows.Count, 7)).Clear
    r = EX_FIRST
    For Each item In items
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).NumberFormat = "@"
        ws.Cells(r, 1).Formula = "=IF(C" & r & "="""","""",ROW()-7)"
        ws.Cells(r, 2).Value = Nz(item, "id")
        ws.Cells(r, 3).Value = Nz(item, "name")
        ws.Cells(r, 4).Value = Nz(item, "description")
        ws.Cells(r, 5).Value = Nz(item, "match")
        ws.Cells(r, 6).Value = JoinColl(QuotedNames(Nz(item, "match")), ",")
        ws.Cells(r, 7).Value = ContainerOf(item)
        r = r + 1
    Next
    If r > EX_FIRST Then
        With ws.Range(ws.Cells(EX_FIRST, 1), ws.Cells(r - 1, 7))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    Range("EXP_FETCHED_AT").Value = Now
End Sub

' 証明書プロファイル名 → HIP_選択肢 の CERT 列（プルダウン候補）
Private Sub WriteCertList(ByVal items As Collection)
    Dim c As Range, item As Object, r As Long
    Set c = Range("CERT_TOP")
    c.Worksheet.Range(c, c.Offset(1000, 0)).ClearContents
    r = 0
    For Each item In items
        c.Offset(r, 0).NumberFormat = "@"
        c.Offset(r, 0).Value = Nz(item, "name")
        r = r + 1
    Next
End Sub

'----------------------------------------------------------
' HIPオブジェクト JSON → 行（項目キー → 表示文字列）
'----------------------------------------------------------
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

' {not_available:{}} / {within:{days:n}} / {not_within:{hours:n}} → 条件・単位・数値
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

' vendor 配列 → 「ベンダー | 製品1 ; 製品2」を1行1ベンダー
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

'----------------------------------------------------------
' 行（項目キー → 文字列）→ HIPオブジェクト JSON（Dictionary）
'   raw : 変更時の既存JSON（このツールで編集しない mobile_device を引き継ぐ）
'   errs : 書式エラーを追記（空ならOK）
'----------------------------------------------------------
Private Function RowToObj(ByVal nm As String, ByVal d As Object, ByVal raw As Object, ByRef errs As String) As Object
    Dim o As Object, c As Object, x As Object, cat As Variant, p As String, ln As Variant, a As Variant
    Dim neg As Boolean, s As String, t As String, list As Collection, k As Variant
    Set o = NewD()
    o("name") = nm
    If d("description") <> "" Then o("description") = d("description")

    ' ホスト情報
    If AnySet(d, "hi_") Then
        Set c = NewD()
        If d("hi_os_vendor") <> "" Then
            Set c("os") = D1("contains", D1(d("hi_os_vendor"), IIf(d("hi_os_value") = "", "All", d("hi_os_value"))))
        End If
        For Each k In Array("domain|hi_domain", "host_name|hi_hostname", "client_version|hi_clientver", _
                            "serial_number|hi_serial", "host_id|hi_hostid")
            p = Split(k, "|")(1)
            If d(p & "_op") <> "" Then Set c(Split(k, "|")(0)) = D1(d(p & "_op"), d(p & "_val"))
        Next
        If d("hi_managed") <> "" Then c("managed") = (d("hi_managed") = "yes")
        Set o("host_info") = D1("criteria", c)
    End If

    ' ベンダー系カテゴリ
    For Each cat In Array("anti_malware|am", "disk_backup|db", "disk_encryption|de", "firewall|fw", _
                          "patch_management|pm", "data_loss_prevention|dlp")
        p = Split(cat, "|")(1)
        If AnySet(d, p & "_") Then
            Set c = NewD()
            If d(p & "_installed") <> "" Then c("is_installed") = (d(p & "_installed") = "yes")
            If d.Exists(p & "_enabled") Then
                If d(p & "_enabled") <> "" Then c("is_enabled") = d(p & "_enabled")
            End If
            Select Case p
                Case "am"
                    If d("am_rtp") <> "" Then c("real_time_protection") = d("am_rtp")
                    If d("am_virdef_op") <> "" Then Set c("virdef_version") = TimeObj(d, "am_virdef", errs, "定義ファイル")
                    If d("am_scan_op") <> "" Then Set c("last_scan_time") = TimeObj(d, "am_scan", errs, "最終スキャン")
                    If d("am_pv_op") = "within" Or d("am_pv_op") = "not_within" Then
                        Set c("product_version") = D1(d("am_pv_op"), D1("versions", NumVal(d("am_pv_val"), 1, 65535, "製品バージョン", errs)))
                    ElseIf d("am_pv_op") <> "" Then
                        Set c("product_version") = D1(d("am_pv_op"), d("am_pv_val"))
                    End If
                Case "db"
                    If d("db_time_op") <> "" Then Set c("last_backup_time") = TimeObj(d, "db_time", errs, "最終バックアップ")
                Case "de"
                    If d("de_locations") <> "" Then
                        Set list = New Collection
                        For Each ln In Lines(d("de_locations"))
                            t = ln
                            neg = TakeNeg(t)
                            a = Parts(t)
                            Set x = NewD()
                            x("name") = a(0)
                            If a(0) = "" Then errs = AddLine(errs, "暗号化場所：場所名が空の行があります")
                            s = "encrypted"
                            If UBound(a) >= 1 Then If a(1) <> "" Then s = a(1)
                            If InStr(",encrypted,unencrypted,partial,unknown,", "," & s & ",") = 0 Then
                                errs = AddLine(errs, "暗号化場所：状態は encrypted / unencrypted / partial / unknown（" & ln & "）")
                            End If
                            Set x("encryption_state") = D1(IIf(neg, "is_not", "is"), s)
                            list.Add x
                        Next
                        Set c("encrypted_locations") = list
                    End If
                Case "pm"
                    If d("pm_check") <> "" Or d("pm_sev_op") <> "" Or d("pm_patches") <> "" Then
                        Set x = NewD()
                        x("check") = IIf(d("pm_check") = "", "has-any", d("pm_check"))
                        If d("pm_sev_op") <> "" Then Set x("severity") = D1(d("pm_sev_op"), NumVal(d("pm_sev_val"), 0, 100000, "重要度", errs))
                        If d("pm_patches") <> "" Then Set x("patches") = Lines(d("pm_patches"))
                        Set c("missing_patches") = x
                    End If
            End Select
            Set x = NewD()
            If c.Count > 0 Then Set x("criteria") = c
            If d(p & "_vendor") <> "" Then Set x("vendor") = VendorObj(d(p & "_vendor"), errs)
            If d(p & "_excl") <> "" Then x("exclude_vendor") = (d(p & "_excl") = "yes")
            Set o(Split(cat, "|")(0)) = x
        End If
    Next

    ' 証明書
    If AnySet(d, "ce_") Then
        Set c = NewD()
        If d("ce_profile") <> "" Then c("certificate_profile") = d("ce_profile")
        If d("ce_attrs") <> "" Then
            Set list = New Collection
            For Each ln In Lines(d("ce_attrs"))
                a = Parts(ln)
                Set x = NewD()
                x("name") = a(0)
                If UBound(a) >= 1 Then If a(1) <> "" Then x("value") = a(1)
                list.Add x
            Next
            Set c("certificate_attributes") = list
        End If
        Set o("certificate") = D1("criteria", c)
    End If

    ' ネットワーク
    If AnySet(d, "nw_") Then
        Set x = NewD()
        If d("nw_type") = "wifi" And d("nw_val") <> "" Then x("ssid") = d("nw_val")
        If d("nw_type") = "mobile" And d("nw_val") <> "" Then x("carrier") = d("nw_val")
        Set c = D1(IIf(d("nw_op") = "", "is", d("nw_op")), D1(d("nw_type"), x))
        Set o("network_info") = D1("criteria", D1("network", c))
    End If

    ' カスタムチェック
    If AnySet(d, "cc_") Then
        Set c = NewD()
        If d("cc_process") <> "" Then
            Set list = New Collection
            For Each ln In Lines(d("cc_process"))
                t = ln
                neg = TakeNeg(t)
                Set x = NewD()
                x("name") = t
                x("running") = Not neg
                list.Add x
            Next
            Set c("process_list") = list
        End If
        If d("cc_registry") <> "" Then Set c("registry_key") = KeyList(d("cc_registry"), True, errs)
        If d("cc_plist") <> "" Then Set c("plist") = KeyList(d("cc_plist"), False, errs)
        Set o("custom_checks") = D1("criteria", c)
    End If

    ' このツールで編集しない条件（モバイルデバイス）は既存の内容を引き継ぐ
    If Not raw Is Nothing Then
        If raw.Exists("mobile_device") Then
            If IsObject(raw("mobile_device")) Then Set o("mobile_device") = raw("mobile_device")
        End If
    End If
    Set RowToObj = o
End Function

' 条件・単位・数値 → {not_available:{}} / {within:{days:n}}
Private Function TimeObj(ByVal d As Object, ByVal k As String, ByRef errs As String, ByVal label As String) As Object
    Dim op As String
    op = d(k & "_op")
    If op = "not_available" Then
        Set TimeObj = D1(op, NewD())
    Else
        If d(k & "_unit") = "" Then errs = AddLine(errs, label & "：単位が未選択")
        Set TimeObj = D1(op, D1(IIf(d(k & "_unit") = "", "days", d(k & "_unit")), NumVal(d(k & "_n"), 1, 65535, label, errs)))
    End If
End Function

' 「ベンダー | 製品1 ; 製品2」の行 → vendor 配列
Private Function VendorObj(ByVal s As String, ByRef errs As String) As Collection
    Dim list As New Collection, ln As Variant, a As Variant, x As Object, pr As Collection, q As Variant
    For Each ln In Lines(s)
        a = Parts(ln)
        Set x = NewD()
        x("name") = a(0)
        If a(0) = "" Then errs = AddLine(errs, "ベンダー：ベンダー名が空の行があります")
        If Len(a(0)) > 103 Then errs = AddLine(errs, "ベンダー：ベンダー名が103文字超（" & a(0) & "）")
        If UBound(a) >= 1 Then
            Set pr = New Collection
            For Each q In Split(a(1), ";")
                If Trim$(q) <> "" Then pr.Add Trim$(q)
            Next
            If pr.Count > 0 Then Set x("product") = pr
        End If
        list.Add x
    Next
    Set VendorObj = list
End Function

' レジストリ／plist の行 → registry_key / plist 配列（同じキー名の行は1つにまとめる）
'   キーだけ    : [!]キー
'   値の条件    : キー | [!]値名 | データ
'   既定値(reg) : キー | (既定) | データ
Private Function KeyList(ByVal s As String, ByVal isReg As Boolean, ByRef errs As String) As Collection
    Dim list As New Collection, idx As Object, ln As Variant, a As Variant, neg As Boolean, nm As String
    Dim kd As Object, vd As Object, vn As String, subKey As String, valKey As String
    Set idx = CreateObject("Scripting.Dictionary")
    subKey = IIf(isReg, "registry_value", "key")
    valKey = IIf(isReg, "value_data", "value")
    For Each ln In Lines(s)
        a = Parts(ln)
        nm = a(0)
        neg = TakeNeg(nm)
        If nm = "" Then
            errs = AddLine(errs, IIf(isReg, "レジストリ", "plist") & "：キー名が空の行があります")
        Else
            If Not idx.Exists(nm) Then
                Set kd = NewD()
                kd("name") = nm
                idx.Add nm, kd
                list.Add kd
            End If
            Set kd = idx(nm)
            If UBound(a) = 0 Then
                If neg Then kd("negate") = True
            ElseIf isReg And a(1) = "(既定)" Then
                If UBound(a) >= 2 Then kd("default_value_data") = a(2) Else kd("default_value_data") = ""
            Else
                vn = a(1)
                Set vd = NewD()
                If TakeNeg(vn) Then vd("negate") = True
                vd("name") = vn
                If vn = "" Then errs = AddLine(errs, IIf(isReg, "レジストリ", "plist") & "：値名が空（" & ln & "）")
                If UBound(a) >= 2 Then If a(2) <> "" Then vd(valKey) = a(2)
                If Not kd.Exists(subKey) Then Set kd(subKey) = New Collection
                kd(subKey).Add vd
            End If
        End If
    Next
    Set KeyList = list
End Function

' 数値チェック（範囲外・非数値は errs に追記して 0）
Private Function NumVal(ByVal s As String, ByVal mn As Long, ByVal mx As Long, ByVal label As String, ByRef errs As String) As Long
    If s = "" Or Not IsNumeric(s) Then
        errs = AddLine(errs, label & "：数値が未入力または数値以外（" & s & "）")
        Exit Function
    End If
    If CDbl(s) < mn Or CDbl(s) > mx Or CDbl(s) <> Int(CDbl(s)) Then
        errs = AddLine(errs, label & "：" & mn & "～" & mx & " の整数（" & s & "）")
        Exit Function
    End If
    NumVal = CLng(s)
End Function

'----------------------------------------------------------
' ③⑦ 変更行に既存値を読込（入力欄が空の 変更 行だけ。既に入力がある行はそのまま）
'----------------------------------------------------------
Public Sub LoadExistingObj()
    Dim ws As Worksheet, wsE As Worksheet, tgt As Collection, missing As String, filled As Long, mode As Long
    Dim it As Variant, r As Long, eRow As Long, c As Long, n As Long
    Set ws = Worksheets(SH_INPUT)
    Set wsE = Worksheets(SH_EXIST)
    Set tgt = ChangeRows(ws, wsE, F_FIRST, F_LAST, missing, filled)
    If tgt Is Nothing Then Exit Sub
    mode = AskLoadMode(filled)
    If mode = 0 Then Exit Sub
    For Each it In tgt
        r = it(0): eRow = it(1)
        If mode = 1 Then
            ws.Range(ws.Cells(r, F_FIRST), ws.Cells(r, F_LAST)).Value = _
                wsE.Range(wsE.Cells(eRow, F_FIRST), wsE.Cells(eRow, F_LAST)).Value
        Else
            For c = F_FIRST To F_LAST
                If SafeStr(ws.Cells(r, c).Value) = "" Then ws.Cells(r, c).Value = SafeStr(wsE.Cells(eRow, c).Value)
            Next
        End If
        n = n + 1
    Next
    LoadDoneMsg n, missing, ""
End Sub

Public Sub LoadExistingProfile()
    Dim ws As Worksheet, wsE As Worksheet, tgt As Collection, missing As String, filled As Long, mode As Long
    Dim it As Variant, r As Long, eRow As Long, n As Long, m As String, pos As Collection, ng As Collection, md As String
    Set ws = Worksheets(SH_INPUTP)
    Set wsE = Worksheets(SH_EXISTP)
    Set tgt = ChangeRows(ws, wsE, 4, 8, missing, filled)
    If tgt Is Nothing Then Exit Sub
    mode = AskLoadMode(filled)
    If mode = 0 Then Exit Sub
    For Each it In tgt
        r = it(0): eRow = it(1)
        If mode = 1 Then ws.Range(ws.Cells(r, 4), ws.Cells(r, 8)).ClearContents
        If SafeStr(ws.Cells(r, 4).Value) = "" Then ws.Cells(r, 4).Value = SafeStr(wsE.Cells(eRow, 4).Value)
        ' match は組み方～直接記述（E～H列）が空のときだけ展開する
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 5), ws.Cells(r, 8))) = 0 Then
            m = SafeStr(wsE.Cells(eRow, 5).Value)
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

' 操作=変更 の行と、②⑥での行番号の組 Array(入力行, 既存行) の一覧。
'   missing : 既存一覧に無い name、filled : 入力欄（c1～c2列）に入力済みの行数
'   変更の行が無ければメッセージを出して Nothing
Private Function ChangeRows(ByVal ws As Worksheet, ByVal wsE As Worksheet, ByVal c1 As Long, ByVal c2 As Long, _
                            ByRef missing As String, ByRef filled As Long) As Collection
    Dim res As New Collection, idx As Object, r As Long, nm As String, total As Long
    Set idx = CreateObject("Scripting.Dictionary")
    For r = EX_FIRST To LastRow(wsE, 3)
        nm = Trim$(SafeStr(wsE.Cells(r, 3).Value))
        If nm <> "" And Not idx.Exists(nm) Then idx.Add nm, r
    Next
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
        s = s & vbLf & vbLf & "次の name は既存一覧に無いため読み込めませんでした（既存一覧が古い場合は「既存情報取得」を実行してください）。" & missing
    End If
    MsgBox s, IIf(missing <> "", vbExclamation, vbInformation)
End Sub

' match が ⑦の組み方（and / or と not）で表せるなら分解する
Private Function ParseSimpleMatch(ByVal m As String, ByRef mode As String, ByRef pos As Collection, ByRef ng As Collection) As Boolean
    Dim names As Collection, i As Long, q1 As Long, q2 As Long, before As String, nm As String
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

' ⑦ I列の関数と同じ規則で match を作る
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

'----------------------------------------------------------
' ③⑦ JSON発行（HipObj_登録変更削除 + HipProfile_登録変更削除 → JSON_CURL）
'   どちらのボタンでも③と⑦の両方から④を作り直す。
'   実行順：HIPオブジェクト登録・変更 → HIPプロファイル登録・変更 → HIPプロファイル削除 → HIPオブジェクト削除
'----------------------------------------------------------
Public Sub GenerateJson()
    GenerateAll
End Sub

Public Sub GenerateJsonProfile()
    GenerateAll
End Sub

Private Sub GenerateAll()
    Dim wsA As Worksheet, wsP As Worksheet, wsE As Worksheet, wsOut As Worksheet
    Dim r As Long, o As Long, ng As Long, nm As String, op As String, e As String, errs As String, warns As String
    Dim bodies As Object, d As Object, raw As Object, exRow As Object, known As Object, delA As Object
    Dim names As Collection, v As Variant, posC As Collection, negC As Collection, seen As Object
    If Not CheckContainer() Then Exit Sub
    Application.Calculate
    Set wsA = Worksheets(SH_INPUT)
    Set wsP = Worksheets(SH_INPUTP)
    Set wsE = Worksheets(SH_EXIST)
    Set wsOut = Worksheets(SH_JSON)
    ng = Application.WorksheetFunction.CountIf(wsA.Range(wsA.Cells(IN_FIRST, IN_JUDGE), wsA.Cells(IN_LAST, IN_JUDGE)), "NG") + _
         Application.WorksheetFunction.CountIf(wsP.Range(wsP.Cells(IN_FIRST, ING_JUDGE), wsP.Cells(IN_LAST, ING_JUDGE)), "NG")
    If ng > 0 Then
        If MsgBox("③⑦に判定NGが " & ng & " 件あります。NG行を除外してJSONを発行しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    ' ②の行（name → 行番号）
    Set exRow = CreateObject("Scripting.Dictionary")
    For r = EX_FIRST To LastRow(wsE, 3)
        If SafeStr(wsE.Cells(r, 3).Value) <> "" Then exRow(SafeStr(wsE.Cells(r, 3).Value)) = r
    Next
    ' 使えるHIPオブジェクト名（②＋③で登録）と ③で削除するもの
    Set known = CreateObject("Scripting.Dictionary")
    Set delA = CreateObject("Scripting.Dictionary")
    For Each v In exRow.Keys
        known(v) = True
    Next

    ' ③：本文を作りながら書式を確認
    Set bodies = CreateObject("Scripting.Dictionary")
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) Then
            nm = SafeStr(wsA.Cells(r, 3).Value)
            op = SafeStr(wsA.Cells(r, 2).Value)
            If op = "削除" Then
                delA(nm) = True
            Else
                If op = "登録" Then known(nm) = True
                Set d = ReadRow(wsA, r)
                Set raw = Nothing
                If op = "変更" And exRow.Exists(nm) Then
                    If SafeStr(wsE.Cells(exRow(nm), EX_JSON).Value) <> "" Then Set raw = ParseJson(SafeStr(wsE.Cells(exRow(nm), EX_JSON).Value))
                End If
                e = ""
                bodies(r) = JsonOf(RowToObj(nm, d, raw, e))
                If e <> "" Then errs = errs & vbLf & "・③ " & r & "行目 " & nm & "：" & Replace(e, vbLf, " / ")
            End If
        End If
    Next

    ' ⑦：match の参照先と、対象・除外の重複を確認
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsP, r, ING_JUDGE) And SafeStr(wsP.Cells(r, 2).Value) <> "削除" Then
            nm = SafeStr(wsP.Cells(r, 3).Value)
            Set names = QuotedNames(SafeStr(wsP.Cells(r, ING_MATCH).Value))
            For Each v In names
                If delA.Exists(v) Then
                    errs = errs & vbLf & "・⑦ " & r & "行目 " & nm & "：③で削除するHIPオブジェクト「" & v & "」を参照しています"
                ElseIf Not known.Exists(v) Then
                    warns = warns & vbLf & "・⑦ " & r & "行目 " & nm & "：「" & v & "」が②既存一覧にも③の登録にもありません"
                End If
            Next
            If SafeStr(wsP.Cells(r, 5).Value) <> MODE_RAW Then
                Set seen = CreateObject("Scripting.Dictionary")
                For Each v In SplitMembers(SafeStr(wsP.Cells(r, 6).Value))
                    seen(CStr(v)) = True
                Next
                For Each v In SplitMembers(SafeStr(wsP.Cells(r, 7).Value))
                    If seen.Exists(CStr(v)) Then errs = errs & vbLf & "・⑦ " & r & "行目 " & nm & "：「" & v & "」が対象と除外の両方にあります"
                Next
            End If
        End If
    Next
    If errs <> "" Then
        MsgBox "次の入力を直してから、もう一度JSON発行してください。" & errs, vbExclamation
        Exit Sub
    End If
    If warns <> "" Then
        If MsgBox("確認してください（このまま発行するとAPIでエラーになる可能性があります）。" & warns & vbLf & vbLf & "続けますか？", _
                  vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    ClearJsonRows wsOut
    o = JS_FIRST
    ' 1. HIPオブジェクト 登録・変更
    For r = IN_FIRST To IN_LAST
        If bodies.Exists(r) Then
            nm = SafeStr(wsA.Cells(r, 3).Value)
            If SafeStr(wsA.Cells(r, 2).Value) = "登録" Then
                WriteJsonRow wsOut, o, "登録", nm, "POST", WithQuery(BaseUrl(), ContainerQuery()), bodies(r)
            Else
                WriteJsonRow wsOut, o, "変更", nm, "PUT", BaseUrl() & "/" & wsA.Cells(r, IN_ID).Value, bodies(r)
            End If
        End If
    Next
    ' 2. HIPプロファイル 登録・変更 → 3. 削除
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsP, r, ING_JUDGE) And SafeStr(wsP.Cells(r, 2).Value) <> "削除" Then WriteJsonRowProfile wsP, r, wsOut, o
    Next
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsP, r, ING_JUDGE) And SafeStr(wsP.Cells(r, 2).Value) = "削除" Then WriteJsonRowProfile wsP, r, wsOut, o
    Next
    ' 4. HIPオブジェクト 削除
    For r = IN_FIRST To IN_LAST
        If IsTarget(wsA, r, IN_JUDGE) And SafeStr(wsA.Cells(r, 2).Value) = "削除" Then
            WriteJsonRow wsOut, o, "削除", SafeStr(wsA.Cells(r, 3).Value), "DELETE", BaseUrl() & "/" & wsA.Cells(r, IN_ID).Value, ""
        End If
    Next
    If o > JS_FIRST Then
        With wsOut.Range(wsOut.Cells(JS_FIRST, 1), wsOut.Cells(o - 1, 8))
            .Borders.LineStyle = xlContinuous
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    wsOut.Activate
    MsgBox "③⑦から " & (o - JS_FIRST) & " 件のJSONを④に生成しました。", vbInformation
End Sub

' JSON発行の対象行か（nameあり、判定がOKか警告）
Private Function IsTarget(ByVal ws As Worksheet, ByVal r As Long, ByVal judgeCol As Long) As Boolean
    Dim judge As String
    judge = SafeStr(ws.Cells(r, judgeCol).Value)
    IsTarget = (SafeStr(ws.Cells(r, 3).Value) <> "" And (judge = "OK" Or judge = "警告"))
End Function

' ③の1行を 項目キー → 文字列 に
Private Function ReadRow(ByVal ws As Worksheet, ByVal r As Long) As Object
    Dim d As Object, keys As Variant, v As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    keys = FieldKeys()
    v = ws.Range(ws.Cells(r, F_FIRST), ws.Cells(r, F_LAST)).Value
    For i = 0 To UBound(keys)
        d(keys(i)) = Trim$(Replace(SafeStr(v(1, i + 1)), vbCr, ""))
    Next
    Set ReadRow = d
End Function

Private Sub WriteJsonRowProfile(ByVal wsIn As Worksheet, ByVal r As Long, ByVal wsOut As Worksheet, ByRef o As Long)
    Dim op As String, nm As String, body As String
    op = SafeStr(wsIn.Cells(r, 2).Value)
    nm = SafeStr(wsIn.Cells(r, 3).Value)
    If op <> "削除" Then
        body = "{""name"":" & JsonStr(nm)
        If SafeStr(wsIn.Cells(r, 4).Value) <> "" Then body = body & ",""description"":" & JsonStr(SafeStr(wsIn.Cells(r, 4).Value))
        body = body & ",""match"":" & JsonStr(SafeStr(wsIn.Cells(r, ING_MATCH).Value)) & "}"
    End If
    Select Case op
        Case "登録"
            WriteJsonRow wsOut, o, "登録(Profile)", nm, "POST", WithQuery(EpUrl(EP_PROFILE), ContainerQuery()), body
        Case "変更"
            WriteJsonRow wsOut, o, "変更(Profile)", nm, "PUT", EpUrl(EP_PROFILE) & "/" & wsIn.Cells(r, ING_ID).Value, body
        Case "削除"
            WriteJsonRow wsOut, o, "削除(Profile)", nm, "DELETE", EpUrl(EP_PROFILE) & "/" & wsIn.Cells(r, ING_ID).Value, ""
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
        Prog "API実行中 " & (r - JS_FIRST + 1) & " / " & n
        st = HttpCall(CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), CStr(ws.Cells(r, 6).Value), resText)
        ws.Cells(r, 8).Value = st
        If st >= 200 And st < 300 Then ok = ok + 1 Else ng = ng + 1
        WriteLog CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), CStr(ws.Cells(r, 4).Value), CStr(ws.Cells(r, 5).Value), st, ErrSummary(resText)
    Next
    Application.StatusBar = False
    ' 実行後のSCMの状態を取得し、②⑥既存一覧と比較して⑧⑨差分確認に出力
    BuildDiffCore
    Worksheets(SH_DIFF).Activate
    MsgBox "成功 " & ok & " 件／失敗 " & ng & " 件" & vbLf & _
           "HipObj_差分確認 と HipProfile_差分確認 に、現在の設定一覧と既存一覧（実行前）からの変化を出力しました。内容を確認してから、SCMの画面でPushしてください。", vbInformation
End Sub

'----------------------------------------------------------
' クリア系
'----------------------------------------------------------
Public Sub ClearInput()
    If MsgBox("入力（B・C列と F～BM列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    With Worksheets(SH_INPUT)
        .Range(.Cells(IN_FIRST, 2), .Cells(IN_LAST, 3)).ClearContents
        .Range(.Cells(IN_FIRST, F_FIRST), .Cells(IN_LAST, F_LAST)).ClearContents
    End With
End Sub

Public Sub ClearInputProfile()
    If MsgBox("入力（B～H列）をクリアします。よろしいですか？", vbOKCancel) <> vbOK Then Exit Sub
    Worksheets(SH_INPUTP).Range("B" & IN_FIRST & ":H" & IN_LAST).ClearContents
End Sub

Public Sub ClearJson()
    ClearJsonRows Worksheets(SH_JSON)
End Sub

Private Sub ClearJsonRows(ByVal ws As Worksheet)
    ws.Range(ws.Cells(JS_FIRST, 1), ws.Cells(ws.Rows.Count, 8)).Clear
End Sub

'==========================================================
' ⑧⑨ 差分確認（②⑥既存一覧＝実行前 と SCMの最新状態＝実行後 を比較）
'   HIPオブジェクトは description と条件のカテゴリ単位で比べる（中身は「項目: 値」を1行ずつ）
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
    Dim ws As Worksheet, r As Long, n As Long, opName As String, js As String
    Set before = CreateObject("Scripting.Dictionary")
    Set after = CreateObject("Scripting.Dictionary")
    Set ops = CreateObject("Scripting.Dictionary")
    Application.Calculate
    ' 実行前：②⑥既存一覧シート（②は保存してある JSON から）
    Set ws = Worksheets(SH_EXIST)
    For r = EX_FIRST To LastRow(ws, 3)
        If SafeStr(ws.Cells(r, 3).Value) <> "" Then
            js = SafeStr(ws.Cells(r, EX_JSON).Value)
            If js <> "" Then Set item = ParseJson(js) Else Set item = Nothing
            before("HipObj|" & ws.Cells(r, 3).Value) = ObjDiffValues(item)
        End If
    Next
    Set ws = Worksheets(SH_EXISTP)
    For r = EX_FIRST To LastRow(ws, 3)
        If SafeStr(ws.Cells(r, 3).Value) <> "" Then
            before("Profile|" & ws.Cells(r, 3).Value) = Array(SafeStr(ws.Cells(r, 4).Value), SafeStr(ws.Cells(r, 5).Value))
        End If
    Next
    ' 実行後：SCMから取得（シートは書き換えない）
    Set items = FetchAllItems(BaseUrl(), "差分確認", False, False)
    If items Is Nothing Then Exit Function
    For Each item In items
        after("HipObj|" & Nz(item, "name")) = ObjDiffValues(item)
    Next
    Set items = FetchAllItems(EpUrl(EP_PROFILE), "差分確認", False, False)
    If items Is Nothing Then Exit Function
    For Each item In items
        after("Profile|" & Nz(item, "name")) = Array(Nz(item, "description"), Nz(item, "match"))
    Next
    ' ④の操作
    Set ws = Worksheets(SH_JSON)
    For r = JS_FIRST To LastRow(ws, 4)
        opName = SafeStr(ws.Cells(r, 2).Value)
        If opName <> "" Then
            If InStr(opName, "(Profile)") > 0 Then
                ops("Profile|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), SafeStr(ws.Cells(r, 8).Value))
            Else
                ops("HipObj|" & ws.Cells(r, 3).Value) = Array(Left$(opName, 2), SafeStr(ws.Cells(r, 8).Value))
            End If
        End If
    Next
    Application.ScreenUpdating = False
    n = WriteDiffSheet(Worksheets(SH_DIFF), "HipObj", DiffCats(), before, after, ops)
    n = n + WriteDiffSheet(Worksheets(SH_DIFFP), "Profile", Array("description", "match"), before, after, ops)
    Application.ScreenUpdating = True
    WriteLog "差分確認", "", "GET", BaseUrl() & " / " & EpUrl(EP_PROFILE), 200, n & "件の変化"
    BuildDiffCore = True
End Function

' HIPオブジェクト1件 → DiffCats の順の表示文字列
Private Function ObjDiffValues(ByVal item As Object) As Variant
    Dim cats As Variant, v() As String, i As Long
    cats = DiffCats()
    ReDim v(UBound(cats))
    If Not item Is Nothing Then
        v(0) = Nz(item, "description")
        For i = 1 To UBound(cats)
            v(i) = FlatText(Child(item, CStr(cats(i))))
        Next
    End If
    ObjDiffValues = v
End Function

' 1種類分（HipObj / Profile）：SCMの現在の全件と削除されたものを一覧表示し、既存一覧との変化を付ける。戻り値は変化のあった件数
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
        If chg = "変更" Then ws.Cells(r, cF + 1).Value = Describe(fieldNames, before(k), after(k))
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

' 変更前の内容（変わった項目だけ「[項目]」の見出し付きで改行区切り）
Private Function Describe(ByVal fieldNames As Variant, ByVal b As Variant, ByVal a As Variant) As String
    Dim i As Long, s As String
    For i = 0 To UBound(b)
        If CStr(b(i)) <> CStr(a(i)) Then
            If s <> "" Then s = s & vbLf
            If CStr(b(i)) = "" Then
                s = s & "[" & fieldNames(i) & "] （なし）"
            Else
                s = s & "[" & fieldNames(i) & "]" & vbLf & b(i)
            End If
        End If
    Next
    Describe = s
End Function

'----------------------------------------------------------
' JSON 書き出し・表示
'----------------------------------------------------------
' Dictionary / Collection / 値 → JSON文字列
Private Function JsonOf(ByVal v As Variant) As String
    Dim k As Variant, s As String, x As Variant
    If IsObject(v) Then
        If v Is Nothing Then
            JsonOf = "null"
        ElseIf TypeName(v) = "Dictionary" Then
            For Each k In v.Keys
                If s <> "" Then s = s & ","
                s = s & JsonStr(CStr(k)) & ":" & JsonOf(v(k))
            Next
            JsonOf = "{" & s & "}"
        Else
            For Each x In v
                If s <> "" Then s = s & ","
                s = s & JsonOf(x)
            Next
            JsonOf = "[" & s & "]"
        End If
    Else
        Select Case VarType(v)
            Case vbBoolean: JsonOf = IIf(v, "true", "false")
            Case vbNull, vbEmpty: JsonOf = "null"
            Case vbString: JsonOf = JsonStr(v)
            Case Else: JsonOf = Replace(CStr(v), ",", ".")
        End Select
    End If
End Function

' 差分確認用：オブジェクトを「項目: 値」の行にする（キーは並べ替え、criteria は省略、name付き配列は [name]）
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

'----------------------------------------------------------
' 補助
'----------------------------------------------------------
Private Function NewD() As Object
    Set NewD = CreateObject("Scripting.Dictionary")
End Function

' 要素1つの Dictionary {k: v}
Private Function D1(ByVal k As String, ByVal v As Variant) As Object
    Dim t As Object
    Set t = CreateObject("Scripting.Dictionary")
    If IsObject(v) Then Set t(k) = v Else t(k) = v
    Set D1 = t
End Function

' 子オブジェクト（無い・オブジェクトでないときは Nothing）
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

' 接頭辞（hi_ など）の項目に1つでも入力があるか
Private Function AnySet(ByVal d As Object, ByVal prefix As String) As Boolean
    Dim k As Variant
    For Each k In d.Keys
        If Left$(k, Len(prefix)) = prefix Then
            If d(k) <> "" Then AnySet = True: Exit Function
        End If
    Next
End Function

' 先頭の「!」（否定）を取り除き、あれば True
Private Function TakeNeg(ByRef s As String) As Boolean
    s = Trim$(s)
    If Left$(s, 1) = "!" Then
        s = Trim$(Mid$(s, 2))
        TakeNeg = True
    End If
End Function

' セル内の行（空行は除く）
Private Function Lines(ByVal s As String) As Collection
    Dim c As New Collection, a() As String, i As Long
    a = Split(Replace(s, vbCr, ""), vbLf)
    For i = LBound(a) To UBound(a)
        If Trim$(a(i)) <> "" Then c.Add Trim$(a(i))
    Next
    Set Lines = c
End Function

' 「a | b | c」→ 前後の空白を除いた配列
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

' GET（ページ送りしながら全件取得）。失敗時は Nothing。quiet=True なら失敗してもメッセージを出さない
Private Function FetchAllItems(ByVal baseU As String, ByVal opLabel As String, ByVal logOk As Boolean, ByVal quiet As Boolean) As Collection
    Dim url As String, resText As String, st As Long, res As Object, item As Object
    Dim offset As Long, cnt As Long, total As Long, items As New Collection
    offset = 0
    Do
        url = WithQuery(baseU, ContainerQuery() & "&limit=" & PAGE_LIMIT & "&offset=" & offset)
        st = HttpCall("GET", url, "", resText)
        If st <> 200 Then
            WriteLog opLabel, "", "GET", url, st, ErrSummary(resText)
            Application.ScreenUpdating = True
            If Not quiet Then MsgBox opLabel & "でSCMからの取得に失敗しました（HTTP " & st & "）。実行ログを確認してください。", vbExclamation
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

'==========================================================
' メンバー選択画面（ダブルクリックでチェックボックス一覧）
'   必要モジュール : modHip（このモジュール） / clsPickerEvents / frmMemberPick
'   ブックを開くと Auto_Open で自動有効化されます。
'   ダブルクリックが効かない時は PickerStart を実行（または Ctrl+Shift+M）
'   対象：HipProfile_登録変更削除 の F列（対象HIPオブジェクト）・G列（除外HIPオブジェクト）
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
    mVendor = False
    If Not PickerRule(Target.Worksheet, Target, ttl) Then Exit Function
    TryPick = True
    ShowPicker Target, ttl
End Function

' 対象列の定義
'   AddSrc シート名, name列, 開始行, 種別ラベル [, 条件列, 条件値]
'   AddMark シート名, name列, 開始行, 条件列, 条件値, 付記
Private Function PickerRule(ByVal ws As Worksheet, ByVal c As Range, ByRef ttl As String) As Boolean
    Dim key As String, cat As String
    Select Case ws.Name
        Case SH_INPUT
            ' ベンダー欄（am_vendor など）：HIP_ベンダー一覧 の同じカテゴリから選ぶ
            If c.Row < IN_FIRST Or c.Row > IN_LAST Or c.Column < F_FIRST Or c.Column > F_LAST Then Exit Function
            key = FieldKeys()(c.Column - F_FIRST)
            If Right$(key, 7) <> "_vendor" Then Exit Function
            cat = CatLabel(Left$(key, InStr(key, "_") - 1))
            AddVendorSrc cat
            mVendor = True
            ttl = cat & " のベンダー・製品選択：" & SafeStr(ws.Cells(c.Row, 3).Value)
            PickerRule = True
        Case SH_INPUTP
            If c.Column < 6 Or c.Column > 7 Or c.Row < IN_FIRST Or c.Row > IN_LAST Then Exit Function
            If SafeStr(ws.Cells(c.Row, 5).Value) = MODE_RAW Then Exit Function
            ' 候補 = ②既存一覧 + ③で登録予定（O列のメンバー確認と同じ範囲）
            AddSrc SH_EXIST, 3, EX_FIRST, "HIPオブジェクト"
            AddSrc SH_INPUT, 3, IN_FIRST, "HIPオブジェクト(③で登録)", 2, "登録"
            AddMark SH_INPUT, 3, IN_FIRST, 2, "削除", " ※③で削除予定"
            ttl = IIf(c.Column = 6, "対象HIPオブジェクト選択：", "除外HIPオブジェクト(not)選択：") & SafeStr(ws.Cells(c.Row, 3).Value)
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
    If mVendor Then
        f.Setup ttl, mNames, VendorItems(cur), vbLf     ' ベンダー名にはカンマを含むものがあるため行単位
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
    Dim c As New Collection, parts() As String, i As Long, t As String
    parts = Split(Replace(Replace(s, vbCr, ""), vbLf, ","), ",")
    For i = LBound(parts) To UBound(parts)
        t = Trim$(parts(i))
        If t <> "" Then c.Add t
    Next
    Set SplitItems = c
End Function

'----------------------------------------------------------
' ベンダー欄のピッカー
'   候補 = HIP_ベンダー一覧 の同じカテゴリの行
'     「ベンダー」              … 製品を指定しない（そのベンダーのすべての製品）
'     「ベンダー | 製品」       … 製品を指定
'   書き戻しはベンダーごとに1行「ベンダー | 製品1 ; 製品2」（製品指定なしを選んだベンダーは「ベンダー」だけ）
'----------------------------------------------------------
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
    Set ws = Worksheets(SH_VENDOR)
    lastR = LastRow(ws, 3)
    If lastR < VD_FIRST Then Exit Sub
    v = ToArray2(ws.Range(ws.Cells(VD_FIRST, 2), ws.Cells(lastR, 4)).Value)
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

Private Function ToArray2(ByVal v As Variant) As Variant
    ToArray2 = v      ' 3列の範囲は常に2次元配列
End Function

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

' 既存HIPオブジェクトで使われているベンダー・製品を HIP_ベンダー一覧 に追加（同じものがあれば追加しない）
Private Sub MergeVendors(ByVal items As Collection)
    Dim ws As Worksheet, have As Object, item As Object, o As Object, vd As Object, cat As Variant
    Dim r As Long, lbl As String, pr As Variant, n As Long
    Set ws = Worksheets(SH_VENDOR)
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
    If n > 0 Then WriteLog "既存取得", "", "", SH_VENDOR, 200, "既存のHIPオブジェクトからベンダー・製品を " & n & " 件追加"
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

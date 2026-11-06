# SCM API連携 Excelツール

Palo Alto Strata Cloud Manager（Prisma Access Configuration API）用の Excel + VBA ツール一式です。ツールごとにフォルダを分けています。

| フォルダ | ツール | Excel | ツール用モジュール |
|---|---|---|---|
| scm-excel | アドレスオブジェクト／アドレスグループ | SCM_AddressObject.xlsx | modScm.bas |
| scm-urlcat | カスタムURLカテゴリ | SCM_URLCategory.xlsx | modUrlCat.bas |
| scm-service | サービス／サービスグループ | SCM_ServiceObject.xlsx | modSvc.bas |
| scm-application | アプリケーション／アプリケーショングループ | SCM_Application.xlsx | modApp.bas |
| scm-policy | セキュリティポリシー | SCM_SecurityPolicy.xlsx | modSecRule.bas |
| scm-decryption | 復号ポリシー | SCM_DecryptionPolicy.xlsx | modDecRule.bas |
| scm-tag | タグ | SCM_Tag.xlsx | modTag.bas |
| scm-hip | HIPオブジェクト／HIPプロファイル | SCM_HIP.xlsx | modHip.bas |
| scm-edl | 外部動的リスト（EDL） | SCM_EDL.xlsx | modEdl.bas |
| api-spec | 公式 OpenAPI 仕様（参照用） | | |

## 取り込み方

1. Excel を開き、VBE で `modScmCommon.bas` とツール用モジュール（上の表）をインポートします。
2. メンバー選択ピッカーがあるツールは `clsPickerEvents.cls` もインポートし、`modPickerSetup.bas` の `CreatePickerForm` を実行するか、空の UserForm `frmMemberPick` に `frmMemberPick_コード.txt` を貼り付けます。

`.bas` / `.cls` / `.txt` は Shift_JIS（cp932）＋ CRLF です。`tools/` は Excel 生成・パッチ用の Python スクリプトと VBA の UTF-8 ソースです。

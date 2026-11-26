# 時事Q 正式問題データ保存仕様

対象: MsCafetech/app_mscafe_net / main。新規・過去時事問題収集の共通仕様。関連知識・画像の保存仕様ではない。対象期間・収集目標は各タスクの指示に従う。

## 1. 正式データと日付
- 正式データは current_affairs/manifest.json が管理する current_affairs/questions/YYYY/MM/DD.json。
- news_date は出典本文・公式資料で確認した出来事の日。公式発表が主題なら発表日、報道そのものが主題なら報道日。実行日・収集日・単なる記事公開日を代用しない。年月しか分からない場合は登録しない。
- news_date から JSON.date、保存パス、ID の YYYYMMDD、manifest entry の date/path、question 冒頭「YYYY年M月D日、」を機械的に決める。複数日分は日付ごとに振り分ける。
- 既存正式問題は保持し、無断で修正・移動・差替え・削除しない。ただし1年以上経過した正式日付別データは削除し manifest に反映する。
- questions_remake/ および manifest_remake.json は廃止済み。取得・検索・参照・採番・重複判定・集計・変更の対象外。schema.json、test.json、id_remap.json、related/、images/、退避・中間ファイルなども変更しない。中間ファイルを作らない。その他の退避データはID履歴確認のみに使用可能。

## 2. JSON 形式
日付別 JSON:
```json
{
  "schema_version": 1,
  "date": "YYYY-MM-DD",
  "language": "ja",
  "question_count": 0,
  "questions": [
    {
      "id": "YYYYMMDD-001",
      "news_date": "YYYY-MM-DD",
      "category": "politics",
      "difficulty": 1,
      "question_type": "current_affairs",
      "news_title": "",
      "news_summary": "",
      "question": "",
      "choices": ["", "", "", ""],
      "correct_index": 0,
      "explanation": "",
      "keywords": [],
      "sources": [{"name": "", "url": ""}]
    }
  ]
}
```
上は構造例（空値、question_count は実データに置換する）。category: politics, international, economy, society, environment, science, space, geography, history_culture, sports。difficulty: 1～3。correct_index: 0～3。question_type: current_affairs または related_knowledge（時事ニュース収集タスクの新規問題は current_affairs）。

manifest:
```json
{
  "schema_version": 1,
  "last_updated": "YYYY-MM-DDTHH:mm:ss+09:00",
  "language": "ja",
  "total_questions": 0,
  "total_files": 0,
  "files": [{"date": "YYYY-MM-DD", "path": "questions/YYYY/MM/DD.json", "question_count": 0, "revision": 1}]
}
```
構造例。last_updated は日本時間 ISO8601。files は日付降順。変更日のみ revision を1加算、新規は1。パスは current_affairs からの相対パス。

## 3. 採番・重複防止
- ID は YYYYMMDD-NNN。対象日の正式JSON内の最大連番、id_remap.json の old_id/new_id の同日最大連番、既知の削除済みIDの同日最大連番の最大値+1。該当なしは001。欠番、移動済み、削除済みID（20261125-025を含む）を再使用しない。
- 同一ニュースは報道機関、記事、報道日、質問、教科が異なっても1問のみ。法案提出・成立・公布・施行等の独立した出来事は別扱い可。過去30日の実質同一知識の出題も避ける。
- 正式manifestと関連する正式JSONで確認する。検索は関連語句で絞り、候補ファイルを実取得して確認する。廃止・退避データは正式問題の重複判定・件数集計に使わない。

## 4. 問題品質
- 小学4～6年生の中学受験向け。question は日付で始め、背景（何が起きたか、重要性）と質問を空行で区切る。報道日と出来事の日が異なる場合は明確に区別し、相対日付は避ける。
- ニュースの主題に直接関係する重要知識を1問。同分野のもっともらしい4択、正解は一つ。解説80～200文字程度で正解理由と教科知識を示す。出典名、実際に本文を確認した直接URL、キーワードを付ける。架空の事実・日付・URLは禁止。
- フリガナ：小学校の教育漢字には付けず、教育漢字外を含む語のみ語全体の正しい読みを全角括弧で付ける。対象は news_title, question, choices, explanation, news_summary, 表示用keywords。同語同読みは表示順（見出し、背景・質問、選択肢、解説）で初出1回のみ。非表示要約等の出現で表示本文の初出を抑制しない。別問題ではリセット。正解だけに付けてヒントにしない。
- 関連知識の別生成、グラフ、年表、比較表は問題収集タスクでは不要。

## 5. GitHub保存と検証
1. 利用可能な現行GitHub操作を確認し、mainの正式manifest、対象日のJSON、採番に必要なID履歴を取得する。旧機能名が見つからないだけで停止せず fetch_file、search、create_tree、create_commit、update_ref 等の利用可能な同等機能を使う。対象日JSONの404は新規作成。manifest不存在時は正式ファイルを確認して再構築。認証・権限エラーを不存在扱いしない。
2. 新規問題を追記し、manifestを更新。対象日の全問題について日付一致、ID一意、4択、正解一意、出典との一致を検証。実件数=question_count=manifest該当entry.question_count、total_questions=全entry件数合計、total_files=filesの長さ、manifest日付・パスの重複なしを検証する。
3. 最新mainのHEADを基準に日付別JSONとmanifestを同一Commitで更新する。例えば create_tree → create_commit → update_ref。競合時は最新HEADと変更対象を再取得し、重複・ID・件数を再検証してマージ・再試行する。既存更新を上書きしない。
4. mainから変更したJSONとmanifestを再取得し、保存内容、日付規則、件数、revisionを照合する。照合完了を正式登録の条件とする。
5. 既存の不整合を発見した場合、その日付への追加は中止して対象ID・項目を報告し、別日付の正常な処理を続ける。新規問題の不整合は出典に照らして修正。修正不能なデータは保存しない。

## 6. エラー報告
GitHub操作失敗ごとに日本時間の日時、実際の操作名、リポジトリ・ブランチ・対象パスまたはSHA、処理内容、実際のHTTPステータス・エラーコード・メッセージ、再試行回数と結果を報告する。番号が返らない場合は「エラー番号なし（応答に含まれず）」。機能検索の空結果、操作未実行、接続失敗、認証・権限、404、競合、タイムアウトを区別する。回復可能なエラーは再試行し、復旧分も報告する。認証情報は伏せる。エラーなしの場合も明記し、未実行を成功や接続エラーと報告しない。

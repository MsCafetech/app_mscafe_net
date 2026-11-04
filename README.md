# app.mscafe.net

MsCafe.tech（App Café by M.A.C.Y）のサイト。1ファイル構成です。

- `index.html` … 公開用（サーバーに置くのはこれだけ）。コメントなし、Shift_JIS にない文字はエスケープ済み。
- `source/index.commented.html` … 編集用（コメントあり）。
- `source/strip-comments.mjs` … 編集用から公開用を作るスクリプト。

## 公開用の作り直し

```sh
npm i acorn
node source/strip-comments.mjs source/index.commented.html index.html
```

## サーバーでの更新

```sh
curl -fsSL -o index.html https://raw.githubusercontent.com/MsCafetech/app_mscafe_net/main/index.html
```

## FuRUSH のデータ（data/furush/）

アプリ FuRUSH は起動時に、ここを直接見に来ます（サーバーは使いません）。

```
https://raw.githubusercontent.com/MsCafetech/app_mscafe_net/main/data/furush/index.json
https://raw.githubusercontent.com/MsCafetech/app_mscafe_net/main/data/furush/<モード>000.zip …
```

- 作り方：FuRUSH-KMP の `tools/pack_baksok.py` でできた `server_upload/baksok/` の中身を、ここに置いて main に push。
- 配信は5分ほどキャッシュされるので、push してからアプリに届くまで数分かかります。
- ファイル名・場所を変えるとアプリが取れなくなるので、`data/furush/` のまま置いてください。

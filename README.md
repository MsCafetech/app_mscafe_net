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

## FuRUSH のデータ（data/baksok/）

アプリ FuRUSH が `https://app.mscafe.net/data/baksok/` から取るデータ（index.json とモードごとの zip）です。
作り方は FuRUSH-KMP の `tools/pack_baksok.py`（できた `server_upload/baksok/` の中身をここに置いて push）。

サーバーへは `tools/sync-baksok.sh` で取り込みます。変わった zip だけを取り、sha256 を確かめて、
最後に index.json を差し替えるので、index.json と zip が食い違った状態は公開されません。

```sh
# 1回だけ：スクリプトを置く
curl -fsSL -o ~/sync-baksok.sh https://raw.githubusercontent.com/MsCafetech/app_mscafe_net/main/tools/sync-baksok.sh

# 手で取り込む（DEST は公開フォルダの data/baksok）
DEST=$HOME/public_html/data/baksok sh ~/sync-baksok.sh

# cron（10分ごと）
*/10 * * * * DEST=$HOME/public_html/data/baksok sh $HOME/sync-baksok.sh >> $HOME/sync-baksok.log 2>&1
```

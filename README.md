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

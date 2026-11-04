#!/bin/sh
# FuRUSH のデータ（data/baksok/）を GitHub からサーバーの公開フォルダへ取り込む。
#
# 使い方（サーバーで。cron で 10 分ごとなど）:
#   DEST=$HOME/public_html/data/baksok sh sync-baksok.sh
#   （リポジトリは公開なのでトークン不要）
#
# ・index.json に書かれた sha256 と違う zip だけを取る（取ったら sha256 を確かめる）
# ・zip を全部置き終えてから、最後に index.json を差し替える
#   （アプリが新しい index.json を見たときには、必ず対応する zip が置いてある）
# ・index.json が前回と同じなら何もしない
# ・raw は数分キャッシュされるので、push 直後に zip が古くて sha256 が合わなければ何も変えずに終わる（次の回で取り込む）
set -eu

: "${DEST:?DEST（公開フォルダの data/baksok）を指定してください}"
REPO="${REPO:-MsCafetech/app_mscafe_net}"
BRANCH="${BRANCH:-main}"
RAW="${RAW:-https://raw.githubusercontent.com/$REPO/$BRANCH/data/baksok}"

sha() { sha256sum "$1" | cut -d' ' -f1; }
fetch() { # fetch <ファイル名> <保存先>
  curl -fsSL -o "$2" "$RAW/$1"
}

mkdir -p "$DEST"
TMP="$DEST/.sync-tmp"
rm -rf "$TMP"; mkdir -p "$TMP"
trap 'rm -rf "$TMP"' EXIT

fetch index.json "$TMP/index.json"
if [ -f "$DEST/index.json" ] && [ "$(sha "$TMP/index.json")" = "$(sha "$DEST/index.json")" ]; then
  exit 0
fi

# "file": "x.zip" と、その次の "sha256": "…" を組にして読む
sed -n 's/.*"file": *"\([^"]*\)".*/F \1/p; s/.*"sha256": *"\([0-9a-f]*\)".*/S \1/p' "$TMP/index.json" |
while read -r kind value; do
  if [ "$kind" = F ]; then file="$value"; continue; fi
  want="$value"
  if [ -f "$DEST/$file" ] && [ "$(sha "$DEST/$file")" = "$want" ]; then continue; fi
  fetch "$file" "$TMP/$file"
  got="$(sha "$TMP/$file")"
  if [ "$got" != "$want" ]; then
    echo "sha256 が合いません（push 直後なら次の回で取り込みます）: $file" >&2
    exit 1
  fi
  mv "$TMP/$file" "$DEST/$file"
  echo "更新: $file"
done

mv "$TMP/index.json" "$DEST/index.json"
echo "更新: index.json"

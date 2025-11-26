# file_loader_bot

Discord 上で Office 文書を PDF/画像へ変換するボットです。メッセージに添付されたファイルを自動で変換するほか、スラッシュコマンド `convert` でも手動変換できます。Excel など横長・縦長のシートは 1 ページに収まるよう自動的に縮小されるため、ページ分割されずに PDF/PNG を共有できます。

## 主な機能

- `.doc(x)`, `.xls(x/m)`, `.ppt(x)`, `.od*`, `.rtf`, `.pdf` などを PDF に変換
- PDF をページごとの PNG へ分割し、10 枚単位で Discord へ送信
- メッセージ添付の自動変換と `/convert` コマンドの 2 通りに対応
- スプレッドシートを PDF 化する際は各シートを 1 ページにフィット

## 前提条件

### システム依存ツール

- [LibreOffice](https://www.libreoffice.org/)（`soffice` コマンド）
- [poppler-utils](https://poppler.freedesktop.org/) の `pdftoppm`

いずれも `PATH` に含まれている必要があります。

### Python

- Python 3.11+
- pip で以下をインストールしてください:

```bash
python -m pip install -U pip
python -m pip install py-cord python-dotenv openpyxl
```

## セットアップ

1. `.env` を作成し、Discord のボットトークンを設定します。

   ```env
   DISCORD_BOT_TOKEN=ここにボットトークン
   ```

2. 依存パッケージをインストール（上記参照）。
3. LibreOffice と poppler-utils を OS にインストールし、`soffice` と `pdftoppm` をコマンドラインから実行できるようにします。

## 実行方法

```bash
python bot.py
```

起動するとログに `Logged in as ...` が表示され、ボットがオンライン状態になります。

## 使い方

- **自動変換**: サーバー内でサポートされているファイルを添付すると、PDF とページごとの PNG が返信されます。
- **スラッシュコマンド**: `/convert` でファイルを指定すると、PDF とプレビュー画像／ZIP が返却されます。

## トラブルシュート

- `Conversion failed: Missing dependency 'soffice' ...`: LibreOffice がインストールされ、`PATH` に入っているか確認してください。
- `Missing dependency 'openpyxl'`: `python -m pip install openpyxl` を実行してください。
- 変換に異常がある場合は、ログに出力されるエラーメッセージを確認し、必要に応じて依存ツールを再インストールしてください。

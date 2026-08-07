# doc2md-drop.ini 編集ガイド

`doc2md-drop.ini` は、`doc2md-drop.exe` の既定動作を決める設定ファイルです。配布時は `doc2md-drop.exe` と同じフォルダに置いてください。

通常は `output`、`output_format`、`use_pandoc`、`pdf_ocr`、`pdf_ocr_lang` だけ確認すれば十分です。

## 基本ルール

設定は `[defaults]` セクションに書きます。

```ini
[defaults]
output = output_data
output_format = adoc
use_pandoc = true
```

`#` で始まる行はコメントです。

真偽値は `true` / `false` を使ってください。`yes` / `no`、`on` / `off`、`1` / `0` も使えます。

`auto` と書ける項目は、標準インストール先や PATH から自動検出します。配布用には `auto` のままにするのが推奨です。

## よく編集する項目

### output

変換結果の出力先フォルダです。

```ini
output = output_data
```

相対パスの場合は、`doc2md-drop.exe` があるフォルダを基準に作成されます。

例:

```ini
output = converted
```

### output_format

出力形式です。

選択肢:

- `md`: Markdown 形式で出力します。
- `adoc`: AsciiDoc 形式で出力します。

```ini
output_format = adoc
```

### overwrite

同名の出力ファイルがある場合に上書きするかどうかです。

選択肢:

- `false`: 上書きせず、`ファイル名 (2).adoc` のように別名で保存します。
- `true`: 既存ファイルと対応する assets フォルダを上書きします。

```ini
overwrite = false
```

### recursive

フォルダを指定したとき、サブフォルダまで変換対象に含めるかどうかです。

選択肢:

- `false`: 指定フォルダ直下だけを対象にします。
- `true`: サブフォルダ内のファイルも対象にします。

```ini
recursive = false
```

## Pandoc 設定

Pandoc は DOCX / XLSX / Markdown / AsciiDoc 変換の品質を上げるために使います。ユーザー側に Pandoc 本体のインストールが必要です。

### use_pandoc

Pandoc を使うかどうかです。

選択肢:

- `true`: Pandoc を使います。見つからない場合は通常の Python 実装にフォールバックします。
- `false`: Pandoc を使いません。

```ini
use_pandoc = true
```

### pandoc_bin

Pandoc 実行ファイルの場所です。

選択肢:

- `auto`: 標準インストール先や PATH から自動検出します。配布用の推奨値です。
- `pandoc`: PATH 上の `pandoc` を使います。
- 絶対パス: 標準外の場所にインストールした場合に指定します。

```ini
pandoc_bin = auto
```

絶対パス指定の例:

```ini
pandoc_bin = C:\Users\user\AppData\Local\Pandoc\pandoc.exe
```

## PDF 設定

### pdf_engine

PDF の本文抽出エンジンです。

選択肢:

- `auto`: `pymupdf4llm` を優先し、失敗時は PyMuPDF のテキスト抽出にフォールバックします。通常はこれを使います。
- `pymupdf4llm`: レイアウトを考慮した Markdown 抽出を試みます。
- `pymupdf_text`: PyMuPDF の素朴なテキスト抽出を使います。ページ見出しを安定して付けたい場合に向きます。

```ini
pdf_engine = auto
```

### pdf_force_page_headings

PDF にページ見出しを強制的に付けるかどうかです。

選択肢:

- `false`: エンジンの出力に任せます。
- `true`: `## Page 1` のようなページ見出しを強制します。OCR 統合や RAG 用途では有効です。

```ini
pdf_force_page_headings = false
```

## PDF OCR 設定

OCR を使う場合は、ユーザー側に Tesseract OCR 本体と必要な言語データのインストールが必要です。

### pdf_ocr

PDF OCR を使うかどうかです。

選択肢:

- `true`: PDF に対して OCR を使います。
- `false`: OCR を使いません。

```ini
pdf_ocr = true
```

### pdf_ocr_strategy

OCR 結果を本文にどう入れるかです。

選択肢:

- `fallback`: 抽出済み本文が少ない場合だけ OCR します。通常の推奨値です。
- `integrate`: 常に OCR し、ページ見出し単位で本文へ統合します。
- `append`: OCR 結果を文書末尾に追記します。
- `replace`: 通常抽出の本文を使わず、OCR 結果だけを出力します。

```ini
pdf_ocr_strategy = fallback
```

### pdf_ocr_lang

Tesseract OCR の言語です。

よく使う値:

- `jpn`: 日本語横書き向けです。
- `jpn+jpn_vert`: 日本語横書きと縦書きの両方を使います。
- `eng`: 英語向けです。
- `jpn+eng`: 日本語と英語を使います。

```ini
pdf_ocr_lang = jpn+jpn_vert
```

指定した言語データが Tesseract に入っていない場合、OCR は失敗します。

### pdf_ocr_dpi

OCR 用に PDF ページを画像化するときの解像度です。

目安:

- `200`: 速いですが、認識精度は落ちる場合があります。
- `300`: 標準的な推奨値です。
- `400` 以上: 精度が上がる場合がありますが、処理時間とメモリ使用量が増えます。

```ini
pdf_ocr_dpi = 300
```

### pdf_ocr_psm

Tesseract のページ分割モードです。

よく使う値:

- `3`: 自動ページ分割です。
- `6`: 1つの均一なテキストブロックとして扱います。既定値です。
- `11`: まばらなテキストを拾います。

```ini
pdf_ocr_psm = 6
```

### pdf_ocr_oem

Tesseract の OCR エンジンモードです。

よく使う値:

- `3`: Tesseract の既定エンジンを使います。推奨値です。
- `1`: LSTM エンジンを使います。

```ini
pdf_ocr_oem = 3
```

### pdf_ocr_tesseract

Tesseract 実行ファイルの場所です。

選択肢:

- `auto`: 標準インストール先や PATH から自動検出します。配布用の推奨値です。
- `tesseract`: PATH 上の `tesseract` を使います。
- 絶対パス: 標準外の場所にインストールした場合に指定します。

```ini
pdf_ocr_tesseract = auto
```

絶対パス指定の例:

```ini
pdf_ocr_tesseract = C:\Program Files\Tesseract-OCR\tesseract.exe
```

### pdf_ocr_tessdata

Tesseract の言語データフォルダです。

選択肢:

- `auto`: 標準インストール先、`TESSDATA_PREFIX`、Tesseract 実行ファイル付近から自動検出します。配布用の推奨値です。
- 空欄: Tesseract の既定設定に任せます。
- 絶対パス: 標準外の場所に言語データがある場合に指定します。

```ini
pdf_ocr_tessdata = auto
```

絶対パス指定の例:

```ini
pdf_ocr_tessdata = C:\Program Files\Tesseract-OCR\tessdata
```

## Excel 設定

### excel_max_rows

各シートから出力する最大行数です。

```ini
excel_max_rows = 500
```

### excel_max_cols

各シートから出力する最大列数です。

```ini
excel_max_cols = 50
```

### excel_cell_max_chars

1セルあたりの最大文字数です。長いセルは省略されます。

```ini
excel_cell_max_chars = 200
```

`.xlsx` に貼り込まれた画像は、画像がある場合だけ `*_assets/images/` に抽出されます。画像がない場合は assets フォルダは作られません。

## Word 設定

### word_heading_style

Word を Pandoc なしで変換するときの Markdown 見出しスタイルです。

選択肢:

- `ATX`: `# 見出し` 形式です。通常はこちらを使います。
- `SETEXT`: 下線付きの見出し形式です。

```ini
word_heading_style = ATX
```

### word_strip_empty_lines / word_keep_empty_lines

Word を Pandoc なしで変換するとき、空行を除去するかどうかです。

選択肢:

- `word_strip_empty_lines = true`: 空行を除去します。
- `word_keep_empty_lines = true`: 空行を保持します。

両方を `true` にしないでください。配布用の既定では両方 `false` にして、ツール本体の既定動作に任せます。

```ini
word_strip_empty_lines = false
word_keep_empty_lines = false
```

## メタ情報と画像

### no_front_matter

出力ファイル先頭のメタ情報を省略するかどうかです。

選択肢:

- `false`: メタ情報を付けます。
- `true`: メタ情報を付けません。

```ini
no_front_matter = false
```

### no_sha256

メタ情報に入力ファイルの SHA-256 を含めるかどうかです。

選択肢:

- `false`: SHA-256 を含めます。
- `true`: SHA-256 を含めません。大きなファイルで少し高速になります。

```ini
no_sha256 = false
```

### include_source_path

メタ情報に入力ファイルの絶対パスを含めるかどうかです。

選択肢:

- `false`: 絶対パスを含めません。配布用の推奨値です。
- `true`: 絶対パスを含めます。

```ini
include_source_path = false
```

### no_images

画像抽出を無効にするかどうかです。

選択肢:

- `false`: 画像を抽出します。画像が実際にある場合だけ assets フォルダを作ります。
- `true`: 画像抽出を無効にします。

```ini
no_images = false
```

## default_args

INI に専用項目がない CLI オプションを追加できます。

例:

```ini
default_args =
    --pdf-max-pages 10
    --json-summary summary.json
```

1行に1つのオプション、または1組のオプションを書いてください。`#` で始まる行はコメントとして無視されます。

## トラブル時の確認

設定ファイルを一時的に無視する場合:

```powershell
.\doc2md-drop.exe --no-config sample.pdf
```

処理後にウィンドウを閉じず、ログを確認したい場合は通常どおり実行します。ターミナル実行で Enter 待ちを省く場合:

```powershell
.\doc2md-drop.exe --no-pause sample.pdf
```

Pandoc や Tesseract が見つからない場合は、`auto` の代わりに絶対パスを指定してください。

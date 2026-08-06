# doc2md_asciidoc_pandoc.py

PDF / Word(.docx) / Excel(.xlsx/.xls) を **Markdown(.md)** または **AsciiDoc(.adoc)** に変換する単一ファイルスクリプトです。  
PDF は **OCR（Tesseract）** を統合できます。`--use-pandoc` を使うと、可能な範囲で **Pandoc による高品質変換**を優先し、利用不可なら **起動時に1回だけ警告して従来方式にフォールバック**します。

> この README ではスクリプト名を `doc2md_asciidoc_pandoc.py` としています。  
> ファイル名を変更している場合は、コマンド中のファイル名を読み替えてください。

---

## 主な機能

- **PDF → Markdown / AsciiDoc**
  - PyMuPDF / pymupdf4llm による抽出（画像切り出し対応）
  - 必要に応じて **Tesseract OCR** をページ単位で統合（`--pdf-ocr`）
  - `--pdf-force-page-headings` / `--pdf-page-heading-template` で **ページ見出しを強制生成**（OCR統合やRAG用途向け）
- **Word(.docx) → Markdown / AsciiDoc**
  - `--use-pandoc` 有効時は Pandoc で直接変換（画像抽出も可）
  - それ以外は mammoth + markdownify
  - `--word-strip-empty-lines`（既定: 有効）/ `--word-keep-empty-lines` で空行処理を制御
- **Excel(.xlsx/.xls) → Markdown / AsciiDoc**
  - `.xlsx` は `--use-pandoc` 有効時に Pandoc で直接変換を優先
  - それ以外は pandas で表抽出
- **運用オプション**
  - `--fail-fast`：最初の失敗で中断
  - `--json-summary <path>`：結果サマリーをJSONで保存

---

## 動作要件

### Python
- Python **3.10+** 推奨

### Pythonパッケージ（pip）
最低限、以下をインストールしてください。

- PyMuPDF（PDF処理の基盤）
- pymupdf4llm（PDF→Markdown変換）
- pymupdf-layout
- Word変換: mammoth / markdownify
- Excel変換: pandas / openpyxl / tabulate
- OCR: pillow / pytesseract（＋別途 Tesseract 本体）

> **重要（必須）**: `pymupdf-layout` は pip のパッケージ名が **`pymupdf-layout`** で、モジュールは `pymupdf.layout` として提供されます。  
> `import pymupdf.layout` が失敗する場合、これが未導入の可能性が高いです。

### 追加ツール（任意）
- **Pandoc**: `--use-pandoc` 利用時に必要（システムにインストールして `pandoc` が実行できること）
- **Tesseract OCR**: `--pdf-ocr` 利用時に必要（システムにインストールして `tesseract` が実行できること）

---

## セットアップ手順（推奨）

### 1) 仮想環境の作成
```bash
python -m venv .venv
# Windows (PowerShell)
.venv\Scripts\Activate.ps1
# macOS / Linux
source .venv/bin/activate
```

### 2) pip を更新
```bash
python -m pip install -U pip setuptools wheel
```

### 3) 依存パッケージをインストール（必須）
```bash
# PyMuPDF + pymupdf4llm
python -m pip install -U pymupdf pymupdf4llm

# layout機能 / pymupdf.layout.activate() 用
python -m pip install -U pymupdf-layout

# Word/Excel/OCR周辺（必要に応じて）
python -m pip install -U pandas openpyxl mammoth markdownify pillow pytesseract tabulate
```

#### まとめて入れる方法（任意）
```bash
python -m pip install -U "pymupdf4llm[ocr,layout]"
# 念のため layout パッケージも明示（「入ってない」事故を防ぐ）
python -m pip install -U pymupdf-layout
```

---

## Pandoc を使う（`--use-pandoc`）

### 1) Pandoc をインストール

#### Windows（winget）
```bash
winget install --source winget --exact --id JohnMacFarlane.Pandoc
```
- インストール後、PowerShellを再起動してください
#### macOS（Homebrew）
Pandoc だけ入れる場合：
```bash
brew install pandoc
```

Pandoc とあわせて、連携ツール（SVG変換/フィルタ/TeX など）も入れる場合（よく使われる構成）：
```bash
brew install pandoc librsvg python homebrew/cask/basictex
```

Homebrew の現在の書き方で分けて実行する場合（こちらを推奨）：
```bash
brew install pandoc librsvg python
brew install --cask basictex
```

> 補足:
> - このツールで **Word/Excel → Markdown/AsciiDoc** を行うだけなら、基本的に **pandoc だけ**で十分です。
> - `basictex`（TeX）は、Pandoc 側で **PDFを生成する用途**で必要になります（本ツールの md/adoc 出力自体には必須ではありません）。

### 2) 動作確認
```bash
pandoc --version
```

### 3) スクリプト実行
```bash
python doc2md_asciidoc_pandoc.py --use-pandoc --output-format adoc input.docx
```

> **仕様**: `--use-pandoc` が指定されていても、起動時に pandoc が見つからない / 起動できない場合は  
> **警告を1回だけ出して、自動で従来方式にフォールバック**します。

---

## OCR を使う（`--pdf-ocr`）

### 1) Tesseract をインストール
OSに合わせてインストールしてください（例）:
- macOS: Homebrew（`brew install tesseract`）
- Ubuntu/Debian: apt（`sudo apt install tesseract-ocr`）
- Windows: インストーラ（UB Mannheim版など）/ winget 等

### 2) 動作確認
```bash
tesseract --version
```

### 3) スクリプト実行例
```bash
# 文字抽出が弱いPDFだけOCRする（fallback）
python doc2md_asciidoc_pandoc.py input.pdf --pdf-ocr --pdf-ocr-strategy fallback

# 常にOCRを統合（ページ単位）
python doc2md_asciidoc_pandoc.py input.pdf --pdf-ocr --pdf-ocr-strategy integrate
```

Windowsで `tesseract` がPATHに無い場合は、実行ファイルを指定できます。
```bash
python doc2md_asciidoc_pandoc.py input.pdf --pdf-ocr --pdf-ocr-tesseract "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

---

## 使い方（基本）

```bash
# Markdown出力（デフォルト）
python doc2md_asciidoc_pandoc.py input.pdf

# AsciiDoc出力
python doc2md_asciidoc_pandoc.py input.pdf --output-format adoc

# 出力先フォルダ変更（デフォルト: output_data）
python doc2md_asciidoc_pandoc.py input.pdf --output outdir

# 画像を出力しない
python doc2md_asciidoc_pandoc.py input.pdf --no-images
```

### 新オプション例：fail-fast / json-summary
```bash
# 最初の失敗で中断
python doc2md_asciidoc_pandoc.py docs/ --recursive --fail-fast

# 結果サマリーをJSONで保存
python doc2md_asciidoc_pandoc.py docs/ --recursive --json-summary summary.json
```

### 新オプション例：Wordの空行
```bash
# 既定: 空行は除去（--word-strip-empty-lines が有効）
python doc2md_asciidoc_pandoc.py input.docx

# 空行を保持したい場合
python doc2md_asciidoc_pandoc.py input.docx --word-keep-empty-lines
```

---

## 主なオプション一覧（完成版）

### 共通
- `-o, --output <dir>`: 出力先フォルダ（既定: `output_data`）
- `--output-format {md,adoc}`: 出力フォーマット（既定: `md`）
- `--overwrite`: 同名出力がある場合に上書き
- `--recursive`: フォルダ入力時に再帰探索
- `--no-front-matter`: フロントマター（YAML/AsciiDoc属性）を付けない
- `--no-sha256`: フロントマターに sha256 を含めない
- `--include-source-path`: フロントマターに source_path（絶対パス）を含める
- `--no-images`: 画像抽出（assets 出力）を無効化
- `--fail-fast`: 最初の失敗で処理を中断
- `--json-summary <path>`: 処理結果サマリーを JSON で保存（パス指定）
- `-v, --verbose`: 詳細ログ
- `-q, --quiet`: 警告以上のみ
- `--debug`: 失敗時に例外を再送出（スタックトレース確認用）
- `--version`: バージョン表示

### Pandoc
- `--use-pandoc`: Pandoc による高品質変換を有効化（利用不可なら起動時に警告→フォールバック）
- `--pandoc-bin <path>`: pandoc 実行ファイル名/パス（既定: `pandoc`）

### PDF
- `--pdf-engine {auto,pymupdf4llm,pymupdf_text}`: PDF 変換エンジン（既定: auto）
- `--pdf-max-pages <N>`: 変換する最大ページ数（未指定=全ページ）
- `--pdf-force-page-headings`: ページ見出し（Page N）を強制生成（OCR統合やRAG用途向け）
- `--pdf-page-heading-template <tpl>`: ページ見出しテンプレート（既定: `## Page {n}`）

### PDF OCR
- `--pdf-ocr`: OCR を有効化
- `--pdf-ocr-strategy {fallback,integrate,append,replace}`
  - `fallback`: 本文が少ない時だけ OCR（既定）
  - `integrate`: 常に OCR してページ単位で統合
  - `append`: OCR結果を本文末尾へ追記
  - `replace`: OCR結果のみ出力
- `--pdf-ocr-min-chars <N>`: fallback 判定のしきい値（既定: 200）
- `--pdf-ocr-lang <lang>`: OCR 言語（例: `jpn` / `jpn+jpn_vert`）
- `--pdf-ocr-dpi <dpi>`: ページレンダリングDPI（既定: 300）
- `--pdf-ocr-psm <N>`: Tesseract PSM（既定: 6）
- `--pdf-ocr-oem <N>`: Tesseract OEM（既定: 3）
- `--pdf-ocr-tesseract <path>`: `tesseract.exe` のパス（必要な場合）
- `--pdf-ocr-tessdata <dir>`: `tessdata` のディレクトリ（任意）
- `--pdf-ocr-debug-images`: 前処理後画像を assets に保存（精度調整用）
- `--pdf-ocr-threshold <N>`: 2値化しきい値（前処理パラメータ）
- `--pdf-ocr-contrast <R>`: コントラスト強調率（前処理パラメータ）

### Excel
- `--excel-max-rows <N>`: 各シートの最大行数（既定: 500）
- `--excel-max-cols <N>`: 各シートの最大列数（既定: 50）
- `--excel-cell-max-chars <N>`: セル文字数上限（既定: 200）

### Word(.docx)
- `--word-heading-style {ATX,SETEXT}`: 見出しスタイル（既定: ATX）
- `--word-strip-empty-lines`: 空行を除去（既定: 有効）
- `--word-keep-empty-lines`: 空行を保持（空行除去を無効化）

---

## 出力物の構成（例）

```
output_data/
  input.md          または input.adoc
  input_assets/
    images/
      input-1-1.png ...
```

---

## よくあるエラーと対処

### `import pymupdf.layout` が失敗する
`pymupdf-layout` を必ずインストールしてください。
```bash
python -m pip install -U pymupdf-layout
```

### `--use-pandoc` なのに pandoc が使われない
起動時に警告が出ていれば、pandoc が見つからない/起動不可です。  
`pandoc --version` が通る状態にしてください。

### OCRが動かない / `tesseract` が見つからない
- `tesseract --version` が通るか確認
- Windows は `--pdf-ocr-tesseract` でフルパス指定

### Q. OCRを指定したのに、Markdownに画像が含まれていません

- **A. 変換エンジンが「テキスト抽出モード」にフォールバックしている可能性があります。**

本ツールは通常、画像も抽出できる `pymupdf4llm` エンジンを使用しますが、ライブラリのエラー（ONNXRuntimeErrorなど）が発生した場合、処理を止めずに **`pymupdf_text`（テキスト抽出のみを行うモード）** に自動的に切り替わります。
このモードでは、仕様上 **画像は抽出されず、assetsフォルダも生成されません**。

- **対処法:**
`pymupdf4llm` や `numpy` のバージョン不整合が原因であることが多いです。以下を実行してライブラリを更新してください。

```bash
pip install --upgrade pymupdf pymupdf4llm numpy
```
### Q. OCRが読み取った画像（処理過程の画像）を確認したい

- **A. OCR処理のために一時的に生成された画像（コントラスト調整・2値化済み）を確認したい場合は、--pdf-ocr-debug-images オプションを使用してください。**

```bash
# OCRデバッグ画像を output_data/..._assets/ocr_debug_images/ に保存
python doc2md_asciidoc_pandoc.py input.pdf --pdf-ocr --pdf-ocr-debug-images
```
- ※ これは「OCRが文字認識に使った画像」を全ページ分保存するものであり、Markdown記事内に埋め込まれる図版とは異なります。

### Q. SyntaxWarning: invalid escape sequence が出る

- **A. Python 3.12以降で表示される警告です。動作に影響はありませんが、気になる場合はスクリプト内の正規表現文字列（r"..."）を確認してください。**
---

## ライセンス / 注意
- 生成された成果物の正確性は入力ファイルの品質（スキャン品質、フォント埋め込み、保護設定など）に依存します。
- 変換品質を重視する場合は `--use-pandoc` の利用を推奨します（利用不可の場合は自動フォールバック）。

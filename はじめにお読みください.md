# はじめにお読みください

## Windows SmartScreen が表示された場合

`doc2md-drop.exe` の初回実行時に、Windows SmartScreen により次のような画面が表示される場合があります。

```text
Windows によって PC が保護されました
Microsoft Defender SmartScreen は認識されないアプリの起動を停止しました。
```

これは、配布実績の少ない自作 exe や未署名 exe で表示されやすい Windows の保護機能です。

配布元が正しいことを確認できている場合は、次の手順で実行できます。

1. SmartScreen 画面の `詳細情報` をクリックします。
2. 表示された発行元とアプリ名を確認します。
3. 問題なければ `実行` をクリックします。

ファイルの入手元が不明な場合、ZIP の内容が改変された可能性がある場合、または想定外のファイル名になっている場合は実行しないでください。

SmartScreen の表示を根本的に減らすには、配布者側で Authenticode コード署名を行い、同じ発行者で配布実績を積む必要があります。

## 配布ファイルのインストールから実行まで

配布先ユーザーは、まず Pandoc と Tesseract OCR をインストールします。

1. Pandoc をインストールします。
   - 通常の Windows インストーラーでインストールしてください。
   - 標準インストール先または PATH から自動検出されます。

2. Tesseract OCR をインストールします。
   - PDF の OCR を使う場合に必要です。
   - インストール時に日本語用の `jpn` を追加してください。
   - 縦書き日本語も扱う場合は `jpn_vert` も追加してください。

3. 配布された次の 2 ファイルを、同じフォルダに置きます。

```text
doc2md-drop.exe
doc2md-drop.ini
```

4. 必要に応じて `doc2md-drop.ini` を編集します。
   - 通常はそのまま使えます。
   - 出力先を変える場合は `output` を変更します。
   - 出力形式を変える場合は `output_format = md` または `output_format = adoc` を指定します。
   - 詳細は `DOC2MD_DROP_INI_GUIDE.md` を参照してください。

5. 変換したい PDF / DOCX / XLSX / XLS ファイルを `doc2md-drop.exe` にドラッグアンドドロップします。

6. 変換結果は、既定では exe と同じフォルダの `output_data` に出力されます。

exe をダブルクリックした場合は、ファイル選択ダイアログから変換対象を選べます。

## 環境設定

ビルド前に、Python パッケージとは別に次の外部ツールをインストールします。

### Tesseract OCR

PDF の OCR を使う場合は Tesseract 本体が必要です。

このフォルダにあるインストーラーを使う場合は、次を実行します。

```powershell
.\tesseract-ocr-w64-setup-5.5.0.20241111.exe
```

インストール時のコンポーネント選択で、`Additional Scripts/language data` から日本語用の `JPN` を追加してください。日本語の縦書き文書も扱う場合は、`jpn_vert` も使える状態にしておくとよいです。

インストール後、必要に応じてパスを確認します。

```powershell
& "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe" --version
& "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe" --list-langs --tessdata-dir "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tessdata"
```

### Pandoc

DOCX / XLSX / Markdown / AsciiDoc 変換品質を上げる場合は Pandoc 本体が必要です。

このフォルダにあるインストーラーを使う場合は、次を実行します。

```powershell
msiexec /i .\pandoc-3.10-windows-x86_64.msi
```

インストール後、必要に応じてパスを確認します。

```powershell
& "$env:LOCALAPPDATA\Pandoc\pandoc.exe" --version
```

## ビルド

通常は単一 exe 形式でビルドします。配布に必要なファイルは基本的に `doc2md-drop.exe` と `doc2md-drop.ini` の 2 つです。

```powershell
.\build_windows_app.ps1
```

フォルダ形式でビルドしたい場合は次を実行します。

```powershell
.\build_windows_app.ps1 -OneDir
```

## 使い方

- `dist\doc2md-drop.exe` に PDF / DOCX / XLSX / XLS ファイルをドラッグアンドドロップすると変換します。
- exe をダブルクリックするとファイル選択ダイアログが開きます。
- 出力は exe と同じ場所を基準にした `output_data` フォルダへ作成されます。
- ドラッグアンドドロップ時の既定オプションは、exe と同じ場所の `doc2md-drop.ini` で変更できます。
- OCR を使う場合は、Python パッケージとは別に Tesseract 本体が必要です。
- Pandoc 変換を使う場合は、Pandoc 本体を別途インストールしてください。

## 配布

配布先ユーザーには、次の 2 ファイルを同じフォルダに置いてもらいます。

```text
doc2md-drop.exe
doc2md-drop.ini
```

ユーザー側では、Pandoc と Tesseract OCR を通常のインストーラーでインストールしてください。`doc2md-drop.ini` の `pandoc_bin = auto` / `pdf_ocr_tesseract = auto` / `pdf_ocr_tessdata = auto` により、標準インストール先と PATH から自動検出します。

## 既定オプションの設定

ビルド後に実際に編集するファイルは、exe と同じ場所にある `dist\doc2md-drop.ini` です。

各項目の意味と選択肢は `DOC2MD_DROP_INI_GUIDE.md` を参照してください。

ルート直下の `doc2md-drop.ini` は、ソースから `doc2md_windows_drop.py` を実行する場合や、次回ビルド時の元ファイルとして使います。ビルド済み exe にドラッグアンドドロップする運用では、ルート直下の INI を変更しても `dist\doc2md-drop.ini` には自動反映されません。

AsciiDoc 出力、Pandoc 使用、PDF OCR 使用にする場合は、最低限次の項目を確認・変更します。

```ini
[defaults]
output = output_data
output_format = adoc
pdf_engine = auto

use_pandoc = true
pandoc_bin = auto

pdf_ocr = true
pdf_ocr_strategy = fallback
pdf_ocr_lang = jpn+jpn_vert
pdf_ocr_tesseract = auto
pdf_ocr_tessdata = auto
```

標準インストール先にない場合だけ、`auto` の代わりに実行ファイルや `tessdata` の絶対パスを指定してください。

例えば、ドラッグアンドドロップ時の出力先だけを `converted` に変える場合は次のようにします。

```ini
[defaults]
output = converted
```

未対応のオプションや細かい指定は `default_args` に CLI と同じ形式で追加できます。

```ini
[defaults]
default_args =
    --pdf-max-pages 10
    --json-summary summary.json
```

設定ファイルを一時的に無視する場合は、ラッパー用オプション `--no-config` を付けます。

```powershell
.\dist\doc2md-drop.exe --no-config sample.pdf
```

コマンドラインから既存オプションも渡せます。

```powershell
.\dist\doc2md-drop.exe --output converted --pdf-ocr sample.pdf
```

ターミナル実行時に最後の Enter 待ちを省く場合は、ラッパー用オプション `--no-pause` を付けます。

```powershell
.\dist\doc2md-drop.exe --no-pause sample.pdf
```

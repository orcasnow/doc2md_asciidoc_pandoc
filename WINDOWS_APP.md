# Windows アプリ化手順

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

通常はフォルダ形式でビルドします。起動が速く、PyMuPDF や pandas などの大きな依存を含めやすいです。

```powershell
.\build_windows_app.ps1
```

単一 exe にしたい場合は次を実行します。

```powershell
.\build_windows_app.ps1 -OneFile
```

## 使い方

- `dist\doc2md-drop\doc2md-drop.exe` に PDF / DOCX / XLSX / XLS ファイルをドラッグアンドドロップすると変換します。
- exe をダブルクリックするとファイル選択ダイアログが開きます。
- 出力は exe と同じ場所を基準にした `output_data` フォルダへ作成されます。
- ドラッグアンドドロップ時の既定オプションは、exe と同じ場所の `doc2md-drop.ini` で変更できます。
- OCR を使う場合は、Python パッケージとは別に Tesseract 本体が必要です。
- Pandoc 変換を使う場合は、Pandoc 本体を別途インストールして `--use-pandoc` を指定してください。

## 既定オプションの設定

ビルド後に実際に編集するファイルは、exe と同じ場所にある `dist\doc2md-drop\doc2md-drop.ini` です。

ルート直下の `doc2md-drop.ini` は、ソースから `doc2md_windows_drop.py` を実行する場合や、次回ビルド時の元ファイルとして使います。ビルド済み exe にドラッグアンドドロップする運用では、ルート直下の INI を変更しても `dist\doc2md-drop\doc2md-drop.ini` には自動反映されません。

AsciiDoc 出力、Pandoc 使用、PDF OCR 使用にする場合は、最低限次の項目を確認・変更します。

```ini
[defaults]
output = output_data
output_format = adoc
pdf_engine = auto

use_pandoc = true
pandoc_bin = C:\Users\<ユーザー名>\AppData\Local\Pandoc\pandoc.exe

pdf_ocr = true
pdf_ocr_strategy = fallback
pdf_ocr_lang = jpn+jpn_vert
pdf_ocr_tesseract = C:\Users\<ユーザー名>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe
pdf_ocr_tessdata = C:\Users\<ユーザー名>\AppData\Local\Programs\Tesseract-OCR\tessdata
```

`C:\Users\<ユーザー名>\...` の `<ユーザー名>` は、この Windows にログインしているユーザー名に置き換えてください。今回の環境では `C:\Users\U743539\...` です。

Tesseract と Pandoc を PATH に登録済みであれば `pandoc_bin = pandoc` のような指定でも動きますが、ドラッグアンドドロップ実行では PATH が期待通りにならないことがあります。Windows アプリとして使う場合は、上のように絶対パスを指定する方が確実です。

他ユーザーに配布する場合も、Tesseract / Pandoc のインストール先パスが変わるだけであれば exe の再ビルドは不要です。配布先のユーザー側で Tesseract / Pandoc をインストールしてもらい、exe と同じ場所の `doc2md-drop.ini` にある次の値だけ、その環境に合わせて変更します。

```ini
pandoc_bin = C:\Users\<ユーザー名>\AppData\Local\Pandoc\pandoc.exe
pdf_ocr_tesseract = C:\Users\<ユーザー名>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe
pdf_ocr_tessdata = C:\Users\<ユーザー名>\AppData\Local\Programs\Tesseract-OCR\tessdata
```

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
.\dist\doc2md-drop\doc2md-drop.exe --no-config sample.pdf
```

コマンドラインから既存オプションも渡せます。

```powershell
.\dist\doc2md-drop\doc2md-drop.exe --output converted --pdf-ocr sample.pdf
```

ターミナル実行時に最後の Enter 待ちを省く場合は、ラッパー用オプション `--no-pause` を付けます。

```powershell
.\dist\doc2md-drop\doc2md-drop.exe --no-pause sample.pdf
```

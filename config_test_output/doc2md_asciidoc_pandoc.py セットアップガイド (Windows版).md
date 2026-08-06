---
tool: doc2md
tool_version: 2.0.1
type: word
source_type: word
source_file: doc2md_asciidoc_pandoc.py セットアップガイド (Windows版).docx
title: doc2md_asciidoc_pandoc.py セットアップガイド (Windows版)
converted_at: 2026-07-17T09:24:56.394499+00:00
engine: mammoth
---

このファイルはPDF、Word (.docx)、Excel (.xlsx/.xls) を **Markdown (.md)** または **AsciiDoc (.adoc)** に変換するためのツールです。Windows環境で、requirements.txt を使って一括でセットアップする手順を解説します。
## **準備するもの (外部ツール)**
Pythonライブラリ以外に、高品質な変換やOCR機能を利用する場合は以下のツールをインストールしてください。
### **Pandoc** (Word/Excelの高精度変換に推奨)
* PowerShellを管理者として開き、以下を実行:
winget install JohnMacFarlane.Pandoc
### **Tesseract OCR** (スキャン済みPDFの文字認識に必要)
* [UB Mannheimのサイト](https://www.google.com/search?q=https://github.com/UB-Mannheim/tesseract/wiki&authuser=1)からインストーラをダウンロードしてインストールしてください。
* インストール先（例: C:/Program Files/Tesseract-OCR/tesseract.exe）を覚えておいてください。
## **セットアップ手順**
Windowsの **PowerShell** または **コマンドプロンプト** で以下の手順を実行してください。
### **1. 仮想環境の作成と有効化**
プロジェクトフォルダに移動し、環境を分離するための仮想環境を作成します。
# 仮想環境の作成
python -m venv .venv
# 仮想環境の有効化 (PowerShell)
./.venv/Scripts/Activate.ps1
### **2. ライブラリの一括インストール**
requirements.txt を使用して、必要なライブラリをすべてインストールします。
python -m pip install -U pip setuptools wheel
python -m pip install -r requirements.txt

## **基本的な使い方**
### **1. Markdownへの変換 (基本)**
python doc2md/_asciidoc/_pandoc.py "パス/への/ファイル.pdf"
### **2. 特定のフォルダ内のファイルを一括変換する 「docs」というフォルダ内の対応ファイルをすべて変換します。**
****python doc2md/_asciidoc/_pandoc.py C:/Users/YourName/Documents/docs
### **3. サブフォルダも含めてすべて変換する（推奨） フォルダが階層構造になっている場合に便利です。**
****python doc2md/_asciidoc/_pandoc.py C:/Users/YourName/Documents/docs --recursive
### **4. ワイルドカードを使用する 特定の拡張子だけを狙ってディレクトリ指定風に動かすこともできます。**
****python doc2md/_asciidoc/_pandoc.py "docs//*.pdf"
### **5. Pandocを使用した高品質変換 (Word/Excel推奨)**
Pandocがインストールされている場合、以下のオプションでより正確な変換が可能です。
python doc2md/_asciidoc/_pandoc.py input.docx --use-pandoc --output-format md
### **6. OCR (文字認識) を使用する場合**
スキャンされたPDFなど、テキストデータを持たないファイルに有効です。
python doc2md/_asciidoc/_pandoc.py input.pdf --pdf-ocr --pdf-ocr-tesseract "C:/Program Files/Tesseract-OCR/tesseract.exe"

## **その他のオプション一覧**
| **オプション** | **説明** |
| --- | --- |
| -o, --output | 出力先フォルダの指定 (既定: output/_data) |
| --output-format | md (Markdown) または adoc (AsciiDoc) |
| --use-pandoc | Pandocを使用して高品質変換を行う |
| --pdf-ocr | OCR機能を有効にする |
| --overwrite | 上書き保存を許可する |
| --no-images | 画像の抽出を行わない |
| --json-summary | 変換結果の統計をJSON形式で保存する |
## **困ったときは**
* **「pymupdf.layout が見つからない」と出る:**pip install pymupdf-layout が成功しているか確認してください。
* **Wordの変換でエラーが出る:**--use-pandoc を外して試すか、Pandocが正しくインストールされパスが通っているか確認してください。
* **OCRの精度が低い:**--pdf-ocr-debug-images オプションを付けて、解析用画像（2値化済み）が正しく生成されているか確認してください。

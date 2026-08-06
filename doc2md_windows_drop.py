"""
Windows drag-and-drop launcher for doc2md_asciidoc_pandoc.py.

Usage:
  - Drop PDF/DOCX/XLSX/XLS files or folders onto the bundled exe.
  - Double-click the exe to choose files from a dialog.
  - Run from a terminal to pass the same CLI options as doc2md_asciidoc_pandoc.py.
"""

from __future__ import annotations

import os
import configparser
import shlex
import sys
from pathlib import Path
from typing import List

from doc2md_asciidoc_pandoc import main as doc2md_main


CONFIG_FILE_NAME = "doc2md-drop.ini"

WRAPPER_FLAGS = {"--no-pause", "--pause", "--no-config"}
OPTIONS_WITH_VALUE = {
    "-o",
    "--output",
    "--output-format",
    "--pandoc-bin",
    "--pdf-engine",
    "--pdf-max-pages",
    "--pdf-page-heading-template",
    "--pdf-ocr-strategy",
    "--pdf-ocr-min-chars",
    "--pdf-ocr-lang",
    "--pdf-ocr-dpi",
    "--pdf-ocr-psm",
    "--pdf-ocr-oem",
    "--pdf-ocr-tesseract",
    "--pdf-ocr-tessdata",
    "--pdf-ocr-threshold",
    "--pdf-ocr-contrast",
    "--excel-max-rows",
    "--excel-max-cols",
    "--excel-cell-max-chars",
    "--word-heading-style",
    "--json-summary",
}
VALUE_OPTION_MAP = {
    "output": "--output",
    "output_format": "--output-format",
    "pandoc_bin": "--pandoc-bin",
    "pdf_engine": "--pdf-engine",
    "pdf_max_pages": "--pdf-max-pages",
    "pdf_page_heading_template": "--pdf-page-heading-template",
    "pdf_ocr_strategy": "--pdf-ocr-strategy",
    "pdf_ocr_min_chars": "--pdf-ocr-min-chars",
    "pdf_ocr_lang": "--pdf-ocr-lang",
    "pdf_ocr_dpi": "--pdf-ocr-dpi",
    "pdf_ocr_psm": "--pdf-ocr-psm",
    "pdf_ocr_oem": "--pdf-ocr-oem",
    "pdf_ocr_tesseract": "--pdf-ocr-tesseract",
    "pdf_ocr_tessdata": "--pdf-ocr-tessdata",
    "pdf_ocr_threshold": "--pdf-ocr-threshold",
    "pdf_ocr_contrast": "--pdf-ocr-contrast",
    "excel_max_rows": "--excel-max-rows",
    "excel_max_cols": "--excel-max-cols",
    "excel_cell_max_chars": "--excel-cell-max-chars",
    "word_heading_style": "--word-heading-style",
    "json_summary": "--json-summary",
}
FLAG_OPTION_MAP = {
    "overwrite": "--overwrite",
    "recursive": "--recursive",
    "no_front_matter": "--no-front-matter",
    "no_sha256": "--no-sha256",
    "include_source_path": "--include-source-path",
    "no_images": "--no-images",
    "use_pandoc": "--use-pandoc",
    "pdf_force_page_headings": "--pdf-force-page-headings",
    "pdf_ocr": "--pdf-ocr",
    "pdf_ocr_debug_images": "--pdf-ocr-debug-images",
    "word_strip_empty_lines": "--word-strip-empty-lines",
    "word_keep_empty_lines": "--word-keep-empty-lines",
    "fail_fast": "--fail-fast",
    "debug": "--debug",
    "verbose": "--verbose",
    "quiet": "--quiet",
}


def _configure_console_encoding() -> None:
    if os.name == "nt":
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleCP(65001)
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass

    for stream_name in ("stdin", "stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except Exception:
                pass


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _choose_files() -> List[str]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        print(f"ファイル選択ダイアログを開けませんでした: {exc}")
        return []

    root = tk.Tk()
    root.withdraw()
    root.update()
    paths = filedialog.askopenfilenames(
        title="変換するファイルを選択",
        filetypes=[
            ("対応ファイル", "*.pdf *.docx *.xlsx *.xls"),
            ("PDF", "*.pdf"),
            ("Word", "*.docx"),
            ("Excel", "*.xlsx *.xls"),
            ("すべてのファイル", "*.*"),
        ],
    )
    root.destroy()
    return list(paths)


def _load_config_args(app_dir: Path) -> List[str]:
    config_path = app_dir / CONFIG_FILE_NAME
    if not config_path.exists():
        return []

    parser = configparser.ConfigParser()
    try:
        parser.read(config_path, encoding="utf-8")
    except Exception as exc:
        print(f"設定ファイルを読み込めませんでした: {config_path} ({exc})")
        return []

    if not parser.has_section("defaults"):
        return []

    section = parser["defaults"]
    args: List[str] = []

    for key, option in VALUE_OPTION_MAP.items():
        value = section.get(key, fallback="").strip()
        if value:
            args.extend([option, value])

    for key, option in FLAG_OPTION_MAP.items():
        raw = section.get(key, fallback="").strip()
        if not raw:
            continue
        try:
            enabled = section.getboolean(key)
        except ValueError:
            print(f"設定値が boolean ではありません: [defaults] {key}={raw}")
            continue
        if enabled:
            args.append(option)

    args.extend(_parse_default_args(section.get("default_args", fallback="")))
    if args:
        print(f"設定ファイル: {config_path}", flush=True)
    return args


def _parse_default_args(raw: str) -> List[str]:
    args: List[str] = []
    for line in raw.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith(("#", ";")):
            continue
        try:
            parts = shlex.split(stripped, posix=False)
        except ValueError as exc:
            print(f"default_args の解析に失敗しました: {stripped} ({exc})")
            continue
        args.extend(_strip_arg_quotes(part) for part in parts)
    return args


def _strip_arg_quotes(value: str) -> str:
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def _split_wrapper_args(argv: List[str]) -> tuple[List[str], bool, bool]:
    pause = True
    use_config = True
    forwarded: List[str] = []
    for arg in argv:
        if arg == "--no-pause":
            pause = False
        elif arg == "--pause":
            pause = True
        elif arg == "--no-config":
            use_config = False
        else:
            forwarded.append(arg)
    return forwarded, pause, use_config


def _normalize_input_paths(argv: List[str], original_cwd: Path) -> List[str]:
    normalized: List[str] = []
    expect_option_value = False
    positional_only = False

    for arg in argv:
        if positional_only:
            normalized.append(_absolute_arg(arg, original_cwd))
            continue

        if expect_option_value:
            normalized.append(arg)
            expect_option_value = False
            continue

        if arg == "--":
            normalized.append(arg)
            positional_only = True
            continue

        if arg.startswith("--") and "=" in arg:
            normalized.append(arg)
            continue

        if arg in OPTIONS_WITH_VALUE:
            normalized.append(arg)
            expect_option_value = True
            continue

        if arg.startswith("-"):
            normalized.append(arg)
            continue

        normalized.append(_absolute_arg(arg, original_cwd))

    return normalized


def _absolute_arg(arg: str, original_cwd: Path) -> str:
    path = Path(arg)
    if path.is_absolute():
        return arg
    return str(original_cwd / path)


def _display_output_dir(argv: List[str], app_dir: Path) -> Path:
    output = "output_data"
    expect_output = False

    for arg in argv:
        if expect_output:
            output = arg
            expect_output = False
            continue
        if arg == "--":
            break
        if arg in {"-o", "--output"}:
            expect_output = True
            continue
        if arg.startswith("--output="):
            output = arg.split("=", 1)[1]
            continue

    output_path = Path(output)
    if output_path.is_absolute():
        return output_path
    return app_dir / output_path


def _should_pause(argv: List[str], pause: bool) -> bool:
    if not pause:
        return False
    if any(arg in {"-h", "--help", "--version"} for arg in argv):
        return False
    return True


def _pause() -> None:
    try:
        input("\n処理が完了しました。Enter キーで閉じます...")
    except (EOFError, KeyboardInterrupt):
        pass


def main() -> int:
    _configure_console_encoding()
    original_cwd = Path.cwd()
    app_dir = _app_dir()

    user_argv, pause, use_config = _split_wrapper_args(sys.argv[1:])
    user_argv = _normalize_input_paths(user_argv, original_cwd)
    config_args = _load_config_args(app_dir) if use_config else []

    os.chdir(app_dir)
    if not user_argv:
        user_argv = _choose_files()
        if not user_argv:
            print("変換対象が選択されませんでした。")
            if pause:
                _pause()
            return 0

    argv = config_args + user_argv

    print("doc2md Windows launcher", flush=True)
    print(f"出力先: {_display_output_dir(argv, app_dir).resolve()}", flush=True)
    print("", flush=True)

    try:
        return doc2md_main(argv)
    finally:
        if _should_pause(argv, pause):
            _pause()


if __name__ == "__main__":
    raise SystemExit(main())

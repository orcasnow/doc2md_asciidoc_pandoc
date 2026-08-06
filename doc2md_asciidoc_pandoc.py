"""
doc2md_asciidoc_pandoc_final.py — PDF / Excel / Word(.docx) → Markdown / AsciiDoc 変換（OCR統合 + Pandoc対応）

主用途: RAG/LLM 向け前処理
- 出力: Markdown(.md) / AsciiDoc(.adoc)
- PDF: PyMuPDF / pymupdf4llm による本文抽出 +（可能なら）画像抽出
- PDF（スキャン等）: オプションで Tesseract OCR を実行し、「ページ見出し単位」で本文に統合
- Word: Pandoc（推奨）または mammoth+markdownify（画像抽出対応）
- Excel: Pandoc（推奨）または pandas

改善（リファクタリング適用）
- Enum化（output_format/pdf_engine/ocr_strategy など）
- OutputPlan による出力計画の一元化
- メタ情報生成の共通化（build_base_meta）
- 依存関係ヒントの一元管理（DEPENDENCY_HINTS）
- --use-pandoc 時の起動直後1回チェック + 変換時の失敗フォールバック
- ページ見出しテンプレート（--pdf-page-heading-template）と強制生成（--pdf-force-page-headings）
- 警告を構造化（WarningEntry）
- fail-fast（--fail-fast）と JSONサマリー（--json-summary）

注: これは単一ファイルで完結する設計を維持しています。
"""

from __future__ import annotations

import argparse
import datetime as _dt
import glob
import hashlib
import importlib
import inspect
import json
import logging
import os
import posixpath
import re
import shutil
import subprocess
import sys
import traceback
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, replace
from enum import Enum
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Pattern, Sequence, Tuple

__version__ = "2.0.1"

SUPPORTED_EXTS = {".pdf", ".docx", ".xlsx", ".xls"}


# -----------------------------
# Logging
# -----------------------------

logger = logging.getLogger("doc2md")


def _configure_logging(verbose: bool, quiet: bool) -> None:
    level = logging.INFO
    if verbose:
        level = logging.DEBUG
    if quiet:
        level = logging.WARNING
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _log_exception(message: str, exc: BaseException) -> None:
    if logger.isEnabledFor(logging.DEBUG):
        logger.exception(message)
    else:
        logger.error(f"{message}: {exc}")


# -----------------------------
# Errors
# -----------------------------


class Doc2MDError(Exception):
    """doc2md 全体の基底例外です。"""


class MissingDependencyError(Doc2MDError):
    """任意依存パッケージが不足している場合に送出します。"""

    def __init__(self, module: str, install_hint: str):
        super().__init__(f"Missing dependency: {module}. {install_hint}")
        self.module = module
        self.install_hint = install_hint


class UnsupportedFormatError(Doc2MDError):
    """未対応フォーマットが指定された場合に送出します。"""


class ConversionError(Doc2MDError):
    """変換処理の途中で発生したエラーを、ステージ情報付きで保持します。"""

    def __init__(self, stage: str, message: str, *, cause: Optional[BaseException] = None):
        super().__init__(f"[{stage}] {message}")
        self.stage = stage
        self.cause = cause


# -----------------------------
# Warning model (structured)
# -----------------------------


@dataclass(frozen=True)
class WarningEntry:
    code: str
    detail: str = ""

    def to_dict(self) -> Dict[str, str]:
        return {"code": self.code, "detail": self.detail}

    def __str__(self) -> str:
        if self.detail:
            return f"{self.code}: {self.detail}"
        return self.code


def _warn(warnings: List[WarningEntry], code: str, detail: str = "") -> None:
    warnings.append(WarningEntry(code=code, detail=detail))


# -----------------------------
# Enums
# -----------------------------


class OutputFormat(str, Enum):
    MD = "md"
    ADOC = "adoc"


class PdfEngine(str, Enum):
    AUTO = "auto"
    PYMUPDF4LLM = "pymupdf4llm"
    PYMUPDF_TEXT = "pymupdf_text"


class OcrStrategy(str, Enum):
    FALLBACK = "fallback"   # 本文が少ないときだけOCR→ページ見出し単位に統合
    INTEGRATE = "integrate" # 常にOCR→ページ見出し単位に統合
    APPEND = "append"       # OCR結果を末尾追記
    REPLACE = "replace"     # OCR結果のみ出力


class WordHeadingStyle(str, Enum):
    ATX = "ATX"       # '#'
    SETEXT = "SETEXT" # underlines


# -----------------------------
# Result / Options / OutputPlan
# -----------------------------


@dataclass
class ConversionResult:
    input_path: Path
    output_path: Optional[Path] = None
    success: bool = False
    warnings: List[WarningEntry] = field(default_factory=list)
    error: Optional[str] = None
    stats: Dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class OutputPlan:
    output_path: Path
    assets_dir: Path
    image_dir_abs: Path
    image_dir_rel: Path


@dataclass(frozen=True)
class ConversionOptions:
    output_dir: Path = Path("output_data")
    output_format: OutputFormat = OutputFormat.MD

    # Pandoc
    use_pandoc: bool = False
    pandoc_bin: str = "pandoc"

    overwrite: bool = False
    recursive: bool = False

    # RAG metadata
    add_front_matter: bool = True
    include_sha256: bool = True
    include_source_path: bool = False
    tool_name: str = "doc2md"
    converted_at_utc: bool = True

    # Batch behavior
    fail_fast: bool = False
    json_summary_path: Optional[Path] = None

    # Debug/behavior on exception
    debug_raise: bool = False

    # Assets/images
    extract_images: bool = True
    assets_dir_suffix: str = "_assets"
    images_subdir: str = "images"

    # PDF
    pdf_engine: PdfEngine = PdfEngine.AUTO
    pdf_max_pages: Optional[int] = None

    # Page heading controls (OCR integrate / RAG chunking)
    pdf_force_page_headings: bool = False
    pdf_page_heading_template: str = "## Page {n}"

    # PDF OCR
    pdf_ocr: bool = False
    pdf_ocr_strategy: OcrStrategy = OcrStrategy.FALLBACK
    pdf_ocr_min_chars: int = 200
    pdf_ocr_lang: str = "jpn"
    pdf_ocr_dpi: int = 300
    pdf_ocr_psm: int = 6
    pdf_ocr_oem: int = 3
    pdf_ocr_tesseract: str = ""       # tesseract.exe path
    pdf_ocr_tessdata_dir: str = ""    # tessdata directory
    pdf_ocr_debug_images: bool = False
    pdf_ocr_threshold: int = 160
    pdf_ocr_contrast: float = 1.8

    # Excel
    excel_max_rows: int = 500
    excel_max_cols: int = 50
    excel_cell_max_chars: int = 200

    # Word
    word_heading_style: WordHeadingStyle = WordHeadingStyle.ATX
    word_strip_empty_lines: bool = True

    @classmethod
    def from_args(cls, args: argparse.Namespace) -> "ConversionOptions":
        def enum_val(enum_cls, raw, default):
            try:
                return enum_cls(str(raw))
            except Exception:
                return default

        json_summary = getattr(args, "json_summary", None)
        json_summary_path = Path(json_summary) if json_summary else None

        return cls(
            output_dir=Path(getattr(args, "output", "output_data")),
            output_format=enum_val(OutputFormat, getattr(args, "output_format", "md"), OutputFormat.MD),
            use_pandoc=bool(getattr(args, "use_pandoc", False)),
            pandoc_bin=str(getattr(args, "pandoc_bin", "pandoc")),
            overwrite=bool(getattr(args, "overwrite", False)),
            recursive=bool(getattr(args, "recursive", False)),
            add_front_matter=not bool(getattr(args, "no_front_matter", False)),
            include_sha256=not bool(getattr(args, "no_sha256", False)),
            include_source_path=bool(getattr(args, "include_source_path", False)),
            debug_raise=bool(getattr(args, "debug", False)),
            extract_images=not bool(getattr(args, "no_images", False)),
            pdf_engine=enum_val(PdfEngine, getattr(args, "pdf_engine", "auto"), PdfEngine.AUTO),
            pdf_max_pages=getattr(args, "pdf_max_pages", None),
            pdf_force_page_headings=bool(getattr(args, "pdf_force_page_headings", False)),
            pdf_page_heading_template=str(getattr(args, "pdf_page_heading_template", "## Page {n}")),
            pdf_ocr=bool(getattr(args, "pdf_ocr", False)),
            pdf_ocr_strategy=enum_val(OcrStrategy, getattr(args, "pdf_ocr_strategy", "fallback"), OcrStrategy.FALLBACK),
            pdf_ocr_min_chars=int(getattr(args, "pdf_ocr_min_chars", 200)),
            pdf_ocr_lang=str(getattr(args, "pdf_ocr_lang", "jpn")),
            pdf_ocr_dpi=int(getattr(args, "pdf_ocr_dpi", 300)),
            pdf_ocr_psm=int(getattr(args, "pdf_ocr_psm", 6)),
            pdf_ocr_oem=int(getattr(args, "pdf_ocr_oem", 3)),
            pdf_ocr_tesseract=str(getattr(args, "pdf_ocr_tesseract", "")),
            pdf_ocr_tessdata_dir=str(getattr(args, "pdf_ocr_tessdata", "")),
            pdf_ocr_debug_images=bool(getattr(args, "pdf_ocr_debug_images", False)),
            pdf_ocr_threshold=int(getattr(args, "pdf_ocr_threshold", 160)),
            pdf_ocr_contrast=float(getattr(args, "pdf_ocr_contrast", 1.8)),
            excel_max_rows=int(getattr(args, "excel_max_rows", 500)),
            excel_max_cols=int(getattr(args, "excel_max_cols", 50)),
            excel_cell_max_chars=int(getattr(args, "excel_cell_max_chars", 200)),
            word_heading_style=enum_val(WordHeadingStyle, getattr(args, "word_heading_style", "ATX"), WordHeadingStyle.ATX),
            word_strip_empty_lines=bool(getattr(args, "word_strip_empty_lines", True)),
            fail_fast=bool(getattr(args, "fail_fast", False)),
            json_summary_path=json_summary_path,
        )


# -----------------------------
# Common helpers
# -----------------------------


_WILDCARD_RE = re.compile(r"[*?\[\]]")


def _ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _now_iso(utc: bool = True) -> str:
    if utc:
        return _dt.datetime.now(tz=_dt.timezone.utc).isoformat()
    return _dt.datetime.now().isoformat()


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _slugify_filename(name: str, max_len: int = 140) -> str:
    s = re.sub(r"\s+", " ", name).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    s = re.sub(r"[^\w\-. ()\[\]]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"_+", "_", s)
    return s[:max_len].strip(" ._") or "document"


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    parent = path.parent
    i = 2
    while True:
        cand = parent / f"{stem} ({i}){suffix}"
        if not cand.exists():
            return cand
        i += 1


def _rewrite_link_destinations(md: str, abs_dir: Path, rel_dir: Path) -> str:
    r"""
    Markdownリンク/画像の参照先を「絶対パス→相対パス」に寄せます。
    （Windowsパスの \ も / に揃える）
    """
    abs_s = str(abs_dir.resolve()).replace("\\", "/")
    rel_s = str(rel_dir).replace("\\", "/")
    if not abs_s.endswith("/"):
        abs_s += "/"
    if rel_s and not rel_s.endswith("/"):
        rel_s += "/"

    # images: ![alt](path "title") / [text](path)
    pattern = re.compile(r'(\!\[[^\]]*\]\(|\[[^\]]*\]\()(\<[^>]+\>|[^\)\s]+)([^\)]*)\)')
    def repl(m: re.Match) -> str:
        prefix = m.group(1)
        url_token = m.group(2)
        rest = m.group(3) or ""

        wrapped = url_token.startswith("<") and url_token.endswith(">")
        url = url_token[1:-1] if wrapped else url_token

        url_norm = url.replace("\\", "/")
        if url_norm.startswith(abs_s):
            url_norm = rel_s + url_norm[len(abs_s):]

        # Markdown の仕様上、URL 部分にスペースが含まれる場合は <...> で囲むのが安全です。
        if wrapped or (" " in url_norm):
            url_out = f"<{url_norm}>"
        else:
            url_out = url_norm

        return f"{prefix}{url_out}{rest})"
    return pattern.sub(repl, (md or "").replace("\\", "/"))


def _expand_inputs(inputs: Sequence[str], recursive: bool) -> List[Path]:
    paths: List[Path] = []

    def add_path(p: Path) -> None:
        if p.is_dir():
            it = p.rglob("*") if recursive else p.glob("*")
            for f in it:
                if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS:
                    paths.append(f)
        elif p.is_file():
            paths.append(p)
        else:
            raise FileNotFoundError(str(p))

    for s in inputs:
        if _WILDCARD_RE.search(s):
            for g in glob.glob(s, recursive=True):
                add_path(Path(g))
        else:
            add_path(Path(s))

    # de-dup keep order
    seen: set[str] = set()
    out: List[Path] = []
    for p in paths:
        rp = str(p.resolve())
        if rp not in seen:
            out.append(p)
            seen.add(rp)
    return out


def build_base_meta(
    options: ConversionOptions,
    source_type: str,
    input_path: Path,
    title: Optional[str] = None,
    extra: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    meta: Dict[str, Any] = {
        "tool": options.tool_name,
        "tool_version": __version__,
        # 互換性のため "type" も残す（旧README/既存データ）
        "type": source_type,
        "source_type": source_type,
        "source_file": input_path.name,
        "title": title or input_path.stem,
        "converted_at": _now_iso(options.converted_at_utc),
    }
    if options.include_source_path:
        meta["source_path"] = str(input_path.resolve())
    if options.include_sha256:
        meta["sha256"] = _sha256(input_path)
    if extra:
        meta.update(extra)
    return meta


# -----------------------------
# Dependencies (single source)
# -----------------------------

DEPENDENCY_HINTS: Dict[str, str] = {
    "fitz": "Install: pip install pymupdf",
    "pymupdf": "Install: pip install pymupdf",
    "pymupdf4llm": "Install: pip install pymupdf4llm pymupdf",
    "pymupdf.layout": "Install: pip install pymupdf-layout",
    "pytesseract": "Install: pip install pytesseract pillow",
    "PIL": "Install: pip install pillow",
    "pandas": "Install: pip install pandas openpyxl",
    "openpyxl": "Install: pip install openpyxl",
    "mammoth": "Install: pip install mammoth markdownify",
    "markdownify": "Install: pip install markdownify",
    "tabulate": "Install: pip install tabulate",
}


def _require(module_name: str, install_hint: Optional[str] = None):
    try:
        return importlib.import_module(module_name)
    except ImportError as e:
        hint = install_hint or DEPENDENCY_HINTS.get(module_name) or "Please install the missing dependency."
        raise MissingDependencyError(module_name, hint) from e


def activate_pymupdf_layout() -> None:
    """
    pymupdf.layout を有効化（入っていない環境でも落とさない）。
    これは PDF レイアウト抽出の品質改善に寄与する場合があります。
    """
    try:
        pymupdf_layout = _require("pymupdf.layout")
        if hasattr(pymupdf_layout, "activate"):
            pymupdf_layout.activate()
            logger.debug("pymupdf.layout.activate() enabled")
    except MissingDependencyError:
        # 明示的な利用箇所がないため、警告は出さない（READMEで案内済み）
        logger.debug("pymupdf.layout not installed; skip activation")
    except Exception as e:
        logger.debug(f"pymupdf.layout activation skipped: {e}")


# -----------------------------
# Pandoc
# -----------------------------


def _pandoc_version(pandoc_bin: str) -> str:
    try:
        cp = subprocess.run(
            [pandoc_bin, "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if cp.returncode != 0:
            return ""
        first = (cp.stdout or "").splitlines()[0].strip()
        return first
    except Exception:
        return ""


def _pandoc_convert_text(
    text: str,
    *,
    from_format: str,
    to_format: str,
    pandoc_bin: str,
    extra_args: Optional[List[str]] = None,
) -> str:
    args = [pandoc_bin, "-f", from_format, "-t", to_format]
    if extra_args:
        args.extend(extra_args)
    cp = subprocess.run(
        args,
        input=text,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if cp.returncode != 0:
        raise ConversionError("pandoc", (cp.stderr or cp.stdout or "pandoc failed").strip())
    return cp.stdout


def _pandoc_convert_file(
    input_path: Path,
    *,
    from_format: Optional[str],
    to_format: str,
    pandoc_bin: str,
    cwd: Optional[Path] = None,
    extract_media: Optional[str] = None,
    extra_args: Optional[List[str]] = None,
) -> str:
    args = [pandoc_bin]
    if from_format:
        args.extend(["-f", from_format])
    args.extend(["-t", to_format, str(input_path)])
    if extract_media:
        args.extend(["--extract-media", extract_media])
    if extra_args:
        args.extend(extra_args)
    cp = subprocess.run(
        args,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        cwd=str(cwd) if cwd else None,
    )
    if cp.returncode != 0:
        raise ConversionError("pandoc", (cp.stderr or cp.stdout or "pandoc failed").strip())
    return cp.stdout


# -----------------------------
# Writers (md/adoc)
# -----------------------------


def _write_markdown(output_path: Path, md_body: str, meta: Dict[str, Any], options: ConversionOptions) -> None:
    parts: List[str] = []
    if options.add_front_matter:
        parts.append("---")
        for k, v in meta.items():
            if v is None:
                continue
            parts.append(f"{k}: {json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v}")
        parts.append("---")
        parts.append("")
    parts.append((md_body or "").rstrip())
    output_path.write_text("\n".join(parts).rstrip() + "\n", encoding="utf-8")


def _write_asciidoc_raw(output_path: Path, adoc_body: str, meta: Dict[str, Any], options: ConversionOptions) -> None:
    parts: List[str] = []
    if options.add_front_matter:
        for k, v in meta.items():
            if v is None:
                continue
            if isinstance(v, (dict, list)):
                val = json.dumps(v, ensure_ascii=False)
            else:
                val = str(v)
            val = val.replace("\n", " ").strip()
            parts.append(f":{k}: {val}")
        parts.append("")  # header/body separator
    parts.append((adoc_body or "").rstrip())
    output_path.write_text("\n".join(parts).rstrip() + "\n", encoding="utf-8")


def _write_asciidoc(output_path: Path, md_body: str, meta: Dict[str, Any], options: ConversionOptions) -> None:
    parts: List[str] = []
    if options.add_front_matter:
        for k, v in meta.items():
            if v is None:
                continue
            if isinstance(v, (dict, list)):
                val = json.dumps(v, ensure_ascii=False)
            else:
                val = str(v)
            val = val.replace("\n", " ").strip()
            parts.append(f":{k}: {val}")
        parts.append("")
    parts.append(_md_to_asciidoc(md_body or ""))
    output_path.write_text("\n".join(parts).rstrip() + "\n", encoding="utf-8")


def _write_document(
    output_path: Path,
    content_md: str,
    meta: Dict[str, Any],
    options: ConversionOptions,
    warnings: Optional[List[WarningEntry]] = None,
) -> None:
    warnings = warnings if warnings is not None else []
    if options.output_format == OutputFormat.ADOC:
        if options.use_pandoc:
            try:
                adoc_body = _pandoc_convert_text(
                    content_md,
                    from_format="markdown",
                    to_format="asciidoc",
                    pandoc_bin=options.pandoc_bin,
                )
                _write_asciidoc_raw(output_path, adoc_body, meta, options)
                return
            except Exception as e:
                _warn(warnings, "pandoc_md_to_adoc_fallback", str(e))
        _write_asciidoc(output_path, content_md, meta, options)
    else:
        _write_markdown(output_path, content_md, meta, options)


# -----------------------------
# Lightweight Markdown -> AsciiDoc
# (kept intentionally simple; prefer pandoc when available)
# -----------------------------


def _md_to_asciidoc(md: str) -> str:
    """
    “軽量” Markdown→AsciiDoc 変換。
    - Pandoc が使えない環境向けのフォールバックであり、完全互換は目指しません。
    """
    text = md or ""
    # code fences
    text = re.sub(r"^```(\w+)?\s*$", lambda m: f"[source,{m.group(1)}]\n----" if m.group(1) else "----", text, flags=re.MULTILINE)
    text = re.sub(r"^```$", "----", text, flags=re.MULTILINE)

    # headings
    def h_repl(m: re.Match) -> str:
        level = len(m.group(1))
        title = m.group(2).strip()
        return f"{'=' * level} {title}"
    text = re.sub(r"^(#{1,6})\s+(.*)$", h_repl, text, flags=re.MULTILINE)

    # bold/italic (best-effort)
    text = re.sub(r"\*\*(.+?)\*\*", r"*\1*", text)
    text = re.sub(r"__(.+?)__", r"*\1*", text)
    text = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"_\1_", text)
    text = re.sub(r"(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)", r"_\1_", text)

    # links/images
    text = re.sub(r"!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)", r"image::\2[\1]", text)
    text = re.sub(r"\[([^\]]+)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)", r"link:\2[\1]", text)

    # tables (very naive): keep markdown tables as-is; AsciiDoc will show them as text
    return text.rstrip() + "\n"


# -----------------------------
# Page headings / OCR integration
# -----------------------------


@dataclass
class PageSection:
    heading_line: str
    level: int
    page_no: int
    content: str


def format_page_heading(template: str, n: int) -> str:
    if "{n}" not in template:
        # 互換: 無い場合は末尾に番号を付与
        return f"{template.strip()} {n}"
    return template.format(n=n).strip()


def compile_page_heading_re(template: str) -> Pattern[str]:
    """
    template例:
      - "## Page {n}"
      - "## ページ {n}"
    """
    # templateからラベル部分を推定（{n} の前後）
    t = template.strip()
    # 先頭の '#'
    t_wo_hash = re.sub(r"^#{1,6}\s*", "", t)
    if "{n}" not in t_wo_hash:
        label = t_wo_hash.strip()
    else:
        label = t_wo_hash.split("{n}", 1)[0].strip()

    if label in ("Page", "ページ"):
        label_pat = r"(?:Page|ページ)"
    else:
        label_pat = re.escape(label) if label else r"(?:Page|ページ)"

    # Heading: allow any level 1-6; accept optional ":"/"："
    return re.compile(
        rf"^(?P<h>#{1,6})\s*{label_pat}\s*(?P<n>\d+)\s*[:：]?\s*$",
        re.MULTILINE,
    )


def parse_page_sections(md_body: str, page_heading_re: Pattern[str]) -> Tuple[str, List[PageSection]]:
    md_body = md_body or ""
    matches = list(page_heading_re.finditer(md_body))
    if not matches:
        return md_body, []
    preamble = md_body[: matches[0].start()]
    sections: List[PageSection] = []
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(md_body)
        heading = m.group("h")
        n = int(m.group("n"))
        content = md_body[start:end].lstrip("\n")
        sections.append(PageSection(heading_line=m.group(0).strip(), level=len(heading), page_no=n, content=content))
    return preamble, sections


def integrate_ocr_into_sections(
    preamble: str,
    sections: List[PageSection],
    ocr_pages: Dict[int, str],
    *,
    label: str = "OCR（Tesseract）",
) -> str:
    parts: List[str] = []
    if preamble.strip():
        parts.append(preamble.rstrip() + "\n")
    for sec in sections:
        parts.append(sec.heading_line.rstrip() + "\n\n")
        body = sec.content.rstrip() + "\n"
        parts.append(body + "\n")
        if sec.page_no in ocr_pages:
            parts.append(f"{'#' * min(sec.level + 1, 6)} {label}\n\n")
            parts.append(ocr_pages[sec.page_no].rstrip() + "\n\n")
    return "".join(parts).rstrip() + "\n"


# -----------------------------
# OCR
# -----------------------------


@dataclass(frozen=True)
class OcrConfig:
    lang: str = "jpn"
    dpi: int = 300
    psm: int = 6
    oem: int = 3
    tesseract_cmd: str = ""
    tessdata_dir: str = ""
    debug_images: bool = False
    threshold: int = 160
    contrast: float = 1.8


def _run_ocr_pages(
    input_path: Path,
    assets_dir: Path,
    options: ConversionOptions,
    warnings: List[WarningEntry],
) -> Tuple[Dict[int, str], Dict[str, Any]]:
    """
    Tesseract OCR をページ単位で実行して {page_no: text} を返します（page_no は1始まり）。
    """
    pytesseract = _require("pytesseract")
    pil = _require("PIL")
    Image = pil.Image
    ImageEnhance = pil.ImageEnhance

    fitz = _require("fitz")
    cfg = OcrConfig(
        lang=options.pdf_ocr_lang,
        dpi=options.pdf_ocr_dpi,
        psm=options.pdf_ocr_psm,
        oem=options.pdf_ocr_oem,
        tesseract_cmd=options.pdf_ocr_tesseract,
        tessdata_dir=options.pdf_ocr_tessdata_dir,
        debug_images=options.pdf_ocr_debug_images,
        threshold=options.pdf_ocr_threshold,
        contrast=options.pdf_ocr_contrast,
    )

    if cfg.tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = cfg.tesseract_cmd

    doc = fitz.open(str(input_path))
    page_count = int(doc.page_count)
    max_pages = options.pdf_max_pages
    used_pages = page_count if max_pages is None else min(page_count, int(max_pages))

    debug_dir = assets_dir / "ocr_debug_images"
    if cfg.debug_images:
        _ensure_dir(debug_dir)

    ocr_pages: Dict[int, str] = {}

    tconf = f"--oem {cfg.oem} --psm {cfg.psm}"
    if cfg.tessdata_dir:
        tconf += f" --tessdata-dir \"{cfg.tessdata_dir}\""

    for i in range(used_pages):
        page = doc.load_page(i)
        zoom = cfg.dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # preprocess: contrast + threshold (simple)
        try:
            enh = ImageEnhance.Contrast(img)
            img2 = enh.enhance(cfg.contrast)
            img2 = img2.convert("L")  # grayscale
            img2 = img2.point(lambda p: 255 if p > cfg.threshold else 0)
        except Exception as e:
            _warn(warnings, "ocr_preprocess_failed", str(e))
            img2 = img.convert("L")

        if cfg.debug_images:
            try:
                img2.save(debug_dir / f"page_{i+1:04d}.png")
            except Exception as e:
                _warn(warnings, "ocr_debug_image_save_failed", str(e))

        try:
            text = pytesseract.image_to_string(img2, lang=cfg.lang, config=tconf)
        except Exception as e:
            _warn(warnings, "ocr_failed", str(e))
            text = ""
        ocr_pages[i + 1] = (text or "").strip()

    doc.close()

    stats = {
        "ocr_pages": used_pages,
        "ocr_lang": cfg.lang,
        "ocr_dpi": cfg.dpi,
        "ocr_psm": cfg.psm,
        "ocr_oem": cfg.oem,
    }
    return ocr_pages, stats


def _estimate_text_chars(md: str) -> int:
    s = re.sub(r"\s+", "", md or "")
    s = re.sub(r"[#`>*_\-\[\]\(\)!|]+", "", s)
    return len(s)


# -----------------------------
# Output planning
# -----------------------------


def plan_output(input_path: Path, options: ConversionOptions) -> OutputPlan:
    out_dir = options.output_dir
    _ensure_dir(out_dir)

    safe_stem = _slugify_filename(input_path.stem)
    ext = ".adoc" if options.output_format == OutputFormat.ADOC else ".md"
    out_path = out_dir / f"{safe_stem}{ext}"
    assets_dir = out_dir / f"{safe_stem}{options.assets_dir_suffix}"

    if out_path.exists() and not options.overwrite:
        out_path = _unique_path(out_path)
        safe_stem2 = out_path.stem
        assets_dir = out_dir / f"{safe_stem2}{options.assets_dir_suffix}"

    image_dir_abs = assets_dir / options.images_subdir
    image_dir_rel = Path(assets_dir.name) / options.images_subdir
    return OutputPlan(output_path=out_path, assets_dir=assets_dir, image_dir_abs=image_dir_abs, image_dir_rel=image_dir_rel)


# -----------------------------
# Converters
# -----------------------------


class BaseConverter:
    kind: str = "unknown"

    def convert(self, input_path: Path, plan: OutputPlan, options: ConversionOptions) -> ConversionResult:
        raise NotImplementedError


class PDFConverter(BaseConverter):
    kind = "pdf"

    def _pdf_metadata(self, input_path: Path) -> Dict[str, Any]:
        try:
            fitz = _require("fitz")
            doc = fitz.open(str(input_path))
            meta = doc.metadata or {}
            out = {
                "pdf_title": meta.get("title") or "",
                "pdf_author": meta.get("author") or "",
                "pdf_subject": meta.get("subject") or "",
                "pdf_keywords": meta.get("keywords") or "",
                "page_count": int(doc.page_count),
            }
            doc.close()
            return out
        except Exception:
            return {}

    def _pymupdf_text(self, input_path: Path, options: ConversionOptions) -> Tuple[str, Dict[str, Any]]:
        fitz = _require("fitz")
        doc = fitz.open(str(input_path))
        max_pages = options.pdf_max_pages
        page_count = int(doc.page_count)
        used_pages = page_count if max_pages is None else min(page_count, int(max_pages))

        parts: List[str] = []
        for i in range(used_pages):
            page = doc.load_page(i)
            text = page.get_text("text") or ""
            parts.append("\n\n" + format_page_heading(options.pdf_page_heading_template, i + 1) + "\n\n")
            parts.append(text.strip() + "\n")
        doc.close()
        stats = {"engine": "pymupdf_text", "page_count": page_count, "pages_converted": used_pages}
        return "".join(parts).strip() + "\n", stats

    def _pymupdf4llm(self, input_path: Path, plan: OutputPlan, options: ConversionOptions) -> Tuple[str, Dict[str, Any]]:
        pymupdf4llm = _require("pymupdf4llm")
        sig = inspect.signature(pymupdf4llm.to_markdown)
        params = sig.parameters

        kwargs: Dict[str, Any] = {}
        image_output_requested = False
        if "write_images" in params:
            kwargs["write_images"] = bool(options.extract_images)
        if "image_path" in params:
            kwargs["image_path"] = str(plan.image_dir_abs)
            image_output_requested = image_output_requested or bool(options.extract_images)
        elif "image_dir" in params:
            kwargs["image_dir"] = str(plan.image_dir_abs)
            image_output_requested = image_output_requested or bool(options.extract_images)
        elif "output_dir" in params:
            kwargs["output_dir"] = str(plan.image_dir_abs)
            image_output_requested = image_output_requested or bool(options.extract_images)
        if "pages" in params and options.pdf_max_pages is not None:
            kwargs["pages"] = list(range(int(options.pdf_max_pages)))

        if image_output_requested:
            _ensure_dir(plan.image_dir_abs)

        logger.debug(f"pymupdf4llm.to_markdown kwargs: {kwargs}")
        md = pymupdf4llm.to_markdown(str(input_path), **kwargs)
        return md, {"engine": "pymupdf4llm"}

    def convert(self, input_path: Path, plan: OutputPlan, options: ConversionOptions) -> ConversionResult:
        res = ConversionResult(input_path=input_path, output_path=plan.output_path)
        warnings: List[WarningEntry] = []

        try:
            meta_extra = self._pdf_metadata(input_path)
            title = meta_extra.get("pdf_title") or input_path.stem
            meta = build_base_meta(options, self.kind, input_path, title=title, extra=meta_extra)

            # Pandoc direct (best effort) when OCR is not required
            if options.use_pandoc and not options.pdf_ocr:
                try:
                    to_fmt = "gfm" if options.output_format == OutputFormat.MD else "asciidoc"
                    body = _pandoc_convert_file(
                        input_path.resolve(),
                        from_format="pdf",
                        to_format=to_fmt,
                        pandoc_bin=options.pandoc_bin,
                        cwd=options.output_dir,
                    )
                    meta["engine"] = "pandoc"
                    meta["pandoc_version"] = _pandoc_version(options.pandoc_bin)

                    if options.output_format == OutputFormat.ADOC:
                        _write_asciidoc_raw(plan.output_path, body, meta, options)
                    else:
                        _write_markdown(plan.output_path, body, meta, options)
                    res.success = True
                    res.warnings = warnings
                    res.stats.update({"engine": "pandoc"})
                    return res
                except Exception as e:
                    _warn(warnings, "pandoc_pdf_fallback", str(e))

            # Extract body
            engine = options.pdf_engine
            body_md = ""
            stats: Dict[str, Any] = {}

            # If force page headings, prefer text engine (guaranteed headings)
            if options.pdf_force_page_headings and engine in (PdfEngine.AUTO, PdfEngine.PYMUPDF4LLM):
                _warn(warnings, "pdf_engine_forced_pymupdf_text", f"requested headings via template={options.pdf_page_heading_template!r}")
                engine = PdfEngine.PYMUPDF_TEXT

            if engine in (PdfEngine.AUTO, PdfEngine.PYMUPDF4LLM):
                try:
                    body_md, stats = self._pymupdf4llm(input_path, plan, options)
                except MissingDependencyError as e:
                    if engine == PdfEngine.PYMUPDF4LLM:
                        raise
                    _warn(warnings, "pymupdf4llm_missing_fallback_to_text", str(e))
                except Exception as e:
                    if engine == PdfEngine.PYMUPDF4LLM:
                        raise
                    _warn(warnings, "pymupdf4llm_failed_fallback_to_text", str(e))

            if not body_md:
                body_md, stats = self._pymupdf_text(input_path, options)

            # Normalize image links to relative
            if options.extract_images:
                body_md = _rewrite_link_destinations(body_md, plan.image_dir_abs, plan.image_dir_rel)

            # OCR
            page_heading_re = compile_page_heading_re(options.pdf_page_heading_template)

            if options.pdf_ocr:
                # Fallback decision
                need_ocr = True
                if options.pdf_ocr_strategy == OcrStrategy.FALLBACK:
                    txt_chars = _estimate_text_chars(body_md)
                    need_ocr = txt_chars < options.pdf_ocr_min_chars
                    res.stats["text_chars_estimate"] = txt_chars
                    res.stats["ocr_triggered"] = bool(need_ocr)

                if need_ocr or options.pdf_ocr_strategy in (OcrStrategy.INTEGRATE, OcrStrategy.APPEND, OcrStrategy.REPLACE):
                    ocr_pages, ocr_stats = _run_ocr_pages(input_path, plan.assets_dir, options, warnings)
                    res.stats.update(ocr_stats)

                    label = "OCR（Tesseract）"

                    if options.pdf_ocr_strategy in (OcrStrategy.FALLBACK, OcrStrategy.INTEGRATE):
                        preamble, sections = parse_page_sections(body_md, page_heading_re)
                        if sections:
                            body_md = integrate_ocr_into_sections(preamble, sections, ocr_pages, label=label)
                            res.stats["ocr_integrated"] = True
                        else:
                            _warn(warnings, "ocr_integrate_failed_no_page_headings", "fallback to append")
                            options2 = replace(options, pdf_ocr_strategy=OcrStrategy.APPEND)
                            options = options2
                            # fallthrough to append

                    if options.pdf_ocr_strategy == OcrStrategy.APPEND:
                        # append all pages (with headings for readability)
                        parts: List[str] = [body_md.rstrip() + "\n\n", f"## {label}\n\n"]
                        for pno in sorted(ocr_pages.keys()):
                            parts.append(format_page_heading(options.pdf_page_heading_template, pno) + "\n\n")
                            parts.append((ocr_pages[pno] or "").rstrip() + "\n\n")
                        body_md = "".join(parts).rstrip() + "\n"
                        res.stats["ocr_appended"] = True

                    if options.pdf_ocr_strategy == OcrStrategy.REPLACE:
                        parts: List[str] = []
                        for pno in sorted(ocr_pages.keys()):
                            parts.append(format_page_heading(options.pdf_page_heading_template, pno) + "\n\n")
                            parts.append((ocr_pages[pno] or "").rstrip() + "\n\n")
                        body_md = "".join(parts).rstrip() + "\n"
                        res.stats["ocr_replaced"] = True

            meta["engine"] = stats.get("engine", "unknown")
            if options.use_pandoc:
                meta["pandoc_version"] = _pandoc_version(options.pandoc_bin) or ""

            _write_document(plan.output_path, body_md, meta, options, warnings)

            res.success = True
            res.warnings = warnings
            res.stats.update(stats)
            return res

        except Exception as e:
            res.success = False
            res.error = str(e)
            res.warnings = warnings
            if options.debug_raise:
                raise
            return res


class WordConverter(BaseConverter):
    kind = "word"

    def _docx_has_media(self, input_path: Path) -> bool:
        try:
            with zipfile.ZipFile(input_path) as zf:
                return any(
                    name.startswith("word/media/") and not name.endswith("/")
                    for name in zf.namelist()
                )
        except Exception as e:
            logger.debug(f"docx media detection skipped: {e}")
            return False

    def convert(self, input_path: Path, plan: OutputPlan, options: ConversionOptions) -> ConversionResult:
        res = ConversionResult(input_path=input_path, output_path=plan.output_path)
        warnings: List[WarningEntry] = []

        try:
            meta = build_base_meta(options, self.kind, input_path, title=input_path.stem)

            # Pandoc direct (preferred)
            if options.use_pandoc:
                try:
                    to_fmt = "gfm" if options.output_format == OutputFormat.MD else "asciidoc"
                    extract_media = None
                    if options.extract_images and self._docx_has_media(input_path):
                        extract_media = str(Path(plan.assets_dir.name) / options.images_subdir).replace("\\", "/")

                    body = _pandoc_convert_file(
                        input_path.resolve(),
                        from_format="docx",
                        to_format=to_fmt,
                        pandoc_bin=options.pandoc_bin,
                        cwd=options.output_dir,
                        extract_media=extract_media,
                    )
                    meta["engine"] = "pandoc"
                    meta["pandoc_version"] = _pandoc_version(options.pandoc_bin)

                    if options.output_format == OutputFormat.ADOC:
                        _write_asciidoc_raw(plan.output_path, body, meta, options)
                    else:
                        _write_markdown(plan.output_path, body, meta, options)
                    res.success = True
                    res.warnings = warnings
                    res.stats.update({"engine": "pandoc"})
                    return res
                except Exception as e:
                    _warn(warnings, "pandoc_docx_fallback", str(e))

            # Fallback: mammoth + markdownify
            mammoth = _require("mammoth")
            markdownify = _require("markdownify")

            def _convert_image(image):
                # image.open() returns bytes-like object
                with image.open() as img_bytes:
                    data = img_bytes.read()
                # determine ext
                content_type = getattr(image, "content_type", "") or ""
                ext = ".png"
                if "jpeg" in content_type or "jpg" in content_type:
                    ext = ".jpg"
                elif "gif" in content_type:
                    ext = ".gif"
                fname = f"image_{hashlib.md5(data).hexdigest()[:12]}{ext}"
                out_abs = plan.image_dir_abs / fname
                out_rel = plan.image_dir_rel / fname
                _ensure_dir(plan.image_dir_abs)
                out_abs.write_bytes(data)
                return {"src": str(out_rel).replace("\\", "/")}

            with input_path.open("rb") as f:
                if options.extract_images:
                    result = mammoth.convert_to_html(f, convert_image=_convert_image)
                else:
                    result = mammoth.convert_to_html(f)

            if getattr(result, "messages", None):
                for m in result.messages:
                    _warn(warnings, "mammoth_message", str(m))

            html = result.value
            md = markdownify.markdownify(html, heading_style=options.word_heading_style.value)
            if options.word_strip_empty_lines:
                md = "\n".join([ln.rstrip() for ln in md.splitlines() if ln.strip() != ""]).strip() + "\n"

            md = _rewrite_link_destinations(md, plan.image_dir_abs, plan.image_dir_rel)

            meta["engine"] = "mammoth"
            _write_document(plan.output_path, md, meta, options, warnings)

            res.success = True
            res.warnings = warnings
            res.stats.update({"engine": "mammoth"})
            return res

        except Exception as e:
            res.success = False
            res.error = str(e)
            res.warnings = warnings
            if options.debug_raise:
                raise
            return res


class ExcelConverter(BaseConverter):
    kind = "excel"

    _XLSX_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    _XLSX_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    _XLSX_OFFICE_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    def _xlsx_rels_path(self, part_path: str) -> str:
        part_dir = posixpath.dirname(part_path)
        part_name = posixpath.basename(part_path)
        return posixpath.join(part_dir, "_rels", f"{part_name}.rels")

    def _xlsx_join_target(self, source_dir: str, target: str) -> str:
        target = (target or "").replace("\\", "/")
        if target.startswith("/"):
            return posixpath.normpath(target.lstrip("/"))
        return posixpath.normpath(posixpath.join(source_dir, target))

    def _xlsx_relationships(self, zf: zipfile.ZipFile, rels_path: str) -> Dict[str, Dict[str, str]]:
        if rels_path not in zf.namelist():
            return {}
        source_dir = posixpath.dirname(posixpath.dirname(rels_path))
        try:
            root = ET.fromstring(zf.read(rels_path))
        except Exception:
            return {}

        rels: Dict[str, Dict[str, str]] = {}
        for rel in root.findall(f"{self._XLSX_REL_NS}Relationship"):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")
            if not rid or not target or rel.attrib.get("TargetMode") == "External":
                continue
            rels[rid] = {
                "target": self._xlsx_join_target(source_dir, target),
                "type": rel.attrib.get("Type", ""),
            }
        return rels

    def _xlsx_sheet_paths(self, zf: zipfile.ZipFile) -> Dict[str, str]:
        if "xl/workbook.xml" not in zf.namelist():
            return {}
        try:
            root = ET.fromstring(zf.read("xl/workbook.xml"))
        except Exception:
            return {}

        workbook_rels = self._xlsx_relationships(zf, "xl/_rels/workbook.xml.rels")
        sheets: Dict[str, str] = {}
        for sheet in root.findall(f".//{self._XLSX_MAIN_NS}sheet"):
            rid = sheet.attrib.get(f"{self._XLSX_OFFICE_REL_NS}id")
            name = sheet.attrib.get("name", "")
            target = workbook_rels.get(rid or "", {}).get("target")
            if target and name:
                sheets[target] = name
        return sheets

    def _xlsx_cell_ref(self, col_zero_based: int, row_zero_based: int) -> str:
        n = col_zero_based + 1
        letters = ""
        while n:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        return f"{letters}{row_zero_based + 1}"

    def _xlsx_drawing_occurrences(
        self,
        zf: zipfile.ZipFile,
        drawing_path: str,
        sheet_name: str,
    ) -> List[Dict[str, str]]:
        drawing_rels = self._xlsx_relationships(zf, self._xlsx_rels_path(drawing_path))
        if drawing_path not in zf.namelist():
            return []

        try:
            root = ET.fromstring(zf.read(drawing_path))
        except Exception:
            return []

        ns = {
            "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        occurrences: List[Dict[str, str]] = []
        for anchor in root:
            tag = anchor.tag.rsplit("}", 1)[-1]
            if tag not in {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"}:
                continue

            cell = ""
            from_el = anchor.find("xdr:from", ns)
            if from_el is not None:
                col_el = from_el.find("xdr:col", ns)
                row_el = from_el.find("xdr:row", ns)
                try:
                    col = int(col_el.text) if col_el is not None and col_el.text is not None else 0
                    row = int(row_el.text) if row_el is not None and row_el.text is not None else 0
                    cell = self._xlsx_cell_ref(col, row)
                except ValueError:
                    cell = ""

            blip = anchor.find(".//a:blip", ns)
            rid = ""
            if blip is not None:
                rid = blip.attrib.get(f"{self._XLSX_OFFICE_REL_NS}embed", "") or blip.attrib.get(f"{self._XLSX_OFFICE_REL_NS}link", "")
            media_path = drawing_rels.get(rid, {}).get("target")
            if media_path and media_path.startswith("xl/media/"):
                occurrences.append({"sheet": sheet_name, "cell": cell, "source": media_path})
        return occurrences

    def _xlsx_image_occurrences(self, zf: zipfile.ZipFile) -> List[Dict[str, str]]:
        sheet_paths = self._xlsx_sheet_paths(zf)
        occurrences: List[Dict[str, str]] = []
        seen: set[Tuple[str, str, str]] = set()

        for sheet_path, sheet_name in sheet_paths.items():
            sheet_rels = self._xlsx_relationships(zf, self._xlsx_rels_path(sheet_path))
            for rel in sheet_rels.values():
                if "drawing" not in rel.get("type", ""):
                    continue
                drawing_path = rel.get("target", "")
                for occ in self._xlsx_drawing_occurrences(zf, drawing_path, sheet_name):
                    key = (occ.get("sheet", ""), occ.get("cell", ""), occ.get("source", ""))
                    if key not in seen:
                        occurrences.append(occ)
                        seen.add(key)

        if occurrences:
            return occurrences

        return [
            {"sheet": "", "cell": "", "source": name}
            for name in zf.namelist()
            if name.startswith("xl/media/") and not name.endswith("/")
        ]

    def _extract_xlsx_images(
        self,
        input_path: Path,
        plan: OutputPlan,
        options: ConversionOptions,
        warnings: List[WarningEntry],
    ) -> Tuple[str, str, Dict[str, Any]]:
        if input_path.suffix.lower() != ".xlsx" or not options.extract_images:
            return "", "", {"excel_images": 0}

        try:
            with zipfile.ZipFile(input_path) as zf:
                occurrences = self._xlsx_image_occurrences(zf)
                media_sources = []
                seen_sources: set[str] = set()
                for occ in occurrences:
                    source = occ.get("source", "")
                    if source and source not in seen_sources:
                        media_sources.append(source)
                        seen_sources.add(source)

                if not media_sources:
                    return "", "", {"excel_images": 0}

                _ensure_dir(plan.image_dir_abs)
                rel_by_source: Dict[str, str] = {}
                for source in media_sources:
                    suffix = Path(source).suffix or ".bin"
                    stem = _slugify_filename(Path(source).stem)
                    out_abs = _unique_path(plan.image_dir_abs / f"{stem}{suffix}")
                    out_abs.write_bytes(zf.read(source))
                    rel_by_source[source] = str((plan.image_dir_rel / out_abs.name)).replace("\\", "/")
        except Exception as e:
            _warn(warnings, "excel_image_extract_failed", str(e))
            return "", "", {"excel_images": 0}

        md_parts = ["## Extracted Images\n"]
        adoc_parts = ["== Extracted Images\n"]
        current_sheet: Optional[str] = None
        for i, occ in enumerate(occurrences, start=1):
            source = occ.get("source", "")
            rel_path = rel_by_source.get(source)
            if not rel_path:
                continue
            sheet = occ.get("sheet", "")
            cell = occ.get("cell", "")
            if sheet and sheet != current_sheet:
                md_parts.append(f"### Sheet: {sheet}\n")
                adoc_parts.append(f"=== Sheet: {sheet}\n")
                current_sheet = sheet
            label_bits = [bit for bit in (sheet, cell, f"image {i}") if bit]
            label = " ".join(label_bits)
            md_parts.append(f"![{label}]({rel_path})\n\n")
            adoc_parts.append(f"image::{rel_path}[{label}]\n\n")

        return (
            "\n".join(md_parts).rstrip() + "\n",
            "\n".join(adoc_parts).rstrip() + "\n",
            {"excel_images": len(media_sources)},
        )

    def convert(self, input_path: Path, plan: OutputPlan, options: ConversionOptions) -> ConversionResult:
        res = ConversionResult(input_path=input_path, output_path=plan.output_path)
        warnings: List[WarningEntry] = []

        try:
            meta = build_base_meta(options, self.kind, input_path, title=input_path.stem)
            image_md, image_adoc, image_stats = self._extract_xlsx_images(input_path, plan, options, warnings)
            if image_stats.get("excel_images"):
                meta["image_count"] = image_stats["excel_images"]

            # Pandoc direct for xlsx when enabled
            if options.use_pandoc and input_path.suffix.lower() == ".xlsx":
                try:
                    to_fmt = "gfm" if options.output_format == OutputFormat.MD else "asciidoc"
                    body = _pandoc_convert_file(
                        input_path.resolve(),
                        from_format="xlsx",
                        to_format=to_fmt,
                        pandoc_bin=options.pandoc_bin,
                        cwd=options.output_dir,
                    )
                    if not body.strip():
                        raise ConversionError("pandoc", "xlsx conversion returned empty output")
                    meta["engine"] = "pandoc"
                    meta["pandoc_version"] = _pandoc_version(options.pandoc_bin)
                    if options.output_format == OutputFormat.ADOC:
                        if image_adoc:
                            body = body.rstrip() + "\n\n" + image_adoc
                    else:
                        if image_md:
                            body = body.rstrip() + "\n\n" + image_md
                    if options.output_format == OutputFormat.ADOC:
                        _write_asciidoc_raw(plan.output_path, body, meta, options)
                    else:
                        _write_markdown(plan.output_path, body, meta, options)
                    res.success = True
                    res.warnings = warnings
                    res.stats.update({"engine": "pandoc", **image_stats})
                    return res
                except Exception as e:
                    _warn(warnings, "pandoc_xlsx_fallback", str(e))

            # Fallback: pandas (xlsx/xls)
            pd = _require("pandas")
            # pandas.DataFrame.to_markdown は内部で tabulate に依存します。
            # tabulate が無い場合は例外に頼らず明示的に to_string へフォールバックします。
            tabulate_ok = True
            try:
                _require("tabulate")
            except MissingDependencyError:
                tabulate_ok = False
                logger.debug("tabulate not installed; fallback to DataFrame.to_string for Excel tables")

            # read excel
            try:
                xl = pd.read_excel(input_path, sheet_name=None, dtype=str, na_filter=False)
            except Exception as e:
                raise ConversionError("excel", f"read_excel failed: {e}", cause=e)

            parts: List[str] = []
            sheet_names = list(xl.keys())
            for sname, df in xl.items():
                # limit
                df2 = df.iloc[: options.excel_max_rows, : options.excel_max_cols].copy()
                # truncate cells
                def trunc(v: Any) -> str:
                    sv = "" if v is None else str(v)
                    if len(sv) > options.excel_cell_max_chars:
                        return sv[: options.excel_cell_max_chars] + "…"
                    return sv
                df2 = df2.map(trunc) if hasattr(df2, "map") else df2.applymap(trunc)
                parts.append(f"## Sheet: {sname}\n")
                if tabulate_ok:
                    try:
                        parts.append(df2.to_markdown(index=False) + "\n\n")
                    except Exception as e:
                        # fallback to plain table
                        _warn(warnings, "excel_to_markdown_failed_fallback_to_string", str(e))
                        parts.append(df2.to_string(index=False) + "\n\n")
                else:
                    parts.append(df2.to_string(index=False) + "\n\n")

            meta["engine"] = "pandas"
            meta["sheet_names"] = sheet_names

            md = "\n".join(parts).rstrip() + "\n"
            if image_md:
                md = md.rstrip() + "\n\n" + image_md
            _write_document(plan.output_path, md, meta, options, warnings)

            res.success = True
            res.warnings = warnings
            res.stats.update({"engine": "pandas", "sheets": len(sheet_names), **image_stats})
            return res

        except Exception as e:
            res.success = False
            res.error = str(e)
            res.warnings = warnings
            if options.debug_raise:
                raise
            return res


# -----------------------------
# Orchestrator
# -----------------------------


class DocumentConverter:
    def __init__(self) -> None:
        self._handlers: Dict[str, BaseConverter] = {
            ".pdf": PDFConverter(),
            ".docx": WordConverter(),
            ".xlsx": ExcelConverter(),
            ".xls": ExcelConverter(),
        }

    def convert_one(self, input_path: Path, options: ConversionOptions) -> ConversionResult:
        suffix = input_path.suffix.lower()
        if suffix not in self._handlers:
            return ConversionResult(input_path=input_path, success=False, error=f"Unsupported extension: {suffix}")

        plan = plan_output(input_path, options)

        # overwrite cleanup
        if options.overwrite:
            try:
                if plan.output_path.exists():
                    plan.output_path.unlink()
            except Exception:
                pass
            try:
                if plan.assets_dir.exists():
                    shutil.rmtree(plan.assets_dir, ignore_errors=True)
            except Exception:
                pass

        logger.info(f"Convert: {input_path} -> {plan.output_path.name} ({suffix[1:]})")
        return self._handlers[suffix].convert(input_path, plan, options)

    def convert_many(self, inputs: Sequence[str], options: ConversionOptions) -> List[ConversionResult]:
        files = _expand_inputs(inputs, options.recursive)
        if not files:
            raise ConversionError("input", "No input files found.")
        results: List[ConversionResult] = []
        for p in files:
            r = self.convert_one(p, options)
            results.append(r)
            if (not r.success) and options.fail_fast:
                break
        return results


# -----------------------------
# CLI
# -----------------------------


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="doc2md",
        description="PDF / Excel / Word(.docx) を Markdown(.md) / AsciiDoc(.adoc) に変換（OCR統合 + Pandoc対応）",
    )
    p.add_argument("inputs", nargs="+", help="入力ファイル/フォルダ/グロブ（例: docs/*.pdf）")

    g_out = p.add_argument_group("Output")
    g_out.add_argument("-o", "--output", default="output_data", help="出力先フォルダ (default: output_data)")
    g_out.add_argument("--output-format", default="md", choices=[e.value for e in OutputFormat], help="出力フォーマット (md|adoc)")
    g_out.add_argument("--overwrite", action="store_true", help="同名出力がある場合に上書き")
    g_out.add_argument("--recursive", action="store_true", help="フォルダ入力時に再帰探索")
    g_out.add_argument("--no-front-matter", action="store_true", help="フロントマター（YAML/AsciiDoc属性）を付けない")
    g_out.add_argument("--no-sha256", action="store_true", help="フロントマターに sha256 を含めない")
    g_out.add_argument("--include-source-path", action="store_true", help="フロントマターに source_path（絶対パス）を含める")
    g_out.add_argument("--no-images", action="store_true", help="画像抽出（assets出力）を無効化")

    g_pandoc = p.add_argument_group("Pandoc (optional)")
    g_pandoc.add_argument("--use-pandoc", action="store_true", help="Pandoc による高品質変換を有効化（利用不可なら起動時に警告→フォールバック）")
    g_pandoc.add_argument("--pandoc-bin", default="pandoc", help="pandoc 実行ファイル名/パス (default: pandoc)")

    g_pdf = p.add_argument_group("PDF")
    g_pdf.add_argument("--pdf-engine", default="auto", choices=[e.value for e in PdfEngine], help="PDF 変換エンジン (auto|pymupdf4llm|pymupdf_text)")
    g_pdf.add_argument("--pdf-max-pages", type=int, default=None, help="変換する最大ページ数（未指定=全ページ）")
    g_pdf.add_argument("--pdf-force-page-headings", action="store_true", help="ページ見出し（Page N）を強制生成（OCR統合やRAG用途向け）")
    g_pdf.add_argument("--pdf-page-heading-template", default="## Page {n}", help='ページ見出しテンプレート（例: "## Page {n}" / "## ページ {n}"）')

    g_ocr = p.add_argument_group("PDF OCR (optional)")
    g_ocr.add_argument("--pdf-ocr", action="store_true", help="OCR を有効化")
    g_ocr.add_argument("--pdf-ocr-strategy", default="fallback", choices=[e.value for e in OcrStrategy], help="OCR 統合戦略 (fallback|integrate|append|replace)")
    g_ocr.add_argument("--pdf-ocr-min-chars", type=int, default=200, help="fallback 判定のしきい値（既定: 200）")
    g_ocr.add_argument("--pdf-ocr-lang", default="jpn", help="OCR 言語（例: jpn / jpn+jpn_vert）")
    g_ocr.add_argument("--pdf-ocr-dpi", type=int, default=300, help="レンダリングDPI（既定: 300）")
    g_ocr.add_argument("--pdf-ocr-psm", type=int, default=6, help="Tesseract PSM（既定: 6）")
    g_ocr.add_argument("--pdf-ocr-oem", type=int, default=3, help="Tesseract OEM（既定: 3）")
    g_ocr.add_argument("--pdf-ocr-tesseract", default="", help="tesseract 実行ファイルのパス（Windows等で必要な場合）")
    g_ocr.add_argument("--pdf-ocr-tessdata", default="", help="tessdata ディレクトリ（任意）")
    g_ocr.add_argument("--pdf-ocr-debug-images", action="store_true", help="前処理後の画像を assets に保存")
    g_ocr.add_argument("--pdf-ocr-threshold", type=int, default=160, help="2値化しきい値（既定: 160）")
    g_ocr.add_argument("--pdf-ocr-contrast", type=float, default=1.8, help="コントラスト強調率（既定: 1.8）")

    g_excel = p.add_argument_group("Excel (fallback: pandas)")
    g_excel.add_argument("--excel-max-rows", type=int, default=500, help="各シートの最大行数（既定: 500）")
    g_excel.add_argument("--excel-max-cols", type=int, default=50, help="各シートの最大列数（既定: 50）")
    g_excel.add_argument("--excel-cell-max-chars", type=int, default=200, help="セル文字数上限（既定: 200）")

    g_word = p.add_argument_group("Word (fallback: mammoth)")
    g_word.add_argument("--word-heading-style", default="ATX", choices=[e.value for e in WordHeadingStyle], help="見出しスタイル（ATX|SETEXT）")
    # 空行の扱いは「既定: 除去（True）」で、明示的に ON/OFF を切り替えられるようにします。
    # store_true に default=True を付けるよりも、mutually exclusive + set_defaults の方が読みやすいです。
    g_word_me = g_word.add_mutually_exclusive_group()
    g_word_me.add_argument("--word-strip-empty-lines", dest="word_strip_empty_lines", action="store_true", help="空行を除去（既定: 有効）")
    g_word_me.add_argument("--word-keep-empty-lines", dest="word_strip_empty_lines", action="store_false", help="空行を保持（空行除去を無効化）")
    g_word.set_defaults(word_strip_empty_lines=True)

    g_run = p.add_argument_group("Run/Debug")
    g_run.add_argument("--fail-fast", action="store_true", help="最初の失敗で処理を中断")
    g_run.add_argument("--json-summary", default=None, help="処理結果サマリーを JSON で保存（パス指定）")
    g_run.add_argument("--debug", action="store_true", help="失敗時に例外を再送出（スタックトレース確認用）")
    g_run.add_argument("-v", "--verbose", action="store_true", help="詳細ログ（DEBUG）")
    g_run.add_argument("-q", "--quiet", action="store_true", help="警告以上のみ")
    g_run.add_argument("--version", action="version", version=f"doc2md {__version__}")

    return p


# -----------------------------
# JSON summary
# -----------------------------


def _write_json_summary(path: Path, results: List[ConversionResult], options: ConversionOptions) -> None:
    out: Dict[str, Any] = {
        "tool": options.tool_name,
        "tool_version": __version__,
        "converted_at": _now_iso(options.converted_at_utc),
        "output_dir": str(options.output_dir.resolve()),
        "success": sum(1 for r in results if r.success),
        "failed": sum(1 for r in results if not r.success),
        "results": [],
    }
    for r in results:
        out["results"].append(
            {
                "input": str(r.input_path),
                "success": r.success,
                "output": str(r.output_path) if r.output_path else None,
                "warnings": [w.to_dict() for w in r.warnings],
                "error": r.error,
                "stats": r.stats,
            }
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")


# -----------------------------
# main
# -----------------------------


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    _configure_logging(bool(getattr(args, "verbose", False)), bool(getattr(args, "quiet", False)))

    # enable layout (after logging configured)
    activate_pymupdf_layout()

    options = ConversionOptions.from_args(args)

    # --use-pandoc 指定時は「起動直後に一度だけ」pandoc の可用性チェック
    if options.use_pandoc:
        ver = _pandoc_version(options.pandoc_bin)
        if not ver:
            logger.warning(
                "--use-pandoc が指定されていますが、pandoc が見つからない/起動できないため、従来方式へフォールバックします。"
                f" (pandoc_bin={options.pandoc_bin})"
            )
            options = replace(options, use_pandoc=False)
        else:
            logger.info(f"pandoc available: {ver}")

    conv = DocumentConverter()

    try:
        results = conv.convert_many(args.inputs, options)
    except Exception as e:
        _log_exception("Conversion aborted", e)
        return 1

    ok = sum(1 for r in results if r.success)
    ng = len(results) - ok

    for r in results:
        if r.success:
            logger.info(f"OK  : {r.input_path.name} -> {r.output_path}")
            if r.warnings:
                wtxt = " / ".join(str(w) for w in r.warnings[:8])
                logger.warning(f"Warnings ({r.input_path.name}): {wtxt}" + (" ..." if len(r.warnings) > 8 else ""))
        else:
            logger.error(f"FAIL: {r.input_path.name} : {r.error}")
            if r.warnings:
                wtxt = " / ".join(str(w) for w in r.warnings[:8])
                logger.warning(f"Warnings ({r.input_path.name}): {wtxt}" + (" ..." if len(r.warnings) > 8 else ""))

    if options.json_summary_path:
        try:
            _write_json_summary(options.json_summary_path, results, options)
            logger.info(f"JSON summary saved: {options.json_summary_path}")
        except Exception as e:
            logger.warning(f"Failed to write JSON summary: {e}")

    logger.info(f"Done. success={ok}, failed={ng}, output_dir={options.output_dir.resolve()}")
    return 0 if ng == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

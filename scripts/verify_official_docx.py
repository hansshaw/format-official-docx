#!/usr/bin/env python3
"""Audit and render an official-format DOCX in one final verification pass."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from official_docx import audit_document


DEFAULT_DPI = 170
DEFAULT_TIMEOUT_SECONDS = 120


for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


def _prepend_workspace_runtime_bin() -> None:
    dependencies_dir = Path(sys.executable).resolve().parent.parent
    runtime_bin = dependencies_dir / "bin" / "override"
    if runtime_bin.is_dir():
        current = os.environ.get("PATH", "")
        entries = [entry for entry in current.split(os.pathsep) if entry]
        os.environ["PATH"] = os.pathsep.join(
            [str(runtime_bin), *[entry for entry in entries if Path(entry) != runtime_bin]]
        )


def _find_soffice(explicit: Path | None) -> Path:
    candidates: list[Path] = []
    if explicit is not None:
        candidates.append(explicit.expanduser())
    for name in ("soffice", "soffice.exe"):
        resolved = shutil.which(name)
        if resolved:
            candidates.append(Path(resolved))

    program_files = [
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
        "C:/Program Files",
        "C:/Program Files (x86)",
    ]
    for root in program_files:
        if root:
            candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.exe")

    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).casefold()
        if key in seen:
            continue
        seen.add(key)
        if candidate.is_file():
            return candidate.resolve()
    raise FileNotFoundError(
        "未找到 LibreOffice soffice。请安装 LibreOffice，或使用 --soffice 指定 soffice.exe。"
    )


def _convert_to_pdf(
    input_path: Path,
    soffice: Path,
    work_dir: Path,
    timeout_seconds: int,
) -> tuple[Path, subprocess.CompletedProcess[str]]:
    profile_dir = work_dir / "profile"
    convert_dir = work_dir / "convert"
    profile_dir.mkdir()
    convert_dir.mkdir()

    env = os.environ.copy()
    env["HOME"] = str(profile_dir)
    profile_argument = f"-env:UserInstallation={profile_dir.resolve().as_uri()}"
    command = [
        str(soffice),
        profile_argument,
        "--headless",
        "--invisible",
        "--nologo",
        "--nodefault",
        "--norestore",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
        str(convert_dir),
        str(input_path),
    ]
    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        timeout=timeout_seconds,
    )

    expected = convert_dir / f"{input_path.stem}.pdf"
    if expected.is_file() and expected.stat().st_size > 0:
        return expected, result
    generated = [path for path in convert_dir.glob("*.pdf") if path.stat().st_size > 0]
    if len(generated) == 1:
        return generated[0], result
    details = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
    raise RuntimeError(
        f"LibreOffice 未生成有效 PDF（退出码 {result.returncode}）。"
        + (f"\n{details}" if details else "")
    )


def _find_pdftoppm(explicit_dir: Path | None) -> Path:
    candidates: list[Path] = []
    if explicit_dir is not None:
        candidates.extend(
            [explicit_dir.expanduser() / "pdftoppm.exe", explicit_dir.expanduser() / "pdftoppm"]
        )

    dependencies_dir = Path(sys.executable).resolve().parent.parent
    candidates.extend(
        [
            dependencies_dir / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe",
            dependencies_dir / "native" / "poppler" / "bin" / "pdftoppm.exe",
        ]
    )
    resolved = shutil.which("pdftoppm")
    if resolved:
        candidates.append(Path(resolved))
    resolved_exe = shutil.which("pdftoppm.exe")
    if resolved_exe:
        candidates.append(Path(resolved_exe))

    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).casefold()
        if key in seen:
            continue
        seen.add(key)
        if candidate.is_file() and candidate.suffix.casefold() != ".cmd":
            return candidate.resolve()
    raise FileNotFoundError(
        "未找到 Poppler 的 pdftoppm。请使用 Codex 工作区 Python，或通过 --poppler-dir 指定其目录。"
    )


def _render_pages(
    pdf_path: Path,
    output_dir: Path,
    dpi: int,
    timeout_seconds: int,
    poppler_dir: Path | None,
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    for stale in output_dir.glob("page-*.png"):
        stale.unlink()
    for stale in output_dir.glob("official-page-*.png"):
        stale.unlink()

    pdftoppm = _find_pdftoppm(poppler_dir)
    prefix = output_dir / "official-page"
    result = subprocess.run(
        [str(pdftoppm), "-png", "-r", str(dpi), str(pdf_path), str(prefix)],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout_seconds,
    )
    raw_paths = sorted(
        output_dir.glob("official-page-*.png"),
        key=lambda path: int(path.stem.rsplit("-", 1)[-1]),
    )
    if result.returncode != 0 or not raw_paths:
        details = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
        raise RuntimeError(
            f"Poppler 未生成 PNG（退出码 {result.returncode}）。"
            + (f"\n{details}" if details else "")
        )

    pages: list[Path] = []
    for index, raw_path in enumerate(raw_paths, start=1):
        destination = output_dir / f"page-{index}.png"
        os.replace(raw_path, destination)
        pages.append(destination.resolve())
    if not pages:
        raise RuntimeError("PDF 已生成，但未渲染出任何 PNG 页面。")
    return pages


def verify_document(args: argparse.Namespace) -> dict[str, object]:
    input_path = args.input.expanduser().resolve()
    if not input_path.is_file() or input_path.suffix.lower() != ".docx":
        raise ValueError(f"找不到有效的 DOCX 文件：{input_path}")
    if args.dpi < 72 or args.dpi > 600:
        raise ValueError("--dpi 必须在 72 到 600 之间。")
    if args.timeout < 1:
        raise ValueError("--timeout 必须大于 0。")

    report = audit_document(
        input_path,
        title_index=args.title_index,
        check_fonts=not args.skip_font_check,
    )
    result: dict[str, object] = {
        "path": str(input_path),
        "ok": False,
        "audit": report.as_dict(),
        "render": None,
    }
    if not report.ok:
        return result

    _prepend_workspace_runtime_bin()
    soffice = _find_soffice(args.soffice)
    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir
        else input_path.with_name(f"{input_path.stem}-qa")
    )

    with tempfile.TemporaryDirectory(prefix="official_docx_verify_") as temp_dir:
        pdf_path, conversion = _convert_to_pdf(
            input_path,
            soffice,
            Path(temp_dir),
            args.timeout,
        )
        pages = _render_pages(
            pdf_path,
            output_dir,
            args.dpi,
            args.timeout,
            args.poppler_dir,
        )
        emitted_pdf: Path | None = None
        if args.emit_pdf:
            emitted_pdf = output_dir / f"{input_path.stem}.pdf"
            shutil.copy2(pdf_path, emitted_pdf)
            emitted_pdf = emitted_pdf.resolve()

    result["ok"] = True
    result["render"] = {
        "soffice": str(soffice),
        "conversion_exit_code": conversion.returncode,
        "dpi": args.dpi,
        "page_count": len(pages),
        "pages": [str(path) for path in pages],
        "pdf": str(emitted_pdf) if emitted_pdf else None,
    }
    return result


def _print_result(result: dict[str, object], as_json: bool) -> None:
    if as_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    audit = result["audit"]
    if not isinstance(audit, dict):
        raise TypeError("内部审计结果无效。")
    print(("PASS" if result["ok"] else "FAIL") + f": {result['path']}")
    for error in audit.get("errors", []):
        print(f"  ERROR: {error}")
    for warning in audit.get("warnings", []):
        print(f"  WARNING: {warning}")
    render = result.get("render")
    if isinstance(render, dict):
        print(f"  RENDER: {render['page_count']} page(s) at {render['dpi']} DPI")
        print(f"  OUTPUT: {Path(render['pages'][0]).parent}")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="一次完成中文公文 DOCX 结构审计和 LibreOffice 最终渲染。"
    )
    parser.add_argument("input", type=Path)
    parser.add_argument("--output-dir", type=Path, default=None)
    parser.add_argument("--title-index", type=int, default=None)
    parser.add_argument("--soffice", type=Path, default=None)
    parser.add_argument("--poppler-dir", type=Path, default=None)
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI)
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--emit-pdf", action="store_true")
    parser.add_argument("--skip-font-check", action="store_true")
    parser.add_argument("--json", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        result = verify_document(args)
        _print_result(result, args.json)
        return 0 if result["ok"] else 1
    except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as exc:
        if args.json:
            print(json.dumps({"ok": False, "errors": [str(exc)]}, ensure_ascii=False, indent=2))
        else:
            print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

"""Command-line entry point."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .app import run
from .logging_config import configure_logging
from .paths import default_base_dir, ensure_runtime_dirs


def build_parser() -> argparse.ArgumentParser:
    """Create the CLI argument parser."""

    parser = argparse.ArgumentParser(
        prog="csv-to-formatted-xlsx",
        description="批次將 CSV 安全轉換為格式化 XLSX，並保護條碼等識別碼。",
    )
    parser.add_argument("--input", type=Path, help="CSV 輸入資料夾（預設：程式旁的 in）")
    parser.add_argument("--output", type=Path, help="XLSX 輸出資料夾（預設：程式旁的 out）")
    parser.add_argument(
        "--overwrite",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="安全覆蓋同名 XLSX（預設：啟用）",
    )
    parser.add_argument("--no-pause", action="store_true", help="結束後不等待 Enter")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the command-line application and return its process exit code."""

    args = build_parser().parse_args(argv)
    base_dir = default_base_dir()
    default_input, default_output, log_dir = ensure_runtime_dirs(base_dir)
    input_dir = (args.input or default_input).resolve()
    output_dir = (args.output or default_output).resolve()
    logger = configure_logging(log_dir)
    logger.info("啟動 version=%s input=%s output=%s", __version__, input_dir, output_dir)

    try:
        exit_code, _ = run(
            input_dir,
            output_dir,
            overwrite=args.overwrite,
            logger=logger,
        )
    except Exception as exc:
        logger.exception("不可恢復的程式錯誤：%s", exc)
        print(f"[程式錯誤] {exc}")
        exit_code = 2

    if getattr(sys, "frozen", False) and not args.no_pause:
        try:
            input("按 Enter 結束...")
        except EOFError:
            pass
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .analyzer import analyze_workbook
from .converter import convert_many


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="离线 BOM Excel 格式转换工具")
    sub = parser.add_subparsers(dest="command", required=True)
    analyze = sub.add_parser("analyze", help="只分析字段映射，不写入 Excel")
    analyze.add_argument("sources", nargs="+", type=Path)
    analyze.add_argument("--mode", choices=("quick", "audit"), default="audit")
    analyze.add_argument("--json", action="store_true", help="输出完整 JSON")
    convert = sub.add_parser("convert", help="转换一份或多份已知格式 BOM")
    convert.add_argument("sources", nargs="+", type=Path)
    convert.add_argument("--template", type=Path, required=True)
    convert.add_argument("--output-dir", type=Path, required=True)
    convert.add_argument("--mode", choices=("quick", "audit"), default="quick")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "analyze":
            for source in args.sources:
                analysis = analyze_workbook(source, mode=args.mode)
                if args.json:
                    print(json.dumps(analysis.to_dict(), ensure_ascii=False, indent=2))
                else:
                    print(f"{source.name}: {analysis.profile_name}, 表头 {analysis.header_rows}, 数据 {analysis.data_row_count} 行, 图片 {analysis.image_count} 张")
                    for mapping in analysis.mappings:
                        target = mapping.target_header or f"[{mapping.status}]"
                        print(f"  {mapping.source_col:>2} {mapping.source_header} -> {target} ({mapping.confidence:.0%})")
            return 0
        results = convert_many(args.sources, args.template, args.output_dir, mode=args.mode)
        failed = False
        for result in results:
            error_count = sum(issue.severity == "error" for issue in result.issues)
            warning_count = sum(issue.severity == "warning" for issue in result.issues)
            print(f"{result.source_path.name} -> {result.output_path} | 行 {result.output_rows}/{result.source_rows} | 图片 {result.output_images}/{result.source_images} | 警告 {warning_count} | 错误 {error_count}")
            print(f"  报告: {result.report_path}")
            failed = failed or error_count > 0
        return 2 if failed else 0
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

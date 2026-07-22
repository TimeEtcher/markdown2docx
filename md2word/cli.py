"""命令行入口：md2word input.md -c config.yaml -o out.docx；无参数时启动 GUI。"""

from __future__ import annotations

import argparse
import os
import sys


def convert(md_path: str, config_path: str | None, output_path: str | None) -> str:
    from .config import load_config
    from .converter import Converter

    if not output_path:
        output_path = os.path.splitext(os.path.abspath(md_path))[0] + ".docx"
    config = load_config(config_path)
    Converter(config).convert_file(md_path, output_path)
    return output_path


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="md2word", description="Markdown 转 Word（支持 YAML 样式配置）")
    parser.add_argument("input", nargs="?", help="输入 Markdown 文件路径（省略则启动 GUI）")
    parser.add_argument("-c", "--config", help="YAML 配置文件路径（可选）")
    parser.add_argument("-o", "--output", help="输出 docx 路径（默认与输入同名）")
    args = parser.parse_args(argv)

    if not args.input:
        from .gui import run_gui
        run_gui()
        return 0

    try:
        out = convert(args.input, args.config, args.output)
    except Exception as exc:
        print(f"转换失败: {exc}", file=sys.stderr)
        return 1
    print(f"已生成: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

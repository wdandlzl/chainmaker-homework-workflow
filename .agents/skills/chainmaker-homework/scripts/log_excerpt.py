"""Read selected lines from one experiment log, with original line numbers."""

import argparse
import re
from pathlib import Path

DEFAULT_MATCH = (
    r"^\s*(?:\$\s+|\+\s+|command\s*[:=])"
    r"|\b(?:version|error|failed|failure|timeout|refused|rejected|passed|verified|success)\b"
    r"|^\s*(?:go version|chainmaker\.org/\S+\s+v\d)"
    r"|\b(?:result_code|exit_code|return_code)\s*="
    r"|成功|失败|错误|超时|拒绝"
)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--log", required=True, type=Path)
    selectors = parser.add_mutually_exclusive_group()
    selectors.add_argument("--match", action="append", metavar="REGEX",
                           help="只匹配给定表达式；可重复指定")
    selectors.add_argument("--lines", metavar="START-END", help="逐行回读指定范围，不截断行")
    parser.add_argument("--context", type=int, default=1)
    parser.add_argument("--limit", type=int, default=24, help="匹配模式的最大输出行数")
    args = parser.parse_args()
    if args.context < 0 or args.limit < 1:
        parser.error("--context 须非负，--limit 须大于零")

    try:
        with args.log.open(encoding="utf-8-sig", errors="replace") as source:
            lines = [line.rstrip("\r\n") for line in source]
    except OSError as exc:
        parser.error(str(exc))

    print(f"日志：{args.log.resolve()}（{len(lines)} 行）")
    if args.lines:
        match = re.fullmatch(r"(\d+)(?:-(\d+))?", args.lines)
        if not match:
            parser.error("--lines 应为单个行号或 START-END")
        start, end = int(match[1]), int(match[2] or match[1])
        if not 1 <= start <= end <= len(lines):
            parser.error(f"行号须在 1–{len(lines)} 内")
        for number in range(start, end + 1):
            print(f"{number}: {lines[number - 1]}")
        return

    try:
        patterns = [re.compile(pattern, re.IGNORECASE) for pattern in args.match or [DEFAULT_MATCH]]
    except re.error as exc:
        parser.error(str(exc))
    matches = [
        i for i, line in enumerate(lines) if any(pattern.search(line) for pattern in patterns)
    ]
    # Prioritize matched lines when adjacent context would consume the output budget.
    selected = set(matches[:args.limit])
    remaining = args.limit - len(selected)
    for i in matches:
        if not remaining:
            break
        neighbors = range(max(0, i - args.context), min(len(lines), i + args.context + 1))
        for neighbor in neighbors:
            if neighbor not in selected:
                selected.add(neighbor)
                remaining -= 1
                if not remaining:
                    break

    if not matches:
        print("未匹配。请使用实验专用 --match 或 --lines 回读。")
        return
    print(f"匹配 {len(matches)} 行，显示 {len(selected)} 行（含上下文）。")
    for i in sorted(selected):
        line = lines[i]
        suffix = " …（使用 --lines 回读）" if len(line) > 320 else ""
        print(f"{i + 1}: {line[:320]}{suffix}")
    if len(matches) > args.limit:
        print(f"另有 {len(matches) - args.limit} 条匹配未显示；可缩小 --match 或回读原始行。")


if __name__ == "__main__":
    main()

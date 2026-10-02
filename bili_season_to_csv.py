#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bili_season_to_csv.py —— B站课程列表导出工具

【功能】
  给一个 B站视频链接或若干 BV 号，自动抓标题、展开分 P，
  输出「序号 + 课程名称」两列 CSV。

【双击使用】
  直接双击本文件，按屏幕提示操作即可。

【命令行使用】
  python bili_season_to_csv.py bvlist.txt
  python bili_season_to_csv.py BV1xx411c7mD BV1yy411c7mE
  python bili_season_to_csv.py "https://www.bilibili.com/video/BV1dr4y1n7vA"
"""

import csv
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

API = "https://api.bilibili.com/x/web-interface/view"
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
    "Referer": "https://www.bilibili.com/",
}
FIELDS = ["序号", "课程名称"]


# ---------------------------------------------------------------------------
# 网络请求
# ---------------------------------------------------------------------------

def fetch(bvid: str) -> dict | None:
    """拉单个视频信息。返回 None 表示拿不到（失效 / 充电专属 / 被风控）。"""
    url = f"{API}?bvid={urllib.parse.quote(bvid)}"
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            j = json.load(resp)
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError) as e:
        print(f"  [跳过] {bvid} 请求失败: {e}")
        return None

    if j.get("code") != 0:
        print(f"  [跳过] {bvid} 接口返回 {j.get('code')}: {j.get('message')}")
        return None

    d = j["data"]
    return {
        "title": d.get("title", "").strip(),
        "pages": [
            {"page": p.get("page", 1), "part": (p.get("part") or "").strip()}
            for p in d.get("pages", [])
        ],
    }


# ---------------------------------------------------------------------------
# 名称处理
# ---------------------------------------------------------------------------

def clean_suffix(name: str) -> str:
    """去掉录屏/转码留下的 '_高清'、'_高清 720P'、'_超清' 等尾巴。"""
    return re.sub(r"\s*[_-]?(高清|超清|720P|1080P|4K).*?$", "", name, flags=re.I).strip()


def page_label(page: int, part: str, video_title: str = "") -> str:
    """
    分 P 显示名。
    - 只负责去掉 part 里与视频总标题重复的前缀，以及空 part 兜底。
    - 不额外添加 P{n}，避免把「C语言简史」变成「P1 C语言简史」。
    """
    part = (part or "").strip()

    # 去掉 part 里与视频标题重复的前缀
    vt = video_title.strip()
    if vt and part.lower().startswith(vt.lower()):
        part = part[len(vt):].strip()
        for sep in ("|", "-", "—", ":", "："):
            if part.startswith(sep):
                part = part[1:].strip()
                break

    # part 为空：用视频标题兜底（单个视频只有一集时）
    if not part:
        return vt or f"P{page}"

    return part


# ---------------------------------------------------------------------------
# 核心处理
# ---------------------------------------------------------------------------

def parse_entry(raw: str) -> tuple[str, int | None]:
    """
    把用户输入解析成 (BV号, 目标分P或None)。
    支持：BVxxx、BVxxx?p=2、https://.../video/BVxxx?p=2
    """
    raw = raw.strip()
    if not raw or raw.startswith("#"):
        return "", None

    # 从链接里抓 BV 号
    m = re.search(r"BV[0-9A-Za-z]{10}", raw)
    if not m:
        return "", None
    bvid = m.group(0)

    # 抓 ?p=N
    want_page = None
    qs_match = re.search(r"[?&]p=(\d+)", raw)
    if qs_match:
        try:
            want_page = int(qs_match.group(1))
        except ValueError:
            want_page = None
    return bvid, want_page


def build_rows(entries: list[str]) -> list[str]:
    """把输入条目转成课程名称列表。"""
    names = []
    parsed = [parse_entry(e) for e in entries]
    parsed = [(bv, p) for bv, p in parsed if bv]
    total = len(parsed)

    for i, (bvid, want_page) in enumerate(parsed, 1):
        print(f"[{i}/{total}] 抓取 {bvid}" + (f" 第 {want_page} P" if want_page else ""))
        info = fetch(bvid)
        if info is None:
            continue

        title = info["title"]
        pages = info["pages"]

        if want_page is not None:
            match = next((p for p in pages if p["page"] == want_page), None)
            if match:
                names.append(page_label(match["page"], match["part"], title))
            continue

        if len(pages) > 1:
            for p in pages:
                names.append(page_label(p["page"], p["part"], title))
        else:
            names.append(title)

        time.sleep(0.4)  # 别打太快，B 站会风控

    # 去掉所有高清/超清后缀
    names = [clean_suffix(n) for n in names]
    return names


def save_csv(rows: list[str], path: str) -> None:
    """utf-8-sig = UTF-8 带 BOM，Excel / Notion 打开中文都不乱码。"""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(FIELDS)
        for i, name in enumerate(rows, 1):
            w.writerow([i, name])


# ---------------------------------------------------------------------------
# 交互界面
# ---------------------------------------------------------------------------

def print_header():
    os.system("cls" if os.name == "nt" else "clear")
    print("=" * 50)
    print("  B站课程列表导出工具")
    print("=" * 50)
    print()


def read_entries_from_file(path: str) -> list[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到文件：{path}")
    with open(path, encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]


def interactive_mode():
    print_header()

    print("请选择输入方式：")
    print("  [1] 粘贴 B站视频链接（自动提取 BV 号，单视频分 P 自动展开）")
    print("  [2] 输入一个或多个 BV 号（空格分隔）")
    print("  [3] 从 txt 文件读取 BV 号")
    print("  [0] 退出")
    print()
    choice = input("请输入数字 [0-3]: ").strip()

    entries: list[str] = []
    if choice == "1":
        print()
        print("提示：粘贴后会自动提取 BV 号。如果视频有多个分 P，会全部展开。")
        link = input("请粘贴链接: ").strip()
        entries = [link]
    elif choice == "2":
        print()
        line = input("请输入 BV 号（多个用空格分隔）: ").strip()
        entries = line.split()
    elif choice == "3":
        print()
        print("提示：txt 文件每行一个 BV 号或链接，# 开头为注释。")
        path = input("请输入文件路径（可拖放到此处）: ").strip().strip('"')
        try:
            entries = read_entries_from_file(path)
        except FileNotFoundError as e:
            print(f"\n错误：{e}")
            pause_and_exit(1)
            return
    elif choice == "0":
        pause_and_exit(0)
        return
    else:
        print("\n错误：无效选项，请输入 0-3。")
        pause_and_exit(1)
        return

    if not entries:
        print("\n错误：没有拿到任何 BV 号。")
        pause_and_exit(1)
        return

    print()
    names = build_rows(entries)
    if not names:
        print("\n错误：没有抓到任何课程名称，请检查 BV 号或网络。")
        pause_and_exit(1)
        return

    print()
    default_name = "bili_season.csv"
    out = input(f"输出文件名（直接回车使用默认 {default_name}）: ").strip().strip('"')
    if not out:
        out = default_name
    if not out.lower().endswith(".csv"):
        out += ".csv"

    save_csv(names, out)
    print()
    print(f"完成：{len(names)} 行 → {os.path.abspath(out)}")
    print("表头：序号, 课程名称")
    pause_and_exit(0)


def pause_and_exit(code: int):
    print()
    try:
        input("按回车键退出...")
    except EOFError:
        pass
    sys.exit(code)


# ---------------------------------------------------------------------------
# 命令行兼容模式
# ---------------------------------------------------------------------------

def cli_mode(args: list[str]):
    out = "bili_season.csv"
    if "-o" in args:
        pos = args.index("-o")
        out = args[pos + 1]
        args = args[:pos] + args[pos + 2:]

    entries = []
    for a in args:
        if a.lower().endswith(".txt"):
            entries.extend(read_entries_from_file(a))
        else:
            entries.append(a)

    names = build_rows(entries)
    if not names:
        print("没有抓到任何数据，检查 BV 号或网络。")
        sys.exit(1)

    save_csv(names, out)
    print(f"\n完成：{len(names)} 行 → {out}")
    print("表头：序号, 课程名称")


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def main():
    args = sys.argv[1:]
    if not args:
        interactive_mode()
    else:
        cli_mode(args)


if __name__ == "__main__":
    main()

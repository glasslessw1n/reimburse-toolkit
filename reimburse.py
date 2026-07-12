#!/usr/bin/env python3
"""
报销单据整理工具 v2.0 - 通用版
功能：
  scan     - 扫描目录，列出所有文件及其分类
  summary  - 按行程汇总文件数量
  report   - 生成费用汇总报告(Markdown)
  check    - 核对滴滴发票与行程报销单配对
  text     - 批量提取PDF文本内容

用法:
  python3 reimburse.py scan --base-dir /path/to/dir
  python3 reimburse.py summary --base-dir /path/to/dir
  python3 reimburse.py report --base-dir /path/to/dir
  python3 reimburse.py check --base-dir /path/to/dir
"""

import os, sys, re, json
from collections import defaultdict
from pathlib import Path


def get_base_dir():
    """获取工作目录：优先 --base-dir 参数，否则当前目录"""
    for i, arg in enumerate(sys.argv):
        if arg == '--base-dir' and i + 1 < len(sys.argv):
            return Path(sys.argv[i + 1])
    for arg in sys.argv:
        if arg.startswith('--base-dir='):
            return Path(arg.split('=', 1)[1])
    return Path.cwd()


# ── 文件类型识别 ──────────────────────────────────
AMOUNT_PATTERNS = [
    re.compile(r"(\d+\.?\d*)\.pdf$"),
    re.compile(r"(\d+\.?\d*)元?\.pdf$"),
]


def extract_amount(filename):
    for pat in AMOUNT_PATTERNS:
        m = pat.search(filename)
        if m:
            return float(m.group(1))
    return None


def classify_file(filename):
    """根据文件名分类"""
    name = os.path.basename(filename)
    if "滴滴出行行程报销单" in name:
        letter = re.search(r"单([A-Z]?)", name)
        return ("滴滴行程单", letter.group(1) if letter and letter.group(1) else "")
    elif "滴滴电子发票" in name:
        letter = re.search(r"票([A-Z]?)", name)
        return ("滴滴发票", letter.group(1) if letter and letter.group(1) else "")
    elif "住宿水单" in name or ("住宿" in name and name.lower().endswith('.jpg')):
        return ("住宿水单", "")
    elif "住宿发票" in name:
        return ("住宿发票", "")
    elif "登机牌" in name:
        return ("交通票", "登机牌")
    elif "餐饮发票" in name:
        amt = extract_amount(name)
        return ("餐饮发票", f"¥{amt}" if amt else "")
    elif "通信发票" in name:
        amt = extract_amount(name)
        return ("通信发票", f"¥{amt}" if amt else "")
    elif re.search(r"\d{4}\s+\S+-\S+\s+\d+", name):
        amt = extract_amount(name)
        route = re.search(r"(\S+-\S+)", name)
        return ("交通票", f"{route.group(1) if route else ''} ¥{amt}" if amt else route.group(1) if route else "")
    else:
        return ("其他", "")


SKIP_DIRS = {'.claude', '__pycache__', '_extracted_texts', '_报销整合输出', '.git'}


def scan_directory(root_dir=None):
    """扫描目录返回文件清单"""
    if root_dir is None:
        root_dir = get_base_dir()
    result = []
    for root, dirs, files in os.walk(root_dir):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith('.')]
        for f in sorted(files):
            if f.startswith('.') or f.endswith('.py'):
                continue
            fullpath = os.path.join(root, f)
            relpath = os.path.relpath(fullpath, root_dir)
            parts = relpath.split(os.sep)
            trip = parts[0] if len(parts) > 1 else "(根目录)"
            cat, sub = classify_file(f)
            amt = extract_amount(f)
            result.append({
                "trip": trip,
                "filename": f,
                "category": cat,
                "subcategory": sub,
                "amount_from_name": amt,
                "path": relpath,
                "size": os.path.getsize(fullpath)
            })
    return result


def scan_summary(root_dir=None):
    """按行程汇总"""
    if root_dir is None:
        root_dir = get_base_dir()
    files = scan_directory(root_dir)
    trips = defaultdict(lambda: defaultdict(list))
    for f in files:
        trips[f["trip"]][f["category"]].append(f)

    print(f"工作目录: {root_dir}\n")
    print("=" * 75)
    print(f"{'行程/目录':<30} {'文件数':>6} {'分类明细'}")
    print("=" * 75)
    total = 0
    for trip, cats in sorted(trips.items()):
        count = sum(len(v) for v in cats.values())
        total += count
        detail = " | ".join(f"{cat}:{len(items)}" for cat, items in sorted(cats.items()))
        print(f"{trip:<30} {count:>6}  {detail}")
    print("=" * 75)
    print(f"{'合计':<30} {total:>6}")
    return trips


def check_didi_pairing(root_dir=None):
    """核对滴滴行程单与电子发票的配对关系"""
    if root_dir is None:
        root_dir = get_base_dir()
    files = scan_directory(root_dir)
    trips = defaultdict(lambda: {"行程单": [], "发票": []})

    for f in files:
        if f["category"] == "滴滴行程单":
            trips[f["trip"]]["行程单"].append(f)
        elif f["category"] == "滴滴发票":
            trips[f["trip"]]["发票"].append(f)

    print(f"工作目录: {root_dir}\n")
    print("滴滴行程单 ↔ 电子发票 配对检查")
    print("=" * 55)
    all_ok = True
    for trip, data in sorted(trips.items()):
        if not data["行程单"] and not data["发票"]:
            continue
        forms = sorted(data["行程单"], key=lambda x: x["filename"])
        invoices = sorted(data["发票"], key=lambda x: x["filename"])

        form_letters = set()
        for f in forms:
            m = re.search(r"单([A-Z]?)", f["filename"])
            if m:
                form_letters.add(m.group(1))
        inv_letters = set()
        for f in invoices:
            m = re.search(r"票([A-Z]?)", f["filename"])
            if m:
                inv_letters.add(m.group(1))

        only_forms = form_letters - inv_letters
        only_inv = inv_letters - form_letters
        status = "✓ 全部配对" if not only_forms and not only_inv else f"✗ 缺发票:{only_forms} 缺行程单:{only_inv}"
        if only_forms or only_inv:
            all_ok = False

        print(f"  {trip}")
        print(f"    行程单: {[f['filename'] for f in forms]}")
        print(f"    发票:   {[f['filename'] for f in invoices]}")
        print(f"    状态:   {status}")
        print()

    if all_ok:
        print("  全部滴滴行程单与发票配对正常 ✓")

    return all_ok


def generate_report(root_dir=None):
    """生成完整报销汇总报告"""
    if root_dir is None:
        root_dir = get_base_dir()
    files = scan_directory(root_dir)
    trips = defaultdict(lambda: defaultdict(list))

    for f in files:
        trips[f["trip"]][f["category"]].append(f)

    lines = []
    lines.append("# 报销汇总报告")
    lines.append(f"生成日期: {__import__('datetime').datetime.now().strftime('%Y-%m-%d')}")
    lines.append("")
    lines.append("## 差旅行程一览")
    lines.append("")
    lines.append("| 行程 | 交通票 | 住宿水单 | 住宿发票 | 滴滴行程单 | 滴滴发票 | 餐饮发票 | 通信发票 |")
    lines.append("|------|--------|----------|----------|-----------|---------|---------|---------|")

    for trip, cats in sorted(trips.items()):
        counts = {cat: len(items) for cat, items in cats.items()}
        line = f"| {trip} | {counts.get('交通票',0)} | {counts.get('住宿水单',0)} | {counts.get('住宿发票',0)} | {counts.get('滴滴行程单',0)} | {counts.get('滴滴发票',0)} | {counts.get('餐饮发票',0)} | {counts.get('通信发票',0)} |"
        lines.append(line)

    lines.append("")
    lines.append("## 各类费用明细")
    lines.append("")

    for trip, cats in sorted(trips.items()):
        lines.append(f"### {trip}")
        lines.append("")
        for cat in ["交通票", "住宿水单", "住宿发票", "滴滴行程单", "滴滴发票", "餐饮发票", "通信发票"]:
            items = cats.get(cat, [])
            if items:
                lines.append(f"**{cat}** ({len(items)}个):")
                for item in items:
                    note = f"  (¥{item['amount_from_name']})" if item['amount_from_name'] else ""
                    lines.append(f"  - {item['filename']}{note}")
                lines.append("")
        lines.append("---")
        lines.append("")

    report_path = root_dir / "报销汇总报告.md"
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"报告已生成: {report_path}")
    return report_path


def extract_text_from_pdfs(root_dir=None):
    """批量提取PDF文本并保存"""
    import PyPDF2
    if root_dir is None:
        root_dir = get_base_dir()

    output_dir = root_dir / "_extracted_texts"
    output_dir.mkdir(exist_ok=True)

    count = 0
    for root, dirs, files in os.walk(root_dir):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith('.')]
        if '_extracted_texts' in root:
            continue
        for f in sorted(files):
            if not f.endswith('.pdf'):
                continue
            fullpath = os.path.join(root, f)
            try:
                reader = PyPDF2.PdfReader(fullpath)
                text = ""
                for page in reader.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n---PAGE---\n"

                if text.strip():
                    outname = os.path.splitext(f)[0] + '.txt'
                    outpath = output_dir / outname
                    with open(outpath, 'w', encoding='utf-8') as out:
                        out.write(text)
                    count += 1
            except Exception as e:
                print(f"  提取失败: {f} - {e}")

    print(f"已提取 {count} 个PDF的文本到 {output_dir}")
    return output_dir


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1]
    root = get_base_dir()

    if cmd == "scan":
        files = scan_directory(root)
        print(f"工作目录: {root}")
        for f in files:
            print(f"[{f['category']:<8s}] {f['trip']:<28s} {f['filename']}")

    elif cmd == "report":
        generate_report(root)

    elif cmd == "check":
        check_didi_pairing(root)

    elif cmd == "text":
        extract_text_from_pdfs(root)

    elif cmd == "summary":
        scan_summary(root)

    else:
        print(f"未知命令: {cmd}")
        print("可用: scan, summary, report, check, text")


if __name__ == "__main__":
    main()

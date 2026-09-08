#!/usr/bin/env python3
"""
报销文件分类排序引擎 - 通用版
用法:
  python3 organize.py                  # 在当前目录运行
  python3 organize.py --base-dir /path/to/reimbursements
  python3 organize.py --dry-run        # 仅打印不生成文件

输出: <base_dir>/报销单据整合排列.md
"""

import os, re, sys
from collections import defaultdict
from pathlib import Path


def get_base_dir():
    """获取工作目录：优先 --base-dir 参数，否则当前脚本所在目录"""
    for i, arg in enumerate(sys.argv):
        if arg == '--base-dir' and i + 1 < len(sys.argv):
            return Path(sys.argv[i + 1])
    for arg in sys.argv:
        if arg.startswith('--base-dir='):
            return Path(arg.split('=', 1)[1])
    return Path.cwd()

DRY_RUN = "--dry-run" in sys.argv


# ── 目录识别 ───────────────────────────────────
SKIP_DIRS = {'.claude', '__pycache__', '_extracted_texts', '_报销整合输出', '.git', '.DS_Store'}

def discover_trips(base_dir):
    """
    自动发现行程目录。
    规则：base_dir 下的一级子目录，排除系统目录。
    如果目录中包含至少一个报销相关文件，则认为是行程目录。
    """
    trips = []
    local_dirs = []

    for entry in sorted(base_dir.iterdir()):
        if not entry.is_dir():
            continue
        if entry.name in SKIP_DIRS or entry.name.startswith('.'):
            continue

        # 检查是否包含报销相关文件
        has_files = any(
            f.endswith(('.pdf', '.jpg')) and not f.startswith('.')
            for f in os.listdir(entry)
        )
        if not has_files:
            continue

        # 识别本地目录
        if '本地' in entry.name:
            local_dirs.append(entry.name)
        else:
            trips.append(entry.name)

    return trips, local_dirs


# ── 辅助函数 ───────────────────────────────────
def parse_date_mmdd(date_str):
    m = re.match(r"(\d{2})(\d{2})", date_str)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return (99, 99)


def parse_date_range(name):
    """从文件名提取日期范围 MMDD-MMDD"""
    m = re.search(r"(\d{4})[^\d](\d{4})", name)
    if m:
        return parse_date_mmdd(m.group(1)), parse_date_mmdd(m.group(2))
    m = re.search(r"(\d{4})", name)
    if m:
        d = parse_date_mmdd(m.group(1))
        return (d, d)
    return ((99, 99), (99, 99))


def sort_key_hotel_by_checkin(f):
    (cin, _) = parse_date_range(f["filename"])
    return cin


def _toll_nn(name):
    """从通行费文件名提取序号 NN (01/02/03)"""
    m = re.search(r"通行费(\d{2})", name)
    return int(m.group(1)) if m else 0


def _self_drive_sort_key(name):
    """自驾车加油费/通行费排序键：按日期，同日期时加油费在前、通行费按NN"""
    is_gas = "加油费" in name
    return (parse_date_range(name)[0], 0 if is_gas else 1, _toll_nn(name))


# ── 核心：文件分类排序 ──────────────────────────
def classify_and_sort(trip_dir, base_dir=None):
    """
    读取一个行程目录，返回按规则排列的数据结构。

    参数:
      trip_dir: 行程目录名（相对于 base_dir）
      base_dir: 基础目录（Path 对象）。不传则从 get_base_dir() 获取。
    """
    if base_dir is None:
        base_dir = get_base_dir()
    trip_path = base_dir / trip_dir
    if not trip_path.exists():
        return None

    files = []
    for f in os.listdir(trip_path):
        if f.startswith('.') or f.endswith(('.py', '.md')):
            continue
        full = trip_path / f
        if full.is_file():
            files.append({
                "filename": f,
                "fullpath": str(full),
                "size": full.stat().st_size
            })

    transport = []
    hotel_singles = []
    didi_forms = {}
    didi_invoices = {}
    dining = []
    self_drive_sheet = None   # 自驾车高速路行程单
    gas_invoices = []          # 加油费发票
    toll_invoices = []         # 通行费发票
    other = []

    for f in files:
        name = f["filename"]
        # 自驾车出差（优先识别，避免被交通票规则误匹配）
        if "高速路" in name and "行程单" in name:
            self_drive_sheet = f
        elif "加油费" in name:
            gas_invoices.append(f)
        elif "通行费" in name:
            toll_invoices.append(f)
        elif "登机牌" in name or re.match(r"\d{4}\s+\S+-\S+\s+\d+", name):
            transport.append(f)
        elif "住宿水单" in name or ("住宿" in name and name.endswith('.jpg')):
            hotel_singles.append(("水单", f))
        elif "住宿发票" in name:
            hotel_singles.append(("发票", f))
        elif "滴滴出行行程报销单" in name:
            m = re.search(r"单([A-Z]?)", name)
            letter = m.group(1) if m else ""
            didi_forms[letter] = f
        elif "滴滴电子发票" in name:
            m = re.search(r"票([A-Z]?)", name)
            letter = m.group(1) if m else ""
            didi_invoices[letter] = f
        elif "餐饮发票" in name:
            dining.append(f)
        else:
            other.append(f)

    # --- 酒店配对 ---
    hotel_by_city = defaultdict(list)
    for typ, f in hotel_singles:
        name = f["filename"]
        m = re.search(r"(\S+)住宿(水单|发票)", name)
        city = m.group(1) if m else "未知"
        hotel_by_city[city].append((typ, f))

    paired_hotels = []
    for city, items in hotel_by_city.items():
        waters = [(f, None) for t, f in items if t == "水单"]
        invs = [(f, None) for t, f in items if t == "发票"]
        paired = set()
        for wi, (w, _) in enumerate(waters):
            wname = w["filename"]
            (w_start, w_end) = parse_date_range(wname)
            for ii, (inv, _) in enumerate(invs):
                if ii in paired:
                    continue
                iname = inv["filename"]
                (i_start, i_end) = parse_date_range(iname)
                if w_start == i_start and w_end == i_end:
                    paired_hotels.append((w, inv))
                    paired.add(ii)
                    break
        for wi, (w, _) in enumerate(waters):
            if not any(w is p[0] for p in paired_hotels):
                paired_hotels.append((w, None))
        for ii, (inv, _) in enumerate(invs):
            if ii not in paired:
                paired_hotels.append((None, inv))

    paired_hotels.sort(key=lambda p: sort_key_hotel_by_checkin(p[0] or p[1]))

    # --- 滴滴配对 ---
    didi_paired = []
    all_letters = sorted(set(list(didi_forms.keys()) + list(didi_invoices.keys())))
    for letter in all_letters:
        form = didi_forms.get(letter)
        inv = didi_invoices.get(letter)
        didi_paired.append((letter, form, inv))

    # --- 排序 ---
    def transport_date(f):
        name = f["filename"]
        m = re.search(r"(\d{4})", name)
        if m:
            return parse_date_mmdd(m.group(1))
        return (99, 99)

    transport.sort(key=transport_date)
    dining.sort(key=lambda f: parse_date_range(f["filename"])[0])

    # --- 自驾车加油费/通行费排序并合并 ---
    gas_invoices.sort(key=lambda f: parse_date_range(f["filename"])[0])
    toll_invoices.sort(key=lambda f: (parse_date_range(f["filename"])[0], _toll_nn(f["filename"])))
    self_drive_invoices = gas_invoices + toll_invoices
    self_drive_invoices.sort(key=lambda f: _self_drive_sort_key(f["filename"]))

    return {
        "trip_dir": trip_dir,
        "transport": transport,
        "hotel_pairs": paired_hotels,
        "didi_paired": didi_paired,
        "dining": dining,
        "self_drive": {
            "sheet": self_drive_sheet,          # 高速路行程单（可能为 None）
            "invoices": self_drive_invoices,    # 合并排序后的加油费+通行费
            "gas_count": len(gas_invoices),
            "toll_count": len(toll_invoices),
        },
        "other": other
    }


def classify_local(base_dir=None):
    """本地目录特殊处理：按固定费用→出行→餐饮排列"""
    if base_dir is None:
        base_dir = get_base_dir()
    local_path = base_dir / "本地"

    # 尝试找到包含"本地"的目录
    if not local_path.exists():
        for entry in base_dir.iterdir():
            if entry.is_dir() and '本地' in entry.name:
                local_path = entry
                break
        else:
            return None

    transport = []
    hotel_pairs = []
    didi_forms = {}
    didi_invoices = {}
    dining = []
    telecom = []
    other = []

    for f in os.listdir(local_path):
        if f.startswith('.') or f.endswith(('.py', '.md')):
            continue
        full = local_path / f
        if not full.is_file():
            continue
        item = {"filename": f, "fullpath": str(full), "size": full.stat().st_size}

        if "通信发票" in f:
            telecom.append(item)
        elif "滴滴出行行程报销单" in f:
            m = re.search(r"单([A-Z]?)", f)
            didi_forms[m.group(1) if m else ""] = item
        elif "滴滴电子发票" in f:
            m = re.search(r"票([A-Z]?)", f)
            didi_invoices[m.group(1) if m else ""] = item
        elif "餐饮发票" in f:
            dining.append(item)
        else:
            other.append(item)

    telecom.sort(key=lambda x: parse_date_range(x["filename"])[0])
    dining.sort(key=lambda x: parse_date_range(x["filename"])[0])

    didi_paired = []
    all_letters = sorted(set(list(didi_forms.keys()) + list(didi_invoices.keys())))
    for letter in all_letters:
        didi_paired.append((letter, didi_forms.get(letter), didi_invoices.get(letter)))

    return {
        "trip_dir": local_path.name,
        "transport": transport,
        "hotel_pairs": hotel_pairs,
        "didi_paired": didi_paired,
        "dining": dining,
        "telecom": telecom,
        "other": other,
    }


# ── 输出格式化 ─────────────────────────────────
def generate_page_plan(base_dir=None):
    """为所有行程生成整合排列的逐页计划"""
    if base_dir is None:
        base_dir = get_base_dir()

    trip_dirs, local_dirs = discover_trips(base_dir)

    page_num = 0
    all_trips = []

    for trip_dir in trip_dirs:
        data = classify_and_sort(trip_dir, base_dir)
        if not data:
            continue
        all_trips.append(data)

        trip_pages = [f"━━━ 行程段: {trip_dir} ━━━"]

        # 自驾车出差：行程单单独一页 → 加油费+通行费(2×2)
        sd = data.get("self_drive", {})
        if sd.get("sheet"):
            page_num += 1
            trip_pages.append(f"  P{page_num:3d}  [自驾车行程单] {sd['sheet']['filename']}")
        for inv in sd.get("invoices", []):
            page_num += 1
            tag = "自驾车加油" if "加油费" in inv["filename"] else "自驾车通行"
            amt_match = re.search(r"(\d+\.?\d*)\.pdf", inv["filename"])
            amt_str = f"  ¥{amt_match.group(1)}" if amt_match else ""
            trip_pages.append(f"  P{page_num:3d}  [{tag}]   {inv['filename']}{amt_str}")

        for t in data["transport"]:
            page_num += 1
            name = t["filename"]
            amt_match = re.search(r"(\d+\.?\d*)\.pdf", name)
            amt_str = f"  ¥{amt_match.group(1)}" if amt_match else ""
            trip_pages.append(f"  P{page_num:3d}  [交通票]     {name}{amt_str}")

        for water, invoice in data["hotel_pairs"]:
            if water:
                page_num += 1
                trip_pages.append(f"  P{page_num:3d}  [住宿水单]   {water['filename']}")
            if invoice:
                page_num += 1
                trip_pages.append(f"  P{page_num:3d}  [住宿发票]   {invoice['filename']}")

        for letter, form, invoice in data["didi_paired"]:
            if form:
                page_num += 1
                trip_pages.append(f"  P{page_num:3d}  [滴滴行程]   {form['filename']}")
            if invoice:
                page_num += 1
                trip_pages.append(f"  P{page_num:3d}  [滴滴发票]   {invoice['filename']}")

        for d in data["dining"]:
            page_num += 1
            amt_match = re.search(r"(\d+\.?\d*)\.pdf", d["filename"])
            amt_str = f"  ¥{amt_match.group(1)}" if amt_match else ""
            trip_pages.append(f"  P{page_num:3d}  [餐饮发票]   {d['filename']}{amt_str}")

        for o in data["other"]:
            page_num += 1
            trip_pages.append(f"  P{page_num:3d}  [其他]       {o['filename']}")

        yield trip_dir, trip_pages

    # ── 本地费用 ──
    local_data = classify_local(base_dir)
    if local_data:
        trip_pages = [f"━━━ 本地费用 ━━━"]
        pn = page_num

        for t in local_data["telecom"]:
            pn += 1
            amt = re.search(r"(\d+\.?\d*)\.pdf", t["filename"])
            amt_str = f"  ¥{amt.group(1)}" if amt else ""
            trip_pages.append(f"  P{pn:3d}  [通信发票]   {t['filename']}{amt_str}")

        for letter, form, invoice in local_data["didi_paired"]:
            if form:
                pn += 1
                trip_pages.append(f"  P{pn:3d}  [滴滴行程]   {form['filename']}")
            if invoice:
                pn += 1
                trip_pages.append(f"  P{pn:3d}  [滴滴发票]   {invoice['filename']}")

        for d in local_data["dining"]:
            pn += 1
            amt = re.search(r"(\d+\.?\d*)\.pdf", d["filename"])
            amt_str = f"  ¥{amt.group(1)}" if amt else ""
            trip_pages.append(f"  P{pn:3d}  [餐饮发票]   {d['filename']}{amt_str}")

        for o in local_data["other"]:
            pn += 1
            trip_pages.append(f"  P{pn:3d}  [其他]       {o['filename']}")

        yield "本地", trip_pages


def main():
    base_dir = get_base_dir()
    trip_dirs, local_dirs = discover_trips(base_dir)

    print(f"工作目录: {base_dir}")
    print(f"发现 {len(trip_dirs)} 个行程目录, {len(local_dirs)} 个本地目录")
    print()

    all_trip_data = list(generate_page_plan(base_dir))

    lines = []
    lines.append("# 报销单据整合排列")
    lines.append("")
    lines.append("> 按整合排列规则生成：交通→酒店[水单→发票]→滴滴[行程单→发票]→餐饮")
    lines.append("")
    lines.append(f"> 工作目录: `{base_dir}`")
    lines.append("")
    lines.append("---")
    lines.append("")

    grand_total = 0
    for trip_dir, pages in all_trip_data:
        lines.append(f"## {trip_dir}")
        lines.append("")
        for p in pages:
            lines.append(p)
        lines.append("")
        grand_total += len(pages) - 1

    lines.append("---")
    lines.append(f"")
    lines.append(f"**总计 {grand_total} 页单据** (按整合排列规则生成)")
    lines.append("")

    if DRY_RUN:
        print('\n'.join(lines))
        return

    output_path = base_dir / "报销单据整合排列.md"
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    print(f"整合排列已生成: {output_path}")
    print(f"共 {grand_total} 页")
    print()

    print("=" * 60)
    print("验证清单")
    print("=" * 60)

    for trip_dir, pages in all_trip_data:
        print(f"\n{trip_dir}:")
        counts = defaultdict(int)
        for p in pages:
            for tag in ["交通票", "住宿水单", "住宿发票", "滴滴行程", "滴滴发票", "餐饮发票", "通信发票"]:
                if tag in p:
                    counts[tag] += 1

        w = counts["住宿水单"]
        i = counts["住宿发票"]
        if w > 0 or i > 0:
            status = "✓" if w == i else f"✗ 水单{w}≠发票{i}"
            print(f"  酒店: {w}水单 vs {i}发票  {status}")

        f_count = counts["滴滴行程"]
        e_count = counts["滴滴发票"]
        if f_count > 0 or e_count > 0:
            dd_status = "✓" if f_count == e_count else f"✗ 行程单{f_count}≠发票{e_count}"
            print(f"  滴滴: {f_count}行程单 vs {e_count}发票  {dd_status}")

    return output_path


if __name__ == "__main__":
    main()

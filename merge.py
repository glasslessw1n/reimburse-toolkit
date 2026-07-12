#!/usr/bin/env python3
"""
合并所有报销文件为一份整合PDF，严格按排列规则排序 - 通用版
用法:
  python3 merge.py                           # 在当前目录运行
  python3 merge.py --base-dir /path/to/reimbursements
  python3 merge.py --base-dir . --output 我的报销.pdf

输出: <base_dir>/报销单整合_yyyymm-mm.pdf (自动从目录名提取日期)
"""

import os, sys, re, tempfile, shutil
from pathlib import Path
from PyPDF2 import PdfMerger

sys.path.insert(0, str(Path(__file__).parent))
from organize import classify_and_sort, classify_local, discover_trips, get_base_dir


def parse_args():
    base_dir = get_base_dir()
    output_file = None
    i = 1
    while i < len(sys.argv):
        arg = sys.argv[i]
        if arg == '--base-dir' and i + 1 < len(sys.argv):
            base_dir = Path(sys.argv[i + 1])
            i += 2
        elif arg.startswith('--base-dir='):
            base_dir = Path(arg.split('=', 1)[1])
            i += 1
        elif arg == '--output' and i + 1 < len(sys.argv):
            output_file = Path(sys.argv[i + 1])
            i += 2
        elif arg.startswith('--output='):
            output_file = Path(arg.split('=', 1)[1])
            i += 1
        else:
            i += 1
    return base_dir, output_file


BASE_DIR, OUTPUT_FILE = parse_args()


def generate_output_filename(base_dir):
    trip_dirs, _ = discover_trips(base_dir)
    trip_dirs_sorted = sorted(trip_dirs)
    if trip_dirs_sorted:
        first_dates = re.findall(r'(\d{4})', trip_dirs_sorted[0])
        last_dates = re.findall(r'(\d{4})', trip_dirs_sorted[-1])
        if first_dates:
            start = first_dates[0]
            end = last_dates[-1] if last_dates else first_dates[-1]
            return base_dir / f"报销单整合_{start[:2]}{start[2:]}-{end[:2]}{end[2:]}.pdf"
    return base_dir / "报销单整合.pdf"


if OUTPUT_FILE is None:
    OUTPUT_FILE = generate_output_filename(BASE_DIR)


def build_ordered_file_list():
    """返回按排列规则排序的完整文件路径列表"""
    ordered = []
    trip_dirs, _ = discover_trips(BASE_DIR)

    for trip_dir in trip_dirs:
        data = classify_and_sort(trip_dir, BASE_DIR)
        if not data:
            continue

        for t in data["transport"]:
            ordered.append((trip_dir, "交通票", t["fullpath"]))

        for water, invoice in data["hotel_pairs"]:
            if water:
                ordered.append((trip_dir, "住宿水单", water["fullpath"]))
            if invoice:
                ordered.append((trip_dir, "住宿发票", invoice["fullpath"]))

        for letter, form, inv in data["didi_paired"]:
            if form:
                ordered.append((trip_dir, "滴滴行程单", form["fullpath"]))
            if inv:
                ordered.append((trip_dir, "滴滴发票", inv["fullpath"]))

        for d in data["dining"]:
            ordered.append((trip_dir, "餐饮发票", d["fullpath"]))

        for o in data["other"]:
            ordered.append((trip_dir, "其他", o["fullpath"]))

    local = classify_local(BASE_DIR)
    if local:
        for t in local.get("telecom", []):
            ordered.append(("本地", "通信发票", t["fullpath"]))
        for letter, form, inv in local.get("didi_paired", []):
            if form:
                ordered.append(("本地", "滴滴行程单", form["fullpath"]))
            if inv:
                ordered.append(("本地", "滴滴发票", inv["fullpath"]))
        for d in local.get("dining", []):
            ordered.append(("本地", "餐饮发票", d["fullpath"]))
        for o in local.get("other", []):
            ordered.append(("本地", "其他", o["fullpath"]))

    return ordered


def merge_pdfs(ordered_files):
    merger = PdfMerger()
    temp_dir = Path(tempfile.mkdtemp())
    success = 0
    failed = []

    for idx, (trip, typ, filepath) in enumerate(ordered_files, 1):
        source = Path(filepath)
        if not source.exists():
            failed.append(f"[不存在] {filepath}")
            continue

        merge_path = None
        if source.suffix.lower() == '.jpg':
            pdf_path = temp_dir / (source.stem + '.pdf')
            os.system(f'sips -s format pdf "{source}" --out "{pdf_path}" 2>/dev/null')
            if pdf_path.exists():
                merge_path = pdf_path
            else:
                failed.append(f"[JPG转换失败] {filepath}")
                continue
        elif source.suffix.lower() == '.pdf':
            merge_path = source
        else:
            failed.append(f"[不支持格式] {filepath}")
            continue

        try:
            merger.append(str(merge_path))
            success += 1
            if idx % 20 == 0:
                print(f"  已合并 {idx}/{len(ordered_files)} ...")
        except Exception as e:
            failed.append(f"[合并失败] {filepath}: {e}")

    print(f"  成功合并 {success} 个文件")

    if failed:
        print(f"  失败 {len(failed)} 个:")
        for f in failed:
            print(f"    {f}")

    print(f"  正在写入 {OUTPUT_FILE} ...")
    merger.write(str(OUTPUT_FILE))
    merger.close()

    shutil.rmtree(temp_dir, ignore_errors=True)
    return OUTPUT_FILE, success, failed


def main():
    print("=" * 60)
    print(f"工作目录: {BASE_DIR}")
    print(f"输出文件: {OUTPUT_FILE}")
    print("=" * 60)
    print()
    print("正在按排列规则生成文件列表...")
    ordered = build_ordered_file_list()
    print(f"共 {len(ordered)} 个文件待合并\n")

    print("正在合并PDF...")
    output, success, failed = merge_pdfs(ordered)

    print()
    print("=" * 60)
    if output.exists():
        size_mb = output.stat().st_size / (1024 * 1024)
        print(f"生成成功: {output}")
        print(f"文件大小: {size_mb:.1f} MB")
        print(f"合并文件数: {success} 个")
    else:
        print("生成失败!")
    print("=" * 60)


if __name__ == "__main__":
    main()

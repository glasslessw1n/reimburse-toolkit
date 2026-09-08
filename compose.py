#!/usr/bin/env python3
"""
按排版规则生成整合报销PDF - 通用版
用法:
  python3 compose.py                           # 在当前目录运行
  python3 compose.py --base-dir /path/to/reimbursements
  python3 compose.py --base-dir . --output 我的报销.pdf
  python3 compose.py --base-dir . --zoom 1.5   # 调整渲染质量

输出: <base_dir>/报销单整合_yyyymm-mm.pdf (自动从目录名提取日期)
"""

import os, sys, re, tempfile, shutil, math, atexit
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, str(Path(__file__).parent))
from organize import classify_and_sort, classify_local, discover_trips, get_base_dir

import fitz  # PyMuPDF

# ── 参数解析 ───────────────────────────────────
def parse_args():
    base_dir = get_base_dir()
    output_file = None
    zoom = 2.0
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
        elif arg == '--zoom' and i + 1 < len(sys.argv):
            zoom = float(sys.argv[i + 1])
            i += 2
        elif arg.startswith('--zoom='):
            zoom = float(arg.split('=', 1)[1])
            i += 1
        else:
            i += 1
    return base_dir, output_file, zoom

BASE_DIR, OUTPUT_FILE, RENDER_ZOOM = parse_args()

# ── 自动生成输出文件名 ─────────────────────────
def generate_output_filename(base_dir):
    """从目录名中提取日期范围生成文件名"""
    trip_dirs, _ = discover_trips(base_dir)
    trip_dirs_sorted = sorted(trip_dirs)

    if trip_dirs_sorted:
        # 尝试从第一个和最后一个目录名中提取日期
        first_dates = re.findall(r'(\d{4})', trip_dirs_sorted[0])
        last_dates = re.findall(r'(\d{4})', trip_dirs_sorted[-1])
        if first_dates:
            start = first_dates[0]  # 如 "0303"
            end = last_dates[-1] if last_dates else first_dates[-1]  # 如 "0428"
            suffix = f"{start[:2]}{end[:2]}"  # "0304"
            return base_dir / f"报销单整合_{start[:2]}{start[2:]}-{end[:2]}{end[2:]}.pdf"

    return base_dir / "报销单整合.pdf"


if OUTPUT_FILE is None:
    OUTPUT_FILE = generate_output_filename(BASE_DIR)

# ── 页面几何参数 ──────────────────────────────
A4_W, A4_H = 595.28, 841.89
MARGIN = 20.0
GAP = 8.0


def cell_rect(col, row, cols, rows, page_w=A4_W, page_h=A4_H):
    total_gap_x = GAP * (cols - 1)
    total_gap_y = GAP * (rows - 1)
    cell_w = (page_w - 2 * MARGIN - total_gap_x) / cols
    cell_h = (page_h - 2 * MARGIN - total_gap_y) / rows
    x0 = MARGIN + col * (cell_w + GAP)
    y0 = MARGIN + row * (cell_h + GAP)
    return fitz.Rect(x0, y0, x0 + cell_w, y0 + cell_h)


# ── 渲染工具 ───────────────────────────────────
def render_page_to_pixmap(doc, page_idx=0):
    page = doc[page_idx]
    mat = fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)
    return page.get_pixmap(matrix=mat)


def render_page_rotated(doc, page_idx=0):
    page = doc[page_idx]
    mat = fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM).prerotate(-90)
    return page.get_pixmap(matrix=mat)


def place_pixmap(page, pix, rect):
    cell_w = rect.width
    cell_h = rect.height
    pix_w_pt = pix.width / RENDER_ZOOM
    pix_h_pt = pix.height / RENDER_ZOOM
    scale = min(cell_w / pix_w_pt, cell_h / pix_h_pt)
    fit_w = pix_w_pt * scale
    fit_h = pix_h_pt * scale
    x0 = rect.x0 + (cell_w - fit_w) / 2
    y0 = rect.y0 + (cell_h - fit_h) / 2
    page.insert_image(fitz.Rect(x0, y0, x0 + fit_w, y0 + fit_h), pixmap=pix)


def open_doc_safe(filepath):
    path = Path(filepath)
    if path.suffix.lower() == '.jpg':
        tmpdir = Path(tempfile.mkdtemp())
        pdf_path = tmpdir / (path.stem + '.pdf')
        os.system(f'sips -s format pdf "{path}" --out "{pdf_path}" 2>/dev/null')
        if pdf_path.exists():
            doc = fitz.open(str(pdf_path))
            doc._tmpdir = tmpdir
            return doc
        return None
    return fitz.open(str(filepath))


def close_doc(doc):
    try:
        doc.close()
    except:
        pass
    if hasattr(doc, '_tmpdir'):
        try:
            shutil.rmtree(doc._tmpdir, ignore_errors=True)
        except:
            pass


# ── 排版函数 ──────────────────────────────────

def compose_transport_pair(outbound_path, return_path, output_doc):
    """行程证明页(P1): 去程+回程 1列2行(上下)"""
    page = output_doc.new_page(width=A4_W, height=A4_H)
    for row, filepath in enumerate([outbound_path, return_path]):
        doc = open_doc_safe(filepath)
        if not doc:
            continue
        pix = render_page_to_pixmap(doc, 0)
        close_doc(doc)
        r = cell_rect(0, row, 1, 2)
        place_pixmap(page, pix, r)


def compose_transport_3x2(sources, output_doc):
    """全部非登机牌交通票: 3×2, 每页最多6张"""
    page = None
    for i, (label, filepath) in enumerate(sources):
        col = i % 3
        row = (i % 6) // 3
        if col == 0 and row == 0:
            page = output_doc.new_page(width=A4_W, height=A4_H)
        doc = open_doc_safe(filepath)
        if not doc:
            continue
        pix = render_page_to_pixmap(doc, 0)
        close_doc(doc)
        r = cell_rect(col, row, 3, 2)
        place_pixmap(page, pix, r)


def compose_hotel_water_bill(filepath, output_doc):
    """酒店水单: 单页满版"""
    doc = open_doc_safe(filepath)
    if not doc:
        return
    page = output_doc.new_page(width=A4_W, height=A4_H)
    pix = render_page_to_pixmap(doc, 0)
    close_doc(doc)
    full_rect = fitz.Rect(MARGIN, MARGIN, A4_W - MARGIN, A4_H - MARGIN)
    place_pixmap(page, pix, full_rect)


def compose_hotel_invoice_1x2(filepath, output_doc):
    """酒店发票: 1列2行, 同张发票上下各一份"""
    doc = open_doc_safe(filepath)
    if not doc:
        return
    pix = render_page_to_pixmap(doc, 0)
    close_doc(doc)
    page = output_doc.new_page(width=A4_W, height=A4_H)
    for row in range(2):
        r = cell_rect(0, row, 1, 2)
        place_pixmap(page, pix, r)


def compose_didi_forms(sources, output_doc):
    """滴滴行程单: =1张→整页不旋转; >1张→左旋90°, 1列2行"""
    if len(sources) == 1:
        for label, filepath in sources:
            doc = open_doc_safe(filepath)
            if not doc:
                continue
            pix = render_page_to_pixmap(doc, 0)
            close_doc(doc)
            page = output_doc.new_page(width=A4_W, height=A4_H)
            full_rect = fitz.Rect(MARGIN, MARGIN, A4_W - MARGIN, A4_H - MARGIN)
            place_pixmap(page, pix, full_rect)
    else:
        page = None
        for i, (label, filepath) in enumerate(sources):
            row = i % 2
            if row == 0:
                page = output_doc.new_page(width=A4_W, height=A4_H)
            doc = open_doc_safe(filepath)
            if not doc:
                continue
            pix = render_page_rotated(doc, 0)
            close_doc(doc)
            r = cell_rect(0, row, 1, 2)
            place_pixmap(page, pix, r)


def compose_didi_invoices(sources, output_doc):
    """滴滴发票: =1张→左旋90°整页; >1张→左旋90°, 2×2"""
    if len(sources) == 1:
        label, filepath = sources[0]
        doc = open_doc_safe(filepath)
        if doc:
            pix = render_page_rotated(doc, 0)
            close_doc(doc)
            page = output_doc.new_page(width=A4_W, height=A4_H)
            full_rect = fitz.Rect(MARGIN, MARGIN, A4_W - MARGIN, A4_H - MARGIN)
            place_pixmap(page, pix, full_rect)
    else:
        page = None
        for i, (label, filepath) in enumerate(sources):
            col = i % 2
            row = (i % 4) // 2
            if col == 0 and row == 0:
                page = output_doc.new_page(width=A4_W, height=A4_H)
            doc = open_doc_safe(filepath)
            if not doc:
                continue
            pix = render_page_rotated(doc, 0)
            close_doc(doc)
            r = cell_rect(col, row, 2, 2)
            place_pixmap(page, pix, r)


def compose_dining_2x2(sources, output_doc):
    """餐饮发票: 左旋90°, 2×2"""
    page = None
    for i, (label, filepath) in enumerate(sources):
        col = i % 2
        row = (i % 4) // 2
        if col == 0 and row == 0:
            page = output_doc.new_page(width=A4_W, height=A4_H)
        doc = open_doc_safe(filepath)
        if not doc:
            continue
        pix = render_page_rotated(doc, 0)
        close_doc(doc)
        r = cell_rect(col, row, 2, 2)
        place_pixmap(page, pix, r)


def compose_self_drive_sheet(filepath, output_doc):
    """自驾车高速路行程单: 单独一页, 整页满版, 不旋转"""
    doc = open_doc_safe(filepath)
    if not doc:
        return
    page = output_doc.new_page(width=A4_W, height=A4_H)
    pix = render_page_to_pixmap(doc, 0)
    close_doc(doc)
    full_rect = fitz.Rect(MARGIN, MARGIN, A4_W - MARGIN, A4_H - MARGIN)
    place_pixmap(page, pix, full_rect)


def compose_self_drive_2x2(sources, output_doc):
    """自驾车加油费+通行费发票: 左旋90°, 2×2, 每页最多4张"""
    page = None
    for i, (label, filepath) in enumerate(sources):
        col = i % 2
        row = (i % 4) // 2
        if col == 0 and row == 0:
            page = output_doc.new_page(width=A4_W, height=A4_H)
        doc = open_doc_safe(filepath)
        if not doc:
            continue
        pix = render_page_rotated(doc, 0)
        close_doc(doc)
        r = cell_rect(col, row, 2, 2)
        place_pixmap(page, pix, r)


def compose_1x2_vertical(sources, output_doc, duplicate=False):
    """1列2行通用排版(上下)"""
    if duplicate:
        for label, filepath in sources:
            doc = open_doc_safe(filepath)
            if not doc:
                continue
            pix = render_page_to_pixmap(doc, 0)
            close_doc(doc)
            page = output_doc.new_page(width=A4_W, height=A4_H)
            for row in range(2):
                r = cell_rect(0, row, 1, 2)
                place_pixmap(page, pix, r)
    else:
        page = None
        for i, (label, filepath) in enumerate(sources):
            row = i % 2
            if row == 0:
                page = output_doc.new_page(width=A4_W, height=A4_H)
            doc = open_doc_safe(filepath)
            if not doc:
                continue
            pix = render_page_to_pixmap(doc, 0)
            close_doc(doc)
            r = cell_rect(0, row, 1, 2)
            place_pixmap(page, pix, r)


# ── 主流程 ─────────────────────────────────────
def build_all():
    """主入口: 生成完整整合PDF"""
    output = fitz.open()
    trip_dirs, _ = discover_trips(BASE_DIR)

    for trip_dir in trip_dirs:
        data = classify_and_sort(trip_dir, BASE_DIR)
        if not data:
            continue

        transport = data["transport"]
        hotel_pairs = data["hotel_pairs"]
        didi_paired = data["didi_paired"]
        dining = data["dining"]
        self_drive = data.get("self_drive", {})

        # 自驾车出差：行程单单独一页 → 加油费+通行费 2×2 (左旋90°)
        sd_sheet = self_drive.get("sheet")
        sd_invoices = self_drive.get("invoices", [])
        if sd_sheet:
            compose_self_drive_sheet(sd_sheet["fullpath"], output)
        if sd_invoices:
            sd_src = [(f["filename"], f["fullpath"]) for f in sd_invoices]
            compose_self_drive_2x2(sd_src, output)

        transport_sorted = sorted(transport,
            key=lambda t: re.search(r"(\d{4})", t["filename"]).group(1)
                          if re.search(r"(\d{4})", t["filename"]) else "9999")

        outbound = transport_sorted[0] if len(transport_sorted) > 0 else None
        returnt = transport_sorted[-1] if len(transport_sorted) > 1 else None

        # P1: 去程+回程(1列2行)
        if outbound and returnt:
            compose_transport_pair(outbound["fullpath"], returnt["fullpath"], output)

        # 全部非登机牌交通票 → 3×2
        non_boarding = [t for t in transport_sorted if "登机牌" not in t["filename"]]
        if non_boarding:
            nb_src = [(t["filename"], t["fullpath"]) for t in non_boarding]
            compose_transport_3x2(nb_src, output)

        # 酒店: 水单 → 发票(1列2行), 按对穿插
        for water, invoice in hotel_pairs:
            if water:
                compose_hotel_water_bill(water["fullpath"], output)
            if invoice:
                compose_hotel_invoice_1x2(invoice["fullpath"], output)

        # 滴滴行程单
        forms = [(f"行程单{l}", f["fullpath"]) for l, f, _ in didi_paired if f]
        if forms:
            compose_didi_forms(forms, output)

        # 滴滴发票
        dinvs = [(f"发票{l}", i["fullpath"]) for l, _, i in didi_paired if i]
        if dinvs:
            compose_didi_invoices(dinvs, output)

        # 餐饮: 2×2
        if dining:
            dining_src = [(d["filename"], d["fullpath"]) for d in dining]
            compose_dining_2x2(dining_src, output)

    # ── 本地费用 ──
    local = classify_local(BASE_DIR)
    if local:
        if local.get("telecom"):
            telecom_src = [(t["filename"], t["fullpath"]) for t in local["telecom"]]
            compose_1x2_vertical(telecom_src, output, duplicate=False)

        lf = [(f"行程单{l}", f["fullpath"]) for l, f, _ in local.get("didi_paired", []) if f]
        if lf:
            compose_didi_forms(lf, output)

        ld = [(f"发票{l}", i["fullpath"]) for l, _, i in local.get("didi_paired", []) if i]
        if ld:
            compose_didi_invoices(ld, output)

        if local.get("dining"):
            dining_src = [(d["filename"], d["fullpath"]) for d in local["dining"]]
            compose_dining_2x2(dining_src, output)

    # 保存
    output.save(str(OUTPUT_FILE), garbage=3, deflate=True, clean=True)
    output.close()

    info = fitz.open(str(OUTPUT_FILE))
    pc = len(info)
    info.close()
    sz = OUTPUT_FILE.stat().st_size / (1024 * 1024)
    print(f"整合PDF: {OUTPUT_FILE}")
    print(f"总页数: {pc}, 大小: {sz:.1f} MB")
    print(f"渲染倍数: {RENDER_ZOOM}x")
    return pc


if __name__ == "__main__":
    print("=" * 60)
    print(f"工作目录: {BASE_DIR}")
    print(f"输出文件: {OUTPUT_FILE}")
    print(f"RENDER_ZOOM: {RENDER_ZOOM}")
    print("=" * 60)
    pages = build_all()
    print("完成.")


@atexit.register
def cleanup_tmp():
    # 清理残留临时文件
    pass

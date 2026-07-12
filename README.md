# 报销单据整合整理工具包

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![License MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

将散乱的差旅报销文件（发票、行程单、水单、登机牌等）按照财务规范自动**分类、配对、排版**，生成可直接打印提交的整合 PDF，同时支持导出 Excel 明细表。

---

## 目录

- [快速开始](#快速开始)
- [目录结构要求](#目录结构要求)
- [文件命名规范](#文件命名规范)
- [工具说明](#工具说明)
- [整合排列规则](#整合排列规则)
- [排版布局规则](#排版布局规则)
- [生成 Excel 明细](#生成-excel-明细)
- [AI Agent / 智能体集成](#ai-agent--智能体集成)
- [依赖](#依赖)

---

## 快速开始

```bash
# 1. 克隆仓库
git clone https://github.com/your-org/reimburse-toolkit.git
cd reimburse-toolkit

# 2. 安装依赖
pip install -r requirements.txt

# 3. 扫描报销目录
python3 reimburse.py summary --base-dir /path/to/bills

# 4. 检查配对
python3 reimburse.py check --base-dir /path/to/bills

# 5. 生成整合 PDF
python3 compose.py --base-dir /path/to/bills

# 6. 导出 Excel 明细
python3 gen_excel.py --base-dir /path/to/bills
```

---

## 目录结构要求

报销文件需按以下结构组织，工具会**自动发现**所有行程目录：

```
bills/yyyyMM-yyMM/                         # 按报销月份归档
├── MMDD-MMDD 城市1 城市2 .../              # 差旅行程目录
│   ├── MMDD 出发-到达 登机牌.pdf
│   ├── MMDD 出发-到达 金额.pdf             # 火车票/机票
│   ├── MMDD-MMDD 城市住宿水单.pdf/.jpg
│   ├── MMDD-MMDD 城市住宿发票.pdf
│   ├── 滴滴出行行程报销单A.pdf
│   ├── 滴滴电子发票A.pdf
│   └── MMDD 餐饮发票 金额.pdf
├── 本地/                                   # 非差旅消费
│   ├── MM 通信发票 金额.pdf
│   ├── MMDD 餐饮发票 金额.pdf
│   ├── 滴滴出行行程报销单X.pdf
│   └── 滴滴电子发票X.pdf
└── ...
```

- 每个差旅一个子目录，命名建议 `MMDD-MMDD 城市1 城市2 ...`
- 本地消费放在名称含「本地」的目录中
- 支持 `.pdf` 和 `.jpg` 格式

---

## 文件命名规范

工具通过文件名关键词自动识别类型，请按以下规范命名：

### 交通票
```
登机牌:  MMDD 出发-到达 登机牌.pdf      例: 0303 重庆-北京 登机牌.pdf
火车票:  MMDD 出发-到达 金额.pdf        例: 0402 广州-南宁 333.pdf
```

### 酒店
```
水单:    MMDD-MMDD 城市住宿水单.pdf/.jpg  例: 0303-0304 北京住宿水单.jpg
发票:    MMDD-MMDD 城市住宿发票.pdf      例: 0303-0304 北京住宿发票.pdf
```

### 滴滴出行
```
行程单:  滴滴出行行程报销单A.pdf           (多段行程用 A/B/C 后缀)
电子发票: 滴滴电子发票A.pdf               (与行程单字母一一对应)
```

### 其他
```
餐饮:    MMDD 餐饮发票 金额.pdf          例: 0306 餐饮发票 1356.pdf
通信:    MM 通信发票 金额.pdf            例: 03 通信发票 197.5.pdf
```

---

## 工具说明

| 工具 | 用途 | 输出 |
|------|------|------|
| `reimburse.py` | 扫描诊断：文件概况、配对核对、文本提取 | 终端输出 / Markdown 报告 |
| `organize.py` | 分类排序：按规则排列并生成逐页计划 | `报销单据整合排列.md` |
| `compose.py` | 排版整合：网格布局 + 旋转适配 | `报销单整合_MMDD-MMDD.pdf` |
| `merge.py` | 简单合并：按规则排序直接拼接 | `报销单整合_MMDD-MMDD.pdf` |
| `gen_excel.py` | 生成明细表：包含各类别金额逐项明细 | `报销明细_yyyymm-mm.xlsx` |

所有工具均支持 `--base-dir <目录>` 指定工作目录，不指定则使用当前目录。

### reimburse.py 子命令

```bash
python3 reimburse.py scan --base-dir <目录>     # 逐文件列出分类
python3 reimburse.py summary --base-dir <目录>  # 按行程汇总统计
python3 reimburse.py check --base-dir <目录>    # 核对滴滴配对
python3 reimburse.py report --base-dir <目录>   # 生成 Markdown 报告
python3 reimburse.py text --base-dir <目录>     # 提取所有 PDF 文本
```

### compose.py 参数

```bash
python3 compose.py --base-dir <目录> [--output 文件名.pdf] [--zoom 2.0]
```

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--base-dir` | 报销文件所在目录 | 当前目录 |
| `--output` | 输出 PDF 文件名 | 自动生成 `报销单整合_MMDD-MMDD.pdf` |
| `--zoom` | 渲染质量倍数 | 2.0 |

### gen_excel.py 参数

```bash
python3 gen_excel.py --base-dir <目录> [--output 文件名.xlsx]
```

生成的 Excel 包含每个差旅区间的：
- 铁路交通：日期、金额、起始站（登机牌标注无金额）
- 酒店住宿：入住-退房日期、PDF 提取金额、城市、是否缺发票
- 滴滴出行：段数、发票总额、返回城市及日期
- 餐饮/通信：日期、月份、金额

---

## 整合排列规则

工具按以下 9 条规则自动整理：

| # | 规则 | 说明 |
|---|------|------|
| 1 | **行程分段** | 每个行程独立成段，按时间排序 |
| 2 | **四层顺序** | 交通 → 酒店[水单→发票] → 滴滴[行程单→发票] → 餐饮 |
| 3 | **酒店紧耦合** | 每家酒店水单和发票成对出现，水单在前发票在后 |
| 4 | **日期排序** | 同类单据按日期升序排列 |
| 5 | **滴滴配对** | 行程单和发票按字母编号 (A/B/C...) 配对 |
| 6 | **交通集中** | 交通票集中在行程段首 |
| 7 | **餐饮在最后** | 餐饮票据放在每段末尾 |
| 8 | **代订酒店** | 第三方代订(无发票)仅展示水单；有发票同时展示水单和发票 |
| 9 | **跨月连续性** | 跨月行程保持完整不间断 |

---

## 排版布局规则

| 类型 | 布局 | 旋转 | 说明 |
|------|------|------|------|
| 行程证明页 P1 | 1列2行 | 无 | 去程在上，回程在下 |
| 交通票（非登机牌） | 3×2 | 无 | 每页最多 6 张 |
| 酒店水单 | 整页满版 | 无 | - |
| 酒店发票 | 1列2行 | 无 | 同一张发票上下各排一份 |
| 滴滴行程单（仅1张） | 整页满版 | 无 | - |
| 滴滴行程单（多张） | 1列2行 | 左旋90° | 上下排列 |
| 滴滴发票（仅1张） | 整页满版 | 左旋90° | - |
| 滴滴发票（多张） | 2×2 | 左旋90° | 每页最多 4 张 |
| 餐饮发票 | 2×2 | 左旋90° | 每页最多 4 张 |
| 通信发票 | 1列2行 | 无 | 上下排列 |

---

## 生成 Excel 明细

```bash
python3 gen_excel.py --base-dir /path/to/bills
```

自动从 PDF 发票中提取价税合计金额，按差旅区间逐项列出明细。

---

## AI Agent / 智能体集成

### 集成方式

本工具包可集成到 AI Agent（如 Claude、GPT 等）的工作流中，实现自动化报销处理。

#### 方式一：命令行调用（推荐）

在你的 Agent 配置中添加工具调用：

```json
{
  "tools": [
    {
      "name": "scan_reimbursements",
      "command": "python3 /path/to/reimburse-toolkit/reimburse.py",
      "args": ["summary", "--base-dir", "${BILLS_DIR}"]
    },
    {
      "name": "check_pairing",
      "command": "python3 /path/to/reimburse-toolkit/reimburse.py",
      "args": ["check", "--base-dir", "${BILLS_DIR}"]
    },
    {
      "name": "generate_pdf",
      "command": "python3 /path/to/reimburse-toolkit/compose.py",
      "args": ["--base-dir", "${BILLS_DIR}"]
    },
    {
      "name": "generate_excel",
      "command": "python3 /path/to/reimburse-toolkit/gen_excel.py",
      "args": ["--base-dir", "${BILLS_DIR}"]
    }
  ]
}
```

#### 方式二：Python 模块导入

```python
import sys
sys.path.insert(0, '/path/to/reimburse-toolkit')

from organize import discover_trips, classify_and_sort, classify_local
from reimburse import scan_directory, check_didi_pairing

# 扫描分类
trips, local = discover_trips(base_dir)

# 按行程处理
for trip_dir in trips:
    data = classify_and_sort(trip_dir, base_dir)
    # data["transport"], data["hotel_pairs"], data["didi_paired"], data["dining"]
```

#### 方式三：Claude Code / Cursor Skill

将工具包复制到项目的 `.claude/skills/` 目录，即可通过自然语言触发：

```
"帮我整理 /path/to/bills 的报销单据"
```

Agent 会自动调用本工具包的 `reimburse.py` → `organize.py` → `compose.py` 流程。

### Agent 收取邮件附件工作流

结合 `collect_rule.md` 中的识别规则，Agent 可自动处理邮件附件：

```
收到邮件附件 →
  1. 打开附件，pdftotext 提取 PDF 内容 / OCR 识别图片
  2. 按 collect_rule.md 的识别规则判断票据类型
  3. 从内容中提取：日期、城市、金额、行程信息
  4. 按命名规范生成文件名并保存到对应行程目录
  5. 记录操作日志
  6. 调用 compose.py 更新整合 PDF
```

详细规则见 [collect_rule.md](collect_rule.md)。

---

## 依赖

| 依赖 | 用途 | 安装 |
|------|------|------|
| Python 3.8+ | 运行环境 | - |
| PyMuPDF (fitz) | compose.py 网格排版渲染 | `pip install PyMuPDF` |
| PyPDF2 | merge.py PDF 合并 | `pip install PyPDF2` |
| openpyxl | gen_excel.py Excel 生成 | `pip install openpyxl` |
| pdftotext | 金额提取（poppler-utils） | macOS: `brew install poppler` / Linux: `apt install poppler-utils` |

```bash
pip install -r requirements.txt
```

macOS 系统自带 `sips` 命令用于 JPG 转 PDF，无需额外安装。

---

## 常见问题

**Q: 酒店发票金额提取失败？**
A: 确保安装了 `poppler-utils`（提供 `pdftotext`）。某些 PDF 的金额被排版拆分，工具已内置拼接修复逻辑。

**Q: 滴滴行程单和发票如何配对？**
A: 通过文件名中的字母后缀自动匹配。`滴滴出行行程报销单A.pdf` ↔ `滴滴电子发票A.pdf`。

**Q: 登机牌散落在单独目录怎么办？**
A: 将登机牌按日期手动归位到对应行程目录，或使用 `collect_rule.md` 的规则让 Agent 自动归位。

---

##  License

MIT License - 详见 [LICENSE](LICENSE) 文件。

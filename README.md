# 发票助手（PDF 批量提取 + GUI 导出）

用于批量读取发票 PDF，并导出为 `CSV` / `XLSX` 的桌面工具。

## 功能

- 批量添加 PDF，支持调整处理顺序（上移/下移）
- 合并整理（一个总表）或分开整理（每个 PDF 一个文件）
- 自动提取销售方、购买方税号、发票号码、开票日期、货物明细等字段
- 兼容明细表头拆分或字段缺失的版式（会尽量保留“产品名称 / 金额 / 税额”）
- XLSX 导出支持表头样式与列宽优化



## 模板化提取策略（稳定优先）

当前版本优先使用 **固定版式模板 + 锚点行提取**，不依赖预置“单位词典”来猜测字段。

### 已内置模板

- `cn_e_invoice_common_tall_v1`：普通发票长版（页面高度 > 430）
- `cn_vat_special_compact_v1`：增值税专用发票短版
- `cn_e_invoice_common_compact_v1`：普通发票短版（通用）
- `cn_e_invoice_common_compact_v2_firegear`：普通发票短版（消防器材类样本）
- `cn_e_invoice_common_compact_v3_shoes`：普通发票短版（鞋类样本）
- `cn_e_invoice_common_compact_v4_packaging`：普通发票短版（塑料包装样本）

### 结构解剖方法（pdfplumber + pandas）

1. **模板识别**：先读标题区（如“普通发票 / 增值税专用发票”）和页面高度，再在候选模板中评分选最优。
2. **固定字段提取**：对票号、日期、购销方、价税合计使用固定 `bbox` 裁剪。
3. **明细行定位**：用数量列数字作为锚点（anchor row），按相邻锚点生成行区间。
4. **按列切片**：每行再按固定 `x` 边界提取 `item_name/model/unit/quantity/unit_price/amount/tax_rate/tax_amount`。
5. **结构化落表**：明细使用 `pandas.DataFrame.from_records` 思路构造记录（GUI 中最终导出 CSV/XLSX）。

### 为什么这样更稳

- 先过滤页外脏字符（`0 <= top/bottom <= page.height`），避免隐藏文字污染字段。
- 发票明细区域常见“看起来像表格，但列靠文字对齐”的情况，直接 `extract_table()` 容易糊行。
- 锚点 + 列边界可以兼容：
  - 明细换行
  - 部分列缺失
  - 轻微偏移的版式
- 不依赖硬编码单位词列表，减少模板外文件的误拆风险。

### 失败回退

若模板识别失败，会自动回退到原有文本/表格提取逻辑，以保持兼容性。

## 环境要求

- Python 3.9+
- Windows / macOS / Linux

## 安装

```bash
pip install -r requirements.txt
```

## 运行

```bash
python invoice_gui_extractor.py
```

## 使用步骤

1. 点击 **添加PDF** 选择需要处理的文件。
2. 通过 **上移/下移** 调整处理顺序。
3. 在“导出设置”中选择：
   - 合并整理（一个总表）
   - 分开整理（每个 PDF 单独文件）
4. 选择导出格式（推荐 `XLSX`）。
5. 点击 **开始整理并导出**，选择输出目录。
6. 完成后程序会直接打开生成的文件（合并模式；分开模式仅单文件时自动打开）。

## 导出列顺序（当前版本）

1. 空
2. 公司名称(销售方)
3. 纳税人识别号(购买方税号)
4. 发票编码(发票号码)
5. 开票日期
6. 空2
7. 产品名称
8. 型号(规格型号)
9. 数量
10. 金额
11. 税额
12. 总价

## 打包 EXE（Windows，可在**无 Python 环境**电脑直接运行）

> 如果你的目标机器包含 **Windows 7**，请优先看后面的「**Win7 兼容打包方案**」。

### 1) 安装打包工具

```bash
pip install pyinstaller
```

### 2) 直接打包为**单个 exe**（推荐）

在项目根目录运行：

```bash
pyinstaller --noconfirm --clean --onefile --windowed --name fapiaozhushou invoice_gui_extractor.py
```

> 说明：`--onefile` 会把依赖全部打进一个 `fapiaozhushou.exe`，目标电脑即使没有安装 Python 也可以直接运行。

### 3) 产物位置

- 最终可执行文件：`dist/fapiaozhushou.exe`
- 这是单文件交付方式，直接把这个 exe 发给别人即可（无需再带整个文件夹）

### 4) 常用补充参数（可选）

- 添加图标：`--icon app.ico`
- 关闭控制台黑窗：`--windowed`（GUI 程序建议保留）
- 每次重打包前清理缓存：`--clean`（已包含）

例如：

```bash
pyinstaller --noconfirm --clean --onefile --windowed --name fapiaozhushou --icon app.ico invoice_gui_extractor.py
```

### 5) 首次运行说明

- 单文件模式首次启动会先在临时目录解包，可能比目录模式稍慢，这是正常现象。
- 若杀毒软件拦截，请将 `dist/fapiaozhushou.exe` 加白名单后再测试。

### 6) 自检（在打包机器上）

1. 双击 `dist/fapiaozhushou.exe`，确认界面可打开。
2. 用 1 份样例 PDF 测试导出 `XLSX`。
3. 再拷贝到**没有 Python** 的电脑上复测一次。

## Win7 兼容打包方案（重点）

你反馈的现象本质上是：

- 新版 Python / 打包器生成的程序，可能依赖 Win7 不具备的系统 API；
- 或依赖了较新的 VC++ 运行库，而 Win7 上安装困难。

下面给两种可行方案。

### 方案 A（推荐）：使用 Python 3.8 + 旧版 PyInstaller，在 Win7 环境打包

这是最稳妥的方式，核心原则是：**用“最低目标系统”来打包**。

#### 1) 准备环境

- Python：`3.8.10`（最后一个 3.8 版本，兼容 Win7）
- PyInstaller：`5.13.2`
- 依赖按 `requirements.txt` 安装

示例：

```bash
py -3.8 -m pip install -U pip
py -3.8 -m pip install pyinstaller==5.13.2 -r requirements.txt
```

#### 2) 优先使用 `--onedir`（比 onefile 更稳）

```bash
py -3.8 -m PyInstaller --noconfirm --clean --onedir --windowed --name fapiaozhushou_win7 invoice_gui_extractor.py
```

生成目录在：`dist/fapiaozhushou_win7/`

> 建议先交付 `onedir` 版本给 Win7 用户测试，稳定后再考虑 `--onefile`。

#### 3) 如果必须单文件，再尝试 `--onefile`

```bash
py -3.8 -m PyInstaller --noconfirm --clean --onefile --windowed --name fapiaozhushou_win7 invoice_gui_extractor.py
```

若单文件在个别 Win7 机器启动慢或被拦截，回退到 `--onedir`。

#### 4) 关键注意事项

1. **尽量在 Win7 虚拟机或真机里打包**，不要在 Win10/11 打包后直接丢给 Win7。
2. Win7 建议先安装系统补丁：`KB2999226 (Universal CRT)` 与 `SHA-2` 相关更新。
3. 若报缺少运行库 DLL，把打包目录内（或 Python 安装目录中的）`vcruntime140.dll` 一并带上。

---

### 方案 B：Nuitka + MinGW64（避免依赖用户安装 VC++）

当用户环境无法安装 VC++ 运行库时，可尝试 Nuitka 的 MinGW 构建链，运行时不要求用户再手动装 VC++。

> 仍建议使用 Python 3.8，并在 Win7 环境实机验证。

```bash
py -3.8 -m pip install -U nuitka zstandard ordered-set
py -3.8 -m nuitka --standalone --mingw64 --enable-plugin=tk-inter --windows-disable-console --output-dir=build_nuitka --assume-yes-for-downloads invoice_gui_extractor.py
```

输出目录中的可执行文件可直接分发测试。

---

### Win7 打包验收清单

1. 在 Win7 干净环境（无 Python）直接启动程序。
2. 导入 1 份 PDF，导出 XLSX 成功。
3. 再测 2~3 份不同版式 PDF（含多页）。
4. 若失败，优先回退到：`Python 3.8 + PyInstaller 5.13.2 + onedir + Win7 本机打包`。

## 打包后功能不缺失建议

1. 打包与运行使用相同 Python 版本。
2. 打包前先确认本地可正常导出 `CSV/XLSX`。
3. 确认依赖齐全：`pdfplumber`、`pandas`、`openpyxl`。
4. 若被安全软件拦截，请将 `dist` 目录加入白名单后再验证。
5. 用 2~3 份真实样本（不同模板）回归测试：
   - 明细正常模板
   - 明细字段不完整模板
   - 多页 PDF 模板

## 常见问题

### 1) 明细识别不完整

- 原因：部分 PDF 为扫描件或表格结构特殊。
- 处理：
  - 扫描件先做 OCR 再导入；
  - 提供可复现样本后可继续补充模板规则。

### 2) XLSX 导出失败

安装依赖：

```bash
pip install pandas openpyxl
```

### 3) PDF 无法读取

安装依赖：

```bash
pip install pdfplumber
```

### 4) 报错 `no such group`

- 该问题通常由正则表达式使用了“非捕获分组 `(?:...)`”，但代码仍尝试读取 `group(1)` 引起。
- 当前版本已修复：匹配函数会在“无捕获组”时自动回退到完整匹配，避免此类崩溃。

# 发票助手（PDF 批量提取 + GUI 导出）

用于批量读取发票 PDF，并导出为 `CSV` / `XLSX` 的桌面工具。

## 功能

- 批量添加 PDF，支持调整处理顺序（上移/下移）
- 合并整理（一个总表）或分开整理（每个 PDF 一个文件）
- 自动提取销售方、购买方税号、发票号码、开票日期、货物明细等字段
- 兼容明细表头拆分或字段缺失的版式（会尽量保留“产品名称 / 金额 / 税额”）
- XLSX 导出支持表头样式与列宽优化

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

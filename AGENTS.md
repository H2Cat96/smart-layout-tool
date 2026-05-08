# 人文课后巩固排版工具

本项目用于将 Word 文档生成排版成品（PDF 或 IDML）。用户不熟悉终端，请用自然语言引导完成全流程。

## 使用流程

当用户要求排版、生成 PDF、生成 IDML，或提到"课后巩固"时，按以下步骤引导：

### 1. 环境检查

先确认本项目根目录位置（包含 `templates/` 的目录），然后静默检查依赖：

```bash
python3 -c "import reportlab, docx, lxml, numpy" 2>&1
```

如果报 ModuleNotFoundError，自动安装：

```bash
pip3 install -r requirements.txt
```

安装完成后告知用户"环境已就绪"，失败则展示错误并协助排查。

### 2. 确定 Word 文档

帮用户找到要排版的 docx 文件：

```bash
find ~/Desktop ~/Downloads -name "*.docx" -maxdepth 2 -mtime -30 2>/dev/null | head -20
```

把找到的文件列出来让用户确认。

### 3. 选择输出格式

问用户要哪种输出：
- **PDF** — 直接出成品，可打印
- **IDML** — InDesign 格式，可精调后再出印刷文件

### 4. 执行生成

PDF：
```bash
cd <项目根目录> && python3 templates/gonggu-neiye/legacy/generate_print_pdf.py --docx "<docx路径>" --output "<输出目录>"
```

IDML：
```bash
cd <项目根目录> && python3 templates/gonggu-neiye/legacy/generate_idml.py --docx "<docx路径>" --output "<输出目录>"
```

输出目录默认放在 docx 同级目录，以文件名命名。

### 5. 报告结果

告诉用户文件在哪里。IDML 模式额外说明：
- 用 InDesign 打开 `.idml` 文件
- `Links/` 里的图片需要重新链接
- `Document Fonts/` 是排版用字体

## 注意事项

- docx 文件路径包含空格时务必用引号包裹
- 当前模板仅支持"巩固内页"（gonggu-neiye）

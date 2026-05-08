---
description: 人文课后巩固排版工具——从 Word 生成 PDF 或 IDML
allowed-tools: Bash, Read, AskUserQuestion
argument-hint: [docx文件路径]
---

# 人文课后巩固排版生成

引导用户完成从 Word 文档到排版成品的全流程。用户不熟悉终端，用自然语言交流即可。

## 流程

### 1. 确定 Word 文档

如果 `$ARGUMENTS` 有值，直接作为 docx 路径。否则帮用户找到文件：

```bash
find ~/Desktop ~/Downloads -name "*.docx" -maxdepth 2 -mtime -30 2>/dev/null | head -20
```

把找到的文件列出来让用户确认。

### 2. 选择输出格式

问用户要哪种输出：
- **PDF** — 直接出成品，可打印
- **IDML** — InDesign 格式，可精调后再出印刷文件

### 3. 执行

工作目录：`/Users/tal/Desktop/teaching-layout-cli`

PDF：
```bash
cd /Users/tal/Desktop/teaching-layout-cli && python3 templates/gonggu-neiye/legacy/generate_print_pdf.py --docx "$DOCX" --output "$OUTPUT_DIR"
```

IDML：
```bash
cd /Users/tal/Desktop/teaching-layout-cli && python3 templates/gonggu-neiye/legacy/generate_idml.py --docx "$DOCX" --output "$OUTPUT_DIR"
```

输出目录默认放在 docx 同级目录，以文件名命名。

### 4. 结果

告诉用户文件在哪里。IDML 模式额外说明：
- 用 InDesign 打开 `.idml` 文件
- `Links/` 里的图片需要重新链接
- `Document Fonts/` 是排版用字体

## 依赖

如果报 ModuleNotFoundError，执行：
```bash
pip3 install reportlab python-docx lxml numpy
```

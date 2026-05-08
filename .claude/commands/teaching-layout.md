---
description: 智能排版工具——从 Word 生成 PDF 或 IDML
allowed-tools: Bash, Read, AskUserQuestion
argument-hint: [docx文件路径]
---

# 智能排版工具

引导用户完成从 Word 文档到排版成品的全流程。用户不熟悉终端，用自然语言交流即可。

## 流程

### 0. 环境检查

先确认项目目录位置。用 `find / -type d -name "teaching-layout-cli" 2>/dev/null | head -5` 找到项目根目录，后续命令都基于该路径。

然后静默检查 Python 依赖是否就绪：

```bash
python3 -c "import reportlab, docx, lxml, numpy" 2>&1
```

如果报 ModuleNotFoundError，自动安装：

```bash
pip3 install -r <项目根目录>/requirements.txt
```

安装完成后告知用户"环境已就绪"，失败则展示错误信息并协助排查。

### 1. 选择模板

列出可用模板：

```bash
ls <项目根目录>/templates/
```

然后读取每个模板的 `template-package.json` 获取显示名称，列出让用户选：

```
可用模板：
1. 人文-课后巩固（gonggu-neiye）
2. 数学-课后内页（math-neiye）
...
```

如果只有一个模板，直接使用，不需要问。

### 2. 确定 Word 文档

如果 `$ARGUMENTS` 有值，直接作为 docx 路径。否则帮用户找到文件：

```bash
find ~/Desktop ~/Downloads -name "*.docx" -maxdepth 2 -mtime -30 2>/dev/null | head -20
```

把找到的文件列出来让用户确认。

### 3. 选择输出格式

问用户要哪种输出：
- **PDF** — 直接出成品，可打印
- **IDML** — InDesign 格式，可精调后再出印刷文件

注意：只有模板目录下存在 `legacy/generate_idml.py` 时才提供 IDML 选项。

### 4. 执行

工作目录为本项目根目录。根据选定的模板和格式执行：

PDF：
```bash
cd <项目根目录> && python3 templates/<模板id>/legacy/generate_print_pdf.py --docx "$DOCX" --output "$OUTPUT_DIR"
```

IDML：
```bash
cd <项目根目录> && python3 templates/<模板id>/legacy/generate_idml.py --docx "$DOCX" --output "$OUTPUT_DIR"
```

输出目录默认放在 docx 同级目录，以文件名命名。

### 5. 结果

告诉用户文件在哪里。IDML 模式额外说明：
- 用 InDesign 打开 `.idml` 文件
- `Links/` 里的图片需要重新链接
- `Document Fonts/` 是排版用字体

## 注意事项

- 如果用户的 docx 文件路径包含空格，务必用引号包裹

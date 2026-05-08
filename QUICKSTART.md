# 智能排版工具 · 快速开始

当前内置模板是 `人文-课后巩固`，内部目录为 `templates/gonggu-neiye`。

## 1. 安装依赖

在项目根目录运行：

```bash
python3 -m pip install -r requirements.txt
```

## 2. 检查模板包

```bash
python3 -m teaching_layout inspect-template --template "人文-课后巩固"
```

如果 `status` 是 `portable`，说明模板运行所需的 JSON、SVG 和字体都在包内，且没有发现本机绝对路径。

## 3. 生成 PDF

```bash
python3 -m teaching_layout build \
  --template "人文-课后巩固" \
  --docx "/path/to/input.docx" \
  --out "/path/to/output-folder" \
  --pdf-name "output.pdf" \
  --preview-name "output_预览.png" \
  --title "课后巩固"
```

生成结果会包含 PDF、预览图和 `generation-manifest.json`。

默认输出是 IDML 单页尺寸 PDF，并使用程序绘制 CMYK 颜色。需要保留旧的横向对页输出时，在 `build` 命令里加：

```bash
--page-mode spread
```

需要旧 RGB 颜色时，加：

```bash
--color-mode rgb
```

## 4. 校验 PDF

```bash
python3 -m teaching_layout validate \
  --pdf "/path/to/output-folder/output.pdf" \
  --manifest "/path/to/output-folder/generation-manifest.json"
```

## AI 助手用户（Claude Code / Codex / OpenClaw）

用 AI 助手打开本项目文件夹，输入：

```
/teaching-layout
```

AI 会自动引导你完成：
1. 选择 Word 文档
2. 选择输出格式（PDF 或 IDML）
3. 执行生成
4. 报告结果和输出路径

无需手动敲命令，全程自然语言交互。

**IDML 输出说明：** 生成的包含 `.idml` 文件、`Links/`（图片）和 `Document Fonts/`（字体）。在 InDesign 中打开 IDML 后需要重新链接图片。

## 分享注意

- 当前模板包随带字体，方便内部流转。对外分发前请确认字体和参考材料授权。
- 知音楼/Yach 文档需要先导出为 DOCX，再交给 CLI 生成。
- 新模板不要复用 `人文-课后巩固` 的规则当作通用标准；每个模板应有独立模板包。

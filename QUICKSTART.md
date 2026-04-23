# Quickstart

这是一份可分享的本地教辅排版 CLI。当前内置模板是 `人文-课后巩固`，内部目录为 `templates/gonggu-neiye`。

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

## 4. 校验 PDF

```bash
python3 -m teaching_layout validate \
  --pdf "/path/to/output-folder/output.pdf" \
  --manifest "/path/to/output-folder/generation-manifest.json"
```

## 分享注意

- 当前模板包随带字体，方便内部流转。对外分发前请确认字体和参考材料授权。
- 知音楼/Yach 文档需要先导出为 DOCX，再交给 CLI 生成。
- 新模板不要复用 `人文-课后巩固` 的规则当作通用标准；每个模板应有独立模板包。

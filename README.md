# teaching-layout-cli

本地教辅排版 CLI 引擎（PDF + IDML）。

目标是把已经验证过的“模板材料”生成流程收敛成稳定入口，并让模板以独立模板包的形式分享：

- `build`：使用模板包和 Word 生成印刷 PDF。
- `parse-docx`：统计 Word 内容结构，确认答题横线没有丢。
- `validate`：检查 PDF 页数、字体和行首标点。
- `import-template`：把用户提供的原始模板素材导入成模板包骨架。
- `inspect-template`：检查模板包是否缺文件、缺字体或仍含本机绝对路径。

快速交给别人使用时，先看 [QUICKSTART.md](QUICKSTART.md)。

## AI 助手用户（推荐）

用 Claude Code / Codex / OpenClaw 打开本项目文件夹，输入：

```
/teaching-layout
```

AI 会引导你选择输出格式（PDF / IDML）、确认 Word 文档、执行生成并报告结果，全程无需手动敲命令。

## 命令行使用

```bash
cd teaching-layout-cli

# 解析 Word 结构
python3 -m teaching_layout parse-docx \
  --docx "/path/to/input.docx"

# 生成 PDF
python3 -m teaching_layout build \
  --template "人文-课后巩固" \
  --docx "/path/to/input.docx" \
  --out "/path/to/output-folder" \
  --pdf-name "output.pdf" \
  --preview-name "output_预览.png" \
  --title "课后巩固"

# 生成 IDML（InDesign 格式）
python3 templates/gonggu-neiye/legacy/generate_idml.py \
  --docx "/path/to/input.docx" \
  --output "/path/to/output-folder"

# 校验 PDF
python3 -m teaching_layout validate \
  --pdf "/path/to/output-folder/output.pdf" \
  --manifest "/path/to/output-folder/generation-manifest.json"
```

`build` 默认输出 IDML 单页尺寸 PDF，并使用程序绘制 CMYK 颜色。需要保留旧的横向对页调试输出时，加 `--page-mode spread`；需要旧 RGB 颜色时，加 `--color-mode rgb`。

分享前可以检查模板包：

```bash
python3 -m teaching_layout inspect-template --template "人文-课后巩固"
```

导入新模板素材时，用户只需要准备这样的原始材料目录：

```text
某个模板材料/
  assets/
    svg/
    images/
    fonts/
    references/
```

然后运行：

```bash
python3 -m teaching_layout import-template \
  --id new-template \
  --name "新模板" \
  --source "/path/to/某个模板材料"
```

`import-template` 会复制这些素材，并生成 `template-package.json`、`layout-rules.json`、`asset-map.json`、`extracted/font-map.json`、`extracted/template.json`、`extracted/style-map.json` 和 `extracted/geometry-map.json`。如果 `assets/references/` 里没有 PDF 或 IDML，导入结果会标记为 `needs_layout_reference`。

如果使用系统 Python，需要先安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

安装依赖后可以直接用系统 Python 运行：

```bash
python3 -m teaching_layout parse-docx \
  --docx "/path/to/input.docx"
```

`--template` 支持两种写法：

- 模板名：`人文-课后巩固` 或 `【人文-课后巩固】`，从 `templates/gonggu-neiye/` 读取模板包配置。
- 兼容旧内部 id：`gonggu-neiye`。
- 旧式材料路径，兼容早期直接指向模板材料目录的用法。

## 第一版验证结果

当前“人文-课后巩固”模板包的校验目标：

- PDF 页数：默认按 IDML 单页输出；旧对页模式可用 `--page-mode spread` 保留
- 练习段落：32/32
- 答案段落：13/13
- Word 答题横线：6 条可识别
- 行首标点：0
- 嵌入字体：方正楷体、方正颜宋准/中/粗

## 当前模板包约定

现在推荐每个模板一个独立模板包：

```text
templates/
  gonggu-neiye/                # 内部目录 id，用户可见名称是“人文-课后巩固”
    template-package.json
    layout-rules.json
    asset-map.json
    legacy/generate_print_pdf.py
    extracted/template.json
    extracted/font-map.json
    extracted/style-map.json
    extracted/geometry-map.json
    assets/svg/
    assets/images/
    assets/fonts/
    assets/references/
```

`template-package.json` 负责声明这个模板实例的源材料位置和渲染后端；`layout-rules.json`、`asset-map.json` 放该模板自己的差异化规则。不要把某个模板的字体、颜色、题号位置当成全局默认。

当前 `gonggu-neiye` 模板包的用户可见名称是“人文-课后巩固”，已经把旧生成脚本、抽取出的版面 JSON、SVG、参考 PDF/IDML 和字体资产放进仓库目录；`extracted/font-map.json` 使用相对路径指向 `assets/fonts/`，便于内部分享。

后续版本会继续把 `generate_print_pdf.py` 拆成真正的模板无关模块，让模板包配置直接驱动核心引擎。

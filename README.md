# teaching-layout-cli

本地教辅 PDF 排版 CLI 引擎雏形。

第一版目标是把已经验证过的“模板材料”生成流程收敛成稳定入口：

- `build`：使用模板包和 Word 生成印刷 PDF。
- `parse-docx`：统计 Word 内容结构，确认答题横线没有丢。
- `validate`：检查 PDF 页数、字体和行首标点。
- `import-template`：把用户提供的原始模板素材导入成模板包骨架。

## 使用

```bash
cd /Users/tal/Desktop/teaching-layout-cli

python3 -m teaching_layout parse-docx \
  --docx "/Users/tal/Desktop/模板材料/【改1】贾平凹短篇：标题含义理解-课后题.docx"

python3 -m teaching_layout build \
  --template gonggu-neiye \
  --docx "/Users/tal/Desktop/模板材料/【改1】贾平凹短篇：标题含义理解-课后题.docx" \
  --out "/Users/tal/Desktop/模板材料/正式生成输出" \
  --pdf-name "贾平凹标题含义_SVG白底样式版.pdf" \
  --preview-name "贾平凹标题含义_SVG白底样式版_预览.png" \
  --title "贾平凹标题含义理解"

python3 -m teaching_layout validate \
  --pdf "/Users/tal/Desktop/模板材料/正式生成输出/贾平凹标题含义_SVG白底样式版.pdf" \
  --manifest "/Users/tal/Desktop/模板材料/正式生成输出/generation-manifest.json"
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
  --source "/Users/tal/Desktop/某个模板材料"
```

`import-template` 会复制这些素材，并生成 `template-package.json`、`layout-rules.json`、`asset-map.json`、`extracted/font-map.json`、`extracted/template.json`、`extracted/style-map.json` 和 `extracted/geometry-map.json`。如果 `assets/references/` 里没有 PDF 或 IDML，导入结果会标记为 `needs_layout_reference`。

如果使用系统 Python，需要先安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

本机 Codex bundled Python 已带齐当前流程依赖，可直接这样运行：

```bash
/Users/tal/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m teaching_layout parse-docx \
  --docx "/Users/tal/Desktop/模板材料/【改1】贾平凹短篇：标题含义理解-课后题.docx"
```

`--template` 支持两种写法：

- 模板名：`gonggu-neiye`，从 `templates/gonggu-neiye/` 读取模板包配置。
- 旧式材料路径：`/Users/tal/Desktop/模板材料`，兼容早期直接指向模板材料目录的用法。

## 第一版验证结果

使用当前“巩固内页”模板包生成的 CLI 测试产物：

```text
/Users/tal/Desktop/模板材料/CLI生成输出/贾平凹标题含义_CLI测试版.pdf
/Users/tal/Desktop/模板材料/CLI生成输出/贾平凹标题含义_CLI测试版_预览.png
```

校验结果：

- PDF 页数：3 个双页展开
- 练习段落：32/32
- 答案段落：13/13
- Word 答题横线：6 条可识别
- 行首标点：0
- 嵌入字体：方正楷体、方正颜宋准/中/粗

## 当前模板包约定

现在推荐每个模板一个独立模板包：

```text
templates/
  gonggu-neiye/
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

当前 `gonggu-neiye` 模板包已经把旧生成脚本、抽取出的版面 JSON 和 SVG 资产放进仓库目录；字体文件仍由 `extracted/font-map.json` 指向本机模板材料中的字体文件，避免仓库直接纳入大体积字体包。

后续版本会继续把 `generate_print_pdf.py` 拆成真正的模板无关模块，让模板包配置直接驱动核心引擎。

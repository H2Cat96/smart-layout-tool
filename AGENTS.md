# 智能排版工具

本项目用于将 Word 文档生成排版成品（PDF 或 IDML）。用户不熟悉终端，请用自然语言引导完成全流程。

## 使用流程

当用户要求排版、生成 PDF、生成 IDML 时，按以下步骤引导：

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

### 2. 选择模板

列出可用模板：

```bash
ls <项目根目录>/templates/
```

读取每个模板目录下的 `template-package.json` 获取显示名称，列出让用户选。
如果只有一个模板，直接使用不需要问。

### 3. 确定 Word 文档

帮用户找到要排版的 docx 文件：

```bash
find ~/Desktop ~/Downloads -name "*.docx" -maxdepth 2 -mtime -30 2>/dev/null | head -20
```

把找到的文件列出来让用户确认。

### 4. 选择输出格式

问用户要哪种输出：
- **PDF** — 直接出成品，可打印
- **IDML** — InDesign 格式，可精调后再出印刷文件

只有模板目录下存在 `legacy/generate_idml.py` 时才提供 IDML 选项。

### 5. 执行生成

PDF：
```bash
cd <项目根目录> && python3 templates/<模板id>/legacy/generate_print_pdf.py --docx "<docx路径>" --output "<输出目录>"
```

IDML：
```bash
cd <项目根目录> && python3 templates/<模板id>/legacy/generate_idml.py --docx "<docx路径>" --output "<输出目录>"
```

输出目录默认放在 docx 同级目录，以文件名命名。

### 6. 报告结果

告诉用户文件在哪里。IDML 模式额外说明：
- 用 InDesign 打开 `.idml` 文件
- `Links/` 里的图片需要重新链接
- `Document Fonts/` 是排版用字体

## 注意事项

- docx 文件路径包含空格时务必用引号包裹

---

## 新建模板

当用户要求创建新模板、导入模板、或提到"做一个新模版"时，按以下步骤引导：

### 1. 确认用户已准备的素材

新模板需要以下材料，问用户已准备好哪些：

| 素材 | 说明 | 必需 |
|------|------|------|
| IDML 或 PDF 参考 | 已排好的样例文件，用于提取版面结构 | 至少一个 |
| 字体文件 | 模板用到的所有 .ttf/.otf 字体 | ✓ |
| SVG 装饰 | 标题角标、栏目底纹等装饰元素 | 可选 |

让用户把这些素材放到一个文件夹里，结构如下：

```
模板素材/
  assets/
    references/   ← IDML 和 PDF 放这里
    fonts/        ← 字体放这里
    svg/          ← SVG 装饰放这里
    images/       ← 其他图片（可选）
```

### 2. 运行导入

确认素材目录路径后执行：

```bash
cd <项目根目录> && python3 -m teaching_layout import-template \
  --id "<模板id>" \
  --name "<模板显示名>" \
  --source "<素材目录路径>"
```

- `--id`：英文标识，如 `yuedu-neiye`
- `--name`：中文显示名，如 `"阅读-课后内页"`

导入完成后告知用户生成了哪些文件。

### 3. 提取版面结构

导入只是复制素材和生成骨架文件。接下来需要从参考文件中提取版面结构，填充到 `extracted/template.json`。

**有 IDML 时：** 解压 IDML（zip 格式），分析 designmap.xml、Spreads/*.xml、Stories/*.xml，提取页面尺寸、TextFrame 布局、Story ID、段落样式。

**只有 PDF 时：** 用 PyMuPDF 分析 PDF：

```python
import fitz
doc = fitz.open("参考.pdf")
page = doc[0]
# 页面尺寸、文字块 bbox（推算版心）、字体字号、图片位置
```

从 PDF 多页中提取：页面尺寸、版心边距、文字区域位置（→ flow frame）、字体/字号、配色。
如果缺少 fitz：`pip3 install PyMuPDF`

告知用户："骨架已生成，接下来需要我分析参考文件来提取版面结构，要继续吗？"

### 4. 配置 layout-rules.json

根据提取到的样式信息，帮用户填写：
- 字体/字号/行距配置
- 颜色定义
- 内容检测正则（标题、题号等模式）
- 如果有 IDML：还需配置段落样式映射和 IDML 字体名

可参考现有模板 `templates/gonggu-neiye/layout-rules.json` 的结构。

### 5. 生成文档撰写指南

分析模板实际用到的内容结构，基于项目根目录 `文档撰写规范.md` 通用底本，生成该模板专属的文档撰写指南：

- 只保留该模板实际用到的内容类型
- 补充模板特有的结构和规则
- 附一份简短的示例文档片段

输出到：`templates/<模板id>/文档撰写指南.md`

### 6. 验证

```bash
cd <项目根目录> && python3 -m teaching_layout inspect-template --template "<模板显示名>"
```

检查 status 是否为 `portable`，如果有缺失项目帮用户补齐。

---
description: 引导创建新排版模板
allowed-tools: Bash, Read, AskUserQuestion
argument-hint: [素材目录路径]
---

# 新建排版模板

引导用户从零创建一个新的排版模板包。

## 流程

### 1. 确认素材

问用户已准备好哪些材料：

- **IDML 或 PDF 参考文件**（至少一个）— 已排好的样例，用于提取版面结构
- **字体文件**（必需）— 模板用到的 .ttf/.otf
- **SVG 装饰**（可选）— 标题角标、栏目底纹

素材应放在一个目录下：

```
素材目录/
  assets/
    references/   ← IDML 或 PDF
    fonts/        ← 字体
    svg/          ← 装饰
    images/       ← 其他图片
```

如果用户还没整理好，帮他把散落的文件归到这个结构里。

### 2. 确定模板标识

问用户两个信息：
- **模板 ID**：英文标识（如 `yuedu-neiye`），只能用字母、数字、横杠、下划线
- **显示名称**：中文名（如 `"阅读-课后内页"`）

### 3. 运行导入

```bash
cd <项目根目录> && python3 -m teaching_layout import-template \
  --id "$TEMPLATE_ID" \
  --name "$TEMPLATE_NAME" \
  --source "$SOURCE_DIR"
```

展示导入结果（生成了哪些文件、status）。

### 4. 提取版面结构

告知用户骨架已生成，询问是否继续分析参考文件来提取版面结构。

#### 路径 A：有 IDML

解压 IDML（它是 zip），分析：
- designmap.xml → 页面尺寸、Spread 列表
- Spreads/*.xml → TextFrame 位置和串联关系
- Stories/*.xml → 段落样式名称

将结果填入 `extracted/template.json`。

#### 路径 B：只有 PDF（无 IDML）

用 Python 分析 PDF 提取版面信息：

```python
import fitz  # PyMuPDF

doc = fitz.open("参考文件.pdf")
page = doc[0]

# 1. 页面尺寸
width_pt, height_pt = page.rect.width, page.rect.height

# 2. 文字区域边界（所有文字块的 bbox 并集 → 推算版心）
blocks = page.get_text("dict")["blocks"]
text_blocks = [b for b in blocks if b["type"] == 0]

# 3. 字体和字号
for b in text_blocks:
    for line in b["lines"]:
        for span in line["spans"]:
            font = span["font"]
            size = span["size"]

# 4. 图片区域
image_blocks = [b for b in blocks if b["type"] == 1]
```

从多页 PDF 中提取：
- 页面尺寸（宽 × 高，pt）
- 版心区域（上下左右边距）
- 文字区域数量和位置（推算 TextFrame 布局）
- 使用的字体和字号
- 图片位置和尺寸

然后根据分析结果构建 `extracted/template.json`：
- 每个排版区域对应一个 flow frame
- 用 PDF 中实际测量的 bbox 定义 frame 位置
- 练习页和答案页分别定义 spread 模板

如果缺少 `fitz`，先安装：`pip3 install PyMuPDF`

### 5. 配置 layout-rules.json

参考 `templates/gonggu-neiye/layout-rules.json` 的结构，帮用户填写：
- `styles`：各角色的字体/字号/行距（从 PDF 分析中提取的字体信息作为参考）
- `colors`：配色方案（从 PDF 中提取的文字颜色）
- `content_detection`：标题、题号等正则模式

如果是 IDML 路径，额外配置：
- `idml.paragraph_style_mapping`：内容类型 → InDesign 样式
- `idml.font_names`：字体标识 → IDML 内字体名

### 6. 生成文档撰写指南

分析模板实际用到的内容结构（从参考文件中识别到的层级、题型、特殊格式），自动生成一份该模板专属的文档撰写指南。

基于项目根目录的 `文档撰写规范.md` 作为通用底本，只保留该模板实际用到的内容类型，并补充该模板特有的规则。

输出到模板目录：`templates/$TEMPLATE_ID/文档撰写指南.md`

内容包含：
- 该模板的文档整体结构示意（用到了哪些层级）
- 各内容类型的写法说明（只列该模板实际用到的）
- 内联格式支持列表
- 一份简短的示例文档片段

### 7. 验证

```bash
cd <项目根目录> && python3 -m teaching_layout inspect-template --template "$TEMPLATE_NAME"
```

确认 status 为 `portable`，列出任何缺失项并协助补齐。

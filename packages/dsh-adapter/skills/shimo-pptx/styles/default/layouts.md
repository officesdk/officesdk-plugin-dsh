# 版式说明：default（蓝色品牌风格）

数值（颜色、字体、字号）见同目录 `style.json`；用法见 `../../references/style.md`。

## 实时方式（template.pptx）

页面 720 × 405 pt（16:9），坐标单位都是 pt，从左上角算。`presentation_create` 返回的 `layouts` 按占位符区分版式（编辑器转换时会丢掉版式名称）：

| 占位符 | 版式 | 用途 |
|---|---|---|
| `title` + `subtitle` | 封面 | 主色底，右上副标题（必填），中部细横线，左下两行大标题（框底对齐） |
| `title` + `body` | 标题和内容 | 白底，顶部标题，下面正文 |
| `title` | 仅标题 | 白底，顶部标题，其余自己放文本框、形状、图片 |
| `body` | 结束页 | 主色底，细横线，左下大字 THANKS（与封面标题同一框） |
| 无 | 空白 | 白底空页；`photo-statement` 先把本页背景设成主色 |

新建时用 `slideLayouts` 指定各页角色。创建返回的每一页已带有所选模板的占位符，用形状 ID 填写。下面的版式和位置也在 `presentation_create` 返回的 `style.layouts` 里（`layouts` 是各版式占位符的位置，`slots` 是常用的文字区和图片框，单位 pt），两处出自同一份 `layouts.json`。常用位置：

| 元素 | left, top, width, height |
|---|---|
| 内容页标题（版式自带，在顶部） | 36, 25, 648, 65 |
| 标签色块（主色小块 + 反白字） | 36, 100, 宽按字数, 26 |
| 正文区 | 36, 140, 648, 230 |
| 左右两栏 | 36, 140, 312, 230 / 372, 140, 312, 230 |
| 三栏 | 36, 140, 192, 230 / 264, 140, 192, 230 / 492, 140, 192, 230 |

- 模板主题就是这套颜色和字体：文本框默认是主色字；主色底上的字设成 `onPrimary`。
- `addShape` 用于标签色块（`Rectangle`，主色填充，`line: "none"`，反白字）、分隔线（细长 `Rectangle`，高 1）、编号圆点（`Ellipse`）。
- 字号用 `style.sizesPt`：页标题 `page_title`，正文 `body`，标签 `label`，指标数字 `metric`。

图片位置（`addImage` 按框拉伸，先用 `find_images.py fit --box <宽>x<高>` 裁成同样比例；主体不在框中时加 `--anchor top|bottom|left|right` 按同一 `--out` 重裁；生成配图用 `--size` 出同比例的图）：

| 用法 | 版式 | left, top, width, height | 生成配图 `--size` | 说明 |
|---|---|---|---|---|
| 全幅图文 | 空白 | 0, 50.4, 720, 354.6 | 1920x945 | 城市/风景横图用这个。本页 `setBackground` 主色；`fit --box 720x354.6 --darken 0.3`；标题 43.2, 72.0, 496.8, 108.0 两行反白；说明 518.4, 201.6, 180.0, 136.8 一段一个框 |
| 左图右文 | 仅标题 | 36, 100, 312, 270 | 1248x1080 | 只给竖构图。右栏文字放 372, 100, 312, 270 |
| 横幅图 | 仅标题 | 36, 100, 648, 180 | 1920x533 | 下方 36, 296, 648, 80 放一行说明 |
| 全幅背景图 | 空白 | 0, 0, 720, 405 | 1920x1080 | 先 `fit --darken 0.45`，再放白字标题（36, 250, 648, 80） |

封面 `placeholders` 必须同时给 title 和 subtitle：title 两行、每行 2–6 个汉字，垂直落在标题框下沿；subtitle 一句定位，不能空。结束页只写 THANKS，位置与封面标题同一框。封面和结束页不放图；要照片开场就在封面后加一页全幅图文。城市/风景横图不要塞进左图右文。

## 离线方式（scripts/brand_deck.py）

deck JSON 为 UTF-8 JSON，完整样例见同目录 `example.json`。
路径可写绝对路径，或相对 deck JSON 所在目录。

```json
{
  "title": "公司名称\n业务介绍",
  "subtitle": "产品与业务概览",
  "palette": "color-1",
  "year": 2026,
  "brand": {
    "website": "company.example",
    "logo": "assets/logo-blue.png",
    "logo_on_dark": "assets/logo-white.png",
    "logo_missing_reason": "没有 Logo 文件时写原因"
  },
  "slides": [ { "layout": "cover" }, "…", { "layout": "closing" } ]
}
```

- `year` 默认当前年份，是封面右下的制作年份；数据年份写在正文、图表或备注里，不跟它走。
- `logo` 用于白底页，`logo_on_dark` 用于主色底和照片上；只给一个时两处都用它，注意对比度。
- 每页可加 `notes`（备注）和 `sources`（出处数组，写进备注）。

### 各页型字段

文字里的 `\n` 是手动换行。`?` 表示可选。

| layout | 字段 |
|---|---|
| `cover` | `title?`（默认用顶层 title）、`subtitle?`、`year?` |
| `overview` | `title`（两行，每行一个高亮色块）、`lead`、`intro`、`tag?`（默认"阅读重点"）、`items`（≤5 条短句） |
| `statement` | `title`、`items`：≤2 个 `{heading, text}`、`side`：`{heading, items[]}`、`image?`（左竖图）、`image2?`、`image3?`（标题右侧方图与横图） |
| `columns` | `title`、`subtitle`、`columns`：2–4 个 `{label, text}`（label 是短标签，2–6 字） |
| `composition` | `title`、`text`、`note?`、`chart`：`{title, categories[], values[]}`（真实可加总的组成）、`panel`：`{heading, metrics: [{label, value}]×≤3}` |
| `image-statement` | `image`、`title`（两行：第二行带主色高亮）、`text` |
| `numbered` | `lead`、`tag`、`title`、`image?`、`items`（≤5 条，可两行） |
| `metrics` | `title`（≤2 行高亮）、`text`、`metrics`：2 个 `{value, label}`、`panel`：3 个 `{value, label}`、`image?`（背景图，自动加 55% 黑色蒙版） |
| `chart` | `chart`：`{type: column|bar|line, unit, categories[], series: [{name, values[]}], number_format?}`、`title`、`text`、`tag?`、`note?` |
| `risks` | `title`（两行高亮）、`lead`、`intro`、`tag`、`items[]`、`image?` |
| `table` | `title`、`subtitle?`、`table`：`{columns[], rows[[...]], column_widths?[], row_labels?}`（表头和首列默认主色底反白，`row_labels: false` 只强调表头）、`note?`；超过 8 行请拆页 |
| `closing` | `title?`（默认 THANKS）、`subtitle?`、`year?` |

图表（`chart`、`composition`）最多 6 个系列或类别：超过 3 个会警告并用主色浅色区分，超过 6 个报错。各页型条目数有上限（如 `overview` 5 项、`columns` 4 栏），超出部分会被截掉并在输出里警告。

### 选页型

| 内容 | 页型 |
|---|---|
| 目录、要点概览 | `overview` |
| 产品/业务体系、有照片的主题 | `statement`、`image-statement` |
| 3–4 个并列项 | `columns` |
| 收入/成本组成（有真实占比） | `composition` |
| 时间趋势、类别比较 | `chart`（趋势用 line/column，类别多用 bar） |
| 规模、关键数字 | `metrics` |
| 计划、重点、步骤（≤5 项） | `numbered` |
| 多方案、多产品对比 | `table` |
| 风险、问题、边界 | `risks` |

### 来源

页型与几何取自 brand-ppt-12 v1.4.1 的 13 页基线（`master-19.json` 第 1、2、4、6–13 页）与扩展第 16 页（对比表）；坐标沿用 960×540 画布（1px = 0.75pt）。未收录原第 3、5、14、15、17–19 页（多照片拼版），因为 DSH 默认没有找图能力，用户不提供照片时这些页型无法成立。

### 已知限制

文字宽度按字符宽度估算，没有逐字实测字形；改字号或文案后要看渲染。复杂图（桑基、旭日等）不支持；常规柱、条、线、环图为原生图表。

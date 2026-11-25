# 样式包格式（shimo-pptx `styles/<id>/`）

写给做蒸馏的开发者。流程见 [README.md](README.md)；校验规则以 `validate.mjs` 为准，本文改了要同步改它。

## 1. 目录

```
packages/dsh-adapter/skills/shimo-pptx/styles/<id>/
  style.json      配色、字体、字号、实时模板参数、meta（必需）
  layouts.json    版式描述：离线每页的对象，实时版式与槽位（必需）
  template.pptx   实时方式模板，scripts/build_template.py 由 style.json 生成（必需）
  layouts.md      给模型看的版式说明（必需；实时坐标要与 layouts.json 一致）
  example.json    示例 deck，回归、缩略图用（必需）
  thumb.png       缩略图 480×270（必需；tooling/style-distill/thumb.py 生成）
  preview-<n>.png 页面样张 480×270（可选，2–4 张，meta.previews 列出；thumb.py --previews 生成）
```

- `<id>` 只能用小写字母、数字和连字符：`^[a-z0-9]+(?:-[a-z0-9]+)*$`。Host 和 `brand_deck.py` 只把 `style.json`、`template.pptx`、`layouts.json` 三者齐全的目录当成样式包（插件选择器也只列这些）；`build_template.py` 只要求有 `style.json`。
- 整个 `skills/` 随插件发给客户：包里不能放源包原件、参考截图、不能分发的字体。`tooling/` 不进插件包。
- **暂存**：`meta.design_source.cleared` 还是 false 的包放 `tooling/style-distill/staging/<id>/`（格式与上面相同），各脚本用 `--styles-dir tooling/style-distill/staging`、校验用 `--styles tooling/style-distill/staging`。用户或法务确认后把 `cleared` 改 true、整目录挪进 `styles/`，就算登记。测试包可以带上暂存包（`TEST_PACKAGE_STAGING_STYLES`，见 README），正式包不会带。

## 2. `style.json`

| 字段 | 说明 |
|---|---|
| `name` | 等于目录名 |
| `title` | 中文标题（日志、报错里用） |
| `meta` | 见 §2.1 |
| `canvas` | `width_px` 960、`height_px` 540、`slide_width_in` 10、`slide_height_in` 5.625、`emu_per_px` 9525、`pt_per_px` 0.75 |
| `default_palette` | `palettes` 里的键 |
| `palettes.<id>` | `name` 加颜色角色：`primary ink onPrimary paper secondary onSecondary inkOnSecondary chartSecondary chartTertiary lineOnPaper lineOnPrimary`，可选 `dividerOnPrimary` 等；版式里按角色名引用 |
| `fonts` | `cjk`、`latin`、`latin_light`、`rule`（一行有中文就整行用 `cjk`）；可选 `roles: {<角色>: {latin, cjk, latin_light?}}`，`text` 对象写 `font: <角色>` 时用（衬线、等宽这类第二、第三套字体）；用到的字体都要列进 `meta.fonts` |
| `type_scale_px` | 离线字号层级（文档用；版式里写实际数值） |
| `template` | `surface`、`sizes_pt`（实时方式字号，`presentation_create` 返回给模型）；可选 `colors: {cover_bg, closing_bg, content_bg, cover_text, title, body}`（配色角色名，缺省：封面/结束页主色底、`onPrimary` 字，内容页标题主色）和 `dark: true`（母版按深色主题映射：文字 lt1、背景 dk1）；可选 `cover: "brand"`（品牌封面/结束页：副标题右上、2.3in 处 1pt `lineOnPrimary` 细横线、标题左下锚底，结束页文字同框；不写时封面标题左上、副标题在下、结束页居中），boxes 须按对应几何写。深色风格（主色是黑色）必须写，否则会出黑底黑字 |
| `mask_opacity` | 版式里用 `style['mask_opacity'][...]` 读 |
| `source` | 提炼来源说明 |

### 2.1 `meta`

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `id` | string | 是 | 等于目录名和 `name` |
| `name` | `{zh, en}` | 是 | 显示名（选择器用） |
| `aliases` | string[] | 是 | 用户可能怎么叫它；所有包之间不重复，也不能等于别的包的 `id` |
| `summary` | `{zh, en}` | 是 | 一句话说明（选择器用） |
| `scenarios` | `{zh: string[1–6], en: string[1–6]}` | 是 | 适用场景 |
| `requires.photos` | `{slots, required, per_slide_max}` | 是 | 由校验按 layouts.json 计算比对：所有图片槽、`needed: true` 的图片槽、单页最多图片槽 |
| `requires.note` | `{zh, en}` | 是 | 前提条件的话（照片多少、找不到图怎么办） |
| `thumb` | string | 是 | 包内文件名（不含 `/`、`..`），PNG，宽 ≤ 640，16:9（±2%），≤ 150 KB |
| `previews` | string[] | 否 | 2–4 张页面样张的包内文件名，规则同 `thumb`；插件模板画廊在缩略图下方显示。`thumb.py --previews 3,4,8` 生成 `preview-<页>.png` 并打印这一项（不改写 style.json）。挑缩略图里没有的页 |
| `fonts[]` | `{family, role, license, source, bundled}` | 是 | 每个用到的字体；`license` 不在 `OFL-1.1 / Apache-2.0 / UFL-1.0 / MIT` 时 `bundled` 必须为 false |
| `design_source` | `{from, note, cleared}` | 是 | 设计出处；`cleared` 为 true 表示用户或法务已确认可以给客户，false 时校验不通过 |

字段只加不改名。插件选择器、模板画廊和推荐卡片读 `id / name / aliases / summary / scenarios / requires.note / thumb / previews`。

## 3. `layouts.json`

```json
{ "version": 1,
  "deck":    { "first": "cover", "last": "closing", "warnings": { "first": "…", "last": "…" } },
  "offline": { "canvas": {"width": 960, "height": 540, "unit": "px"}, "order": ["cover", "…"], "layouts": { "<id>": Layout } },
  "live":    { "canvas": {"width": 720, "height": 405, "unit": "pt"}, "layouts": [LiveLayout], "slots": [LiveSlot] } }
```

### 3.1 离线 Layout

```json
{ "title": "封面", "description": "brand_deck.py layouts 打印的一句话（含字段）",
  "background": "primary",
  "vars":   { "closing": false },
  "fields": { "title": {"type": "text", "required": true}, "items": {"type": "list", "max": 5}, "image": {"type": "image", "needed": true} },
  "let":    { "size": "fit(d['title'], 66, 640)" },
  "objects": [ … ] }
```

- `fields`：`type` 是 `text | list | object | image | chart | table`。`required` 的字段缺失时报错（`第N页(layout)：缺少字段 x`）；`needed` 只用于图片，表示这页应当有图（计入 `meta.requires.photos.required`，缺了是警告不是错误）；`max` 写列表上限（与 repeat 的 `max` 一致）。
- `extends: "<id>"`：变体，继承对象表，`vars`、`let` 合并，其余键子版式优先（例：closing 继承 cover，`vars.closing = true`）。
- 对象按顺序绘制（z 顺序）。任何对象可带 `when`（为假不画）和 `let`（局部名字）。

| type | 参数（带 * 的可写表达式） | 说明 |
|---|---|---|
| `text` | `text`*（必为表达式）`x* y* w* size* color weight* lh* align wrap max_h* anchor name`，可选 `font`（`fonts.roles` 的角色）`italic`（中文行不加）`bold` `spacing*`（字距 pt）`para_colors`（颜色数组：第 N 段用第 N 个，其余用 `color`；如白色标题加一行金色强调，仍是一个文本框） | 一个文本框；字符串里 `\n` 分段，或给字符串列表，每项一段。**一段文字一个文本框**，不要一行一个 |
| `rect` | `x* y* w* h* fill alpha* outline name`，可选 `gradient: {angle, stops: [[位置 0–1, 颜色, 不透明度 0–1 或 null], …]}` | `fill: null` 为无填充；`gradient` 是线性渐变，`angle` 按 CSS（180 = 从上到下） |
| `shape` | `kind`（`triangle` / `rtTriangle` / `parallelogram` / `ellipse`）`x* y* w* h* fill alpha* outline name flip_h flip_v gradient` | 预设形状，如斜切遮罩 |
| `hline` / `vline` | `x* y* w*` / `h*`、`color` | 1px 细线 |
| `arrow` | `x* y* size* color` | "→" |
| `tag` | `text* x* y* size* fill color h*`，可选 `pad_ratio`(0.6154) `h_ratio`(1.3) `text_extra`(20) | 色块宽 = 字宽 + 左右留白 |
| `highlight` | `text* x* y* size* fill color row_h pad*`，可选 `max_w`(520) `text_right`(900) `base`(60) `pad_ratio`(0.09) `dy_ratio`(0.02) `weight`(300) | 每行一个色块 + 一个多段文本框 |
| `image` | `slot`（字段名）`x* y* w* h* panel mask*` | 按框等比裁切；缺图画 `panel` 色面并警告；同图整份只能用一次 |
| `chart` | `chart*`、`x* y* w* h* bg`、可选 `kind`、`fonts{base,labels,doughnut_labels,value_axis,category_axis}`（px） | 原生图表：column / bar / line / doughnut，行为固定 |
| `legend` | `chart* x* y* w* label_x* label_w*`，可选 `size swatch step swatch_dy rule_dy color rule` | 环图图例 |
| `table` | `table* x* y* w* bottom*`，可选 `max_row_h font font_small small_after_rows head_fill head_color body_fill body_color border style_id pad_x pad_y` | 原生表格，行为固定 |
| `header` | `left* right* dark` | 页眉官网 + Logo |
| `logo` | `x* y* w* max_h* dark align` | Logo（只用用户给的文件） |
| `group` | `objects` | 一组对象共用 `when`/`let` |
| `repeat` | `items*` `max`（必填，正整数）`label` `objects` | 循环，名字 `item i n`；超过 `max` 截掉并警告 `<label> 有 N 项…`；校验按 `max` 算文本框上限 |

颜色写调色板角色名、`a|b`（a 没有时用 b；校验要求每套配色都至少有其中一个）、`#RRGGBB`，或结果为 `#RRGGBB` 的表达式（同样过白名单）。

### 3.2 表达式

Python 表达式语法的子集，渲染器逐节点解释（不用 `eval`），Host 侧 `src/host/style-expr.mjs` 用同一规则检查。

- 允许：数字、字符串、`True/False/None`；列表、元组、字典；`+ - * / // %`、一元 `-`、`not`；比较（含链式、`in`、`not in`）；`and/or`；`a if c else b`；下标与切片；名字；调用下面的函数或 `x.get(k[, 默认])`。
- 不允许：其他属性、lambda、推导式、星号展开、关键字参数、f-string、赋值、导入。名字只用 ASCII，整数不以 0 开头（Host 侧比 Python 严，写包时按严的来）。
- 读本页字段只用 `d['x']` 或 `d.get('x')`，`x` 必须在 `fields` 里声明（`layout`、`notes`、`sources` 除外），校验会查。
- 名字：`d`（本页）、`deck`（整份）、`style`（style.json）、`P`（当前配色）、版式 `vars`、`let`，repeat 内 `item i n`。
- 函数：`len str int float min max abs round`；`lines(t)` 行列表；`line(t, k)` 第 k 行（越界为空）；`tw(t, size)` 估字宽（px）；`fit(t, size, max_w[, min_ratio])` 放不下时缩字号（下限 0.62 倍，到下限会警告）；`pad2(n)` 两位编号；`fmt(v, spec)`；`year()`；`pct(values, k)`。
- `let` 在第一次用到时求值并缓存，所以 `fit` 的警告出现在用到它的对象那里。

### 3.3 实时部分

```json
"layouts": [ { "id": "cover", "name": "封面", "placeholders": ["title", "subtitle"], "background": "primary", "boxes": { "title": [l, t, w, h] } } ],
"slots":   [ { "id": "left-image", "layout": "title-only", "kind": "image|text", "box": [l, t, w, h], "art_size": "1248x1080", "text_box": [l, t, w, h], "darken": 0.45, "note": "左图右文" } ]
```

- 单位 pt，页面 720×405。`name` 是 template.pptx 里版式的名字，`placeholders` 用 `title / subtitle / body`；`boxes` 必须与模板占位符几何一致（±0.5pt）。
- 每个槽位的 `box` 要在 `layouts.md`「实时方式」节以 `l, t, w, h` 原样出现。
- Host 把这部分原样放进 `presentation_create` 返回的 `style.layouts`。

## 4. 校验

```sh
node tooling/style-distill/validate.mjs                 # 所有包；找不到带 python-pptx 的 Python 时失败
node tooling/style-distill/validate.mjs --style <id> --python <python>
```

检查项与阈值（字号写不出静态下限算错误；`highlight` 按 0.62 倍缩字下限算）：文件齐全；`style.json` 与 `meta` 字段；模板主题颜色、字体与 style.json 一致，模板版式占位符与 `live.layouts` 一致；离线文字 ≥ 9pt（`fit` 按 0.62 倍下限算），图表文字 ≥ 8pt，实时字号 ≥ 9pt；每页文本框 ≤ 30；图片槽与 `meta.requires.photos` 一致；表达式通过白名单且名字都有定义；Host 能读 layouts.json；渲染器能读且能出 example；thumb.png 合规。

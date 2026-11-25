# 风格蒸馏工具集

把别人做好的 PPT 风格（一个 skill 包、一套模板、一份原稿，或者只有几张截图、一份 PDF）整理成 shimo-pptx 的样式包 `styles/<id>/`。只给开发用：`tooling/` 不进插件包；客户拿到的只是 `skills/shimo-pptx/` 里的样式包和渲染器。

样式包有哪些文件、每个字段什么意思，见 [pack-format.md](pack-format.md)。

## 准备

| 要什么 | 用在哪 | 本机怎么有 |
|---|---|---|
| Python 3.9+（只要标准库） | 预检 `preflight.py` | 系统自带的 `python3` |
| 带 python-pptx、lxml、Pillow 的 Python | 渲染器、校验、对比、缩略图、回归 | DSH 自带：`/Applications/DeepSeek Harness.app/Contents/Resources/runtime/primary-runtime/dependencies/python/bin/python3.12`（下文记作 `$PY`）；或设 `SHIMO_PYTHON` |
| Node 22+ | 校验 `validate.mjs`、单测 | 开发机已有 |
| Microsoft PowerPoint for Mac | PPTX 转图片（对比、缩略图） | 装在 `/Applications` |
| poppler 的 `pdftoppm`、`pdfinfo`、`pdftotext`、`pdffonts` | PDF 转 PNG、读页数/文字层/字体 | `brew install poppler` |

**PowerPoint 转图第一次运行**会弹「是否允许终端控制 Microsoft PowerPoint」，要有人在电脑前点允许（系统设置 › 隐私与安全性 › 自动化 里也能开）。没点或被挡住时，`pptx2png.py` 等 180 秒后报错退出（退出码 3），不会一直卡着。转图时文件会先拷进 PowerPoint 的沙盒目录，转完就删；只关闭自己打开的那份，不会退出 PowerPoint。

```sh
PY="/Applications/DeepSeek Harness.app/Contents/Resources/runtime/primary-runtime/dependencies/python/bin/python3.12"
T=tooling/style-distill
```

## 五个阶段

### 1. 预检：这个包值不值得做，有什么拦路的

```sh
python3 -I $T/preflight.py ~/Downloads/某风格.zip --out preflight.md --json preflight.json
```

- **产出**：一页报告：产物格式（PPTX / HTML）、运行时依赖、字体和图片的授权与出处（标出不能分发的）、设计来源摘录、照片需求、最小字号，最后是「能否进入提取」和「要人决定的事」。
- **安全**：源包当作不可信数据。zip 先检查（绝对路径、`..`、符号链接、条目数、解压大小），通过后才解到临时目录、设成只读、用完删除。包里的脚本只读不跑。
- **通过**：报告结论是「能」；「要人决定」的每一条都已经问过用户，并有答复。有一条「不能分发」的字体，就要在适配阶段换掉（只写字体名，不打包）。

**只有参考稿**（截图、总览图、PDF、PPTX，或只放着这些文件的目录）时，预检自动按「参考稿」口径出报告：

```sh
python3 -I $T/preflight.py 参考稿.pdf --out preflight.md --json preflight.json      # 也可以是 .pptx、一张图、一个目录
```

- 报告首行写明用了哪种模式和理由。单个图片 / PDF / PPTX → 参考稿；目录里只有图片、PDF、PPTX、字体和 README/txt/md → 参考稿；有脚本、json、css、`SKILL.md` 等 → 源包（和以前一样）；zip → 源包。`--ref` / `--package` 可以强制。**注意**：单个 `.pptx` 以前会被当成源包解开，现在按参考稿处理。
- 参考稿报告：每个文件的类型、sha256、页数、尺寸和画幅；PDF 的文字层有无、字体名和是否嵌入（要 poppler；没有时页数按 `/Type /Page` 粗数，数不出就写「页数未知」）；PPTX 的页数、页面尺寸、母版、用到的版式、主题字体和正文字体名、媒体和嵌入字体数（`.pptm` 等带宏的不读）；字体文件的授权判定和 **OS/2 fsType**（可安装 / 限制嵌入 / 仅预览打印 / 可编辑 / 不可子集 / 仅位图）。源包报告的字体表也多了 fsType 一列。
- 「要人决定」第一条永远是**设计归属待用户确认**：这套设计归谁、能不能给客户用。确认前 `meta.design_source.cleared` 保持 `false`。

### 2. 提取：把版式变成 `layouts.json`

没有通用脚本，因为每个源包的版式数据形态不一样。按来源分三种做法：

| 源包里有什么 | 怎么做 |
|---|---|
| 对象级母版 JSON（brand-ppt、松烟这一类：每页对象带 x/y/w/h、字号、`maxLines`） | `python3 -I $T/master2layouts.py master.json --out draft.json --report draft.md`：文字 → `text`，色块 → `rect`/细线，三角 → `shape`，图 → `image`；几何照抄（同为 960×540 画布），`row1_*`、`step1_*` 这类编号对象合成 `repeat` |
| 只有 HTML/CSS（Noir、Genspark 这一类） | `node $T/measure_html.mjs --root <源包 assets> --layouts <layouts.json> --slides slides --out measure.json` 量出每个区域和文字槽，再 `python3 -I $T/gs2layouts.py measure.json --layouts <layouts.json> --out draft.json --notes notes.md` 出草稿；1920 宽折成 960，小字抬到 9pt，CSS 才有的效果（滤镜、局部变色、图示）在 notes 里逐条列出替代或放弃。密的页要人工删字、改表格 |
| 只有 PPTX 原稿 | `"$PY" -I $T/ref_intake.py 原稿.pptx --out ref/` 盘点版式、字体名、媒体并出逐页图（经 PowerPoint；`--no-render` 只盘点）；`$PY -I $T/inspect_pptx.py dump 原稿.pptx` 列出每页形状的位置和字号，照着写 |
| 只有图片 / 总览图 / PDF | 见下面「只有图片或 PDF」 |

#### 只有图片或 PDF

1. **归一成逐页图**：

   ```sh
   "$PY" -I $T/ref_intake.py <截图目录 | 总览图.png | 参考稿.pdf> --out ref/
   ```

   得到 `ref/pages/01.png…` 和 `ref/intake.json`（每页尺寸、sha256、来自哪个文件/哪一页/总览图的哪个框；PDF 的文字层与字体）。目录按文件名自然排序（`2.png` 在 `10.png` 前）。PDF 按 `--width`（默认 1920）出图。
   - **单张图**：16:9 或 4:3 的图当一页。其他比例的图按页间的缝切成逐页：背景色取图的外框，页面要是底色实心、大小比例一致的矩形。某页底色和背景几乎一样时，如果其他页排成规整网格，就按网格补上这一页，并在 `notes` 里提示核对。判不清时退出 3，不会猜：这时用 `--boxes boxes.json`（`[[x0,y0,x1,y1],…]`，源图像素）给框，或 `--single` 当一页。2×2、3×3 这类 16:9 页面拼成的总览图，整体也是 16:9，会被当成一页并给出警告，要加 `--montage`。带投影或描边的总览图，切出的页框会把阴影算进去，这时也用 `--boxes`。
2. **量**：

   ```sh
   "$PY" -I $T/measure_image.py ref/intake.json --out measure.json                         # 每页底色、主色；PDF 有文字层时附行框
   "$PY" -I $T/measure_image.py ref/intake.json --page 3 --roi card=50,160,432,262 --roi title=50,82,450,50
   ```

   ROI 用 960 画布坐标（`--units src` 改成源像素），先看图手画、留一点边。每个 ROI 给出：区域底色、墨迹框（画布与源像素）、墨迹主色、逐行上下沿、行距，以及字号估计（按大写字高、按混排字高、按行距三种）。**字号只是估计**：试验里只看图估的字号偏大约 20%，以第 4 阶段对比为准。卡片底色和页面底色很接近时（通道差小于 `--threshold`，默认 24），卡片不算墨迹，它的颜色会出现在框在卡片里的 ROI 的「区域底色」和整页主色里。
3. **识字**：没有 OCR 引擎。PDF 有文字层时用 `measure.json` 里的 `text_lines`（标 `pdf-text-layer`）；没有时由 agent 看图识读。`notes.md` 里每段文字都标来源：PDF 文字层 / OCR / agent 识读 / 低可信。
4. **写**：照着量出的颜色和框手写 `style.json` 与 `layouts.json`。图上量不出的要推断，并在 notes 里写明是推断：字体身份（写进 `meta.fonts` 时 `source` 写「看图推断」）、字段划分、可选/必填、`repeat`、容量。

- **产出**：`styles/<id>/` 下的 `style.json`（配色、字体、`meta`）、`layouts.json`、`example.json`。
- **通过**：`$PY -I packages/dsh-adapter/skills/shimo-pptx/scripts/brand_deck.py layouts --style <id>` 能列出全部页型，`build --spec styles/<id>/example.json` 能出稿。

### 3. 适配：改到符合我们的规矩

逐条对着下面「以往的坑」改，然后：

```sh
"$PY" -I packages/dsh-adapter/skills/shimo-pptx/scripts/build_template.py --style <id>   # 实时方式模板
"$PY" -I $T/thumb.py --style <id>                                                        # 缩略图 thumb.png
node $T/validate.mjs --style <id>
```

- **产出**：`template.pptx`、`thumb.png`、写好实时方式坐标的 `layouts.md`；`meta.requires.photos` 按校验算出的数填。
- **通过**：`validate.mjs` 全绿（它会检查字段、模板与 style.json 一致、最小字号、每页文本框数、图片槽数、表达式、Host 和渲染器都能读、缩略图）。

### 4. 对比：同一份内容，原稿和我们的并排看

```sh
"$PY" -I $T/compare.py --theirs 原稿.pptx --spec styles/<id>/example.json --style <id> --out 对比/ --labels 原稿,样式包
```

`--theirs` 也可以是原 skill 产物导出的 PDF，或一个放着逐页截图的目录（按文件名排序）。

- **产出**：`对比/pages/NN.png`（每页左原稿、右我们）、`对比/compare.png`（总览）、`checks.json` 和 `report.md`：溢出、重叠、小于 9pt 的字、重复图片、文本框和段落数、疑似一行一个文本框、整页图片、外链。
- **通过**：我们这一侧没有溢出、重叠、重复图片、小于 9pt 的字、疑似一行一个文本框（命令退出码 0）；并排图逐页看过，差别都说得清楚。

### 5. 关卡：人来定

- **逐页效果**：把 `compare.png` 和各页并排图给用户，用户逐页确认。没确认的页不算数。
- **设计来源和授权**：预检报告里的出处、借用、字体授权，由用户或法务拍板。同意后把 `meta.design_source.cleared` 设为 `true`，并在 `note` 写清是谁、什么时候同意的。`cleared` 为 `false` 时校验不通过，包进不了 `styles/`。
- **两项都通过后才登记**：样式包放进 `packages/dsh-adapter/skills/shimo-pptx/styles/<id>/`，在本机（有 Python）跑完整校验和全量测试：

```sh
node $T/validate.mjs                       # 所有包；找不到带 python-pptx 的 Python 会失败
node --test packages/dsh-adapter/test/*.test.mjs
```

CI 没有 Python，只跑不需要 Python 的那部分（字段、模板、字号、文本框、图片槽、表达式、Host 读取）；渲染器能否出稿、回归金样只在本机跑，所以这一步不能省。

## 暂存包

设计来源还没确认（`cleared: false`）的包放 `tooling/style-distill/staging/<id>/`，用法和 `styles/` 下的一样，只是各脚本加 `--styles-dir tooling/style-distill/staging`，校验用 `node $T/validate.mjs --styles tooling/style-distill/staging`（`cleared: false` 只记「待登记」，别的检查照常；id 和别名也不能和已登记的包重）。CI 的 `style-packs.test.mjs` 也按这个规则校验暂存包（暂存目录没有包时也算通过）。

打包不读暂存目录，adapter 包的 `files` 也够不到 `tooling/`，所以包里不会带暂存包。确认授权后：`cleared` 改 true、整目录挪进 `styles/`，就算登记。

## 改渲染器时的回归

`brand_deck.py` 或 default 的 `layouts.json` 一改，就跑：

```sh
"$PY" -I $T/regress.py check
```

它用 default 的 `example.json` 和 `fixtures/default-regression*.json` 出稿，逐个部件（规范化后的 XML、嵌入工作簿、图片）和金样 `fixtures/default-golden.json` 比，构建报告（含警告顺序）也比。有意改变产物时，先用 `compare.py --theirs <旧产物>` 看清每页变化、写进说明，再 `regress.py golden` 重录金样。金样用 DSH 运行时的 python-pptx 1.0.2 录的，版本不同会提示。

## 以往蒸馏的坑

| 坑 | 出现过在 | 怎么办 |
|---|---|---|
| **一行一个文本框** | brand-ppt-12、松烟原稿（`body-line0`、`body-line1`…） | 一段文字一个文本框，`\n` 分段；`compare.py` 会报「疑似一行一个文本框」 |
| **字体不嵌入** | 所有 PPTX 样式 | 样式包只写字体名，不打包字体文件；客户电脑没装会被替换，所以别选只有一种字体才好看的风格；`meta.fonts` 写清授权和来源 |
| **字体不能分发** | Noir 的 Times 是从 macOS 系统字体抽出来的 | 预检会标「不能分发」；换成开源字体 |
| **照片太多** | 松烟一份要 19 张不同的照片 | 找图能力有限，缺图时会大量落到生成配图，效果掉很多；照片多的页型要么不收，要么在 `meta.requires.note` 写清，并给出少图的替代页型 |
| **HTML 产物** | Noir 出的是 HTML 演示 | 进不了 DSH 的 PPTX 链路；要从页面反推几何，工作量和失真都大，预检会直接标出来 |
| **小字号** | Noir 最小 3pt（1920 宽画布上 8px） | 不能只把几何减半；校验要求文字 ≥ 9pt、图表 ≥ 8pt，放不下就删字或拆页 |
| **Codex 运行时依赖** | brand-ppt-12、松烟要 `@oai/artifact-tool`、`skia-canvas` | DSH 和客户环境都没有；只取它的母版数据，交给 `brand_deck.py` 渲染 |
| **整页宽居中** | 参考图上的主标题看着偏在一边（被两侧角标挤的） | 主标题按整页宽居中（或按版式的左边距），不要照抄量出来的偏移中心 |
| **只看图估字号偏大** | 图片试点估的字号偏大约 20%（94/76、80/64、42/36） | `measure_image.py` 的字号只当上限参考；对比阶段看溢出和并排图再收 |
| **设计借用** | 松烟重建自 Kimi Slides，Noir 取自 Genspark 模板 | 改名去不掉借用；由用户或法务定，`meta.design_source.cleared` 记录结论 |
| **浅色主色** | 松烟主色是白色，Noir 是香槟金 | `build_template.py` 把封面、结束页设成主色底，会做出白色或金色封面；深色风格要先改模板的背景规则 |
| **默认风格专属规则** | `references/style.md` 里的高亮色块、箭头、封面与 THANKS 同构 | 新风格不一定有这些；这类规则写进该样式自己的 `layouts.md` |
| **增量发版漏文件** | Host 从本工具集起要读 `styles/<id>/layouts.json` | 只换 3 个 JS 的增量发版要连 `skills/shimo-pptx/styles/` 一起带上，否则 `presentation_create` 报 layouts.json 读不了 |

## 文件

| 文件 | 作用 |
|---|---|
| `preflight.py` | 预检报告（标准库）：源包，或参考稿（图片 / PDF / PPTX） |
| `validate.mjs` | 校验样式包；`validatePacks()` 也被单测调用 |
| `compare.py` | 并排对比 + 自动检查 |
| `pptx2png.py` | PPTX / PDF 转逐页 PNG（PowerPoint + pdftoppm） |
| `thumb.py` | 生成 `thumb.png` |
| `ref_intake.py` | 参考图片目录 / 总览图 / PDF / PPTX → 逐页 PNG + `intake.json`（总览图按缝切，判不清就要框） |
| `measure_image.py` | 量参考页：底色、主色、ROI 的墨迹框与颜色、行与字号估计；PDF 文字层行框 |
| `inspect_pptx.py` | 把 PPTX 读成形状清单（`dump`）或部件哈希（`parts`） |
| `master2layouts.py` | 对象级母版 JSON（松烟这类）→ layouts.json 草稿：编号对象合成 repeat，字号不低于 9pt |
| `html_survey.py` | HTML 样式库批量预检（拦路项、照片、最小字号、字体授权）+ 带缩略图短名单 |
| `measure_html.mjs` | 在隔离副本、禁 JS、禁网的无头 Chrome 里量 HTML 页的框和计算样式 |
| `gs2layouts.py` | 量出来的 HTML 页 + 源 layouts.json（Genspark / Noir 这类）→ layouts.json 草稿和转换说明 |
| `live_preview.py` | 实时方式的本机预览：模板每个版式一页 + 每个槽位一页，出 PPTX 和图 |
| `regress.py` | 渲染器回归（金样） |
| `fixtures/` | 回归夹具、金样、表达式两端一致的用例 |
| `pack-format.md` | 样式包格式 |
| `staging/` | 设计来源还没确认（`cleared: false`）的样式包，格式同 `styles/<id>/`；确认后挪进 `styles/` |

单测：`packages/dsh-adapter/test/style-packs.test.mjs`（CI 跑）、`packages/dsh-adapter/test/styles.test.mjs`（要 DSH 的依赖包，本机跑）。

---
name: shimo-pptx
description: 新建 PPT（.pptx）：默认在 DSH 右侧的 OfficeSDK 编辑器里实时生成（presentation_create 按总页数新建，再分批 presentation_edit，用户在幻灯片浏览视图看每页生成）；工作汇报的环形图和柱形支持随模板实时出稿；其他需要新插入表格或图表、批量出多份、当前环境未提供实时工具或用户要脚本出稿时，改用离线脚本 brand_deck.py。两种方式共用同一个样式包（用户没点名时用工作区默认样式，没有就是 default：蓝紫主色、Noto Sans SC / Poppins 的品牌风格）。用户没给图时，用随附脚本上网找可商用的照片，找不到就按主题生成配图，并写明图片来源。
whenToUse: 用户要做新的演示文稿/PPT/幻灯片，且没有给自己的模板或品牌规范时加载。默认走实时方式；需要模板以外的新表格或图表、要批量出多份、当前环境没有 presentation_create（headless、命令行）、用户明确要脚本出稿或要与 brand-ppt 版式像素级一致时走离线方式。编辑已有 PPT 时不用本 skill，直接用 presentation_read / presentation_edit。
metadata:
  version: 0.0.11
  distilledFrom: brand-ppt-12 v1.4.1
---

# 新建 PPT

收到保存成功回执（含专用表格工具返回 saved）后，用 DSH 的 `present` 交付这份最终文件一次。只交付本次完成的文件，创建占位和写入失败时保留现场。

桌面端切换会话时，编辑器会等待原会话重新显示。实时工具返回 `no-editor` 时，告知用户切回这条会话，回到原路径用读取工具确认文件和写入进度，再继续实时写入。已经创建的文件只读回恢复，不再次新建同名或改名文件，也不要用脚本另出一份。

## 已选基础模板的实时生成

用户要求工作汇报、工作总结或季度复盘 PPT 时，默认用下方“团队工作汇报”原模板；要求培训、员工培训或经验分享 PPT 时用 `slide-008`；要求空白 PPT 时用 `slide-000`。用户明确点名其他样式时按用户要求。用户只给场景也可选择，不要求用户知道模板编号。以下模板优先于本文件其他出稿方式：

- **团队工作汇报**（`slide-002`，8 页）用实时方式。先调用 `presentation_create(path, style="slide-002", slideCount=8, reportingData=…)`。仅准备图表所需的数字，无需先写全稿或运行脚本。`reportingData` 可含 `chart: {title, categories: [三个类别], values: [三个非负数]}`、`bars: [五个非负数]`、`statuses: {page2: [四个状态], page6: [三个状态]}`（green、amber、red、neutral）。只有用户给出或允许模拟时才传对应数据。缺数据时省略该项。原生环形图、图例、可编辑柱条及状态色在创建时准备好，页上文字在右侧逐批生成。根据返回的真实页和形状 ID、`style.textSlots`，一至两页一批用 `presentation_edit` 的 `setText` 填写各页；环形图中央百分比和三行图例已按数据计算，保留其值，柱条数值文字与传入数据一致。用 `completedSlideIds` 标出本批写完的页，八页完成后 `presentation_save`。
- `slide-008`：`presentation_create(path, style="slide-008", slideCount=7)`；从返回的 `deck.slides` 取得每页真实 `slideId` 与各形状 ID，结合 `style.textSlots` 中每页的形状名，按页用 `presentation_edit` 的 `setText` 填写适合该位置的文字。每个文本框独立填写，字数适应原版；同一页分批时中间传 `completedSlideIds: []`，完成该页时传该页 ID。保留模板页面顺序、样式和可编辑对象。第二页原示例折线的数据点已清除；没有用户数据时不填数值，不画折线，不凭空编造数据。确有数据时可用可编辑形状绘制符合实际数据的图。全部完成后 `presentation_save`。
- `slide-000`：`presentation_create(path, style="slide-000", slideCount=1)`。只要求空白页时直接 `presentation_save`；要求写内容时根据用户信息在空白页实时添加文本或形状，之后保存。

以上三套模板都应先打开右侧编辑器，再调用编辑。确定页数后先打开右侧，内容准备好一页就提交，可按线上交互并行生成多页；不要为了确认字数而先检索源代码或读取样例 PPT。不要在对话中要求用户描述内部动画步骤。若没有资料，不把模板的样例数字当作事实。

在无实时工具的环境中，固定页模板也可用 `brand_deck.py build --style <编号> --spec <输入.json> --out <文件.pptx>`。每页使用 `text-slots.json` 中的文本框名称，输入示例见该模板的 `example.json`。工作汇报支持同一组图表数据字段。

## 1. 选出稿方式

| 情况 | 方式 |
|---|---|
| 其余情况（默认） | **实时**：右侧编辑器里分批生成，见第 4 节 |
| 需要新插入表格或图表（工作汇报的既有图表走上述实时方式；实时方式插入的表格、图表目前存不下，office-kit K6） | **离线**：见第 5 节 |
| 要一次出多份稿 | 离线 |
| 没有 `presentation_create` 工具（headless、命令行等没有右侧编辑器的环境） | 离线 |
| 用户明确要脚本出稿，或要与 brand-ppt 版式像素级一致 | 离线 |

拿不准时用实时方式，并告诉用户：要表格或图表时可以改用离线脚本出稿。

## 2. 选样式

样式包在本目录 `styles/<样式名>/`，每个包有 `style.json`（配色、字体、字号）、实时方式模板 `template.pptx`、版式描述 `layouts.json`（实时方式的版式和槽位、离线脚本画的页型）、版式说明 `layouts.md`、一份示例 `example.json`，以及选择器卡片用的缩略图 `thumb.png`；`style.json`、`template.pptx`、`layouts.json` 三者缺一不算样式包。`style.json` 的 `meta` 写着这个包的显示名（`name`）、别名（`aliases`）、一句话特点（`summary`）、适用场景（`scenarios`）和前提条件（`requires`：`photos` 是图片槽数量，`note` 是说明）；包名（`meta.id`）就是目录名。

按这三层选，先对上哪层用哪层：

1. **用户点名了**：消息里的名字对上某个包的包名、`meta.name`（中文或英文）或 `meta.aliases`（不分大小写）就用它。输入框「灵感库」里选了 PPT 模板时，消息末尾是「请使用「<显示名>」样式生成 PPT 演示文稿。」，会话上下文里同时有一条「Inspiration library pick」写明包名；推荐卡片写「用 <包名>（<显示名>）样式」。这些都是点名，按上下文给的包名传（`presentation_create` 也认显示名和别名）。实时方式把包名传给 `presentation_create` 的 `style`，离线方式传给 `--style`。
2. **用户没提样式**：用工作区默认样式。实时方式**不传** `style`，插件会用工作区默认（用户在「灵感库」里勾选记住的，没有就是 `default`），结果里的 `style.name` 是实际用的样式。离线方式：会话上下文里有「Workspace default PPT style」一条时用它的包名，没有就用 `default`。不要因为"没说"就自己传 `default`。
3. **用户的说法对不上任何包**（"好看点""高级感""换个风格""商务一点"）：不自己定，按下面「推荐样式」让用户选；只有一个包时直接说明只有它。

**推荐样式**（`presentation_suggest_styles`）。要**新建**一份 PPT、用户**没点名**样式、会话上下文里**没有**「Workspace default PPT style」、`styles/` 里**不止一个**包时，先推荐再动手；用户只说"做个 PPT"、只给了主题，或第 3 层说法对不上任何包，都算：

- 先读懂内容：主题、给谁看、什么场合、用户有没有照片。再按各包 `meta.summary`、`meta.scenarios`、`meta.requires` 挑 2–4 个最合适的，好的排前面。
- 调 `presentation_suggest_styles`：每个候选写包名和一句**结合用户内容**的理由（用户的语言，例如「重庆夜景照片多，黑底大图最出效果」），`topic` 写几个字的主题。界面会在输入框上方显示带缩略图的卡片，另有「都不要，用默认」「看全部样式」。
- 调完就**结束本轮**，只回一句「选一个版式，或直接告诉我想要的」，不要在文字里再列一遍，也不要先 `presentation_create`。
- 用户的回复：「用 <包名>（<显示名>）样式」就是点名（第 1 层）；「都不要，用默认样式 <包名>」用那个包；「第 N 个」、显示名、别名按推荐的顺序和列表认（会话上下文里会列出上一轮推荐了哪些）；说了别的要求（"再暗一点"）就照那个要求重新挑或再推荐一次；实在对不上就再问一次，不要自己定。
- **不推荐**：用户已经点名、工作区有默认样式、只有一个包、编辑已有 PPT、或是子代理（只在主会话里推荐，子代理的推荐用户看不到）。用户说"随便""你定"时用工作区默认（第 2 层），不推荐。

- 选中的包 `meta.requires.photos.required` 较多（或 `meta.requires.note` 说要照片）、而用户没给图时，先提醒用户（可以照 `note` 说）：第 3 节上网找图不一定凑得齐，凑不齐时页面效果会打折扣；问用户是继续、自己提供图片，还是改用 `default`。
- 用户只说"换个颜色"时不换样式包：离线方式在 deck JSON 里换 `palette`（`style.json` 的 `palettes`），实时方式按 `style.json` 的配色自己设颜色。
- 名字不在 `styles/` 里会报错并列出可用样式；照实告诉用户有哪些样式，不要自己编一个。

风格规则两种方式共用，见 [references/style.md](references/style.md)；各样式的具体数值只在它的 `style.json` 里。

## 3. 配图

带图位的页（离线方式的图像陈述、陈述、编号清单、指标、风险页，实时方式的全幅图文、左图右文、横幅图、全幅背景图）要有真图，不要留空或换成灰色色块。两种方式都按这个顺序找图：

1. **用户给的图，或工作区里与本稿主题相符的图**优先。工作区里别的主题的素材（比如做重庆却只有武汉、华为的图）不能拿来用。
2. **上网找**：用本 skill 的 `scripts/find_images.py`（Python 先按 `office-pptx` 调用 `load_workspace_dependencies` 拿到；命令都在任务工作区里运行，图片和记录写在工作区的 `images/`，不要 cd 到 skill 目录）。它查 Openverse 和 Wikimedia Commons，只收可商用、可修改的授权（CC0、公有领域、CC BY、CC BY-SA），不需要密钥：

   ```sh
   "<python>" "<本 skill 目录>/scripts/find_images.py" search --query "Chongqing skyline" --query "Hongyadong" --orientation landscape
   "<python>" "<本 skill 目录>/scripts/find_images.py" fetch --pick 2 --name cover --pick 5 --name hongyadong
   ```

   关键词用英文地名、事物名最准；把这份稿要用的关键词放进同一次 `search`（重复 `--query`）。`--pick` 的编号只对最近一次 `search` 有效，再搜会换掉候选，所以要分批搜时每次搜完先 `fetch`；`--name` 按顺序对应 `--pick`。你看不到图，只能按标题、作者、尺寸和描述选：选描述与本页内容一致、尺寸够大（全幅图宽 ≥ 1600 px）的。下载的图在任务工作区 `images/`，来源自动记进 `images/sources.json`。不要用 `web_fetch` 下载图片（它只取文字页）；`web_search` 能用时可以帮你想关键词，不能用也不影响。
3. **找不到合适的就生成配图**：`scripts/theme_art.py` 用样式包配色按主题画示意图（不是照片，不带文字）。母题可组合：`skyline`（城市）、`mountains` / `hills`（山）、`bridge`（桥）、`river` / `waves`（江河、海）、`sun`、`towers`、`network` / `grid`（科技）、`circles` / `bars`（商业、数据）、`leaves`（自然）；`--mood day|dusk|night`。

   ```sh
   "<python>" "<本 skill 目录>/scripts/theme_art.py" --motif mountains,skyline,bridge,river --mood night --style <样式名> --size 1920x1080 --out images/cover-art.png --note "重庆：山城、两江、夜景"
   ```

   `--size` 按图框比例出图：离线方式和实时方式的全幅背景图 `1920x1080`；实时方式按该样式 `layouts.md` 里图框的比例，default 是全幅图文 `1920x945`、左图右文 `1248x1080`、横幅图 `1920x533`。

4. **同一张图整份只用一次**；照片上的字反白。
5. **交付时写明来源**：运行 `find_images.py credits`，它写出 `images/sources.md`；离线方式会自动把用到的图的署名写进对应页备注。实时方式在交付说明里只写一行压缩署名（「作者 + 协议」，例如「照片：Ann（CC BY 2.0）、Bob（CC BY-SA 4.0），详见 images/sources.md」），链接留在 `sources.md`，不要把它全文贴进聊天。CC BY / CC BY-SA 的图必须有作者和协议。

## 4. 实时方式

1. **确定页数后立即新建**。用户指定 N 页就建 N 页，含封面和结束页。未指定时沿用默认 8 页。先调用 `presentation_create { path: "<名称>.pptx", slideCount: N }`，再拟写各页或找图。样式按第 2 节。可传 `slideLayouts` 指定每页的模板角色，例如 `["cover","blank","content","closing"]`。右侧打开幻灯片浏览视图，一次创建全部页面。每页有独立工作光标。
2. **按页写入**。创建结果的 `deck.slides` 已按页序列出真实页 ID 和形状 ID。用 `setText` 填模板占位符，用 `addTextBox`、`addImage` 等操作补充内容。引用对应的 `slideId`，不要再用 `addSlide` 重复建页。尺寸、字号和配色从返回的样式与布局读取。
3. **按准备好的顺序提交**。每批写一页或几页，当页有图时先下载并裁好这一页需要的图。写完一页便可提交；一页分多次填充时，在中间批次传 `completedSlideIds: []`，最后一批传本次已完成的页 ID。未传该字段时，本次写到的页面视为完成。所有页完成后编辑器自动回到普通视图第一页。只通过编辑器写入本文件。
   **图片先裁好再放**：`addImage` 会把图拉伸到给定的框，所以先按第 3 节准备好图，再按框的比例裁。default 样式里城市、风景这类横图用全幅图文（`photo-statement` 槽；其他样式用它 `layouts.md` 里的全幅或半幅图槽）：`find_images.py fit --in images/skyline.jpg --box 720x354.6 --darken 0.3 --out images/skyline-photo.jpg`；左图右文（312×270 近方框）只给竖构图，横图不要塞进去。字要压在照片上时加 `--darken`（蒙版直接做进图里，字用白色）。按图片描述判断主体偏在哪边：主体会被裁掉时加 `--anchor top|bottom|left|right`，用同一个 `--out` 重裁（默认 `center`）。生成配图按第 3 节给的尺寸出图（与图框同比例），尺寸不对时同样用 `fit` 裁。图片框位置见该样式 `layouts.md`「实时方式」。
   **封面、全幅图文、结束页**（default 样式；其他样式看它的 `layouts.md`）：
   - 封面填写已有的标题和副标题占位符。`title` 两行（`\n`），每行 2–6 个汉字，例如「厦门\n城市介绍」「鹭岛\n厦门」，不要只写一个词；`subtitle` 一句真实定位，例如「福建东南沿海的经济特区」，不能空着，也不要把标题再写一遍。
   - 全幅图文：本页 `layout: "blank"`，先 `setBackground` 设成主色（顶部露出主色条），再按 `photo-statement` 槽放图；`title_box` 放两行反白标题，`text_box` 放一段反白短说明。
   - 结束页只写 THANKS，不加别的字。
4. **检查**：一次调用要么整批生效，要么整批撤回——一批里任何一页出错，这一批的几页都会撤回，之前的批次保留。没报错就继续下一批；报错时按提示改这一批，整批重新提交。用户在聊天中停止生成（结果是 `cancelled`）时，正在做的这一批同样整批撤回，之后的修改会先弹卡片让用户确认——这时不要自动重提，停下来问用户要不要继续。
5. **保存**：全部完成后调用 `presentation_save`。

**文案**（写进幻灯片的字）：
- 大标题（封面、全幅图文这类给了两行高度的标题框）两行、每行 2–6 个汉字（或同样短的英文）；内容页的页标题只占一行，写一个短句，不要换行。一段意思放一个文本框；放不下就缩短，不要把一句话拆成两行、两个框或两页。
- 成品页只写内容。生成过程里的事（限流、裁切失败、找不到图、未写入、重做了哪页）不要写进任何一页。
- 用户没给的数字（人口、面积、游客数、GDP 等）不要编，也不要在页上写「暂无数据」。

**对话**（聊天里的字）：用户在右侧看着页面出现，过程卡片会记下每一步。
- 生成过程中**不要往聊天里写进度**：不写「封面已出现」「裁切没问题」「八页都已写入」「再检查一遍」这类话，直接调用下一步工具。
- `presentation_edit` 的 `summary` 一句话，例如「封面」「第 2–4 页：城市概况、鹭岛、鼓浪屿」。
- 完成后只回两三句：几页、文件名（路径）、用了网上的图时一行压缩署名（见第 3 节）。不要列页目录，不要讲检查和返工过程，不要贴 `sources.md` 全文。

实时方式现在做不到：表格、图表（存不下，见第 1 节，改用离线方式）；放映动画和切换效果；同一文本框里只让部分文字加粗或变色。

## 5. 离线方式

先按 `office-pptx` 的要求调用 `load_workspace_dependencies` 拿到 Python 路径；然后**用本 skill 的生成脚本出稿**，不要自己从零写 python-pptx 版式。脚本只读本目录资源，输出写到任务工作区。

1. **定页数与页型**。用户说 N 页就正好 N 页（含封面和结束页）；没说时默认 13 页。第 1 页 `cover`，最后一页 `closing`（THANKS），中间按内容选页型，可重复使用。一页只讲一件事。页型和字段见该样式的 `layouts.md`「离线方式」一节，完整样例见 `example.json`。
2. **准备图片**（第 3 节），**写 deck JSON**（放任务工作区，例如 `deck.json`），把图片路径写进各页的 `image` 字段。
3. **生成**：

   ```sh
   "<python>" "<本 skill 目录>/scripts/brand_deck.py" build --spec deck.json --out <名称>.pptx --style <样式名>
   ```

   `--style` 按第 2 节定（脚本不写 `--style` 就是 `default`，它不知道工作区默认样式）。输出 JSON 列出样式、页型、Logo 是否放入、写进备注的图片署名（`image_credits`）和警告。有"超出文本框""标题过长"警告就精简文案后重跑（加 `--overwrite`）；有"未提供 image"警告就按第 3 节补图后重跑——脚本缺图时只是临时换成色面，不能这样交付。
4. **检查与交付**：按 `office-pptx` 回读页数和内容；要看效果时加载 `office-render` 渲染检查。用 `present` 交付最终 .pptx。离线生成的文件之后可以用 `presentation_read` 在右侧打开给用户看。

## 不做

不上网找 Logo，也不用生成图、文字或形状仿 Logo（Logo 只用用户给的文件）；配图不用 NC（禁止商用）、ND（禁止修改）授权或来源不明的图；不做桑基、旭日等复杂图；不嵌入字体；不改用户没让改的已有文件。

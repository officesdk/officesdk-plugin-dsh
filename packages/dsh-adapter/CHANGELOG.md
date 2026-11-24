# OfficeSDK 更新日志

每个版本一节，写给安装和使用插件的人看。没定版的改动先记在「未发布」下，定版时改成版本号和日期。

版本节里可以加一个「### 更新要点」小节：1–5 条用户一看就懂的话（说用户能感到的变化，不写内部实现）；紧跟一个「### Highlights」小节，逐条写同样要点的英文（条数相同，DSH 界面为英文时显示这一份；有更新要点却没有英文，生成脚本会报错）。插件更新后第一次打开文档时，会在标签页右上角弹出「本次更新」小提示，列出从旧版本到新版本之间各版本的要点；没有这个小节的版本不提示。改完要点（包括把「未发布」改成版本号）后运行 `node scripts/whats-new.mjs`，重新生成随包的 `whats-new.json`（单测会检查两者一致）。

## 0.0.2 — 2026-11-24

### 更新要点

- 聊天框输入文字时，占位提示会隐藏，中文输入法组字也能看清。
- 带编辑引用的消息会显示文档名和修改指令，重新打开已有会话也能看到。
- 聊天框里有未发送的草稿时，仍可从文档选区发送局部修改，草稿会保留。
- 滚动聊天时，「正在编辑」一行会模糊下方经过的文字，在深色和浅色主题中都清晰可读。

### Highlights

- The chat placeholder disappears as you type, including while composing Chinese characters.
- Messages with an edit reference show the filename and your request, even when you reopen an existing conversation.
- Send an edit from a document selection while keeping an unfinished draft in chat.
- The editing status above the chat box stays readable as messages scroll behind it in both light and dark themes.

## 0.0.1 — 2026-11-24

### 更新要点

- 在 Word、Excel 或 PPT 中选中内容即可发起局部修改。聊天消息显示文档名，悬停可查看引用原文。
- 可以同时修改不同位置，或在处理期间切换文件。修改详情按文档保留，直到你采纳或撤销。
- 专注模式下可在编辑器右下角继续对话、回答 Agent 的问题。
- 创建多页 PPT 时，可以在幻灯片浏览视图中看到各页逐个生成标题、文字和其他元素。

### Highlights

- Select content in Word, Excel or PowerPoint to request an edit. Chat messages show the filename. Hover over it to read the original selection.
- Edit several passages at once and switch files while edits are in progress. Each document keeps its changes for review until you accept or undo them.
- In focus mode, continue the conversation and answer the agent’s questions in the lower-right chat panel.
- When creating a presentation, watch titles, text and other elements appear on each slide in Slide Sorter view.

## 0.0.1-rc.1 — 2026-11-23

初始开源版本。

### 更新要点

- 初始开源版本：在 DeepSeek Harness 右侧标签页打开、编辑和保存 Word、Excel、PPT 文件，对话里的 AI 能在编辑器里实时新建和修改文档。
- 自带 PPT 蓝色品牌风格和 Word 公文版式，装好就能用；文档引擎在第一次打开文档时自动下载。

### Highlights

- First open-source release: open, edit and save Word, Excel and PPT files in a tab on the right of DeepSeek Harness, and the AI in the chat creates and edits documents live in the editor.
- Comes with the PPT Blue Brand style and the Word official-document layout, ready to use after install; the document engine downloads automatically the first time you open a document.

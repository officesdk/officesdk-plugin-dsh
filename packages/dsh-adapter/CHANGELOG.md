# OfficeSDK 更新日志

每个版本一节，写给安装和使用插件的人看。没定版的改动先记在「未发布」下，定版时改成版本号和日期。

新版本按[发布前确认更新说明](../../docs/NPM.md#发布前确认更新说明)执行。每个明确功能单列一条，纯修复或维护版采用约定的一句说明。发布前确认版本号、渠道和完整中英文文案，将确认稿放在该版本的「### 更新要点」和「### Highlights」中，两种语言逐条对应。运行 `node scripts/whats-new.mjs` 同步随包说明。

## 1.0.1 — 2026-11-27

### 更新要点

- Word 和 PPT 生成效果优化。
- 修复部分已知问题。

### Highlights

- Improved document generation in Word and PowerPoint.
- Fixed some known issues.

## 1.0.0 — 2026-11-26

### 更新要点

- 新增 Word 默认风格，会议纪要、工作周报和项目计划等常用文档可直接按场景起草。
- 新增 PPT 默认风格，可用于团队工作汇报、员工培训和经验分享，也可从空白幻灯片开始。

### Highlights

- New default styles for Word help you draft meeting minutes, weekly reports, project plans, and other everyday documents.
- New default styles for PowerPoint let you create team updates, training decks, and knowledge-sharing presentations, or start with a blank slide.

## 0.0.3 — 2026-11-25

### 更新要点

- 部分已知问题的修复

### Highlights

- Fixes for some known issues.

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

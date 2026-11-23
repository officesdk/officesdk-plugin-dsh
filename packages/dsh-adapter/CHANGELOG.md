# OfficeSDK 更新日志

每个版本一节，写给安装和使用插件的人看。没定版的改动先记在「未发布」下，定版时改成版本号和日期。

版本节里可以加一个「### 更新要点」小节：1–5 条用户一看就懂的话（说用户能感到的变化，不写内部实现）；紧跟一个「### Highlights」小节，逐条写同样要点的英文（条数相同，DSH 界面为英文时显示这一份；有更新要点却没有英文，生成脚本会报错）。插件更新后第一次打开文档时，会在标签页右上角弹出「本次更新」小提示，列出从旧版本到新版本之间各版本的要点；没有这个小节的版本不提示。改完要点（包括把「未发布」改成版本号）后运行 `node scripts/whats-new.mjs`，重新生成随包的 `whats-new.json`（单测会检查两者一致）。

## 0.0.1-rc.1 — 2026-11-23

初始开源版本。

### 更新要点

- 初始开源版本：在 DeepSeek Harness 右侧标签页打开、编辑和保存 Word、Excel、PPT 文件，对话里的 AI 能在编辑器里实时新建和修改文档。
- 自带 PPT 蓝色品牌风格和 Word 公文版式，装好就能用；文档引擎在第一次打开文档时自动下载。

### Highlights

- First open-source release: open, edit and save Word, Excel and PPT files in a tab on the right of DeepSeek Harness, and the AI in the chat creates and edits documents live in the editor.
- Comes with the PPT Blue Brand style and the Word official-document layout, ready to use after install; the document engine downloads automatically the first time you open a document.

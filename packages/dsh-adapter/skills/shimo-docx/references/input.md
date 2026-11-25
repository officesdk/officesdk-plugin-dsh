# doc JSON 字段

UTF-8 JSON 对象。单行字段不能含换行。

## 通用

| 字段 | 说明 |
|---|---|
| `mode` | `auto`（默认）/ `redhead` / `plain` |
| `title` | 标题；`title_lines` 可手动分行，拼起来必须与 `title` 完全一致 |
| `subtitle`、`author` | plain 可选：副标题、署名行（居中楷体） |
| `recipient` | 主送（plain 可选，如"各部门"） |
| `blocks` | 正文块数组，见下 |
| `signing_unit`、`date` | 落款署名与日期（如"2026年11月2日"） |
| `fonts` | 覆盖字体：`title`/`body`/`heading`/`secondary`/`footer` |
| `font_policy` | `named`（默认，写标准字体名）或 `installed`（选本机已装的回退字体） |
| `draft_label` | 是否在页眉写"草稿 · 待核定"：redhead 默认 true，plain 默认 false |

## 正文块

```json
{"type": "heading", "level": 1, "text": "一、总体情况"}
{"type": "paragraph", "text": "单段文字，不含换行。"}
{"type": "list", "items": ["第一项。", "第二项。"], "numbered": true}
{"type": "table", "title": "表1 项目进展", "columns": ["项目", "进度"], "rows": [["甲", "100%"]], "column_widths": [2, 1], "note": "注：……"}
```

- `level`：1 一、（黑体） 2（一）（楷体） 3 1.（仿宋） 4（1）（仿宋）。编号写进 `text`。
- `list` 默认自动加"1. 2."；条目已以"1.""一、""（一）"等编号开头时不再加（以年份、数量开头的不算编号）。
- 表格每行单元格数要等于列数；`column_widths` 给的话必须与列数相同且都是正数；空单元格写 `""`；跨行合并单元格不支持。

## redhead 专用

| 字段 | 说明 |
|---|---|
| `document_type` | 文种：通知、函、请示、报告、批复、通报、决定、公告、通告、意见、纪要、命令、议案、其他。只给公文填；内部工作报告等不填。auto 模式下"报告""意见"需同时有发文字号或发文机关标志才出红头 |
| `redhead` | 发文机关标志，如"示例市数据局文件" |
| `document_number` | 发文字号，如"示数〔2026〕18号" |
| `signer` | 签发人（单人）；请示、报告缺签发人时预检提示 |
| `serial`、`security`、`urgency` | 份号、密级、紧急程度；用户明确给了才写 |
| `format_mode` | `ordinary`（默认）或 `letter`（信函：首页双红线、首页无页码） |
| `seal_mode` | `none` 或 `reserved`（只预留盖章位置） |
| `attachments` | `[{"title": "附件名", "blocks": [...]}]`；没有正文的附件只写附件说明 |
| `colophon` | `{"copy_to": "…", "issue_authority": "…", "issue_date": "2026年11月2日"}`；抄送≤24 字，印发机关与日期各≤14 字且成对 |

## 预检输出

`--check-only` 不写文件，输出 `mode`、`missing`（关键要素缺失或仍是占位）、`warnings`（标题与文种不符、日期无法核验、主送过长可能挤掉首页正文、版记需人工核对等）、`fonts`（每个角色用的字体）。缺项不阻止生成，生成的文档里用【待填写…】标出。

# capabilities.json

插件安装包根目录的 `capabilities.json`，给对接方的程序读取：这个版本是什么、在哪些平台、需要什么、能做什么、这一版改了什么。

## 怎么生成

- 数据源：`packages/dsh-adapter/capabilities.source.json`（手写：插件名、平台及状态、通用能力、各产品能力与 AI 工具）。以后加能力、改平台状态，只改这个文件。
- 其余字段从各自的文件读出，不手写：`package.json` 的 `engines`、`components/office-kit.json`（office-kit 版本、提交与平台）、`skills/*/SKILL.md` 的 `metadata.version`、PPT 样式包目录、Word/表格模板目录、`CHANGELOG.md` 的「更新要点」。
- 运行 `node scripts/whats-new.mjs`，同时重新生成 `whats-new.json` 和 `capabilities.json`（两者都入库、都随包）；`--check` 只比对，CI 和单测会检查。数据源缺字段、AI 工具名在 `src/tools` 里找不到、`components/office-kit.json` 里读不出 office-kit 版本时，生成直接报错。
- 入库的文件里 `version`、`channel`、`highlights` 为 `null`；`tooling/suite/pack-universal.mjs` 打包时写入，`scripts/verify-universal-package.mjs` 核对与 `package.json` 一致；回滚重发（`scripts/npm-release.mjs rollback`）只重写这三个字段，其余逐字节不变。
- Windows 转为正式（`stable`）时，只把数据源里 `win32-x64` 的 `status` 改为 `stable`；Windows 与 Mac 用同一个包、同一个标签。

## 字段（v: 2）

| 字段 | 类型 | 说明 |
|---|---|---|
| `v` | 2 | 格式版本；不兼容的改动才加 1（0.1.0 起为 2：一个包通用所有平台，去掉 `platform`） |
| `plugin.id` | string | npm 包名 `@officesdk/officesdk-plugin-dsh` |
| `plugin.name` | `{ zh, en }` | 显示名 |
| `version` | string | 本包版本，与 `package.json` 一致，如 `0.1.0`、`0.1.0-test.12` |
| `channel` | `alpha` / `beta` / `rc` / `test` / `stable` | 渠道：预览阶段 `X.Y.Z-alpha.N` / `-beta.N` / `-rc.N` 为对应阶段（npm dist-tag 同名），`-test.N` 为 `test`，`X.Y.Z` 为 `stable` |
| `platforms[]` | `{ name, status }` | 支持的平台（`darwin-arm64` / `win32-x64`）与成熟度（`test` / `stable`），与渠道无关 |
| `installVia` | `npm` | 这个版本从哪里安装：npm（0.0.0-alpha.1 起） |
| `install` | `{ stable, test[, alpha, beta, rc] }` | 现在能用的安装写法，每个渠道一个；`dsh plugin --profile <p> add <值>`：`@officesdk/officesdk-plugin-dsh@latest` / `@next`，以及预览阶段 `alpha` / `beta` / `rc`（`@officesdk/officesdk-plugin-dsh@alpha` 等） |
| `registry` | string | 发布这个包的 npm registry（`https://registry.npmjs.org/`，或打包时 `--registry` 指定的）；公网读法见 docs/NPM.md |
| `npm` | `{ status, install, registry }` | npm 安装方式：`status` 为 `enabled`（与 `install` 相同） |
| `tagPattern` | `{ stable, test }` | 发布标签的格式说明（`vX.Y.Z`、`vX.Y.Z-test.N`），不是实际值 |
| `requires.dsh` | semver range | 需要的 DSH 版本范围（`engines.dsh`） |
| `requires.node` | semver range | `engines.node` |
| `officeKit` | `{ version, commit?, platforms }` | 文档引擎 office-kit 的发布版本（发布带提交号时另有 `commit`，0.3 起不带），及有可执行 Kit 的平台；插件首次使用时按 `kit-manifest.json` 下载 |
| `common[]` | capability | 三种文档通用的能力 |
| `products[]` | product | `presentation`（PPT）/ `document`（Word）/ `spreadsheet`（表格） |
| `products[].name` | `{ zh, en }` | 显示名 |
| `products[].extensions` | string[] | 处理的扩展名 |
| `products[].capabilities[]` | capability | 该产品能力 |
| `products[].aiTools` | string[] | 模型可调用的工具名 |
| `products[].stylePacks` | string[] | 仅 PPT：本版本随附 skill 的样式包 |
| `products[].templates` | string[] | 仅 Word / 表格：AI 新建用的模板 |
| `skills[]` | `{ name, version }` | 本版本随附的 skill 与版本 |
| `skillsBaseline` | true | `skills`、`stylePacks`、`templates` 是随插件发布的基线；skill 之后可单独更新，用户实际可用的版本见仓库 `skills` 分支的 `skills-catalog.json` |
| `highlights` | string[] | 本版本「更新要点」（同 `whats-new.json`，可能为空数组） |
| `changelog` | string | 包内完整更新日志的文件名 |
| `whatsNew` | string | 包内各版本更新要点的文件名 |

## 对接方怎么拿到更新

0.0.0-alpha.1 起插件只走 npm：`capabilities.json`、`whats-new.json`、`CHANGELOG.md` 都在 npm 包里，公网读法见 [NPM.md](NPM.md)「版本与更新信息」。

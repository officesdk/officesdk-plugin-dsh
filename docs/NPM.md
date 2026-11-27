# npm 发布与安装

`@officesdk/officesdk-plugin-dsh` 只通过 npm 给用户安装（设计：spec 261111-npm-publish-prep）。本文件说明渠道、两种构建、上架门禁、
发布（GitHub Actions 工作流或本机）、移动 dist-tag、GitHub Release、回滚和公网读法。

## 渠道与安装

| 渠道 | dist-tag | 版本号 | 来自 |
|---|---|---|---|
| 预览 alpha / beta / rc | `alpha` / `beta` / `rc` | `X.Y.Z-alpha.N` → `X.Y.Z-beta.N` → `X.Y.Z-rc.N` | `scripts/npm-release.mjs publish`（经上架门禁；GitHub Actions 的 npm-publish 工作流或本机，见「发布」） |
| 正式版 | `latest` | `X.Y.Z` | 同上 |
| 测试版（git 时期） | `next` | `X.Y.Z-test.N` | 不再发布，仍可解析 |

版本与 dist-tag 的规则在 `packages/dsh-adapter/src/host/version.mjs`：同一 X.Y.Z 里 alpha < beta < rc < test < 正式版。
首个正式版之前，`latest` 可以用 npm-publish 工作流的 dist-tag 模式指向预览版（见「移动 dist-tag」）：这时不带 dist-tag 的安装和正式渠道
（`capabilities.json` 里 stable 为 `@latest`）拿到的是那个预览版。
一键更新按安装时的阶段查 dist-tag，取本阶段及更成熟阶段里最新的一版（alpha 用户也会收到 beta、rc 与正式版）。

```sh
# 先完全退出桌面版
dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh@rc               # 预览版（alpha、beta 同理）
dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh                  # latest 指向的版本（首个正式版之前可能是预览版，见下文）
dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh@<版本>           # 指定版本（回滚到旧版也用这一条）
```

DSH 把依赖写成精确版本（不记 registry、不记渠道）。插件自己的「一键更新」按渠道查 npm 的 dist-tag，再执行
`dsh plugin add @officesdk/officesdk-plugin-dsh@<新版本>`（`dsh plugin update` 不会越过精确版本）。渠道记在
`~/.dsh/plugins/officesdk/npm/channels.json`（按 profile），测试渠道收到正式版后仍是测试渠道；`OFFICESDK_CHANNEL=next|latest` 可改。

registry 顺序：`OFFICESDK_NPM_REGISTRY` → profile 的 `.npmrc`（`@officesdk:registry=` 或 `registry=`）→ npmjs，连不上再 npmmirror。
只有来自 `OFFICESDK_NPM_REGISTRY` 的地址会写进 profile `.npmrc`（只增改 `@officesdk:registry` 一行，其他行保留）。

### 包名从 `@officesdk/officesdk` 改为 `@officesdk/officesdk-plugin-dsh`（0.0.0-alpha.3）

0.0.0-alpha.1、alpha.2 以 `@officesdk/officesdk` 发布；从 alpha.3 起包名是 `@officesdk/officesdk-plugin-dsh`
（与公开仓 github.com/officesdk/officesdk-plugin-dsh 一致），scope 仍是 `@officesdk`，profile 的 `.npmrc` 那一行不用改。
旧名的一键更新只查旧名，收不到新版本，所以已装旧名的机器手动换一次（先完全退出桌面版）：

```sh
dsh plugin --profile desktop remove @officesdk/officesdk                   # 先卸旧名：两个同时装会插入同名组件、抢同一个路由
dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh@alpha
```

插件数据目录 `~/.dsh/plugins/officesdk`（文档引擎、渠道记录、设置）不变。插件不再自动改写旧包名留下的 `cordis.patch.yml` 组件行：
之前在插件页关掉或配置过的组件，换名后在插件页重新设一次。

新包名第一次发布时即使只带 `--tag alpha`，registry 也会把 `latest` 一起指向它（npm 的规则：包的第一版总是 `latest`），
所以不带 dist-tag 的 `add @officesdk/officesdk-plugin-dsh` 在有正式版之前也会装到预览版。

## 构建

`tooling/suite/pack-universal.mjs --version V --sources public --obfuscate light` 打 npm 包（`public` 是唯一的下载档位）：

| | |
|---|---|
| `download-sources.json` | Kit：`release`（office-kit 的 GitHub 发布页，即 kit-manifest 的 `sources`）；editor：`npm`（kit-manifest 的 `editor.npm`，GitHub 发布页不放编辑器）；设了 `OFFICESDK_KIT_MIRROR` 时镜像在前，首次打开文档时下载；skill：`bundled`（随包自带，不下载） |
| skill | 完整随包：shimo-pptx 只有 `styles/default`（蓝色品牌风格），shimo-docx 只有 `templates/gongwen.docx`；仓库里也只有这两样，`scripts/public-skills.mjs` 打包时再核对一遍；`skills-catalog.json` 为 `layer: full`、`remoteCatalog: false` |
| package.json | `publishConfig`（npmjs 或 `--registry`、public）、无 maintainers/scripts；`repository/homepage/bugs` 来自 `--repository` 或 `OFFICESDK_PUBLIC_REPOSITORY` |
| 混淆 | `lib/client.js`、`lib/tools`、`lib/skills`、`resources/kit-page` 轻量混淆；`lib/host` 只 minify |

来源顺序与超时在包根的 `download-sources.json`，改它不用改代码；环境变量照旧覆盖：`OFFICESDK_KIT_MIRROR`、
`OFFICESDK_EDITOR_NPM`、`OFFICESDK_SKILLS_BASE`（skill 镜像基址，只在 `download-sources.json` 写了 `public` 来源时有用，演练用）。

skill（spec 261113-bundle-public-skills）：只有 PPT 蓝色品牌风格（`default`）和 Word 公文版式（`gongwen`），随包自带、不联网下载；
`capabilities.json`、`skills-catalog.json`、组件说明也只列这两样。`spreadsheet_create` 不注册（读、改已有 xlsx 照常），灵感库 Word
只显示公文版式、不显示 Excel 模板；模型按 `SKILL.md` 传 `redhead`/`basic` 时 `document_create` 改用 `gongwen` 新建并在结果里说明（`templateFallback`）。

editor 首次使用下载（spec 261113-bundle-public-skills §6，撤回 261112-bundle-editor）：editor 页面包 `offline-editor.js`
（office-kit 0.3.0-rc.1 起 `@officesdk/officesdk-core-editor` 自带的 `dist/bundle/offline-editor.js`，K1/K16/K18 已在其中）与
`sheet-host-fonts.css` 随包，在 `resources/editor`、`resources/host-fonts`；`dist/resources`（office-kit 0.3.0-rc.1：1038 个文件，
tgz 约 33 MB）不进主包，第一次打开文档时和 Kit 一起按上表来源下载，首开因此要联网、会慢一些（下载期间文档标签页显示进度，
来源都不通时显示失败并可重试）。npm 是 editor 的最后一个来源，默认连接 15 秒、无数据 30 秒、试 3 次；`OFFICESDK_EDITOR_NPM`
可换成别的 registry（如国内镜像，照样按 sha512 校验）；设 `off` 时 editor 只剩 `OFFICESDK_KIT_MIRROR`，两者都没有就无处下载。
editor 与 Kit 是外部组件 office-kit 的一部分，不在本仓库里：随包的 editor 文件构建时按 `components/office-kit.json` 从钉住的公开发布取得（`tooling/kit/vendor.mjs`）。

## 上架门禁 `scripts/npm-gate.mjs`

```sh
node scripts/npm-gate.mjs artifacts/npm/officesdk-officesdk-plugin-dsh-<V>.tgz --target ci
node scripts/npm-gate.mjs <tgz> --target npmjs --registry https://registry.npmjs.org/ --tag latest
```

任一项不过不发布：files 白名单（含禁带 `.npmrc`、`.env`、密钥、`*.map`、`src/`、`test/`）、扫密、内部地址（私网 IP、发布方自己组织的域名、
kit-manifest 来源若是存储桶，桶里发布目录以外的路径（GitHub 这类公共代码托管不算）；kit-manifest 的来源与 editor registry 不得是组织域名或私网地址；下载档位只能列公开默认来源；其余命中必须列在 `scripts/npm-gate-allow.json` 并写原因；发往 npmjs 以外 registry 的构建，`package.json` 与 `capabilities.json` 里该包自己 `publishConfig.registry` 的主机不算）、许可（`license`、`LICENSE`、`THIRD-PARTY-NOTICES.md`、打进 bundle 的第三方包许可白名单）、
依赖审计（有 dependencies 才跑 `npm audit`，无 high/critical）、版本（`latest` 只收 `X.Y.Z`、`alpha` / `beta` / `rc` 只收 `X.Y.Z-<阶段>.N`、`next` 只收 `X.Y.Z-test.N`、同号同字节视为已发、
同号不同字节失败、registry 还没有 `latest` 时 `next` 失败）、元数据（`publishConfig.access=public`、engines、无 os/cpu/private、无安装脚本、
公网 `repository`）、通用包自检（包须恰好带两样 skill，主包不得带 `resources/office-kit`）。`--target ci` 把需要公网设施或配置的项（`repository`、非 public 档位、未配置的域名清单）降为告警；
`--target npmjs` 全部是错误。混淆不编码字符串，门禁照常能扫到（单测在混淆代码里埋点验证）。stderr 打印包内文件清单（等同 `npm pack --dry-run` 核对）。

**provenance**：不加 `--provenance`。经 npm trusted publishing 从**公开**仓库发布时 npm 自动生成 provenance；私有仓库不支持
（显式加会被 npm 拒绝），所以仓库公开前发的版本没有 provenance。仓库翻成公开后，经 trusted publishing 发布的版本自动带 provenance，
npm 包页会显示并链回源仓库、提交与工作流运行；私有期发的版本（含 0.0.1-rc.1）没有，它们与仓库的关联靠 GitHub Release（见「GitHub Release」）。
`artifacts/npm-publish.json` 记录发布结果、包 integrity、认证方式（`auth`：`oidc` 或 `token`，即 `npm-release.mjs` 提供的认证）与来源：
源提交（Actions 上为 `GITHUB_SHA`，本机为 HEAD）、该提交的 tree（main 会重建为无历史提交，tree 仍可追溯）、Actions 运行地址。

组织域名清单不进仓库：发布方用环境变量 `NPM_GATE_INTERNAL_HOSTS`（逗号分隔的域名），或 `--internal-hosts <json>` / `NPM_GATE_INTERNAL_HOSTS_FILE`，
或 git 忽略的 `scripts/npm-gate-internal.local.json`（`{ "hosts": ["corp.example"] }`）提供；都没有时 `--target npmjs` 不放行。
GitHub Actions 上它来自仓库 Secret `NPM_GATE_INTERNAL_HOSTS`（见「发布」）。

## 发布

### 发布前确认更新说明

1. 在发布对话里列出拟发布的版本号、渠道、完整中文说明和对应英文，请产品负责人逐项确认。得到明确确认后才执行公开发布。版本号、渠道或对外说明有变化，先发修订稿重新确认。
2. 有明确功能更新时，每项独立写一条，直接写用户可以做什么或会看到什么。同一批中常规修复可以省略。只有修复、维护或没有明确功能变化时，中文只写「本次更新包含重要的安全修复和改进」，英文只写「This update includes important security fixes and improvements.」。
3. 将确认稿写入插件 CHANGELOG 对应版本的「更新要点」和「Highlights」，运行 `node scripts/whats-new.mjs`。npm 包元数据、DSH 安装前预览、安装后说明和 GitHub Release 均从该版本的文案生成。发布演练时对照确认稿核对两个语言版本。
4. 在支持更新预览的 DSH 插件中，打开更新入口先显示版本及说明。用户可以关闭面板，点击说明下方的更新按钮才开始安装。说明读取失败时提示重试。已安装 0.0.2 的用户首次过渡升级可在对应版本的 GitHub Release 查看说明。

定版：`packages/dsh-adapter/package.json` 的 `version` 与 CHANGELOG 的「## <版本> — <日期>」节一起改，再 `node scripts/whats-new.mjs`。
只发 `package.json` 当前定的版本；`--tag` 必须是版本的阶段（rc 版本只发 `rc`）；发预览版不会移动 `latest`（要移动见「移动 dist-tag」）。

三条路径，同一个版本只走其中一条：CI 出的包与本机出的包字节不同（至少包里记的源提交不同），一条发过后另一条会被门禁按「同号不同字节」拒绝。

| 路径 | 认证 | 什么时候用 |
|---|---|---|
| A. GitHub Actions + npm trusted publishing | GitHub OIDC，无长期 token | 推荐 |
| B. GitHub Actions + `NPM_TOKEN` secret | granular access token | A 配不了或临时不可用时的后备 |
| C. 本机手动发 | 本机临时 `NPM_TOKEN=…` | 不走 CI 时 |

### 一次性配置

GitHub 仓库（A、B 都要）：Settings → Secrets and variables → Actions → **Secrets** 加 `NPM_GATE_INTERNAL_HOSTS`（组织域名清单，逗号分隔）。
它不是凭据，放 Secret（不放 Variables）只为日志遮蔽：公开仓的 Actions 日志人人可读，Variable 的值会明文出现在每一步日志里，Secret 会被遮成 `***`。
GitHub 只遮整串值，门禁报错时会写出命中的单个域名，所以工作流两个 job 的第一步再把清单里每个域名单独遮蔽（`::add-mask::`）。

A. npm trusted publishing（推荐）：npmjs.com 上包的 Settings → Trusted Publisher → GitHub Actions，填
Organization `officesdk`、Repository `officesdk-plugin-dsh`、Workflow filename `npm-publish.yml`（只填文件名，区分大小写；工作流改名就发不了）、
Environment 留空（**启用工作流里注释掉的 `release` 环境时，publish 与 dist-tag 两个 job 都要启用，这里同时填 `release`，否则 trusted publishing 失败**）、
Allowed actions 勾上 `npm publish`（新建配置默认只允许 `npm stage publish`），用 dist-tag 模式还要勾 **Allow npm dist-tag**
（同一条配置：npm 按文件名 `npm-publish.yml` 匹配，dist-tag 也从这个工作流走）。
配了 trusted publisher 就不要再设 `NPM_TOKEN`：npm CLI 有 OIDC 时先走它，publisher 配错（例如没勾 `npm publish`）时换来的 token 会盖掉
`NPM_TOKEN`，发布被拒且不会回落到 token。注意：

- 新建的配置要在 2 天内完成第一次成功发布，否则过期，只能删掉重建；配好后尽快跑一次正式发布。已验证的这一条同时服务 dist-tag 模式，无需为它另建配置。
- 包里 `package.json` 的 `repository.url` 须与仓库一致：工作流出包时用 `--repository` 取本仓地址，不用另配。
- 需要 npm CLI 11.5.1 及以上、Node 22.14 及以上、GitHub 托管 runner：工作流用 Node 24，发布前检查 npm 版本，不够即失败。
- 首次经 trusted publishing 发布成功后，建议在包的 Settings → Publishing access 选「Require two-factor authentication and disallow tokens」，并吊销旧 token。

B. `NPM_TOKEN`（后备）：npmjs.com → Access Tokens 建 granular access token（只给 `@officesdk/officesdk-plugin-dsh` 读写、勾 bypass two-factor
authentication、设到期日并按期轮换；classic automation token 已停用），在 GitHub 的 Settings → Secrets and variables → Actions → Secrets 加
`NPM_TOKEN`。只在 npmjs 上没有配置 trusted publisher 时用它（见上一条）。两样都没配时工作流的发布 job 失败，不会「绿了但没发」：
publish job 总有 OIDC 环境，不设 `NPM_TOKEN` 就由 npm CLI 做 OIDC 交换，没有 trusted publisher 时 `npm publish` 报认证错误（如 `ENEEDAUTH`）。
发布失败时报错里的 token 会替换成 `***`；没有 token（走 trusted publishing）时报错附上要核对的 trusted publisher 字段。

### A / B：运行工作流

Actions → **npm-publish** → Run workflow（分支选 `main`），填：

- `mode`：`publish`（默认；`dist-tag` 见「移动 dist-tag」）
- `version`：要发的版本（默认 `1.0.1`，须等于 `packages/dsh-adapter/package.json`）
- `tag`：dist-tag（alpha / beta / rc / next / latest，默认 `latest`，须是版本的阶段）
- `dry_run`：默认开。先开着跑一次，确认通过后关掉再跑一次正式发布。第一次演练时顺便看一眼运行日志：
  每一步开头的 env 里 `NPM_GATE_INTERNAL_HOSTS` 显示为 `***`，日志里不出现任何组织域名。

工作流（`.github/workflows/npm-publish.yml`）的 publish 模式分两个 job（dist-tag 模式只跑第三个 job `dist-tag`，见「移动 dist-tag」）：

1. **build**（不带任何凭据）：逐个遮蔽域名（清单没配即失败）→ 检查输入（版本格式、等于 package.json、whats-new.json 有该节、tag 是版本的阶段、正式发布只从 `main`）→
   `npm ci --prefix tooling --ignore-scripts` → `vendor.mjs fetch` → `npm run check` → 出 public 包（`--repository` 为本仓）→ 记字节与 sha256 →
   `verify-universal-package` → `npm-gate --target npmjs` → `npm-release publish --dry-run`（结果 `artifacts/npm-dry-run.json`，写进 job summary）→
   把 tgz 交给 publish job（artifact `npm-package`，保留 7 天）。
2. **publish**（只在 `dry_run` 关闭时跑；权限 `contents: read` + `id-token: write`）：逐个遮蔽域名 → 检查 npm 版本 → `vendor.mjs fetch` / `check`
   （门禁比对的 `kit-manifest.json` 由本 job 按钉住的 sha256 自己取）→ 取 build 的 tgz 并核 sha256 →
   `npm-release publish`（以 npmjs 目标再过一次门禁；全工作流只有这一步引用 `secrets.NPM_TOKEN`，没设就走 trusted publishing）→
   上传发布记录 artifact `npm-publish`（`artifacts/npm-publish.json` 与 `artifacts/npm-release-result.json`）。
   结果只认 `published` 与 `already-published`（同号同字节已在 npm 上，job summary 会醒目标出）；`npm publish` 报错（如 trusted publisher 没配好）
   或其他结果 job 都失败。

### C：本机手动发

在仓库根目录出包、先演练、再发布（以 0.0.3 为例；组织域名清单照上节给）：

```sh
node tooling/suite/pack-universal.mjs --version 0.0.3 --sources public --obfuscate light \
  --repository https://github.com/officesdk/officesdk-plugin-dsh --out artifacts/npm
node scripts/npm-release.mjs publish --version 0.0.3 --tag latest --dry-run   # 不要 token、不发布，跑同样的选包与门禁
NPM_TOKEN=<你的 token> node scripts/npm-release.mjs publish --version 0.0.3 --tag latest
```

token 只经环境变量传入，写在用完即删的临时 userconfig 里；没有 token（也不在 GitHub Actions 的 OIDC 环境里）时只打 WARNING、什么都不发、退出码 0。
包的 Publishing access 改成 disallow tokens 后，本机用 token 发布（这一条与「回滚」的 `publish --rollback`）会被 npm 拒绝：
需要时在 Publishing access 临时放开、发完改回（工作流只发 `--version` 定版的构建，不发回滚包）。

### 公开仓的安全约束

仓库公开后任何人都能看工作流文件和运行日志。改动 `npm-publish.yml` 与 `github-release.yml` 时保持：

- 只 `workflow_dispatch` 手动触发；绝不加 `pull_request` / `pull_request_target`，不在发布工作流里跑来自 PR 或 fork 的代码。
- 顶层权限只 `contents: read`；只有 publish 与 dist-tag 两个 job 加 `id-token: write`；`github-release.yml` 只有它的 job 有 `contents: write`（建 tag 与 Release），`GH_TOKEN` 只在建 Release 那一步。
- build job 不带凭据；`NPM_TOKEN` 只在 publish job 的发布那一步引用；dist-tag job 与 `github-release.yml` 不引用任何 secret；装依赖一律 `--ignore-scripts`，publish job 不装依赖，dist-tag job 只装钉版的 npm 11.21.3（核 lock integrity）。
- dist-tag job 与发布用同一条 trusted publisher，换到的令牌按这条的 Allowed actions 也能发布：它和 publish job 一样不跑第三方代码，只有钉 SHA 的 action、仓库脚本和核过 integrity 的 npm。
- 构建与发布分 job，publish 只发 build 交来的 tgz。核 sha256 防的是传错或拿错包，不防 build job 被攻破（那时 tgz 与 sha256 可一起被改）；
  真正的防线是 publish job 不跑第三方代码、用自己取的 kit-manifest 以 npmjs 目标重跑门禁。
- 两个 job 第一步逐个遮蔽组织域名；job output 里不放域名或清单片段（含遮蔽串的 output 会被 GitHub 丢弃）。
- 第三方 action 一律钉 commit SHA（行尾注释版本；升级时 SHA 与注释一起改）；不用缓存；checkout 不保留 git 凭据（`persist-credentials: false`）。
- inputs 只经 env 进脚本，run 脚本里不写 `${{ }}`。
- 仓库计划支持 environment 后，启用 publish 与 dist-tag 两个 job 上注释掉的 `environment: release` 并设 required reviewers，同时在 npm trusted publisher 填同名 Environment。

仓库翻成公开前核对：

1. 私有期的 Actions 运行日志与 artifact 翻转后对外可见：先在 Actions 页审查，删除不想公开的运行记录。
2. `NPM_GATE_INTERNAL_HOSTS` 在 **Secrets**（不在 Variables）；仓库与包内无组织域名、无 token（门禁与全树扫描）。
3. 认证只留一种：trusted publisher 配好并首发成功后不设 `NPM_TOKEN`，包的 Publishing access 改为 disallow tokens，吊销旧 token。
4. 公开后第一次发布检查 npm 包页出现 provenance。
5. 每个 `v*` tag 钉住一份完整的旧仓库快照（main 每次重建，tag 仍指向当时的提交）：翻转前审查 tag 与 Release，不想公开的先删除。

## 移动 dist-tag（OIDC，免 token）

`node scripts/npm-release.mjs dist-tag --version <版本> --tag <dist-tag> [--dry-run]` 把一个 dist-tag 指向 npm 上**已发布**的版本，不重新发布。
典型用法：首个正式版之前把 `latest` 指向预览版（如 0.0.1-rc.1），让不带 dist-tag 的 `add @officesdk/officesdk-plugin-dsh` 装到它。

规则（`distTagProblem`）：`next` 只收 `X.Y.Z-test.N`、测试版只进 `next`；`alpha` / `beta` / `rc` 只收本阶段或更成熟的版本（rc 用户不会收到 alpha）；
`latest` 收正式版或任一预览版。版本必须已在 registry；tag 已是该版本时什么都不做（`unchanged`）；tag 现指向更新的版本时拒绝往回挪
（已更新的用户不会降级，退回旧版用「回滚」）。

### 配置

不用新配置：dist-tag 和发布用同一条 trusted publisher（Workflow filename `npm-publish.yml`），只要它的 Allowed actions 勾了 **Allow npm dist-tag**。
npm 按工作流文件名匹配，所以 dist-tag 也只能从 `npm-publish.yml` 走（工作流改名，发布与 dist-tag 都会断）。
`npm dist-tag` 从 npm 11.21.0 起（npm 12 为 12.2.0）才用 OIDC 换令牌，Node 自带的 npm 不够：dist-tag job 自己把 npm 11.21.3 装到独立目录
（`--ignore-scripts`，核对 lock 里的 integrity 与钉值一致）。不用 `NPM_TOKEN`：包的 Publishing access 要求 2FA 且禁止 token。

### 运行

Actions → **npm-publish** → Run workflow（分支选 `main`）：

- `mode`：`dist-tag`
- `version`：npm 上已有的版本（如 `0.0.1-rc.1`）
- `tag`：要指向它的 dist-tag。**挪 latest 要选 `latest`**：默认的 `rc` 若已指向该版本，只会得到 `unchanged`
- `dry_run`：先开着跑一次，确认后关掉再跑一次正式移动

命令行也可以：`gh workflow run npm-publish.yml --ref main -f mode=dist-tag -f version=0.0.1-rc.1 -f tag=latest -f dry_run=true`（正式时 `dry_run=false`）。
网页表单里 boolean 输入是同名的隐藏 input 与 checkbox 两个元素：用脚本或浏览器自动化读它的值，要读**可见 checkbox 的 `.checked`**。

dist-tag 模式只跑 dist-tag job（build、publish 跳过；权限 `contents: read` + `id-token: write`，不引用任何 secret）：
检查输入（上面的规则；正式移动只从 `main`）→ 装 npm 11.21.3 → `npm-release.mjs dist-tag`：

- 演练：读 registry 当前的 dist-tags、跑 `npm dist-tag ls`，job summary 写出 `from` → 版本、`change`（`move` / `none`）与 OIDC 交换结果。
  `npm dist-tag ls` 会先做一次 OIDC 交换（只读），所以演练就能看出 trusted publisher 认不认这个工作流：交换失败（`failed: …`）演练即失败；
  交换成功只说明身份对上了，是否勾了 Allow npm dist-tag 要到正式移动才知道。
- 正式：`npm dist-tag add @officesdk/officesdk-plugin-dsh@<版本> <tag>`，再重读 registry 的 dist-tags 核对（最多 5 次、间隔 2 秒），结果 `moved`。

结果只写进 job summary（不上传 artifact）。失败排查：

| npm 报错 | 含义 |
|---|---|
| `E401` / `ENEEDAUTH` | trusted publishing 没认上：核对 trusted publisher 的 Workflow filename `npm-publish.yml`、仓库、Environment（留空；启用 `release` 环境时 publish 与 dist-tag 两个 job 都要启用）；npm 须 11.21.0 及以上 |
| `E403` | 这条 trusted publisher 没勾 Allow npm dist-tag |
| `EOTP` | npm 要一次性密码：trusted publishing 不应需要，检查包的 Publishing access 与 trusted publisher |

本机手动：用自己的 npm 登录（会要 2FA 一次性密码）`npm dist-tag add @officesdk/officesdk-plugin-dsh@<版本> <tag>`。包禁止 token，
`NPM_TOKEN=… node scripts/npm-release.mjs dist-tag …` 只适用于演练用的本地 registry；既无 token 又不在有 OIDC 的 GitHub Actions job 里时，脚本只打 WARNING、不改动、退出码 0。

注意：

- `latest` 挪到预览版后，再发下一个预览版（如 rc.2，`tag` 为 `rc`）时 `latest` 不会跟着走，要再跑一次 dist-tag 模式；首个正式版以 `latest` 发布后由发布接管。
- dist-tag 与发布共用 concurrency 组 `npm-publish`：同一时刻只跑一个，已在排队的运行会被后排进来的顶掉（取消），先后点两次时留意。

## GitHub Release（关联 npm 版本）

`.github/workflows/github-release.yml` 给 npm 上已发布的版本建 GitHub Release `v<版本>`，附上 npm 发布的**同一个** tarball，说明里链接 npm 版本页。
它不碰 npm 认证：只要 `contents: write` 和内置的 `GITHUB_TOKEN`（不用个人 token），不引用任何 secret。

Actions → **github-release** → Run workflow（分支选 `main`）：`version`（如 `0.0.1-rc.1`）、`dry_run`（先开后关）。命令行：
`gh workflow run github-release.yml --ref main -f version=0.0.1-rc.1 -f dry_run=true`。

步骤：检查输入（非测试版的版本号；正式只从 `main`）→ `npm pack @officesdk/officesdk-plugin-dsh@<版本>` 取 npm 上的包 →
`npm-release.mjs release-notes`（核对包名、版本与 registry 的 integrity，不符即失败；生成说明：npm 版本页
`https://www.npmjs.com/package/@officesdk/officesdk-plugin-dsh/v/<版本>`、安装命令、附件 integrity、tag 说明、包内 CHANGELOG 该版本节）→
建 Release（全工作流只有这一步拿 `GH_TOKEN`）：

- 同名 Release 已存在且带同名 tgz：提示后跳过，不改动；是 draft（多半是上次中途失败留下的）或缺附件：失败并说明怎么处理（删掉 draft / 补传附件后重跑）。
- tag `v<版本>` 已存在：沿用它（`--verify-tag`）；不存在：打在本次运行的提交（`main` 当前的 HEAD）上。只有确认不存在（404）才新建，其他读取错误一律失败。
- 预览版（alpha / beta / rc）标为 Pre-release，正式版不标。
- 演练只打印将要执行的 `gh release create` 命令与说明，不建 tag、不建 Release。结果只写进 job summary。

tag 与 main 重建：main 每次发布都重建为单一无父提交。`v<版本>` 钉住建 Release 时 main 的那个提交，成为不可变快照；之后 main 再重建，
老 tag 指向的提交不在 main 上了，但由 tag 保住、不会丢。每个版本一个 Release 的用法不受影响。附件是 npm 上发布的字节（integrity 为证），
tag 是建 Release 时的仓库快照，不一定是构建该版本的那份源码。

组织策略把 `GITHUB_TOKEN` 限成只读时，建 Release 会报 403（Resource not accessible by integration）：到 Settings → Actions → General → Workflow permissions 检查。

## 只从 npm 更新

插件只从 npm 自更新：没有 git 安装、换仓库或切换渠道的逻辑。用别的方式装的副本（本地 `.tgz`、路径）不会自更新，按包名重装即可。

## 回滚

`node scripts/npm-release.mjs rollback --to <旧版本> --out artifacts/npm-rollback.tgz`，再
`NPM_TOKEN=<token> node scripts/npm-release.mjs publish --rollback --package artifacts/npm-rollback.tgz --tag <阶段>`：把旧版本的文件按其阶段的
下一个号重新发布（正式版为下一个补丁号），所有人经一键更新拿到；office-kit 低于当前兼容下限、或新号低于更成熟阶段已发版本时拒绝。
单台机器：`dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh@<旧版本>`。

回滚目标只能是 office-kit 0.3.0-rc.1 布局以后的版本：0.0.0-alpha.3 及更早的包从已停止对外的旧下载地址取 Kit 与编辑器，下载档位也不是现在门禁认可的公开来源，`publish --rollback` 会被门禁拒绝；要退回旧引擎，用新代码配旧 kit-manifest 另出一个版本。

## 版本与更新信息（公网读法）

更新检查读 npm packument 的 dist-tags，不再 `git ls-remote`。`capabilities.json`、`whats-new.json`、`CHANGELOG.md`、`kit-manifest.json`、
`skills-catalog.json` 都在包里，公网可读：

- `https://registry.npmjs.org/@officesdk/officesdk-plugin-dsh/latest`（或 `/next`、`/<版本>`）：package.json 元数据
- `https://unpkg.com/@officesdk/officesdk-plugin-dsh@latest/capabilities.json`、`https://registry.npmmirror.com/@officesdk/officesdk-plugin-dsh/<版本>/files/capabilities.json`

`capabilities.json` 的 `install` 跟随出包渠道（`scripts/capabilities.mjs` `installFields`）：`installVia: npm`，`install` 每个渠道一条（`stable` `@latest`、`test` `@next`、`alpha` / `beta` / `rc` 同名 dist-tag），`registry` 为这次构建的 registry（`--registry` 给的地址，否则 npmjs），`npm.status: enabled`。

## 混淆

production minify 一直在做（host：空白与语法；client、kit-page：完整 minify）。npm 构建再加一层 `javascript-obfuscator` 低强度
（`tooling/suite/obfuscate.mjs`：改局部标识符、字符串移入打乱的数组但**不编码**；不做控制流平坦化、死代码、自卫、反调试）。
`lib/host/index.mjs` 不混淆：它的 Remote 服务类方法会被改成计算键，而 Typert 靠源码文本找 Remote 方法和参数名。
实测（0.1.1-test.7）：tgz 436 KB → 494 KB（+13%）；client 冷启动约 +2 ms；错误栈函数名保留（esbuild keepNames），消息不变。
这只提高阅读成本，不是加密。


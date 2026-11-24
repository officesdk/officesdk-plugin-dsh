# OfficeSDK 总体说明：项目、架构与发布流程

面向研发与维护者，是本仓库的入口。这里只讲全局，细节在下表列的专题文档里。

| 想看 | 去哪里 |
|---|---|
| 构建、测试、出包 | [BUILD.md](BUILD.md) |
| npm 渠道、两种构建、上架门禁、回滚 | [NPM.md](NPM.md) |
| `capabilities.json` 怎么生成、字段说明 | [capabilities.md](capabilities.md) |
| office-kit（外部组件）用哪一版、怎么取 | [../components/README.md](../components/README.md) |
| 更新日志 | [../packages/dsh-adapter/CHANGELOG.md](../packages/dsh-adapter/CHANGELOG.md) |

## 1. 项目是什么

OfficeSDK（npm 包 `@officesdk/officesdk-plugin-dsh`，DSH 插件列表里显示为 OfficeSDK）是 DeepSeek Harness（DSH）的插件。装上以后，DSH
右侧的文件标签页可以直接打开、编辑、保存 Word（.docx）、Excel（.xlsx）、PPT（.pptx），模型也可以通过插件的 AI 工具在编辑器里实时新建和
修改这三类文件。插件还带一个「灵感库」：在 DSH 左侧导航里按 Word / Excel / PPT 浏览模板，选好后带入对话框。

- **一个通用的纯 JS 包**：Mac 和 Windows 装同一个包（不到 1 MB），里面没有可执行文件。
- **引擎首启下载**：第一次打开文档时，按平台下载 office-kit 的 Kit（文档服务）和编辑器资源，之后不再下载。
- **本机完成**：打开、编辑、保存都在用户电脑上进行，不需要任何服务端。
- **office-kit 是外部组件**：Kit 与编辑器页面属于外部组件 office-kit，不在本仓库里（见 [../NOTICE](../NOTICE)）。

DSH 版本要求写在 `packages/dsh-adapter/package.json` 的 `engines.dsh`（现为 `>=0.2.1-rc.3 <0.3.0-0`）；平台成熟度写在
`packages/dsh-adapter/capabilities.source.json`（现在 darwin-arm64 为 `stable`，win32-x64 为 `test`）。

## 2. 架构

### 2.1 交付物与来源

```text
            ┌──────────────────────────────────────────────────────┐
  安装/更新 │ 插件主包 @officesdk/officesdk-plugin-dsh（npm，纯 JS）  │
            │  lib/host · lib/client · lib/tools · skills/          │
            │  resources/editor（office-kit 编辑器页面，钉住版本）    │
            │  kit-manifest.json · download-sources.json            │
            │  capabilities.json · whats-new.json · CHANGELOG        │
            └───────────────┬──────────────────────────────────────┘
                            │ 首次使用时按 download-sources.json 依次尝试，按清单 sha512 / 逐文件 hash 校验
                ┌───────────┴───────────┐
                ▼                       ▼
          Kit（按平台）            editor 资源
          office-kit GitHub 发布页  office-kit npm 包
                └───────────┬───────────┘
                 存到 ~/.dsh/plugins/officesdk/office-kit/，多个窗口和 profile 共用
```

| 部分 | 是什么 | 来自 |
|---|---|---|
| 插件主包 | Host（Node 侧服务）、Client（DSH 页面里的 React 组件）、AI 工具、skill | npm `@officesdk/officesdk-plugin-dsh`（本仓库构建） |
| Kit | office-kit 官方各平台 Kit 归档（文档服务） | office-kit 的 GitHub 发布页（officesdk/officesdk-core）；不走 npm |
| editor | office-kit 编辑器资源；页面包 `offline-editor.js` 与表格图标字体随主包 | 随主包的部分按 `components/office-kit.json` 从钉住的公开发布取得；其余首次使用时从 npm（`@officesdk/officesdk-core-editor`）下载 |
| skills | `shimo-pptx`（PPT，蓝色品牌风格 `default`）、`shimo-docx`（Word 公文版式 `gongwen`） | 本仓库 `packages/dsh-adapter/skills/`，随主包 |

**下载来源可配置**：顺序与超时是包根的数据文件 `download-sources.json`，改它不用改代码。`public` 档位（npm 公网包）的 Kit
从 GitHub 发布页、editor 从 npm 下载，skill 随包自带；这是唯一的档位。环境变量可覆盖：`OFFICESDK_KIT_MIRROR`、`OFFICESDK_EDITOR_NPM`、
`OFFICESDK_SKILLS_BASE`。不管从哪个来源下载，都以包内 `kit-manifest.json` 的 sha512 与逐文件 hash 为准。

### 2.2 代码结构

```text
packages/dsh-adapter/        插件本身（打包后就是 @officesdk/officesdk-plugin-dsh）
  src/host/                  Node 侧：文档会话与 Kit 进程、下载器（fetch/）、一键更新（updater、npm）、版本与渠道（version）、
                             样式目录与推荐（style-*）、模板（template-pick）、本次更新（whats-new）
  src/client/                DSH 页面侧：编辑器标签页（plugin、controller、toolbar）、灵感库（inspiration*、template-covers）、
                             风格选择与带入（style-picker、style-flows、style-cards）、AI 实时生成视图（agent、tool-views）、选中文字改写（selection-bridge）
  src/tools/                 模型可调用的 AI 工具：presentation_* / document_* / spreadsheet_*
  src/skills/                skill 清单、下载与缓存（catalog、store）
  skills/                    随包 skill 源：shimo-pptx（styles/default）、shimo-docx（templates/gongwen.docx）
  capabilities.source.json   能力与平台状态的手写数据源（生成 capabilities.json）
  download-sources.json      下载来源与超时
packages/core/               与宿主无关的核心：AI 编辑模型、Kit 编辑器页面与浏览器侧代理
packages/kit-host/           启动并驱动 office-kit 的 Kit 进程、文档视图与 HTTP 路由
components/                  office-kit 钉住版本（office-kit.json）
tooling/kit/                 按钉住版本取 office-kit 文件（vendor.mjs）
tooling/suite/               打包（pack-universal.mjs）、混淆、依赖分层检查
tooling/style-distill/       样式包蒸馏工具
scripts/                     门禁（npm-gate）、发布（npm-release）、能力与更新要点生成、包校验
```

### 2.3 客户端功能模块

| 模块 | 入口 | 要点 |
|---|---|---|
| 编辑器标签页 | DSH 文件列表或对话里的文件链接 | 编辑 / 阅读模式；保存到原文件前需授权；另存为；Word 可打印；文件被外部改写时询问用哪一版 |
| AI 实时生成 | 模型调用 `*_create` / `*_edit` | 右侧编辑器边生成边看；可停止；PPT 可撤销 AI 修改 |
| 灵感库 | DSH 左侧导航「灵感库」 | 整页浏览 Word / Excel / PPT 模板，「添加并使用」后回到对话框，用 DSH 原发送按钮开始 |
| 样式与 skill | 对话框里的风格选择；工具栏「⋯ → Skills 管理」 | 可预下载、检查更新、清除缓存 |
| 一键更新与本次更新 | 工具栏「插件更新」 | 按渠道查 npm dist-tag；更新后第一次打开文档弹「本次更新」要点（`whats-new.json`） |

## 3. 发布流程

### 3.1 版本与渠道

| 阶段 | 版本号 | npm dist-tag |
|---|---|---|
| 预览 | `X.Y.Z-alpha.N` → `X.Y.Z-beta.N` → `X.Y.Z-rc.N` | `alpha` / `beta` / `rc` |
| 正式 | `X.Y.Z` | `latest` |
| （旧）测试版 | `X.Y.Z-test.N` | `next`：旧测试渠道，不再发布 |

- 版本来自源码 `packages/dsh-adapter/package.json`，同时在 `CHANGELOG.md` 写 `## <版本> — <日期>` 一节（可带「### 更新要点」），
  再运行 `node scripts/whats-new.mjs` 重新生成 `whats-new.json` 与 `capabilities.json`。
- 规则在 `packages/dsh-adapter/src/host/version.mjs`：同一 X.Y.Z 里 alpha < beta < rc < 正式版；dist-tag 等于阶段，门禁会核对。

### 3.2 一键更新与回滚

- **一键更新**：插件按安装时的渠道查 npm dist-tag，取「本阶段及更成熟阶段」里最新的一版，再执行
  `dsh plugin add @officesdk/officesdk-plugin-dsh@<新版本>`，DSH 重新组合加载，通常不用重启。registry：设了 `OFFICESDK_NPM_REGISTRY` 就用它，
  否则用 profile `.npmrc` 里的 `@officesdk:registry`（或 `registry`），都没有才用 npmjs，连不上再 npmmirror。
- **回滚**：见 [NPM.md](NPM.md)「回滚」。

### 3.3 发一个版本

1. 改 `package.json` 版本、写 CHANGELOG 版本节、`node scripts/whats-new.mjs`；跑测试、出包、门禁（[BUILD.md](BUILD.md)）。
2. `scripts/npm-release.mjs publish` 发布到 npm（经门禁）：通常在 GitHub Actions 手动运行 `npm-publish` 工作流（先演练、再正式发；
   认证用 npm trusted publishing，`NPM_TOKEN` secret 为后备），也可在本机发（[NPM.md](NPM.md)「发布」）。
3. 冒烟后推进阶段：Mac 与 Windows 各走一遍（首次下载引擎、三类文档各开一个、一键更新）。
4. 按需：用 `npm-publish` 工作流的 dist-tag 模式移动 dist-tag（如首个正式版之前把 `latest` 指向预览版），用 `github-release` 工作流给该版本建
   GitHub Release、附 npm 上的同一个包（[NPM.md](NPM.md)「移动 dist-tag」「GitHub Release」）。

## 4. 对外信息

每个包里都带 `capabilities.json`（字段见 [capabilities.md](capabilities.md)）、`whats-new.json`、`CHANGELOG.md`、`kit-manifest.json`。

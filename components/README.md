# office-kit（外部组件）

OfficeSDK 插件打开、编辑和保存 Word / Excel / PPT 靠 office-kit：它的文档引擎（Kit）和编辑器页面是独立的外部组件，
**不在本仓库里**（见根目录 [NOTICE](../NOTICE)）。本仓库只记录用哪一版、从哪里取：

| 文件 | 进 git | 内容 |
|---|---|---|
| `office-kit.json` | 是 | 钉住的 office-kit 版本和插件能运行 Kit 的平台；构建时要取的三个文件各自来自哪个公开来源（`sources`：npm 包带 tarball 的 sha512，或 GitHub 发布页目录带 `SHA256SUMS`）、在来源里的路径和 sha256 |
| `vendor/editor/bundle/offline-editor.js` | 否 | office-kit 编辑器页面包（`@officesdk/officesdk-core-editor` 的 `dist/bundle/offline-editor.js`，K1/K16/K18 已在其中），随插件包发出 |
| `vendor/host-fonts/sheet-host-fonts.css` | 否 | 表格图标字体，随插件包发出（0.3.0-rc.1 起页面包已自带，Host 仍注入、无害；暂取自插件已发布的 0.0.0-alpha.3） |
| `kit-manifest.json` | 否 | office-kit 官方各平台 Kit 归档与编辑器资源的大小与 hash；插件第一次打开文档时按它下载并逐文件校验：Kit 来自 office-kit 的 GitHub 发布页，编辑器资源来自 npm |

## 取文件

```sh
node tooling/kit/vendor.mjs fetch   # 按 office-kit.json 从各自来源取三个文件，校验 sha512 / sha256 / SHA256SUMS 后写到上表位置
node tooling/kit/vendor.mjs check   # 只核对磁盘上的文件就是钉住的那几个（并核对页面包就是 kit-manifest 的 editor.bundle）
node tooling/kit/kit-smoke.mjs --data <目录>   # 像插件首开那样下载本机 Kit 与编辑器资源、对 SHA256SUMS、跑 Kit 自检
```

`tooling/suite/pack-universal.mjs` 在 `check` 不过时拒绝出包。下载的来源缓存在 `.vendor-cache/`（git 忽略）；本机没有网络时
可以把同样命名的文件放进一个目录，用 `OFFICESDK_VENDOR_DIR=<目录>` 指给 `fetch`（照样校验）。

## 升级 office-kit

office-kit 的新版本（及其编辑器补丁）由 office-kit 维护方构建发布。拿到新的公开发布后：

1. 改 `office-kit.json`：`officeKit.version` / `platforms`（kit-manifest 中带 `binary` 的平台），`sources` 里各来源的版本、
   GitHub 发布页地址与 npm 的 `integrity`（`npm view <包>@<版本> dist.integrity`），三个文件的 `sha256`（kit-manifest.json 的
   sha256 对照发布页 `SHA256SUMS`；页面包的 sha256 等于 kit-manifest 的 `editor.bundle.sha256`）；
2. `node tooling/kit/vendor.mjs fetch`，再 `node scripts/whats-new.mjs` 重新生成 `capabilities.json`；
3. 跑测试并出包验证（[docs/BUILD.md](../docs/BUILD.md)）。

插件只调用 office-kit 公开的接口；需要的新能力向 office-kit 提需求，不在本仓库里改它的代码。

# Build and verify

Everything here runs from this repository alone, on macOS, Linux or Windows with Node 22.19 or later (Node 24 is what
the package is tested with).

```sh
npm ci --prefix tooling --ignore-scripts        # esbuild and javascript-obfuscator for the build
node tooling/kit/vendor.mjs fetch               # office-kit's editor bundle, host fonts and kit-manifest.json (components/README.md)
node tooling/suite/pack-universal.mjs --version 0.0.3 --sources public --obfuscate light \
  --repository https://github.com/officesdk/officesdk-plugin-dsh --out artifacts/npm
node scripts/verify-universal-package.mjs artifacts/npm/officesdk-officesdk-plugin-dsh-0.0.3.tgz
node scripts/npm-gate.mjs artifacts/npm/officesdk-officesdk-plugin-dsh-0.0.3.tgz --target ci
```

`--repository <url>` (or `OFFICESDK_PUBLIC_REPOSITORY`) names the public repository in the package's `package.json`.

The gate checks that the package names no private address and no host of the publisher's own organisation. That host
list is configuration, never committed: set `NPM_GATE_INTERNAL_HOSTS` (comma-separated domains) or write
`scripts/npm-gate-internal.local.json` (`{ "hosts": ["corp.example"] }`, git-ignored). Without one, `--target npmjs`
refuses and `ci`/`local` warn.

## The universal package

One `@officesdk/officesdk-plugin-dsh` serves every platform (spec 261110-universal-package). It is pure JS: no Kit
executables, no office-kit editor resources, no `os`/`cpu`. It carries `kit-manifest.json` (office-kit's official Kit
archive per platform and its editor resources: size, sha256, sha512, every file's hash) and downloads both on first use
into `~/.dsh/plugins/officesdk/office-kit/`, from `OFFICESDK_KIT_MIRROR` when set, then office-kit's public release: the Kit
from its GitHub release (`release`: the manifest's own `sources`), the editor resources from its npm package (`npm`:
the manifest's `editor.npm`; the GitHub release does not carry them).
The first open of a document therefore needs the network. When no source can be reached the document tab shows the
download failure with a retry. Whatever the source, the manifest's sha512 and file hashes decide. The source order and
limits are data: `download-sources.json` at the package root (`public` for the npm build; see docs/NPM.md).

The office-kit editor page (`resources/editor/bundle/offline-editor.js`, `resources/host-fonts`) ships in the package;
it comes from the pinned office-kit sources (`components/office-kit.json`), not from this repository.

`node tooling/kit/kit-smoke.mjs --data <dir>` downloads this machine's Kit and the editor resources the same way the
plugin does, checks the manifest against the release's `SHA256SUMS` and runs the Kit's self-test; it prints
`kitDir` for `OFFICESDK_KIT_DIR`. `.github/workflows/windows-kit-smoke.yml` runs it on a GitHub-hosted Windows x64 runner
(by hand; win32-arm64 has no hosted runner and is not covered).

Office documents open through office-kit's Kit only. The manifest lists every platform office-kit releases a runnable
Kit for (`KIT_TRIPLES` in `scripts/platform-package.mjs`); a platform not listed in `capabilities.source.json` is
`status: test`. On a platform without a Kit, opening a document says so plainly (`本平台暂不支持…`).

The package carries its skills complete: the repository holds exactly the PPT `default` style (the blue house style)
and the Word `gongwen` template (`scripts/public-skills.mjs` checks the build against that list). `public` is the only
download profile.

## Tests

```sh
node --test packages/core/test/*.test.mjs packages/kit-host/test/*.test.mjs scripts/test/*.test.mjs tooling/kit/test/*.test.mjs tooling/suite/test/*.test.mjs
node --test packages/dsh-adapter/test/*.test.mjs   # needs DeepSeek Harness's @deepseek-ai/* packages (below)
node scripts/whats-new.mjs --check
node tooling/suite/check-deps.mjs
```

The adapter's host and client tests load DeepSeek Harness packages (`@deepseek-ai/cordis`, `dsh-typert-protocol`,
`dsh-atomic-write`, …) that are peer dependencies provided by DSH at run time. Point a `node_modules` at the repository
root to an installed DSH's packages to run them. `packages/kit-host/test/real-kit.test.mjs` runs against a real Kit
when `OFFICESDK_KIT_DIR` names an unpacked office-kit Kit release; it is skipped otherwise.

## Releases

Versions come from `packages/dsh-adapter/package.json`, with a `## <version> — <date>` section in
`packages/dsh-adapter/CHANGELOG.md`, then `node scripts/whats-new.mjs`. Versions are `X.Y.Z-alpha.N` → `X.Y.Z-beta.N`
→ `X.Y.Z-rc.N` → `X.Y.Z`; the npm dist-tag is the stage (`alpha`, `beta`, `rc`; `latest` for a release),
`packages/dsh-adapter/src/host/version.mjs`. Publishing goes through the gate: `scripts/npm-release.mjs publish`,
run by the `npm-publish` GitHub Actions workflow (`.github/workflows/npm-publish.yml`: run by hand, a dry run first;
npm trusted publishing, or the NPM_TOKEN secret as a fallback) or by hand (docs/NPM.md). The same workflow's mode
`dist-tag` points a dist-tag at a published version, and the `github-release` workflow (`.github/workflows/github-release.yml`)
makes its GitHub Release with the npm package attached (docs/NPM.md). Before a stage moves on:
smoke-test on a Mac and on Windows (first download of the engine, open a document of each kind, one-click update).

## Install in DSH

```text
dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh@rc               preview (rc; alpha and beta the same way)
dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh@<version>        a pinned version
dsh plugin --profile desktop add /path/to/officesdk-officesdk-plugin-dsh-<V>.tgz  a local build
```

The plugin updates itself from npm only: the panel's update checks the channel's dist-tag (`alpha`, `beta`, `rc`,
`next`, `latest`) on the registry the profile uses and installs that exact version with `dsh plugin add`
(`src/host/updater.mjs`, `src/host/npm.mjs`). A copy installed any other way (a local `.tgz`, a path) does not update
itself: reinstall it by package name.

Host output lives at lib/host/index.mjs to retain its resource root relative to import.meta.url. Typert looks up Remote
arguments by source parameter names: workspaceFileScope, path, revision, signal and id must stay unmangled. Property
names are protocol/SDK contracts and are not mangled. Legal notices remain in the bundles. The light obfuscation of the
npm build raises analysis cost; it does not encrypt shipped code or make it impossible to inspect.

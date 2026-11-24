# OfficeSDK for DeepSeek Harness

OfficeSDK for DeepSeek Harness (`@officesdk/officesdk-plugin-dsh`) opens, edits and
saves Word (.docx), Excel (.xlsx) and PowerPoint (.pptx) files in the host’s right-side file tab, lets the model create and edit
those files live through AI tools, with bundled document and presentation skills. Document rendering runs on the user's
machine; no server is needed.

Documents open through [office-kit](NOTICE), a separate external component: its document engine (Kit) and editor
resources download on first use from office-kit's public release, and its editor page ships in the package from a
pinned release. office-kit is not part of this repository.

> 中文说明：[docs/OVERVIEW.md](docs/OVERVIEW.md)（项目、架构与发布流程）。

## Install

[OfficeSDK website](https://officesdk.com/) · [Installation guide](https://officesdk.com/en/install/deepseek-harness) · [中文安装指南](https://officesdk.com/zh/install/deepseek-harness)

In DeepSeek Harness Desktop, open **Plugins → Add plugin** and enter `@officesdk/officesdk-plugin-dsh`. To install from a terminal, use the desktop app’s **Manage dsh Command…** menu first. The command uses the bundled runtime. The package requires Node.js 22.19.0 or later. See the [current package metadata](https://registry.npmjs.org/@officesdk/officesdk-plugin-dsh/latest) for compatible host versions.

```sh
dsh plugin --profile desktop add @officesdk/officesdk-plugin-dsh   # quit the desktop app first
```

It updates itself from the plugin panel. The first open of a document downloads office-kit's engine for this platform
(darwin-arm64 is stable, win32-x64 is in test).

## Build

```sh
npm ci --prefix tooling --ignore-scripts
node tooling/kit/vendor.mjs fetch
node tooling/suite/pack-universal.mjs --version 0.0.1 --sources public --obfuscate light \
  --repository https://github.com/officesdk/officesdk-plugin-dsh --out artifacts/npm
```

Details, tests and releases: [docs/BUILD.md](docs/BUILD.md), [docs/NPM.md](docs/NPM.md). Releases to npm run from the
`npm-publish` GitHub Actions workflow (run by hand, npm trusted publishing), which also moves dist-tags (mode `dist-tag`);
the `github-release` workflow makes the GitHub Release of a published version, with its npm package attached.

## Repository layout

| Path | What |
|---|---|
| `packages/dsh-adapter/` | the plugin: Host service, DSH client, AI tools, bundled skills |
| `packages/core/` | host-neutral core: AI edit model, the Kit editor page and its in-browser agents |
| `packages/kit-host/` | runs office-kit's Kit process and serves its document views |
| `components/office-kit.json` | the pinned office-kit release ([components/README.md](components/README.md)) |
| `tooling/`, `scripts/` | packaging, the npm release gate, capabilities and release notes |

## Model editing tools

The `officesdk/tools` component (「OfficeSDK · 模型编辑工具」 on the plugin page) lets the model create, read, edit and save PPTX, DOCX and XLSX files; it needs the `officesdk` component enabled. The plugin page keeps its text short; the details are here:

- Every change appears step by step in the OfficeSDK tab.
- Edits to an existing file apply only after you confirm. A file the model creates is shown as it is built, and you can stop it at any time.
- PPTX edits can be undone. Word and Excel edits cannot be undone automatically: steps finished before a stop or an error stay.
- The result can be saved back to the original file.

## Bundled skills

The `officesdk/skills` component (「OfficeSDK · 默认风格 skill」 on the plugin page) registers two skills with DSH. The plugin page shows a one-line summary from `locale/skills/*.json`; the model chooses a skill by the `description` in its `SKILL.md`, so trimming the plugin page text does not change which skill the model picks:

- `shimo-pptx` — new decks. By default it builds the deck live in the right-side editor (`presentation_create`, then `presentation_edit` in batches of up to four new slides that are animated together). It uses the offline script `scripts/brand_deck.py` instead when the deck needs tables or charts (live tables and charts cannot be saved yet, office-kit K6), when several decks are wanted, when there is no right-side editor (headless, command line), or when the user asks for script output or pixel parity with brand-ppt.
  Pictures: the user's own, or workspace pictures that fit the topic, come first. Otherwise `scripts/find_images.py` searches Openverse and Wikimedia Commons (no API key; CC0, public domain, CC BY and CC BY-SA only), downloads into the task's `images/` and records title, creator, licence and source in `images/sources.json` / `sources.md`; when nothing fits, `scripts/theme_art.py` draws a theme illustration in the style's colours. Credits go into the slide notes (offline) or the delivery message (live).
- `shimo-docx` — new Word documents in the official-document layout: live in the right-side editor (`document_create` from the `gongwen` template, then `document_edit` section by section) or by the offline script when the document needs tables or several copies (live tables cannot be saved yet, office-kit K14).

Existing spreadsheets are read and edited live (`spreadsheet_read`, `spreadsheet_edit`); no spreadsheet skill or template ships, so the model does not create new workbooks in the editor.

Since plugin 0.0.37 they replace `dsh-pptx-style` and `dsh-docx-style` (0.0.36) and continue their 1.0.0 as 0.0.2; skill versions now increase like the plugin's, last number only. A manual copy under an old name in `~/.dsh/skills`, `~/.agents/skills` or a project's `.dsh/skills` / `.agents/skills` no longer shadows the bundled skill; it shows up as a separate skill beside it, so delete it. Turning the component off is kept across the rename: DSH records it against the component id `officesdk-skills`, which did not change.

### Adding a PPT style

A style is a folder `packages/dsh-adapter/skills/shimo-pptx/styles/<name>/` (lower-case letters, digits and hyphens). Both modes pick it by name (`presentation_create { style }`, `brand_deck.py --style`); without one they use `default`. Adding a style changes no code:

1. Copy `styles/default` to `styles/<name>`.
2. Edit `style.json`: `palettes` and `default_palette`, `fonts`, `type_scale_px` (offline sizes on the 960 × 540 px canvas) and `template` (`surface` and `sizes_pt`, the live-mode sizes).
3. Rebuild its live template with the python that has python-pptx (DSH's runtime python), from the repository root: `python packages/dsh-adapter/skills/shimo-pptx/scripts/build_template.py --style <name>` writes `styles/<name>/template.pptx`.
4. Update `layouts.md` and `example.json` if its layouts differ, and tell the model when to choose it in `SKILL.md` 「选样式」.
5. From the repository root, run `node --test packages/dsh-adapter/test/*.test.mjs`: it builds the example with the offline script and checks that every pack's template carries its `style.json` colours and fonts.

## License

The source code in this repository is licensed under the [MIT License](LICENSE). office-kit (the Kit, the
editor and its resources) is a separate external component with its own owner and terms and is not covered by
that license; see [NOTICE](NOTICE).

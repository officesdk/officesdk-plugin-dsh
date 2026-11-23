# 暂存样式包

设计来源还没经用户或法务确认（`meta.design_source.cleared: false`）的样式包放这里，一个包一个目录，格式与 `packages/dsh-adapter/skills/shimo-pptx/styles/<id>/` 相同。

- 各脚本加 `--styles-dir tooling/style-distill/staging`，校验用 `node tooling/style-distill/validate.mjs --styles tooling/style-distill/staging`（`cleared: false` 只记「待登记」）。
- 确认后把 `cleared` 改为 true、整目录挪进 `styles/`，就算登记。

现在没有暂存包。

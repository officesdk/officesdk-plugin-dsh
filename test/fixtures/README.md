Office test inputs for the Kit host and the e2e tooling. `sample.pptx` (and `acceptance/acceptance.pptx`, the same deck) is a neutral two-slide deck generated with python-pptx.

| File | SHA256 |
| --- | --- |
| sample.docx | `0af955a34886817b7d0cc7430cf3203c784e934149617af2c6b10931da65a56a` |
| sample.pptx | `e81721118c3547754785bab7e609f61c49f2f9ebfa528a4d668198fab250b996` |
| sample.xlsx | `ce39a1e4098ab9163390682deb84a8cd6f747c28c32ff2fd928e95ae9fdac2ef` |

PPTX has a title slide and a bulleted slide with one picture. DOCX includes text and tracked revisions, XLSX has `hello` and `42` on Sheet1. Fixtures are only test inputs and must not ship in the production plugin.

Kit editor fixtures (spec §13.7), generated from the files above:

| File | SHA256 | Made by |
| --- | --- | --- |
| kit-edit.docx | `ec4d8c5f092938730a75a31dc9e43d458bf882b3d1a37025cb6033774514ce9e` | sample.docx with `<w:trackRevisions/>` removed from word/settings.xml (the Writer SDK turns tracking off while mounting, which the editor package treats as a fault) |
| kit-edit.xlsx | `7626a2af6261958d53ae00c095305f75ca576605238e5178b5ef9e56f1194f43` | openpyxl: A1 `hello`, B2 `42`, A3 `2`, B3 `=A3*3` with no cached value |

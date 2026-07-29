# Headless Codex exact-answer accuracy

Each repetition is shown separately. Agreement compares parsed final JSON, not prose.

| Task | Arm | Rep 1 | Rep 2 | Agreement |
|---|---|---:|---:|---:|
| B1 | excel-lsp | pass | pass | yes |
| B1 | naive-dump | pass | pass | yes |
| B2 | excel-lsp | pass | pass | yes |
| B2 | naive-dump | pass | pass | yes |
| B3 | excel-lsp | pass | pass | yes |
| B3 | naive-dump | pass | pass | yes |
| B4 | excel-lsp | pass | pass | yes |
| B4 | naive-dump | fail | fail | yes |
| B5 | excel-lsp | pass | pass | yes |
| B5 | naive-dump | pass | fail | no |
| B6 | excel-lsp | pass | pass | yes |
| B6 | naive-dump | pass | pass | yes |

| Arm | Exact answers | Accuracy |
|---|---:|---:|
| excel-lsp | 12/12 | 100.0% |
| naive-dump | 9/12 | 75.0% |

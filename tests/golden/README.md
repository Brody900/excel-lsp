# Golden tests

Workbook-map snapshots begin in P2 with F03 and F20. Runtime `indexed_at` is
normalized to `<indexed_at>` before comparison; every other field is exact.
P3 adds formula-semantic snapshots for F07/F19. P4's
`p4-graph-semantics.json` freezes dependent/precedent traces and path results
for F03/F04/F05/F15/F19 plus the exact F09a/F09b circular-diagnostic split.
All snapshots must remain deterministic and their measured budgets must link
to committed evidence.

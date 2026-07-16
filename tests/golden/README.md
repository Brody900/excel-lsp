# Golden tests

Workbook-map snapshots begin in P2 with F03 and F20. Runtime `indexed_at` is
normalized to `<indexed_at>` before comparison; every other field is exact.
Later phases add diagnostics and trace snapshots. All snapshots must remain
deterministic and their measured budgets must link to committed evidence.

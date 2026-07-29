# Benchmark tasks

The six frozen tasks use deterministic fixtures and exact JSON answers. Every
prompt ends with the same final-line contract, and `benchmarks/check.py` parses
only that final `ANSWER:` line. Prose before the answer is ignored; fuzzy
matching, reordered arrays, omitted fields, and trailing text all fail.

`benchmarks/workloads.py` adds the disclosed 1,000-row archive used by both
arms. Original members remain byte-identical except package declarations and
F03's deliberately extended Summary XML; a regression proves its pre-existing
cells, formulas, and caches remain unchanged. Generated workload paths are
named in each task contract.

- [B1](B1.md): transitive input lineage on F03.
- [B2](B2.md): copied-formula audit on F07.
- [B3](B3.md): cached-error census on F08.
- [B4](B4.md): ordered schema and dtypes on F13.
- [B5](B5.md): conservative block-level impact on F03.
- [B6](B6.md): bounded value lookup on F03.

# Property tests

P2 introduces the first non-vacuous Hypothesis suite in
`test_region_properties.py`:

- I4: every generated non-empty cell belongs to exactly one pairwise-disjoint
  region;
- I5: planted ListObject rectangles are reproduced exactly and never intersect
  heuristic output;
- I6: separated planted rectangular grids retain exact ordinals and symbol IDs
  across stream chunking and reversed table/merge metadata ordering.

The same suite compares the sparse row-sweep overlap index with a brute-force
rectangle scan, including wide rectangles, so the performance optimization is
also checked for exactness rather than only speed. For merge-bearing geometry,
it compares the optimized later-primitive sweep and component-local table BSP
with a separate brute reference. The reference literally enumerates
cell-to-cell table-free witnesses, retains primitive membership through
bounding-box closure, applies tables in fixed rectangle order, and recomputes
each directional child separately. Randomized cells, non-overlapping merges,
both table metadata orders, and gap tolerances 0–3 are covered.

P4 adds `test_edge_store_properties.py`: random dependency-edge rectangles on
multiple sheets are queried through both the R*Tree and documented interval
fallback and compared with an independent brute-force overlap scan. The cases
cover point and range lookups, keyset page sizes, Excel row/column boundaries,
and explicit whole-column rectangles.

P3's formula/reference, structured-context, and R1C1 properties live in their
corresponding files. P5's central non-whitelist error contract is instead
protected by the exact F08 OOXML fixture: all specified errors plus an
unrecognized `t="e"` value are asserted end to end, so a synthetic random-text
property would add no independent semantic dimension.

P6 adds `test_editor_preservation.py`: 50 generated edit scripts mix unique
coordinates, formulas, finite numbers, booleans, strings, and nulls over small
authored workbooks. Every untouched OOXML member must remain byte-identical,
and reparsing must recover every requested value or formula exactly. This is
the randomized I18 proof alongside the exact F16/F21 part diffs. A property
counts as evidence only when it asserts the named invariant rather than merely
proving that random input does not crash.

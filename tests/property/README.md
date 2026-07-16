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

Later phases add formula/reference, R1C1, graph, diagnostics, and surgical-edit
properties with their corresponding production modules. A property counts as
evidence only when it asserts the named invariant rather than merely proving
that random input does not crash.

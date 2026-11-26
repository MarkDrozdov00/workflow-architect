# Composition architecture audit

## Findings

The canonical builder previously established graph ranks before it established a business reading order. Pagination then grouped ranks or phases, and routing optimized the resulting geometry. Recovery and shared-outcome edges could therefore influence rank order as strongly as ordinary progression. A candidate could pass collision checks while requiring backward navigation or placing a terminal outcome inside the main journey.

Process areas constrained vertical placement but did not identify a successful journey. Adaptive recovery could also select a geometrically valid legacy ordering without measuring whether ordinary business progression remained forward.

## Implemented composition model

`analyze_main_path` now derives a semantic order from graph structure, explicit edge classes, process areas, and explicit outcome metadata. It does not infer outcome meaning from labels or node IDs. Explicit `main` and `parallel` relationships lead the order; alternative and exception branches remain adjacent to their originating decision; recovery relationships remain secondary. If metadata does not establish a successful terminal, the analysis records no invented primary terminal.

The semantic index now influences initial layout, process-map pagination, phase ordering, row ordering, and composition scoring. Candidate scoring measures ordinary and primary backward slide edges, primary cross-slide edges, and split immediate outcomes in addition to geometry, utilization, crossings, and connector length. Hard semantic and geometry validation remains unchanged.

Large graphs retain a bounded legacy-order recovery candidate. This prevents the new ordering from replacing a previously viable composition when routing constraints dominate.

## Text serialization

PowerPoint paragraph properties now explicitly disable Latin and East Asian mid-word line breaking. Diamond sizing and existing safe-region padding remain in force. The generated PPTX XML was checked for these properties.


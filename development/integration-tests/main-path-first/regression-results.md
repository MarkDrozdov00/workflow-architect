# Regression results

## Automated tests

- Full suite: **185 passed, 1 failed** in 386.38 seconds.
- New main-path composition tests: **4 passed** in 25.32 seconds.
- The only failure is `test_continuations_avoid_swimlane_labels_and_use_local_routes`: the 34-node banking deck reports 26 connector crossings against the historical target of 18.
- The same test confirms no continuation overlaps a swimlane label and zero branch-label association errors.

The crossing threshold was not changed. The remaining defect is documented in the root-cause and visual-inspection reports.

## Semantic acceptance

Eleven of twelve acceptance fixtures rendered successfully. Every successful fixture reports complete edge coverage and no missing node IDs. All successful fixtures report zero primary backward slide edges. The exact 68-node operational credit-card fixture reached the 110-second routing bound and produced no partial PPTX.

## Isolation and deployment artifacts

The builder remains a self-contained canonical file using Python 3.11, the standard library, and `python-pptx`. The full isolation tests passed as part of the 185 successful tests. No repository-local renderer import was introduced.


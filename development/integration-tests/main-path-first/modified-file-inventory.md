# Modified file inventory

## Canonical deployment files

- `final/ppt-designer/workflow_ppt_builder.py`: main-path analysis, semantic ordering and pagination diagnostics, readability-aware candidate scoring, bounded routing recovery, and explicit no-mid-word PowerPoint text properties.
- `final/ppt-designer/instructions.md`: documents main-path-first composition behavior and semantic pagination.
- `final/deployment-manifest.json`: updated SHA-256 hashes for the two modified Presentation Designer artifacts.

No Workflow Designer instruction, shared schema, theme definition, or source workflow fixture was changed.

## Development-only verification files

- `tests/test_main_path_first_composition.py`
- `development/integration-tests/main-path-first/run_acceptance.py`
- `development/integration-tests/main-path-first/acceptance-results.json`
- Generated PPTX, PDF, PNG, contact-sheet, and report artifacts under `development/integration-tests/main-path-first/`.


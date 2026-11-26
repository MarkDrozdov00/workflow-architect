# ChatPwC file delivery and execution-efficiency report

## Scope

This audit changes agent instructions and deployment documentation only. The canonical builder and all three workflow-schema copies remain byte-identical. No renderer composition, routing, pagination, geometry, visual styling, or workflow-reasoning rule changed.

## File-delivery root cause

The previous instructions told both agents to create a Markdown link using the generated path. That is safe only when the path is already a platform-issued user-download reference. In delegated execution, `/mnt/data/...` or another sandbox path belongs to the subagent runtime and may not be addressable from the parent conversation. Reconstructing a parent Markdown link from that path can therefore navigate incorrectly or fail.

The prompts also asked both the subagent and parent to present the file. Although they said to reuse the reference, the emphasis on a new authored Markdown link could encourage the parent to republish or reupload the same PPTX. ChatPwC may separately append generated-file UI elements, which agents cannot reliably suppress.

The canonical policy now requires one platform file-return operation, preservation of the exact returned reference, and one agent-controlled publication. Agents may not construct a link from a filename, sandbox path, guessed identifier, or another conversation's reference. If ChatPwC exposes only an attachment, the response refers to that attachment without manufacturing a link.

## Execution-efficiency audit

### Previous delegated path

The observed behavior could include schema loading and printing, workflow re-analysis by the subagent, duplicate validation calls, intermediate serialization or files, one builder call, and parent-side file republication. Exact ChatPwC call counts and latency were not available in local logs, so this is an observed-process assessment rather than a measured baseline.

### Updated delegated path

1. Workflow Designer retains one complete validated workflow object.
2. It transfers that object once, unchanged, with a separate compact presentation-preferences object.
3. Presentation Designer enters delegated structured-workflow mode.
4. It skips requirements discovery, workflow reconstruction, and open-ended business analysis.
5. It does not print the schema or run a duplicate Python schema-validation pass when builder validation is sufficient.
6. It invokes the builder once and returns the PPTX once.
7. The parent preserves the returned file reference without copying or reuploading the file.

The shared maximum remains one initial render plus one distinct recovery attempt. Builder-internal bounded recovery does not increase this agent-level allowance.

## Local measurements

A seven-node employee-leave workflow was processed through the unchanged canonical builder on 26 November 2026:

| Operation | Measured time |
|---|---:|
| Compact handoff serialization | 0.000141 s |
| Direct workflow validation | 0.000165 s |
| Builder execution and one-slide PPTX creation | 0.911 s |

The compact payload was 1,254 bytes. The run used one rendering attempt and created one user deliverable. Workflow-generation model time, subagent scheduling, file-upload time, and ChatPwC UI publication latency cannot be measured locally and are excluded.

## Expected tool-call comparison

| Activity | Previous observed possibility | Updated instruction contract |
|---|---|---|
| Load or print schema | Repeated or full-schema output | Load only when needed, cache in context, never print in full |
| Workflow reconstruction | Possible in both agents | Parent creates once; delegated subagent preserves it |
| Separate pre-render Python validation | Potentially duplicated | Builder validation is authoritative unless diagnosing a specific error |
| Workflow intermediate files | Possible | None unless the platform requires file transfer |
| Builder calls | Could multiply across agents | One initial call, at most one distinct recovery across the chain |
| PPTX publication | Parent and subagent could both publish | One platform return and one preserved reference |

These are contract-level reductions. ChatPwC telemetry is required to prove actual production call-count and latency changes.

## Recommended model configuration

- Workflow Designer: medium reasoning for incomplete or materially ambiguous workflows; low or medium reasoning when requirements are complete. Keep output concise and retain JSON internally.
- Workflow PPT Designer: low reasoning for delegated validated JSON. Use medium reasoning only for standalone natural-language workflow creation or a material ambiguity.
- Do not increase reasoning merely for a renderer retry. The deterministic builder supplies validation and layout diagnostics.

No specific ChatPwC model name or latency estimate is asserted because tenant model availability and performance were not verified.

## Remaining platform uncertainties

- Whether a subagent's returned file reference is reusable verbatim in a parent authored response.
- Whether ChatPwC always adds a generated-file UI entry in addition to an authored file reference.
- Whether attached-file references persist across conversational refinements.
- Whether schema references remain cached across separate Run Code executions.
- ChatPwC model scheduling, token, subagent, and attachment latency.

These behaviors require tenant acceptance testing. The updated instructions choose a safe fallback: refer to the actual attached PPTX rather than inventing a link.

## Existing renderer baseline

The builder hash remains `16304930ec14c186c43c0b4068862b95521be733f2b33ec2fa8ff952543f2b91`. The prior regression result remains 185 passed and one known banking connector-crossing threshold failure. Large operational graphs may still exhaust bounded routing, and connector crossings or additional slides remain possible. No threshold or renderer behavior changed in this task.

Targeted prompt-parity, agent-registry, CLI, deployment-isolation, and Presentation Designer tests completed with **52 passed**. Four new tests specifically verify file-reference preservation, delegated structured-workflow mode, manifest integrity, frozen builder hash, and byte-identical schema copies.

## Modified files

Canonical production behavior:

- `final/workflow-designer/instructions.md`
- `final/ppt-designer/instructions.md`
- `final/deployment-manifest.json`

Deployment documentation and verification:

- `development/deployment/chatpwc-deployment-guide.md`
- `development/deployment/chatpwc-acceptance-tests.md`
- `development/deployment/final-artifact-verification.json`
- `development/deployment/chatpwc-integration-efficiency-report.md`
- `tests/test_chatpwc_integration_efficiency.py`


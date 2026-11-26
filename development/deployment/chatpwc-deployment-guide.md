# Workflow Architect: ChatPwC deployment guide

## Deployment status

The canonical agents are ready for controlled ChatPwC use. Workflow PPT Designer builder version is **1.6.7**. Do not treat the local Codex harness as part of the deployment.

The Workflow Designer now contains the approved conditional delegation contract for the subagent named **Workflow PPT Designer**. Both agents remain independently usable.

## 1. Final artifact inventory

| Path relative to `final/` | Purpose | Deployment treatment | Runtime dependency | SHA-256 |
|---|---|---|---|---|
| `workflow-designer/instructions.md` | Canonical Workflow Designer behavior, business-readable delivery, efficient internal JSON retention, conditional presentation delegation, file-reference preservation, and shared render-attempt limit | Copy entire contents into the agent instructions field | None | `7269066cc5fd3b32133523a34a6fd30e539844e34f7c08297963d694cefab08f` |
| `workflow-designer/description.md` | User-facing description | Copy into the agent description field | None | `89006aba584d68d82174473c754fb9cf4fec6b6a5a89671032c6812848050d84` |
| `workflow-designer/workflow.schema.json` | Approved Workflow Architect JSON contract | Upload as an agent reference/attachment | JSON-readable reference support | `9ce37e89a6c325e1ab7d32997d3d27e1a4cee49fdc63f55ff09103b94e69fb30` |
| `workflow-designer/agent.json` | Deployment-neutral capability metadata used by local discovery | Do not upload unless ChatPwC supports importing this exact repository-specific metadata format; otherwise configure fields manually | None | `34f6d2ea1f8ba807f2cdf0eff8ee2630326097b7a64432cd02670b6070d1e27c` |
| `ppt-designer/instructions.md` | Canonical Workflow PPT Designer behavior, standalone and delegated modes, runtime contract, single-publication PPTX delivery, and shared render-attempt limit | Copy entire contents into the agent instructions field | Python execution for rendering | `9122f97479ea09c26ceafc382c17fd8d73c91a02aae2fe7e6bc0e234f968d26c` |
| `ppt-designer/description.md` | User-facing description | Copy into the agent description field | None | `7816537b337a112eb7582b9fa1e1321362fc426c4fdb80773e0e2ffd568a211d` |
| `ppt-designer/workflow.schema.json` | Approved Workflow Architect JSON contract | Upload as an agent reference/attachment | JSON-readable reference support | `9ce37e89a6c325e1ab7d32997d3d27e1a4cee49fdc63f55ff09103b94e69fb30` |
| `ppt-designer/workflow_ppt_builder.py` | Complete self-contained PowerPoint renderer | Upload as the agent's Python/runtime file | Python 3.11+ and `python-pptx>=1.0.0` | `16304930ec14c186c43c0b4068862b95521be733f2b33ec2fa8ff952543f2b91` |
| `ppt-designer/agent.json` | Deployment-neutral capability and script metadata using the deployment name Workflow PPT Designer | Configure equivalent capabilities manually unless ChatPwC explicitly supports this format | None | `86c0f3e1b3e51959c4d81739b3fb79d26d000c03ea6cd8af8330cd133537fba9` |
| `shared/workflow.schema.json` | Reference copy proving the shared contract | Supporting artifact only; do not upload when each agent receives its own identical schema | None | `9ce37e89a6c325e1ab7d32997d3d27e1a4cee49fdc63f55ff09103b94e69fb30` |
| `deployment-manifest.json` | Local inventory, hashes, entry points, dependencies and simulation differences | Deployment metadata only; retain for verification, not as an agent attachment | None | `329efb6f529276f020f8cde29e035dd199fdc910eab5b2bcb3526d430ece9187` |

## 2. Workflow Designer configuration

1. Create a ChatPwC agent using the exact name **Workflow Designer**.
2. Copy [description.md](../../final/workflow-designer/description.md) into the description field.
3. Copy [instructions.md](../../final/workflow-designer/instructions.md) without additions or development wrappers into the instructions field.
4. Upload [workflow.schema.json](../../final/workflow-designer/workflow.schema.json) as the approved workflow-schema reference.
5. Configure only the capabilities supported by the target tenant that correspond to conversation, workflow JSON authoring and workflow validation. The names in `agent.json` are repository metadata, not confirmed ChatPwC configuration keys.
6. Do not enable Python solely for this agent. Workflow Designer produces conversation text and structured JSON; it does not render PowerPoint.
7. Start a fresh conversation and run the standalone acceptance tests before connecting a subagent.

Expected inputs include natural-language process requirements, clarification answers, existing workflow JSON for review, requests for business overview/detailed business/operational detail, and presentation-ready JSON. Expected outputs include targeted clarification questions, assumptions/open questions, workflow explanations, and one canonical JSON object when structured output is requested.

By default, Workflow Designer presents business-readable objectives, requirements, numbered process steps, decisions, exceptions, assumptions, and prioritized questions. It retains the complete validated JSON internally and does not expose raw JSON, IDs, schemas, validation reports, or implementation details unless requested. Explicit JSON requests should receive a downloadable JSON file by default; inline JSON is reserved for an explicit inline or raw-JSON request.

## 3. Workflow PPT Designer configuration

1. Create a ChatPwC agent using the exact name **Workflow PPT Designer**.
2. Copy [description.md](../../final/ppt-designer/description.md) into the description field.
3. Copy [instructions.md](../../final/ppt-designer/instructions.md) without a development wrapper into the instructions field.
4. Upload [workflow.schema.json](../../final/ppt-designer/workflow.schema.json) as the approved schema reference.
5. Upload [workflow_ppt_builder.py](../../final/ppt-designer/workflow_ppt_builder.py) as the executable Python file.
6. Enable the target tenant's supported Python/Run Code capability. Confirm Python 3.11+ and `python-pptx>=1.0.0`. Do not install packages inside the rendering call.
7. Confirm the uploaded module is importable as `workflow_ppt_builder`. The canonical instructions do not specify the physical upload directory because ChatPwC file placement has not been verified.
8. Confirm `/mnt/data/` is writable in the target runtime. This is the documented output location, but tenant compatibility still requires testing.
9. Execute the callable below. Do not generate replacement drawing code or import a repository renderer.
10. On success, return the file at `output_path` once using the platform's supported file-return mechanism. Preserve the returned file reference. On failure, return the structured error and do not claim that a presentation exists.

The successful response must publish one agent-controlled PowerPoint deliverable. When the file-return operation supplies a reusable download reference, use that exact reference once near the beginning. Never construct a link from `output_path`, a sandbox path, a filename, or a guessed identifier. When ChatPwC exposes only an attached generated file, tell the user to download that attachment. Do not copy or reupload the file merely to obtain a link. ChatPwC may append interface-generated attachment entries; agent instructions cannot reliably suppress those elements. Do not expose the complete workflow JSON, schema, traceability mapping, rendering report, or technical diagnostics unless requested.

Allow at most two agent-level render calls per user request across Workflow Designer and Workflow PPT Designer together: one initial attempt and one non-semantic recovery attempt. Do not repeat identical failures. The builder's internal bounded candidate evaluation is unchanged and does not expand this agent-level allowance. After the recovery fails, preserve the graph, explain the limitation, offer overview/subprocess/revised-detail alternatives, and wait for approval before changing scope.

```python
from workflow_ppt_builder import generate_workflow

result = generate_workflow(
    workflow=workflow_data,
    output_path="/mnt/data/workflow.pptx",
    theme="minimal",
    layout="auto",
    include_exceptions="none",
    continuation_style="minimal",
    include_assumptions=False,
    include_open_questions=False,
    exception_ids=None,
    density="compact",
    include_overview=False,
    style=None,
    write_traceability=False,
)
```

The return object contains `status`, `output_path`, `details` and `warnings`. Treat success as valid only when `status == "success"` and the returned file exists. Traceability is always constructed and validated internally. With `write_traceability=False`, the normal ChatPwC delivery writes only the requested PPTX; set it to `True` only when a separate traceability artifact is explicitly required.

### Execution modes

- **Standalone natural-language mode:** Presentation Designer establishes a minimal complete graph, retains it internally, validates it, invokes the builder once, and returns the PPTX.
- **Delegated structured-workflow mode:** Presentation Designer accepts Workflow Designer's complete validated JSON as authoritative, skips requirements discovery and graph reconstruction, applies presentation options separately, and invokes the builder directly. It does not print or reload the full schema when the builder's validation is sufficient.

### Verified instruction/runtime alignment

- Builder filename: explicit.
- Callable and parameters: explicit and match the Python signature.
- Output path: explicit as `/mnt/data/workflow.pptx`.
- Workflow schema: explicitly required as the supplied Workflow Architect schema.
- Failure behavior: explicit.
- Alternative renderer prohibition: explicit (`Do not generate arbitrary PowerPoint drawing code`).
- Return of actual generated files: platform-dependent; instructions require preservation of the platform-returned reference and prohibit converting a subagent path into a parent link.
- User-facing delivery: explicit; the PPTX is primary, one agent-controlled publication is allowed, and technical artifacts are hidden by default.

## 4. Subagent integration

Both agents remain independently usable. The confirmed shared contract is the byte-identical Workflow Architect schema.

When ChatPwC supports connecting one configured agent to another:

1. The user works with Workflow Designer to confirm requirements.
2. Workflow Designer produces one complete valid workflow JSON object.
3. Workflow Designer delegates only when the user explicitly requests a PowerPoint deliverable. Transfer the complete JSON object unchanged to Workflow PPT Designer. Do not summarize, reorder, relabel or omit nodes and edges.
4. Transfer presentation preferences separately: layout, theme, density, exception mode, continuation style, optional sections and validated style overrides.
5. Workflow PPT Designer validates and renders the unchanged graph.
6. Preserve and return the actual PPTX file reference supplied by Workflow PPT Designer. Do not replace a subagent reference with its filesystem path, copy the file, or publish it again.
7. For workflow design, analysis, refinement or JSON-only requests, Workflow Designer operates independently and does not delegate.

This describes the intended handoff, not a confirmed ChatPwC API. Subagent selection, authorization, context transfer, file attachment and Python execution must be configured and tested using features available in the target tenant.

### Configured delegation behavior

The canonical Workflow Designer instructions now explicitly:

- name **Workflow PPT Designer** as the presentation subagent;
- delegate only when the user explicitly requests a PowerPoint deliverable;
- require semantic review before delegation;
- transfer the complete workflow JSON unchanged;
- transfer presentation requirements and styling preferences separately;
- preserve the underlying workflow during presentation-only refinements;
- return a presentation only when an actual output file is available;
- return successful presentations through the existing platform file reference, or refer to the attached file when the reference cannot be reused;
- publish one agent-controlled PPTX deliverable rather than copying, regenerating, reuploading, or reconstructing its link;
- share a maximum of two agent-level render attempts with Workflow PPT Designer;
- preserve and report the workflow when rendering or delegation fails; and
- continue independently when presentation generation is not requested or the subagent is unavailable.

Workflow PPT Designer platform-dependent items:

- The instructions identify the filename and import, but not a confirmed ChatPwC physical upload location.
- They explain `/mnt/data` output and structured results, but the exact reusable ChatPwC file-reference type remains platform behavior.
- The schema is used as an approved reference and the builder validates independently; no tenant-specific attachment lookup mechanism is specified.

Proposed future clarification, subject to target-runtime confirmation:

> Use the uploaded `workflow_ppt_builder.py` from the runtime location configured by ChatPwC. Import `generate_workflow` from that module. Save outputs under `/mnt/data` when that path is available, and return the file through the platform's supported file-output mechanism.

Do not add a physical path or tool name until the target tenant confirms it.

## 5. Runtime dependency

The manifest and canonical Workflow PPT Designer instructions consistently require Python 3.11+ and `python-pptx>=1.0.0`. Local verification used Python 3.11.1 with `python-pptx 1.0.2`.

## 6. Manual ChatPwC configuration

The following settings must be configured through capabilities actually available in the target ChatPwC tenant:

- Create the two agents with the exact names **Workflow Designer** and **Workflow PPT Designer**.
- Copy each canonical description and instruction file into the corresponding configuration fields.
- Upload each agent's schema reference; upload `workflow_ppt_builder.py` only to Workflow PPT Designer.
- Enable Python execution for Workflow PPT Designer and verify the documented dependency and writable output location.
- Select Workflow PPT Designer as an available subagent for Workflow Designer.
- Configure the tenant-supported file-return mechanism so generated PPTX files can be returned to users.

These are configuration outcomes, not claims about ChatPwC field names or invocation APIs. Local paths, tool names, import paths, file attachment behavior and subagent invocation must be confirmed in the target tenant.

## 7. Final isolation verification

- No development prompt is required.
- No development renderer is required.
- Builder imports only Python standard-library modules and `python-pptx`.
- All three schema copies are byte-identical.
- Every manifest-listed hash matches its file.
- Builder version is 1.6.7.
- The local CLI, presentation bridge, fixtures, preview tools and tests are not deployment dependencies.
- The tested ChatPwC tenant successfully executes the builder, delegates to the presentation subagent, and exposes generated PPTX attachments. Physical upload paths and reusable file-reference behavior across parent/subagent authored responses remain platform-dependent.

## 8. Known limitations

- Large operational graphs may fail to find a valid layout, but V1.6.7 terminates with controlled diagnostics.
- Some workflows may use more slides than preferred.
- Sparse slides and connector crossings remain possible.
- Font substitution and preview rendering may differ in ChatPwC.
- ChatPwC rendering and file attachment work in the tested tenant, but preservation of a subagent file reference in a parent authored response remains tenant/platform dependent.
- ChatPwC model behavior may differ from the locally tested Codex model.

## 9. Recommended model configuration

- **Workflow Designer:** use a capable general model with medium reasoning for ambiguous or detailed business workflows. Prefer low or medium reasoning for complete requirements and presentation-only handoffs. Keep response verbosity concise so internal JSON and schema material do not consume user-facing tokens.
- **Workflow PPT Designer:** use low reasoning for delegated validated JSON, because the builder performs deterministic validation and rendering. Use medium reasoning only for standalone natural-language workflow creation or material ambiguity. Do not use high reasoning merely for rendering retries.

These are configuration recommendations, not measured ChatPwC latency claims. Model names and exact reasoning controls depend on the tenant.


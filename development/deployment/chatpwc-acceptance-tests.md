# ChatPwC deployment acceptance tests

Run each test in a fresh conversation unless the test explicitly requires refinement or handoff.

## 1. Workflow Designer standalone conversation

**Prompt**

> We need an employee leave approval process, but we have not decided whether HR approval is mandatory or merely notified. Help me design it. Ask only the questions that materially affect the workflow and do not generate slides.

**Expected behavior**

Workflow Designer asks concise questions about approval authority, rejection, correction/cancellation, notification and the desired detail level. It does not invent an HR approval rule or attempt Python execution.

**Acceptance criteria**

- Confirmed requirements, assumptions and open questions are distinct.
- The workflow is explained with business-readable headings and steps rather than raw JSON or internal IDs.
- Consequential questions are prioritized and distinguish blocking decisions from configurable rules.
- No PowerPoint is generated.
- No unsupported policy is presented as fact.

## 2. Workflow Designer JSON generation

**Prompt**

> Create a detailed business workflow JSON for employee leave. The employee submits a request, the manager approves, rejects, or requests correction, corrections return to the manager, and HR is notified after approval. Rejection closes the request. Use the supplied schema and return one complete JSON object.

**Expected behavior**

The agent creates schema-compatible JSON with labelled decision branches, correction recovery, successful and unsuccessful outcomes, owners and semantic outcome metadata. Because this prompt explicitly requests JSON, it supplies a downloadable JSON file unless inline JSON is also explicitly requested.

**Acceptance criteria**

- JSON validates against the uploaded schema.
- The downloadable JSON contains the complete graph; the conversational summary remains business-readable.
- Every edge references an existing node.
- Decision outcomes are labelled.
- Pending correction is not classified as terminal.

## 3. Workflow PPT Designer standalone natural-language request

**Prompt**

> Create an editable PowerPoint workflow for employee leave approval. The employee submits a request. The manager approves, rejects, or requests correction. Corrections return to manager review. HR is notified after approval.

**Expected behavior**

Workflow PPT Designer creates a minimal valid workflow, requests rendering through the uploaded builder and returns the actual PPTX only after successful execution.

**Acceptance criteria**

- Output file exists and opens.
- The file appears once as the agent-controlled PowerPoint deliverable.
- If the platform returns a reusable download reference, the response uses that exact reference once. If only an attachment is available, the response refers to the attachment and does not manufacture a sandbox-path link.
- The response does not repeat the file under a manually authored file heading.
- The complete workflow JSON and technical diagnostics are not displayed by default.
- Shapes and connectors are editable.
- All three manager outcomes are visible.
- The graph and presentation diagnostics report complete coverage.

## 4. Workflow PPT Designer with supplied JSON

**Prompt**

> Generate a corporate-theme PowerPoint from the following Workflow Architect JSON. Preserve every node, edge, owner, decision label and outcome. Do not change business logic.
>
> [Paste the validated JSON from test 2 here.]

**Expected behavior**

The agent treats the JSON as authoritative, skips requirements discovery and workflow reconstruction, invokes `workflow_ppt_builder.generate_workflow`, and returns the generated file once.

**Acceptance criteria**

- Source JSON remains unchanged.
- The response preserves the platform-returned PPTX reference or refers to the actual attachment; it does not construct a link from `/mnt/data` or another sandbox path.
- The supplied JSON and complete schema are not printed.
- Node and edge coverage are complete.
- No missing or duplicated terminal outcome.
- One render call is used unless the first attempt fails and a distinct recovery configuration is justified.

## 5. Conversational styling refinement

Run after test 4.

**Prompt**

> Keep the workflow logic unchanged. Use monochrome styling, square activity boxes, thicker connectors, and make only the rejected outcome border 2.5 points.

**Expected behavior**

The agent classifies this as presentation-only, retains the workflow, merges the style settings and renders a new PPTX.

**Acceptance criteria**

- Node and edge sets are unchanged.
- Monochrome, square activities and thicker connectors persist.
- Only the targeted rejected node receives the 2.5-point border.
- Element verification reports the actual serialized property.
- The response returns the updated PPTX once without repeating workflow JSON or technical reports.
- The exact platform-returned file reference is reused when available; no sandbox-path link is reconstructed.

## 6. Workflow Designer-to-Workflow PPT Designer handoff

Use only after configuring a supported ChatPwC subagent relationship.

**Prompt**

> Design a business-level supplier invoice workflow and create the PowerPoint. Invoices undergo validation and duplicate checking in parallel. PO invoices require three-way matching. Non-PO invoices require cost-owner approval. Exceptions remain pending for correction. Rejection closes the invoice. Payment failure requires reconciliation and safe retry.

**Expected behavior**

Workflow Designer clarifies material gaps, produces canonical JSON, transfers it unchanged to Workflow PPT Designer with presentation preferences kept separate, and returns the actual PPTX only after confirmed success.

**Acceptance criteria**

- The JSON sent to Workflow PPT Designer is byte-equivalent or canonically equivalent to Workflow Designer output.
- No edge, decision or outcome is removed during handoff.
- Subagent/render failure is reported without claiming a file exists.
- The parent reuses the subagent's returned file reference rather than reuploading or regenerating it.
- The parent does not translate the subagent's filesystem path into a new Markdown link.
- Workflow PPT Designer uses delegated structured-workflow mode: no requirements rediscovery, JSON regeneration, or duplicate general semantic analysis.
- The schema is not printed to tool output, and unchanged JSON is transferred once without intermediate workflow files unless the platform requires one.
- Across the delegation chain, no more than one initial render and one recovery render occur.

Automatic delegation occurs only for an explicit PowerPoint request. A JSON-only or workflow-analysis request must remain within Workflow Designer.

## 7. File-reference and duplicate-publication test

Run after a successful delegated presentation.

**Inspection**

Review the Workflow PPT Designer tool result, the parent Workflow Designer response, and ChatPwC's generated-file area.

**Acceptance criteria**

- The subagent generated the PPTX once.
- The parent preserved the returned file reference and did not copy, regenerate, or reupload the file.
- The authored parent response contains at most one file reference for the PPTX.
- No authored “Files uploaded by the tool” section appears.
- A platform-generated duplicate UI entry, if any, is recorded separately as a platform limitation rather than attributed to a second render.
- Clicking the authored reference downloads or opens the same PPTX. If ChatPwC does not expose a reusable parent reference, the response points to the attached file without a fabricated link.

## 8. Controlled complex-workflow handling

**Prompt**

> Render this complete operational workflow without simplifying it. If it cannot be rendered within the execution budget, return the exact diagnostic and do not produce a partial PowerPoint.
>
> [Supply the validated 71-node operational banking JSON.]

**Expected behavior**

V1.6.7 either generates a fully validated presentation or terminates within the configured budget with routing statistics, affected continuations and suggested next actions.

**Acceptance criteria**

- No process node or edge is silently removed.
- No partial file is presented as successful.
- Failure reports elapsed time, budget/candidate data and the affected relationship.
- The user is offered a less detailed workflow or separate subprocess presentations rather than automatic simplification.
- No third agent-level render is attempted; any change in scope waits for user approval.


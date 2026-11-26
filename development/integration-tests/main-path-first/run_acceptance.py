"""Render main-path-first acceptance fixtures with the canonical final builder."""
from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
HERE = Path(__file__).resolve().parent
OUT = HERE / "after"


def load_builder():
    path = ROOT / "final/ppt-designer/workflow_ppt_builder.py"
    spec = importlib.util.spec_from_file_location("main_path_acceptance_builder", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def employee_leave():
    return {
        "title": "Employee Leave Request", "description": "One-slide approval workflow.",
        "domain": "Human Resources", "version": "1.0",
        "actors": [{"id": "employee", "label": "Employee"},
                   {"id": "manager", "label": "Manager"},
                   {"id": "hr", "label": "HR"}],
        "nodes": [
            {"id": "start", "type": "start", "label": "Start request", "owner": "employee"},
            {"id": "submit", "type": "manual_action", "label": "Submit leave request", "owner": "employee"},
            {"id": "review", "type": "decision", "label": "Manager decision?", "owner": "manager"},
            {"id": "correct", "type": "manual_action", "label": "Correct request", "owner": "employee"},
            {"id": "notify_hr", "type": "process", "label": "Notify HR", "owner": "hr"},
            {"id": "approved", "type": "end", "label": "Leave approved", "owner": "employee"},
            {"id": "rejected", "type": "end", "label": "Leave rejected", "owner": "employee"},
        ],
        "edges": [
            {"from": "start", "to": "submit"}, {"from": "submit", "to": "review"},
            {"from": "review", "to": "notify_hr", "label": "Approve"},
            {"from": "review", "to": "rejected", "label": "Reject", "path": "alternative"},
            {"from": "review", "to": "correct", "label": "Corrections", "path": "recovery"},
            {"from": "correct", "to": "review", "label": "Resubmit", "path": "recovery"},
            {"from": "notify_hr", "to": "approved"},
        ],
        "metadata": {"outcome_classification": {
            "approved": "successful", "rejected": "unsuccessful"}},
    }


def corporate_expansion(base):
    workflow = copy.deepcopy(base)
    workflow["title"] = "Existing-Customer Credit Card Application — Corporate Verification Detail"
    workflow["version"] = "1.1"
    workflow["nodes"] = [node for node in workflow["nodes"] if node["id"] != "authority_check"]
    insertion = [
        {"id": "company_profile", "type": "process", "label": "Validate company profile and status",
         "owner": "bank", "phase": "verification", "process_area": "review"},
        {"id": "representative", "type": "process", "label": "Identify authorized representative",
         "owner": "bank", "phase": "verification", "process_area": "review"},
        {"id": "authority_evidence", "type": "process", "label": "Review signing-authority evidence",
         "owner": "bank", "phase": "verification", "process_area": "review"},
        {"id": "authority_valid", "type": "decision", "label": "Signing authority verified?",
         "owner": "bank", "phase": "verification", "process_area": "review"},
        {"id": "authority_information", "type": "manual_action",
         "label": "Provide missing authority information", "owner": "customer",
         "phase": "verification", "process_area": "review"},
    ]
    applicant_index = next(index for index, node in enumerate(workflow["nodes"])
                           if node["id"] == "applicant_type")
    workflow["nodes"][applicant_index + 1:applicant_index + 1] = insertion
    workflow["edges"] = [edge for edge in workflow["edges"]
                         if edge["from"] != "authority_check" and edge["to"] != "authority_check"]
    workflow["edges"].extend([
        {"from": "applicant_type", "to": "company_profile", "label": "Corporate", "path": "alternative"},
        {"from": "company_profile", "to": "representative", "path": "alternative"},
        {"from": "representative", "to": "authority_evidence", "path": "alternative"},
        {"from": "authority_evidence", "to": "authority_valid", "path": "alternative"},
        {"from": "authority_valid", "to": "assess", "label": "Yes", "path": "main"},
        {"from": "authority_valid", "to": "authority_information", "label": "More information",
         "path": "recovery"},
        {"from": "authority_information", "to": "authority_evidence", "label": "Resubmit",
         "path": "recovery"},
        {"from": "authority_valid", "to": "assisted_pending", "label": "Unresolved",
         "path": "exception"},
    ])
    workflow["metadata"]["refinement_scope"] = "corporate_verification_and_signing_authority"
    return workflow


def load(path):
    return json.loads((ROOT / path).read_text(encoding="utf-8"))


def main():
    builder = load_builder(); OUT.mkdir(parents=True, exist_ok=True)
    credit_overview = load("development/integration-tests/chatpwc-credit-card-failure/credit-card-business-overview.workflow.json")
    expanded = corporate_expansion(credit_overview)
    (HERE / "corporate-verification-expanded.workflow.json").write_text(
        json.dumps(expanded, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    cases = {
        "employee-leave": employee_leave(),
        "qes-overview": load("development/integration-tests/cross-domain-acceptance/qes/business-overview.workflow.json"),
        "qes-detailed": load("development/integration-tests/cross-domain-acceptance/qes/detailed-business.workflow.json"),
        "expense-overview": load("development/integration-tests/cross-domain-acceptance/expense/business-overview.workflow.json"),
        "expense-detailed": load("development/integration-tests/cross-domain-acceptance/expense/detailed-business.workflow.json"),
        "supplier-invoice-overview": load("development/integration-tests/cross-domain-acceptance/supplier_invoice/business-overview.workflow.json"),
        "supplier-invoice-detailed": load("development/integration-tests/cross-domain-acceptance/supplier_invoice/detailed-business.workflow.json"),
        "banking-34": load("development/agent-testing/sessions/20261021T123551Z-7dad7a96/latest_workflow.json"),
        "banking-51": load("development/agent-testing/sessions/20261124T100400Z-5e5377db/exports/workflow.json"),
        "credit-card-68": load("development/integration-tests/chatpwc-credit-card-failure/credit_card_workflow_internal.json"),
        "credit-card-business": credit_overview,
        "credit-card-corporate-expanded": expanded,
    }
    summary = {}
    for name, workflow in cases.items():
        output = OUT / f"{name}.pptx"; started = time.perf_counter()
        options = {"layout": "auto", "theme": "corporate",
                   "include_exceptions": "inline", "write_traceability": False}
        if name == "banking-34":
            options.update(json.loads((ROOT / "development/agent-testing/sessions/"
                                       "20261021T123551Z-7dad7a96/session.json").read_text())[
                                           "presentation_config"])
            options["write_traceability"] = False
        elif name == "credit-card-68":
            options.update(load(
                "development/integration-tests/chatpwc-credit-card-failure/"
                "banking_presentation_config_internal.json"))
            options["layout"] = "auto"
            options["write_traceability"] = False
        elif name == "credit-card-corporate-expanded":
            options.update({"layout": "swimlane", "include_exceptions": "none"})
        result = builder.generate_workflow(
            workflow, output, **options)
        elapsed = round(time.perf_counter() - started, 3)
        details = result.get("details", {}); pagination = details.get("pagination_quality") or {}
        summary[name] = {
            "status": result.get("status"), "elapsed_seconds": elapsed,
            "nodes": len(workflow["nodes"]), "edges": len(workflow["edges"]),
            "source_sha256": hashlib.sha256(json.dumps(workflow, sort_keys=True).encode()).hexdigest(),
            "slides": details.get("slides"), "diagram_slides": details.get("diagram_slides"),
            "node_coverage": details.get("node_coverage"), "edge_coverage": details.get("edge_coverage"),
            "ordinary_backward_edges": pagination.get("ordinary_backward_edges"),
            "primary_backward_edges": pagination.get("primary_backward_edges"),
            "primary_cross_slide_edges": pagination.get("primary_cross_slide_edges"),
            "split_immediate_outcomes": pagination.get("split_immediate_outcomes"),
            "connector_crossings": (details.get("layout_quality") or {}).get("connector_crossings"),
            "geometry_status": details.get("geometry_status"),
            "errors": details.get("errors", []), "warnings": result.get("warnings", []),
            "pptx": str(output.relative_to(ROOT)) if output.exists() else None,
        }
        print(name, summary[name]["status"], elapsed, summary[name]["slides"], flush=True)
    (HERE / "acceptance-results.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()

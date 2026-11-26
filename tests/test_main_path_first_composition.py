import importlib.util
import json
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BUILDER = ROOT / "final/ppt-designer/workflow_ppt_builder.py"


def _builder():
    spec = importlib.util.spec_from_file_location("canonical_main_path_builder", BUILDER)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def test_main_path_uses_explicit_edges_and_outcome_metadata_only():
    builder = _builder()
    workflow = {
        "title": "Approval", "description": "", "domain": "test", "version": "1",
        "actors": [{"id": "a", "label": "Actor"}],
        "nodes": [
            {"id": "start", "type": "start", "label": "Start", "owner": "a"},
            {"id": "review", "type": "decision", "label": "Review?", "owner": "a"},
            {"id": "done", "type": "end", "label": "Completed", "owner": "a"},
            {"id": "rejected", "type": "end", "label": "Rejected", "owner": "a"},
        ],
        "edges": [
            {"from": "start", "to": "review", "path": "main"},
            {"from": "review", "to": "done", "label": "Approved", "path": "main"},
            {"from": "review", "to": "rejected", "label": "Rejected", "path": "alternative"},
        ],
        "metadata": {"outcome_classification": {
            "done": "successful", "rejected": "unsuccessful"}},
    }
    analysis = builder.analyze_main_path(workflow)
    assert analysis["successful_terminal_ids"] == ["done"]
    assert analysis["primary_edge_ids"] == [0, 1]
    assert analysis["primary_node_ids"] == ["start", "review", "done"]
    assert "rejected" not in analysis["primary_node_ids"]


def test_candidate_scoring_penalizes_ordinary_backward_navigation():
    builder = _builder()
    workflow = {
        "nodes": [
            {"id": "start", "type": "start"}, {"id": "next", "type": "process"},
            {"id": "done", "type": "end"}],
        "edges": [
            {"from": "start", "to": "next", "path": "main"},
            {"from": "next", "to": "done", "path": "main"}],
        "metadata": {"outcome_classification": {"done": "successful"}},
    }
    analysis = builder.analyze_main_path(workflow)
    forward = builder._main_path_page_metrics(
        workflow, {"start": 0, "next": 1, "done": 1}, analysis)
    backward = builder._main_path_page_metrics(
        workflow, {"start": 1, "next": 0, "done": 1}, analysis)
    assert forward["ordinary_backward_edges"] == 0
    assert backward["ordinary_backward_edges"] == 1
    assert backward["primary_backward_edges"] == 1


def test_supplier_invoice_overview_regression_renders(tmp_path):
    builder = _builder()
    workflow = json.loads((ROOT / "development/integration-tests/cross-domain-acceptance/"
                           "supplier_invoice/business-overview.workflow.json").read_text())
    output = tmp_path / "supplier-invoice.pptx"
    result = builder.generate_workflow(
        workflow, output, layout="auto", theme="corporate",
        include_exceptions="inline", write_traceability=False)
    assert result["status"] == "success", result
    assert output.is_file()
    assert result["details"]["node_coverage"]["missing"] == []
    assert result["details"]["edge_coverage"]["missing_indices"] == []


def test_powerpoint_serializes_whole_word_line_break_settings(tmp_path):
    builder = _builder()
    workflow = {
        "title": "Long decision test", "description": "", "domain": "test", "version": "1",
        "actors": [{"id": "corporate_customer", "label": "Corporate Customer"}],
        "nodes": [
            {"id": "start", "type": "start", "label": "Begin", "owner": "corporate_customer"},
            {"id": "decision", "type": "decision",
             "label": "Is corporate signing authority sufficiently verified?",
             "owner": "corporate_customer"},
            {"id": "done", "type": "end", "label": "Completed", "owner": "corporate_customer"},
            {"id": "stop", "type": "end", "label": "Stopped", "owner": "corporate_customer"},
        ],
        "edges": [
            {"from": "start", "to": "decision"},
            {"from": "decision", "to": "done", "label": "Yes"},
            {"from": "decision", "to": "stop", "label": "No", "path": "alternative"},
        ],
        "metadata": {"outcome_classification": {
            "done": "successful", "stop": "unsuccessful"}},
    }
    output = tmp_path / "text-properties.pptx"
    result = builder.generate_workflow(workflow, output, layout="swimlane")
    assert result["status"] == "success", result
    with zipfile.ZipFile(output) as archive:
        slide_xml = "".join(archive.read(name).decode("utf-8")
                            for name in archive.namelist()
                            if name.startswith("ppt/slides/slide") and name.endswith(".xml"))
    assert 'latinLnBrk="0"' in slide_xml
    assert 'eaLnBrk="0"' in slide_xml
    records = result["details"]["text_layout_validation"]["records"]
    decision = next(item for item in records if item["shape_name"] == "workflow-node::decision")
    assert decision["word_fragmentation_risk"] is False


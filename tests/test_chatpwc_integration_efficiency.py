import hashlib
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def _text(agent):
    return (ROOT / "final" / agent / "instructions.md").read_text(encoding="utf-8")


def test_parent_preserves_subagent_file_reference_without_path_reconstruction():
    text = _text("workflow-designer")
    assert "exactly as returned" in text
    assert "subagent filesystem path is not evidence" in text
    assert "Do not copy, regenerate, reupload" in text
    assert "manually reproduce a “Files uploaded by the tool” section" in text


def test_presentation_designer_has_fast_delegated_structured_mode():
    text = _text("ppt-designer")
    assert "delegated structured-workflow mode" in text
    assert "Do not repeat requirements discovery" in text
    assert "do not run a separate duplicate Python schema-validation pass" in text
    assert "Do not reconstruct a link from a filename" in text
    assert "tell the user to download the attached PowerPoint" in text


def test_manifest_matches_changed_instructions_and_frozen_builder():
    manifest = json.loads((ROOT / "final" / "deployment-manifest.json").read_text())
    for agent in ("workflow-designer", "ppt-designer"):
        for name, expected in manifest["agents"][agent]["files"].items():
            actual = hashlib.sha256((ROOT / "final" / agent / name).read_bytes()).hexdigest()
            assert actual == expected
    builder = ROOT / "final" / "ppt-designer" / "workflow_ppt_builder.py"
    assert hashlib.sha256(builder.read_bytes()).hexdigest() == (
        "16304930ec14c186c43c0b4068862b95521be733f2b33ec2fa8ff952543f2b91"
    )


def test_all_deployed_schema_copies_remain_identical():
    copies = [
        ROOT / "final" / "shared" / "workflow.schema.json",
        ROOT / "final" / "workflow-designer" / "workflow.schema.json",
        ROOT / "final" / "ppt-designer" / "workflow.schema.json",
    ]
    assert len({path.read_bytes() for path in copies}) == 1


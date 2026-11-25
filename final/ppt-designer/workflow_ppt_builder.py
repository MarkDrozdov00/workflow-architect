"""Self-contained Workflow Architect Presentation Designer builder.

Deployment dependency: python-pptx>=1.0.0.
This file contains validation, layout, routing, styling, pagination, and PPTX
generation. It does not import repository-local renderer modules.
"""

from __future__ import annotations

BUILDER_VERSION = "1.6.7"
import time

_ROUTING_STATE = {
    "deadline": None,
    "grid_expansions": 0,
    "grid_budget_exhaustions": 0,
    "continuation_candidates": 0,
    "fallback_uses": 0,
}


def _routing_budget_check(context: str = "routing"):
    deadline = _ROUTING_STATE.get("deadline")
    if deadline is not None and time.monotonic() > deadline:
        raise ValueError(
            "Routing execution budget exhausted after 110 seconds; "
            f"affected operation: {context}; grid expansions: "
            f"{_ROUTING_STATE['grid_expansions']}; continuation candidates: "
            f"{_ROUTING_STATE['continuation_candidates']}. "
            "No partial PowerPoint was accepted. Consider a less detailed workflow "
            "or separate subprocess presentations."
        )
# ---- Consolidated from renderer/themes.py ----
from dataclasses import dataclass, replace


@dataclass(frozen=True)
class Theme:
    name: str
    background: str = "FFFFFF"
    text: str = "202124"
    muted: str = "667085"
    line: str = "98A2B3"
    lane_fill: str = "F7F8FA"
    accent: str = "2F6FED"
    accent_soft: str = "EAF0FF"
    decision: str = "FFF3D6"
    exception: str = "C2413A"
    exception_soft: str = "FDECEA"
    success: str = "287D5A"
    font: str = "Aptos"


THEMES = {
    "minimal": Theme("minimal"),
    "minimal_process_map": Theme(
        "minimal_process_map", text="25282D", muted="5F6672", line="7E8794",
        lane_fill="FAFAFB", accent="315F86", accent_soft="F7F8FA",
        decision="FFF9E8", exception="B4473D", exception_soft="FFF5F3",
        success="2C7054",
    ),
    "monochrome": Theme(
        "monochrome", text="111111", muted="4A4A4A", line="333333",
        lane_fill="F2F2F2", accent="222222", accent_soft="F7F7F7",
        decision="E2E2E2", exception="333333", exception_soft="EAEAEA",
        success="222222", font="Arial",
    ),
    "corporate": Theme(
        "corporate", text="172033", muted="526079", line="65758B",
        lane_fill="F4F7FA", accent="245C8A", accent_soft="EAF2F8",
        decision="FFF2D9", exception="A94442", exception_soft="FCEBEA",
        success="24705A", font="Aptos",
    ),
    "high_contrast": Theme(
        "high_contrast", text="000000", muted="202020", line="000000",
        lane_fill="F1F1F1", accent="0037A6", accent_soft="FFFFFF",
        decision="FFF2A8", exception="8A0000", exception_soft="FFFFFF",
        success="005B35", font="Arial",
    ),
}


def get_theme(name: str) -> Theme:
    if name not in THEMES:
        raise ValueError(f"Unknown theme: {name}. Available: {', '.join(sorted(THEMES))}")
    return THEMES[name]

# ---- Consolidated from renderer/style.py ----
from copy import deepcopy
from typing import Any


DEFAULT_STYLE = {
    "outcome_colors": {
        "successful": "green",
        "unsuccessful": "red",
        "recoverable": "blue",
        "unclassified": "neutral",
    },
    "outcome_overrides": {},
    "decision_typography": "standard",
    "decision_padding": "standard",
    "branch_label_position": "near_source",
    "connector_label_offset": "standard",
    "spacing": "standard",
    "colors": {},
    "typography": {},
    "shapes": {},
    "connectors": {},
    "elements": {},
    "edges": {},
    "layout_preferences": {},
    "reset_to_theme": False,
}

OUTCOME_CATEGORIES = {"successful", "unsuccessful", "recoverable", "unclassified"}
COLOR_VALUES = {"green", "red", "blue", "neutral"}
ENUMS = {
    "decision_typography": {"compact", "standard", "large"},
    "decision_padding": {"compact", "standard", "generous"},
    "branch_label_position": {"near_source", "centered", "near_target"},
    "connector_label_offset": {"tight", "standard", "wide"},
    "spacing": {"compact", "standard", "spacious"},
}

NAMED_COLORS = {
    "black": "000000", "white": "FFFFFF", "gray": "808080", "grey": "808080",
    "light_gray": "E6E6E6", "light_grey": "E6E6E6", "dark_gray": "333333",
    "dark_grey": "333333", "red": "C2413A", "green": "287D5A", "blue": "2F6FED",
    "orange": "D97706", "yellow": "F5C542", "purple": "7A5AF8", "neutral": "D9DEE7",
}
COLOR_KEYS = {
    "background", "title", "subtitle", "text", "activity_fill", "activity_border",
    "decision_fill", "decision_border", "connector", "connector_label", "swimlane_fill",
    "swimlane_alt_fill", "swimlane_header", "accent", "successful", "unsuccessful",
    "recoverable", "unclassified",
}
FONT_FAMILIES = {"Aptos", "Arial", "Calibri", "Helvetica", "Verdana", "Tahoma",
                 "Trebuchet MS", "Georgia", "Times New Roman"}


def _validated_color(value: Any, path: str) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{path} must be a named color or hexadecimal RGB value.")
    normalized = value.strip().lower().replace(" ", "_")
    if normalized in NAMED_COLORS:
        return NAMED_COLORS[normalized]
    candidate = value.strip().lstrip("#").upper()
    if len(candidate) == 6 and all(char in "0123456789ABCDEF" for char in candidate):
        return candidate
    raise ValueError(f"{path} must be a supported named color or #RRGGBB value.")


def _bounded_number(value: Any, path: str, minimum: float, maximum: float) -> float:
    if not isinstance(value, (int, float)) or isinstance(value, bool) or not minimum <= float(value) <= maximum:
        raise ValueError(f"{path} must be between {minimum:g} and {maximum:g}.")
    return float(value)


def validate_style(style: dict[str, Any] | None, workflow: dict[str, Any] | None = None) -> dict[str, Any]:
    if style is None:
        style = {}
    if not isinstance(style, dict):
        raise ValueError("style must be an object.")
    if style.get("reset_to_theme") is True:
        style = {"reset_to_theme": True}
    unknown = sorted(set(style) - set(DEFAULT_STYLE))
    if unknown:
        raise ValueError("Unsupported style settings: " + ", ".join(unknown) + ".")
    result = deepcopy(DEFAULT_STYLE)
    result["reset_to_theme"] = bool(style.get("reset_to_theme", False))
    for key, allowed in ENUMS.items():
        if key in style:
            if style[key] not in allowed:
                raise ValueError(f"style.{key} must be one of: {', '.join(sorted(allowed))}.")
            result[key] = style[key]
    if "outcome_colors" in style:
        colors = style["outcome_colors"]
        if not isinstance(colors, dict):
            raise ValueError("style.outcome_colors must be an object.")
        unknown_colors = sorted(set(colors) - OUTCOME_CATEGORIES)
        if unknown_colors:
            raise ValueError("Unsupported outcome color categories: " + ", ".join(unknown_colors) + ".")
        for category, value in colors.items():
            if value not in COLOR_VALUES:
                raise ValueError(f"style.outcome_colors.{category} must be one of: {', '.join(sorted(COLOR_VALUES))}.")
            result["outcome_colors"][category] = value
    if "outcome_overrides" in style:
        overrides = style["outcome_overrides"]
        if not isinstance(overrides, dict) or not all(isinstance(key, str) for key in overrides):
            raise ValueError("style.outcome_overrides must map node IDs to semantic categories.")
        invalid = {key: value for key, value in overrides.items() if value not in OUTCOME_CATEGORIES}
        if invalid:
            raise ValueError("Outcome override values must be successful, unsuccessful, recoverable, or unclassified.")
        if workflow is not None:
            nodes = {node["id"]: node for node in workflow.get("nodes", [])}
            missing = sorted(set(overrides) - set(nodes))
            if missing:
                raise ValueError("Outcome overrides reference unknown nodes: " + ", ".join(missing) + ".")
        result["outcome_overrides"] = dict(overrides)
    colors = style.get("colors", {})
    if not isinstance(colors, dict):
        raise ValueError("style.colors must be an object.")
    unknown_colors = sorted(set(colors) - COLOR_KEYS)
    if unknown_colors:
        raise ValueError("Unsupported style.colors settings: " + ", ".join(unknown_colors) + ".")
    result["colors"] = {key: _validated_color(value, f"style.colors.{key}")
                        for key, value in colors.items()}

    typography = style.get("typography", {})
    if not isinstance(typography, dict):
        raise ValueError("style.typography must be an object.")
    typography_keys = {"font_family", "title_size", "subtitle_size", "activity_size",
                       "decision_size", "branch_label_size", "swimlane_label_size",
                       "font_weight", "text_alignment", "text_padding"}
    unknown_typography = sorted(set(typography) - typography_keys)
    if unknown_typography:
        raise ValueError("Unsupported style.typography settings: " + ", ".join(unknown_typography) + ".")
    validated_typography = {}
    if "font_family" in typography:
        if typography["font_family"] not in FONT_FAMILIES:
            raise ValueError("style.typography.font_family is unsupported. Supported fonts: " +
                             ", ".join(sorted(FONT_FAMILIES)) + ".")
        validated_typography["font_family"] = typography["font_family"]
    for key, limits in {"title_size": (16, 32), "subtitle_size": (7, 16),
                        "activity_size": (8, 16), "decision_size": (8, 14),
                        "branch_label_size": (8, 12), "swimlane_label_size": (8, 14),
                        "text_padding": (0.02, 0.20)}.items():
        if key in typography:
            validated_typography[key] = _bounded_number(typography[key], f"style.typography.{key}", *limits)
    for key, allowed in {"font_weight": {"regular", "bold"},
                         "text_alignment": {"left", "center", "right"}}.items():
        if key in typography:
            if typography[key] not in allowed:
                raise ValueError(f"style.typography.{key} must be one of: {', '.join(sorted(allowed))}.")
            validated_typography[key] = typography[key]
    result["typography"] = validated_typography

    shapes = style.get("shapes", {})
    if not isinstance(shapes, dict):
        raise ValueError("style.shapes must be an object.")
    shape_keys = {"activity_corner_style", "border_width", "shadows", "shape_padding", "terminal_shape"}
    unknown_shapes = sorted(set(shapes) - shape_keys)
    if unknown_shapes:
        raise ValueError("Unsupported style.shapes settings: " + ", ".join(unknown_shapes) + ".")
    validated_shapes = {}
    if "activity_corner_style" in shapes:
        if shapes["activity_corner_style"] not in {"rounded", "square"}:
            raise ValueError("style.shapes.activity_corner_style must be rounded or square.")
        validated_shapes["activity_corner_style"] = shapes["activity_corner_style"]
    if "terminal_shape" in shapes:
        if shapes["terminal_shape"] not in {"oval", "rounded", "square"}:
            raise ValueError("style.shapes.terminal_shape must be oval, rounded, or square.")
        validated_shapes["terminal_shape"] = shapes["terminal_shape"]
    if "border_width" in shapes:
        validated_shapes["border_width"] = _bounded_number(shapes["border_width"], "style.shapes.border_width", 0.5, 4)
    if "shape_padding" in shapes:
        validated_shapes["shape_padding"] = _bounded_number(shapes["shape_padding"], "style.shapes.shape_padding", 0.02, 0.20)
    if "shadows" in shapes:
        if not isinstance(shapes["shadows"], bool):
            raise ValueError("style.shapes.shadows must be true or false.")
        validated_shapes["shadows"] = shapes["shadows"]
    result["shapes"] = validated_shapes

    connectors = style.get("connectors", {})
    if not isinstance(connectors, dict):
        raise ValueError("style.connectors must be an object.")
    connector_keys = {"line_color", "line_width", "arrowhead", "branch_label_position",
                      "label_clearance", "spacing", "recovery_line_style"}
    unknown_connectors = sorted(set(connectors) - connector_keys)
    if unknown_connectors:
        raise ValueError("Unsupported style.connectors settings: " + ", ".join(unknown_connectors) + ".")
    validated_connectors = {}
    if "line_color" in connectors:
        validated_connectors["line_color"] = _validated_color(connectors["line_color"], "style.connectors.line_color")
    if "line_width" in connectors:
        validated_connectors["line_width"] = _bounded_number(connectors["line_width"], "style.connectors.line_width", 0.5, 4)
    for key, allowed in {"arrowhead": {"triangle", "none"},
                         "branch_label_position": {"near_source", "centered", "near_target"},
                         "label_clearance": {"tight", "standard", "wide"},
                         "spacing": {"compact", "standard", "spacious"},
                         "recovery_line_style": {"solid", "dashed"}}.items():
        if key in connectors:
            if connectors[key] not in allowed:
                raise ValueError(f"style.connectors.{key} must be one of: {', '.join(sorted(allowed))}.")
            validated_connectors[key] = connectors[key]
    result["connectors"] = validated_connectors
    if "branch_label_position" in validated_connectors:
        result["branch_label_position"] = validated_connectors["branch_label_position"]
    if "label_clearance" in validated_connectors:
        result["connector_label_offset"] = validated_connectors["label_clearance"]
    if "spacing" in validated_connectors:
        result["spacing"] = validated_connectors["spacing"]

    nodes = {node["id"]: node for node in (workflow or {}).get("nodes", [])}
    elements = style.get("elements", {})
    if not isinstance(elements, dict):
        raise ValueError("style.elements must map node IDs to style objects.")
    missing_elements = sorted(set(elements) - set(nodes)) if workflow is not None else []
    if missing_elements:
        raise ValueError("Element style overrides reference unknown nodes: " + ", ".join(missing_elements) + ".")
    validated_elements = {}
    for node_id, values in elements.items():
        if not isinstance(values, dict):
            raise ValueError(f"style.elements.{node_id} must be an object.")
        unknown_values = sorted(set(values) - {"fill", "border", "text_color", "border_width", "text_size"})
        if unknown_values:
            raise ValueError(f"Unsupported element settings for {node_id}: {', '.join(unknown_values)}.")
        item = {}
        for key in ("fill", "border", "text_color"):
            if key in values: item[key] = _validated_color(values[key], f"style.elements.{node_id}.{key}")
        if "border_width" in values: item["border_width"] = _bounded_number(values["border_width"], f"style.elements.{node_id}.border_width", 0.5, 4)
        if "text_size" in values: item["text_size"] = _bounded_number(values["text_size"], f"style.elements.{node_id}.text_size", 8, 16)
        validated_elements[node_id] = item
    result["elements"] = validated_elements

    edge_pairs = {f"{edge['from']}->{edge['to']}" for edge in (workflow or {}).get("edges", [])}
    edge_styles = style.get("edges", {})
    if not isinstance(edge_styles, dict):
        raise ValueError("style.edges must map source->target references to style objects.")
    missing_edge_styles = sorted(set(edge_styles) - edge_pairs) if workflow is not None else []
    if missing_edge_styles:
        raise ValueError("Edge style overrides reference unknown edges: " + ", ".join(missing_edge_styles) + ".")
    validated_edges = {}
    for edge_id, values in edge_styles.items():
        if not isinstance(values, dict) or set(values) - {"color", "width", "line_style"}:
            raise ValueError(f"style.edges.{edge_id} supports only color, width, and line_style.")
        item = {}
        if "color" in values: item["color"] = _validated_color(values["color"], f"style.edges.{edge_id}.color")
        if "width" in values: item["width"] = _bounded_number(values["width"], f"style.edges.{edge_id}.width", 0.5, 4)
        if "line_style" in values:
            if values["line_style"] not in {"solid", "dashed"}:
                raise ValueError(f"style.edges.{edge_id}.line_style must be solid or dashed.")
            item["line_style"] = values["line_style"]
        validated_edges[edge_id] = item
    result["edges"] = validated_edges
    preferences = style.get("layout_preferences", {})
    if not isinstance(preferences, dict):
        raise ValueError("style.layout_preferences must be an object.")
    preference_keys = {"orientation", "preferred_slide_count", "content_margins", "activity_spacing"}
    unknown_preferences = sorted(set(preferences) - preference_keys)
    if unknown_preferences:
        raise ValueError("Unsupported style.layout_preferences settings: " + ", ".join(unknown_preferences) + ".")
    validated_preferences = {}
    if "orientation" in preferences:
        if preferences["orientation"] != "landscape":
            raise ValueError("style.layout_preferences.orientation currently supports landscape only.")
        validated_preferences["orientation"] = "landscape"
    if "preferred_slide_count" in preferences:
        validated_preferences["preferred_slide_count"] = int(_bounded_number(
            preferences["preferred_slide_count"], "style.layout_preferences.preferred_slide_count", 1, 20))
    for key in ("content_margins", "activity_spacing"):
        if key in preferences:
            if preferences[key] not in {"compact", "standard", "spacious"}:
                raise ValueError(f"style.layout_preferences.{key} must be compact, standard, or spacious.")
            validated_preferences[key] = preferences[key]
    if "activity_spacing" in validated_preferences:
        result["spacing"] = validated_preferences["activity_spacing"]
    result["layout_preferences"] = validated_preferences
    return result


def workflow_outcome_metadata(workflow: dict[str, Any]) -> dict[str, str]:
    metadata = workflow.get("metadata") or {}
    values = metadata.get("outcome_classification") or {}
    if not isinstance(values, dict):
        return {}
    return {str(key): str(value) for key, value in values.items() if value in OUTCOME_CATEGORIES}


def _apply_style_to_theme(theme: Theme, style: dict[str, Any]) -> Theme:
    colors = style.get("colors", {}); typography = style.get("typography", {})
    return replace(
        theme,
        background=colors.get("background", theme.background),
        text=colors.get("text", theme.text),
        muted=colors.get("subtitle", colors.get("connector_label", theme.muted)),
        line=colors.get("connector", colors.get("activity_border", theme.line)),
        lane_fill=colors.get("swimlane_fill", theme.lane_fill),
        accent=colors.get("accent", colors.get("decision_border", theme.accent)),
        accent_soft=colors.get("activity_fill", theme.accent_soft),
        decision=colors.get("decision_fill", theme.decision),
        exception=colors.get("unsuccessful", theme.exception),
        success=colors.get("successful", theme.success),
        font=typography.get("font_family", theme.font),
    )


def _style_value(style: dict[str, Any], section: str, key: str, default: Any) -> Any:
    return style.get(section, {}).get(key, default)


def _node_visual(node: dict[str, Any], theme: Theme, style: dict[str, Any],
                 fill: str, border: str, text_color: str, border_width: float, text_size: float):
    colors = style.get("colors", {}); node_style = style.get("elements", {}).get(node["id"], {})
    if node.get("type") == "decision":
        fill = colors.get("decision_fill", fill); border = colors.get("decision_border", border)
    elif node.get("type") not in {"start", "end", "error"}:
        fill = colors.get("activity_fill", fill); border = colors.get("activity_border", border)
    fill = node_style.get("fill", fill); border = node_style.get("border", border)
    text_color = node_style.get("text_color", text_color)
    border_width = node_style.get("border_width", style.get("shapes", {}).get("border_width", border_width))
    text_size = node_style.get("text_size", text_size)
    return fill, border, text_color, float(border_width), float(text_size)


def _edge_visual(edge: dict[str, Any], theme: Theme, style: dict[str, Any], special: bool,
                 default_width: float):
    connectors = style.get("connectors", {}); colors = style.get("colors", {})
    edge_style = style.get("edges", {}).get(f"{edge['from']}->{edge['to']}", {})
    default_color = theme.exception if special else theme.line
    color = edge_style.get("color", connectors.get("line_color", colors.get("connector", default_color)))
    # A semantic palette may intentionally use a very pale unsuccessful color
    # as the node fill. Reusing that tint for a transparent connector and label
    # makes the route disappear against white swimlanes. Preserve its hue but
    # enforce stroke contrast unless the caller explicitly chose a connector
    # or per-edge line color.
    if special and "color" not in edge_style and "line_color" not in connectors and "connector" not in colors:
        color = _legible_semantic_stroke(color)
    width = edge_style.get("width", connectors.get("line_width", default_width))
    line_style = edge_style.get("line_style",
                 connectors.get("recovery_line_style", "solid") if edge.get("path") == "recovery" else "solid")
    return color, float(width), line_style, connectors.get("arrowhead", "triangle")


def _legible_semantic_stroke(value: str) -> str:
    """Darken pale semantic fill tints while preserving their visual hue."""
    import colorsys
    raw = str(value).lstrip("#")
    if len(raw) != 6:
        return value
    try:
        red, green, blue = (int(raw[index:index + 2], 16) / 255 for index in (0, 2, 4))
    except ValueError:
        return value
    hue, lightness, saturation = colorsys.rgb_to_hls(red, green, blue)
    if lightness <= 0.62:
        return raw.upper()
    red, green, blue = colorsys.hls_to_rgb(hue, 0.42, max(0.45, saturation))
    return "".join(f"{round(channel * 255):02X}" for channel in (red, green, blue))


def _set_dash(line, line_style: str):
    if line_style != "dashed":
        return
    ln = line.line._get_or_add_ln()
    for child in list(ln):
        if child.tag.endswith("prstDash"):
            ln.remove(child)
    dash = OxmlElement("a:prstDash"); dash.set("val", "dash"); ln.append(dash)


def _apply_shadow(shape, enabled: bool):
    sp_pr = shape._element.spPr
    for child in list(sp_pr):
        if child.tag.endswith("effectLst"):
            sp_pr.remove(child)
    if not enabled:
        return
    effects = OxmlElement("a:effectLst")
    shadow = OxmlElement("a:outerShdw")
    shadow.set("blurRad", "38100"); shadow.set("dist", "25400"); shadow.set("dir", "2700000")
    shadow.set("algn", "ctr"); shadow.set("rotWithShape", "0")
    color = OxmlElement("a:srgbClr"); color.set("val", "000000")
    alpha = OxmlElement("a:alpha"); alpha.set("val", "18000"); color.append(alpha); shadow.append(color)
    effects.append(shadow); sp_pr.append(effects)

# ---- Consolidated from renderer/validation.py ----
from collections import Counter, defaultdict, deque
from dataclasses import dataclass, field
from typing import Any

NODE_TYPES = {
    "start", "end", "process", "decision", "subprocess", "manual_action",
    "external_system", "error",
}


@dataclass
class ValidationReport:
    valid: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)


class WorkflowValidationError(ValueError):
    def __init__(self, report: ValidationReport):
        self.report = report
        super().__init__("; ".join(report.errors))


def validate_workflow(workflow: Any, *, raise_on_error: bool = True) -> ValidationReport:
    errors: list[str] = []
    warnings: list[str] = []
    if not isinstance(workflow, dict):
        report = ValidationReport(False, ["Workflow must be a JSON object."])
        if raise_on_error:
            raise WorkflowValidationError(report)
        return report

    for key in ("title", "domain", "version", "actors", "nodes", "edges"):
        if key not in workflow:
            errors.append(f"Missing required field: {key}.")
    nodes = workflow.get("nodes", [])
    edges = workflow.get("edges", [])
    actors = workflow.get("actors", [])
    if not isinstance(nodes, list) or not nodes:
        errors.append("nodes must be a non-empty array.")
        nodes = []
    if not isinstance(edges, list):
        errors.append("edges must be an array.")
        edges = []
    if not isinstance(actors, list) or not actors:
        errors.append("actors must be a non-empty array.")
        actors = []

    actor_ids = {a.get("id") for a in actors if isinstance(a, dict) and a.get("id")}
    process_areas = workflow.get("process_areas", [])
    area_ids = [area.get("id") for area in process_areas if isinstance(area, dict) and area.get("id")]
    if len(area_ids) != len(set(area_ids)):
        errors.append("Process area IDs must be unique.")
    node_ids: list[str] = []
    node_by_id: dict[str, dict[str, Any]] = {}
    for i, node in enumerate(nodes):
        if not isinstance(node, dict):
            errors.append(f"Node at index {i} must be an object.")
            continue
        nid = node.get("id")
        if not isinstance(nid, str) or not nid.strip():
            errors.append(f"Node at index {i} has no valid id.")
            continue
        if nid in node_by_id:
            errors.append(f"Duplicate node id: {nid}.")
        node_ids.append(nid)
        node_by_id[nid] = node
        if node.get("type") not in NODE_TYPES:
            errors.append(f"Node {nid} has unsupported type: {node.get('type')}.")
        if not node.get("label"):
            errors.append(f"Node {nid} has no label.")
        if node.get("owner") not in actor_ids:
            errors.append(f"Node {nid} references unknown owner: {node.get('owner')}.")
        hint = node.get("layout_hint") or {}
        requested_area = hint.get("preferred_process_area") or node.get("process_area") or node.get("lane")
        if requested_area and area_ids and requested_area not in area_ids:
            warnings.append(f"Node {nid} references unknown process area: {requested_area}.")
        relative = hint.get("preferred_relative_position") or {}
        for relation, target in relative.items():
            if target == nid:
                warnings.append(f"Node {nid} has a self-referential {relation} layout hint.")

    outcome_metadata = (workflow.get("metadata") or {}).get("outcome_classification") or {}
    if not isinstance(outcome_metadata, dict):
        errors.append("metadata.outcome_classification must be an object.")
    else:
        unknown_outcomes = sorted(set(outcome_metadata) - set(node_by_id))
        if unknown_outcomes:
            errors.append("Outcome classification references unknown nodes: " + ", ".join(unknown_outcomes) + ".")
        allowed_outcomes = {"successful", "unsuccessful", "recoverable", "unclassified"}
        invalid_outcomes = sorted(node_id for node_id, value in outcome_metadata.items()
                                  if value not in allowed_outcomes)
        if invalid_outcomes:
            errors.append("Outcome classification has unsupported values for: " + ", ".join(invalid_outcomes) + ".")

    outgoing: dict[str, list[dict[str, Any]]] = defaultdict(list)
    incoming: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for i, edge in enumerate(edges):
        if not isinstance(edge, dict):
            errors.append(f"Edge at index {i} must be an object.")
            continue
        src, dst = edge.get("from"), edge.get("to")
        if src not in node_by_id:
            errors.append(f"Edge {i} references missing source node: {src}.")
        if dst not in node_by_id:
            errors.append(f"Edge {i} references missing target node: {dst}.")
        if src in node_by_id and dst in node_by_id:
            outgoing[src].append(edge)
            incoming[dst].append(edge)

    starts = sorted(n for n, v in node_by_id.items() if v.get("type") == "start")
    ends = sorted(n for n, v in node_by_id.items() if v.get("type") in {"end", "error"})
    if not starts:
        errors.append("Workflow must contain at least one start node.")
    if not ends:
        errors.append("Workflow must contain at least one terminal outcome (end or error).")

    for nid, node in node_by_id.items():
        if node.get("type") == "decision":
            branches = outgoing[nid]
            if len(branches) < 2:
                errors.append(f"Decision node {nid} must have at least two outgoing branches.")
            if any(not str(e.get("label", "")).strip() for e in branches):
                errors.append(f"Decision node {nid} has an unlabeled outgoing branch.")

    reachable: set[str] = set()
    queue = deque(starts)
    while queue:
        nid = queue.popleft()
        if nid in reachable:
            continue
        reachable.add(nid)
        queue.extend(e["to"] for e in outgoing[nid])
    unreachable = sorted(set(node_by_id) - reachable)
    if unreachable:
        warnings.append("Unreachable nodes: " + ", ".join(unreachable) + ".")

    can_terminate: set[str] = set(ends)
    queue = deque(ends)
    while queue:
        nid = queue.popleft()
        for e in incoming[nid]:
            if e["from"] not in can_terminate:
                can_terminate.add(e["from"])
                queue.append(e["from"])
    stranded = sorted(reachable - can_terminate)
    if stranded:
        errors.append("Reachable nodes without a path to a terminal outcome: " + ", ".join(stranded) + ".")

    report = ValidationReport(
        not errors, errors, warnings,
        {"node_count": len(node_by_id), "edge_count": len(edges), "actors": len(actor_ids),
         "starts": starts, "terminals": ends, "unreachable": unreachable},
    )
    if errors and raise_on_error:
        raise WorkflowValidationError(report)
    return report

# ---- Consolidated from renderer/layout.py ----
from collections import defaultdict, deque
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class Placement:
    node_id: str
    page: int
    rank: int
    lane: int
    slot: int
    column: int = 0


@dataclass(frozen=True)
class DecisionGroup:
    decision_id: str
    node_ids: tuple[str, ...]
    edge_labels: tuple[str, ...]
    min_rank: int
    max_rank: int
    feasible: bool


@dataclass
class LayoutResult:
    placements: dict[str, Placement]
    pages: int
    back_edges: set[tuple[str, str]]
    ranks_per_page: int
    page_ranks: list[list[int]]
    decision_groups: list[DecisionGroup]
    split_decision_groups: list[str]
    rendering_mode: str = "compact"
    unused_space: float = 0.0
    split_edges: int = 0
    connector_crossings: int = 0
    graph_analysis: dict[str, Any] | None = None
    candidate_scores: tuple[dict[str, Any], ...] = ()
    continuation_space_reserve: float = 0.0


@dataclass(frozen=True)
class SlideCompositionPage:
    """Semantic content assigned to one diagram slide before drawing begins."""
    slide_number: int
    node_ids: tuple[str, ...]
    local_edge_ids: tuple[int, ...]
    incoming_cross_edge_ids: tuple[int, ...]
    outgoing_cross_edge_ids: tuple[int, ...]
    decision_ids: tuple[str, ...]
    recovery_edge_ids: tuple[int, ...]
    terminal_ids: tuple[str, ...]
    reading_order: tuple[str, ...]


@dataclass(frozen=True)
class SlideCompositionPlan:
    pages: tuple[SlideCompositionPage, ...]
    edge_records: tuple[dict[str, Any], ...]
    slide_relationships: tuple[dict[str, Any], ...]
    strategy: str


def build_slide_composition_plan(workflow: dict[str, Any], page_of: dict[str, int],
                                 ordered_pages: list[list[str]], strategy: str) -> SlideCompositionPlan:
    """Create a coordinate-independent, lossless presentation plan."""
    nodes = {node["id"]: node for node in workflow["nodes"]}
    edge_records = []
    for edge_id, edge in enumerate(workflow["edges"]):
        source_page = page_of.get(edge["from"])
        target_page = page_of.get(edge["to"])
        representation = "local" if source_page == target_page else "cross_slide"
        edge_records.append({
            "edge_id": edge_id, "source": edge["from"], "target": edge["to"],
            "label": str(edge.get("label", "")), "path": edge.get("path", "main"),
            "source_page": source_page, "target_page": target_page,
            "representation": representation,
        })
    pages = []
    for page_index, node_ids in enumerate(ordered_pages):
        node_set = set(node_ids)
        local = tuple(record["edge_id"] for record in edge_records
                      if record["source"] in node_set and record["target"] in node_set)
        incoming = tuple(record["edge_id"] for record in edge_records
                         if record["target"] in node_set and record["source"] not in node_set)
        outgoing = tuple(record["edge_id"] for record in edge_records
                         if record["source"] in node_set and record["target"] not in node_set)
        pages.append(SlideCompositionPage(
            slide_number=page_index + 1,
            node_ids=tuple(node_ids), local_edge_ids=local,
            incoming_cross_edge_ids=incoming, outgoing_cross_edge_ids=outgoing,
            decision_ids=tuple(node_id for node_id in node_ids if nodes[node_id].get("type") == "decision"),
            recovery_edge_ids=tuple(record["edge_id"] for record in edge_records
                                    if record["path"] == "recovery" and
                                    (record["source"] in node_set or record["target"] in node_set)),
            terminal_ids=tuple(node_id for node_id in node_ids
                               if nodes[node_id].get("type") in {"end", "error"}),
            reading_order=tuple(node_ids),
        ))
    relationships = []
    grouped: dict[tuple[int, int], list[int]] = defaultdict(list)
    for record in edge_records:
        if record["representation"] == "cross_slide":
            grouped[(record["source_page"], record["target_page"])].append(record["edge_id"])
    for (source_page, target_page), edge_ids in sorted(grouped.items()):
        relationships.append({"from_slide": source_page + 1, "to_slide": target_page + 1,
                              "edge_ids": edge_ids, "direction": "recovery" if target_page < source_page else "forward"})
    return SlideCompositionPlan(tuple(pages), tuple(edge_records), tuple(relationships), strategy)


def validate_composition_plan(workflow: dict[str, Any], plan: SlideCompositionPlan) -> dict[str, Any]:
    """Compare a slide plan with its source graph before any PowerPoint is saved."""
    critical: list[str] = []; warnings: list[str] = []
    source_nodes = {node["id"]: node for node in workflow["nodes"]}
    planned_nodes = [node_id for page in plan.pages for node_id in page.node_ids]
    missing_nodes = sorted(set(source_nodes) - set(planned_nodes))
    duplicate_nodes = sorted(node_id for node_id in set(planned_nodes) if planned_nodes.count(node_id) > 1)
    if missing_nodes: critical.append("Missing source nodes: " + ", ".join(missing_nodes) + ".")
    if duplicate_nodes: critical.append("Duplicated source nodes: " + ", ".join(duplicate_nodes) + ".")

    records = list(plan.edge_records)
    record_ids = [record["edge_id"] for record in records]
    expected_ids = list(range(len(workflow["edges"])))
    missing_edges = sorted(set(expected_ids) - set(record_ids))
    duplicate_edges = sorted(edge_id for edge_id in set(record_ids) if record_ids.count(edge_id) > 1)
    if missing_edges: critical.append("Missing source edges: " + ", ".join(map(str, missing_edges)) + ".")
    if duplicate_edges: critical.append("Duplicated source edges: " + ", ".join(map(str, duplicate_edges)) + ".")
    for edge_id, edge in enumerate(workflow["edges"]):
        matching = [record for record in records if record["edge_id"] == edge_id]
        if not matching:
            continue
        record = matching[0]
        expected = (edge["from"], edge["to"], str(edge.get("label", "")), edge.get("path", "main"))
        actual = (record["source"], record["target"], record["label"], record["path"])
        if actual != expected:
            critical.append(f"Edge {edge_id} mapping differs from the source graph: expected {expected}, got {actual}.")
        if record["representation"] == "cross_slide" and not any(
                edge_id in relationship["edge_ids"] for relationship in plan.slide_relationships):
            critical.append(f"Cross-slide edge {edge_id} has no slide relationship.")
        if (record["representation"] == "cross_slide" and
                record["target_page"] < record["source_page"] and
                record["path"] != "recovery"):
            # Slide order is a presentation choice, not process direction.
            # Mapping and cross-slide coverage remain hard requirements.
            warnings.append(f"Non-recovery edge {edge_id} uses backward cross-slide continuity.")

    outgoing: dict[str, list[tuple[str, str]]] = defaultdict(list)
    planned_outgoing: dict[str, list[tuple[str, str]]] = defaultdict(list)
    incoming: dict[str, list[str]] = defaultdict(list)
    for edge in workflow["edges"]:
        outgoing[edge["from"]].append((str(edge.get("label", "")), edge["to"]))
        incoming[edge["to"]].append(edge["from"])
    for record in records:
        planned_outgoing[record["source"]].append((record["label"], record["target"]))
    decision_mappings = {}
    for node_id, node in source_nodes.items():
        if node.get("type") != "decision":
            continue
        expected = sorted(outgoing[node_id]); actual = sorted(planned_outgoing[node_id])
        decision_mappings[node_id] = [{"label": label, "target": target} for label, target in actual]
        if actual != expected:
            critical.append(f"Decision {node_id} branch mapping differs from the source graph.")
    orphaned = sorted(node_id for node_id, node in source_nodes.items()
                      if node.get("type") in {"end", "error"} and not incoming[node_id])
    if orphaned:
        critical.append("Terminal outcomes have no source predecessor: " + ", ".join(orphaned) + ".")
    return {
        "status": "failed" if critical else "verified",
        "critical": critical, "warnings": warnings,
        "informational": {
            "slides": len(plan.pages), "nodes": len(planned_nodes), "edges": len(records),
            "cross_slide_edges": sum(record["representation"] == "cross_slide" for record in records),
            "layout_strategy": plan.strategy,
        },
        "checks": {"missing_nodes": missing_nodes, "duplicate_nodes": duplicate_nodes,
                   "missing_edges": missing_edges, "duplicate_edges": duplicate_edges,
                   "orphaned_terminals": orphaned, "decision_branch_mappings": decision_mappings},
    }


def calculate_layout(workflow: dict[str, Any], layout: str = "flowchart",
                     ranks_per_page: int | None = None, rendering_mode: str = "compact",
                     preferred_slide_count: int | None = None,
                     allow_adaptive_compression: bool = True,
                     continuation_space_reserve: float = 0.0) -> LayoutResult:
    if layout == "standard":
        layout = "flowchart"
    if layout == "process_map":
        layout = "flowchart"
    if layout not in {"swimlane", "flowchart"}:
        raise ValueError("layout must be 'flowchart', 'swimlane', or 'process_map'")
    if rendering_mode not in {"compact", "balanced", "detailed"}:
        raise ValueError("rendering_mode must be 'compact', 'balanced', or 'detailed'")
    nodes = {n["id"]: n for n in workflow["nodes"]}
    if ranks_per_page is not None:
        capacity = ranks_per_page
    elif rendering_mode == "compact":
        # Evaluate the complete graph before introducing a rank boundary.  The
        # renderer can use several lanes/rows, so medium graphs should not be
        # fragmented merely because they have more than seven logical ranks.
        capacity = 12 if len(nodes) <= 12 and len(workflow["edges"]) <= 15 else 10
    else:
        capacity = {"balanced": 9, "detailed": 6}[rendering_mode]
    actors = [a["id"] for a in workflow["actors"]]
    lane_index = {aid: i for i, aid in enumerate(actors)}
    adjacency: dict[str, list[str]] = defaultdict(list)
    indegree = {nid: 0 for nid in nodes}
    for e in workflow["edges"]:
        # Recovery edges intentionally point back to a prior activity. They
        # must be routed as loops, not included in the forward topological
        # order used for pagination; otherwise one retry cycle can place a
        # valid main-path merge before its own predecessor on another slide.
        if e.get("path") == "recovery" and len(nodes) > 24:
            continue
        adjacency[e["from"]].append(e["to"])
        indegree[e["to"]] += 1
    for values in adjacency.values():
        values.sort()

    rank = {nid: 0 for nid, n in nodes.items() if n["type"] == "start"}
    queue = deque(sorted(n for n, d in indegree.items() if d == 0))
    # Recovery-only entry points also have zero forward indegree. Seed them at
    # rank zero so their non-recovery fallback edges participate in normal
    # ordering instead of being assigned an artificial final rank later.
    for node_id in queue:
        rank.setdefault(node_id, 0)
    remaining = indegree.copy()
    processed: set[str] = set()
    while queue:
        src = queue.popleft()
        processed.add(src)
        for dst in adjacency[src]:
            rank[dst] = max(rank.get(dst, 0), rank.get(src, 0) + 1)
            remaining[dst] -= 1
            if remaining[dst] == 0:
                queue.append(dst)

    # Cycles remain after Kahn's algorithm. Assign stable shortest-path ranks and
    # classify edges pointing to the same or an earlier rank as retry/back edges.
    bfs = deque(sorted(rank, key=lambda n: (rank[n], n)))
    seen = set(bfs)
    while bfs:
        src = bfs.popleft()
        for dst in adjacency[src]:
            candidate = rank.get(src, 0) + 1
            if dst not in seen:
                rank[dst] = candidate
                seen.add(dst)
                bfs.append(dst)
    for nid in sorted(nodes):
        rank.setdefault(nid, max(rank.values(), default=0) + 1)

    # A recovery-only node (for example, "Retry ID scan?") has no incoming
    # edge in the forward DAG, but it is not a true process entry point. Anchor
    # it immediately after the activity that invokes the recovery choice, then
    # propagate its ordinary fallback path forward. This keeps recovery choices
    # beside their cause without allowing the retry edge itself to distort the
    # main topological order.
    recovery_incoming: dict[str, list[str]] = defaultdict(list)
    for edge in workflow["edges"]:
        if edge.get("path") == "recovery" and len(nodes) > 24:
            recovery_incoming[edge["to"]].append(edge["from"])
    anchored_recovery_nodes = {
        node_id for node_id in nodes
        if indegree.get(node_id, 0) == 0 and nodes[node_id].get("type") != "start" and
        recovery_incoming.get(node_id) and (len(nodes) > 24 or adjacency.get(node_id))
    }
    for node_id in sorted(anchored_recovery_nodes):
        rank[node_id] = min(rank[source_id] + 1 for source_id in recovery_incoming[node_id])
    # The non-recovery graph is acyclic after the cycle handling above. A
    # bounded relaxation is deterministic and updates any fallback activities
    # downstream of the newly anchored recovery choices.
    if anchored_recovery_nodes:
        for _ in range(len(nodes)):
            changed = False
            for source_id in sorted(adjacency, key=lambda value: (rank[value], value)):
                for target_id in adjacency[source_id]:
                    candidate = rank[source_id] + 1
                    if rank[target_id] < candidate:
                        rank[target_id] = candidate
                        changed = True
            if not changed:
                break

    back_edges = {(e["from"], e["to"]) for e in workflow["edges"]
                  if e.get("path") == "recovery" or rank[e["to"]] <= rank[e["from"]]}
    rank = _separate_lane_collisions(rank, nodes, workflow["edges"], lane_index, layout, back_edges,
                                     allow_terminal_stacking=(rendering_mode == "compact" and
                                                              (len(actors) <= 4 or
                                                               continuation_space_reserve <= 0)))
    back_edges = {(e["from"], e["to"]) for e in workflow["edges"]
                  if e.get("path") == "recovery" or rank[e["to"]] <= rank[e["from"]]}
    buckets: dict[tuple[int, int], list[str]] = defaultdict(list)
    for nid, node in nodes.items():
        lane = lane_index.get(node.get("owner"), 0) if layout == "swimlane" else 0
        buckets[(rank[nid], lane)].append(nid)
    ordered_ranks = sorted(set(rank.values()))
    decision_groups = _build_decision_groups(workflow, nodes, rank, ordered_ranks, capacity)
    phase_breaks = _phase_breaks(nodes, rank, ordered_ranks)
    graph_analysis = analyze_workflow_graph(workflow, rank)
    rank_widths = {value: max((_node_size(nodes[node_id], rendering_mode)[0]
                              for node_id in nodes if rank[node_id] == value), default=0.9)
                   for value in ordered_ranks}
    rank_floor_widths = {}
    for value in ordered_ranks:
        floors = []
        for node_id, node in nodes.items():
            if rank[node_id] != value:
                continue
            node_type = node.get("type")
            floors.append(1.48 if node_type in {"end", "error"} else
                          1.55 if node_type == "decision" else 0.84)
        rank_floor_widths[value] = max(floors, default=0.84)
    # Match the narrowest supported page content width. Pages that cannot fit
    # their actual node widths plus routing gaps are rejected before placement.
    width_budget = (11.10 if layout == "swimlane" else 11.55) - max(
        0.0, min(2.0, continuation_space_reserve))
    page_ranks = _paginate_ranks(ordered_ranks, decision_groups, capacity, phase_breaks,
                                 workflow["edges"], rank, rendering_mode,
                                 rank_widths, width_budget, rank_floor_widths,
                                 preferred_slide_count, graph_analysis, allow_adaptive_compression)
    rank_page = {r: (page, column) for page, ranks in enumerate(page_ranks) for column, r in enumerate(ranks)}
    placements = {}
    for (r, lane), ids in sorted(buckets.items()):
        for slot, nid in enumerate(sorted(ids)):
            page, column = rank_page[r]
            placements[nid] = Placement(nid, page, r, lane, slot, column)
    # Repair semantic groups only when pagination actually separates them.
    # First keep non-recovery decision outcomes with their source decision so
    # mutually exclusive branches cannot appear as a backward slide sequence.
    non_recovery_incoming: dict[str, list[str]] = defaultdict(list)
    for edge in workflow["edges"]:
        if edge.get("path") != "recovery":
            non_recovery_incoming[edge["to"]].append(edge["from"])
    composition_moved = False
    page_node_counts = Counter(placement.page for placement in placements.values())
    # Decision-group repair is a preference, not permission to overfill a
    # slide after the width-aware pagination pass. A dense slide leaves no
    # routing corridors and causes distinct branches to collapse onto the same
    # fallback route. Compact mode allows at most the rank capacity in nodes;
    # other modes retain the same conservative bound.
    composition_node_limit = capacity if len(nodes) > 24 else len(nodes)
    for edge in workflow["edges"]:
        if edge.get("path") == "recovery" or nodes[edge["from"]].get("type") != "decision":
            continue
        # A shared merge is downstream of several branches. Pulling it onto
        # one decision's page makes the other valid forward branches appear to
        # run backward across slides. Keep shared destinations at their
        # graph-derived downstream rank instead of treating them as a private
        # immediate outcome.
        if len(nodes) > 24 and len(non_recovery_incoming.get(edge["to"], [])) > 1:
            continue
        source = placements[edge["from"]]; target = placements[edge["to"]]
        if source.page == target.page:
            continue
        if page_node_counts[source.page] >= composition_node_limit:
            continue
        target_rank = source.rank + 1
        occupied = {(placement.rank, placement.lane) for node_id, placement in placements.items()
                    if node_id != edge["to"] and placement.page == source.page}
        while (target_rank, target.lane) in occupied:
            target_rank += 1
        placements[edge["to"]] = Placement(edge["to"], source.page, target_rank,
                                            target.lane, target.slot, source.column + 1)
        page_node_counts[target.page] -= 1
        page_node_counts[source.page] += 1
        composition_moved = True
    # Keep a recovery-only activity with the activity that invokes it. This is
    # safe only when the target has no ordinary incoming path; shared retry
    # destinations such as the original review or provider call remain on
    # their main-process page and are represented as cross-slide recovery.
    if allow_adaptive_compression:
        for edge in workflow["edges"]:
            if edge.get("path") != "recovery" or non_recovery_incoming.get(edge["to"]):
                continue
            source = placements[edge["from"]]; target = placements[edge["to"]]
            if source.page == target.page:
                continue
            target_rank = source.rank if target.lane != source.lane else source.rank + 1
            occupied = {(placement.rank, placement.lane) for node_id, placement in placements.items()
                        if node_id != edge["to"] and placement.page == source.page}
            while (target_rank, target.lane) in occupied:
                target_rank += 1
            placements[edge["to"]] = Placement(edge["to"], source.page, target_rank,
                                                target.lane, target.slot, source.column + 1)
            page_node_counts[target.page] -= 1
            page_node_counts[source.page] += 1
            composition_moved = True
    # Then keep terminal outcomes with a single immediate predecessor. Same-page
    # terminals retain their original geometry, avoiding broad layout changes.
    incoming_for_terminal: dict[str, list[str]] = defaultdict(list)
    for edge in workflow["edges"]:
        if edge.get("path") != "recovery":
            incoming_for_terminal[edge["to"]].append(edge["from"])
    recovery_only_nodes = {
        node_id for node_id in nodes
        if recovery_incoming.get(node_id) and not non_recovery_incoming.get(node_id)
    }
    recovery_terminal_offsets: dict[str, int] = defaultdict(int)
    for node_id, node in sorted(nodes.items()):
        predecessors = incoming_for_terminal.get(node_id, [])
        if (node.get("type") not in {"end", "error"} or len(predecessors) != 1 or
                predecessors[0] not in placements):
            continue
        source = placements[predecessors[0]]; target = placements[node_id]
        if source.page == target.page:
            continue
        if predecessors[0] in recovery_only_nodes:
            # A reconciliation/result decision belongs with the recovery
            # activity that invoked it. Never pull that decision backward to
            # an early terminal page. Keep each immediate terminal beside the
            # decision in a distinct column so the branch remains local and
            # visually traceable.
            offset = recovery_terminal_offsets[predecessors[0]]
            recovery_terminal_offsets[predecessors[0]] += 1
            placements[node_id] = Placement(
                node_id, source.page, source.rank + 1 + offset,
                target.lane, target.slot, source.column + 1 + offset)
            page_node_counts[target.page] -= 1
            page_node_counts[source.page] += 1
            composition_moved = True
            continue
        # If the predecessor already occupies the last usable column, pulling
        # the terminal onto that page would place it beyond the slide margin.
        # Move the single predecessor beside the terminal instead; its incoming
        # relationship remains an explicit cross-slide edge and the terminal
        # keeps a visible local connector.
        last_source_column = max(
            (placement.column for placement in placements.values()
             if placement.page == source.page), default=source.column)
        if source.column >= capacity - 1 or source.column >= last_source_column:
            source_rank = target.rank - 1
            occupied = {(placement.rank, placement.lane)
                        for other_id, placement in placements.items()
                        if other_id != predecessors[0] and placement.page == target.page}
            while (source_rank, source.lane) in occupied:
                source_rank -= 1
            placements[predecessors[0]] = Placement(
                predecessors[0], target.page, source_rank, source.lane, source.slot,
                max(0, target.column - 1))
            page_node_counts[source.page] -= 1
            page_node_counts[target.page] += 1
            composition_moved = True
            continue
        placements[node_id] = Placement(node_id, source.page, source.rank + 1, target.lane,
                                        target.slot, source.column + 1)
        page_node_counts[target.page] -= 1
        page_node_counts[source.page] += 1
        composition_moved = True
    if not allow_adaptive_compression and continuation_space_reserve > 0:
        # Recovery result decisions belong with the activity that produces the
        # result.  The previous recovery implementation always created a new
        # page for these hubs.  That avoided a difficult continuation placement
        # but produced late two-node slides divorced from the review that caused
        # them.  Move the complete local result group beside its invoking
        # activity first; geometry validation can still reject the candidate
        # and let the bounded recovery planner try a less dense composition.
        for recovery_id in sorted(recovery_only_nodes):
            if nodes[recovery_id].get("type") != "decision":
                continue
            outgoing_edges = [edge for edge in workflow["edges"] if edge["from"] == recovery_id]
            if len(outgoing_edges) < 2:
                continue
            terminal_targets = sorted({
                edge["to"] for edge in outgoing_edges
                if nodes[edge["to"]].get("type") in {"end", "error"}
            })
            if not terminal_targets:
                continue
            if len(nodes) < 45:
                # Preserve the established composition of medium workflows.
                # Their recovery hubs already fit the validated continuation
                # model; the consolidation below is reserved for genuinely
                # large graphs where it removes otherwise unjustified sparse
                # trailing slides.
                new_page = max(placement.page for placement in placements.values()) + 1
                source = placements[recovery_id]
                placements[recovery_id] = Placement(
                    recovery_id, new_page, 0, source.lane, source.slot, 0)
                for offset, target_id in enumerate(terminal_targets, start=1):
                    target = placements[target_id]
                    placements[target_id] = Placement(
                        target_id, new_page, offset, target.lane, target.slot, offset)
                composition_moved = True
                continue
            invoking = sorted(recovery_incoming.get(recovery_id, []),
                              key=lambda node_id: (placements[node_id].page,
                                                   placements[node_id].column,
                                                   node_id))
            if not invoking:
                continue
            invocation = placements[invoking[-1]]
            destination_page = invocation.page
            source = placements[recovery_id]
            occupied = {(placement.rank, placement.lane)
                        for node_id, placement in placements.items()
                        if node_id not in {recovery_id, *terminal_targets} and
                        placement.page == destination_page}
            result_rank = invocation.rank + 1
            while (result_rank, source.lane) in occupied:
                result_rank += 1
            placements[recovery_id] = Placement(
                recovery_id, destination_page, result_rank, source.lane,
                source.slot, invocation.column + 1)
            for offset, target_id in enumerate(terminal_targets, start=1):
                target = placements[target_id]
                target_rank = result_rank + offset
                while (target_rank, target.lane) in occupied:
                    target_rank += 1
                occupied.add((target_rank, target.lane))
                placements[target_id] = Placement(
                    target_id, destination_page, target_rank, target.lane,
                    target.slot, invocation.column + 1 + offset)
            composition_moved = True
    # A keep-together repair can make an otherwise valid final page wider than
    # its hard content bounds. For smaller diagrams, move the trailing
    # predecessor/terminal pair to a compact continuation page rather than
    # compressing every shape or clipping the outcome.
    if len(nodes) <= 24 and not allow_adaptive_compression:
        for page in sorted({placement.page for placement in placements.values()}):
            rank_width = {}
            for other_id, placement in placements.items():
                if placement.page == page:
                    rank_width[placement.rank] = max(
                        rank_width.get(placement.rank, 0.0),
                        _node_size(nodes[other_id], rendering_mode)[0])
            required_width = sum(rank_width.values()) + 0.12 * max(0, len(rank_width) - 1)
            if required_width <= width_budget + 1e-6:
                continue
            candidates = []
            for terminal_id, terminal_node in nodes.items():
                terminal = placements.get(terminal_id)
                predecessors = incoming_for_terminal.get(terminal_id, [])
                if (terminal_node.get("type") not in {"end", "error"} or terminal is None or
                        terminal.page != page or len(predecessors) != 1):
                    continue
                predecessor_id = predecessors[0]
                predecessor = placements.get(predecessor_id)
                if predecessor and predecessor.page == page:
                    candidates.append((terminal.column, terminal_id, predecessor_id))
            if not candidates:
                continue
            _, terminal_id, predecessor_id = max(candidates)
            new_page = max(placement.page for placement in placements.values()) + 1
            predecessor = placements[predecessor_id]; terminal = placements[terminal_id]
            placements[predecessor_id] = Placement(
                predecessor_id, new_page, 0, predecessor.lane, predecessor.slot, 0)
            placements[terminal_id] = Placement(
                terminal_id, new_page, 1, terminal.lane, terminal.slot, 1)
            composition_moved = True
    if composition_moved:
        active_pages = sorted({placement.page for placement in placements.values()})
        page_remap = {old: new for new, old in enumerate(active_pages)}
        rebuilt_ranks = []
        for old_page in active_pages:
            values = sorted({placement.rank for placement in placements.values()
                             if placement.page == old_page})
            rebuilt_ranks.append(values)
            columns = {value: index for index, value in enumerate(values)}
            for node_id, placement in list(placements.items()):
                if placement.page == old_page:
                    placements[node_id] = Placement(node_id, page_remap[old_page], placement.rank,
                                                    placement.lane, placement.slot,
                                                    columns[placement.rank])
        page_ranks = rebuilt_ranks
    pages = len(page_ranks) or 1
    split_groups = []
    for group in decision_groups:
        group_pages = {placements[nid].page for nid in group.node_ids if nid in placements}
        if len(group_pages) > 1:
            split_groups.append(group.decision_id)
    split_edges = sum(placements[e["from"]].page != placements[e["to"]].page for e in workflow["edges"])
    unused = sum(max(0, capacity - len(values)) / capacity for values in page_ranks) / pages
    crossings = _estimate_crossings(workflow["edges"], placements, back_edges)
    candidate_scores = (_composition_candidate_score(
        page_ranks, workflow, rank, decision_groups, rank_widths, rank_floor_widths,
        width_budget, preferred_slide_count, graph_analysis),)
    return LayoutResult(placements, pages, back_edges, capacity, page_ranks or [[]],
                        decision_groups, split_groups, rendering_mode, round(unused, 3),
                        split_edges, crossings, graph_analysis, candidate_scores,
                        continuation_space_reserve)


def analyze_workflow_graph(workflow: dict[str, Any], rank: dict[str, int] | None = None) -> dict[str, Any]:
    """Return deterministic, label-independent topology used by composition."""
    nodes = {node["id"]: node for node in workflow["nodes"]}
    edges = list(workflow["edges"])
    outgoing: dict[str, list[dict[str, Any]]] = defaultdict(list)
    incoming: dict[str, list[dict[str, Any]]] = defaultdict(list)
    forward: dict[str, list[str]] = defaultdict(list)
    for edge_id, edge in enumerate(edges):
        record = dict(edge, edge_id=edge_id)
        outgoing[edge["from"]].append(record); incoming[edge["to"]].append(record)
        if edge.get("path") != "recovery":
            forward[edge["from"]].append(edge["to"])
    for values in outgoing.values(): values.sort(key=lambda item: (item["to"], item.get("label", "")))
    for values in incoming.values(): values.sort(key=lambda item: (item["from"], item.get("label", "")))

    def descendants(start: str) -> dict[str, int]:
        distance = {start: 0}; queue = deque([start])
        while queue:
            source = queue.popleft()
            for target in sorted(forward.get(source, [])):
                if target not in distance:
                    distance[target] = distance[source] + 1; queue.append(target)
        return distance

    convergence = sorted(node_id for node_id in nodes if len(incoming[node_id]) > 1)
    forks = []
    for source in sorted(nodes):
        targets = sorted(set(forward.get(source, [])))
        if len(targets) < 2:
            continue
        reach = [descendants(target) for target in targets]
        common = set.intersection(*(set(values) for values in reach)) if reach else set()
        common.discard(source)
        join = min(common, key=lambda node_id: (max(values[node_id] for values in reach),
                                                rank.get(node_id, 10**6) if rank else 0, node_id)) if common else None
        forks.append({"source": source, "targets": targets, "convergence": join,
                      "explicit_parallel": any(edge.get("path") == "parallel"
                                               for edge in outgoing[source])})

    groups = []
    for node_id, node in sorted(nodes.items()):
        member_ids = [node_id]
        reason = None
        if node.get("type") == "decision":
            member_ids += [edge["to"] for edge in outgoing[node_id] if edge.get("path") != "recovery"]
            reason = "decision_outcomes"
        elif node.get("type") in {"end", "error"} and len(incoming[node_id]) == 1:
            member_ids.insert(0, incoming[node_id][0]["from"]); reason = "terminal_predecessor"
        elif node.get("type") == "manual_action":
            resolutions = [edge["to"] for edge in outgoing[node_id]
                           if nodes[edge["to"]].get("type") == "decision"]
            if resolutions:
                member_ids += resolutions; reason = "manual_review_resolution"
        if reason:
            groups.append({"type": reason, "anchor": node_id,
                           "node_ids": list(dict.fromkeys(member_ids))})
    for edge_id, edge in enumerate(edges):
        if edge.get("path") == "recovery":
            groups.append({"type": "recovery_loop", "anchor": edge["from"],
                           "node_ids": [edge["from"], edge["to"]], "edge_ids": [edge_id]})

    stage_members: dict[str, list[str]] = defaultdict(list)
    for node_id, node in nodes.items():
        stage = node.get("phase") or node.get("process_area") or "workflow"
        stage_members[str(stage)].append(node_id)
    return {
        "node_count": len(nodes), "edge_count": len(edges),
        "branch_points": sorted(node_id for node_id in nodes if len(outgoing[node_id]) > 1),
        "convergence_points": convergence,
        "parallel_workstreams": forks,
        "recovery_edge_ids": [index for index, edge in enumerate(edges) if edge.get("path") == "recovery"],
        "manual_review_ids": sorted(node_id for node_id, node in nodes.items()
                                    if node.get("type") == "manual_action"),
        "terminal_ids": sorted(node_id for node_id, node in nodes.items()
                               if node.get("type") in {"end", "error"}),
        "logical_groups": groups,
        "stages": [{"id": stage, "node_ids": sorted(values, key=lambda value: (rank or {}).get(value, 0))}
                   for stage, values in sorted(stage_members.items(),
                                               key=lambda item: min((rank or {}).get(value, 0)
                                                                    for value in item[1]))],
    }


def _separate_lane_collisions(rank, nodes, edges, lane_index, layout, back_edges,
                              allow_terminal_stacking=False):
    """Give nodes sharing a rank and lane their own columns, then preserve forward ordering."""
    result = dict(rank)
    for _ in range(len(nodes)):
        changed = False
        buckets: dict[tuple[int, int], list[str]] = defaultdict(list)
        for nid, node in nodes.items():
            lane = lane_index.get(node.get("owner"), 0) if layout == "swimlane" else 0
            buckets[(result[nid], lane)].append(nid)
        if layout == "swimlane":
            for (bucket_rank, _), ids in sorted(buckets.items()):
                # Terminal outcomes that converge at the same logical point can
                # share a column and stack vertically in their owner lane.  The
                # former one-node-per-column rule pushed short fulfilment paths
                # onto sparse trailing slides.
                if (allow_terminal_stacking and len(ids) <= 3 and
                        all(nodes[nid].get("type") in {"end", "error"} for nid in ids)):
                    continue
                for offset, nid in enumerate(sorted(ids)):
                    if offset and result[nid] < bucket_rank + offset:
                        result[nid] = bucket_rank + offset; changed = True
        for edge in edges:
            pair = (edge["from"], edge["to"])
            if pair not in back_edges and result[edge["to"]] <= result[edge["from"]]:
                result[edge["to"]] = result[edge["from"]] + 1; changed = True
        if not changed: break
    return result


def _build_decision_groups(workflow, nodes, rank, ordered_ranks, capacity):
    rank_position = {value: index for index, value in enumerate(ordered_ranks)}
    outgoing: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for edge in workflow["edges"]:
        outgoing[edge["from"]].append(edge)
    groups = []
    for nid, node in sorted(nodes.items()):
        if node.get("type") != "decision":
            continue
        branches = sorted(outgoing[nid], key=lambda edge: (edge.get("label", ""), edge["to"]))
        # A shared outcome may already be placed earlier because another
        # decision also targets it. It cannot be co-located with both decisions;
        # treating that backward/shared target as part of both pagination units
        # creates sparse pages. Keep the decision with its forward outcomes and
        # represent the shared target with an explicit routed transition.
        forward_targets = [edge["to"] for edge in branches if rank[edge["to"]] >= rank[nid]]
        node_ids = tuple(dict.fromkeys([nid] + forward_targets))
        group_ranks = [rank[value] for value in node_ids]
        low, high = min(group_ranks), max(group_ranks)
        span = rank_position[high] - rank_position[low] + 1
        feasible = span <= capacity and len(node_ids) <= 8
        groups.append(DecisionGroup(nid, node_ids, tuple(str(edge.get("label", "")) for edge in branches),
                                    low, high, feasible))
    return groups


def _phase_breaks(nodes, rank, ordered_ranks):
    phases: dict[int, set[str]] = defaultdict(set)
    for nid, node in nodes.items():
        if node.get("phase"):
            phases[rank[nid]].add(node["phase"])
    breaks = set()
    for left, right in zip(ordered_ranks, ordered_ranks[1:]):
        if phases[left] and phases[right] and phases[left].isdisjoint(phases[right]):
            breaks.add((left, right))
    return breaks


def _paginate_ranks(ranks: list[int], decision_groups: list[DecisionGroup], capacity: int,
                    phase_breaks: set[tuple[int, int]] | None = None, edges=None, rank=None,
                    rendering_mode: str = "compact", rank_widths=None,
                    width_budget: float = 11.10, rank_floor_widths=None,
                    preferred_slide_count: int | None = None,
                    graph_analysis: dict[str, Any] | None = None,
                    allow_adaptive_compression: bool = True) -> list[list[int]]:
    """Globally balance pages while strongly preferring intact feasible decision groups."""
    if not ranks:
        return [[]]
    phase_breaks = phase_breaks or set()
    count = len(ranks)
    best: list[tuple[float, list[list[int]]] | None] = [None] * (count + 1)
    best[count] = (0.0, [])
    for start in range(count - 1, -1, -1):
        candidates = []
        for end in range(start + 1, min(count, start + capacity) + 1):
            candidate_ranks = ranks[start:end]
            if rank_widths:
                required_width = (sum(rank_widths[value] for value in candidate_ranks) +
                                  0.12 * max(0, len(candidate_ranks) - 1))
                floor_width = (sum((rank_floor_widths or rank_widths)[value]
                                   for value in candidate_ranks) +
                               0.12 * max(0, len(candidate_ranks) - 1))
                # The renderer already performs bounded per-node width recovery.
                # Accept a candidate when that established readable floor fits,
                # rather than paginating from its uncompressed estimate alone.
                effective_width = (required_width if (not allow_adaptive_compression or
                                                       (graph_analysis or {}).get("node_count", 0) <= 24)
                                   else floor_width)
                if effective_width > width_budget + 1e-6:
                    continue
            tail = best[end]
            if tail is None:
                continue
            size = end - start
            # A slide has a real cost; compact mode therefore evaluates the complete
            # graph before accepting any page break. Empty area is a smaller penalty.
            slide_cost = {"compact": 90.0, "balanced": 55.0, "detailed": 24.0}[rendering_mode]
            cost = slide_cost + ((capacity - size) / capacity) ** 2 * 18.0
            if end < count:
                left, right = ranks[end - 1], ranks[end]
                crossing_edges = sum(1 for edge in (edges or [])
                                     if rank[edge["from"]] <= left and rank[edge["to"]] >= right)
                cost += crossing_edges * 12.0
                for group in decision_groups:
                    if group.min_rank <= left and group.max_rank >= right:
                        cost += 1000.0 if group.feasible else 35.0
                if (left, right) in phase_breaks:
                    cost -= 4.0
                if any(
                    group.feasible and group.min_rank == right
                    for group in decision_groups
                ):
                    cost -= 3.0
            candidates.append((cost + tail[0], [candidate_ranks] + tail[1]))
        if not candidates:
            # A single oversized rank is still emitted so geometry validation
            # can report the exact offending element instead of crashing.
            tail = best[start + 1]
            best[start] = (10000.0 + (tail[0] if tail else 0.0),
                           [[ranks[start]]] + (tail[1] if tail else []))
        else:
            best[start] = min(candidates, key=lambda item: (item[0], len(item[1]), item[1]))
    pages = best[0][1]
    # Preferred slide count is soft. If the normal optimum exceeds it, compare
    # deterministic contiguous partitions at the preference and one slide above
    # it. Invalid width floors are rejected by the common scorer.
    alternatives = [pages]
    for page_count in sorted({preferred_slide_count, (preferred_slide_count + 1)
                              if preferred_slide_count else None} - {None}):
        if 0 < page_count <= len(ranks):
            for offsets in _balanced_rank_partitions(ranks, page_count, capacity):
                alternatives.append(offsets)
    scored = [(_composition_candidate_score(candidate, {"edges": edges or []}, rank,
                                               decision_groups, rank_widths or {},
                                               rank_floor_widths or rank_widths or {}, width_budget,
                                               preferred_slide_count, graph_analysis or {}), candidate)
              for candidate in alternatives]
    valid = [(score, candidate) for score, candidate in scored if score["valid"]]
    return min(valid, key=lambda item: (item[0]["total_cost"], len(item[1]), item[1]))[1] if valid else pages


def _balanced_rank_partitions(ranks: list[int], page_count: int, capacity: int) -> list[list[list[int]]]:
    """Return a small deterministic set of balanced contiguous partitions."""
    if page_count <= 0 or page_count > len(ranks):
        return []
    results = []
    base = len(ranks) // page_count; remainder = len(ranks) % page_count
    for bias in range(min(page_count, 3)):
        sizes = [base + (1 if (index - bias) % page_count < remainder else 0)
                 for index in range(page_count)]
        if any(size <= 0 or size > capacity for size in sizes):
            continue
        pages = []; cursor = 0
        for size in sizes:
            pages.append(ranks[cursor:cursor + size]); cursor += size
        if pages not in results:
            results.append(pages)
    return results


def _composition_candidate_score(pages, workflow, rank, decision_groups, rank_widths,
                                 rank_floor_widths, width_budget, preferred_slide_count,
                                 graph_analysis):
    page_of = {value: page for page, values in enumerate(pages) for value in values}
    valid = all((sum(rank_floor_widths.get(value, 0.84) for value in values) +
                 0.12 * max(0, len(values) - 1)) <= width_budget + 1e-6
                for values in pages)
    cross_edges = sum(page_of.get(rank[edge["from"]]) != page_of.get(rank[edge["to"]])
                      for edge in workflow.get("edges", []))
    split_decisions = sum(len({page_of.get(rank[node_id]) for node_id in group.node_ids
                               if node_id in rank}) > 1 for group in decision_groups)
    split_recovery = 0; split_terminals = 0
    for group in graph_analysis.get("logical_groups", []):
        group_pages = {page_of.get(rank[node_id]) for node_id in group["node_ids"] if node_id in rank}
        if len(group_pages) > 1:
            split_recovery += group["type"] == "recovery_loop"
            split_terminals += group["type"] == "terminal_predecessor"
    utilization = []
    for values in pages:
        used = sum(rank_widths.get(value, 0.84) for value in values) + 0.12 * max(0, len(values) - 1)
        utilization.append(min(1.0, used / width_budget))
    sparse = sum(value < 0.38 for value in utilization)
    preference_gap = abs(len(pages) - preferred_slide_count) if preferred_slide_count else 0
    total = (len(pages) * 95 + cross_edges * 13 + split_decisions * 90 +
             split_recovery * 70 + split_terminals * 65 + sparse * 80 +
             preference_gap * 12 + sum((1 - value) ** 2 for value in utilization) * 18)
    return {"valid": valid, "slide_count": len(pages), "nodes_per_slide": [],
            "slide_utilization": [round(value, 3) for value in utilization],
            "cross_slide_edges": cross_edges, "split_decision_groups": split_decisions,
            "split_recovery_loops": split_recovery, "isolated_terminal_groups": split_terminals,
            "sparse_slides": sparse, "preferred_slide_count": preferred_slide_count,
            "total_cost": round(total, 3)}


def _adaptive_layout_candidate_score(layout_result: LayoutResult,
                                     workflow: dict[str, Any],
                                     preferred_slide_count: int | None) -> dict[str, Any]:
    """Score a fully repaired layout before the expensive render/validation pass."""
    counts = [sum(placement.page == page
                  for placement in layout_result.placements.values())
              for page in range(layout_result.pages)]
    single_node = sum(count == 1 for count in counts)
    sparse = sum(count <= 2 for count in counts)
    cross_edges = sum(
        layout_result.placements[edge["from"]].page !=
        layout_result.placements[edge["to"]].page
        for edge in workflow.get("edges", [])
        if edge["from"] in layout_result.placements and edge["to"] in layout_result.placements
    )
    split_groups = 0
    for group in layout_result.graph_analysis.get("logical_groups", []):
        pages = {layout_result.placements[node_id].page
                 for node_id in group.get("node_ids", [])
                 if node_id in layout_result.placements}
        split_groups += len(pages) > 1
    mean = sum(counts) / max(1, len(counts))
    imbalance = sum((count - mean) ** 2 for count in counts) / max(1, len(counts))
    preference_gap = abs(layout_result.pages - preferred_slide_count) if preferred_slide_count else 0
    total = (single_node * 1200 + sparse * 500 + split_groups * 85 +
             cross_edges * 18 + layout_result.pages * 95 +
             layout_result.unused_space * 160 + imbalance * 4 + preference_gap * 10)
    return {
        "total_cost": round(total, 3), "slide_count": layout_result.pages,
        "nodes_per_slide": counts, "single_node_slides": single_node,
        "sparse_slides": sparse, "cross_slide_edges": cross_edges,
        "split_logical_groups": split_groups,
        "unused_space": layout_result.unused_space,
    }


def _estimate_crossings(edges, placements, back_edges):
    """Count rank-order inversions between forward edges on the same page."""
    forward = [e for e in edges if (e["from"], e["to"]) not in back_edges and
               placements[e["from"]].page == placements[e["to"]].page]
    crossings = 0
    for index, first in enumerate(forward):
        a1, a2 = placements[first["from"]], placements[first["to"]]
        for second in forward[index + 1:]:
            b1, b2 = placements[second["from"]], placements[second["to"]]
            if (a1.page == b1.page and a1.rank == b1.rank and a2.rank == b2.rank and
                    (a1.lane - b1.lane) * (a2.lane - b2.lane) < 0):
                crossings += 1
    return crossings

# ---- Consolidated from renderer/process_map.py ----
from collections import defaultdict, deque
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class ProcessArea:
    id: str
    label: str
    order: int
    y: float
    h: float


@dataclass(frozen=True)
class ProcessMapNode:
    node_id: str
    area_id: str
    row: int
    column: int
    x: float
    y: float
    w: float
    h: float


@dataclass
class ProcessMapPlan:
    areas: list[ProcessArea]
    nodes: dict[str, ProcessMapNode]
    warnings: list[str]
    unused_space: float
    long_paths: int = 0
    small_shapes: int = 0


@dataclass(frozen=True)
class ProcessMapComplexity:
    score: float
    node_count: int
    edge_count: int
    decision_count: int
    process_area_count: int
    branch_points: int
    merge_points: int
    cycle_edges: int
    estimated_text_units: int


@dataclass(frozen=True)
class ProcessMapPagination:
    pages: tuple[tuple[str, ...], ...]
    page_labels: tuple[str, ...]
    complexity: ProcessMapComplexity
    split_edges: int
    warnings: tuple[str, ...]


def analyze_process_map_complexity(workflow: dict[str, Any]) -> ProcessMapComplexity:
    nodes = workflow["nodes"]; edges = workflow["edges"]
    outgoing = defaultdict(int); incoming = defaultdict(int)
    for edge in edges:
        outgoing[edge["from"]] += 1; incoming[edge["to"]] += 1
    decisions = sum(node.get("type") == "decision" for node in nodes)
    branches = sum(value > 1 for value in outgoing.values())
    merges = sum(value > 1 for value in incoming.values())
    cycles = sum(edge.get("path") == "recovery" for edge in edges)
    text_units = sum(max(1, (len(node.get("label", "")) + len(node.get("description", ""))) // 24)
                     for node in nodes)
    areas = len(workflow.get("process_areas") or workflow.get("actors") or [])
    score = (len(nodes) + len(edges) * 0.32 + decisions * 1.25 + areas * 0.8 +
             branches * 0.7 + merges * 0.45 + cycles * 1.5 + text_units * 0.22)
    return ProcessMapComplexity(round(score, 2), len(nodes), len(edges), decisions, areas,
                                branches, merges, cycles, text_units)


def paginate_process_map(workflow: dict[str, Any], density: str = "compact") -> ProcessMapPagination:
    """Partition a process map before geometry is calculated."""
    complexity = analyze_process_map_complexity(workflow)
    nodes = workflow["nodes"]
    single_limit = {"compact": 48.0, "balanced": 40.0, "detailed": 32.0}[density]
    node_limit = {"compact": 18, "balanced": 15, "detailed": 12}[density]
    if complexity.score <= single_limit and len(nodes) <= node_limit:
        return ProcessMapPagination((tuple(node["id"] for node in nodes),), ("Complete workflow",),
                                    complexity, 0, ())

    phase_defs = workflow.get("phases") or []
    phase_order = [phase["id"] for phase in phase_defs]
    phase_labels = {phase["id"]: phase.get("label", phase["id"]) for phase in phase_defs}
    assigned = {node.get("phase") for node in nodes if node.get("phase")}
    use_phases = bool(phase_order) and all(node.get("phase") in phase_order for node in nodes)
    groups: list[tuple[str, list[str]]] = []
    if use_phases:
        for phase_id in phase_order:
            members = [node["id"] for node in nodes if node.get("phase") == phase_id]
            if members:
                groups.append((phase_labels[phase_id], members))
    else:
        ranks = _graph_ranks(workflow)
        ordered = sorted(nodes, key=lambda node: (ranks[node["id"]], node["id"]))
        fallback_limit = min(node_limit, max(4, int(node_limit * single_limit / max(single_limit, complexity.score))))
        for index in range(0, len(ordered), fallback_limit):
            groups.append((f"Steps {index + 1}–{min(len(ordered), index + fallback_limit)}",
                           [node["id"] for node in ordered[index:index + fallback_limit]]))

    pages: list[list[str]] = []; labels: list[str] = []
    for label, members in groups:
        group_complexity = analyze_process_map_complexity(_workflow_view(workflow, set(members)))
        effective_limit = min(node_limit, max(4, int(node_limit * single_limit /
                              max(single_limit, group_complexity.score))))
        if len(members) > effective_limit:
            chunks = _split_group_with_decisions(workflow, members, effective_limit)
            for part, chunk in enumerate(chunks, 1):
                pages.append(chunk); labels.append(f"{label} · {part} of {len(chunks)}")
            continue
        if pages and len(pages[-1]) + len(members) <= node_limit:
            trial_ids = set(pages[-1] + members)
            trial = _workflow_view(workflow, trial_ids)
            if analyze_process_map_complexity(trial).score <= single_limit:
                pages[-1].extend(members); labels[-1] += f" / {label}"; continue
        pages.append(list(members)); labels.append(label)

    # Keep terminal outcomes beside an immediate predecessor when capacity
    # permits. This is a visual move only; the source node and edge remain
    # unchanged and each node is still rendered exactly once.
    node_by_id = {node["id"]: node for node in nodes}
    incoming: dict[str, list[str]] = defaultdict(list)
    for edge in workflow["edges"]:
        if edge.get("path") != "recovery":
            incoming[edge["to"]].append(edge["from"])
    changed = True
    while changed:
        changed = False
        page_of = {node_id: page for page, values in enumerate(pages) for node_id in values}
        for node_id, node in node_by_id.items():
            if node.get("type") not in {"end", "error"} or len(incoming[node_id]) != 1:
                continue
            source = incoming[node_id][0]
            if source not in page_of or node_id not in page_of or page_of[source] == page_of[node_id]:
                continue
            source_page, terminal_page = page_of[source], page_of[node_id]
            if len(pages[source_page]) >= node_limit:
                continue
            pages[terminal_page].remove(node_id)
            insert_at = pages[source_page].index(source) + 1
            pages[source_page].insert(insert_at, node_id)
            changed = True
            break
    retained = [(page, label) for page, label in zip(pages, labels) if page]
    pages = [page for page, _ in retained]
    labels = [label for _, label in retained]

    page_of = {node_id: page for page, values in enumerate(pages) for node_id in values}
    split_edges = sum(page_of[edge["from"]] != page_of[edge["to"]] for edge in workflow["edges"])
    warnings = (f"Process map paginated before routing: complexity score {complexity.score} exceeds "
                f"the {density} one-slide limit {single_limit}.",)
    return ProcessMapPagination(tuple(tuple(values) for values in pages), tuple(labels), complexity,
                                split_edges, warnings)


def process_map_page_view(workflow: dict[str, Any], node_ids: set[str]) -> dict[str, Any]:
    return _workflow_view(workflow, node_ids)


def _workflow_view(workflow, node_ids):
    view = dict(workflow)
    view["nodes"] = [node for node in workflow["nodes"] if node["id"] in node_ids]
    view["edges"] = [edge for edge in workflow["edges"]
                     if edge["from"] in node_ids and edge["to"] in node_ids]
    active_areas = {((node.get("layout_hint") or {}).get("preferred_process_area") or
                    node.get("process_area") or node.get("lane")) for node in view["nodes"]}
    if workflow.get("process_areas"):
        view["process_areas"] = [area for area in workflow["process_areas"] if area["id"] in active_areas]
    return view


def _split_group_with_decisions(workflow, members, limit):
    member_set = set(members); outgoing = defaultdict(list)
    for edge in workflow["edges"]:
        if edge["from"] in member_set and edge["to"] in member_set:
            outgoing[edge["from"]].append(edge["to"])
    units = []; consumed = set()
    for node_id in members:
        if node_id in consumed:
            continue
        node = next(node for node in workflow["nodes"] if node["id"] == node_id)
        unit = [node_id]
        if node.get("type") == "decision":
            unit.extend(target for target in outgoing[node_id] if target not in consumed)
        unit = list(dict.fromkeys(unit)); consumed.update(unit); units.append(unit)
    pages = []; current = []
    for unit in units:
        if current and len(current) + len(unit) > limit:
            pages.append(current); current = []
        current.extend(unit)
    if current: pages.append(current)
    return pages


def plan_process_map(workflow: dict[str, Any], density: str = "compact",
                     spacing: str = "standard") -> ProcessMapPlan:
    """Plan one landscape slide using process areas and multiple rows per area."""
    nodes = {node["id"]: node for node in workflow["nodes"]}
    areas, warnings = _areas(workflow)
    area_ids = {area["id"] for area in areas}
    rank = _graph_ranks(workflow)
    membership: dict[str, str] = {}
    for node_id, node in nodes.items():
        hint = node.get("layout_hint") or {}
        requested = hint.get("preferred_process_area") or node.get("process_area") or node.get("lane")
        if requested and requested not in area_ids:
            warnings.append(f"Node {node_id} references unknown process area {requested}; owner mapping was used.")
            requested = None
        owner_area = node.get("owner") if node.get("owner") in area_ids else None
        membership[node_id] = requested or owner_area or areas[0]["id"]
        direction = hint.get("preferred_direction")
        if direction and direction not in {"right", "left", "up", "down"}:
            warnings.append(f"Node {node_id} has unsupported preferred_direction {direction}.")

    assigned_rows: dict[str, int] = {}
    group_rows: dict[tuple[str, str], int] = {}
    for node_id in sorted(nodes, key=lambda value: (rank[value], value)):
        node = nodes[node_id]; area_id = membership[node_id]; hint = node.get("layout_hint") or {}
        preferred = hint.get("preferred_row")
        if preferred is not None and (not isinstance(preferred, int) or preferred < 1 or preferred > 8):
            warnings.append(f"Node {node_id} has invalid preferred_row {preferred}; automatic placement was used.")
            preferred = None
        group = hint.get("keep_together") or hint.get("group")
        if group and (area_id, str(group)) in group_rows:
            row = group_rows[(area_id, str(group))]
        else:
            row = (preferred - 1) if preferred else rank[node_id] // 5
            if group:
                group_rows[(area_id, str(group))] = row
        assigned_rows[node_id] = row

    area_rows = {area["id"]: max((assigned_rows[nid] for nid in nodes if membership[nid] == area["id"]), default=0) + 1
                 for area in areas}
    weights = [max(1.0, area_rows[area["id"]] * 0.92) for area in areas]
    top, bottom = 1.02, 7.22
    total = sum(weights); cursor = top; planned_areas = []
    for area, weight in zip(areas, weights):
        height = (bottom - top) * weight / total
        planned_areas.append(ProcessArea(area["id"], area["label"], area["order"], cursor, height))
        cursor += height

    node_plans: dict[str, ProcessMapNode] = {}
    bounds = {"compact": (1.12, 13.00), "standard": (1.25, 12.92), "spacious": (1.42, 12.75)}
    left, right = bounds[spacing]
    size_scale = {"compact": 0.92, "standard": 1.0, "spacious": 1.08}[spacing]
    for area in planned_areas:
        row_count = area_rows[area.id]
        row_height = area.h / row_count
        for row in range(row_count):
            members = [nid for nid in nodes if membership[nid] == area.id and assigned_rows[nid] == row]
            members.sort(key=lambda value: (rank[value], value))
            if not members:
                continue
            gap = (right - left) / max(1, len(members))
            for index, node_id in enumerate(members):
                node = nodes[node_id]; w, h = process_map_node_size(node, density)
                w *= size_scale; h *= size_scale
                x = left + (index + 0.5) * gap - w / 2
                y = area.y + row * row_height + (row_height - h) / 2
                node_plans[node_id] = ProcessMapNode(node_id, area.id, row, index, x, y, w, h)

    used = sum(plan.w * plan.h for plan in node_plans.values())
    usable = (right - left) * (bottom - top)
    unused = max(0.0, 1.0 - used / usable)
    small = sum(plan.w < 0.72 or plan.h < 0.34 for plan in node_plans.values())
    return ProcessMapPlan(planned_areas, node_plans, warnings, round(unused, 3), small_shapes=small)


def process_map_node_size(node: dict[str, Any], density: str) -> tuple[float, float]:
    length = len(node.get("label", "")) + len(node.get("description", ""))
    if node.get("type") == "decision":
        return min(1.42, max(1.00, 0.88 + length * 0.016)), min(1.00, max(0.90, 0.78 + length * 0.008))
    if node.get("type") in {"start", "end", "error"}:
        # Terminal ellipses lose substantially more usable text area than a
        # rectangle. Size from the visible label and its longest whole word,
        # then reserve enough height for two or three centered lines.
        label = str(node.get("label", ""))
        words = label.split()
        longest = max((len(word) for word in words), default=1)
        width = min(2.05, max(1.18, 0.76 + len(label) * 0.026, 0.48 + longest * 0.09))
        chars_per_line = max(longest, int(width * 10.0))
        lines = max(1, math.ceil(max(1, len(label)) / max(1, chars_per_line)))
        height = min(0.86, max(0.62, 0.43 + lines * 0.13))
        if density == "detailed":
            return width * 1.06, height * 1.06
        return width, height
    width = min(1.62, max(0.88, 0.76 + length * 0.012))
    height = min(0.76, max(0.56, 0.44 + length * 0.005))
    if density == "detailed":
        return width * 1.08, height * 1.08
    return width, height


def _areas(workflow):
    warnings = []
    raw = workflow.get("process_areas")
    if raw:
        areas = sorted(raw, key=lambda area: (area.get("order", 0), area.get("id", "")))
    else:
        areas = [{"id": actor["id"], "label": actor.get("label", actor["id"]), "order": index + 1}
                 for index, actor in enumerate(workflow["actors"])]
        warnings.append("process_areas were not provided; actor definitions were used as process areas.")
    return areas, warnings


def _graph_ranks(workflow):
    nodes = {node["id"] for node in workflow["nodes"]}
    adjacency = defaultdict(list); indegree = {node_id: 0 for node_id in nodes}
    for edge in workflow["edges"]:
        if edge.get("path") == "recovery":
            continue
        adjacency[edge["from"]].append(edge["to"]); indegree[edge["to"]] += 1
    queue = deque(sorted(node_id for node_id, degree in indegree.items() if degree == 0))
    rank = {node_id: 0 for node_id in queue}
    while queue:
        source = queue.popleft()
        for target in sorted(adjacency[source]):
            rank[target] = max(rank.get(target, 0), rank[source] + 1)
            indegree[target] -= 1
            if indegree[target] == 0:
                queue.append(target)
    for node_id in sorted(nodes):
        rank.setdefault(node_id, max(rank.values(), default=0) + 1)
    return rank

# ---- Consolidated from renderer/pptx_renderer.py ----
from collections import defaultdict
from dataclasses import dataclass
import heapq
import math
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


SLIDE_W, SLIDE_H = 13.333, 7.5
ROUTING_ENGINE_VERSION = "1.2"


@dataclass(frozen=True)
class Box:
    x: float; y: float; w: float; h: float
    @property
    def right(self): return self.x + self.w
    @property
    def bottom(self): return self.y + self.h
    @property
    def cx(self): return self.x + self.w / 2
    @property
    def cy(self): return self.y + self.h / 2


def _diamond_boundary_port(box: Box, side: str, offset: float = 0.0) -> tuple[float, float]:
    """Return a point on a diamond edge, never on its rectangular envelope.

    ``offset`` is normalized to the half-axis parallel to the requested side.
    Zero selects a vertex. Values in (-1, 1) select a point on one of the two
    sloping edges while preserving an orthogonal outward escape direction.
    """
    offset = max(-0.82, min(0.82, float(offset)))
    half_w, half_h = box.w / 2, box.h / 2
    if side in {"right", "left"}:
        y = box.cy + offset * half_h
        x_delta = half_w * (1.0 - abs(offset))
        return ((box.cx + x_delta if side == "right" else box.cx - x_delta), y)
    if side in {"top", "bottom"}:
        x = box.cx + offset * half_w
        y_delta = half_h * (1.0 - abs(offset))
        return (x, (box.cy - y_delta if side == "top" else box.cy + y_delta))
    raise ValueError(f"Unknown diamond side: {side}")


def _diamond_boundary_value(point: tuple[float, float], box: Box) -> float:
    """Return the normalized L1 radius of a point relative to a diamond."""
    half_w, half_h = max(box.w / 2, 1e-9), max(box.h / 2, 1e-9)
    return abs(point[0] - box.cx) / half_w + abs(point[1] - box.cy) / half_h


def _point_on_diamond_boundary(point: tuple[float, float], box: Box, tolerance: float = 0.025) -> bool:
    return abs(_diamond_boundary_value(point, box) - 1.0) <= tolerance


def _point_inside_diamond(point: tuple[float, float], box: Box, clearance: float = 0.0) -> bool:
    return _diamond_boundary_value(point, box) < 1.0 - clearance


def _box_overlaps_diamond(candidate: Box, diamond: Box, gap: float = 0.0) -> bool:
    """Conservatively test a text/arrow box against the visible diamond polygon."""
    expanded = Box(candidate.x - gap, candidate.y - gap,
                   candidate.w + 2 * gap, candidate.h + 2 * gap)
    points = (
        (expanded.x, expanded.y), (expanded.right, expanded.y),
        (expanded.x, expanded.bottom), (expanded.right, expanded.bottom),
        (expanded.cx, expanded.cy),
    )
    if any(_point_inside_diamond(point, diamond, -0.01) for point in points):
        return True
    vertices = ((diamond.cx, diamond.y), (diamond.right, diamond.cy),
                (diamond.cx, diamond.bottom), (diamond.x, diamond.cy))
    return any(expanded.x <= x <= expanded.right and expanded.y <= y <= expanded.bottom
               for x, y in vertices)


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def _text_fit_estimate(text: str, width: float, height: float, size: float,
                       padding_x: float, padding_y: float, *, diamond: bool = False,
                       bold: bool = False) -> dict[str, Any]:
    """Conservatively estimate whole-word wrapping inside a PowerPoint shape."""
    safe_width = max(0.05, width - 2 * padding_x)
    safe_height = max(0.05, height - 2 * padding_y)
    if diamond:
        # A diamond's useful center is much smaller than its bounding box.
        safe_width *= 0.68
        safe_height *= 0.66
    char_factor = 0.56 + (0.08 if diamond else 0.0) + (0.035 if bold else 0.0)
    char_width = max(0.035, size / 72.0 * char_factor)
    line_height = max(0.11, size / 72.0)
    capacity = max(1, int(safe_width / char_width))
    words = str(text).replace("\n", " \n ").split()
    lines: list[str] = []; current = ""; fragmented = False
    for word in words:
        if word == "\n":
            if current: lines.append(current); current = ""
            continue
        if len(word) > capacity:
            fragmented = True
        candidate = word if not current else current + " " + word
        if current and len(candidate) > capacity:
            lines.append(current); current = word
        else:
            current = candidate
    if current or not lines:
        lines.append(current)
    required_height = len(lines) * line_height
    return {"fits": not fragmented and required_height <= safe_height + 0.01,
            "lines": lines, "line_count": len(lines), "capacity": capacity,
            "required_height": round(required_height, 3),
            "safe_width": round(safe_width, 3), "safe_height": round(safe_height, 3),
            "word_fragmentation_risk": fragmented}


def _fitted_font_size(shape, text: str, requested: float, padding_x: float, padding_y: float,
                      *, diamond: bool, bold: bool, minimum: float = 8.0) -> tuple[float, dict[str, Any]]:
    width, height = shape.width / 914400.0, shape.height / 914400.0
    size = float(requested)
    estimate = _text_fit_estimate(text, width, height, size, padding_x, padding_y,
                                  diamond=diamond, bold=bold)
    while not estimate["fits"] and size > minimum:
        size = max(minimum, size - 0.5)
        estimate = _text_fit_estimate(text, width, height, size, padding_x, padding_y,
                                      diamond=diamond, bold=bold)
    return size, estimate


def _set_text(shape, text: str, theme: Theme, size: float = 12, bold: bool = False,
              color: str | None = None, style: dict[str, Any] | None = None, role: str | None = None,
              explicit_size: bool = False):
    style = style or {}; typography = style.get("typography", {})
    role_sizes = {"title": "title_size", "subtitle": "subtitle_size", "activity": "activity_size",
                  "decision": "decision_size", "branch_label": "branch_label_size",
                  "swimlane": "swimlane_label_size"}
    if role in role_sizes and not explicit_size:
        size = typography.get(role_sizes[role], size)
    frame = shape.text_frame; frame.clear(); frame.word_wrap = True
    padding = typography.get("text_padding", style.get("shapes", {}).get("shape_padding", 0.08))
    vertical_padding = min(padding, 0.08)
    if role in {"title", "subtitle", "branch_label", "swimlane"}:
        # Text boxes already have explicit placement; use compact vertical
        # margins while retaining the configured horizontal breathing room.
        vertical_padding = 0.02
    if role == "decision":
        margin_map = {"compact": (0.04, 0.02), "standard": (0.08, 0.04), "generous": (0.13, 0.07)}
        padding, vertical_padding = margin_map[style.get("decision_padding", "standard")]
        custom_padding = typography.get("text_padding")
        if custom_padding is not None:
            padding, vertical_padding = custom_padding, min(custom_padding, 0.10)
    fitted_size, fit = _fitted_font_size(shape, text, size, padding, vertical_padding,
                                         diamond=role == "decision", bold=bold)
    frame.margin_left = frame.margin_right = Inches(padding)
    frame.margin_top = frame.margin_bottom = Inches(vertical_padding); frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    alignments = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    p = frame.paragraphs[0]; p.alignment = alignments.get(typography.get("text_alignment"), PP_ALIGN.CENTER)
    run = p.add_run(); run.text = text; run.font.name = typography.get("font_family", theme.font); run.font.size = Pt(fitted_size)
    run.font.bold = typography.get("font_weight", "bold" if bold else "regular") == "bold"
    run.font.color.rgb = _rgb(color or theme.text)
    shape._element.nvSpPr.cNvPr.set("textFit", json.dumps({"requested": size, "applied": fitted_size,
                                                            **fit}, separators=(",", ":")))


def _decision_text_style(shape, style):
    sizes = {"compact": 8, "standard": 9, "large": 10}
    margins = {"compact": (0.04, 0.02), "standard": (0.08, 0.04), "generous": (0.13, 0.07)}
    horizontal, vertical = margins[style["decision_padding"]]
    custom_padding = style.get("typography", {}).get("text_padding")
    if custom_padding is not None:
        horizontal, vertical = custom_padding, min(custom_padding, 0.10)
    shape.text_frame.margin_left = shape.text_frame.margin_right = Inches(horizontal)
    shape.text_frame.margin_top = shape.text_frame.margin_bottom = Inches(vertical)
    return style.get("typography", {}).get("decision_size", sizes[style["decision_typography"]])


def _presentation_text_audit(prs) -> dict[str, Any]:
    """Audit the conservative fit metadata embedded on editable text shapes."""
    records = []; critical = 0; diamond_violations = 0; small_fonts = 0
    for slide_number, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False) or not shape.text.strip():
                continue
            raw = shape._element.nvSpPr.cNvPr.get("textFit")
            if not raw:
                continue
            try:
                fit = json.loads(raw)
            except (TypeError, json.JSONDecodeError):
                continue
            is_node = shape.name.startswith("workflow-node::")
            try:
                is_diamond = shape.auto_shape_type == MSO_SHAPE.DIAMOND
            except (ValueError, AttributeError):
                is_diamond = False
            clipped = not bool(fit.get("fits", True))
            critical += int(is_node and clipped)
            diamond_violations += int(is_diamond and clipped)
            small_fonts += int(is_node and float(fit.get("applied", 99)) < 8.0)
            records.append({"slide": slide_number, "shape_name": shape.name,
                            "shape_id": shape.shape_id, "text": shape.text,
                            "is_workflow_node": is_node, "is_diamond": is_diamond,
                            **fit})
    return {"status": "clean" if critical == 0 else "critical",
            "critical_text_clipping": critical,
            "diamond_safe_region_violations": diamond_violations,
            "fonts_below_minimum": small_fonts,
            "measurement_method": "conservative font-size and whole-word estimate; preview review remains required",
            "records": records}


def _actual_shape_properties(shape) -> dict[str, Any]:
    runs = [run for paragraph in shape.text_frame.paragraphs for run in paragraph.runs] if shape.has_text_frame else []
    return {"border_width": round(shape.line.width.pt, 3) if shape.line.width is not None else None,
            "border": str(shape.line.color.rgb) if shape.line.color.type is not None else None,
            "fill": str(shape.fill.fore_color.rgb) if shape.fill.type is not None and shape.fill.fore_color.type is not None else None,
            "text_color": (str(runs[0].font.color.rgb) if runs and runs[0].font.color.type is not None else None),
            "text_size": (round(runs[0].font.size.pt, 3) if runs and runs[0].font.size is not None else None)}


def _verify_element_styles(pptx_path: str | Path, style: dict[str, Any]) -> list[dict[str, Any]]:
    """Reopen the written artifact and prove each requested element property reached it."""
    requested = style.get("elements", {})
    if not requested:
        return []
    reopened = Presentation(str(pptx_path)); results = []
    for node_id, properties in sorted(requested.items()):
        matches = [(slide_number, shape) for slide_number, slide in enumerate(reopened.slides, start=1)
                   for shape in slide.shapes if shape.name == f"workflow-node::{node_id}"]
        actual = _actual_shape_properties(matches[0][1]) if len(matches) == 1 else {}
        checks = {}
        for key, expected in properties.items():
            actual_value = actual.get(key)
            if key in {"border_width", "text_size"}:
                checks[key] = actual_value is not None and abs(float(actual_value) - float(expected)) <= 0.02
            else:
                checks[key] = actual_value == expected
        results.append({"requested_style": properties, "resolved_source_node_id": node_id,
                        "rendered_shape_identifier": f"workflow-node::{node_id}",
                        "rendered_shape_id": matches[0][1].shape_id if len(matches) == 1 else None,
                        "slide_number": matches[0][0] if len(matches) == 1 else None,
                        "actual_powerpoint_properties": actual,
                        "property_verification": checks,
                        "verification_result": "verified" if len(matches) == 1 and all(checks.values()) else "failed"})
    return results


def _pptx_connector_endpoints(shape) -> tuple[tuple[float, float], tuple[float, float]]:
    """Recover directed connector endpoints from a saved PowerPoint segment."""
    xfrm = shape._element.spPr.xfrm
    flip_h = str(xfrm.get("flipH", "0")).lower() in {"1", "true"}
    flip_v = str(xfrm.get("flipV", "0")).lower() in {"1", "true"}
    emu = 914400.0
    left, top = shape.left / emu, shape.top / emu
    right, bottom = (shape.left + shape.width) / emu, (shape.top + shape.height) / emu
    start = (right if flip_h else left, bottom if flip_v else top)
    end = (left if flip_h else right, top if flip_v else bottom)
    return start, end


def _verify_pptx_diamond_attachments(pptx_path: str | Path, nodes: dict[str, dict[str, Any]],
                                     indexed_edges: list[dict[str, Any]],
                                     slide_limit: int | None = None) -> dict[str, Any]:
    """Reopen the artifact and verify native connector geometry at every local diamond."""
    reopened = Presentation(str(pptx_path)); records = []; errors = []
    edge_by_index = {int(edge.get("_render_index", index)): edge
                     for index, edge in enumerate(indexed_edges)}
    slides = list(reopened.slides)[:slide_limit] if slide_limit is not None else list(reopened.slides)
    for slide_number, slide in enumerate(slides, start=1):
        node_shapes = {shape.name.removeprefix("workflow-node::"): shape for shape in slide.shapes
                       if shape.name.startswith("workflow-node::")}
        grouped = defaultdict(list)
        for shape in slide.shapes:
            if not shape.name.startswith("workflow-edge::"):
                continue
            parts = shape.name.split("::")
            if len(parts) < 3 or not parts[1].isdigit() or "segment-" not in parts[2]:
                continue
            grouped[int(parts[1])].append(shape)
        for edge_index, shapes in grouped.items():
            edge = edge_by_index.get(edge_index)
            if edge is None or edge["from"] not in node_shapes or edge["to"] not in node_shapes:
                continue
            shapes.sort(key=lambda shape: int(shape.name.rsplit("-", 1)[1]))
            source_shape, target_shape = node_shapes[edge["from"]], node_shapes[edge["to"]]
            source_box = Box(source_shape.left / 914400.0, source_shape.top / 914400.0,
                             source_shape.width / 914400.0, source_shape.height / 914400.0)
            target_box = Box(target_shape.left / 914400.0, target_shape.top / 914400.0,
                             target_shape.width / 914400.0, target_shape.height / 914400.0)
            first_start, first_end = _pptx_connector_endpoints(shapes[0])
            last_start, last_end = _pptx_connector_endpoints(shapes[-1])
            checks = {}
            if nodes[edge["from"]].get("type") == "decision":
                checks["source_on_boundary"] = _point_on_diamond_boundary(first_start, source_box, 0.035)
                checks["source_escape_outside"] = not _point_inside_diamond(first_end, source_box, 0.005)
            if nodes[edge["to"]].get("type") == "decision":
                checks["target_on_boundary"] = _point_on_diamond_boundary(last_end, target_box, 0.035)
                checks["arrowhead_approach_outside"] = not _point_inside_diamond(last_start, target_box, 0.005)
            # Verify the serialized connector, not only the planned route.
            # Every diamond is an obstacle except for the first source segment
            # or final destination segment at its legitimate boundary port.
            reentries = []
            for diamond_id, diamond_shape in node_shapes.items():
                if nodes.get(diamond_id, {}).get("type") != "decision":
                    continue
                diamond_box = Box(diamond_shape.left / 914400.0,
                                  diamond_shape.top / 914400.0,
                                  diamond_shape.width / 914400.0,
                                  diamond_shape.height / 914400.0)
                for segment_index, segment_shape in enumerate(shapes):
                    allowed = ((diamond_id == edge["from"] and segment_index == 0) or
                               (diamond_id == edge["to"] and
                                segment_index == len(shapes) - 1))
                    if allowed:
                        continue
                    segment = _pptx_connector_endpoints(segment_shape)
                    if _segment_hits_diamond(segment, diamond_box, 0.002):
                        reentries.append({"diamond": diamond_id,
                                          "segment": segment_index})
            checks["no_diamond_reentry"] = not reentries
            if not checks:
                continue
            verified = all(checks.values())
            record = {"edge_index": edge_index, "source": edge["from"], "target": edge["to"],
                      "slide_number": slide_number, "checks": checks,
                      "diamond_reentries": reentries,
                      "verification_result": "verified" if verified else "failed"}
            records.append(record)
            if not verified:
                errors.append(record)
    return {"status": "verified" if not errors else "failed",
            "verified_edges": len(records) - len(errors), "failed_edges": len(errors),
            "records": records}


def _verify_pptx_cross_slide_continuations(pptx_path: str | Path,
                                           plan: SlideCompositionPlan,
                                           nodes: dict[str, dict[str, Any]]) -> dict[str, Any]:
    """Verify visible, connected, business-readable cross-slide references."""
    reopened = Presentation(str(pptx_path)); records = []; errors = []
    indexed_slides = []
    for slide_number, slide in enumerate(reopened.slides, start=1):
        shapes = {shape.name: shape for shape in slide.shapes}
        workflow_boxes = [
            Box(shape.left / 914400.0, shape.top / 914400.0,
                shape.width / 914400.0, shape.height / 914400.0)
            for shape in slide.shapes if shape.name.startswith("workflow-node::")
        ]
        protected_boxes = [
            Box(shape.left / 914400.0, shape.top / 914400.0,
                shape.width / 914400.0, shape.height / 914400.0)
            for shape in slide.shapes if shape.name.startswith("workflow-swimlane-label::")
        ]
        indexed_slides.append((slide_number, shapes, workflow_boxes, protected_boxes))
    for record in plan.edge_records:
        if record["representation"] != "cross_slide":
            continue
        edge_id = int(record["edge_id"]); source_slide = int(record["source_page"]) + 1
        target_slide = int(record["target_page"]) + 1
        _, source_shapes, source_node_boxes, source_protected_boxes = indexed_slides[source_slide - 1]
        outgoing = source_shapes.get(f"workflow-continuation::{edge_id}::outgoing")
        links = [shape for name, shape in source_shapes.items()
                 if name.startswith(f"workflow-continuation-link::{edge_id}::outgoing::")]
        destination_label = str(nodes[record["target"]].get("label", "")).rstrip("?")
        branch_label = str(record.get("label", "")).strip()
        text = outgoing.text.strip() if outgoing is not None and hasattr(outgoing, "text") else ""
        slide_reference_matches = bool(outgoing and f"slide {target_slide}" in text)
        branch_reference_matches = bool(not branch_label or branch_label in text)
        # Detailed references name the destination. Compact/minimal references
        # deliberately use only the branch meaning and matching slide number;
        # the connected source node supplies the remaining business context.
        text_matches = bool(slide_reference_matches and branch_reference_matches and
                            (destination_label in text or len(text) <= 80))
        no_node_overlap = False; no_protected_overlap = False
        if outgoing is not None:
            marker_box = Box(outgoing.left / 914400.0, outgoing.top / 914400.0,
                             outgoing.width / 914400.0, outgoing.height / 914400.0)
            no_node_overlap = not any(_boxes_overlap(marker_box, box, 0.01)
                                      for box in source_node_boxes)
            no_protected_overlap = not any(_boxes_overlap(marker_box, box, 0.01)
                                           for box in source_protected_boxes)
        _, target_shapes, _, target_protected_boxes = indexed_slides[target_slide - 1]
        incoming = target_shapes.get(f"workflow-continuation::{edge_id}::incoming")
        incoming_protected_overlap = False
        if incoming is not None:
            incoming_box = Box(incoming.left / 914400.0, incoming.top / 914400.0,
                               incoming.width / 914400.0, incoming.height / 914400.0)
            incoming_protected_overlap = any(
                _boxes_overlap(incoming_box, box, 0.01) for box in target_protected_boxes
            )
        represented = bool(outgoing and links and text_matches and no_node_overlap and
                           no_protected_overlap and not incoming_protected_overlap)
        item = {
            "edge_id": edge_id, "source_node": record["source"],
            "destination_node": record["target"], "source_slide": source_slide,
            "destination_slide": target_slide, "branch_label": branch_label,
            "outgoing_indicator_shape_ids": [outgoing.shape_id] if outgoing else [],
            "outgoing_link_shape_ids": [shape.shape_id for shape in links],
            "incoming_indicator_shape_ids": [incoming.shape_id] if incoming else [],
            "visually_represented": represented,
            "destination_reference_understandable": text_matches,
            "indicator_overlaps_workflow_node": not no_node_overlap,
            "indicator_overlaps_swimlane_label": (not no_protected_overlap or
                                                   incoming_protected_overlap),
        }
        records.append(item)
        if not represented:
            errors.append(item)
    return {"status": "verified" if not errors else "failed",
            "cross_slide_edges": len(records), "visually_represented": len(records) - len(errors),
            "unrepresented": len(errors), "records": records, "errors": errors}


def _outcome_category(node, workflow, style):
    metadata = workflow_outcome_metadata(workflow)
    if node["id"] in style.get("outcome_overrides", {}):
        return style["outcome_overrides"][node["id"]]
    if node["id"] in metadata:
        return metadata[node["id"]]
    if node.get("type") == "error":
        return "recoverable"
    if node.get("type") == "end":
        return "unclassified"
    return None


def _outcome_paint(category, theme, style):
    custom = style.get("colors", {}).get(category)
    if custom:
        text = "000000" if _relative_luminance(custom) > 0.56 else "FFFFFF"
        return custom, text, custom
    palette = {
        "green": (theme.success, "FFFFFF", theme.success),
        "red": (theme.exception, "FFFFFF", theme.exception),
        "blue": (theme.accent, "FFFFFF", theme.accent),
        "neutral": (theme.lane_fill, theme.text, theme.line),
    }
    return palette[style["outcome_colors"][category]]


def _relative_luminance(color: str) -> float:
    values = [int(color[index:index + 2], 16) / 255 for index in (0, 2, 4)]
    linear = [value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4
              for value in values]
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]


def _contrast_ratio(first: str, second: str) -> float:
    lighter, darker = sorted((_relative_luminance(first), _relative_luminance(second)), reverse=True)
    return (lighter + 0.05) / (darker + 0.05)


def _style_contrast_warnings(theme: Theme, style: dict[str, Any]) -> list[str]:
    colors = style.get("colors", {})
    pairs = {
        "slide title": (colors.get("title", theme.text), theme.background),
        "subtitle": (colors.get("subtitle", theme.muted), theme.background),
        "activity text": (colors.get("text", theme.text), colors.get("activity_fill", theme.accent_soft)),
        "decision text": (colors.get("text", theme.text), colors.get("decision_fill", theme.decision)),
    }
    return [f"Color contrast for {label} is {ratio:.2f}:1; 4.5:1 is recommended."
            for label, pair in pairs.items() if (ratio := _contrast_ratio(*pair)) < 4.5]


def _add_title(slide, title: str, subtitle: str, theme: Theme, style: dict[str, Any] | None = None):
    style = style or {}; colors = style.get("colors", {}); typography = style.get("typography", {})
    box = slide.shapes.add_textbox(Inches(0.45), Inches(0.12), Inches(12.4), Inches(0.56))
    _set_text(box, title, theme, typography.get("title_size", 21), True, colors.get("title", theme.text), style, "title")
    box.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    title_run = box.text_frame.paragraphs[0].runs[0]
    title_run.font.name = typography.get("font_family", "Arial")
    title_run._r.get_or_add_rPr().set("spc", "0")
    if subtitle.strip():
        subtitle_size = typography.get("subtitle_size", 7.5 if len(subtitle) > 125 else 8 if len(subtitle) > 95 else 9)
        sub = slide.shapes.add_textbox(Inches(0.47), Inches(0.62), Inches(12.2), Inches(0.37))
        _set_text(sub, subtitle, theme, subtitle_size, False, colors.get("subtitle", theme.muted), style, "subtitle")
        sub.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        subtitle_run = sub.text_frame.paragraphs[0].runs[0]
        subtitle_run.font.name = typography.get("font_family", "Arial")
        subtitle_run._r.get_or_add_rPr().set("spc", "0")


def _add_slide_background(slide, theme: Theme):
    background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SLIDE_W), Inches(SLIDE_H))
    background.name = "presentation-background"
    background.fill.solid(); background.fill.fore_color.rgb = _rgb(theme.background)
    background.line.fill.background()


def _concise_stage_label(label: str) -> str:
    """Reduce a compound process-area label to its primary business stage."""
    value = str(label).strip()
    for separator in (" and ", " & ", " / "):
        if separator in value:
            value = value.split(separator, 1)[0].strip()
            break
    return value


def _infer_stage_labels(nodes: list[dict[str, Any]]) -> list[str]:
    """Infer short, domain-neutral stage names from visible activity verbs."""
    categories = (
        ("Submission", ("submit", "request", "initiat", "intake", "prepare", "collect")),
        ("Review", ("review", "verify", "validat", "check", "assess", "identify")),
        ("Decision", ("decision", "approve", "reject", "eligible", "outcome", "authoriz")),
        ("Completion", ("pay", "reimburse", "issue", "record", "notify", "complete", "activat", "sign")),
    )
    found = []
    for node in nodes:
        label = str(node.get("label", "")).lower()
        for category, tokens in categories:
            if any(token in label for token in tokens) and category not in found:
                found.append(category)
                break
    return found


def _slide_content_title(workflow: dict[str, Any], node_ids: list[str], page: int,
                         total_pages: int, section_label: str | None = None) -> str:
    """Create a concise business title from the content actually shown."""
    node_by_id = {node["id"]: node for node in workflow["nodes"]}
    visible = [node_by_id[node_id] for node_id in node_ids if node_id in node_by_id]
    phase_defs = {phase["id"]: phase.get("label", phase["id"])
                  for phase in workflow.get("phases", [])}
    phases = list(dict.fromkeys(phase_defs.get(node.get("phase"), node.get("phase"))
                                for node in visible if node.get("phase")))
    area_defs = {area["id"]: area.get("label", area["id"])
                 for area in workflow.get("process_areas", [])}
    area_counts = Counter(node.get("process_area") for node in visible if node.get("process_area"))
    area_order = {area_id: index for index, area_id in enumerate(area_defs)}
    visible_area_ids = (sorted(area_counts,
                               key=lambda area_id: (-area_counts[area_id],
                                                    area_order.get(area_id, 10**6),
                                                    str(area_id)))
                        if len(workflow.get("nodes", [])) >= 45 else
                        [area_id for area_id in area_defs if area_id in area_counts])
    areas = [area_defs.get(area_id, str(area_id)) for area_id in visible_area_ids]
    # Process areas are normally more specific than broad phases on a
    # multi-slide diagram and avoid repeated titles such as "Verification &
    # Decision" on adjacent slides. A one-slide workflow keeps its phase name.
    if total_pages > 1 and areas:
        title = " & ".join(_concise_stage_label(value) for value in areas[:2])
    elif phases:
        title = " & ".join(_concise_stage_label(value) for value in phases[:2])
    elif areas:
        title = " & ".join(_concise_stage_label(value) for value in areas[:2])
    elif section_label and not section_label.startswith("Steps ") and section_label != "Complete workflow":
        title = section_label.replace(" / ", " and ").split(" · ")[0]
    else:
        inferred = _infer_stage_labels(visible)
        meaningful = [node for node in visible if node.get("type") not in {"start", "end", "error"}]
        if not meaningful:
            meaningful = visible
        labels = [str(node.get("label", "")).strip().rstrip("?") for node in meaningful]
        labels = [label for label in labels if label]
        if inferred:
            title = " & ".join(inferred[:2])
        elif len(labels) >= 2:
            title = f"{labels[0]} and {labels[-1]}"
        elif labels:
            title = labels[0]
        else:
            title = workflow["title"]
    if len(title) > 68:
        title = title[:65].rsplit(" ", 1)[0] + "…"
    return title


def _cross_slide_context(page: int, plan: SlideCompositionPlan,
                         nodes: dict[str, dict[str, Any]]) -> str:
    """Describe boundary relationships in business terms without graph IDs."""
    records = {record["edge_id"]: record for record in plan.edge_records}
    phrases = []
    for relationship in plan.slide_relationships:
        if page + 1 not in {relationship["from_slide"], relationship["to_slide"]}:
            continue
        for edge_id in relationship["edge_ids"]:
            record = records[edge_id]
            source = str(nodes[record["source"]].get("label", record["source"])).rstrip("?")
            target = str(nodes[record["target"]].get("label", record["target"])).rstrip("?")
            label = record["label"].strip()
            if record["path"] == "recovery":
                phrase = f"{source} returns to {target}"
            elif label:
                phrase = f"{source}: {label} → {target}"
            else:
                phrase = f"{source} → {target}"
            if phrase not in phrases:
                phrases.append(phrase)
    return " · Handoffs: " + "; ".join(phrases) if phrases else ""


def _shape_kind(node_type: str):
    return {"start": MSO_SHAPE.OVAL, "end": MSO_SHAPE.OVAL, "decision": MSO_SHAPE.DIAMOND,
            "subprocess": MSO_SHAPE.ROUNDED_RECTANGLE,
            "manual_action": MSO_SHAPE.ROUNDED_RECTANGLE,
            "external_system": MSO_SHAPE.ROUNDED_RECTANGLE, "error": MSO_SHAPE.OCTAGON}.get(
                node_type, MSO_SHAPE.ROUNDED_RECTANGLE)


def _styled_shape_kind(node_type: str, style: dict[str, Any]):
    shapes = style.get("shapes", {})
    if node_type in {"process", "subprocess", "manual_action", "external_system"}:
        return MSO_SHAPE.RECTANGLE if shapes.get("activity_corner_style") == "square" else _shape_kind(node_type)
    if node_type in {"start", "end", "error"} and "terminal_shape" in shapes:
        return {"oval": MSO_SHAPE.OVAL, "rounded": MSO_SHAPE.ROUNDED_RECTANGLE,
                "square": MSO_SHAPE.RECTANGLE}[shapes["terminal_shape"]]
    return _shape_kind(node_type)


def _node_size(node: dict[str, Any], rendering_mode: str = "compact") -> tuple[float, float]:
    # Compact and balanced diagrams show the action label; the full source
    # description remains attached to the editable shape as alternative text.
    # Detailed mode reserves visible space for both.
    length = len(node.get("label", ""))
    if rendering_mode == "detailed":
        length += len(node.get("description", ""))
    longest_word = max((len(word) for word in str(node.get("label", "")).split()), default=1)
    if node.get("type") in {"end", "error"}:
        words = str(node.get("label", "")).split()
        longest = max((len(word) for word in words), default=1)
        width = min(2.30, max(1.48, 0.74 + length * 0.024, 0.52 + longest * 0.095))
        lines = max(1, math.ceil(max(1, length) / max(12, int(width * 11))))
        height = min(1.02, max(0.66, 0.48 + lines * 0.15))
        return width, height
    if rendering_mode == "compact":
        if node.get("type") == "decision":
            return min(2.35, max(1.65, 1.48 + length * 0.015)), min(1.42, max(0.96, 0.86 + length * 0.007))
        if length > 55:
            return min(2.40, 1.20 + length * 0.012), min(1.15, 0.54 + length * 0.006)
        return min(1.70, max(0.84, 0.76 + length * 0.011,
                             0.34 + longest_word * 0.078)), min(0.68, max(0.56, 0.47 + length * 0.004))
    if rendering_mode == "balanced":
        if node.get("type") == "decision":
            return min(1.85, max(1.38, 1.18 + length * 0.020)), min(1.15, max(0.94, 0.84 + length * 0.009))
        return min(1.78, max(1.12, 1.00 + length * 0.012,
                             0.34 + longest_word * 0.078)), min(0.76, max(0.56, 0.52 + length * 0.004))
    if node.get("type") == "decision":
        return min(2.35, max(1.65, 1.35 + length * 0.024)), min(1.42, max(1.12, 0.98 + length * 0.010))
    return min(2.05, max(1.45, 1.30 + length * 0.012,
                         0.34 + longest_word * 0.078)), min(0.92, max(0.64, 0.58 + length * 0.005))


def render_presentation(workflow: dict[str, Any], layout_result: LayoutResult, output_path: str,
                        theme: Theme, layout: str, include_exceptions: str,
                        continuation_style: str = "minimal", rendering_mode: str = "compact",
                        include_assumptions: bool = False, include_open_questions: bool = False,
                        include_overview: bool = False, style: dict[str, Any] | None = None) -> dict[str, Any]:
    style = style or {}
    if include_exceptions not in {"separate", "inline", "omit", "exception_flowchart"}:
        raise ValueError("include_exceptions must be 'separate', 'inline', 'omit', or 'exception_flowchart'")
    if continuation_style not in {"minimal", "compact", "detailed"}:
        raise ValueError("continuation_style must be 'minimal', 'compact', or 'detailed'")
    prs = Presentation(); prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    blank = prs.slide_layouts[6]; nodes = {n["id"]: n for n in workflow["nodes"]}
    actors = workflow["actors"] if layout == "swimlane" else [{"id": "workflow", "label": "Workflow"}]
    edges = workflow["edges"]
    # Recovery edges are operational return paths in the main graph.  Hiding
    # them leaves correction activities disconnected and can make a decision
    # appear to have fewer outcomes than the source workflow.  The exception
    # setting controls exception-only paths and supporting slides.
    if include_exceptions == "inline":
        normal_edges = list(edges)
    else:
        normal_edges = [edge for edge in edges if edge.get("path") != "exception"]
        # A retained terminal must never be left floating merely because its
        # final incoming edge is tagged as an exception. If both endpoint nodes
        # remain in the presentation, retain that final visual connection.
        retained_ids = set(nodes)
        terminal_ids_for_edges = {node_id for node_id, node in nodes.items()
                                  if node.get("type") in {"end", "error"}}
        for edge in edges:
            if (edge.get("path") == "exception" and edge["from"] in retained_ids and
                    edge["to"] in terminal_ids_for_edges and edge not in normal_edges):
                normal_edges.append(edge)
    continuation_count = 0
    quality = {key: 0 for key in ("node_overlaps", "connector_node_intersections", "connector_diamond_intersections", "connector_text_intersections",
                                   "diamond_attachment_errors", "diamond_interior_attachment_segments",
                                   "arrowhead_diamond_interior", "connector_shape_reentry_intersections",
                                   "connector_segments_off_slide", "arrowhead_shape_overlaps",
                                   "arrowheads_off_slide", "label_collisions", "label_connector_overlaps",
                                   "branch_label_association_errors", "unnecessary_connector_detours",
                                   "ambiguous_destination_overlaps",
                                   "connector_crossings", "parallel_segment_overlaps",
                                   "text_overflow_warnings", "terminal_text_overflow", "off_slide_elements", "excessive_empty_space",
                                   "very_long_connector_paths", "unreadably_small_shapes")}
    layout_warnings = _style_contrast_warnings(theme, style)
    classified = set(workflow_outcome_metadata(workflow)) | set(style.get("outcome_overrides", {}))
    unclassified_terminals = sorted(node["id"] for node in workflow["nodes"]
                                    if node.get("type") == "end" and node["id"] not in classified)
    if unclassified_terminals:
        layout_warnings.append("Terminal outcomes rendered neutral because no semantic classification was supplied: " +
                               ", ".join(unclassified_terminals) + ".")
    process_map_unused_space = None; process_map_pagination = None
    rendered_node_ids: list[str] = []; covered_edge_indices: set[int] = set()
    exception_split_groups: list[str] = []
    indexed_edges = [dict(edge, _render_index=index) for index, edge in enumerate(normal_edges)]
    if layout == "process_map":
        map_workflow = dict(workflow); map_workflow["edges"] = indexed_edges
        process_map_pagination = paginate_process_map(map_workflow, rendering_mode)
        map_pages = process_map_pagination.pages
        map_page_of = {node_id: page for page, values in enumerate(map_pages) for node_id in values}
        main_pages = len(map_pages)
        layout_warnings.extend(process_map_pagination.warnings)
        if include_overview:
            slide = prs.slides.add_slide(blank)
            _add_slide_background(slide, theme)
            overview_mappings = _render_process_map_overview(slide, workflow, theme, style)
        else:
            overview_mappings = {}
    else:
        main_pages = layout_result.pages; overview_mappings = {}
        map_page_of = {node_id: placement.page for node_id, placement in layout_result.placements.items()}
    if layout == "process_map":
        ordered_pages = [list(values) for values in map_pages]
    else:
        ordered_pages = []
        for page in range(main_pages):
            values = [node_id for node_id, placement in layout_result.placements.items()
                      if placement.page == page]
            values.sort(key=lambda node_id: (layout_result.placements[node_id].column,
                                              layout_result.placements[node_id].lane, node_id))
            ordered_pages.append(values)
    planned_node_ids = set(map_page_of)
    composition_workflow = dict(workflow)
    composition_workflow["nodes"] = [node for node in workflow["nodes"] if node["id"] in planned_node_ids]
    composition_workflow["edges"] = [edge for edge in normal_edges
                                     if edge["from"] in planned_node_ids and edge["to"] in planned_node_ids]
    composition_plan = build_slide_composition_plan(
        composition_workflow, map_page_of, ordered_pages,
        "process_map" if layout == "process_map" else f"{layout}_{rendering_mode}",
    )
    semantic_validation = validate_composition_plan(composition_workflow, composition_plan)
    if semantic_validation["critical"]:
        raise ValueError("Critical semantic validation failed; no PowerPoint was written: " +
                         " ".join(semantic_validation["critical"]))
    preferred_slides = style.get("layout_preferences", {}).get("preferred_slide_count")
    if preferred_slides and preferred_slides != main_pages:
        layout_warnings.append(
            f"Preferred slide count was {preferred_slides}; {main_pages} diagram slides were required "
            "to preserve readability and semantic integrity."
        )
    incoming_edges: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for edge in normal_edges:
        if edge["from"] in map_page_of and edge["to"] in map_page_of:
            incoming_edges[edge["to"]].append(edge)
    terminal_ids = [node_id for node_id, node in nodes.items()
                    if node.get("type") in {"end", "error"} and node_id in map_page_of]
    orphaned_terminals = []
    terminal_predecessors = {}
    cross_slide_terminal_predecessors = {}
    for node_id in terminal_ids:
        predecessors = [edge["from"] for edge in incoming_edges.get(node_id, [])]
        terminal_predecessors[node_id] = predecessors
        cross_slide_terminal_predecessors[node_id] = [
            source for source in predecessors if map_page_of[source] != map_page_of[node_id]
        ]
        # A terminal with only cross-slide predecessors is not visually
        # orphaned when its source edges receive explicit, verified
        # continuation references. The post-write continuation verifier is the
        # hard check for those relationships.
        if not predecessors:
            orphaned_terminals.append(node_id)
    if orphaned_terminals:
        raise ValueError("Critical terminal connectivity failure; terminal outcomes lack a visible incoming "
                         "connection on their slide: " + ", ".join(sorted(orphaned_terminals)) + ".")
    terminal_connectivity = {
        "terminals_checked": len(terminal_ids),
        "visually_orphaned_terminals": [],
        "same_slide_predecessors": {
            node_id: [source for source in predecessors
                      if map_page_of[source] == map_page_of[node_id]]
            for node_id, predecessors in terminal_predecessors.items()
        },
        "cross_slide_predecessors": cross_slide_terminal_predecessors,
        "cross_slide_edges_accounted": sum(
            map_page_of[edge["from"]] != map_page_of[edge["to"]]
            for edge in normal_edges if edge["from"] in map_page_of and edge["to"] in map_page_of
        ),
    }
    unused_values = []; slide_titles: list[str] = []
    # Cross-slide handoffs are operational diagnostics, not slide furniture.  Keep
    # the business relationships available to callers without forcing a long,
    # technical transition inventory into the visible subtitle.
    cross_slide_contexts = [
        _cross_slide_context(page, composition_plan, nodes).removeprefix(" · ")
        for page in range(main_pages)
    ]
    for page, plan_page in enumerate(composition_plan.pages):
        section_label = (process_map_pagination.page_labels[page] if process_map_pagination else None)
        slide_titles.append(workflow["title"] if main_pages == 1 else
                            _slide_content_title(workflow, list(plan_page.node_ids), page,
                                                 main_pages, section_label))
    duplicate_titles = Counter(slide_titles)
    if any(count > 1 for count in duplicate_titles.values()):
        node_by_id = {node["id"]: node for node in workflow["nodes"]}
        used_titles = set()
        for page, plan_page in enumerate(composition_plan.pages):
            title = slide_titles[page]
            if duplicate_titles[title] <= 1:
                used_titles.add(title)
                continue
            visible = [node_by_id[node_id] for node_id in plan_page.node_ids
                       if node_id in node_by_id]
            visible_text = " ".join(str(node.get("label", "")).lower()
                                    for node in visible)
            manual_count = sum(node.get("type") == "manual_action" for node in visible)
            automated_count = sum(node.get("type") == "external_system" for node in visible)
            recovery_count = sum(node.get("type") in {"error", "end"} for node in visible)
            if any(token in visible_text for token in ("eligib", "scope")):
                qualifier = "Eligibility and Application Setup"
            elif "terms" in visible_text or "combined status" in visible_text:
                qualifier = "Verification Completion and Terms"
            elif manual_count >= 2:
                qualifier = "Manual Review and Recovery"
            elif automated_count >= 2:
                qualifier = "Automated Checks"
            elif recovery_count >= 2:
                qualifier = "Outcomes and Resolution"
            else:
                qualifier = "Decision Paths"
            candidate = f"{title}: {qualifier}"
            if candidate in used_titles:
                candidate = f"{title}: {qualifier} {page + 1}"
            slide_titles[page] = candidate
            used_titles.add(candidate)
    for page in range(main_pages):
        slide = prs.slides.add_slide(blank)
        _add_slide_background(slide, theme)
        section = (f" · {process_map_pagination.page_labels[page]}" if process_map_pagination else "")
        slide_title = slide_titles[page]
        # Domain, version, edge inventories, and cross-slide handoffs belong in
        # traceability, not in ordinary business-slide headers. A future
        # explicit subtitle option can opt content back in without making
        # technical metadata the default.
        subtitle = ""
        _add_title(slide, slide_title, subtitle, theme, style)
        if layout == "process_map":
            page_workflow = process_map_page_view(map_workflow, set(map_pages[page]))
            page_result = _render_process_map_page(slide, page_workflow, indexed_edges, theme, rendering_mode,
                                                   page, map_page_of, main_pages, style)
            layout_warnings.extend(page_result.get("warnings", []))
            unused_values.append(page_result.get("unused_space", 0.0))
            rendered_node_ids.extend(page_result.get("rendered_node_ids", []))
            covered_edge_indices.update(page_result.get("covered_edge_indices", []))
        else:
            page_result = _render_page(slide, nodes, actors, layout_result, page, theme, layout,
                                       indexed_edges, continuation_style=continuation_style,
                                       rendering_mode=rendering_mode, style=style, workflow=workflow)
            layout_warnings.extend(page_result.get("warnings", []))
        continuation_count += page_result["continuation_count"]
        _merge_quality(quality, page_result["quality"])
    if layout != "process_map":
        rendered_node_ids = list(layout_result.placements)
        # Every source edge is either drawn locally or retained as an explicit
        # cross-slide relationship in the layout model.
        covered_edge_indices = set(range(len(indexed_edges)))
    if unused_values:
        process_map_unused_space = round(sum(unused_values) / len(unused_values), 3)

    if layout == "process_map":
        expected_nodes = {node["id"] for node in workflow["nodes"]}
        rendered_counts = {node_id: rendered_node_ids.count(node_id) for node_id in expected_nodes}
        missing_nodes = sorted(node_id for node_id, count in rendered_counts.items() if count == 0)
        duplicate_nodes = sorted(node_id for node_id, count in rendered_counts.items() if count > 1)
        missing_edges = sorted(set(range(len(indexed_edges))) - covered_edge_indices)
        if missing_nodes or duplicate_nodes or missing_edges:
            raise ValueError("Process-map integrity failure: "
                             f"missing nodes={missing_nodes}, duplicate nodes={duplicate_nodes}, "
                             f"missing edge indices={missing_edges}.")

    special_edges = [e for e in edges if e.get("path") in {"exception", "recovery"}]
    if include_exceptions == "separate" and special_edges:
        slide = prs.slides.add_slide(blank); _add_slide_background(slide, theme)
        _add_title(slide, "Exceptions and recovery", "Illustrative exception register", theme, style)
        _render_exception_summary(slide, special_edges, nodes, theme)
    elif include_exceptions == "exception_flowchart" and special_edges:
        exception_workflow = _exception_view(workflow)
        exception_layout = calculate_layout(exception_workflow, layout, rendering_mode=rendering_mode)
        exception_split_groups = exception_layout.split_decision_groups
        exception_pages = 1 if layout == "process_map" else exception_layout.pages
        for page in range(exception_pages):
            slide = prs.slides.add_slide(blank)
            _add_slide_background(slide, theme)
            suffix = f" · Part {page + 1} of {exception_layout.pages}" if exception_layout.pages > 1 else ""
            _add_title(slide, "Exception and recovery flow" + suffix,
                       "Errors, recovery actions, retries, escalation, and outcomes", theme, style)
            if layout == "process_map":
                page_result = _render_process_map_page(slide, exception_workflow,
                                                       exception_workflow["edges"], theme, rendering_mode,
                                                       style=style)
                layout_warnings.extend(page_result.get("warnings", []))
            else:
                page_result = _render_page(slide, nodes, actors, exception_layout, page, theme, layout,
                                           exception_workflow["edges"], True, continuation_style, rendering_mode,
                                           style, exception_workflow)
            continuation_count += page_result["continuation_count"]
            _merge_quality(quality, page_result["quality"])

    _add_notes_slides(prs, workflow, blank, theme, include_assumptions, include_open_questions)
    text_audit = _presentation_text_audit(prs)
    quality["text_overflow_warnings"] = max(
        quality.get("text_overflow_warnings", 0), text_audit["critical_text_clipping"])
    critical_keys = (
        "node_overlaps", "connector_node_intersections", "connector_diamond_intersections", "connector_text_intersections",
        "diamond_attachment_errors", "diamond_interior_attachment_segments", "arrowhead_diamond_interior",
        "connector_shape_reentry_intersections",
        "connector_segments_off_slide", "arrowhead_shape_overlaps", "arrowheads_off_slide",
        "label_collisions", "label_connector_overlaps", "ambiguous_destination_overlaps",
        "text_overflow_warnings", "terminal_text_overflow",
        "off_slide_elements",
    )
    critical_geometry = {key: quality[key] for key in critical_keys if quality.get(key, 0)}
    if critical_geometry:
        summary = ", ".join(f"{key}={value}" for key, value in sorted(critical_geometry.items()))
        if quality.get("node_overlap_details"):
            summary += "; node_overlap_details=" + json.dumps(quality["node_overlap_details"])
        if quality.get("label_collision_details"):
            summary += "; label_collision_details=" + json.dumps(
                quality["label_collision_details"][:12])
        if quality.get("ambiguous_destination_overlap_details"):
            summary += "; ambiguous_destination_overlap_details=" + json.dumps(
                quality["ambiguous_destination_overlap_details"][:12])
        if quality.get("connector_shape_reentry_details"):
            summary += "; connector_shape_reentry_details=" + json.dumps(
                quality["connector_shape_reentry_details"][:12])
        if quality.get("connector_node_intersection_details"):
            summary += "; connector_node_intersection_details=" + json.dumps(
                quality["connector_node_intersection_details"][:12])
        raise ValueError("Critical geometry validation failed; no PowerPoint was written: " + summary + ".")
    geometry_status = "clean"
    if quality.get("branch_label_association_errors"):
        geometry_status = "clean_with_warnings"
        layout_warnings.append(
            f"Branch-label proximity warnings: {quality['branch_label_association_errors']}; "
            "labels remain clear of connector strokes and slide boundaries."
        )
    if layout_warnings:
        geometry_status = "clean_with_warnings"
    prs.save(output_path)
    pptx_diamond_verification = _verify_pptx_diamond_attachments(
        output_path, nodes, indexed_edges, main_pages)
    if pptx_diamond_verification["status"] != "verified":
        Path(output_path).unlink(missing_ok=True)
        raise ValueError(
            "PowerPoint diamond attachment verification failed after writing the artifact: "
            f"failed_edges={pptx_diamond_verification['failed_edges']}."
        )
    cross_slide_verification = _verify_pptx_cross_slide_continuations(
        output_path, composition_plan, nodes)
    if cross_slide_verification["status"] != "verified":
        Path(output_path).unlink(missing_ok=True)
        raise ValueError(
            "PowerPoint cross-slide semantic verification failed after writing the artifact: "
            f"unrepresented_edges={cross_slide_verification['unrepresented']}."
        )
    element_style_verification = _verify_element_styles(output_path, style)
    failed_element_styles = [item for item in element_style_verification
                             if item["verification_result"] != "verified"]
    if failed_element_styles:
        Path(output_path).unlink(missing_ok=True)
        failed_ids = ", ".join(item["resolved_source_node_id"] for item in failed_element_styles)
        raise ValueError("Element styling verification failed after writing the PowerPoint for: " + failed_ids + ".")
    connector_names = {}
    for slide_index, slide in enumerate(list(prs.slides)[:main_pages], start=1):
        for shape in slide.shapes:
            if shape.name.startswith("workflow-edge::"):
                edge_token = shape.name.split("::", 2)[1]
                connector_names.setdefault(int(edge_token), []).append(shape.name)
    traceability = {
        "builder_version": BUILDER_VERSION,
        "workflow_title": workflow.get("title", ""),
        "layout_strategy": composition_plan.strategy,
        "nodes": [{
            "source_node_id": node_id, "source_label": nodes[node_id].get("label", ""),
            "slide_number": map_page_of[node_id] + 1,
            "rendered_shape_identifier": f"workflow-node::{node_id}",
            "node_type": nodes[node_id].get("type"), "owner": nodes[node_id].get("owner"),
        } for node_id in sorted(map_page_of)],
        "edges": [{
            "source_node_id": record["source"], "destination_node_id": record["target"],
            "original_branch_label": record["label"],
            "source_edge_identity": f"edge-{record['edge_id']:03d}",
            "representation": record["representation"],
            "slide_number": record["source_page"] + 1 if record["representation"] == "local" else None,
            "slide_transition": ([record["source_page"] + 1, record["target_page"] + 1]
                                 if record["representation"] == "cross_slide" else None),
            "rendered_connector_identifiers": connector_names.get(record["edge_id"], []),
            "continuation_indicator_shape_ids": next(
                (item["outgoing_indicator_shape_ids"] + item["incoming_indicator_shape_ids"]
                 for item in cross_slide_verification["records"]
                 if item["edge_id"] == record["edge_id"]), []),
            "visually_represented": (record["representation"] == "local" or next(
                (item["visually_represented"] for item in cross_slide_verification["records"]
                 if item["edge_id"] == record["edge_id"]), False)),
            "destination_reference_understandable": (record["representation"] == "local" or next(
                (item["destination_reference_understandable"]
                 for item in cross_slide_verification["records"]
                 if item["edge_id"] == record["edge_id"]), False)),
            "path": record["path"],
        } for record in composition_plan.edge_records],
        "decisions": semantic_validation["checks"]["decision_branch_mappings"],
        "terminals": {node_id: [{"source": edge["from"], "label": str(edge.get("label", ""))}
                                for edge in normal_edges if edge["to"] == node_id]
                      for node_id in terminal_ids},
        "slide_relationships": list(composition_plan.slide_relationships),
        "cross_slide_contexts": [
            {"slide_number": index + 1, "context": context}
            for index, context in enumerate(cross_slide_contexts) if context
        ],
        "semantic_validation": semantic_validation,
    }
    traceability_path = str(Path(output_path).with_suffix(".traceability.json"))
    Path(traceability_path).write_text(json.dumps(traceability, indent=2, ensure_ascii=False) + "\n",
                                       encoding="utf-8")
    quality["slides_generated"] = len(prs.slides)
    page_node_counts = []
    page_terminal_only = []
    for page in range(main_pages):
        if layout == "process_map":
            page_ids = list(map_pages[page])
        else:
            page_ids = [node_id for node_id, placement in layout_result.placements.items()
                        if placement.page == page]
        page_node_counts.append(len(page_ids))
        page_terminal_only.append(bool(page_ids) and all(nodes[node_id].get("type") in {"end", "error"}
                                                         for node_id in page_ids))
    cross_slide_edges = (process_map_pagination.split_edges if process_map_pagination else
                         layout_result.split_edges)
    return {"slides": len(prs.slides), "diagram_slides": main_pages,
            "exception_slide": include_exceptions in {"separate", "exception_flowchart"} and bool(special_edges),
            "exception_mode": include_exceptions, "continuation_style": continuation_style,
            "rendering_mode": rendering_mode,
            "diagram_structure": layout,
            "slide_titles": slide_titles,
            "continuation_markers": continuation_count,
            "cross_slide_contexts": [
                {"slide_number": index + 1, "context": context}
                for index, context in enumerate(cross_slide_contexts) if context
            ],
            "exception_split_decision_groups": exception_split_groups,
            "layout_quality": quality,
            "geometry_status": geometry_status,
            "critical_geometry_defects": {},
            "text_layout_validation": text_audit,
            "pptx_diamond_attachment_verification": pptx_diamond_verification,
            "cross_slide_continuity": cross_slide_verification,
            "element_style_verification": element_style_verification,
            "layout_warnings": sorted(set(layout_warnings)),
            "process_map_unused_space": process_map_unused_space,
            "process_map_complexity": (process_map_pagination.complexity.__dict__
                                       if process_map_pagination else None),
            "process_map_split_edges": process_map_pagination.split_edges if process_map_pagination else 0,
            "pagination_quality": {
                "page_node_counts": page_node_counts,
                "sparse_slides": sum(count <= 2 for count in page_node_counts),
                "terminal_only_slides": sum(page_terminal_only),
                "cross_slide_relationships": cross_slide_edges,
                "visible_continuation_labels": continuation_count,
                "candidate_scores": list(layout_result.candidate_scores) if layout != "process_map" else [],
            },
            "terminal_connectivity": terminal_connectivity,
            "composition_plan": {
                "strategy": composition_plan.strategy,
                "graph_analysis": layout_result.graph_analysis if layout != "process_map" else None,
                "pages": [{"slide_number": page.slide_number, "node_ids": list(page.node_ids),
                           "local_edge_ids": list(page.local_edge_ids),
                           "incoming_cross_edge_ids": list(page.incoming_cross_edge_ids),
                           "outgoing_cross_edge_ids": list(page.outgoing_cross_edge_ids),
                           "decision_ids": list(page.decision_ids),
                           "recovery_edge_ids": list(page.recovery_edge_ids),
                           "terminal_ids": list(page.terminal_ids),
                           "reading_order": list(page.reading_order)}
                          for page in composition_plan.pages],
                "slide_relationships": list(composition_plan.slide_relationships),
            },
            "semantic_validation": semantic_validation,
            "traceability_path": str(Path(traceability_path).resolve()),
            "traceability_statistics": {"nodes": len(traceability["nodes"]),
                                         "edges": len(traceability["edges"]),
                                         "decisions": len(traceability["decisions"]),
                                         "terminals": len(traceability["terminals"])},
            "node_coverage": {"expected": len(workflow["nodes"]), "rendered": len(set(rendered_node_ids)),
                              "missing": missing_nodes if layout == "process_map" else [],
                              "duplicates": duplicate_nodes if layout == "process_map" else []},
            "edge_coverage": {"expected": len(indexed_edges), "represented": len(covered_edge_indices),
                              "missing_indices": missing_edges if layout == "process_map" else []},
            "overview_included": bool(include_overview and layout == "process_map"),
            "overview_node_mappings": overview_mappings,
            "applied_theme": theme.name.replace("minimal_process_map", "minimal"),
            "applied_style": style,
            "native_shape_count": sum(len(s.shapes) for s in prs.slides)}


def _render_process_map_page(slide, workflow, edges, theme, rendering_mode,
                             page_index=0, page_of=None, total_pages=1, style=None):
    """Render geometry produced by the independent two-dimensional process-map planner."""
    style = style or {}
    plan = plan_process_map(workflow, rendering_mode, style.get("spacing", "standard"))
    nodes = {node["id"]: node for node in workflow["nodes"]}
    for index, area in enumerate(plan.areas):
        background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(area.y),
                                            Inches(12.70), Inches(area.h))
        colors = style.get("colors", {})
        lane_color = colors.get("swimlane_alt_fill", theme.background) if index % 2 else colors.get("swimlane_fill", theme.lane_fill)
        background.fill.solid(); background.fill.fore_color.rgb = _rgb(lane_color)
        background.line.color.rgb = _rgb(colors.get("activity_border", "D8DCE2")); background.line.width = Pt(0.6)
        label = slide.shapes.add_textbox(Inches(0.31), Inches(area.y + 0.05), Inches(0.52), Inches(area.h - 0.10))
        _set_text(label, area.label, theme, 8, True, colors.get("swimlane_header", theme.muted), style, "swimlane")
        body = label.text_frame._txBody.bodyPr
        body.set("rot", "-5400000")

    geometry = {node_id: Box(value.x, value.y, value.w, value.h) for node_id, value in plan.nodes.items()}
    for node_id in sorted(geometry, key=lambda value: (plan.nodes[value].area_id, plan.nodes[value].row,
                                                       plan.nodes[value].column, value)):
        node = nodes[node_id]; box = geometry[node_id]
        shape = slide.shapes.add_shape(_styled_shape_kind(node["type"], style), Inches(box.x), Inches(box.y),
                                       Inches(box.w), Inches(box.h))
        shape.name = f"workflow-node::{node_id}"
        terminal = node["type"] in {"start", "end"}; is_error = node["type"] == "error"
        category = _outcome_category(node, workflow, style)
        if category:
            fill, text_color, outcome_line = _outcome_paint(category, theme, style)
        elif terminal:
            fill, text_color = theme.accent, "FFFFFF"
        else:
            fill = theme.exception_soft if is_error else theme.decision if node["type"] == "decision" else "FFFFFF"
            text_color = theme.text
        default_border = outcome_line if category else theme.exception if is_error else theme.line
        label = node["label"]
        size = _decision_text_style(shape, style) if node["type"] == "decision" else 8
        fill, border, text_color, border_width, size = _node_visual(
            node, theme, style, fill, default_border, text_color, 1.0, size)
        shape.fill.solid(); shape.fill.fore_color.rgb = _rgb(fill)
        shape.line.color.rgb = _rgb(border); shape.line.width = Pt(border_width)
        _apply_shadow(shape, style.get("shapes", {}).get("shadows", False))
        explicit_size = "text_size" in style.get("elements", {}).get(node_id, {})
        role = ("decision" if node["type"] == "decision" else "terminal"
                if node["type"] in {"start", "end", "error"} else "activity")
        _set_text(shape, label, theme, size, terminal or bool(category), text_color, style, role, explicit_size)
        if node["type"] == "decision":
            _decision_text_style(shape, style)
        if node.get("description"):
            shape._element.nvSpPr.cNvPr.set("descr", node["description"])

    local_edges = [edge for edge in edges if edge["from"] in geometry and edge["to"] in geometry]
    outgoing = defaultdict(list)
    for edge in edges:
        if edge["from"] in geometry:
            outgoing[edge["from"]].append(edge)
    for source in outgoing:
        outgoing[source].sort(key=lambda edge: (geometry[edge["to"]].cy if edge["to"] in geometry else 99.0,
                                                edge.get("label", ""), edge["to"]))
    connector_segments = []; pending_labels = []
    for edge in local_edges:
        pending = _route_process_map_edge(
            slide, edge, geometry[edge["from"]], geometry[edge["to"]], geometry, theme,
            outgoing[edge["from"]].index(edge), len(outgoing[edge["from"]]),
            nodes[edge["from"]].get("type") == "decision",
            nodes[edge["to"]].get("type") == "decision", connector_segments, style,
        )
        if pending:
            pending_labels.append(pending)
    label_boxes = []; label_records = []; placement_warnings = []
    for text, points, source_is_decision, branch_index, branch_count, color, source, target in pending_labels:
        box = _place_edge_label(slide, text, points, source_is_decision, branch_index, branch_count,
                                list(geometry.values()), label_boxes, connector_segments, theme, color,
                                style["branch_label_position"], style["connector_label_offset"], source, target,
                                placement_warnings, measured=False, style=style)
        label_records.append((box, source, target, points))
    continuation_count, continuation_edges, continuation_records = _render_process_map_continuations(
        slide, page_index, total_pages, edges, page_of or {}, geometry, label_boxes,
        connector_segments, theme, nodes,
    )
    label_records.extend(continuation_records)
    quality_nodes = {node_id: dict(node, description="") for node_id, node in nodes.items()}
    quality = _quality_report(geometry, connector_segments, label_records, quality_nodes, list(geometry))
    quality["excessive_empty_space"] = int(plan.unused_space > 0.94)
    quality["unreadably_small_shapes"] += plan.small_shapes
    covered = {edge["_render_index"] for edge in local_edges if "_render_index" in edge} | continuation_edges
    return {"continuation_count": continuation_count, "quality": quality,
            "warnings": plan.warnings + placement_warnings,
            "unused_space": plan.unused_space, "rendered_node_ids": list(geometry),
            "covered_edge_indices": sorted(covered)}


def _render_process_map_continuations(slide, page, total_pages, edges, page_of, geometry,
                                      label_boxes, connector_segments, theme, nodes):
    if total_pages <= 1 or not page_of:
        return 0, set(), []
    covered = set(); count = 0; records = []
    cross_edges = [edge for edge in edges if edge["from"] in page_of and edge["to"] in page_of and
                   page_of[edge["from"]] != page_of[edge["to"]] and
                   (page_of[edge["from"]] == page or page_of[edge["to"]] == page)]
    cross_edges.sort(key=lambda edge: edge.get("_render_index", 0))
    # Cross-slide relationships stay traceable through their source edge
    # indices and stage titles.  Ordinary process maps do not expose internal
    # continuation prose or graph references on the canvas.
    if cross_edges:
        return 0, {edge.get("_render_index", 0) for edge in cross_edges}, []
    for edge in cross_edges:
        outgoing = page_of[edge["from"]] == page
        node_id = edge["from"] if outgoing else edge["to"]
        anchor = geometry.get(node_id)
        if anchor is None:
            continue
        index = edge.get("_render_index", 0)
        other = page_of[edge["to"]] + 1 if outgoing else page_of[edge["from"]] + 1
        direction = "To" if outgoing else "From"
        semantic = str(edge.get("label", "")).strip()
        source_label = str(nodes.get(edge["from"], {}).get("label", "previous process step")).rstrip("?")
        target_label = str(nodes.get(edge["to"], {}).get("label", "next process step")).rstrip("?")
        text = (f"{direction} {target_label if outgoing else source_label} · slide {other}" +
                (f" · {semantic}" if semantic else ""))
        width = min(2.35, max(1.05, 0.24 + len(text) * 0.055)); height = 0.22
        candidates = _continuation_candidates(anchor, width, height, outgoing, count)
        # Dense phase boundaries can create several continuations around one
        # node. Reserve deterministic top and bottom channels before accepting
        # any colliding fallback.
        for row_y in (0.92, 1.17, 6.69, 6.94):
            for slot in range(5):
                candidates.append(Box(0.92 + slot * 2.42, row_y, width, height))
        candidates = [_clamp_continuation_box(box) for box in candidates]
        blockers = list(geometry.values()) + label_boxes
        chosen = next((box for box in candidates if _inside_slide(box) and
                       not any(_boxes_overlap(box, other_box, 0.025) for other_box in blockers) and
                       not any(_segment_hits_box(segment, box, 0.02) for segment in connector_segments)), None)
        if chosen is None:
            valid = [box for box in candidates if _inside_slide(box)]
            if not valid:
                raise ValueError(f"No readable continuation-label position for edge {index} on Part {page + 1}.")
            chosen = min(valid, key=lambda box: sum(_boxes_overlap(box, other_box, 0.0)
                                                    for other_box in blockers))
        tag = slide.shapes.add_textbox(Inches(chosen.x), Inches(chosen.y), Inches(chosen.w), Inches(chosen.h))
        tag.fill.background(); tag.line.fill.background(); _set_text(tag, text, theme, 7, False, theme.muted)
        tag.text_frame.margin_left = tag.text_frame.margin_right = Inches(0.01)
        tag.text_frame.margin_top = tag.text_frame.margin_bottom = 0
        label_boxes.append(chosen); covered.add(index); count += 1
        records.append((chosen, edge["from"], edge["to"]))
    return count, covered, records


def _render_process_map_overview(slide, workflow, theme, style=None):
    """Render a phase-only overview whose boxes map to detailed workflow nodes."""
    phases = workflow.get("phases") or []
    if not phases:
        return {}
    style = style or {}
    _add_title(slide, workflow["title"] + " · Overview",
               "Phase overview · each phase maps to detailed workflow nodes", theme, style)
    mappings = {}; count = len(phases); gap = 0.22
    width = min(2.15, (11.8 - gap * max(0, count - 1)) / max(1, count))
    total = width * count + gap * max(0, count - 1); x = (SLIDE_W - total) / 2
    y = 2.55
    for index, phase in enumerate(phases):
        node_ids = [node["id"] for node in workflow["nodes"] if node.get("phase") == phase["id"]]
        mappings[phase["id"]] = node_ids
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                                     Inches(width), Inches(1.05))
        box.fill.solid(); box.fill.fore_color.rgb = _rgb("FFFFFF")
        box.line.color.rgb = _rgb(theme.line); box.line.width = Pt(1.0)
        _set_text(box, f"{phase.get('label', phase['id'])}\n{len(node_ids)} detailed steps", theme, 10, True)
        if index < count - 1:
            connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x + width), Inches(y + 0.52),
                                                   Inches(x + width + gap), Inches(y + 0.52))
            connector.line.color.rgb = _rgb(theme.line); _set_arrowhead(connector)
        x += width + gap
    return mappings


def _route_process_map_edge(slide, edge, a: Box, b: Box, geometry, theme, branch_index,
                            branch_count, source_is_decision, target_is_decision,
                            connector_segments, style=None):
    """Route a process-map edge on an obstacle-aware orthogonal grid."""
    try:
        points = _process_map_route_points(a, b, geometry, branch_index, branch_count,
                                           source_is_decision, connector_segments,
                                           target_is_decision)
    except ValueError as error:
        raise ValueError(f"Unable to route edge {edge['from']} -> {edge['to']}: {error}") from error
    is_special = edge.get("path") in {"exception", "recovery"}
    style = style or {}
    color, width, line_style, arrowhead = _edge_visual(edge, theme, style, is_special,
                                                       1.6 if is_special else 1.05)
    for index, (p1, p2) in enumerate(zip(points, points[1:])):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(p1[0]), Inches(p1[1]),
                                          Inches(p2[0]), Inches(p2[1]))
        line.name = f"workflow-edge::{edge.get('_render_index', 0)}::segment-{index + 1}"
        line.line.color.rgb = _rgb(color); line.line.width = Pt(width); _set_dash(line, line_style)
        if index == len(points) - 2 and arrowhead != "none":
            _set_arrowhead(line)
        connector_segments.append((p1, p2, edge["from"], edge["to"]))
    if edge.get("label"):
        return (str(edge["label"]), points, source_is_decision, branch_index, branch_count,
                color if is_special else theme.muted, edge["from"], edge["to"])
    return None


def _process_map_route_points(a: Box, b: Box, geometry, branch_index: int, branch_count: int,
                              source_is_decision: bool, existing_segments,
                              target_is_decision: bool = False):
    """Find the least-colliding editable Manhattan route between two process-map nodes."""
    clearance = 0.10
    if source_is_decision:
        source_port, source_escape = _decision_exit(a, branch_index, branch_count, b)
        source_escape = (min(13.02, max(0.90, source_escape[0])),
                         min(7.18, max(0.90, source_escape[1])))
        source_candidates = [(source_port, source_escape)]
        dx, dy = source_escape[0] - source_port[0], source_escape[1] - source_port[1]
        distance = max(abs(dx), abs(dy), 0.001)
        short_escape = (source_port[0] + dx / distance * clearance,
                        source_port[1] + dy / distance * clearance)
        if short_escape != source_escape:
            source_candidates.append((source_port, short_escape))
        # A preferred decision escape can land inside a neighboring node in a
        # dense map. Keep the semantic branch index, but evaluate other native
        # diamond ports before declaring the edge unroutable.
        for candidate in _diamond_port_candidates(a, b, clearance):
            port_already_used = any(
                abs(segment[0][0] - candidate[0][0]) < 0.015 and
                abs(segment[0][1] - candidate[0][1]) < 0.015
                for segment in existing_segments
            )
            if candidate not in source_candidates and not port_already_used:
                source_candidates.append(candidate)
    else:
        source_candidates = _port_candidates(a, b, clearance, outgoing=True)
    target_candidates = (_diamond_port_candidates(b, a, clearance)
                         if target_is_decision else
                         _port_candidates(b, a, clearance, outgoing=False))
    obstacles = [box for box in geometry.values() if box not in {a, b}]
    endpoint_diamonds = ([a] if source_is_decision else []) + ([b] if target_is_decision else [])
    candidates = []; preferred_candidates = []
    for source_candidate_index, (source_port, source_escape) in enumerate(source_candidates):
        source_stub = (source_port, source_escape)
        if any(_segment_hits_box(source_stub, obstacle, 0.045) for obstacle in obstacles):
            continue
        for target_port, target_escape in target_candidates:
            target_stub = (target_escape, target_port)
            if any(_segment_hits_box(target_stub, obstacle, 0.045) for obstacle in obstacles):
                continue
            middle = _orthogonal_grid_path(source_escape, target_escape, obstacles, existing_segments,
                                           endpoint_diamonds)
            if middle:
                points = [source_port]
                if source_escape != source_port:
                    points.append(source_escape)
                points.extend(middle[1:])
                if target_port != target_escape:
                    points.append(target_port)
                points = _simplify_route(points)
                source_inner = Box(a.x + 0.01, a.y + 0.01,
                                   max(0.01, a.w - 0.02), max(0.01, a.h - 0.02))
                target_inner = Box(b.x + 0.01, b.y + 0.01,
                                   max(0.01, b.w - 0.02), max(0.01, b.h - 0.02))
                route_segments = list(zip(points, points[1:]))
                source_reentry = any(
                    (_segment_hits_diamond(segment, source_inner, -0.005)
                     if source_is_decision else
                     _segment_hits_box(segment, source_inner, 0.0))
                    for segment in route_segments[1:])
                target_early_entry = any(
                    (_segment_hits_diamond(segment, target_inner, -0.005)
                     if target_is_decision else
                     _segment_hits_box(segment, target_inner, 0.0))
                    for segment in route_segments[:-1])
                if source_reentry or target_early_entry:
                    continue
                candidates.append(points)
                if source_is_decision and source_port == source_candidates[0][0]:
                    preferred_candidates.append(points)
    if not candidates:
        raise ValueError("No obstacle-free orthogonal route found between process-map nodes "
                         f"at ({a.cx:.3f}, {a.cy:.3f}) and ({b.cx:.3f}, {b.cy:.3f}).")
    pool = preferred_candidates or candidates
    return min(pool, key=lambda route: _route_score(route, a, b, geometry, existing_segments))


def _port_candidates(box: Box, other: Box, clearance: float, outgoing: bool):
    ports = {
        "right": ((box.right, box.cy), (box.right + clearance, box.cy)),
        "left": ((box.x, box.cy), (box.x - clearance, box.cy)),
        "bottom": ((box.cx, box.bottom), (box.cx, box.bottom + clearance)),
        "top": ((box.cx, box.y), (box.cx, box.y - clearance)),
    }
    dx, dy = other.cx - box.cx, other.cy - box.cy
    preferred = ("right" if dx >= 0 else "left") if abs(dx) >= abs(dy) else ("bottom" if dy >= 0 else "top")
    order = [preferred] + [side for side in ("right", "left", "bottom", "top") if side != preferred]
    result = []
    for side in order:
        port, escape = ports[side]
        escape = (min(13.00, max(0.30, escape[0])), min(7.16, max(0.94, escape[1])))
        # For incoming paths, the external point precedes the node attachment.
        result.append((port, escape) if outgoing else (port, escape))
    return result


def _diamond_port_candidates(box: Box, other: Box, clearance: float):
    """Return boundary-correct diamond ports with external escape points."""
    dx, dy = other.cx - box.cx, other.cy - box.cy
    preferred = (("right" if dx >= 0 else "left") if abs(dx) >= abs(dy)
                 else ("bottom" if dy >= 0 else "top"))
    order = [preferred] + [side for side in ("right", "left", "bottom", "top")
                           if side != preferred]
    result = []
    for side in order:
        port = _diamond_boundary_port(box, side)
        if side == "right":
            escape = (port[0] + clearance, port[1])
        elif side == "left":
            escape = (port[0] - clearance, port[1])
        elif side == "top":
            escape = (port[0], port[1] - clearance)
        else:
            escape = (port[0], port[1] + clearance)
        escape = (min(13.00, max(0.30, escape[0])),
                  min(7.16, max(0.94, escape[1])))
        result.append((port, escape))
    return result


def _orthogonal_grid_path(start, end, obstacles, existing_segments, diamond_obstacles=()):
    _routing_budget_check("orthogonal obstacle routing")
    if not all(math.isfinite(value) for value in (*start, *end)):
        raise ValueError(f"Routing endpoints must be finite: start={start}, end={end}.")
    normalize = lambda point: (round(float(point[0]), 3), round(float(point[1]), 3))
    start = normalize(start); end = normalize(end)
    if start == end:
        return [start]
    margin = 0.10
    # Bounds expand to include valid attachment escapes. Previously endpoints
    # below y=0.88 were inserted into `points` after the axis arrays had been
    # filtered, so x_index/y_index lookups raised KeyError (for example 0.835).
    min_x, max_x = min(0.88, start[0], end[0]), max(13.00, start[0], end[0])
    min_y, max_y = min(0.88, start[1], end[1]), max(7.16, start[1], end[1])
    xs = {start[0], end[0], 0.92, 13.00}
    ys = {start[1], end[1], 0.94, 7.16}
    for box in obstacles:
        xs.update((round(box.x - margin, 3), round(box.right + margin, 3), round(box.cx, 3)))
        ys.update((round(box.y - margin, 3), round(box.bottom + margin, 3), round(box.cy, 3)))
    for box in diamond_obstacles:
        xs.update((round(box.x - margin, 3), round(box.right + margin, 3), round(box.cx, 3)))
        ys.update((round(box.y - margin, 3), round(box.bottom + margin, 3), round(box.cy, 3)))
    xs = sorted(value for value in xs if min_x <= value <= max_x)
    ys = sorted(value for value in ys if min_y <= value <= max_y)
    points = {(x, y) for x in xs for y in ys
              if not any(box.x - 0.055 < x < box.right + 0.055 and
                         box.y - 0.055 < y < box.bottom + 0.055 for box in obstacles)
              and not any(_point_inside_diamond((x, y), box, 0.015)
                          for box in diamond_obstacles)}
    points.update((start, end))
    x_index = {value: index for index, value in enumerate(xs)}
    y_index = {value: index for index, value in enumerate(ys)}
    queue = [(0.0, 0.0, start, None)]
    best = {(start, None): 0.0}; previous = {}
    final_state = None
    expansions = 0
    while queue:
        expansions += 1
        _ROUTING_STATE["grid_expansions"] += 1
        if expansions > 5000:
            _ROUTING_STATE["grid_budget_exhaustions"] += 1
            return None
        if expansions % 128 == 0:
            _routing_budget_check("orthogonal obstacle routing")
        _, cost, point, direction = heapq.heappop(queue)
        if cost != best.get((point, direction)):
            continue
        if point == end:
            final_state = (point, direction); break
        xi, yi = x_index[point[0]], y_index[point[1]]
        neighbors = []
        if xi: neighbors.append((xs[xi - 1], point[1], "h"))
        if xi + 1 < len(xs): neighbors.append((xs[xi + 1], point[1], "h"))
        if yi: neighbors.append((point[0], ys[yi - 1], "v"))
        if yi + 1 < len(ys): neighbors.append((point[0], ys[yi + 1], "v"))
        for nx, ny, next_direction in neighbors:
            neighbor = (nx, ny)
            if neighbor not in points:
                continue
            segment = (point, neighbor)
            if any(_segment_hits_box(segment, box, 0.045) for box in obstacles):
                continue
            if any(_segment_hits_diamond(segment, box, 0.012) for box in diamond_obstacles):
                continue
            length = _segment_length(point, neighbor)
            relations = [_segment_relation((point, neighbor, "", ""), prior)
                         for prior in existing_segments]
            crossing_cost = relations.count("cross") * 1.7
            overlap_cost = relations.count("overlap") * 7.0
            turn_cost = 0.16 if direction and direction != next_direction else 0.0
            next_cost = cost + length + crossing_cost + overlap_cost + turn_cost
            state = (neighbor, next_direction)
            if next_cost + 1e-9 < best.get(state, float("inf")):
                best[state] = next_cost; previous[state] = (point, direction)
                estimate = abs(neighbor[0] - end[0]) + abs(neighbor[1] - end[1])
                heapq.heappush(queue, (next_cost + estimate, next_cost, neighbor, next_direction))
    if final_state is None:
        return None
    route = []
    state = final_state
    while state:
        route.append(state[0]); state = previous.get(state)
    return list(reversed(route))


def _simplify_route(points):
    result = []
    for point in points:
        if result and point == result[-1]:
            continue
        result.append(point)
        while len(result) >= 3:
            a, b, c = result[-3:]
            vertical = (abs(a[0] - b[0]) < 0.005 and
                        abs(b[0] - c[0]) < 0.005 and
                        min(a[1], c[1]) - 0.005 <= b[1] <= max(a[1], c[1]) + 0.005)
            horizontal = (abs(a[1] - b[1]) < 0.005 and
                          abs(b[1] - c[1]) < 0.005 and
                          min(a[0], c[0]) - 0.005 <= b[0] <= max(a[0], c[0]) + 0.005)
            # Preserve a collinear reversal.  Removing its turnaround point can
            # replace a valid outside approach with a shortcut through the
            # destination shape, which caused the visible decision crossing
            # that V1.6.4's planned-geometry check missed.
            if vertical or horizontal:
                result.pop(-2)
            else:
                break
    return result


def _render_page(slide, nodes, actors, lr, page, theme, layout, edges, exception_mode=False,
                 continuation_style="minimal", rendering_mode="compact", style=None, workflow=None):
    style = style or {}
    workflow = workflow or {"nodes": list(nodes.values())}
    compact = rendering_mode == "compact"
    spacing = style.get("spacing", "standard")
    margin_setting = style.get("layout_preferences", {}).get("content_margins", spacing)
    top, bottom = {"compact": (1.00, 7.24), "standard": ((1.02, 7.22) if compact else (1.08, 7.12)),
                   "spacious": (1.12, 7.08)}[margin_setting]
    left, right = {"compact": (1.25, 12.98), "standard": ((1.35, 12.90) if compact else (1.55, 12.72)),
                   "spacious": (1.62, 12.65)}[margin_setting]
    page_nodes = [nid for nid, p in lr.placements.items() if p.page == page]
    crossing_edges = [
        edge for edge in edges
        if edge["from"] in lr.placements and edge["to"] in lr.placements and
        lr.placements[edge["from"]].page != lr.placements[edge["to"]].page and
        page in {lr.placements[edge["from"]].page, lr.placements[edge["to"]].page}
    ]
    if crossing_edges and lr.continuation_space_reserve > 0:
        # Reserve a deterministic right-side continuation rail before placing
        # workflow nodes. Previously nodes consumed the full content width and
        # continuation search ran only afterward, so a dense but otherwise
        # valid page could have no legal marker location. More than two
        # handoffs receive a wider rail; pagination recovery handles any page
        # that can no longer fit readable node geometry.
        continuation_rail = max(lr.continuation_space_reserve,
                                1.18 if len(crossing_edges) <= 2 else 1.58)
        right -= continuation_rail
    if layout == "swimlane":
        active_ids = {nodes[nid].get("owner") for nid in page_nodes}
        page_actors = [actor for actor in actors if actor["id"] in active_ids]
    else:
        page_actors = actors
    owner_lane = {actor["id"]: index for index, actor in enumerate(page_actors)}
    lane_counts = [sum(nodes[nid].get("owner") == actor["id"] for nid in page_nodes)
                   if layout == "swimlane" else len(page_nodes) for actor in page_actors]
    weights = [1.0 + min(0.35, max(0, count - 2) * 0.08) for count in lane_counts]
    scale = (bottom - top) / sum(weights)
    lane_heights = [weight * scale for weight in weights]
    lane_tops = []
    lane_label_boxes: list[Box] = []
    cursor_y = top
    for i, actor in enumerate(page_actors):
        lane_tops.append(cursor_y); lane_h = lane_heights[i]; y = cursor_y; cursor_y += lane_h
        lane = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(y), Inches(12.7), Inches(lane_h))
        colors = style.get("colors", {})
        lane_color = colors.get("swimlane_alt_fill", theme.background) if i % 2 else colors.get("swimlane_fill", theme.lane_fill)
        lane.fill.solid(); lane.fill.fore_color.rgb = _rgb(lane_color)
        lane.line.color.rgb = _rgb(colors.get("activity_border", "D9DEE7")); lane.line.width = Pt(0.7)
        label = slide.shapes.add_textbox(Inches(0.32), Inches(y + 0.04), Inches(1.02), Inches(lane_h - 0.08))
        label.name = f"workflow-swimlane-label::{actor['id']}"
        _set_text(label, actor.get("label", actor["id"]), theme, 10, True,
                  colors.get("swimlane_header", theme.muted), style, "swimlane")
        lane_label_boxes.append(Box(0.32, y + 0.04, 1.02, lane_h - 0.08))

    columns = max(1, len(lr.page_ranks[page])); usable = right - left
    render_sizes = {nid: _node_size(nodes[nid], rendering_mode) for nid in page_nodes}
    column_widths = []
    for column in range(columns):
        widths = [render_sizes[nid][0] for nid in page_nodes
                  if lr.placements[nid].column == column]
        column_widths.append(max(widths, default=0.9))
    minimum_gap = 0.12
    required = sum(column_widths) + minimum_gap * max(0, columns - 1)
    if required > usable + 1e-6:
        # Post-pagination keep-together repairs can add a final outcome column.
        # Recover the small excess from activity/decision widths while keeping
        # terminal outcomes at their readable minimum; this is bounded and
        # applies only to the overfull slide.
        excess = required - usable
        column_floors = []
        for column in range(columns):
            floors = []
            for nid in page_nodes:
                if lr.placements[nid].column != column:
                    continue
                node_type = nodes[nid].get("type")
                floors.append(1.48 if node_type in {"end", "error"} else
                              1.55 if node_type == "decision" else 0.84)
            column_floors.append(max(floors, default=0.84))
        compressible = [max(0.0, width - floor)
                        for width, floor in zip(column_widths, column_floors)]
        available = sum(compressible)
        if available > 0:
            fraction = min(1.0, excess / available)
            column_widths = [width - room * fraction
                             for width, room in zip(column_widths, compressible)]
            for nid in page_nodes:
                column = lr.placements[nid].column
                width, height = render_sizes[nid]
                render_sizes[nid] = (min(width, column_widths[column]), height)
    total_width = sum(column_widths)
    gap = (usable - total_width) / max(1, columns - 1)
    gap = max(0.12, gap)
    used = total_width + gap * max(0, columns - 1)
    cursor = left + max(0.0, (usable - used) / 2)
    centers = []
    for width in column_widths:
        centers.append(cursor + width / 2)
        cursor += width + gap
    geometry: dict[str, Box] = {}
    lane_boxes: dict[int, list[Box]] = defaultdict(list)
    for nid in sorted(page_nodes, key=lambda n: (lr.placements[n].column, lr.placements[n].lane, n)):
        node, p = nodes[nid], lr.placements[nid]
        lane = owner_lane.get(node.get("owner"), 0) if layout == "swimlane" else 0
        lane_h = lane_heights[lane]; w, h = render_sizes[nid]
        x = centers[p.column] - w / 2; base_y = lane_tops[lane]
        if layout == "flowchart":
            siblings = sorted(value for value in page_nodes if lr.placements[value].column == p.column)
            position = siblings.index(nid)
            y = base_y + (position + 1) * lane_h / (len(siblings) + 1) - h / 2
        else:
            siblings = sorted(value for value in page_nodes
                              if lr.placements[value].column == p.column and
                              nodes[value].get("owner") == node.get("owner"))
            sibling_sizes = [_node_size(nodes[value], rendering_mode)[1] for value in siblings]
            stack_gap = 0.07
            stack_height = sum(sibling_sizes) + stack_gap * max(0, len(siblings) - 1)
            if len(siblings) > 1 and stack_height <= lane_h - 0.12:
                position = siblings.index(nid)
                y = (base_y + (lane_h - stack_height) / 2 +
                     sum(sibling_sizes[:position]) + stack_gap * position)
            else:
                y = base_y + (lane_h - h) / 2
        candidate = Box(x, y, w, h)
        if compact and any(_boxes_overlap(candidate, prior, 0.05) for prior in lane_boxes[lane]):
            upper = base_y + 0.10
            lower = base_y + lane_h - h - 0.10
            alternatives = [Box(x, upper, w, h), Box(x, lower, w, h)]
            candidate = min(alternatives, key=lambda box: sum(_boxes_overlap(box, prior, 0.04)
                                                               for prior in lane_boxes[lane]))
        geometry[nid] = candidate; lane_boxes[lane].append(candidate)
        x, y = candidate.x, candidate.y
        shape = slide.shapes.add_shape(_styled_shape_kind(node["type"], style), Inches(x), Inches(y), Inches(w), Inches(h))
        shape.name = f"workflow-node::{nid}"
        is_error = node["type"] == "error"
        category = _outcome_category(node, workflow, style)
        if category:
            fill, semantic_text_color, outcome_line = _outcome_paint(category, theme, style)
        else:
            fill = (theme.exception_soft if is_error else theme.decision if node["type"] == "decision"
                else "F3EEFF" if node["type"] == "manual_action"
                else "E8F4F0" if node["type"] == "external_system" else theme.accent_soft)
            semantic_text_color = theme.text
        line_color = (outcome_line if category else theme.exception if is_error else "7A5AF8" if node["type"] == "manual_action"
                      else theme.success if node["type"] == "external_system" else theme.accent)
        label = node["label"]
        if rendering_mode == "detailed" and node.get("description"):
            label += "\n" + node["description"]
        minimum = 8 if compact else 9
        size = minimum if compact or len(label) > 70 else 9
        if node["type"] == "decision":
            size = _decision_text_style(shape, style)
        fill, line_color, semantic_text_color, border_width, size = _node_visual(
            node, theme, style, fill, line_color, semantic_text_color,
            1.6 if node["type"] == "decision" else 1.25, size)
        shape.fill.solid(); shape.fill.fore_color.rgb = _rgb(fill)
        shape.line.color.rgb = _rgb(line_color); shape.line.width = Pt(border_width)
        _apply_shadow(shape, style.get("shapes", {}).get("shadows", False))
        explicit_size = "text_size" in style.get("elements", {}).get(nid, {})
        role = ("decision" if node["type"] == "decision" else "terminal"
                if node["type"] in {"start", "end", "error"} else "activity")
        _set_text(shape, label, theme, size, node["type"] in {"start", "end"} or bool(category),
                  semantic_text_color, style, role, explicit_size)
        if node.get("description"):
            shape._element.nvSpPr.cNvPr.set("descr", node["description"])
        if node["type"] == "decision":
            _decision_text_style(shape, style)

    local_edges = [e for e in edges if e["from"] in geometry and e["to"] in geometry]
    outgoing = defaultdict(list)
    # Count off-page outcomes when assigning decision ports.  Otherwise a
    # three-way decision is rendered as two local branches plus an unrelated
    # continuation label, and the visible ports no longer match the graph.
    for edge in edges:
        if edge["from"] in geometry:
            outgoing[edge["from"]].append(edge)
    for src in outgoing:
        outgoing[src].sort(key=lambda e: (geometry[e["to"]].cy if e["to"] in geometry else 99.0,
                                          e.get("label", ""), e["to"]))
    channel_use = defaultdict(int)
    label_boxes: list[Box] = []
    label_records: list[tuple[Box, str, str]] = []
    placement_warnings: list[str] = []
    connector_segments: list[tuple[tuple[float, float], tuple[float, float]]] = []
    pending_labels = []
    for edge in local_edges:
        pending = _route_edge(slide, edge, geometry[edge["from"]], geometry[edge["to"]], geometry, theme,
                              outgoing[edge["from"]].index(edge), len(outgoing[edge["from"]]), channel_use,
                              nodes[edge["from"]].get("type") == "decision",
                              nodes[edge["to"]].get("type") == "decision", connector_segments, style)
        if pending: pending_labels.append(pending)
    def place_local_labels():
        for text, points, source_is_decision, branch_index, branch_count, color, source, target in pending_labels:
            box = _place_edge_label(slide, text, points, source_is_decision, branch_index, branch_count,
                                    list(geometry.values()) + (lane_label_boxes if len(nodes) >= 45 else []),
                                    label_boxes, connector_segments, theme, color,
                                    style["branch_label_position"], style["connector_label_offset"], source, target,
                                    placement_warnings, style=style,
                                    large_graph=len(nodes) >= 45)
            label_records.append((box, source, target, points))

    if lr.continuation_space_reserve > 0:
        # Recovery candidates reserve cross-slide semantics first. Local labels
        # can move to alternate connector segments; continuations cannot vanish.
        continuation_count = _render_continuations(
            slide, page, edges, lr, nodes, geometry, label_boxes, connector_segments,
            theme, continuation_style, rendering_mode, style,
            protected_boxes=lane_label_boxes)
        place_local_labels()
    else:
        # Preserve the approved V1.6.2 placement order for layouts that do not
        # need adaptive continuation space.
        place_local_labels()
        continuation_count = _render_continuations(
            slide, page, edges, lr, nodes, geometry, label_boxes, connector_segments,
            theme, continuation_style, rendering_mode, style,
            protected_boxes=lane_label_boxes)
    quality_nodes = (nodes if rendering_mode == "detailed" else
                     {node_id: dict(node, description="") for node_id, node in nodes.items()})
    return {"continuation_count": continuation_count, "warnings": placement_warnings,
            "quality": _quality_report(geometry, connector_segments, label_records, quality_nodes,
                                       page_nodes, protected_boxes=(lane_label_boxes
                                                                   if len(nodes) >= 45 else []))}


def _route_edge(slide, edge, a: Box, b: Box, geometry, theme, branch_index, branch_count, channel_use,
                source_is_decision, target_is_decision, connector_segments, style=None):
    is_special = edge.get("path") in {"exception", "recovery"}; is_back = b.cx <= a.cx
    points = _compute_route_points(a, b, geometry, branch_index, branch_count, source_is_decision,
                                   is_back, channel_use, connector_segments,
                                   is_recovery=edge.get("path") == "recovery",
                                   target_is_decision=target_is_decision)
    style = style or {}
    color, width, line_style, arrowhead = _edge_visual(edge, theme, style, is_special,
                                                       1.6 if is_special else 1.15)
    for i, (p1, p2) in enumerate(zip(points, points[1:])):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(p1[0]), Inches(p1[1]), Inches(p2[0]), Inches(p2[1]))
        line.name = f"workflow-edge::{edge.get('_render_index', 0)}::segment-{i + 1}"
        line.line.color.rgb = _rgb(color); line.line.width = Pt(width); _set_dash(line, line_style)
        if i == len(points) - 2 and arrowhead != "none": _set_arrowhead(line)
        connector_segments.append((p1, p2, edge["from"], edge["to"]))
    if edge.get("label"):
        return (str(edge["label"]), points, source_is_decision, branch_index, branch_count,
                color if is_special else theme.muted, edge["from"], edge["to"])
    return None


def _compute_route_points(a: Box, b: Box, geometry, branch_index: int, branch_count: int,
                          source_is_decision: bool, is_back: bool, channel_use,
                          connector_segments=None, is_recovery: bool = False,
                          target_is_decision: bool = False) -> list[tuple[float, float]]:
    """Calculate a route while preserving a unique decision exit and the target attachment."""
    connector_segments = connector_segments or []
    # Use the same obstacle-aware grid as process maps.  The older heuristic
    # routes below remain a bounded fallback for unusually constrained pages,
    # but normal swimlane and flowchart connectors now validate every segment
    # against every unrelated node before drawing.
    try:
        routed = _process_map_route_points(a, b, geometry, branch_index, branch_count,
                                           source_is_decision, connector_segments,
                                           target_is_decision)
        # The preferred semantic port is tried first. In a dense graph it can
        # be blocked by an adjacent node; the router may then use another
        # unused native diamond port. Rejecting that valid alternate here used
        # to trigger a heuristic fallback that crossed the blocking diamond.
        return routed
    except ValueError:
        pass
    if is_back:
        # Return paths use an exterior corridor so they cannot be mistaken for
        # a forward approval or rejection branch.  Evaluate both sides because
        # the nearest corridor may already contain another loop.
        bottom = min(7.12, max(a.bottom, b.bottom) + 0.18 + channel_use["bottom"] * 0.12)
        top = max(0.98, min(a.y, b.y) - 0.18 - channel_use["top"] * 0.12)
        bottom_route = [(a.cx, a.bottom), (a.cx, bottom), (b.cx, bottom), (b.cx, b.bottom)]
        top_route = [(a.cx, a.y), (a.cx, top), (b.cx, top), (b.cx, b.y)]
        candidates = [bottom_route, top_route]
        chosen = _choose_endpoint_safe_route(candidates, a, b, source_is_decision,
                                             target_is_decision, geometry,
                                             connector_segments)
        channel_use["bottom" if chosen is bottom_route else "top"] += 1
        return _simplify_route(chosen)
    if source_is_decision:
        start, escape = _decision_exit(a, branch_index, branch_count, b)
        mid = max(escape[0] + 0.18, a.right + 0.34 + branch_index * 0.12, (a.right + b.x) / 2)
        obstacles = [box for box in geometry.values() if box not in {a, b} and box.x < b.x and
                     box.right > escape[0] and box.y - 0.06 <= escape[1] <= box.bottom + 0.06]
        if obstacles:
            upper_branch = branch_index <= (branch_count - 1) / 2
            channel = (max(1.12, min(box.y for box in obstacles) - 0.16) if upper_branch
                       else min(7.02, max(box.bottom for box in obstacles) + 0.16))
            preferred = [start, escape, (escape[0], channel), (b.x - 0.16, channel),
                         (b.x - 0.16, b.cy), (b.x, b.cy)]
        else:
            preferred = [start, escape, (mid, escape[1]), (mid, b.cy), (b.x, b.cy)]
        upper = max(1.04, min(a.y, b.y) - 0.18 - branch_index * 0.06)
        lower = min(7.18, max(a.bottom, b.bottom) + 0.18 + branch_index * 0.06)
        alternatives = [preferred,
                        [start, escape, (escape[0], upper), (b.x - 0.12, upper), (b.x - 0.12, b.cy), (b.x, b.cy)],
                        [start, escape, (escape[0], lower), (b.x - 0.12, lower), (b.x - 0.12, b.cy), (b.x, b.cy)]]
        return _choose_endpoint_safe_route(alternatives, a, b, source_is_decision,
                                           target_is_decision, geometry,
                                           connector_segments)
    branch_offset = (branch_index - (branch_count - 1) / 2) * 0.14
    start_y = a.cy + branch_offset
    obstacles = [box for box in geometry.values() if box not in {a, b} and
                 box.x < b.x and box.right > a.right and box.y - 0.06 <= start_y <= box.bottom + 0.06]
    if obstacles:
        channel = max(1.12, min(box.y for box in obstacles) - 0.16 - branch_index * 0.08)
        preferred = [(a.right, start_y), (a.right + 0.16, start_y), (a.right + 0.16, channel),
                     (b.x - 0.16, channel), (b.x - 0.16, b.cy), (b.x, b.cy)]
        lower = min(7.18, max(a.bottom, b.bottom) + 0.18 + branch_index * 0.06)
        alternate = [(a.right, start_y), (a.right + 0.12, start_y), (a.right + 0.12, lower),
                     (b.x - 0.12, lower), (b.x - 0.12, b.cy), (b.x, b.cy)]
        return _choose_endpoint_safe_route([preferred, alternate], a, b,
                                           source_is_decision, target_is_decision,
                                           geometry, connector_segments)
    mid = (a.right + b.x) / 2 + branch_offset
    preferred = [(a.right, start_y), (mid, start_y), (mid, b.cy), (b.x, b.cy)]
    upper = max(1.04, min(a.y, b.y) - 0.16)
    alternate = [(a.right, start_y), (a.right + 0.12, start_y), (a.right + 0.12, upper),
                 (b.x - 0.12, upper), (b.x - 0.12, b.cy), (b.x, b.cy)]
    return _choose_endpoint_safe_route([preferred, alternate], a, b,
                                       source_is_decision, target_is_decision,
                                       geometry, connector_segments)


def _choose_endpoint_safe_route(candidates, source: Box, target: Box,
                                source_is_decision: bool, target_is_decision: bool,
                                geometry, existing_segments):
    """Choose a route that does not leave and later re-enter either endpoint."""
    safe = []
    source_inner = Box(source.x + 0.01, source.y + 0.01,
                       max(0.01, source.w - 0.02), max(0.01, source.h - 0.02))
    target_inner = Box(target.x + 0.01, target.y + 0.01,
                       max(0.01, target.w - 0.02), max(0.01, target.h - 0.02))
    for points in candidates:
        points = _simplify_route(points)
        segments = list(zip(points, points[1:]))
        unrelated = [box for box in geometry.values()
                     if box is not source and box is not target]
        if any(_segment_hits_box(segment, box, 0.02)
               for segment in segments for box in unrelated):
            continue
        source_reentry = any(
            (_segment_hits_diamond(segment, source_inner, -0.005)
             if source_is_decision else _segment_hits_box(segment, source_inner, 0.0))
            for segment in segments[1:])
        target_early = any(
            (_segment_hits_diamond(segment, target_inner, -0.005)
             if target_is_decision else _segment_hits_box(segment, target_inner, 0.0))
            for segment in segments[:-1])
        if not source_reentry and not target_early:
            safe.append(points)
    if not safe:
        raise ValueError("No endpoint-safe orthogonal route found.")
    return min(safe, key=lambda points: _route_score(
        points, source, target, geometry, existing_segments))


def _route_score(points, source, target, geometry, existing_segments):
    segments = list(zip(points, points[1:]))
    unrelated = [box for box in geometry.values() if box not in {source, target}]
    node_hits = sum(_segment_hits_box(segment, box, 0.05) for segment in segments for box in unrelated)
    crossings = sum(_segment_relation((*segment, "", ""), prior) == "cross"
                    for segment in segments for prior in existing_segments)
    overlaps = sum(_segment_relation((*segment, "", ""), prior) == "overlap"
                   for segment in segments for prior in existing_segments)
    length = sum(_segment_length(*segment) for segment in segments)
    return node_hits * 10000 + overlaps * 250 + crossings * 35 + length


def _decision_exit(box: Box, branch_index: int, branch_count: int,
                   target: Box | None = None):
    """Return source and escape points for a distinct decision branch port."""
    if branch_count == 3:
        if branch_index == 0:
            port = _diamond_boundary_port(box, "top")
            return port, (port[0], port[1] - 0.42)
        if branch_index == 1:
            port = _diamond_boundary_port(box, "right")
            return port, (port[0] + 0.85, port[1])
        port = _diamond_boundary_port(box, "bottom")
        return port, (port[0], port[1] + 0.42)
    if branch_count == 2:
        if branch_index == 0:
            # When both outcomes are geometrically above the decision, sending
            # the second branch out of the bottom forces it to retrace the
            # first branch's vertical corridor. Keep the first outcome on the
            # top port and give the other a distinct right-side exit.
            # Apply the mirror rule when the first outcome lies below the
            # decision: it receives the right port while the second outcome
            # retains the bottom port.
            if target is not None and (
                    target.cy > box.cy + 0.08 or
                    (target.cx > box.cx and abs(target.cy - box.cy) <= 0.08)):
                port = _diamond_boundary_port(box, "right")
                return port, (port[0] + 0.52, port[1])
            port = _diamond_boundary_port(box, "top")
            return port, (port[0], port[1] - 0.38)
        if target is not None and target.cy < box.cy - 0.08:
            port = _diamond_boundary_port(box, "right")
            return port, (port[0] + 0.52, port[1])
        port = _diamond_boundary_port(box, "bottom")
        return port, (port[0], port[1] + 0.38)

    # Four outcomes receive all four native diamond exits.  Choosing every
    # exit from destination direction caused two same-side outcomes to share a
    # visible corridor before diverging, which looked like one ambiguous arrow
    # in the serialized presentation.
    if branch_count == 4:
        # Outgoing edges are ordered by target vertical position. Give the
        # uppermost and lowermost outcomes the top and bottom ports, then use
        # right and left for the two middle outcomes. This preserves four
        # distinct exits without sending a low terminal branch backward
        # through the activity immediately left of the decision.
        side = ("top", "right", "left", "bottom")[branch_index]
        port = _diamond_boundary_port(box, side)
        clearance = 0.48 + branch_index * 0.04
        if side == "right":
            return port, (port[0] + clearance, port[1])
        if side == "left":
            return port, (port[0] - clearance, port[1])
        if side == "top":
            return port, (port[0], port[1] - clearance)
        return port, (port[0], port[1] + clearance)

    # Five or more branches choose the diamond edge facing the destination.
    # Sibling branches on the same side receive separate points on that sloped
    # edge.  Unlike a rectangular anchor at (box.right, y), each port satisfies
    # the diamond boundary equation exactly.
    if target is not None:
        dx, dy = target.cx - box.cx, target.cy - box.cy
        side = (("right" if dx >= 0 else "left") if abs(dx) >= abs(dy)
                else ("bottom" if dy >= 0 else "top"))
    else:
        side = ("top", "right", "bottom", "left")[branch_index % 4]
    centered = branch_index - (branch_count - 1) / 2
    offset = max(-0.62, min(0.62, centered * 0.24))
    port = _diamond_boundary_port(box, side, offset)
    clearance = 0.48 + min(0.28, branch_index * 0.04)
    if side == "right":
        escape = (port[0] + clearance, port[1])
    elif side == "left":
        escape = (port[0] - clearance, port[1])
    elif side == "top":
        escape = (port[0], port[1] - clearance)
    else:
        escape = (port[0], port[1] + clearance)
    return port, escape


def _place_edge_label(slide, text, points, source_is_decision, branch_index, branch_count,
                      node_boxes, label_boxes, connector_segments, theme, color,
                      position="centered", offset="standard", source=None, target=None, warnings=None,
                      measured=True, style=None, large_graph=False):
    """Place a transparent edge label beside a clean segment without hiding the route."""
    width, height = (_edge_label_size(str(text)) if measured else
                     (min(1.10, max(0.44, 0.20 + len(str(text)) * 0.075)), 0.24))
    segments = list(zip(points, points[1:]))
    horizontal = [(i, p1, p2) for i, (p1, p2) in enumerate(segments)
                  if abs(p1[1] - p2[1]) < 0.01 and abs(p2[0] - p1[0]) >= width + 0.12]
    vertical = [(i, p1, p2) for i, (p1, p2) in enumerate(segments)
                if abs(p1[0] - p2[0]) < 0.01 and abs(p2[1] - p1[1]) >= height + 0.12]
    eligible = sorted(horizontal + vertical, key=lambda s: s[0])
    if position == "near_source":
        ordered = eligible
    elif position == "near_target":
        ordered = list(reversed(eligible))
    else:
        ordered = sorted(eligible, key=lambda s: -_segment_length(s[1], s[2]))
    candidates: list[Box] = []
    line_gap = {"tight": 0.06, "standard": 0.10, "wide": 0.18}[offset]
    label_channel = branch_index - (branch_count - 1) / 2
    for _, p1, p2 in ordered:
        if abs(p1[1] - p2[1]) < 0.01:
            midpoint = (p1[0] + p2[0]) / 2
            quarter = p1[0] + (p2[0] - p1[0]) * 0.28
            three_quarter = p1[0] + (p2[0] - p1[0]) * 0.72
            anchors = ([quarter, midpoint, three_quarter] if position == "near_source" else
                       [three_quarter, midpoint, quarter] if position == "near_target" else
                       [midpoint, quarter, three_quarter])
            # Above is preferred. Independent branch offsets keep sibling labels apart.
            candidates.extend([
                Box(anchors[0] - width / 2, p1[1] - height - line_gap - max(0, -label_channel) * 0.05, width, height),
                Box(anchors[0] - width / 2, p1[1] + line_gap + max(0, label_channel) * 0.05, width, height),
                Box(anchors[1] - width / 2, p1[1] - height - line_gap, width, height),
                Box(anchors[2] - width / 2, p1[1] - height - line_gap, width, height),
                Box(anchors[1] - width / 2, p1[1] + line_gap, width, height),
                Box(anchors[2] - width / 2, p1[1] + line_gap, width, height),
                Box(midpoint - width / 2, p1[1] - height - 0.34, width, height),
                Box(midpoint - width / 2, p1[1] + 0.34, width, height),
                Box(quarter - width / 2, p1[1] - height - 0.52, width, height),
                Box(three_quarter - width / 2, p1[1] + 0.52, width, height),
            ])
        else:
            midpoint = (p1[1] + p2[1]) / 2
            candidates.extend([
                Box(p1[0] + line_gap, midpoint - height / 2, width, height),
                Box(p1[0] - width - line_gap, midpoint - height / 2, width, height),
                Box(p1[0] + 0.32, midpoint - height / 2, width, height),
                Box(p1[0] - width - 0.32, midpoint - height / 2, width, height),
            ])
            # The first segment leaving a diamond is intentionally short. Put
            # labels nearer its escape end so their boxes clear the sloping
            # diamond boundary while remaining visibly tied to that branch.
            for fraction in (0.72, 0.88):
                anchor = p1[1] + (p2[1] - p1[1]) * fraction
                candidates.append(Box(p1[0] + line_gap, anchor - height / 2, width, height))
                candidates.append(Box(p1[0] - width - line_gap, anchor - height / 2, width, height))
        if measured:
            # Dense diagrams need more than a midpoint/quarter fallback.  Sample
            # the actual segment at deterministic intervals and on both sides.
            for fraction in (0.16, 0.28, 0.40, 0.60, 0.72, 0.84):
                if abs(p1[1] - p2[1]) < 0.01:
                    anchor = p1[0] + (p2[0] - p1[0]) * fraction
                    for gap in (line_gap, line_gap + 0.10, line_gap + 0.22, line_gap + 0.38):
                        candidates.append(Box(anchor - width / 2, p1[1] - height - gap, width, height))
                        candidates.append(Box(anchor - width / 2, p1[1] + gap, width, height))
                else:
                    anchor = p1[1] + (p2[1] - p1[1]) * fraction
                    for gap in (line_gap, line_gap + 0.10, line_gap + 0.22, line_gap + 0.38):
                        candidates.append(Box(p1[0] + gap, anchor - height / 2, width, height))
                        candidates.append(Box(p1[0] - width - gap, anchor - height / 2, width, height))
    source_x, source_y = points[0]
    nearby_source_candidates = [
        Box(source_x + 0.10, source_y - height - 0.28, width, height),
        Box(source_x + 0.10, source_y + 0.28, width, height),
        Box(source_x - width - 0.10, source_y - height - 0.28, width, height),
        Box(source_x - width - 0.10, source_y + 0.28, width, height),
        Box(source_x + 0.18, source_y - height - 0.62, width, height),
        Box(source_x + 0.18, source_y + 0.62, width, height),
        Box(source_x - width - 0.18, source_y - height - 0.62, width, height),
        Box(source_x - width - 0.18, source_y + 0.62, width, height),
        # Short node-to-node gaps may not have enough room for even a compact
        # label beside the connector.  Search farther along the open vertical
        # space around the source before accepting a colliding fallback.
        Box(source_x + 0.10, source_y - height - 0.88, width, height),
        Box(source_x + 0.10, source_y + 0.88, width, height),
        Box(source_x - width - 0.10, source_y - height - 0.88, width, height),
        Box(source_x - width - 0.10, source_y + 0.88, width, height),
        Box(source_x + 0.34, source_y - height - 1.08, width, height),
        Box(source_x + 0.34, source_y + 1.08, width, height),
        Box(source_x - width - 0.34, source_y - height - 1.08, width, height),
        Box(source_x - width - 0.34, source_y + 1.08, width, height),
    ]
    source_node_box = None
    if source_is_decision and position == "near_source":
        # Anchor a label against the actual diamond bounds as well as the
        # connector point. A top-exit branch near the slide header can leave a
        # narrow but valid pocket above the diamond; fixed offsets from the
        # port may miss that pocket by a few hundredths of an inch and trigger
        # a real label-node collision.
        source_node_box = next((box for box in node_boxes
                                if box.x - 0.02 <= source_x <= box.right + 0.02 and
                                box.y - 0.02 <= source_y <= box.bottom + 0.02), None)
        boundary_candidates = []
        if source_node_box is not None:
            boundary_gap = 0.045
            # Associate each outcome with the first segment leaving its native
            # diamond port.  In a four-way fan these side-specific positions
            # are both closer and less ambiguous than a shared row above or
            # below the decision.
            if large_graph and len(points) > 1:
                first_end = points[1]
                if abs(first_end[0] - source_x) < 0.01:  # top or bottom exit
                    outward_y = min(source_y, first_end[1]) if first_end[1] < source_y else max(source_y, first_end[1])
                    boundary_candidates.extend([
                        Box(source_x + 0.07, outward_y - height / 2, width, height),
                        Box(source_x - width - 0.07, outward_y - height / 2, width, height),
                    ])
                else:  # left or right exit
                    outward_x = min(source_x, first_end[0]) if first_end[0] < source_x else max(source_x, first_end[0])
                    boundary_candidates.extend([
                        Box(outward_x - width / 2, source_y - height - 0.07, width, height),
                        Box(outward_x - width / 2, source_y + 0.07, width, height),
                    ])
            boundary_candidates = [
                *boundary_candidates,
                Box(source_x + boundary_gap, source_node_box.y - height - 0.04, width, height),
                Box(source_x - width - boundary_gap, source_node_box.y - height - 0.04, width, height),
                Box(source_x + boundary_gap, source_node_box.bottom + 0.04, width, height),
                Box(source_x - width - boundary_gap, source_node_box.bottom + 0.04, width, height),
            ]
            # The rectangular corners around a diamond are valid label space.
            # Sample just outside both sloping sides so a label can remain near
            # its branch even when a parallel routing corridor blocks the
            # diamond's top or bottom vertex.
            for fraction in (0.12, 0.22, 0.34, 0.66, 0.78, 0.88):
                anchor_y = source_node_box.y + source_node_box.h * fraction
                label_top = anchor_y - height / 2 - 0.04
                label_bottom = anchor_y + height / 2 + 0.04
                closest_y = min(max(source_node_box.cy, label_top), label_bottom)
                normalized = abs(closest_y - source_node_box.cy) / max(source_node_box.h / 2, 1e-9)
                x_reach = source_node_box.w / 2 * max(0.0, 1.0 - normalized)
                boundary_candidates.append(
                    Box(source_node_box.cx + x_reach + boundary_gap,
                        anchor_y - height / 2, width, height))
                boundary_candidates.append(
                    Box(source_node_box.cx - x_reach - boundary_gap - width,
                        anchor_y - height / 2, width, height))
            # Sibling branches often form a horizontal corridor immediately
            # above or below the diamond. Place the current label just beyond
            # that corridor instead of letting its text box straddle the line.
            for segment in connector_segments:
                if (segment[2], segment[3]) == (source, target):
                    continue
                p1, p2 = segment[:2]
                if abs(p1[1] - p2[1]) >= 0.01:
                    continue
                if max(p1[0], p2[0]) < source_node_box.x - 0.30 or min(p1[0], p2[0]) > source_node_box.right + 1.50:
                    continue
                corridor_y = p1[1]
                for candidate_x in (source_node_box.right + boundary_gap,
                                    source_x + boundary_gap,
                                    source_x - width - boundary_gap):
                    boundary_candidates.append(
                        Box(candidate_x, corridor_y - height - 0.05, width, height))
                    boundary_candidates.append(
                        Box(candidate_x, corridor_y + 0.05, width, height))
        source_candidates = [
            Box(source_x + 0.08, source_y - height - 0.30, width, height),
            Box(source_x + 0.08, source_y + 0.30, width, height),
            Box(source_x - width / 2, source_y - height - 0.34, width, height),
            Box(source_x - width / 2, source_y + 0.34, width, height),
            Box(source_x + 0.18, source_y - height - 0.55, width, height),
            Box(source_x + 0.18, source_y + 0.55, width, height),
            Box(source_x - width - 0.55, source_y - height - 0.55, width, height),
            Box(source_x - width - 0.55, source_y + 0.55, width, height),
            Box(source_x + 0.55, source_y - height - 0.55, width, height),
            Box(source_x + 0.55, source_y + 0.55, width, height),
            Box(source_x - width / 2, source_y - height - 0.78, width, height),
            Box(source_x - width / 2, source_y + 0.78, width, height),
            Box(source_x - width / 2, source_y - height - 1.05, width, height),
            Box(source_x - width / 2, source_y + 1.05, width, height),
            Box(source_x + 0.85, source_y - height - 0.72, width, height),
            Box(source_x - width - 0.85, source_y - height - 0.72, width, height),
        ]
        candidates = boundary_candidates + source_candidates + candidates
    else:
        candidates = nearby_source_candidates + candidates
    # A dense branch fan can block the preferred offset even though a compact
    # position beside another point on the same segment remains unambiguous.
    # Sample both sides of every real segment with a small but visible 3.2 pt
    # gap before accepting a fallback that is closer to another route.
    compact_gap = 0.045
    for p1, p2 in segments:
        for fraction in (0.18, 0.36, 0.64, 0.82):
            anchor_x = p1[0] + (p2[0] - p1[0]) * fraction
            anchor_y = p1[1] + (p2[1] - p1[1]) * fraction
            if abs(p1[1] - p2[1]) < 0.01:
                candidates.append(Box(anchor_x - width / 2, anchor_y - height - compact_gap, width, height))
                candidates.append(Box(anchor_x - width / 2, anchor_y + compact_gap, width, height))
            else:
                candidates.append(Box(anchor_x + compact_gap, anchor_y - height / 2, width, height))
                candidates.append(Box(anchor_x - width - compact_gap, anchor_y - height / 2, width, height))
    # Short merge segments may be too small for the full label width. Search
    # around their real elbows before accepting a fallback over the source
    # node or a sibling branch label. These positions remain anchored to an
    # actual route vertex and therefore preserve edge association.
    for vertex_x, vertex_y in points[1:-1]:
        for radius in (0.18, 0.34, 0.52, 0.72):
            candidates.extend([
                Box(vertex_x + radius, vertex_y - height - radius, width, height),
                Box(vertex_x - width - radius, vertex_y - height - radius, width, height),
                Box(vertex_x + radius, vertex_y + radius, width, height),
                Box(vertex_x - width - radius, vertex_y + radius, width, height),
                Box(vertex_x - width / 2, vertex_y - height - radius, width, height),
                Box(vertex_x - width / 2, vertex_y + radius, width, height),
            ])
        # Dense merges sometimes leave a clear pocket that does not align
        # with the symmetric offsets above. Use a bounded deterministic grid
        # around the elbow; clean() still requires slide, node, label,
        # connector, arrowhead, and edge-association clearance.
        for x_step in range(-11, 7):
            for y_step in range(-8, 11):
                candidates.append(Box(vertex_x + x_step * 0.08,
                                      vertex_y + y_step * 0.08,
                                      width, height))
    # Exhaust the local source neighbourhood before accepting a colliding
    # label.  Dense four-way decisions and continuation rails can block every
    # symmetric candidate even though a clean pocket exists slightly farther
    # away.  The deterministic distance ordering keeps the label associated
    # with its branch while avoiding a global, arbitrary placement.
    if large_graph:
        neighbourhood = []
        for x_step in range(-20, 21):
            for y_step in range(-18, 19):
                if abs(x_step) + abs(y_step) > 28:
                    continue
                neighbourhood.append((abs(x_step) + abs(y_step), abs(y_step), abs(x_step),
                                      Box(source_x + x_step * 0.11 - width / 2,
                                          source_y + y_step * 0.11 - height / 2,
                                          width, height)))
        candidates.extend(item[3] for item in sorted(neighbourhood, key=lambda item: item[:3]))
    if not candidates:
        candidates = [Box(points[0][0] + 0.10, points[0][1] - height - 0.10, width, height)]
    blockers = node_boxes + label_boxes
    own_segments = [segment for segment in connector_segments
                    if (segment[2], segment[3]) == (source, target)]
    def clean(box):
        overlaps_blocker = any(
            (_box_overlaps_diamond(box, other, 0.04)
             if source_is_decision and source_node_box is not None and other is source_node_box
             else _boxes_overlap(box, other, 0.04))
            for other in blockers
        )
        if not _inside_slide(box) or overlaps_blocker:
            return False
        if any(_segment_hits_box(segment, box, 0.025) for segment in own_segments):
            return False
        if any(_segment_hits_box(segment, box, 0.025) for segment in connector_segments
               if (segment[2], segment[3]) != (source, target)):
            return False
        if not measured:
            return True
        # Keep labels clear of the destination arrowhead and require them to
        # remain visibly associated with their own route.
        arrow_x, arrow_y = points[-1]
        arrow_box = Box(arrow_x - 0.10, arrow_y - 0.10, 0.20, 0.20)
        own_distance = _box_route_distance(box, own_segments)
        other_routes = [segment for segment in connector_segments
                        if (segment[2], segment[3]) != (source, target)]
        other_distance = _box_route_distance(box, other_routes) if other_routes else 99.0
        # A label beside a very short segment can require a wider perpendicular
        # offset.  It remains acceptable when its own route is still the
        # nearest route; the quality report records proximity separately.
        max_route_distance = 1.02 if _segment_length(points[0], points[-1]) < width + 0.18 else 0.62
        return (not _boxes_overlap(box, arrow_box, 0.04) and own_distance <= max_route_distance and
                own_distance <= other_distance + 0.16)
    chosen = next((box for box in candidates if clean(box)), None)
    if chosen is None:
        # Preserve visibility without hiding connector segments; use the least crowded alternate.
        valid = [box for box in candidates if _inside_slide(box)]
        chosen = min(valid,
                     key=lambda box: (sum(_boxes_overlap(box, other, 0.0) for other in blockers),
                                      sum(_segment_hits_box(segment, box, 0.015)
                                          for segment in own_segments),
                                      sum(_segment_hits_box(segment, box, 0.0) for segment in connector_segments
                                          if (segment[2], segment[3]) != (source, target)),
                                      (_box_separation(box, source_node_box)
                                       if source_is_decision and source_node_box is not None and
                                       position == "near_source" else 0.0),
                                      max(0.0, _box_route_distance(box, own_segments) -
                                          _box_route_distance(box, [segment for segment in connector_segments
                                                                    if (segment[2], segment[3]) != (source, target)]) - 0.16),
                                      _box_route_distance(box, own_segments)), default=candidates[0])
        if warnings is not None:
            warnings.append(f"No collision-free connector-label position for {source} -> {target}; "
                            "the least ambiguous position was used.")
    tag = slide.shapes.add_textbox(Inches(chosen.x), Inches(chosen.y), Inches(chosen.w), Inches(chosen.h))
    tag.name = f"workflow-edge-label::{source or ''}->{target or ''}"
    style = style or {}
    tag.fill.background(); tag.line.fill.background()
    label_color = style.get("colors", {}).get("connector_label", color)
    _set_text(tag, text, theme, style.get("typography", {}).get("branch_label_size", 9),
              False, label_color, style, "branch_label")
    tag.text_frame.margin_left = tag.text_frame.margin_right = Inches(0.01)
    tag.text_frame.margin_top = tag.text_frame.margin_bottom = 0
    label_boxes.append(chosen)
    return chosen


def _edge_label_size(text: str, font_pt: float = 9.0) -> tuple[float, float]:
    """Estimate a compact text box using whole words and the requested font."""
    words = text.split() or [text]
    longest = max((len(word) for word in words), default=1)
    natural = 0.18 + sum((0.051 if char.islower() else 0.058) for char in text) * font_pt / 9.0
    width = min(1.10, max(0.48, 0.18 + longest * 0.058, natural))
    chars_per_line = max(longest, int((width - 0.10) / 0.055))
    lines = 1; used = 0
    for word in words:
        needed = len(word) + (1 if used else 0)
        if used and used + needed > chars_per_line:
            lines += 1; used = len(word)
        else:
            used += needed
    return width, max(0.24, 0.17 * lines + 0.07)


def _box_route_distance(box: Box, segments) -> float:
    if not segments:
        return 99.0
    cx, cy = box.cx, box.cy
    distances = []
    for segment in segments:
        (x1, y1), (x2, y2) = segment[:2]
        if abs(y1 - y2) < 0.01:
            px = min(max(cx, min(x1, x2)), max(x1, x2)); py = y1
        else:
            px = x1; py = min(max(cy, min(y1, y2)), max(y1, y2))
        distances.append(abs(cx - px) + abs(cy - py))
    return min(distances)


def _segment_length(p1, p2):
    return abs(p2[0] - p1[0]) + abs(p2[1] - p1[1])


def _inside_slide(box: Box):
    return 0.28 <= box.x and box.right <= 13.02 and 0.92 <= box.y and box.bottom <= 7.18


def _inside_slide_point(point):
    return 0.28 <= point[0] <= 13.02 and 0.92 <= point[1] <= 7.18


def _boxes_overlap(a: Box, b: Box, gap: float):
    return not (a.right + gap <= b.x or b.right + gap <= a.x or a.bottom + gap <= b.y or b.bottom + gap <= a.y)


def _box_separation(a: Box, b: Box) -> float:
    dx = max(a.x - b.right, b.x - a.right, 0.0)
    dy = max(a.y - b.bottom, b.y - a.bottom, 0.0)
    return dx + dy


def _segment_hits_box(segment, box: Box, gap: float):
    (x1, y1), (x2, y2) = segment[:2]
    if abs(y1 - y2) < 0.01:
        return box.y - gap <= y1 <= box.bottom + gap and max(min(x1, x2), box.x - gap) <= min(max(x1, x2), box.right + gap)
    if abs(x1 - x2) < 0.01:
        return box.x - gap <= x1 <= box.right + gap and max(min(y1, y2), box.y - gap) <= min(max(y1, y2), box.bottom + gap)
    return False


def _segment_hits_diamond(segment, box: Box, gap: float = 0.0) -> bool:
    """Return whether an orthogonal segment enters a diamond collision envelope."""
    (x1, y1), (x2, y2) = segment[:2]
    half_w, half_h = max(box.w / 2, 1e-9), max(box.h / 2, 1e-9)
    if abs(y1 - y2) < 0.01:
        normalized = abs(y1 - box.cy) / half_h
        if normalized > 1.0 + gap / half_h:
            return False
        reach = half_w * max(0.0, 1.0 - min(1.0, normalized)) + gap
        low, high = box.cx - reach, box.cx + reach
        return max(min(x1, x2), low) < min(max(x1, x2), high) - 1e-6
    if abs(x1 - x2) < 0.01:
        normalized = abs(x1 - box.cx) / half_w
        if normalized > 1.0 + gap / half_w:
            return False
        reach = half_h * max(0.0, 1.0 - min(1.0, normalized)) + gap
        low, high = box.cy - reach, box.cy + reach
        return max(min(y1, y2), low) < min(max(y1, y2), high) - 1e-6
    return False


def _merge_quality(total, page):
    for key, value in page.items():
        if isinstance(value, list):
            if not isinstance(total.get(key), list):
                total[key] = []
            total[key].extend(value)
        else:
            total[key] += value


def _quality_report(geometry, segments, label_records, nodes, page_nodes, protected_boxes=None):
    """Measure rendered geometry; source/target attachment segments are excluded explicitly."""
    protected_boxes = protected_boxes or []
    node_items = list(geometry.items())
    node_overlap_details = [
        [first_id, second_id]
        for index, (first_id, first_box) in enumerate(node_items)
        for second_id, second_box in node_items[index + 1:]
        if _boxes_overlap(first_box, second_box, 0.02)
    ]
    node_overlaps = len(node_overlap_details)
    connector_node = 0; connector_diamond = 0; connector_text = 0; arrowhead_overlap = 0
    connector_node_details = []
    connector_off_slide = sum(not (_inside_slide_point(segment[0]) and _inside_slide_point(segment[1]))
                              for segment in segments)
    for segment in segments:
        _, _, source, target = segment
        intended_nodes = {source, target}
        for token in (source, target):
            if isinstance(token, str) and token.startswith("continuation:"):
                intended_nodes.add(token.split(":", 2)[1])
        for node_id, box in node_items:
            if node_id not in intended_nodes and _segment_hits_box(segment, box, 0.01):
                connector_node += 1
                connector_node_details.append({
                    "source": source, "target": target, "node": node_id,
                    "points": [list(segment[0]), list(segment[1])],
                })
                connector_diamond += int(nodes[node_id].get("type") == "decision")
                text_box = Box(box.x + min(0.16, box.w * 0.12), box.y + min(0.12, box.h * 0.12),
                               max(0.01, box.w - min(0.32, box.w * 0.24)),
                               max(0.01, box.h - min(0.24, box.h * 0.24)))
                connector_text += _segment_hits_box(segment, text_box, 0.0)
    routes = defaultdict(list)
    for segment in segments:
        routes[(segment[2], segment[3])].append(segment)
    arrowheads_off_slide = 0
    diamond_attachment_errors = 0
    diamond_interior_attachment_segments = 0
    arrowhead_diamond_interior = 0
    connector_shape_reentry = 0
    connector_shape_reentry_details = []
    for (source, target), route in routes.items():
        if not route:
            continue
        intended_nodes = {source, target}
        if isinstance(source, str) and source.startswith("continuation:"):
            intended_nodes.add(source.split(":", 2)[1])
        source_node, target_node = nodes.get(source, {}), nodes.get(target, {})
        if source_node.get("type") == "decision" and source in geometry:
            diamond = geometry[source]
            port, escape = route[0][0], route[0][1]
            diamond_attachment_errors += int(not _point_on_diamond_boundary(port, diamond))
            diamond_interior_attachment_segments += int(_point_inside_diamond(escape, diamond, 0.005))
            midpoint = ((port[0] + escape[0]) / 2, (port[1] + escape[1]) / 2)
            diamond_interior_attachment_segments += int(_point_inside_diamond(midpoint, diamond, 0.005))
        if target_node.get("type") == "decision" and target in geometry:
            diamond = geometry[target]
            approach, endpoint = route[-1][0], route[-1][1]
            diamond_attachment_errors += int(not _point_on_diamond_boundary(endpoint, diamond))
            arrowhead_diamond_interior += int(_point_inside_diamond(approach, diamond, 0.005))
            midpoint = ((approach[0] + endpoint[0]) / 2, (approach[1] + endpoint[1]) / 2)
            diamond_interior_attachment_segments += int(_point_inside_diamond(midpoint, diamond, 0.005))
        endpoint = route[-1][1]
        arrowheads_off_slide += int(not _inside_slide_point(endpoint))
        arrow_box = Box(endpoint[0] - 0.08, endpoint[1] - 0.08, 0.16, 0.16)
        arrowhead_overlap += sum(_boxes_overlap(arrow_box, box, 0.01)
                                 for node_id, box in node_items if node_id not in intended_nodes)
        # Source and destination are exempt only at their attachment segment.
        # The former all-segment exemption let a routed path leave a decision
        # correctly and later cross back through the same diamond undetected.
        route_source = (source.split(":", 2)[1]
                        if isinstance(source, str) and source.startswith("continuation:")
                        else source)
        route_target = (target.split(":", 2)[1]
                        if isinstance(target, str) and target.startswith("continuation:")
                        else target)
        if (len(nodes) < 45 and isinstance(source, str) and
                source.startswith("continuation:")):
            # Preserve the previously verified continuation composition for
            # medium workflows. The stricter continuation re-entry audit is a
            # large-graph recovery safeguard; local diamond attachments still
            # undergo the saved-PowerPoint verification below.
            continue
        chains: list[list[tuple]] = []
        for segment in route:
            if (not chains or
                    _segment_length(chains[-1][-1][1], segment[0]) > 0.015):
                chains.append([segment])
            else:
                chains[-1].append(segment)
        for chain in chains:
            for segment_index, segment in enumerate(chain):
                for node_id in {route_source, route_target}:
                    if node_id not in geometry:
                        continue
                    allowed = ((node_id == route_source and segment_index == 0) or
                               (node_id == route_target and segment_index == len(chain) - 1))
                    if allowed:
                        continue
                    box = geometry[node_id]
                    inner = Box(box.x + 0.01, box.y + 0.01,
                                max(0.01, box.w - 0.02),
                                max(0.01, box.h - 0.02))
                    hit = (_segment_hits_diamond(segment, inner, -0.005)
                           if nodes.get(node_id, {}).get("type") == "decision"
                           else _segment_hits_box(segment, inner, 0.0))
                    if hit:
                        connector_shape_reentry += 1
                        connector_shape_reentry_details.append({
                            "source": source, "target": target, "node": node_id,
                            "segment": segment_index,
                            "points": [list(segment[0]), list(segment[1])],
                            "box": [round(box.x, 3), round(box.y, 3),
                                    round(box.w, 3), round(box.h, 3)],
                        })
    def visible_label_box(box: Box) -> Box:
        # PowerPoint text boxes are intentionally a little larger than their
        # glyphs to avoid clipping. Validate the visible text region rather
        # than treating transparent internal padding as painted content.
        inset_x = min(0.035, box.w * 0.08)
        inset_y = min(0.025, box.h * 0.10)
        return Box(box.x + inset_x, box.y + inset_y,
                   max(0.01, box.w - inset_x * 2),
                   max(0.01, box.h - inset_y * 2))

    label_boxes = [visible_label_box(record[0]) for record in label_records]
    label_connector_overlaps = 0
    for record in label_records:
        if len(record) < 3:
            continue
        label, source, target = visible_label_box(record[0]), record[1], record[2]
        label_connector_overlaps += sum(
            _segment_hits_box(segment, label, 0.015)
            for segment in routes.get((source, target), [])
        )
    label_collision_details = []
    for index, first in enumerate(label_records):
        for second in label_records[index + 1:]:
            if _boxes_overlap(visible_label_box(first[0]),
                              visible_label_box(second[0]), 0.02):
                label_collision_details.append({
                    "type": "label_label", "first": [first[1], first[2]],
                    "second": [second[1], second[2]],
                })
    for record in label_records:
        label, source, target = visible_label_box(record[0]), record[1], record[2]
        for node_id, box in node_items:
            overlaps = (_box_overlaps_diamond(label, box, 0.02)
                        if nodes[node_id].get("type") == "decision"
                        else _boxes_overlap(label, box, 0.02))
            if overlaps:
                label_collision_details.append({
                    "type": "label_node", "edge": [source, target], "node": node_id,
                })
        for segment in segments:
            if ((segment[2], segment[3]) != (source, target) and
                    _segment_hits_box(segment, label, 0.01)):
                label_collision_details.append({
                    "type": "label_connector", "edge": [source, target],
                    "connector": [segment[2], segment[3]],
                })
        for protected_index, protected in enumerate(protected_boxes):
            if _boxes_overlap(label, protected, 0.02):
                label_collision_details.append({
                    "type": "label_protected_text", "edge": [source, target],
                    "protected_index": protected_index,
                })
    label_collisions = len(label_collision_details)
    association_errors = 0; association_details = []
    for record in label_records:
        if len(record) < 4:
            # Continuation references do not have a local connector segment.
            continue
        label, source, target = visible_label_box(record[0]), record[1], record[2]
        own = routes.get((source, target), [])
        own_distance = _box_route_distance(label, own)
        other_distances = [_box_route_distance(label, route) for edge, route in routes.items()
                           if edge != (source, target)]
        decision_adjacent = bool(
            nodes.get(source, {}).get("type") == "decision" and source in geometry and
            _box_separation(label, geometry[source]) <= 0.45
        )
        target_adjacent = bool(
            nodes.get(source, {}).get("type") == "decision" and target in geometry and
            _box_separation(label, geometry[target]) <= 0.45 and own_distance <= 1.60
        )
        short_endpoint_adjacent = bool(
            own and sum(_segment_length(segment[0], segment[1]) for segment in own) <= 0.30 and
            min((_box_separation(label, geometry[node_id]) for node_id in (source, target)
                 if node_id in geometry), default=99.0) <= 0.40
        )
        association_error = bool(not own or (not decision_adjacent and not target_adjacent and
                                             not short_endpoint_adjacent and (
            own_distance > 1.35 or
            (other_distances and min(other_distances) + 0.26 < own_distance))))
        # This metric is specifically a decision-branch semantic check. Other
        # labelled transitions still participate in collision and connector
        # proximity validation, but should not be reported as ambiguous
        # decision outcomes.
        decision_error = bool(association_error and
                              nodes.get(source, {}).get("type") == "decision")
        association_errors += int(decision_error)
        if decision_error:
            association_details.append({
                "source": source, "target": target,
                "own_distance": round(own_distance, 3),
                "nearest_other_distance": round(min(other_distances), 3) if other_distances else None,
            })
    crossings = 0; parallel = 0
    for index, first in enumerate(segments):
        for second in segments[index + 1:]:
            if {first[2], first[3]} & {second[2], second[3]}:
                continue
            relation = _segment_relation(first, second)
            crossings += relation == "cross"
            parallel += relation == "overlap"
    overflow = 0; terminal_overflow = 0
    for node_id in page_nodes:
        box = geometry[node_id]; text = nodes[node_id].get("label", "") + nodes[node_id].get("description", "")
        # Approximate PowerPoint's 8–9 pt wrapping area, including the shape's
        # internal margins. This is intentionally conservative but permits
        # normal two- and three-line labels.
        capacity = max(8, int(box.w * 14)) * max(1, int(box.h * 6.5))
        has_overflow = len(text) > capacity
        overflow += has_overflow
        if nodes[node_id].get("type") in {"end", "error"}:
            safe_capacity = max(8, int(box.w * 10.5)) * max(1, int(box.h * 5.2))
            terminal_overflow += len(nodes[node_id].get("label", "")) > safe_capacity
    off_slide = sum(not _inside_slide(box) for _, box in node_items) + sum(not _inside_slide(box) for box in label_boxes)
    paths = defaultdict(float)
    for segment in segments:
        paths[(segment[2], segment[3])] += _segment_length(segment[0], segment[1])
    long_paths = sum(length > 8.0 for length in paths.values())
    detours = 0
    for (source, target), length in paths.items():
        if source not in geometry or target not in geometry:
            continue
        direct = max(0.01, abs(geometry[target].cx - geometry[source].cx) +
                     abs(geometry[target].cy - geometry[source].cy))
        detours += length > direct * 2.35 + 0.75
    ambiguous = 0; ambiguous_details = []
    route_items = list(routes.items())
    for index, ((source_a, target_a), first_route) in enumerate(route_items):
        for (source_b, target_b), second_route in route_items[index + 1:]:
            if source_a != source_b or target_a == target_b:
                continue
            # A tiny shared source stub is acceptable.  Longer overlap makes
            # destinations visually ambiguous even though the edges share a node.
            for first in first_route:
                for second in second_route:
                    if _segment_relation(first, second) == "overlap":
                        overlap = _segment_overlap_length(first, second)
                        if overlap > 0.18:
                            ambiguous += 1
                            ambiguous_details.append({
                                "source": source_a, "first_target": target_a,
                                "second_target": target_b,
                                "overlap": round(overlap, 3),
                            })
    small_shapes = sum(box.w < 0.70 or box.h < 0.32 for _, box in node_items)
    if node_items:
        extent_w = max(box.right for _, box in node_items) - min(box.x for _, box in node_items)
        extent_h = max(box.bottom for _, box in node_items) - min(box.y for _, box in node_items)
        excessive_empty = int(extent_w < 5.2 and extent_h < 1.2 and len(node_items) > 5)
    else:
        excessive_empty = 0
    return {"node_overlaps": node_overlaps,
            "node_overlap_details": node_overlap_details,
            "connector_node_intersections": connector_node,
            "connector_node_intersection_details": connector_node_details,
            "connector_diamond_intersections": connector_diamond,
            "connector_text_intersections": connector_text,
            "diamond_attachment_errors": diamond_attachment_errors,
            "diamond_interior_attachment_segments": diamond_interior_attachment_segments,
            "arrowhead_diamond_interior": arrowhead_diamond_interior,
            "connector_shape_reentry_intersections": connector_shape_reentry,
            "connector_shape_reentry_details": connector_shape_reentry_details,
            "connector_segments_off_slide": connector_off_slide,
            "arrowhead_shape_overlaps": arrowhead_overlap,
            "arrowheads_off_slide": arrowheads_off_slide,
            "label_collisions": label_collisions,
            "label_collision_details": label_collision_details,
            "label_connector_overlaps": label_connector_overlaps,
            "branch_label_association_errors": association_errors,
            "branch_label_association_details": association_details,
            "unnecessary_connector_detours": detours,
            "ambiguous_destination_overlaps": ambiguous,
            "ambiguous_destination_overlap_details": ambiguous_details,
            "connector_crossings": crossings,
            "parallel_segment_overlaps": parallel,
            "text_overflow_warnings": overflow,
            "terminal_text_overflow": terminal_overflow,
            "off_slide_elements": off_slide,
            "excessive_empty_space": excessive_empty,
            "very_long_connector_paths": long_paths,
            "unreadably_small_shapes": small_shapes}


def _segment_overlap_length(first, second):
    (a1, a2), (b1, b2) = first[:2], second[:2]
    if abs(a1[1] - a2[1]) < 0.01 and abs(b1[1] - b2[1]) < 0.01:
        return max(0.0, min(max(a1[0], a2[0]), max(b1[0], b2[0])) -
                   max(min(a1[0], a2[0]), min(b1[0], b2[0])))
    if abs(a1[0] - a2[0]) < 0.01 and abs(b1[0] - b2[0]) < 0.01:
        return max(0.0, min(max(a1[1], a2[1]), max(b1[1], b2[1])) -
                   max(min(a1[1], a2[1]), min(b1[1], b2[1])))
    return 0.0


def _segment_relation(first, second):
    (a1, a2), (b1, b2) = first[:2], second[:2]
    ah, bh = abs(a1[1] - a2[1]) < 0.01, abs(b1[1] - b2[1]) < 0.01
    if ah and bh and abs(a1[1] - b1[1]) < 0.035:
        if max(min(a1[0], a2[0]), min(b1[0], b2[0])) < min(max(a1[0], a2[0]), max(b1[0], b2[0])) - 0.03:
            return "overlap"
    if not ah and not bh and abs(a1[0] - b1[0]) < 0.035:
        if max(min(a1[1], a2[1]), min(b1[1], b2[1])) < min(max(a1[1], a2[1]), max(b1[1], b2[1])) - 0.03:
            return "overlap"
    horizontal, vertical = (first, second) if ah and not bh else ((second, first) if bh and not ah else (None, None))
    if horizontal:
        h1, h2 = horizontal[:2]; v1, v2 = vertical[:2]
        if min(h1[0], h2[0]) + 0.02 < v1[0] < max(h1[0], h2[0]) - 0.02 and \
                min(v1[1], v2[1]) + 0.02 < h1[1] < max(v1[1], v2[1]) - 0.02:
            return "cross"
    return "none"


def _continuation_route_metrics(points, existing_segments):
    """Return crossings, overlaps, length, and bends for a compact continuation route."""
    route_segments = list(zip(points, points[1:]))
    crossings = 0; overlaps = 0
    for segment in route_segments:
        candidate = (segment[0], segment[1], "continuation", "continuation")
        for prior in existing_segments:
            relation = _segment_relation(candidate, prior)
            crossings += relation == "cross"
            overlaps += relation == "overlap"
    length = sum(_segment_length(start, end) for start, end in route_segments)
    bends = max(0, len(points) - 2)
    return crossings, overlaps, length, bends


def _set_arrowhead(connector):
    ln = connector.line._get_or_add_ln(); tail = ln.find("{http://schemas.openxmlformats.org/drawingml/2006/main}tailEnd")
    if tail is None: tail = OxmlElement("a:tailEnd"); ln.append(tail)
    tail.set("type", "triangle"); tail.set("w", "sm"); tail.set("len", "sm")


def _render_continuations(slide, page, edges, lr, nodes, geometry, label_boxes,
                          connector_segments, theme, continuation_style, rendering_mode="compact",
                          style=None, protected_boxes=None):
    style = style or {}
    protected_boxes = protected_boxes or []
    all_crossing = sorted([e for e in edges if e["from"] in lr.placements and e["to"] in lr.placements and
                           lr.placements[e["from"]].page != lr.placements[e["to"]].page],
                          key=lambda e: (e["from"], e["to"], e.get("label", "")))
    groups: dict[tuple[int, int], list[dict[str, Any]]] = defaultdict(list)
    for edge in all_crossing:
        groups[(lr.placements[edge["from"]].page, lr.placements[edge["to"]].page)].append(edge)
    incoming_by_target: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for edge in edges:
        if edge["from"] in lr.placements and edge["to"] in lr.placements:
            incoming_by_target[edge["to"]].append(edge)
    visible = []
    for boundary, group in sorted(groups.items()):
        if page not in boundary:
            continue
        for index, edge in enumerate(group):
            source_page, target_page = boundary
            if source_page == page:
                # Every crossing edge receives a connected source-side
                # indicator. This is the minimum human-readable semantic
                # representation and is never suppressed in compact mode.
                visible.append((edge, boundary, index, len(group), True))
            elif target_page == page:
                target_incoming = incoming_by_target[edge["to"]]
                local_incoming = [item for item in target_incoming
                                  if lr.placements[item["from"]].page == target_page]
                # Add a matching incoming reference when the target would
                # otherwise have no visible predecessor, or when it is an
                # explicit convergence point with several required inputs.
                required_inputs = [item for item in target_incoming
                                   if item.get("path", "main") != "recovery"]
                # Several mutually exclusive outcomes may converge on one
                # activity; repeating an incoming marker for each of those
                # alternatives adds clutter without clarifying continuity.
                # A matching destination-side reference is reserved for a
                # true synchronization across separate process areas.
                input_areas = {
                    str(nodes[item["from"]].get("process_area", "")).strip()
                    for item in required_inputs
                    if str(nodes[item["from"]].get("process_area", "")).strip()
                }
                required_join = (len(required_inputs) > 1 and
                                 len(input_areas) > 1 and
                                 all(item.get("path", "main") == "main"
                                     for item in required_inputs))
                backward_reference = target_page < source_page
                if not local_incoming or required_join or backward_reference:
                    visible.append((edge, boundary, index, len(group), False))

    continuation_boxes: list[Box] = []
    for visible_index, (edge, (source_page, target_page), index, group_size, outgoing) in enumerate(visible):
        node_id = edge["from"] if outgoing else edge["to"]
        anchor = geometry.get(node_id)
        if anchor is None:
            continue
        other_page = target_page + 1 if outgoing else source_page + 1
        source_label = str(nodes[edge["from"]].get("label", "Previous step")).rstrip("?")
        target_label = str(nodes[edge["to"]].get("label", "Next step")).rstrip("?")
        semantic = str(edge.get("label", "")).strip()
        backward = target_page < source_page
        if outgoing:
            action = "Return to" if backward or edge.get("path") == "recovery" else "To"
            label = f"{semantic + ': ' if semantic else ''}{action} {target_label} · slide {other_page}"
        else:
            action = "Retry from" if backward or edge.get("path") == "recovery" else "From"
            label = f"{semantic + ': ' if semantic else ''}{action} {source_label} · slide {other_page}"
        width = min(1.82, max(1.15, 0.28 + min(len(label), 24) * 0.055))
        lines = max(1, math.ceil(len(label) / max(18, int(width * 13))))
        height = min(0.50, 0.17 + lines * 0.13)
        if lr.continuation_space_reserve > 0:
            # Recovery candidates reserve a dedicated continuation rail.  Use
            # it consistently instead of floating prose around process nodes.
            # The off-page connector remains editable, preserves the branch
            # condition and destination, and cannot be confused with an
            # activity rectangle or decision diamond.
            width = min(1.42, max(1.18, width))
            height = min(0.60, max(0.44, height))
            rail_x = 13.00 - width
            rail_candidates = []
            y = 1.02
            while y + height <= 7.16:
                box = Box(rail_x, y, width, height)
                if not any(_boxes_overlap(box, other, 0.05) for other in continuation_boxes):
                    rail_candidates.append(box)
                y += 0.12
            rail_candidates.sort(key=lambda box: (abs(box.cy - anchor.cy), box.y))
            candidates = rail_candidates + _continuation_candidates(
                anchor, width, height, outgoing, index)
        else:
            candidates = _continuation_candidates(anchor, width, height, outgoing, index)
        if outgoing and nodes[node_id].get("type") == "decision":
            candidates = [
                Box(anchor.cx - width / 2, anchor.bottom + 0.12 + index * 0.08, width, height),
                Box(anchor.right + 0.12, anchor.cy - height / 2 + index * 0.08, width, height),
            ] + candidates
        if outgoing and edge.get("path") == "recovery" and nodes[node_id].get("type") == "decision":
            candidates = [Box(anchor.cx - width / 2, anchor.bottom + 0.10 + index * 0.08, width, height),
                          Box(anchor.right + 0.12, anchor.bottom + 0.04 + index * 0.08, width, height)] + candidates
        candidates = [_clamp_continuation_box(box) for box in candidates]
        chosen = None; points = None
        unrelated_nodes = [box for other_id, box in geometry.items() if other_id != node_id]
        def candidate_route(candidate):
            if (not _inside_slide(candidate) or
                    any(_boxes_overlap(candidate, other, 0.04) for other in geometry.values()) or
                    any(_boxes_overlap(candidate, other, 0.015) for other in label_boxes) or
                    any(_boxes_overlap(candidate, other, 0.025) for other in protected_boxes)):
                return None
            candidate_points = _continuation_link_points(
                anchor, candidate, nodes[node_id].get("type") == "decision", outgoing,
                use_rail=lr.continuation_space_reserve > 0)
            candidate_segments = list(zip(candidate_points, candidate_points[1:]))
            node_hit = any(_segment_hits_box(segment, box, 0.025)
                           for segment in candidate_segments for box in unrelated_nodes)
            label_hit = any(_segment_hits_box(segment, box, 0.005)
                            for segment in candidate_segments for box in label_boxes)
            protected_hit = any(_segment_hits_box(segment, box, 0.015)
                                for segment in candidate_segments for box in protected_boxes)
            if node_hit or label_hit or protected_hit:
                # A direct L route may be blocked on a dense slide even when
                # the marker itself has a valid location. Reuse the canonical
                # obstacle-aware router before rejecting that location.
                link_obstacles = unrelated_nodes + label_boxes + protected_boxes
                routed = _orthogonal_grid_path(
                    candidate_points[0], candidate_points[-1], link_obstacles,
                    [(start, end) for start, end, *_ in connector_segments],
                )
                if not routed:
                    return None
                candidate_points = _simplify_route(routed)
                candidate_segments = list(zip(candidate_points, candidate_points[1:]))
                anchor_inner = Box(anchor.x + 0.01, anchor.y + 0.01,
                                   max(0.01, anchor.w - 0.02),
                                   max(0.01, anchor.h - 0.02))
                if len(nodes) >= 45 and any(
                        (_segment_hits_diamond(segment, anchor_inner, -0.005)
                         if nodes[node_id].get("type") == "decision" else
                         _segment_hits_box(segment, anchor_inner, 0.0))
                        for segment in candidate_segments[1:]):
                    return None
                if (any(_segment_hits_box(segment, box, 0.025)
                        for segment in candidate_segments for box in unrelated_nodes) or
                        any(_segment_hits_box(segment, box, 0.005)
                            for segment in candidate_segments for box in label_boxes) or
                        any(_segment_hits_box(segment, box, 0.015)
                            for segment in candidate_segments for box in protected_boxes)):
                    return None
            return candidate_points

        valid_options = []
        primary_limit = 72 if len(nodes) < 45 else 16
        for candidate_order, candidate in enumerate(candidates[:primary_limit]):
            _ROUTING_STATE["continuation_candidates"] += 1
            _routing_budget_check(
                f"continuation {edge['from']} -> {edge['to']} on slide {page + 1}")
            candidate_points = candidate_route(candidate)
            if candidate_points:
                metrics = _continuation_route_metrics(candidate_points, connector_segments)
                valid_options.append((candidate_order, candidate, candidate_points, metrics))
        if valid_options:
            first_order, first_candidate, first_points, first_metrics = valid_options[0]
            first_length = first_metrics[2]
            bounded = [option for option in valid_options
                       if option[3][2] <= first_length * 1.28 + 0.45 and
                       option[3][3] <= first_metrics[3] + 2]
            selected = min(
                bounded or [valid_options[0]],
                key=lambda option: (option[3][0], option[3][1], option[3][2], option[0]),
            )
            _, chosen, points, _ = selected
        if chosen is None:
            # A dense slide can have no free wide label slot even though a
            # narrow multi-line reference remains readable. This bounded
            # fallback changes only the continuation callout, never workflow
            # nodes or their typography.
            fallback_width, fallback_height = 0.94, 0.56
            fallback_candidates = [
                _clamp_continuation_box(box) for box in
                _continuation_candidates(anchor, fallback_width, fallback_height, outgoing, index)
            ]
            fallback_options = []
            _ROUTING_STATE["fallback_uses"] += 1
            fallback_limit = 64 if len(nodes) < 45 else 12
            for candidate_order, candidate in enumerate(fallback_candidates[:fallback_limit]):
                _ROUTING_STATE["continuation_candidates"] += 1
                _routing_budget_check(
                    f"fallback continuation {edge['from']} -> {edge['to']} on slide {page + 1}")
                candidate_points = candidate_route(candidate)
                if candidate_points:
                    fallback_options.append((candidate_order, candidate, candidate_points,
                                             _continuation_route_metrics(candidate_points,
                                                                         connector_segments)))
            if fallback_options:
                _, chosen, points, _ = min(
                    fallback_options,
                    key=lambda option: (option[3][0], option[3][1], option[3][2], option[0]),
                )
        if chosen is None:
            raise ValueError(
                f"No collision-free continuation position for {edge['from']} -> {edge['to']} "
                f"on slide {page + 1}."
            )
        # Use the same transparent editable reference in every layout path.
        # Outlined off-page symbols consumed extra clearance and changed slide
        # counts depending on whether adaptive recovery happened to run.
        marker_kind = None
        tag = (slide.shapes.add_shape(marker_kind, Inches(chosen.x), Inches(chosen.y),
                                      Inches(chosen.w), Inches(chosen.h))
               if marker_kind is not None else
               slide.shapes.add_textbox(Inches(chosen.x), Inches(chosen.y),
                                        Inches(chosen.w), Inches(chosen.h)))
        edge_index = int(edge.get("_render_index", edges.index(edge)))
        role = "outgoing" if outgoing else "incoming"
        tag.name = f"workflow-continuation::{edge_index}::{role}"
        if marker_kind is not None:
            tag.fill.solid(); tag.fill.fore_color.rgb = _rgb("FFFFFF")
            tag.line.color.rgb = _rgb(style.get("colors", {}).get("connector", theme.line))
            tag.line.width = Pt(0.8)
        else:
            tag.fill.background(); tag.line.fill.background()
        _set_text(
            tag, label, theme, 7, False, style.get("colors", {}).get("connector_label", theme.muted),
            style, "branch_label")
        tag.text_frame.margin_left = tag.text_frame.margin_right = Inches(0.01)
        tag.text_frame.margin_top = tag.text_frame.margin_bottom = 0
        color, width_pt, line_style, _ = _edge_visual(edge, theme, style,
                                                       edge.get("path") in {"exception", "recovery"}, 1.0)
        directed = points if outgoing else list(reversed(points))
        for segment_index, (p1, p2) in enumerate(zip(directed, directed[1:]), start=1):
            line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(p1[0]), Inches(p1[1]),
                                              Inches(p2[0]), Inches(p2[1]))
            line.name = f"workflow-continuation-link::{edge_index}::{role}::segment-{segment_index}"
            line.line.color.rgb = _rgb(color); line.line.width = Pt(max(0.9, min(1.2, width_pt)))
            _set_dash(line, line_style)
            if segment_index == len(directed) - 1:
                _set_arrowhead(line)
            continuation_token = f"continuation:{node_id}:{edge_index}:{role}"
            connector_segments.append((p1, p2, continuation_token, continuation_token))
        label_boxes.append(chosen)
        continuation_boxes.append(chosen)
    return len(visible)


def _continuation_candidates(anchor: Box, width: float, height: float, outgoing: bool, index: int):
    offset = index * 0.08
    preferred = [
        Box(anchor.right + 0.10, anchor.y - height - 0.08 - offset, width, height),
        Box(anchor.right + 0.10, anchor.bottom + 0.08 + offset, width, height),
        Box(anchor.cx - width / 2, anchor.y - height - 0.12 - offset, width, height),
        Box(anchor.cx - width / 2, anchor.bottom + 0.12 + offset, width, height),
        Box(anchor.x - width - 0.10, anchor.y - height - 0.08 - offset, width, height),
        Box(anchor.x - width - 0.10, anchor.bottom + 0.08 + offset, width, height),
    ]
    # When the preferred side is at the slide boundary, use a second compact
    # ring around the anchor rather than overlapping a workflow node.
    preferred.extend([
        Box(anchor.right + 0.10, anchor.cy - height / 2 + 0.14 + offset, width, height),
        Box(anchor.x - width - 0.10, anchor.cy - height / 2 - 0.14 - offset, width, height),
        Box(anchor.cx - width / 2 + 0.20, anchor.y - height - 0.18 - offset, width, height),
        Box(anchor.cx - width / 2 - 0.20, anchor.bottom + 0.18 + offset, width, height),
    ])
    for distance in (0.45, 0.85, 1.25, 1.65, 2.05):
        preferred.extend([
            Box(anchor.x - width - distance, anchor.y - height - 0.08 - offset, width, height),
            Box(anchor.x - width - distance, anchor.bottom + 0.08 + offset, width, height),
            Box(anchor.right + distance, anchor.y - height - 0.08 - offset, width, height),
            Box(anchor.right + distance, anchor.bottom + 0.08 + offset, width, height),
        ])
    grid = []
    y = 1.02
    while y + height <= 7.10:
        x = 0.34
        while x + width <= 12.96:
            grid.append(Box(x, y, width, height)); x += 0.48
        y += 0.36
    grid.sort(key=lambda box: (abs(box.cx - anchor.cx) + abs(box.cy - anchor.cy), box.y, box.x))
    preferred.extend(grid)
    return preferred if outgoing else list(reversed(preferred))


def _clamp_continuation_box(box: Box) -> Box:
    x = min(max(0.30, box.x), 13.00 - box.w)
    y = min(max(0.94, box.y), 7.16 - box.h)
    return Box(x, y, box.w, box.h)


def _continuation_link_points(anchor: Box, marker: Box, anchor_is_decision: bool,
                              outgoing: bool, use_rail: bool = False) -> list[tuple[float, float]]:
    """Return a short orthogonal link between a workflow node and an indicator."""
    dx, dy = marker.cx - anchor.cx, marker.cy - anchor.cy
    if use_rail and marker.x >= anchor.right + 0.30:
        anchor_point = (_diamond_boundary_port(anchor, "right") if anchor_is_decision
                        else (anchor.right, anchor.cy))
        marker_point = (marker.x, marker.cy)
        corridor_x = marker.x - 0.08
        return _simplify_route([
            anchor_point, (corridor_x, anchor_point[1]),
            (corridor_x, marker_point[1]), marker_point,
        ])
    if abs(dx) >= abs(dy):
        side = "right" if dx >= 0 else "left"
        anchor_point = (_diamond_boundary_port(anchor, side) if anchor_is_decision else
                        ((anchor.right, anchor.cy) if side == "right" else (anchor.x, anchor.cy)))
        marker_point = (marker.x, marker.cy) if side == "right" else (marker.right, marker.cy)
        elbow = (marker_point[0], anchor_point[1])
    else:
        side = "bottom" if dy >= 0 else "top"
        anchor_point = (_diamond_boundary_port(anchor, side) if anchor_is_decision else
                        ((anchor.cx, anchor.bottom) if side == "bottom" else (anchor.cx, anchor.y)))
        marker_point = (marker.cx, marker.y) if side == "bottom" else (marker.cx, marker.bottom)
        elbow = (anchor_point[0], marker_point[1])
    points = _simplify_route([anchor_point, elbow, marker_point])
    return points if outgoing else points


def _exception_view(workflow):
    special = [e for e in workflow["edges"] if e.get("path") in {"exception", "recovery"}]
    ids = {value for edge in special for value in (edge["from"], edge["to"])}
    terminal_ids = {n["id"] for n in workflow["nodes"] if n["type"] in {"end", "error"}}
    terminal_edges = [e for e in workflow["edges"] if e["from"] in ids and e["to"] in terminal_ids]
    rendered_edges = special + [e for e in terminal_edges if e not in special]
    ids.update(e["to"] for e in terminal_edges); view = dict(workflow)
    view["nodes"] = [n for n in workflow["nodes"] if n["id"] in ids]; view["edges"] = rendered_edges
    return view


def _render_exception_summary(slide, edges, nodes, theme):
    y = 1.15
    for edge in edges[:10]:
        src, dst = nodes[edge["from"]], nodes[edge["to"]]
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(y), Inches(11.9), Inches(0.48))
        box.fill.solid(); box.fill.fore_color.rgb = _rgb(theme.exception_soft); box.line.color.rgb = _rgb(theme.exception)
        _set_text(box, f"{src['label']}  →  {edge.get('label', edge.get('path', 'Exception'))}  →  {dst['label']}", theme, 11); y += 0.58


def _add_notes_slides(prs, workflow, blank, theme, include_assumptions=False, include_open_questions=False):
    items = [("Assumptions", workflow.get("assumptions", []) if include_assumptions else []),
             ("Open questions", workflow.get("open_questions", []) if include_open_questions else [])]
    if not any(v for _, v in items): return
    slide = prs.slides.add_slide(blank); _add_title(slide, "Assumptions and open questions", "Review before using this illustrative workflow as policy", theme)
    y = 1.2
    for heading, values in items:
        if not values: continue
        head = slide.shapes.add_textbox(Inches(0.7), Inches(y), Inches(11.8), Inches(0.35))
        _set_text(head, heading, theme, 16, True); head.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT; y += 0.42
        for value in values:
            box = slide.shapes.add_textbox(Inches(0.9), Inches(y), Inches(11.3), Inches(0.4))
            _set_text(box, "• " + str(value), theme, 12); box.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT; y += 0.42
        y += 0.15

# ---- Consolidated from renderer/generate.py ----
import json
from copy import deepcopy
from pathlib import Path
from typing import Any



def _renderer_generate_workflow(workflow: dict[str, Any] | str | Path, output_path: str | Path,
                      theme: str = "minimal", layout: str = "flowchart",
                      include_exceptions: str = "none", continuation_style: str = "minimal",
                      rendering_mode: str | None = None, include_assumptions: bool = False,
                      include_open_questions: bool = False, exception_ids: list[str] | tuple[str, ...] | None = None,
                      density: str = "compact", include_overview: bool = False,
                      style: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate, lay out, and render a workflow as editable native PowerPoint shapes."""
    render_started = time.monotonic()
    _ROUTING_STATE.update({
        "deadline": render_started + 110.0,
        "grid_expansions": 0,
        "grid_budget_exhaustions": 0,
        "continuation_candidates": 0,
        "fallback_uses": 0,
    })
    try:
        if isinstance(workflow, (str, Path)):
            with Path(workflow).open("r", encoding="utf-8") as handle:
                workflow = json.load(handle)
        report = validate_workflow(workflow)
        source_workflow = workflow
        if layout == "standard":
            layout = "flowchart"
        exception_mode = _normalize_exception_mode(include_exceptions)
        effective_density = rendering_mode or density
        style_config = validate_style(style, source_workflow)
        presentation_workflow, exception_metadata = _select_exceptions(
            source_workflow, exception_ids, preserve_all=exception_mode == "inline" and exception_ids is None,
            preserve_standalone_recovery=True)
        effective_theme = "minimal_process_map" if layout == "process_map" and theme == "minimal" else theme
        theme_obj = get_theme(effective_theme)
        preferred_slides = style_config.get("layout_preferences", {}).get("preferred_slide_count")
        full_layout = calculate_layout(source_workflow, layout, rendering_mode=effective_density,
                                       preferred_slide_count=preferred_slides)
        layout_workflow = (_main_view(presentation_workflow)
                           if exception_mode in {"separate", "omit", "exception_flowchart"}
                           else presentation_workflow)
        layout_result = calculate_layout(layout_workflow, layout, rendering_mode=effective_density,
                                         preferred_slide_count=preferred_slides)
        destination = Path(output_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        if theme == "monochrome":
            monochrome_colors = {
                "background": "FFFFFF", "title": "111111", "subtitle": "444444", "text": "111111",
                "activity_fill": "F7F7F7", "activity_border": "333333",
                "decision_fill": "E2E2E2", "decision_border": "222222",
                "connector": "333333", "connector_label": "333333",
                "swimlane_fill": "F2F2F2", "swimlane_alt_fill": "FFFFFF", "swimlane_header": "333333",
                "accent": "222222", "successful": "E5E5E5", "unsuccessful": "C8C8C8",
                "recoverable": "F3F3F3", "unclassified": "DDDDDD",
            }
            monochrome_colors.update(style_config["colors"])
            style_config["colors"] = monochrome_colors
        theme_obj = _apply_style_to_theme(theme_obj, style_config)
        composition_fallback = None
        recovery_attempts: list[dict[str, Any]] = []
        selected_strategy = "initial"

        def recoverable_layout_failure(error: ValueError) -> bool:
            message = str(error)
            return any(token in message for token in (
                "No collision-free continuation position",
                "No readable continuation-label position",
                "No endpoint-safe orthogonal route found",
                "Critical geometry validation failed",
                "Critical terminal connectivity failure",
                "PowerPoint cross-slide semantic verification failed",
            ))

        # Candidate order is the recovery cost order. The first valid result is
        # therefore the least disruptive composition that satisfies all hard
        # geometry and semantic checks. A preferred slide count remains soft.
        candidates: list[tuple[str, LayoutResult]] = [("initial", layout_result)]
        if layout != "process_map":
            candidate_specs: list[tuple[str, dict[str, Any]]] = []
            if preferred_slides:
                for extra in (1, 2, 3):
                    candidate_specs.append((f"soft_preference_plus_{extra}", {
                        "preferred_slide_count": preferred_slides + extra,
                        "allow_adaptive_compression": False,
                        "continuation_space_reserve": 1.58,
                    }))
            candidate_specs.append(("conservative_same_preference", {
                "preferred_slide_count": preferred_slides,
                "allow_adaptive_compression": False,
                "continuation_space_reserve": 1.58,
            }))
            candidate_specs.append(("unconstrained_conservative", {
                "preferred_slide_count": None,
                "allow_adaptive_compression": False,
                "continuation_space_reserve": 1.58,
            }))
            # Smaller rank capacities reserve additional horizontal routing and
            # continuation space without shrinking workflow shapes or text.
            for capacity in (9, 8, 7, 6, 5, 4):
                candidate_specs.append((f"routing_space_capacity_{capacity}", {
                    "ranks_per_page": capacity,
                    "preferred_slide_count": None,
                    "allow_adaptive_compression": False,
                    "continuation_space_reserve": 1.58,
                }))
            seen_signatures = {
                tuple(sorted((node_id, placement.page, placement.column)
                             for node_id, placement in layout_result.placements.items()))
            }
            for strategy, kwargs in candidate_specs:
                candidate = calculate_layout(
                    layout_workflow, layout, rendering_mode=effective_density, **kwargs)
                signature = tuple(sorted((node_id, placement.page, placement.column)
                                         for node_id, placement in candidate.placements.items()))
                if signature in seen_signatures:
                    continue
                seen_signatures.add(signature)
                candidates.append((strategy, candidate))

        # Compare every bounded composition before rendering.  V1.6.4 used
        # recovery cost order and accepted the first geometrically valid
        # layout, which allowed a ten-slide candidate with a one-node opener
        # to beat a more coherent seven-slide plan.  Geometry remains a hard
        # gate below; this score only determines which valid-looking plans are
        # worth attempting first.
        candidate_quality = {
            strategy: _adaptive_layout_candidate_score(candidate, layout_workflow,
                                                        preferred_slides)
            for strategy, candidate in candidates
        }
        candidates.sort(key=lambda item: (
            candidate_quality[item[0]]["total_cost"], item[1].pages, item[0]))

        details = None
        last_error: ValueError | None = None
        for strategy, candidate in candidates[:6]:
            _routing_budget_check(f"layout attempt {strategy}")
            style_variants = [(strategy, style_config)]
            variant_index = 0
            while variant_index < len(style_variants):
                variant_strategy, candidate_style = style_variants[variant_index]
                variant_index += 1
                try:
                    candidate_details = render_presentation(
                        presentation_workflow, candidate, str(destination), theme_obj, layout,
                        exception_mode, continuation_style, effective_density,
                        include_assumptions, include_open_questions, include_overview, candidate_style)
                except ValueError as error:
                    last_error = error
                    recovery_attempts.append({
                        "strategy": variant_strategy,
                        "slides": candidate.pages,
                        "candidate_score": candidate_quality.get(
                            strategy, candidate_quality.get(variant_strategy)),
                        "status": "failed",
                        "diagnostic": str(error),
                    })
                    if not recoverable_layout_failure(error):
                        raise
                    label_only_failure = ("label_collisions" in str(error) and not any(
                        token in str(error) for token in (
                            "node_overlaps=", "connector_node_intersections=",
                            "connector_diamond_intersections=", "connector_text_intersections=",
                            "connector_segments_off_slide=", "arrowhead_shape_overlaps=",
                            "arrowheads_off_slide=", "off_slide_elements=",
                            "text_overflow_warnings=", "terminal_text_overflow=",
                            "ambiguous_destination_overlaps=",
                        )))
                    if (label_only_failure and
                            not variant_strategy.endswith(("labels_wide", "labels_tight",
                                                           "labels_centered", "labels_near_target"))):
                        for clearance in ("wide", "tight"):
                            alternate_style = deepcopy(candidate_style)
                            alternate_style["connector_label_offset"] = clearance
                            alternate_style.setdefault("connectors", {})[
                                "label_clearance"] = clearance
                            style_variants.append(
                                (f"{strategy}_labels_{clearance}", alternate_style))
                        for label_mode in ("centered", "near_target"):
                            alternate_style = deepcopy(candidate_style)
                            alternate_style["branch_label_position"] = label_mode
                            alternate_style.setdefault("connectors", {})[
                                "branch_label_position"] = label_mode
                            style_variants.append(
                                (f"{strategy}_labels_{label_mode}", alternate_style))
                    continue
                details = candidate_details
                layout_result = candidate
                style_config = candidate_style
                selected_strategy = variant_strategy
                recovery_attempts.append({
                    "strategy": variant_strategy,
                    "slides": candidate.pages,
                    "candidate_score": candidate_quality.get(
                        strategy, candidate_quality.get(variant_strategy)),
                    "status": "selected",
                    "diagnostic": None,
                })
                break
            if details is not None:
                break
        if details is None:
            attempted = "; ".join(
                f"{item['strategy']} ({item['slides']} slides): {item['diagnostic']}"
                for item in recovery_attempts
            )
            raise ValueError(
                "Adaptive layout recovery exhausted all deterministic candidates. "
                f"Attempts: {attempted}"
            ) from last_error
        if selected_strategy != "initial":
            composition_fallback = recovery_attempts[0]["diagnostic"]
            details.setdefault("layout_warnings", []).append(
                "The initial composition failed a hard layout constraint; "
                f"adaptive recovery selected {selected_strategy.replace('_', ' ')}."
            )
        details["adaptive_layout_recovery"] = {
            "applied": selected_strategy != "initial",
            "selected_strategy": selected_strategy,
            "attempts": recovery_attempts,
            "preferred_slide_count": preferred_slides,
            "selected_slide_count": layout_result.pages,
        }
        details["performance_diagnostics"] = {
            "total_render_seconds": round(time.monotonic() - render_started, 3),
            "layout_attempts": len(recovery_attempts),
            "route_grid_expansions": _ROUTING_STATE["grid_expansions"],
            "continuation_candidates_evaluated": _ROUTING_STATE["continuation_candidates"],
            "grid_budget_exhaustions": _ROUTING_STATE["grid_budget_exhaustions"],
            "fallback_uses": _ROUTING_STATE["fallback_uses"],
            "execution_budget_seconds": 110,
        }
        details.update({"pages": details["diagram_slides"] if layout == "process_map" else layout_result.pages,
                        "back_edges": sorted(full_layout.back_edges),
                        "decision_groups": len(layout_result.decision_groups),
                        "split_decision_groups": layout_result.split_decision_groups,
                        "unused_space": (details.get("process_map_unused_space")
                                         if layout == "process_map" else layout_result.unused_space),
                        "split_edges": (details.get("process_map_split_edges", 0)
                                        if layout == "process_map" else layout_result.split_edges),
                        "minimum_font_size": 8 if effective_density == "compact" else 9,
                        "connector_crossings": (details["layout_quality"]["connector_crossings"]
                                                if layout == "process_map" else layout_result.connector_crossings),
                        "composition_candidate_fallback": composition_fallback,
                        "exception_mode_requested": include_exceptions,
                        "density": effective_density,
                        "exception_selection": exception_metadata,
                        "omitted_sections": {
                            "assumptions": list(source_workflow.get("assumptions", [])) if not include_assumptions else [],
                            "open_questions": list(source_workflow.get("open_questions", [])) if not include_open_questions else [],
                            "exception_ids": (exception_metadata["available_ids"] if exception_mode == "omit"
                                              else exception_metadata["omitted_ids"]),
                        }, **report.details})
        generation_warnings = list(report.warnings)
        generation_warnings.extend(details.get("layout_warnings", []))
        if layout != "process_map":
            for decision_id in layout_result.split_decision_groups:
                group = next(group for group in layout_result.decision_groups if group.decision_id == decision_id)
                reason = "group exceeds one-slide capacity" if not group.feasible else "overlapping group constraints"
                generation_warnings.append(f"Decision group {decision_id} spans slides: {reason}.")
        for decision_id in details.get("exception_split_decision_groups", []):
            generation_warnings.append(f"Exception decision group {decision_id} spans slides because its retry path exceeds one-slide capacity.")
        return {"status": "success", "output_path": str(destination.resolve()), "details": details,
                "warnings": generation_warnings}
    except (WorkflowValidationError, json.JSONDecodeError, OSError, ValueError) as exc:
        errors = exc.report.errors if isinstance(exc, WorkflowValidationError) else [str(exc)]
        warnings = exc.report.warnings if isinstance(exc, WorkflowValidationError) else []
        return {"status": "error", "output_path": None, "details": {
            "errors": errors,
            "performance_diagnostics": {
                "total_render_seconds": round(time.monotonic() - render_started, 3),
                "route_grid_expansions": _ROUTING_STATE["grid_expansions"],
                "continuation_candidates_evaluated": _ROUTING_STATE["continuation_candidates"],
                "grid_budget_exhaustions": _ROUTING_STATE["grid_budget_exhaustions"],
                "fallback_uses": _ROUTING_STATE["fallback_uses"],
                "execution_budget_seconds": 110,
            }}, "warnings": warnings}


def _normalize_exception_mode(value: str) -> str:
    aliases = {"none": "omit", "omit": "omit", "register": "separate", "separate": "separate",
               "flowchart": "exception_flowchart", "exception_flowchart": "exception_flowchart",
               "inline": "inline"}
    if value not in aliases:
        raise ValueError("include_exceptions must be none, register, flowchart, inline, omit, separate, or exception_flowchart")
    return aliases[value]


def _select_exceptions(workflow: dict[str, Any], exception_ids, preserve_all: bool = False,
                       preserve_standalone_recovery: bool = False):
    """Create a presentation-only view containing selected exception roots and their recovery closure."""
    view = deepcopy(workflow)
    roots = [(edge.get("id") or edge["to"], edge) for edge in workflow["edges"] if edge.get("path") == "exception"]
    available = [str(identifier) for identifier, _ in roots]
    if preserve_all:
        return view, {"available_ids": available, "selected_ids": available, "omitted_ids": []}
    if exception_ids is None:
        selected = set(available)
    else:
        selected = {str(value) for value in exception_ids}
        unknown = selected - set(available)
        if unknown:
            raise ValueError(f"Unknown exception IDs: {', '.join(sorted(unknown))}. Available: {', '.join(available)}")
    selected_roots = [edge for identifier, edge in roots if str(identifier) in selected]
    reachable = {edge["to"] for edge in selected_roots}
    selected_special = list(selected_roots)
    changed = True
    while changed:
        changed = False
        for edge in workflow["edges"]:
            if edge.get("path") == "recovery" and edge["from"] in reachable and edge not in selected_special:
                selected_special.append(edge); reachable.add(edge["to"]); changed = True
    special = {"exception", "recovery"}
    # Recovery edges can also be ordinary correction loops in the main
    # workflow.  Only recovery edges belonging to an exception closure are
    # controlled by exception selection; standalone recovery loops remain.
    exception_closure = {edge["to"] for _, edge in roots}
    changed = True
    while changed:
        changed = False
        for edge in workflow["edges"]:
            if edge.get("path") == "recovery" and edge["from"] in exception_closure and edge["to"] not in exception_closure:
                exception_closure.add(edge["to"]); changed = True
    view["edges"] = [edge for edge in workflow["edges"]
                     if edge.get("path") not in special or edge in selected_special or
                     (preserve_standalone_recovery and edge.get("path") == "recovery" and
                      edge["from"] not in exception_closure)]
    omitted_closure = {edge["to"] for identifier, edge in roots if str(identifier) not in selected}
    changed = True
    while changed:
        changed = False
        for edge in workflow["edges"]:
            if edge.get("path") == "recovery" and edge["from"] in omitted_closure and edge["to"] not in omitted_closure:
                omitted_closure.add(edge["to"]); changed = True
    retained_refs = {value for edge in view["edges"] for value in (edge["from"], edge["to"])}
    view["nodes"] = [node for node in view["nodes"]
                     if node["id"] not in omitted_closure or node["id"] in retained_refs]
    return view, {"available_ids": available, "selected_ids": sorted(selected),
                  "omitted_ids": sorted(set(available) - selected)}


def _main_view(workflow: dict[str, Any]) -> dict[str, Any]:
    """Return a diagram view without nodes used only by exception/recovery paths."""
    view = deepcopy(workflow)
    # Ordinary correction/resubmission loops are part of the business process,
    # not optional exception documentation.  Remove exception-rooted material
    # only; _select_exceptions has already removed recovery closures belonging
    # to omitted exception roots.
    special = {"exception"}
    normal_incoming = {e["to"] for e in workflow["edges"] if e.get("path", "main") not in special}
    candidates = {e["to"] for e in workflow["edges"] if e.get("path") == "exception"}
    changed = True
    while changed:
        changed = False
        for edge in workflow["edges"]:
            if edge.get("path") == "recovery" and edge["from"] in candidates and edge["to"] not in normal_incoming:
                if edge["to"] not in candidates:
                    candidates.add(edge["to"]); changed = True
    keep = {n["id"] for n in workflow["nodes"] if n["id"] not in candidates}
    view["nodes"] = [n for n in workflow["nodes"] if n["id"] in keep]
    view["edges"] = [e for e in workflow["edges"] if e["from"] in keep and e["to"] in keep and e.get("path") not in special]
    return view

# ---- Stable deployment interface ----
SUPPORTED_LAYOUTS = {"auto", "flowchart", "swimlane", "process_map"}
SUPPORTED_DENSITIES = {"compact", "balanced", "detailed"}
SUPPORTED_EXCEPTION_MODES = {"none", "register", "flowchart", "inline"}
SUPPORTED_THEMES = {"minimal", "monochrome", "corporate", "high_contrast"}

def generate_workflow(
    workflow: dict[str, Any] | str | Path,
    output_path: str | Path = "/mnt/data/workflow.pptx",
    theme: str = "minimal",
    layout: str = "auto",
    include_exceptions: str = "none",
    continuation_style: str = "minimal",
    include_assumptions: bool = False,
    include_open_questions: bool = False,
    exception_ids: list[str] | tuple[str, ...] | None = None,
    density: str = "compact",
    include_overview: bool = False,
    style: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate and render a workflow as an editable PowerPoint presentation."""
    if layout not in SUPPORTED_LAYOUTS:
        return _deployment_error(f"Unsupported layout: {layout}.")
    if density not in SUPPORTED_DENSITIES:
        return _deployment_error(f"Unsupported density: {density}.")
    if include_exceptions not in SUPPORTED_EXCEPTION_MODES:
        return _deployment_error(f"Unsupported exception mode: {include_exceptions}.")
    if theme not in SUPPORTED_THEMES:
        return _deployment_error(f"Unsupported deployment theme: {theme}.")
    effective_layout = "process_map" if layout == "auto" else layout
    result = _renderer_generate_workflow(
        workflow=workflow, output_path=output_path, theme=theme, layout=effective_layout,
        include_exceptions=include_exceptions, continuation_style=continuation_style,
        include_assumptions=include_assumptions, include_open_questions=include_open_questions,
        exception_ids=exception_ids, density=density, include_overview=include_overview, style=style,
    )
    if result.get("status") == "success" and not Path(result["output_path"]).is_file():
        return _deployment_error("Renderer reported success but no PowerPoint file exists.")
    result.setdefault("details", {})["requested_layout"] = layout
    result["details"]["effective_layout"] = effective_layout
    result["details"]["builder_version"] = BUILDER_VERSION
    errors = result["details"].get("errors", [])
    result["details"].setdefault(
        "unsupported_settings", [error for error in errors if "unsupported" in error.lower()]
    )
    return result

def _deployment_error(message: str) -> dict[str, Any]:
    return {"status": "error", "output_path": None,
            "details": {"errors": [message], "builder_version": BUILDER_VERSION},
            "warnings": []}

__all__ = ["generate_workflow", "validate_workflow", "build_slide_composition_plan",
           "validate_composition_plan", "SlideCompositionPlan", "SlideCompositionPage",
           "WorkflowValidationError", "BUILDER_VERSION"]

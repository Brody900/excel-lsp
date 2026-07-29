from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pytest
from PIL import Image

pytestmark = pytest.mark.live

ROOT = Path(__file__).parents[2]
EVIDENCE = ROOT / "docs" / "evidence" / "live-excel"


def _json(name: str) -> dict[str, Any]:
    return json.loads((EVIDENCE / name).read_text(encoding="utf-8"))


def test_required_live_protocol_artifacts_and_assertions() -> None:
    screenshots = [
        "01-index-map.png",
        "02-diagnostics.png",
        "03-no-repair.png",
        "04-vba-stamp.png",
        "05-write-refusal.png",
        "06-chart-intact.png",
        "07-trace-a2.png",
        "08-trace-b2.png",
        "09-trace-c2.png",
    ]
    index = (EVIDENCE / "index.md").read_text(encoding="utf-8")
    for name in screenshots:
        path = EVIDENCE / name
        assert path.stat().st_size > 10_000
        assert f"[{name}]({name})" in index

    roundtrip = _json("roundtrip.json")
    assert roundtrip["normalOpen"] == {"F16": True, "L1": True}
    assert roundtrip["repairDialog"] == {"F16": False, "L1": False}
    assert roundtrip["macro"] == {"MacroModelZ1": 42, "name": "Stamp", "ran": True}
    assert roundtrip["recalculation"]["L1"] == {
        "ModelB2_B4": [60, 80, 120],
        "ModelC2_C4": [75, 100, 150],
    }
    assert roundtrip["writeRefusal"]["results"][0]["error"]["code"] == "E_OPEN_IN_EXCEL"
    assert all(not result["stale"] for result in roundtrip["refresh"].values())

    diagnostics = _json("02-l3-diagnostics.json")
    assert diagnostics["counts"]["code"] == {
        "E_ERRVAL": 5,
        "W_INCONSISTENT_FORMULA": 1,
    }
    chart = _json("chart-preservation.json")
    assert chart["normalOpen"] is True
    assert chart["repairDialog"] is False
    assert chart["chartObjects"] == 1
    assert chart["shapes"] == ["Chart 1", "Image 2"]
    assert chart["chartVisible"] is True
    assert chart["imageVisible"] is True
    assert chart["selectedShape"] == "Image 2"
    assert set(chart["shapeGeometry"]) == {"Chart 1", "Image 2"}
    for geometry in chart["shapeGeometry"].values():
        assert geometry["visible"] is True
        assert geometry["width"] > 0
        assert geometry["height"] > 0


def test_lineage_demo_uses_only_numbered_live_frames() -> None:
    manifest = _json("demo-capture.json")
    output = ROOT / manifest["output"]
    assert hashlib.sha256(output.read_bytes()).hexdigest() == manifest["outputSha256"]
    assert [frame["file"] for frame in manifest["frames"]] == [
        "docs/evidence/live-excel/07-trace-a2.png",
        "docs/evidence/live-excel/08-trace-b2.png",
        "docs/evidence/live-excel/09-trace-c2.png",
    ]
    for frame in manifest["frames"]:
        path = ROOT / frame["file"]
        assert hashlib.sha256(path.read_bytes()).hexdigest() == frame["sha256"]
    with Image.open(output) as demo:
        assert demo.n_frames == 3  # pyright: ignore[reportAttributeAccessIssue]
        assert demo.size == tuple(manifest["dimensions"])
        assert demo.info["duration"] == manifest["frameDurationMs"]

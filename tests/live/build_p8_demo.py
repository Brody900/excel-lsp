"""Build the P8 lineage demo GIF from the committed live Excel captures."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

from PIL import Image

ROOT = Path(__file__).parents[2]
EVIDENCE = ROOT / "docs" / "evidence" / "live-excel"
OUTPUT = ROOT / "docs" / "assets" / "lineage-demo.gif"
MANIFEST = EVIDENCE / "demo-capture.json"
FRAME_NAMES = ("07-trace-a2.png", "08-trace-b2.png", "09-trace-c2.png")
FRAME_DURATION_MS = 1_400
OUTPUT_SIZE = (960, 508)


def build() -> dict[str, Any]:
    frames: list[Image.Image] = []
    sources: list[dict[str, Any]] = []
    for name in FRAME_NAMES:
        path = EVIDENCE / name
        raw = path.read_bytes()
        with Image.open(path) as source:
            frame = source.convert("RGB").resize(OUTPUT_SIZE, Image.Resampling.LANCZOS)
        frames.append(frame)
        sources.append(
            {
                "file": f"docs/evidence/live-excel/{name}",
                "sha256": hashlib.sha256(raw).hexdigest(),
            }
        )

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    frames[0].save(
        OUTPUT,
        save_all=True,
        append_images=frames[1:],
        duration=FRAME_DURATION_MS,
        loop=0,
        optimize=True,
    )
    payload = {
        "artifactVersion": 1,
        "output": "docs/assets/lineage-demo.gif",
        "outputSha256": hashlib.sha256(OUTPUT.read_bytes()).hexdigest(),
        "dimensions": list(OUTPUT_SIZE),
        "frameDurationMs": FRAME_DURATION_MS,
        "frames": sources,
        "provenance": "Assembled only from numbered live desktop-Excel screenshots.",
    }
    MANIFEST.write_text(
        json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8", newline="\n"
    )
    return payload


if __name__ == "__main__":
    print(json.dumps(build(), sort_keys=True))

"""Render reproducible P6 untouched-part evidence for F16 or F21."""

from __future__ import annotations

import hashlib
import json
import shutil
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile

from excel_lsp.core.edit import CellEdit, write_cells

FIXTURES = Path(__file__).parent / "generated"
CASES = {
    "F16": (
        "macro_book.xlsm",
        "MacroModel",
        ("xl/vbaProject.bin",),
    ),
    "F21": (
        "chart_image.xlsx",
        "Dashboard",
        (
            "xl/charts/chart1.xml",
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/media/image1.png",
        ),
    ),
}


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _parts(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {
            info.filename: archive.read(info) for info in archive.infolist() if not info.is_dir()
        }


def render_part_diff(fixture_id: str) -> dict[str, object]:
    """Patch a fresh fixture copy and return its complete part manifest."""
    try:
        filename, sheet, protected_parts = CASES[fixture_id]
    except KeyError as exc:
        raise ValueError(f"unknown P6 part-diff fixture: {fixture_id}") from exc

    source_path = FIXTURES / filename
    with tempfile.TemporaryDirectory(prefix=f"excel-lsp-{fixture_id.casefold()}-") as raw_dir:
        edited_path = Path(raw_dir) / filename
        shutil.copyfile(source_path, edited_path)
        before = _parts(edited_path)
        source_hash = _sha256(edited_path.read_bytes())
        result = write_cells(edited_path, (CellEdit.value(sheet, "A2", 99),))
        after = _parts(edited_path)

        names = sorted(before.keys() | after.keys())
        parts = []
        for name in names:
            before_hash = _sha256(before[name]) if name in before else None
            after_hash = _sha256(after[name]) if name in after else None
            if before_hash is None:
                status = "added"
            elif after_hash is None:
                status = "deleted"
            elif before_hash == after_hash:
                status = "preserved"
            else:
                status = "modified"
            parts.append(
                {
                    "name": name,
                    "beforeSha256": before_hash,
                    "afterSha256": after_hash,
                    "status": status,
                }
            )

        deliberate = set(result.patch.modified_parts) | set(result.patch.deleted_parts)
        untouched_identical = all(
            before.get(name) == after.get(name)
            for name in before.keys() | after.keys()
            if name not in deliberate
        )
        protected_identical = all(before[name] == after[name] for name in protected_parts)
        return {
            "artifactVersion": 1,
            "fixtureId": fixture_id,
            "fixture": filename,
            "edit": {"sheet": sheet, "ref": "A2", "value": 99},
            "sourceSha256": source_hash,
            "editedSha256": _sha256(edited_path.read_bytes()),
            "modifiedParts": list(result.patch.modified_parts),
            "deletedParts": list(result.patch.deleted_parts),
            "partCountBefore": len(before),
            "partCountAfter": len(after),
            "protectedParts": list(protected_parts),
            "protectedPartsByteIdentical": protected_identical,
            "untouchedPartsByteIdentical": untouched_identical,
            "parts": parts,
        }


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: python tests/fixtures/render_part_diff.py F16|F21", file=sys.stderr)
        return 2
    try:
        artifact = render_part_diff(sys.argv[1].upper())
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    print(json.dumps(artifact, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import json
import tomllib
from pathlib import Path

ROOT = Path(__file__).parents[2]


def test_pypi_metadata_is_complete_and_registry_verifiable() -> None:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    assert project["name"] == "excel-lsp"
    assert project["readme"] == "README.md"
    assert project["license"] == "MIT"
    assert project["requires-python"] == ">=3.11"
    assert project["urls"]["Repository"] == "https://github.com/Brody900/excel-lsp"
    assert (ROOT / "src" / "excel_lsp" / "py.typed").is_file()

    registry = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    assert registry["name"] == "io.github.Brody900/excel-lsp"
    assert registry["version"] == "0.1.0"
    assert registry["repository"]["id"] == "1315997024"
    assert registry["packages"] == [
        {
            "registryType": "pypi",
            "identifier": "excel-lsp",
            "version": "0.1.0",
            "runtimeHint": "uvx",
            "packageArguments": [{"type": "positional", "value": "serve"}],
            "transport": {"type": "stdio"},
        }
    ]
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "<!-- mcp-name: io.github.Brody900/excel-lsp -->" in readme


def test_release_workflow_is_safe_without_pypi_configuration() -> None:
    workflow = (ROOT / ".github" / "workflows" / "release.yml").read_text(encoding="utf-8")
    assert "vars.PYPI_PUBLISH_ENABLED == 'true'" in workflow
    assert "id-token: write" in workflow
    assert "pypa/gh-action-pypi-publish@v1.14.1" in workflow
    assert "password:" not in workflow

"""Tests for the installable package boundary."""

from importlib.metadata import version

import excel_lsp.core as core
from excel_lsp import __version__


def test_distribution_and_module_versions_match() -> None:
    assert version("excel-lsp") == __version__


def test_core_layer_is_importable() -> None:
    assert core.__all__ == ()

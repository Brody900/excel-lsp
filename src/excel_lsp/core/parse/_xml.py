"""Small namespace-agnostic OOXML helpers.

OOXML producers do not all use the same namespace prefix (or, in a few exported
files, even the same namespace URI).  The parser therefore matches element and
attribute local names while still requiring the expected package structure.
"""

# pyright: reportAttributeAccessIssue=false, reportMissingTypeStubs=false
# pyright: reportUnknownArgumentType=false, reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false, reportUnknownVariableType=false

from __future__ import annotations

from collections.abc import Iterator

from lxml import etree


def local_name(tag: str) -> str:
    """Return an expanded XML name's local component."""
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    return tag.split(":", 1)[-1]


def attr_by_local(element: etree._Element, name: str) -> str | None:
    """Find an attribute without depending on its namespace prefix or URI."""
    direct = element.get(name)
    if direct is not None:
        return direct
    for key, value in element.attrib.items():
        if local_name(key) == name:
            return value
    return None


def child_by_local(element: etree._Element, name: str) -> etree._Element | None:
    """Return the first direct child with ``name``."""
    for child in element:
        if isinstance(child.tag, str) and local_name(child.tag) == name:
            return child
    return None


def children_by_local(element: etree._Element, name: str) -> Iterator[etree._Element]:
    """Yield direct children with ``name``."""
    for child in element:
        if isinstance(child.tag, str) and local_name(child.tag) == name:
            yield child


def descendants_by_local(element: etree._Element, name: str) -> Iterator[etree._Element]:
    """Yield descendants with ``name``."""
    for child in element.iter():
        if child is not element and isinstance(child.tag, str) and local_name(child.tag) == name:
            yield child


def parse_xml(data: bytes) -> etree._Element:
    """Parse a small OOXML metadata part with entity resolution disabled."""
    parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False)
    return etree.fromstring(data, parser=parser)


def parse_bool(value: str | None, *, default: bool = False) -> bool:
    """Parse the XML Schema boolean spellings used by SpreadsheetML."""
    if value is None:
        return default
    return value.strip().casefold() in {"1", "true", "on"}


def clear_element(element: etree._Element) -> None:
    """Release an iterparse element and already-consumed siblings."""
    parent = element.getparent()
    element.clear(keep_tail=True)
    if parent is not None:
        while element.getprevious() is not None:
            del parent[0]


def text_content(element: etree._Element) -> str:
    """Concatenate SpreadsheetML text runs, excluding phonetic annotations."""
    chunks: list[str] = []
    for text_element in descendants_by_local(element, "t"):
        ancestor = text_element.getparent()
        is_phonetic = False
        while ancestor is not None and ancestor is not element:
            if isinstance(ancestor.tag, str) and local_name(ancestor.tag) == "rPh":
                is_phonetic = True
                break
            ancestor = ancestor.getparent()
        if not is_phonetic:
            chunks.append(text_element.text or "")
    return "".join(chunks)

"""Identity-safe exception evidence for cleanup and finalization boundaries."""

from __future__ import annotations

from collections.abc import Iterable
from typing import cast


def _derive_group(
    source: BaseExceptionGroup[BaseException],
    members: tuple[BaseException, ...],
) -> BaseExceptionGroup[BaseException]:
    rebuilt: BaseExceptionGroup[BaseException] = source.derive(members)
    rebuilt.__cause__ = source.__cause__
    rebuilt.__context__ = source.__context__
    rebuilt.__traceback__ = source.__traceback__
    for note in getattr(source, "__notes__", ()):
        rebuilt.add_note(note)
    # Assigning __cause__ changes this flag, so restore it last.
    rebuilt.__suppress_context__ = source.__suppress_context__
    return rebuilt


def normalize_exception_graph(
    error: BaseException,
    *,
    excluded: Iterable[BaseException] = (),
) -> BaseException | None:
    """Return one acyclic graph in which every exception identity occurs once."""
    owned: set[int] = {id(item) for item in excluded}

    def rebuild_membership(item: BaseException) -> BaseException | None:
        if id(item) in owned:
            return None
        owned.add(id(item))
        if not isinstance(item, BaseExceptionGroup):
            return item
        group = cast(BaseExceptionGroup[BaseException], item)
        members: list[BaseException] = []
        changed = False
        for member in group.exceptions:
            retained = rebuild_membership(member)
            if retained is None:
                changed = True
                continue
            changed = changed or retained is not member
            members.append(retained)
        if not members:
            return None
        if not changed:
            return cast(BaseException, item)
        rebuilt = _derive_group(group, tuple(members))
        owned.add(id(rebuilt))
        return rebuilt

    root = rebuild_membership(error)
    if root is None:
        return None

    normalized_links: set[int] = set()

    def normalize_links(item: BaseException) -> None:
        if id(item) in normalized_links:
            return
        normalized_links.add(id(item))
        suppressed = item.__suppress_context__
        cause = item.__cause__
        context = item.__context__
        item.__cause__ = claim_external(cause) if cause is not None else None
        item.__context__ = claim_external(context) if context is not None else None
        item.__suppress_context__ = suppressed
        if isinstance(item, BaseExceptionGroup):
            group = cast(BaseExceptionGroup[BaseException], item)
            for member in group.exceptions:
                normalize_links(member)

    def claim_external(item: BaseException) -> BaseException | None:
        if id(item) in owned:
            return None
        owned.add(id(item))
        current = item
        if isinstance(item, BaseExceptionGroup):
            group = cast(BaseExceptionGroup[BaseException], item)
            members: list[BaseException] = []
            changed = False
            for member in group.exceptions:
                retained = claim_external(member)
                if retained is None:
                    changed = True
                    continue
                changed = changed or retained is not member
                members.append(retained)
            if not members:
                return None
            if changed:
                current = _derive_group(group, tuple(members))
                owned.add(id(current))
        normalize_links(current)
        return current

    normalize_links(root)
    return root


def _exception_graph_contains(root: BaseException, target: BaseException) -> bool:
    visited: set[int] = set()
    pending = [root]
    while pending:
        item = pending.pop()
        if item is target:
            return True
        if id(item) in visited:
            continue
        visited.add(id(item))
        if item.__cause__ is not None:
            pending.append(item.__cause__)
        if item.__context__ is not None:
            pending.append(item.__context__)
        if isinstance(item, BaseExceptionGroup):
            group = cast(BaseExceptionGroup[BaseException], item)
            pending.extend(group.exceptions)
    return False


def prepare_chained_failure(
    error: BaseException,
    primary_error: BaseException,
) -> BaseException | None:
    """Sanitize cleanup evidence before installing it as a primary's cause."""
    excluded: list[BaseException] = []
    visited: set[int] = set()
    pending = [primary_error]
    while pending:
        item = pending.pop()
        if id(item) in visited:
            continue
        visited.add(id(item))
        excluded.append(item)
        if item.__cause__ is not None:
            pending.append(item.__cause__)
        if item.__context__ is not None:
            pending.append(item.__context__)
        if isinstance(item, BaseExceptionGroup):
            group = cast(BaseExceptionGroup[BaseException], item)
            pending.extend(group.exceptions)
    return normalize_exception_graph(error, excluded=excluded)


def prepare_chained_failure_with_primary_evidence(
    error: BaseException,
    primary_error: BaseException,
    *,
    message: str,
) -> BaseException | None:
    """Preserve visible prior causal evidence before installing cleanup as cause."""
    previous_cause = primary_error.__cause__
    previous_context = primary_error.__context__
    context_was_suppressed = primary_error.__suppress_context__
    previous_error = previous_cause
    if previous_error is None and not context_was_suppressed:
        previous_error = previous_context
    if previous_error is not None:
        if primary_error.__cause__ is previous_error:
            primary_error.__cause__ = None
        if primary_error.__context__ is previous_error:
            primary_error.__context__ = None
        if not _exception_graph_contains(error, previous_error):
            error = BaseExceptionGroup(message, (previous_error, error))
    return prepare_chained_failure(error, primary_error)


__all__ = [
    "normalize_exception_graph",
    "prepare_chained_failure",
    "prepare_chained_failure_with_primary_evidence",
]

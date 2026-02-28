"""This module contains a helper function for flattening a Base object.

This is the standard flatten.py provided by Speckle Automate boilerplate.
It recursively traverses the Speckle object tree and yields every child object.
"""

from collections.abc import Iterable
from specklepy.objects import Base


def flatten_base(base: Base) -> Iterable[Base]:
    """Flatten a Speckle Base object into an iterable of its children.

    Args:
        base: The root Speckle object to flatten.

    Yields:
        Every Base object found in the tree (depth-first).
    """
    yield base

    for key in base.get_dynamic_member_names():
        value = base[key]
        yield from _flatten_value(value)


def _flatten_value(value) -> Iterable[Base]:
    if isinstance(value, Base):
        yield from flatten_base(value)
    elif isinstance(value, list):
        for item in value:
            yield from _flatten_value(item)

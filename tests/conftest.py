"""Shared test helpers: a fake COM connection and mock builders.

Controller tests do not need Office installed - COM objects are replaced
podmieniane na ``MagicMock``, a warstwa watkow/timeoutow na
:class:`FakeConnection`, which runs calls synchronously.
"""

from __future__ import annotations

from typing import Any, Sequence
from unittest.mock import MagicMock

import pytest

from bridge.utils.com_helpers import com_error


class FakeConnection:
    """Zastepuje :class:`bridge.connection_manager.AppConnection` w testach."""

    def __init__(self, app: Any | None = None, key: str = "test") -> None:
        self.app = app if app is not None else MagicMock()
        self.key = key
        self.reset_count = 0

    def run(self, func, *args, timeout: float | None = None):
        return func(*args)

    def application(self) -> Any:
        return self.app

    def reset(self) -> None:
        self.reset_count += 1

    def status(self) -> dict[str, Any]:
        return {"app": self.key, "connected": True, "last_error": None}


def com_collection(items: Sequence[Any]) -> MagicMock:
    """Buduje mock kolekcji COM: ``Count`` + wywolanie z indeksem 1-based."""
    collection = MagicMock()
    collection.Count = len(items)
    collection.side_effect = lambda index, *_: items[int(index) - 1]
    collection.__iter__ = lambda _self: iter(items)
    return collection


def make_com_error(hresult: int = -2147352567, description: str = "COM error") -> Exception:
    """Builds a COM exception with the given HRESULT and ``excepinfo`` text."""
    return com_error(hresult, "Test", (0, "Office", description, None, 0, hresult), None)


def make_text_frame(text: str = "", has_text: bool | None = None) -> MagicMock:
    """Mock ``TextFrame`` with text set and a working ``TextRange``."""
    frame = MagicMock()
    frame.HasText = bool(text) if has_text is None else has_text
    frame.TextRange.Text = text
    return frame


def make_shape(
    shape_id: int = 1,
    name: str = "Shape 1",
    text: str = "",
    has_text_frame: bool = True,
    shape_type: int = 17,
    placeholder_type: int | None = None,
) -> MagicMock:
    """Mock slide shape with the properties that get queried most often."""
    shape = MagicMock()
    shape.Id = shape_id
    shape.Name = name
    shape.Type = 14 if placeholder_type else shape_type
    shape.Left, shape.Top, shape.Width, shape.Height = 10.0, 20.0, 300.0, 100.0
    shape.HasTextFrame = has_text_frame
    shape.HasTable = False
    shape.TextFrame = make_text_frame(text)

    if placeholder_type is not None:
        shape.PlaceholderFormat.Type = placeholder_type

    return shape


@pytest.fixture
def fake_connection() -> FakeConnection:
    return FakeConnection()

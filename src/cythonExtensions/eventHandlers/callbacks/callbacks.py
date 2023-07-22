"""This module contains the event handlers for the keyboard and mouse events."""

from typing import Callable


kbEventHandlers: dict[tuple[int, int], Callable] = {}
"""A dictionary of the keyboard event handlers that does not require any lock keys to be on."""


kbEventHandlersWithLockKeys: dict[tuple[int, int, int], Callable] = {}
"""A dictionary of the keyboard event handlers that requires lock keys to be on."""

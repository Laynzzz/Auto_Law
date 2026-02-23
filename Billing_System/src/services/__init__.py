"""Service layer — shared business logic for CLI and GUI.

Every public service function returns a ServiceResult so callers can
inspect success/failure, read a human-friendly message, and access
structured data without coupling to module internals.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ServiceResult:
    """Standardised return type for all service functions.

    Attributes:
        success: True when the operation completed without expected errors.
        message: Human-readable summary suitable for CLI echo or GUI dialog.
        data:    Structured output (paths, counts, row dicts, etc.).
    """

    success: bool
    message: str
    data: dict | None = field(default=None)

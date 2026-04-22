"""
Diagnostic helpers for AuditPro runtime failures.

Provides a consistent way to capture root-cause context across the UI,
API, and agent layers without duplicating traceback formatting logic.
"""
from __future__ import annotations

from traceback import format_exception


def build_diagnostic(exc: Exception, context: dict | None = None, trace_lines: int = 8) -> dict:
    """Build a compact structured diagnostic payload for an exception."""
    trace = "".join(format_exception(type(exc), exc, exc.__traceback__)).strip().splitlines()
    return {
        "type": type(exc).__name__,
        "message": str(exc),
        "context": context or {},
        "traceback": trace[-trace_lines:],
    }


def format_diagnostic(diag: dict) -> str:
    """Render a structured diagnostic payload into a readable message."""
    parts = [f"{diag.get('type', 'Error')}: {diag.get('message', '')}"]
    context = diag.get("context") or {}
    if context:
        context_text = ", ".join(f"{key}={value}" for key, value in context.items())
        parts.append(f"Contexte: {context_text}")
    trace = diag.get("traceback") or []
    if trace:
        parts.append("Trace:")
        parts.extend(trace)
    return "\n".join(parts)
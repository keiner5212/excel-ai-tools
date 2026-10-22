from __future__ import annotations

from toon_format import dumps


def render(payload: object) -> str:
    """Encode a tool result as TOON text for the model."""
    return dumps(payload)

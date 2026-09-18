"""excel_line_core.llm — harness-agnostic LLM client protocol.

The core package never imports a runtime-specific LLM facade. Instead it
declares the minimal callable contract any harness can satisfy:

    complete(prompt: str, **kwargs) -> str   (empty string => unavailable)

Harnesses inject a concrete client (e.g. a Hermes ctx.llm wrapper) at
construction time. The core only depends on the protocol below.
"""
from __future__ import annotations

from typing import Any, Callable, Protocol, runtime_checkable

AnyLLM = Any


@runtime_checkable
class BaseLLMClient(Protocol):
    """Minimal LLM contract. Return '' (never raise) when unavailable."""

    def complete(self, prompt: str, **kwargs) -> str:
        ...


# A plain callable also satisfies the protocol at runtime.
LLMClient = Callable[..., str]


def classify(prompt: str, client: "AnyLLM", **kwargs) -> str:
    """Classify one record through any client; '' means 'LLM down'.

    Used by the worker so it can stay free of runtime imports: pass in a
    callable/BaseLLMClient that answers a prompt with text (or '' when the
    model is unreachable).
    """
    if client is None:
        return ""
    try:
        if callable(client):
            out = client(prompt, **kwargs)
        else:
            out = client.complete(prompt, **kwargs)
        return (out or "").strip()
    except Exception:
        return ""
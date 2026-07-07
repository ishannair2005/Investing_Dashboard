"""
ai_client.py

Thin, provider-agnostic wrapper around the AI backend used both for
per-headline news classification (news.py) and full investment
analysis (analysis.py). Switching config.AI_PROVIDER between
"anthropic" and "openai" is the only change needed to move providers --
neither calling module imports a provider SDK directly.

Not part of the file list in the original spec, but both news.py and
analysis.py need this same provider-switching behavior; without a
shared module, "swap the AI provider" would mean editing two files
identically instead of one.
"""

import logging
from typing import Optional

from config import AI_MODEL, AI_PROVIDER, ANTHROPIC_API_KEY, OPENAI_API_KEY

logger = logging.getLogger(__name__)


def generate(prompt: str, system: Optional[str] = None, max_tokens: int = 1024) -> str:
    """Send a prompt to the configured AI provider, return its text response.

    Raises RuntimeError for an unsupported AI_PROVIDER or a missing API
    key -- callers should not silently proceed without the AI backend
    they think is configured.
    """
    if AI_PROVIDER == "anthropic":
        return _generate_anthropic(prompt, system, max_tokens)
    if AI_PROVIDER == "openai":
        return _generate_openai(prompt, system, max_tokens)
    raise RuntimeError(f"Unsupported AI_PROVIDER: {AI_PROVIDER!r} (expected 'anthropic' or 'openai')")


def _generate_anthropic(prompt: str, system: Optional[str], max_tokens: int) -> str:
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY is not set (required when AI_PROVIDER=anthropic)")
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    kwargs = {"model": AI_MODEL, "max_tokens": max_tokens, "messages": [{"role": "user", "content": prompt}]}
    if system:
        kwargs["system"] = system
    response = client.messages.create(**kwargs)
    return "".join(block.text for block in response.content if block.type == "text")


def _generate_openai(prompt: str, system: Optional[str], max_tokens: int) -> str:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is not set (required when AI_PROVIDER=openai)")
    import openai

    client = openai.OpenAI(api_key=OPENAI_API_KEY)
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    response = client.chat.completions.create(model=AI_MODEL, max_tokens=max_tokens, messages=messages)
    return response.choices[0].message.content

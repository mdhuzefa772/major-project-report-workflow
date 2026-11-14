"""Thin helper around the workflow-aware Mistral client for JSON-constrained LLM calls."""

import json
import os
import re

from mistralai.workflows.client import get_mistral_client

DEFAULT_MODEL = "mistral-large-latest"


def get_model() -> str:
    return os.environ.get("REPORT_MODEL", DEFAULT_MODEL)


def _strip_fences(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\n", "", text)
        text = re.sub(r"\n```$", "", text)
    return text.strip()


async def complete_json(system: str, user: str, temperature: float = 0.2) -> dict:
    """Call the model and parse a JSON object out of the response."""
    client = get_mistral_client()
    response = await client.chat.complete_async(
        model=get_model(),
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        response_format={"type": "json_object"},
        temperature=temperature,
    )
    content = response.choices[0].message.content or "{}"
    return json.loads(_strip_fences(content))


async def complete_text(system: str, user: str, temperature: float = 0.2) -> str:
    client = get_mistral_client()
    response = await client.chat.complete_async(
        model=get_model(),
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        temperature=temperature,
    )
    return response.choices[0].message.content or ""

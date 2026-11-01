"""Thin OpenAI chat-completions client with retry/backoff for rate limits."""

from __future__ import annotations

import logging
import os
import random
import time
from typing import Callable, List, Optional, Sequence

import requests

from .config import ModelSettings
from .io_utils import encode_image

log = logging.getLogger(__name__)


class LLMClient:
    def __init__(self, settings: Optional[ModelSettings] = None, api_key: Optional[str] = None,
                 session: Optional[requests.Session] = None, sleep: Callable[[float], None] = time.sleep):
        self.settings = settings or ModelSettings()
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        if not self.api_key:
            raise RuntimeError("OPENAI_API_KEY is not set. Copy .env.example to .env and add your key.")
        self.session = session or requests.Session()
        self._sleep = sleep

    # ---------- public ----------
    def describe(self, prompt: str, image_paths: Sequence[str] = ()) -> str:
        """Send the prompt (and up to max_images_per_call images) and return the model's text."""
        content: List[dict] = [{"type": "text", "text": prompt}]
        for path in list(image_paths)[: self.settings.max_images_per_call]:
            try:
                content.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{encode_image(path)}"},
                })
            except OSError as exc:
                log.warning("Could not read image %s: %s", path, exc)

        payload = {
            "model": self.settings.model,
            "messages": [{"role": "user", "content": content}],
            "max_tokens": self.settings.max_tokens,
        }
        if self.settings.temperature is not None:
            payload["temperature"] = self.settings.temperature

        data = self._post_with_retries(payload)
        return data["choices"][0]["message"]["content"]

    # ---------- internals ----------
    def _backoff(self, attempt: int) -> float:
        return self.settings.base_backoff_s * (2 ** attempt) + random.uniform(0, 0.5)

    def _post_with_retries(self, payload: dict) -> dict:
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        retries = self.settings.max_retries
        for attempt in range(retries):
            try:
                resp = self.session.post(self.settings.url, headers=headers, json=payload,
                                         timeout=self.settings.timeout_s)
                if resp.status_code == 429:
                    retry_after = resp.headers.get("Retry-After")
                    wait = float(retry_after) if retry_after else self._backoff(attempt)
                    log.warning("429 rate limited - waiting %.1fs (attempt %d/%d)", wait, attempt + 1, retries)
                    self._sleep(wait)
                    continue
                resp.raise_for_status()
                return resp.json()
            except requests.exceptions.RequestException as exc:
                wait = self._backoff(attempt)
                log.warning("API error (%s) - backing off %.1fs (attempt %d/%d)", exc, wait, attempt + 1, retries)
                self._sleep(wait)
        raise RuntimeError(f"Giving up after {retries} attempts (rate limits / network errors).")

import pytest
import requests

from museum_catalog.config import ModelSettings
from museum_catalog.llm import LLMClient


class FakeResponse:
    def __init__(self, status=200, body=None, headers=None):
        self.status_code = status
        self._body = body or {"choices": [{"message": {"content": "Beschreibung"}}]}
        self.headers = headers or {}

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.exceptions.HTTPError(str(self.status_code))

    def json(self):
        return self._body


class FakeSession:
    def __init__(self, responses):
        self.responses = list(responses)
        self.payloads = []

    def post(self, url, headers, json, timeout):
        self.payloads.append(json)
        return self.responses.pop(0)


def make_client(responses, **settings):
    session = FakeSession(responses)
    sleeps = []
    client = LLMClient(ModelSettings(**settings), api_key="test", session=session, sleep=sleeps.append)
    return client, session, sleeps


def test_text_only_request():
    client, session, _ = make_client([FakeResponse()], temperature=0.0)
    assert client.describe("Prompt") == "Beschreibung"
    payload = session.payloads[0]
    assert payload["temperature"] == 0.0
    assert payload["messages"][0]["content"] == [{"type": "text", "text": "Prompt"}]


def test_images_are_capped_and_encoded(tmp_path):
    images = []
    for i in range(4):
        p = tmp_path / f"{i}.jpg"
        p.write_bytes(b"img")
        images.append(str(p))
    client, session, _ = make_client([FakeResponse()], max_images_per_call=2)
    client.describe("Prompt", images)
    content = session.payloads[0]["messages"][0]["content"]
    assert len(content) == 3  # text + 2 images
    assert content[1]["image_url"]["url"].startswith("data:image/jpeg;base64,")
    assert "temperature" not in session.payloads[0]


def test_retries_on_rate_limit_using_retry_after():
    client, _, sleeps = make_client([FakeResponse(429, headers={"Retry-After": "7"}), FakeResponse()])
    assert client.describe("Prompt") == "Beschreibung"
    assert sleeps == [7.0]


def test_gives_up_after_max_retries():
    client, _, sleeps = make_client([FakeResponse(500)] * 3, max_retries=3)
    with pytest.raises(RuntimeError, match="Giving up"):
        client.describe("Prompt")
    assert len(sleeps) == 3


def test_missing_api_key(monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    with pytest.raises(RuntimeError, match="OPENAI_API_KEY"):
        LLMClient()

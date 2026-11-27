"""Tests for vision-based slide extraction in vision_converter."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from types import SimpleNamespace

import httpx2
import pymupdf
import pytest
from PIL import Image, ImageDraw

import pdf_to_md.vision_converter as vc
from pdf_to_md.vision_converter import DEFAULT_MODEL, convert_presentation_vision


class _FakeSlideModels:
    def __init__(self, handler):
        self._handler = handler
        self.calls: list[dict] = []

    async def generate_content(self, *, model, contents, config):
        parts = contents[0]["parts"]
        prompt_text = parts[0]["text"]
        inline_data = parts[1]["inline_data"]
        self.calls.append(
            {
                "model": model,
                "config": config,
                "prompt": prompt_text,
                "mime_type": inline_data["mime_type"],
                "data_len": len(inline_data["data"]),
            }
        )
        result = self._handler(prompt_text, len(self.calls))
        if isinstance(result, Exception):
            raise result
        if isinstance(result, str):
            return SimpleNamespace(text=result)
        return SimpleNamespace(text=json.dumps(result))


class _FakeAsyncClient:
    def __init__(self, handler):
        self.models = _FakeSlideModels(handler)
        self.closed = False

    async def aclose(self):
        self.closed = True


def _write_slides_pdf(path: Path, count: int = 3) -> None:
    doc = pymupdf.open()
    try:
        for i in range(1, count + 1):
            page = doc.new_page(width=640, height=360)
            page.insert_text((48, 72), f"Slide {i} Topic", fontsize=18)
            page.insert_text((48, 112), f"Details for slide {i}.", fontsize=12)
        doc.save(path)
    finally:
        doc.close()


def _write_slide_png(path: Path, width: int = 240, height: int = 160) -> None:
    img = Image.new("RGB", (width, height), (250, 250, 250))
    draw = ImageDraw.Draw(img)
    draw.rectangle((16, 16, width - 16, height - 16), fill=(30, 80, 170))
    img.save(path, format="PNG")


def test_default_model_and_slide_pdf_exports(tmp_path, monkeypatch):
    assert DEFAULT_MODEL == "gemini-4-hydrogen"

    pdf_path = tmp_path / "deck (v1).pdf"
    _write_slides_pdf(pdf_path, count=3)

    def _handler(prompt: str, call_idx: int) -> dict:
        slide_num = 1
        for line in prompt.splitlines():
            if line.startswith("Slide number: "):
                slide_num = int(line.split(": ", 1)[1])
        return {
            "topic": f"Topic {slide_num}",
            "technologies": ["Python", "PostgreSQL"],
            "markdown_body": f"# Slide {slide_num}\n\nExtracted content {slide_num}.",
        }

    # 1. Default model and default thinking_level (None -> no thinking_config)
    default_client = _FakeAsyncClient(_handler)
    monkeypatch.setattr(vc, "_create_async_client", lambda api_key: default_client)

    out_default = tmp_path / "out_default"
    written_default = convert_presentation_vision(
        pdf_path,
        out_default,
        api_key="test-key",
        dpi=100,
    )
    assert len(written_default) == 3
    assert len(default_client.models.calls) == 3
    assert default_client.closed is True
    assert all(
        c["model"] == "gemini-4-hydrogen"
        and c["config"] == {"response_mime_type": "application/json"}
        for c in default_client.models.calls
    )

    # 2. Custom model, thinking_level="xhigh", export_format="both", and page range
    custom_client = _FakeAsyncClient(_handler)
    monkeypatch.setattr(vc, "_create_async_client", lambda api_key: custom_client)

    out_both = tmp_path / "out_both"
    written_both = convert_presentation_vision(
        pdf_path,
        out_both,
        api_key="test-key",
        model="gemini-3.8-flash",
        thinking_level="xhigh",
        dpi=100,
        export_format="both",
        page_start=2,
        page_end=3,
    )
    assert [p.name for p in written_both] == [
        "deck_v1__slide_002.md",
        "deck_v1__slide_003.md",
        "deck_v1_.json",
    ]
    assert len(custom_client.models.calls) == 2
    assert all(
        c["model"] == "gemini-3.8-flash"
        and c["config"]
        == {
            "response_mime_type": "application/json",
            "thinking_config": {"thinking_level": "XHIGH"},
        }
        for c in custom_client.models.calls
    )

    slide2_text = (out_both / "deck_v1__slide_002.md").read_text(encoding="utf-8")
    assert 'document_source: "deck (v1).pdf"' in slide2_text
    assert "slide_number: 2" in slide2_text
    assert 'topic: "Topic 2"' in slide2_text
    assert 'technologies: ["Python", "PostgreSQL"]' in slide2_text
    assert "vision_analyzed: true" in slide2_text
    assert "# Slide 2" in slide2_text

    unified = json.loads((out_both / "deck_v1_.json").read_text(encoding="utf-8"))
    assert [entry["slide_number"] for entry in unified] == [2, 3]
    assert unified[0]["topic"] == "Topic 2"
    assert unified[1]["vision_analyzed"] is True


def test_single_png_input_and_cli_forwarding(tmp_path, monkeypatch):
    img_path = tmp_path / "slide.png"
    _write_slide_png(img_path)

    client = _FakeAsyncClient(
        lambda prompt, call_idx: {
            "topic": "Diagram",
            "technologies": ["Kafka"],
            "markdown_body": "- Broker topology",
        }
    )
    monkeypatch.setattr(vc, "_create_async_client", lambda api_key: client)

    out_dir = tmp_path / "cli_out"
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "pdf-to-md-vision",
            str(img_path),
            "-o",
            str(out_dir),
            "--api-key",
            "test-key",
            "--model",
            "gemini-4-hydrogen",
            "--thinking-level",
            "HIGH",
        ],
    )
    vc.main()

    assert len(client.models.calls) == 1
    assert client.models.calls[0]["model"] == "gemini-4-hydrogen"
    assert client.models.calls[0]["config"] == {
        "response_mime_type": "application/json",
        "thinking_config": {"thinking_level": "HIGH"},
    }
    assert client.models.calls[0]["mime_type"] == "image/png"
    assert "- Broker topology" in (out_dir / "slide_slide_001.md").read_text(encoding="utf-8")


def test_retries_and_is_retryable(tmp_path, monkeypatch):
    img_path = tmp_path / "retry_slide.png"
    _write_slide_png(img_path)

    sleep_calls: list[float] = []

    async def _fast_sleep(delay: float) -> None:
        sleep_calls.append(delay)

    monkeypatch.setattr(vc.asyncio, "sleep", _fast_sleep)

    # Malformed JSON first, valid JSON on second attempt -> succeeds after 1 retry
    def _flaky_handler(prompt: str, call_idx: int):
        if call_idx == 1:
            return "not valid json"
        return {
            "topic": "Recovered",
            "technologies": [],
            "markdown_body": "Recovered body",
        }

    flaky_client = _FakeAsyncClient(_flaky_handler)
    monkeypatch.setattr(vc, "_create_async_client", lambda api_key: flaky_client)

    written = convert_presentation_vision(
        img_path, tmp_path / "out_retry", api_key="k", max_retries=2
    )
    assert len(written) == 1
    assert len(flaky_client.models.calls) == 2
    assert len(sleep_calls) == 1

    # Non-retryable HTTP 400 error -> fails immediately without sleeping
    class _ApiErr(Exception):
        def __init__(self, code: int):
            super().__init__(f"HTTP {code}")
            self.code = code

    sleep_calls.clear()
    bad_client = _FakeAsyncClient(lambda prompt, call_idx: _ApiErr(400))
    monkeypatch.setattr(vc, "_create_async_client", lambda api_key: bad_client)

    with pytest.raises(RuntimeError, match="failed after 1 attempt"):
        convert_presentation_vision(
            img_path, tmp_path / "out_fail", api_key="k", max_retries=3
        )
    assert len(bad_client.models.calls) == 1
    assert sleep_calls == []

    # Direct checks for _is_retryable
    assert vc._is_retryable(TimeoutError()) is True
    assert vc._is_retryable(ConnectionError()) is True
    assert vc._is_retryable(_ApiErr(429)) is True
    assert vc._is_retryable(_ApiErr(503)) is True
    assert vc._is_retryable(_ApiErr(400)) is False
    assert vc._is_retryable(httpx2.TransportError("network reset")) is True


def test_option_validation(tmp_path):
    img_path = tmp_path / "slide.png"
    _write_slide_png(img_path)

    with pytest.raises(ValueError, match="thinking_level"):
        convert_presentation_vision(img_path, api_key="k", thinking_level="invalid")

    with pytest.raises(ValueError, match="model"):
        convert_presentation_vision(img_path, api_key="k", model="  ")

    with pytest.raises(ValueError, match="export_format"):
        convert_presentation_vision(img_path, api_key="k", export_format="xml")

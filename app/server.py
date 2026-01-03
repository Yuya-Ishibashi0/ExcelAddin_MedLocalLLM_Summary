from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

def load_settings() -> dict[str, Any]:
    settings_path = Path(__file__).with_name("settings.json")
    try:
        with settings_path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
            return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        return {}


SETTINGS = load_settings()

OLLAMA_URL = os.getenv("OLLAMA_URL") or SETTINGS.get("ollama_url") or "http://127.0.0.1:11434"
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL") or SETTINGS.get("ollama_model") or "gemma3:1b-it-qat"
REQUEST_TIMEOUT = float(os.getenv("OLLAMA_TIMEOUT") or SETTINGS.get("timeout", 120))
OLLAMA_PRESET = (os.getenv("OLLAMA_PRESET") or SETTINGS.get("preset") or "").strip().lower()
OLLAMA_NUM_PREDICT_DEFAULT = int(SETTINGS.get("num_predict_default", 512))
OLLAMA_NUM_PREDICT_RAW = os.getenv("OLLAMA_NUM_PREDICT") or SETTINGS.get("num_predict")
OLLAMA_NUM_CTX_RAW = os.getenv("OLLAMA_NUM_CTX") or SETTINGS.get("num_ctx")
OLLAMA_NUM_THREAD_RAW = os.getenv("OLLAMA_NUM_THREAD") or SETTINGS.get("num_thread")
OLLAMA_NUM_BATCH_RAW = os.getenv("OLLAMA_NUM_BATCH") or SETTINGS.get("num_batch")
OLLAMA_TEMPERATURE_RAW = os.getenv("OLLAMA_TEMPERATURE") or SETTINGS.get("temperature")
OLLAMA_TOP_K_RAW = os.getenv("OLLAMA_TOP_K") or SETTINGS.get("top_k")
OLLAMA_TOP_P_RAW = os.getenv("OLLAMA_TOP_P") or SETTINGS.get("top_p")
OLLAMA_REPEAT_PENALTY_RAW = os.getenv("OLLAMA_REPEAT_PENALTY") or SETTINGS.get("repeat_penalty")
OLLAMA_REPEAT_LAST_N_RAW = os.getenv("OLLAMA_REPEAT_LAST_N") or SETTINGS.get("repeat_last_n")
OLLAMA_PRESENCE_PENALTY_RAW = os.getenv("OLLAMA_PRESENCE_PENALTY") or SETTINGS.get("presence_penalty")
OLLAMA_FREQUENCY_PENALTY_RAW = os.getenv("OLLAMA_FREQUENCY_PENALTY") or SETTINGS.get("frequency_penalty")
OLLAMA_TFS_Z_RAW = os.getenv("OLLAMA_TFS_Z") or SETTINGS.get("tfs_z")
OLLAMA_TYPICAL_P_RAW = os.getenv("OLLAMA_TYPICAL_P") or SETTINGS.get("typical_p")
OLLAMA_MIROSTAT_RAW = os.getenv("OLLAMA_MIROSTAT") or SETTINGS.get("mirostat")
OLLAMA_MIROSTAT_TAU_RAW = os.getenv("OLLAMA_MIROSTAT_TAU") or SETTINGS.get("mirostat_tau")
OLLAMA_MIROSTAT_ETA_RAW = os.getenv("OLLAMA_MIROSTAT_ETA") or SETTINGS.get("mirostat_eta")
OLLAMA_PENALIZE_NEWLINE_RAW = os.getenv("OLLAMA_PENALIZE_NEWLINE") or SETTINGS.get("penalize_newline")
OLLAMA_SEED_RAW = os.getenv("OLLAMA_SEED") or SETTINGS.get("seed")
OLLAMA_STOP_RAW = os.getenv("OLLAMA_STOP") or SETTINGS.get("stop")
SYSTEM_PROMPT = os.getenv("SYSTEM_PROMPT") or SETTINGS.get(
    "system_prompt",
    "回答は短く完結。長くなる場合は重要点のみ箇条書き3件以内。続きを書かない。",
)


class ChatMessage(BaseModel):
    role: str
    content: str


class ChatRequest(BaseModel):
    messages: list[ChatMessage]
    selection_context: str | None = None
    model: str | None = None
    max_tokens: int | None = None
    preset: str | None = None


class ChatResponse(BaseModel):
    answer: str


app = FastAPI(title="Local LLM Orchestrator", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"]
)


@app.get("/health")
def health() -> dict[str, bool]:
    return {"ok": True}


def parse_optional_int(value: str | None) -> int | None:
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def parse_optional_float(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def parse_optional_bool(value: str | None) -> bool | None:
    if value is None:
        return None
    lowered = value.strip().lower()
    if lowered in {"1", "true", "yes", "on"}:
        return True
    if lowered in {"0", "false", "no", "off"}:
        return False
    return None


def parse_stop_list(value: str | None) -> list[str] | None:
    if not value:
        return None
    trimmed = value.strip()
    if not trimmed:
        return None
    if trimmed.startswith("["):
        try:
            data = json.loads(trimmed)
        except json.JSONDecodeError:
            return None
        if isinstance(data, list):
            return [str(item) for item in data if str(item).strip()]
        return None
    return [item.strip() for item in trimmed.split(",") if item.strip()]


def resolve_preset_name(value: str | None) -> str | None:
    if not value:
        return None
    name = value.strip().lower()
    if not name or name in {"custom", "none", "default"}:
        return None
    return name


def get_preset_options(name: str) -> dict[str, int] | None:
    presets: dict[str, dict[str, int]] = {
        "fast": {"num_ctx": 1024, "num_predict": 256},
        "balanced": {"num_ctx": 2048, "num_predict": 512},
        "long": {"num_ctx": 4096, "num_predict": 1024},
    }
    return presets.get(name)


def build_payload(req: ChatRequest, stream: bool) -> tuple[dict[str, Any], str]:
    model = (req.model or "").strip() or OLLAMA_MODEL
    messages: list[dict[str, str]] = []
    if SYSTEM_PROMPT:
        messages.append({"role": "system", "content": SYSTEM_PROMPT})
    if req.selection_context and req.selection_context.strip():
        messages.append(
            {
                "role": "system",
                "content": f"選択範囲コンテキスト:\\n{req.selection_context}",
            }
        )

    for message in req.messages:
        if hasattr(message, "model_dump"):
            messages.append(message.model_dump())
        else:
            messages.append(message.dict())

    preset_name = resolve_preset_name(req.preset) or resolve_preset_name(OLLAMA_PRESET)
    preset = get_preset_options(preset_name) if preset_name else None
    num_predict = req.max_tokens if req.max_tokens and req.max_tokens > 0 else None
    if num_predict is None:
        num_predict = parse_optional_int(OLLAMA_NUM_PREDICT_RAW)
    if num_predict is None and preset:
        num_predict = preset.get("num_predict")
    if num_predict is None:
        num_predict = OLLAMA_NUM_PREDICT_DEFAULT
    num_ctx = parse_optional_int(OLLAMA_NUM_CTX_RAW)
    if num_ctx is None and preset:
        num_ctx = preset.get("num_ctx")
    num_thread = parse_optional_int(OLLAMA_NUM_THREAD_RAW)
    num_batch = parse_optional_int(OLLAMA_NUM_BATCH_RAW)
    options: dict[str, Any] = {}
    if num_predict > 0:
        options["num_predict"] = num_predict
    if num_ctx and num_ctx > 0:
        options["num_ctx"] = num_ctx
    if num_thread and num_thread > 0:
        options["num_thread"] = num_thread
    if num_batch and num_batch > 0:
        options["num_batch"] = num_batch
    temperature = parse_optional_float(OLLAMA_TEMPERATURE_RAW)
    if temperature is not None:
        options["temperature"] = temperature
    top_k = parse_optional_int(OLLAMA_TOP_K_RAW)
    if top_k is not None:
        options["top_k"] = top_k
    top_p = parse_optional_float(OLLAMA_TOP_P_RAW)
    if top_p is not None:
        options["top_p"] = top_p
    repeat_penalty = parse_optional_float(OLLAMA_REPEAT_PENALTY_RAW)
    if repeat_penalty is not None:
        options["repeat_penalty"] = repeat_penalty
    repeat_last_n = parse_optional_int(OLLAMA_REPEAT_LAST_N_RAW)
    if repeat_last_n is not None:
        options["repeat_last_n"] = repeat_last_n
    presence_penalty = parse_optional_float(OLLAMA_PRESENCE_PENALTY_RAW)
    if presence_penalty is not None:
        options["presence_penalty"] = presence_penalty
    frequency_penalty = parse_optional_float(OLLAMA_FREQUENCY_PENALTY_RAW)
    if frequency_penalty is not None:
        options["frequency_penalty"] = frequency_penalty
    tfs_z = parse_optional_float(OLLAMA_TFS_Z_RAW)
    if tfs_z is not None:
        options["tfs_z"] = tfs_z
    typical_p = parse_optional_float(OLLAMA_TYPICAL_P_RAW)
    if typical_p is not None:
        options["typical_p"] = typical_p
    mirostat = parse_optional_int(OLLAMA_MIROSTAT_RAW)
    if mirostat is not None:
        options["mirostat"] = mirostat
    mirostat_tau = parse_optional_float(OLLAMA_MIROSTAT_TAU_RAW)
    if mirostat_tau is not None:
        options["mirostat_tau"] = mirostat_tau
    mirostat_eta = parse_optional_float(OLLAMA_MIROSTAT_ETA_RAW)
    if mirostat_eta is not None:
        options["mirostat_eta"] = mirostat_eta
    penalize_newline = parse_optional_bool(OLLAMA_PENALIZE_NEWLINE_RAW)
    if penalize_newline is not None:
        options["penalize_newline"] = penalize_newline
    seed = parse_optional_int(OLLAMA_SEED_RAW)
    if seed is not None:
        options["seed"] = seed
    stop = parse_stop_list(OLLAMA_STOP_RAW)
    if stop:
        options["stop"] = stop
    payload: dict[str, Any] = {
        "model": model,
        "messages": messages,
        "stream": stream,
    }
    if options:
        payload["options"] = options
    return payload, model


@app.post("/chat", response_model=ChatResponse)
def chat(req: ChatRequest) -> ChatResponse:
    if not req.messages:
        return ChatResponse(answer="")

    payload, model = build_payload(req, stream=False)

    try:
        resp = requests.post(
            f"{OLLAMA_URL}/api/chat",
            json=payload,
            timeout=REQUEST_TIMEOUT,
        )
    except requests.exceptions.ConnectionError as exc:
        raise HTTPException(
            status_code=503,
            detail="Ollamaが起動していません。ollama serve を実行してください",
        ) from exc
    except requests.exceptions.Timeout as exc:
        raise HTTPException(
            status_code=504,
            detail="応答が遅延しています。モデルを軽くする/入力範囲を減らす…",
        ) from exc

    if resp.status_code == 404:
        raise HTTPException(
            status_code=404,
            detail=f"モデルが見つかりません。ollama pull {model} を実行してください",
        )

    if resp.status_code >= 400:
        detail = ""
        try:
            detail = resp.json().get("error", "")
        except Exception:
            detail = ""
        raise HTTPException(status_code=resp.status_code, detail=detail or "Ollamaエラー")

    data = resp.json()
    content = data.get("message", {}).get("content", "")
    return ChatResponse(answer=content)


@app.post("/chat/stream")
def chat_stream(req: ChatRequest) -> StreamingResponse:
    if not req.messages:
        return StreamingResponse(iter(()), media_type="text/plain")

    payload, model = build_payload(req, stream=True)

    try:
        resp = requests.post(
            f"{OLLAMA_URL}/api/chat",
            json=payload,
            timeout=REQUEST_TIMEOUT,
            stream=True,
        )
    except requests.exceptions.ConnectionError as exc:
        raise HTTPException(
            status_code=503,
            detail="Ollamaが起動していません。ollama serve を実行してください",
        ) from exc
    except requests.exceptions.Timeout as exc:
        raise HTTPException(
            status_code=504,
            detail="応答が遅延しています。モデルを軽くする/入力範囲を減らす…",
        ) from exc

    if resp.status_code == 404:
        resp.close()
        raise HTTPException(
            status_code=404,
            detail=f"モデルが見つかりません。ollama pull {model} を実行してください",
        )

    if resp.status_code >= 400:
        detail = ""
        try:
            detail = resp.json().get("error", "")
        except Exception:
            detail = ""
        resp.close()
        raise HTTPException(status_code=resp.status_code, detail=detail or "Ollamaエラー")

    def iter_stream() -> Any:
        with resp:
            for line in resp.iter_lines(decode_unicode=True):
                if not line:
                    continue
                try:
                    data = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if data.get("done"):
                    break
                delta = data.get("message", {}).get("content", "")
                if delta:
                    yield delta

    return StreamingResponse(iter_stream(), media_type="text/plain")

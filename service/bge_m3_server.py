"""Minimal BGE-M3 embedding server.

Run:
    HF_HOME=.hf-cache .venv-bge/bin/uvicorn service.bge_m3_server:app --port 18111

Endpoint:
    POST /embed {"texts": ["..."], "return_dense": true, "return_sparse": true}

Response shape matches excel_parser_rag.vector.BgeM3HttpClient.
"""

from __future__ import annotations

import os
from pathlib import Path
from threading import Lock
from typing import Any, Dict, List

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

_ROOT = Path(__file__).resolve().parent.parent
_CACHE = _ROOT / ".hf-cache"
os.environ.setdefault("HF_HOME", str(_CACHE))
os.environ.setdefault("TRANSFORMERS_CACHE", str(_CACHE / "transformers"))
os.environ.setdefault("HF_HUB_DISABLE_TELEMETRY", "1")
os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

MODEL_NAME = os.environ.get("BGE_M3_MODEL", "BAAI/bge-m3")
DEVICE = os.environ.get("BGE_M3_DEVICE", "cpu")
USE_FP16 = os.environ.get("BGE_M3_USE_FP16", "false").lower() in {"1", "true", "yes", "on"}
MAX_BATCH_SIZE = int(os.environ.get("BGE_M3_MAX_BATCH_SIZE", "16"))

app = FastAPI(title="bge-m3-server", version="1.0.0")

_model: Any = None
_model_lock = Lock()


class EmbedRequest(BaseModel):
    texts: List[str] = Field(default_factory=list)
    return_dense: bool = True
    return_sparse: bool = True


def _load_model() -> Any:
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                from FlagEmbedding import BGEM3FlagModel

                _CACHE.mkdir(parents=True, exist_ok=True)
                _model = BGEM3FlagModel(
                    MODEL_NAME,
                    use_fp16=USE_FP16,
                    devices=DEVICE,
                    cache_dir=str(_CACHE),
                )
    return _model


def _sparse_to_dict(value: Any) -> Dict[str, float]:
    if not value:
        return {}
    if isinstance(value, dict):
        return {str(k): float(v) for k, v in value.items()}
    # Some versions return a list of (token, weight) pairs.
    try:
        return {str(k): float(v) for k, v in value}
    except Exception:
        return {}


@app.get("/healthz")
def healthz() -> Dict[str, Any]:
    return {
        "status": "ok",
        "model": MODEL_NAME,
        "loaded": _model is not None,
        "device": DEVICE,
        "cache": str(_CACHE),
    }


@app.post("/embed")
def embed(req: EmbedRequest) -> Dict[str, Any]:
    texts = [str(t) for t in req.texts if str(t).strip()]
    if not texts:
        raise HTTPException(status_code=422, detail="texts must not be empty")
    if len(texts) > MAX_BATCH_SIZE:
        raise HTTPException(status_code=413, detail=f"max batch size is {MAX_BATCH_SIZE}")

    model = _load_model()
    out = model.encode(
        texts,
        return_dense=req.return_dense,
        return_sparse=req.return_sparse,
        return_colbert_vecs=False,
    )
    dense_vecs = out.get("dense_vecs")
    if dense_vecs is None:
        dense_vecs = [[] for _ in texts]
    lexical_weights = out.get("lexical_weights")
    if lexical_weights is None:
        lexical_weights = [{} for _ in texts]

    data = []
    for dense, sparse in zip(dense_vecs, lexical_weights):
        if hasattr(dense, "tolist"):
            dense = dense.tolist()
        data.append(
            {
                "dense": [float(v) for v in dense],
                "sparse": _sparse_to_dict(sparse),
            }
        )
    return {"model": MODEL_NAME, "data": data}

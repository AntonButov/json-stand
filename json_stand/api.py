from __future__ import annotations

import json
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile

from json_stand.pipeline import convert_data

app = FastAPI(title="json-stand", version="0.1.0")
MAX_BYTES = 10 * 1024 * 1024


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/convert")
async def convert(file: UploadFile = File(...)) -> dict:
    name = file.filename or "upload.json"
    if not name.lower().endswith(".json"):
        raise HTTPException(status_code=400, detail="expected a .json file")

    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="empty file")
    if len(raw) > MAX_BYTES:
        raise HTTPException(status_code=400, detail=f"file too large (max {MAX_BYTES} bytes)")

    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise HTTPException(status_code=400, detail=f"invalid JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="root must be a JSON object")

    with tempfile.TemporaryDirectory(prefix="json-stand-api-") as tmp:
        out = Path(tmp) / (Path(name).stem + ".xlsx")
        result = convert_data(payload, out, normalize=True, include_base64=True)
        return result.to_dict()

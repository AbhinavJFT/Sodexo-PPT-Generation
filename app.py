"""FastAPI backend for the PPT-from-reference generator."""
import asyncio
import logging
import os
import re
import uuid
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from groq import Groq
from rewrite import rewrite_deck

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(levelname)s %(message)s")
log = logging.getLogger("app")

BASE = Path(__file__).parent
TEMPLATES = BASE / "templates"
OUTPUTS = BASE / "outputs"

TEMPLATES.mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)

GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
GROQ_MODEL = os.environ.get("GROQ_MODEL", "llama-3.3-70b-versatile")
if not GROQ_API_KEY:
    raise SystemExit("GROQ_API_KEY not set (see .env)")

groq_client = Groq(api_key=GROQ_API_KEY)
app = FastAPI(title="PPT Generator")

SAFE_ID = re.compile(r"^[A-Za-z0-9_.\-]+\.pptx$")


@app.get("/", response_class=HTMLResponse)
async def index():
    return (BASE / "static" / "index.html").read_text()


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(400, "Only .pptx files accepted")
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", file.filename)
    tid = uuid.uuid4().hex[:8]
    dest = TEMPLATES / f"{tid}__{safe_name}"
    data = await file.read()
    if len(data) > 50 * 1024 * 1024:
        raise HTTPException(413, "File too large (50MB max)")
    dest.write_bytes(data)
    log.info("uploaded %s (%d bytes)", dest.name, len(data))
    return {"id": dest.name, "filename": file.filename, "size": dest.stat().st_size}


@app.get("/api/templates")
async def list_templates():
    items = []
    for p in sorted(TEMPLATES.glob("*.pptx"), key=lambda x: x.stat().st_mtime, reverse=True):
        items.append({
            "id": p.name,
            "filename": p.name.split("__", 1)[-1] if "__" in p.name else p.name,
            "size": p.stat().st_size,
        })
    return items


@app.delete("/api/templates/{template_id}")
async def delete_template(template_id: str):
    if not SAFE_ID.match(template_id):
        raise HTTPException(400, "bad id")
    p = TEMPLATES / template_id
    if not p.exists():
        raise HTTPException(404)
    p.unlink()
    return {"ok": True}


@app.post("/api/generate")
async def generate(template_id: str = Form(...), prompt: str = Form(...)):
    if not SAFE_ID.match(template_id):
        raise HTTPException(400, "bad template id")
    src = TEMPLATES / template_id
    if not src.exists():
        raise HTTPException(404, f"template not found: {template_id}")
    if len(prompt.strip()) < 3:
        raise HTTPException(400, "prompt too short")

    out_id = uuid.uuid4().hex[:8]
    out_path = OUTPUTS / f"{out_id}.pptx"
    log.info("generate: template=%s out=%s", template_id, out_path.name)

    try:
        stats = await asyncio.to_thread(
            rewrite_deck, groq_client, GROQ_MODEL, str(src.resolve()), str(out_path.resolve()), prompt
        )
    except Exception as e:
        log.exception("generation error")
        raise HTTPException(500, f"generation error: {type(e).__name__}: {e}")

    if not out_path.exists():
        raise HTTPException(500, "output file was not created")

    return {
        "id": out_path.name,
        "download_url": f"/api/download/{out_path.name}",
        **stats,
    }


@app.get("/api/download/{name}")
async def download(name: str):
    if not SAFE_ID.match(name):
        raise HTTPException(400, "bad id")
    p = OUTPUTS / name
    if not p.exists():
        raise HTTPException(404)
    return FileResponse(
        p, filename=name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

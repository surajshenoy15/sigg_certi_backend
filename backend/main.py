"""
SIGGRAPH BNMIT - Certificate Generator Backend
FastAPI service for generating personalized certificates and emailing them.
"""

import os
import io
import csv
import uuid
import json
import base64
import asyncio
import logging
import urllib.request
from pathlib import Path
from typing import List, Dict, Any
from contextlib import asynccontextmanager

import requests
import openpyxl
import cloudinary
import cloudinary.uploader

from dotenv import load_dotenv
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, EmailStr
from PIL import Image, ImageDraw, ImageFont


# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
load_dotenv()

# Cloudinary reads CLOUDINARY_URL automatically from .env / Render env
cloudinary.config(secure=True)

PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://sigg-certi-backend.onrender.com"
).rstrip("/")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("cert-gen")

BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "generated"
FONTS_DIR = BASE_DIR / "fonts"

for d in (UPLOADS_DIR, OUTPUT_DIR, FONTS_DIR):
    d.mkdir(exist_ok=True)

# In-memory job registry + JSON backup
JOBS: Dict[str, Dict[str, Any]] = {}


# ---------------------------------------------------------------------------
# Job persistence helpers
# ---------------------------------------------------------------------------
def _job_file(job_id: str) -> Path:
    return OUTPUT_DIR / job_id / "job.json"


def _save_job(job_id: str) -> None:
    if job_id not in JOBS:
        return

    path = _job_file(job_id)
    path.parent.mkdir(exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(JOBS[job_id], f, ensure_ascii=False, indent=2)


def _get_job(job_id: str) -> Dict[str, Any]:
    if job_id in JOBS:
        return JOBS[job_id]

    path = _job_file(job_id)

    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            job = json.load(f)

        JOBS[job_id] = job
        return job

    raise HTTPException(
        404,
        f"Job not found: {job_id}. Please upload the template and recipient file again."
    )


def _ensure_template_exists(job: Dict[str, Any]) -> Path:
    template_path = Path(job["template_path"])

    if template_path.exists():
        return template_path

    template_url = job.get("template_url")

    if not template_url:
        raise HTTPException(
            404,
            "Template file missing and no Cloudinary URL found."
        )

    template_path.parent.mkdir(exist_ok=True)

    try:
        urllib.request.urlretrieve(template_url, template_path)
    except Exception as e:
        raise HTTPException(
            500,
            f"Could not restore template from Cloudinary: {str(e)}"
        )

    return template_path


# ---------------------------------------------------------------------------
# Lifespan
# ---------------------------------------------------------------------------
@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("Certificate Generator API starting up")
    logger.info(f"Uploads dir: {UPLOADS_DIR}")
    logger.info(f"Output dir:  {OUTPUT_DIR}")
    yield
    logger.info("Shutting down")


app = FastAPI(
    title="SIGGRAPH BNMIT Certificate Generator",
    description="Generate and email personalized certificates for event attendees.",
    version="1.0.0",
    lifespan=lifespan,
)

ALLOWED_ORIGINS = os.getenv(
    "ALLOWED_ORIGINS",
    "http://localhost:5173,http://localhost:3000,http://127.0.0.1:5173,https://sigg-certi-frontend.vercel.app",
).split(",")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[o.strip() for o in ALLOWED_ORIGINS],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class Recipient(BaseModel):
    name: str
    email: EmailStr


class GenerateRequest(BaseModel):
    job_id: str
    name_x: int
    name_y: int
    font_size: int = 64
    font_color: str = "#5B3FD9"
    font_family: str = "default"


class SendEmailRequest(BaseModel):
    job_id: str
    subject: str
    body: str
    sender_name: str = "SIGGRAPH BNMIT"


# ---------------------------------------------------------------------------
# File parsing helpers
# ---------------------------------------------------------------------------
def _parse_recipients_from_csv(content: bytes) -> List[Dict[str, str]]:
    text = content.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))

    if not reader.fieldnames:
        raise HTTPException(400, "CSV has no header row.")

    name_keys = {"name", "full name", "student name", "participant", "participant name"}
    email_keys = {"email", "email id", "e-mail", "email address", "mail"}

    name_field = next(
        (f for f in reader.fieldnames if f and f.strip().lower() in name_keys),
        None
    )
    email_field = next(
        (f for f in reader.fieldnames if f and f.strip().lower() in email_keys),
        None
    )

    if not name_field or not email_field:
        raise HTTPException(
            400,
            f"CSV must contain 'name' and 'email' columns. Got: {reader.fieldnames}",
        )

    rows = []

    for i, row in enumerate(reader, start=2):
        name = (row.get(name_field) or "").strip()
        email = (row.get(email_field) or "").strip()

        if not name or not email:
            logger.warning(f"Row {i} skipped because name/email missing: {row}")
            continue

        rows.append({
            "name": name,
            "email": email,
        })

    return rows


def _parse_recipients_from_xlsx(content: bytes) -> List[Dict[str, str]]:
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    try:
        header = next(rows_iter)
    except StopIteration:
        raise HTTPException(400, "Spreadsheet is empty.")

    header_norm = [str(h).strip().lower() if h else "" for h in header]

    name_keys = {"name", "full name", "student name", "participant", "participant name"}
    email_keys = {"email", "email id", "e-mail", "email address", "mail"}

    try:
        name_idx = next(i for i, h in enumerate(header_norm) if h in name_keys)
        email_idx = next(i for i, h in enumerate(header_norm) if h in email_keys)
    except StopIteration:
        raise HTTPException(
            400,
            f"Spreadsheet must contain 'name' and 'email' columns. Got: {header}",
        )

    rows = []

    for r in rows_iter:
        if not r:
            continue

        name = str(r[name_idx]).strip() if r[name_idx] is not None else ""
        email = str(r[email_idx]).strip() if r[email_idx] is not None else ""

        if not name or not email:
            continue

        rows.append({
            "name": name,
            "email": email,
        })

    return rows


# ---------------------------------------------------------------------------
# Certificate rendering helpers
# ---------------------------------------------------------------------------
def _load_font(family: str, size: int) -> ImageFont.FreeTypeFont:
    candidates = []
    family = (family or "default").lower()

    if family == "default":
        candidates += [
            FONTS_DIR / "Poppins-Bold.ttf",
            FONTS_DIR / "DejaVuSans-Bold.ttf",
        ]
    elif family == "serif":
        candidates += [
            FONTS_DIR / "PlayfairDisplay-Bold.ttf",
            FONTS_DIR / "DejaVuSerif-Bold.ttf",
        ]
    elif family == "mono":
        candidates += [
            FONTS_DIR / "DejaVuSansMono-Bold.ttf",
        ]
    else:
        candidates.append(FONTS_DIR / family)

    candidates += [
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
        Path("/Library/Fonts/Arial Bold.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf"),
    ]

    for p in candidates:
        try:
            if p.exists():
                return ImageFont.truetype(str(p), size=size)
        except Exception:
            continue

    logger.warning("No TTF font found. Falling back to PIL default font.")
    return ImageFont.load_default()


def _hex_to_rgb(hex_color: str) -> tuple:
    h = hex_color.strip().lstrip("#")

    if len(h) == 3:
        h = "".join(c * 2 for c in h)

    if len(h) != 6:
        return (91, 63, 217)

    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _render_certificate(
    template_path: Path,
    name: str,
    name_x: int,
    name_y: int,
    font_size: int,
    font_color: str,
    font_family: str,
) -> Image.Image:
    img = Image.open(template_path).convert("RGB")
    draw = ImageDraw.Draw(img)
    font = _load_font(font_family, font_size)
    color = _hex_to_rgb(font_color)

    bbox = draw.textbbox((0, 0), name, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    x = name_x - text_w / 2
    y = name_y - text_h / 2

    draw.text((x, y), name, fill=color, font=font)

    return img


# ---------------------------------------------------------------------------
# Brevo email helpers
# ---------------------------------------------------------------------------
def _brevo_config() -> Dict[str, str]:
    return {
        "api_key": os.getenv("BREVO_API_KEY", ""),
        "sender_email": os.getenv("BREVO_SENDER_EMAIL", ""),
        "sender_name": os.getenv("BREVO_SENDER_NAME", "SIGGRAPH BNMIT"),
    }


def _send_one_email(
    to_email: str,
    to_name: str,
    subject: str,
    body: str,
    sender_name: str,
    attachment_path: Path,
) -> None:
    cfg = _brevo_config()

    if not cfg["api_key"]:
        raise Exception("BREVO_API_KEY is not configured.")

    if not cfg["sender_email"]:
        raise Exception("BREVO_SENDER_EMAIL is not configured.")

    personalized_body = body.replace("{{name}}", to_name).replace("{{Name}}", to_name)

    with open(attachment_path, "rb") as f:
        encoded_attachment = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "sender": {
            "name": cfg["sender_name"] or sender_name,
            "email": cfg["sender_email"],
        },
        "to": [
            {
                "email": to_email,
                "name": to_name,
            }
        ],
        "subject": subject,
        "textContent": personalized_body,
        "attachment": [
            {
                "content": encoded_attachment,
                "name": f"{to_name}_certificate.jpg",
            }
        ],
    }

    headers = {
        "accept": "application/json",
        "api-key": cfg["api_key"],
        "content-type": "application/json",
    }

    try:
        response = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            json=payload,
            headers=headers,
            timeout=30,
        )

        if response.status_code not in (200, 201, 202):
            raise Exception(
                f"Brevo failed with status {response.status_code}: {response.text}"
            )

    except requests.exceptions.RequestException as e:
        raise Exception(f"Brevo network/API error: {str(e)}") from e


async def _send_emails_task(job_id: str, subject: str, body: str, sender_name: str):
    job = _get_job(job_id)
    cfg = _brevo_config()

    if not cfg["api_key"]:
        job["email_status"] = "failed"
        job["email_error"] = "BREVO_API_KEY not configured."
        JOBS[job_id] = job
        _save_job(job_id)
        return

    if not cfg["sender_email"]:
        job["email_status"] = "failed"
        job["email_error"] = "BREVO_SENDER_EMAIL not configured."
        JOBS[job_id] = job
        _save_job(job_id)
        return

    job["email_status"] = "sending"
    job["sent"] = 0
    job["failed"] = 0
    job["email_log"] = []
    JOBS[job_id] = job
    _save_job(job_id)

    for cert in job.get("certificates", []):
        try:
            await asyncio.to_thread(
                _send_one_email,
                cert["email"],
                cert["name"],
                subject,
                body,
                sender_name,
                OUTPUT_DIR / job_id / cert["filename"],
            )

            job["sent"] += 1
            job["email_log"].append({
                "email": cert["email"],
                "status": "sent",
            })

            logger.info(f"[{job_id}] sent → {cert['email']}")

        except Exception as e:
            job["failed"] += 1
            job["email_log"].append({
                "email": cert["email"],
                "status": "failed",
                "error": str(e),
            })

            logger.exception(f"[{job_id}] failed → {cert['email']}: {e}")

        JOBS[job_id] = job
        _save_job(job_id)

    job["email_status"] = "completed"
    JOBS[job_id] = job
    _save_job(job_id)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.get("/")
async def root():
    return {
        "service": "Certificate Generator",
        "version": "1.0.0",
        "status": "ok",
    }


@app.get("/api/health")
async def health():
    brevo = _brevo_config()

    return {
        "status": "ok",
        "brevo_configured": bool(brevo["api_key"] and brevo["sender_email"]),
        "brevo_sender": brevo["sender_email"] if brevo["sender_email"] else None,
        "public_base_url": PUBLIC_BASE_URL,
    }


@app.get("/api/jobs")
async def list_jobs():
    return {
        "memory_count": len(JOBS),
        "memory_jobs": list(JOBS.keys()),
        "saved_jobs": [
            p.parent.name
            for p in OUTPUT_DIR.glob("*/job.json")
        ],
    }


@app.get("/api/routes")
async def list_routes():
    return [
        {
            "path": route.path,
            "methods": list(route.methods),
        }
        for route in app.routes
        if hasattr(route, "methods")
    ]


@app.post("/api/upload")
async def upload_files(
    template: UploadFile = File(..., description="Certificate template JPG/PNG"),
    recipients: UploadFile = File(..., description="CSV or XLSX with name and email"),
):
    job_id = uuid.uuid4().hex[:12]
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(exist_ok=True)

    tpl_ext = Path(template.filename or "").suffix.lower() or ".jpg"

    if tpl_ext not in {".jpg", ".jpeg", ".png"}:
        raise HTTPException(400, "Template must be JPG or PNG.")

    tpl_bytes = await template.read()

    tpl_path = UPLOADS_DIR / f"{job_id}_template{tpl_ext}"
    tpl_path.parent.mkdir(exist_ok=True)
    tpl_path.write_bytes(tpl_bytes)

    try:
        upload_result = cloudinary.uploader.upload(
            io.BytesIO(tpl_bytes),
            folder="siggraph_certificate_templates",
            public_id=f"{job_id}_template",
            resource_type="image",
            overwrite=True,
        )

        template_cloudinary_url = upload_result["secure_url"]
        template_public_id = upload_result["public_id"]

    except Exception as e:
        logger.exception(f"Cloudinary upload failed: {e}")
        raise HTTPException(
            500,
            f"Cloudinary upload failed: {str(e)}"
        )

    rec_bytes = await recipients.read()
    rec_name = (recipients.filename or "").lower()

    if rec_name.endswith(".csv"):
        recs = _parse_recipients_from_csv(rec_bytes)
    elif rec_name.endswith((".xlsx", ".xls")):
        recs = _parse_recipients_from_xlsx(rec_bytes)
    else:
        raise HTTPException(400, "Recipients file must be .csv or .xlsx.")

    if not recs:
        raise HTTPException(400, "No valid rows found in recipient list.")

    with Image.open(tpl_path) as im:
        width, height = im.size

    JOBS[job_id] = {
        "job_id": job_id,
        "template_path": str(tpl_path),
        "template_url": template_cloudinary_url,
        "template_public_id": template_public_id,
        "template_width": width,
        "template_height": height,
        "recipients": recs,
        "certificates": [],
        "email_status": "idle",
        "sent": 0,
        "failed": 0,
        "email_log": [],
    }

    _save_job(job_id)

    logger.info(f"Created job {job_id} with {len(recs)} recipients ({width}x{height})")

    return {
        "job_id": job_id,
        "recipient_count": len(recs),
        "recipients": recs[:5],
        "template_width": width,
        "template_height": height,
        "template_url": template_cloudinary_url,
    }


@app.get("/api/template/{job_id}")
async def get_template(job_id: str):
    job = _get_job(job_id)

    return {
        "template_url": job.get("template_url"),
        "template_width": job.get("template_width"),
        "template_height": job.get("template_height"),
    }


@app.post("/api/preview")
async def preview_certificate(req: GenerateRequest):
    job = _get_job(req.job_id)
    template_path = _ensure_template_exists(job)

    sample_name = job["recipients"][0]["name"] if job.get("recipients") else "Sample Name"

    img = _render_certificate(
        template_path,
        sample_name,
        req.name_x,
        req.name_y,
        req.font_size,
        req.font_color,
        req.font_family,
    )

    preview_path = OUTPUT_DIR / req.job_id / "preview.jpg"
    preview_path.parent.mkdir(exist_ok=True)
    img.save(preview_path, "JPEG", quality=92)

    job["config"] = req.dict()
    JOBS[req.job_id] = job
    _save_job(req.job_id)

    return {
        "preview_url": f"{PUBLIC_BASE_URL}/files/{req.job_id}/preview.jpg",
        "sample_name": sample_name,
    }


@app.post("/api/generate")
async def generate_certificates(req: GenerateRequest):
    job = _get_job(req.job_id)
    template_path = _ensure_template_exists(job)

    out_dir = OUTPUT_DIR / req.job_id
    out_dir.mkdir(exist_ok=True)

    certs = []

    for rec in job["recipients"]:
        img = _render_certificate(
            template_path,
            rec["name"],
            req.name_x,
            req.name_y,
            req.font_size,
            req.font_color,
            req.font_family,
        )

        safe = "".join(
            c if c.isalnum() or c in "-_ " else "_"
            for c in rec["name"]
        ).strip()

        filename = f"{safe or 'certificate'}.jpg"

        counter = 1
        while (out_dir / filename).exists():
            filename = f"{safe}_{counter}.jpg"
            counter += 1

        img.save(out_dir / filename, "JPEG", quality=92)

        certs.append({
            "name": rec["name"],
            "email": rec["email"],
            "filename": filename,
            "url": f"{PUBLIC_BASE_URL}/files/{req.job_id}/{filename}",
        })

    job["certificates"] = certs
    job["config"] = req.dict()
    job["email_status"] = "idle"
    job["sent"] = 0
    job["failed"] = 0
    job["email_log"] = []

    JOBS[req.job_id] = job
    _save_job(req.job_id)

    logger.info(f"[{req.job_id}] generated {len(certs)} certificates")

    return {
        "job_id": req.job_id,
        "count": len(certs),
        "certificates": certs,
    }


@app.post("/api/send")
async def send_emails(req: SendEmailRequest, background: BackgroundTasks):
    job = _get_job(req.job_id)

    if not job.get("certificates"):
        raise HTTPException(400, "Generate certificates first.")

    background.add_task(
        _send_emails_task,
        req.job_id,
        req.subject,
        req.body,
        req.sender_name,
    )

    return {
        "status": "queued",
        "job_id": req.job_id,
        "total": len(job["certificates"]),
    }


@app.get("/api/status/{job_id}")
async def get_status(job_id: str):
    job = _get_job(job_id)

    return {
        "job_id": job_id,
        "recipient_count": len(job.get("recipients", [])),
        "generated": len(job.get("certificates", [])),
        "email_status": job.get("email_status", "idle"),
        "sent": job.get("sent", 0),
        "failed": job.get("failed", 0),
        "email_log": job.get("email_log", []),
        "error": job.get("email_error"),
    }


@app.get("/api/download/{job_id}")
async def download_all(job_id: str):
    job = _get_job(job_id)

    import zipfile

    src_dir = OUTPUT_DIR / job_id
    zip_path = OUTPUT_DIR / f"{job_id}_certificates.zip"

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for cert in job.get("certificates", []):
            p = src_dir / cert["filename"]

            if p.exists():
                zf.write(p, cert["filename"])

    return FileResponse(
        zip_path,
        filename=f"certificates_{job_id}.zip",
        media_type="application/zip",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
    )
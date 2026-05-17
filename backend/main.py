"""
SIGGRAPH BNMIT - Certificate Generator Backend
FastAPI service for generating personalized certificates and emailing them.
"""
import os
import io
import csv
import uuid
import asyncio
import smtplib
import logging
from email.message import EmailMessage
from pathlib import Path
from typing import List, Optional, Dict, Any
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, EmailStr
from PIL import Image, ImageDraw, ImageFont
import openpyxl
from dotenv import load_dotenv
import cloudinary
import cloudinary.uploader
# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
load_dotenv()
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

# In-memory job registry. For production swap with Redis/DB.
JOBS: Dict[str, Dict[str, Any]] = {}


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

# CORS - allow your Vite dev server and anything else you configure
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

# Serve generated certs for download/preview
app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class Recipient(BaseModel):
    name: str
    email: EmailStr


class PreviewRequest(BaseModel):
    job_id: str
    name: str = "Sample Name"


class GenerateRequest(BaseModel):
    job_id: str
    name_x: int           # X coordinate (in pixels) where name text is centered
    name_y: int           # Y coordinate (in pixels) where name text baseline sits
    font_size: int = 64
    font_color: str = "#5B3FD9"   # hex
    font_family: str = "default"  # "default" | "serif" | "mono" | uploaded


class SendEmailRequest(BaseModel):
    job_id: str
    subject: str
    body: str             # Plain-text body. {{name}} placeholder supported.
    sender_name: str = "SIGGRAPH BNMIT"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _parse_recipients_from_csv(content: bytes) -> List[Recipient]:
    text = content.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    rows = []
    # normalize header lookup
    fieldnames = [(f or "").strip().lower() for f in (reader.fieldnames or [])]
    if not fieldnames:
        raise HTTPException(400, "CSV has no header row.")

    name_keys = {"name", "full name", "student name", "participant", "participant name"}
    email_keys = {"email", "email id", "e-mail", "email address", "mail"}

    name_field = next((f for f in reader.fieldnames if f.strip().lower() in name_keys), None)
    email_field = next((f for f in reader.fieldnames if f.strip().lower() in email_keys), None)

    if not name_field or not email_field:
        raise HTTPException(
            400,
            f"CSV must contain 'name' and 'email' columns. Got: {reader.fieldnames}",
        )

    for i, row in enumerate(reader, start=2):
        name = (row.get(name_field) or "").strip()
        email = (row.get(email_field) or "").strip()
        if not name or not email:
            logger.warning(f"Row {i} skipped (missing data): {row}")
            continue
        rows.append({"name": name, "email": email})
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
        name = (str(r[name_idx]).strip() if r[name_idx] is not None else "")
        email = (str(r[email_idx]).strip() if r[email_idx] is not None else "")
        if not name or not email:
            continue
        rows.append({"name": name, "email": email})
    return rows


def _load_font(family: str, size: int) -> ImageFont.FreeTypeFont:
    """Try to load a real TTF; fall back gracefully."""
    candidates = []
    family = (family or "default").lower()

    # Allow project-local fonts
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
        candidates += [FONTS_DIR / "DejaVuSansMono-Bold.ttf"]
    else:
        # Treat as a custom filename
        candidates.append(FONTS_DIR / family)

    # System fallbacks
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

    logger.warning("No TTF found, falling back to PIL default font (size ignored).")
    return ImageFont.load_default()


def _hex_to_rgb(hex_color: str) -> tuple:
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


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

    # Measure & center horizontally at name_x
    bbox = draw.textbbox((0, 0), name, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = name_x - text_w / 2
    y = name_y - text_h / 2

    draw.text((x, y), name, fill=color, font=font)
    return img


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------
def _smtp_config() -> Dict[str, Any]:
    return {
        "host": os.getenv("SMTP_HOST", "smtp.gmail.com"),
        "port": int(os.getenv("SMTP_PORT", "587")),
        "user": os.getenv("SMTP_USER", ""),
        "password": os.getenv("SMTP_PASSWORD", ""),
        "from_email": os.getenv("SMTP_FROM", os.getenv("SMTP_USER", "")),
        "use_tls": os.getenv("SMTP_USE_TLS", "true").lower() == "true",
    }


def _send_one_email(
    cfg: Dict[str, Any],
    to_email: str,
    to_name: str,
    subject: str,
    body: str,
    sender_name: str,
    attachment_path: Path,
) -> None:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{cfg['from_email']}>"
    msg["To"] = to_email
    msg.set_content(body.replace("{{name}}", to_name).replace("{{Name}}", to_name))

    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="image",
            subtype="jpeg",
            filename=f"{to_name}_certificate.jpg",
        )

    with smtplib.SMTP(cfg["host"], cfg["port"], timeout=30) as server:
        server.ehlo()
        if cfg["use_tls"]:
            server.starttls()
            server.ehlo()
        server.login(cfg["user"], cfg["password"])
        server.send_message(msg)


async def _send_emails_task(job_id: str, subject: str, body: str, sender_name: str):
    job = JOBS[job_id]
    cfg = _smtp_config()
    if not cfg["user"] or not cfg["password"]:
        job["email_status"] = "failed"
        job["email_error"] = "SMTP not configured. Set SMTP_USER and SMTP_PASSWORD in .env"
        return

    job["email_status"] = "sending"
    job["sent"] = 0
    job["failed"] = 0
    job["email_log"] = []

    for cert in job["certificates"]:
        try:
            await asyncio.to_thread(
                _send_one_email,
                cfg,
                cert["email"],
                cert["name"],
                subject,
                body,
                sender_name,
                OUTPUT_DIR / job_id / cert["filename"],
            )
            job["sent"] += 1
            job["email_log"].append({"email": cert["email"], "status": "sent"})
            logger.info(f"[{job_id}] sent → {cert['email']}")
        except Exception as e:
            job["failed"] += 1
            job["email_log"].append({"email": cert["email"], "status": "failed", "error": str(e)})
            logger.exception(f"[{job_id}] failed → {cert['email']}: {e}")

    job["email_status"] = "completed"


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.get("/")
async def root():
    return {"service": "Certificate Generator", "version": "1.0.0", "status": "ok"}


@app.get("/api/health")
async def health():
    cfg = _smtp_config()
    return {
        "status": "ok",
        "smtp_configured": bool(cfg["user"] and cfg["password"]),
        "smtp_user": cfg["user"] if cfg["user"] else None,
    }


@app.post("/api/upload")
async def upload_files(
    template: UploadFile = File(..., description="Certificate template (JPG/PNG)"),
    recipients: UploadFile = File(..., description="CSV or XLSX with name & email"),
):
    """Upload template + recipient list. Returns a job_id used by subsequent calls."""
    job_id = uuid.uuid4().hex[:12]
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(exist_ok=True)

    # validate template
    tpl_ext = Path(template.filename).suffix.lower() or ".jpg"
    if tpl_ext not in {".jpg", ".jpeg", ".png"}:
        raise HTTPException(400, "Template must be JPG or PNG.")

    # read template bytes
    tpl_bytes = await template.read()

    # save temporary local copy for PIL processing
    tpl_path = UPLOADS_DIR / f"{job_id}_template{tpl_ext}"
    tpl_path.write_bytes(tpl_bytes)

    # upload template to Cloudinary for permanent preview URL
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
        raise HTTPException(500, f"Cloudinary upload failed: {str(e)}")

    # parse recipients
    rec_bytes = await recipients.read()
    rec_name = (recipients.filename or "").lower()

    if rec_name.endswith(".csv"):
        recs = _parse_recipients_from_csv(rec_bytes)
    elif rec_name.endswith((".xlsx", ".xls")):
        recs = _parse_recipients_from_xlsx(rec_bytes)
    else:
        raise HTTPException(400, "Recipients file must be .csv or .xlsx")

    if not recs:
        raise HTTPException(400, "No valid rows found in recipient list.")

    # get image dimensions
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
    }

    logger.info(f"Created job {job_id} with {len(recs)} recipients ({width}x{height})")

    return {
        "job_id": job_id,
        "recipient_count": len(recs),
        "recipients": recs[:5],
        "template_width": width,
        "template_height": height,

        # frontend should use this directly in <img src="">
        "template_url": template_cloudinary_url,
    }

@app.get("/api/template/{job_id}")
async def get_template(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, "Job not found")

    job = JOBS[job_id]

    return {
        "template_url": job.get("template_url"),
        "template_width": job.get("template_width"),
        "template_height": job.get("template_height"),
    }

@app.post("/api/preview")
async def preview_certificate(req: GenerateRequest):
    """Render a single sample cert so the user can verify placement before bulk."""
    if req.job_id not in JOBS:
        raise HTTPException(404, "Job not found")

    job = JOBS[req.job_id]
    sample_name = job["recipients"][0]["name"] if job["recipients"] else "Sample Name"

    img = _render_certificate(
        Path(job["template_path"]),
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

    # Save the chosen config on the job
    job["config"] = req.dict()

    return {
    "preview_url": f"{PUBLIC_BASE_URL}/files/{req.job_id}/preview.jpg",
    "sample_name": sample_name,
}


@app.post("/api/generate")
async def generate_certificates(req: GenerateRequest):
    """Render one certificate per recipient."""
    if req.job_id not in JOBS:
        raise HTTPException(404, "Job not found")

    job = JOBS[req.job_id]
    out_dir = OUTPUT_DIR / req.job_id
    out_dir.mkdir(exist_ok=True)
    template_path = Path(job["template_path"])

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
        # sanitize filename
        safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in rec["name"]).strip()
        filename = f"{safe or 'certificate'}.jpg"
        # avoid collisions
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
    logger.info(f"[{req.job_id}] generated {len(certs)} certificates")
    return {"job_id": req.job_id, "count": len(certs), "certificates": certs}


@app.post("/api/send")
async def send_emails(req: SendEmailRequest, background: BackgroundTasks):
    """Email the generated certificates as attachments. Runs in background."""
    if req.job_id not in JOBS:
        raise HTTPException(404, "Job not found")
    job = JOBS[req.job_id]
    if not job.get("certificates"):
        raise HTTPException(400, "Generate certificates first.")

    background.add_task(_send_emails_task, req.job_id, req.subject, req.body, req.sender_name)
    return {"status": "queued", "job_id": req.job_id, "total": len(job["certificates"])}


@app.get("/api/status/{job_id}")
async def get_status(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, "Job not found")
    job = JOBS[job_id]
    return {
        "job_id": job_id,
        "recipient_count": len(job["recipients"]),
        "generated": len(job.get("certificates", [])),
        "email_status": job.get("email_status", "idle"),
        "sent": job.get("sent", 0),
        "failed": job.get("failed", 0),
        "email_log": job.get("email_log", []),
        "error": job.get("email_error"),
    }


@app.get("/api/download/{job_id}")
async def download_all(job_id: str):
    """Zip everything up for download."""
    if job_id not in JOBS:
        raise HTTPException(404, "Job not found")
    import zipfile
    src_dir = OUTPUT_DIR / job_id
    zip_path = OUTPUT_DIR / f"{job_id}_certificates.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for cert in JOBS[job_id].get("certificates", []):
            p = src_dir / cert["filename"]
            if p.exists():
                zf.write(p, cert["filename"])
    return FileResponse(zip_path, filename=f"certificates_{job_id}.zip", media_type="application/zip")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)

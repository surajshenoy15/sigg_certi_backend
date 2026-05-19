"""
SIGGRAPH BNMIT - Certificate + Attendance Backend
FastAPI service.

Flows:
  A) Events + QR attendance
     - Create event, import Google Forms CSV
     - Email QR invites (Brevo)
     - Admin scans QR -> check-in recorded
     - Export attendance CSV / feed straight into cert flow

  B) Certificates (existing)
     - Upload template + recipient list
     - Render personalized certs
     - Email as attachments
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
from typing import List, Dict, Any, Optional
from contextlib import asynccontextmanager

import requests
import openpyxl
import cloudinary
import cloudinary.uploader

from dotenv import load_dotenv
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, EmailStr, Field
from PIL import Image, ImageDraw, ImageFont

from event_store import EventStore
from qr_service import upload_qr, build_qr_png


# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
load_dotenv()

# Cloudinary reads CLOUDINARY_URL automatically from env
cloudinary.config(secure=True)

PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://sigg-certi-backend.onrender.com"
).rstrip("/")

PUBLIC_FRONTEND_URL = os.getenv(
    "PUBLIC_FRONTEND_URL",
    "https://sigg-certi-frontend.vercel.app"
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
EVENTS_DIR = BASE_DIR / "events"
QR_LOCAL_DIR = BASE_DIR / "qr_local"

for d in (UPLOADS_DIR, OUTPUT_DIR, FONTS_DIR, EVENTS_DIR, QR_LOCAL_DIR):
    d.mkdir(exist_ok=True)

events = EventStore(EVENTS_DIR)

# In-memory job registry + JSON backup on disk
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
    logger.info("Certificate + Attendance API starting up")
    logger.info(f"Uploads dir:        {UPLOADS_DIR}")
    logger.info(f"Output dir:         {OUTPUT_DIR}")
    logger.info(f"Events dir:         {EVENTS_DIR}")
    logger.info(f"Public API URL:     {PUBLIC_BASE_URL}")
    logger.info(f"Public Frontend:    {PUBLIC_FRONTEND_URL}")
    yield
    logger.info("Shutting down")


app = FastAPI(
    title="SIGGRAPH BNMIT · Events + Certificates",
    description="Generate certificates and run QR-based attendance for events.",
    version="2.0.0",
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

# Serve generated certs + locally-stored QR fallbacks
app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")
app.mount("/files-qr", StaticFiles(directory=QR_LOCAL_DIR), name="files-qr")


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


class EventCreate(BaseModel):
    name: str = Field(min_length=2, max_length=120)
    date: str = ""


class SendInvitesRequest(BaseModel):
    subject: str = "Your attendance QR code"
    body_paragraph: str = (
        "Here's your personal QR code for attendance. Please show this at the venue "
        "entrance so we can mark you present."
    )
    only_unsent: bool = True


class CheckInRequest(BaseModel):
    token: str


# ---------------------------------------------------------------------------
# File parsing helpers
# ---------------------------------------------------------------------------
def _parse_recipients_from_csv(content: bytes) -> List[Dict[str, str]]:
    text = content.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))

    if not reader.fieldnames:
        raise HTTPException(400, "CSV has no header row.")

    NAME_KEYS = {
        "name", "full name", "student name", "participant",
        "participant name", "your name", "name of the student",
    }
    EMAIL_KEYS = {
        "email", "email id", "e-mail", "email address", "mail", "your email",
    }
    PHONE_KEYS = {"phone", "phone number", "mobile", "mobile number", "contact"}
    USN_KEYS = {"usn", "roll number", "roll no", "register number"}

    def pick(keys):
        for fn in reader.fieldnames:
            if fn and fn.strip().lower() in keys:
                return fn
        return None

    name_field = pick(NAME_KEYS)
    email_field = pick(EMAIL_KEYS)
    phone_field = pick(PHONE_KEYS)
    usn_field = pick(USN_KEYS)

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

        rec = {"name": name, "email": email}

        if phone_field:
            rec["phone"] = (row.get(phone_field) or "").strip()
        if usn_field:
            rec["usn"] = (row.get(usn_field) or "").strip()

        rows.append(rec)

    return rows


def _parse_recipients_from_xlsx(content: bytes) -> List[Dict[str, str]]:
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    try:
        header = next(rows_iter)
    except StopIteration:
        raise HTTPException(400, "Spreadsheet is empty.")

    headers = [str(h).strip() if h is not None else "" for h in header]

    # Re-encode as CSV bytes and reuse the CSV parser
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(headers)

    for r in rows_iter:
        if r is None:
            continue
        w.writerow(["" if c is None else str(c) for c in r])

    return _parse_recipients_from_csv(out.getvalue().encode("utf-8"))


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


def brevo_ready() -> bool:
    cfg = _brevo_config()
    return bool(cfg["api_key"] and cfg["sender_email"])


def _send_brevo(
    to_email: str,
    to_name: str,
    subject: str,
    html_content: Optional[str] = None,
    text_content: Optional[str] = None,
    sender_name: Optional[str] = None,
    attachment: Optional[dict] = None,
) -> None:
    """
    Generic Brevo send.
    attachment = {"content": <base64>, "name": "file.jpg"}
    """
    cfg = _brevo_config()

    if not cfg["api_key"]:
        raise Exception("BREVO_API_KEY is not configured.")
    if not cfg["sender_email"]:
        raise Exception("BREVO_SENDER_EMAIL is not configured.")

    payload = {
        "sender": {
            "name": sender_name or cfg["sender_name"],
            "email": cfg["sender_email"],
        },
        "to": [{"email": to_email, "name": to_name}],
        "subject": subject,
    }

    if html_content:
        payload["htmlContent"] = html_content
    if text_content:
        payload["textContent"] = text_content
    if attachment:
        payload["attachment"] = [attachment]

    try:
        response = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            json=payload,
            headers={
                "accept": "application/json",
                "api-key": cfg["api_key"],
                "content-type": "application/json",
            },
            timeout=30,
        )

        if response.status_code not in (200, 201, 202):
            raise Exception(
                f"Brevo failed with status {response.status_code}: {response.text[:300]}"
            )

    except requests.exceptions.RequestException as e:
        raise Exception(f"Brevo network/API error: {str(e)}") from e


def _send_one_cert_email(
    to_email: str,
    to_name: str,
    subject: str,
    body: str,
    sender_name: str,
    attachment_path: Path,
) -> None:
    personalized_body = body.replace("{{name}}", to_name).replace("{{Name}}", to_name)

    with open(attachment_path, "rb") as f:
        encoded_attachment = base64.b64encode(f.read()).decode("utf-8")

    _send_brevo(
        to_email=to_email,
        to_name=to_name,
        subject=subject,
        text_content=personalized_body,
        sender_name=sender_name,
        attachment={
            "content": encoded_attachment,
            "name": f"{to_name}_certificate.jpg",
        },
    )


def _render_invite_html(
    name: str,
    event_name: str,
    event_date: str,
    qr_url: str,
    check_in_url: str,
    body_paragraph: str,
    sender_name: str,
) -> str:
    date_block = (
        f'<p style="margin:0 0 14px;font-size:14px;color:#2a283d;">'
        f'<strong>Date:</strong> {event_date}</p>'
        if event_date else ''
    )

    return f"""<!doctype html>
<html>
<body style="margin:0;padding:0;background:#f4f1ea;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;color:#0f0e1a;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#f4f1ea;padding:32px 16px;">
<tr>
<td align="center">

<table role="presentation" width="560" cellspacing="0" cellpadding="0" border="0" style="background:#ffffff;border:1px solid #0f0e1a;max-width:560px;width:100%;">
<tr>
<td style="padding:28px 32px 12px;border-bottom:2px solid #0f0e1a;">
<div style="font-size:10px;letter-spacing:0.2em;text-transform:uppercase;color:#5b3fd9;font-weight:600;">{sender_name}</div>
<div style="font-size:22px;font-weight:700;letter-spacing:-0.01em;margin-top:4px;">You're in — {event_name}</div>
</td>
</tr>

<tr>
<td style="padding:24px 32px 8px;">
<p style="margin:0 0 14px;font-size:15px;line-height:1.6;">Hi {name},</p>
<p style="margin:0 0 14px;font-size:15px;line-height:1.6;">{body_paragraph}</p>
{date_block}
</td>
</tr>

<tr>
<td align="center" style="padding:18px 32px 24px;">
<div style="border:1px solid #0f0e1a;padding:22px;background:#f8f6ef;">
<div style="font-size:18px;font-weight:700;margin-bottom:8px;">Your QR code is attached as a PDF</div>
<div style="font-size:14px;line-height:1.6;color:#2a283d;">
Please open the attached PDF and show the QR code at the venue entrance.
</div>
</div>
</td>
</tr>

<tr>
<td style="padding:0 32px 24px;font-size:13px;line-height:1.6;color:#2a283d;">
<p style="margin:0 0 8px;">
If the PDF does not open, use this check-in link from your phone at the venue:<br/>
<a href="{check_in_url}" style="color:#5b3fd9;word-break:break-all;">{check_in_url}</a>
</p>
</td>
</tr>

<tr>
<td style="padding:16px 32px;border-top:1px solid #0f0e1a;background:#eae5d8;font-size:11px;letter-spacing:0.12em;text-transform:uppercase;color:#2a283d;">
See you at the venue · {sender_name}
</td>
</tr>
</table>

</td>
</tr>
</table>
</body>
</html>"""

def _safe_filename(value: str) -> str:
    safe = "".join(
        c if c.isalnum() or c in "-_ " else "_"
        for c in str(value or "attendance_qr")
    ).strip()

    return safe or "attendance_qr"


def _draw_centered_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    y: int,
    page_w: int,
    font,
    fill,
) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    x = (page_w - text_w) // 2
    draw.text((x, y), text, fill=fill, font=font)


def _build_attendance_qr_pdf(
    name: str,
    event_name: str,
    event_date: str,
    check_in_url: str,
) -> bytes:
    """
    Builds a PDF attachment containing the attendee's QR code.
    Uses PIL only, so no new dependency is required.
    """

    page_w, page_h = 1240, 1754

    page = Image.new("RGB", (page_w, page_h), "white")
    draw = ImageDraw.Draw(page)

    title_font = _load_font("default", 58)
    heading_font = _load_font("default", 42)
    normal_font = _load_font("default", 28)
    small_font = _load_font("default", 22)

    ink = (15, 14, 26)
    purple = (91, 63, 217)
    muted = (72, 70, 90)
    light_bg = (244, 241, 234)

    # Background border
    draw.rectangle((50, 50, page_w - 50, page_h - 50), outline=ink, width=3)

    # Header
    draw.text((90, 90), "SIGGRAPH BNMIT", fill=purple, font=small_font)
    draw.text((90, 135), "Attendance QR Pass", fill=ink, font=title_font)
    draw.line((90, 235, page_w - 90, 235), fill=ink, width=3)

    # Event info block
    y = 300

    draw.text((90, y), "Event", fill=purple, font=small_font)
    y += 34
    draw.text((90, y), event_name, fill=ink, font=heading_font)
    y += 78

    draw.text((90, y), "Participant", fill=purple, font=small_font)
    y += 34
    draw.text((90, y), name, fill=ink, font=normal_font)
    y += 58

    if event_date:
        draw.text((90, y), "Date", fill=purple, font=small_font)
        y += 34
        draw.text((90, y), event_date, fill=ink, font=normal_font)
        y += 68

    # Instruction box
    draw.rectangle((90, y, page_w - 90, y + 90), fill=light_bg, outline=ink, width=2)
    draw.text(
        (120, y + 28),
        "Show this QR code at the venue entrance for attendance check-in.",
        fill=ink,
        font=small_font,
    )

    # QR code
    qr_png = build_qr_png(check_in_url, box_size=14, border=3)
    qr_img = Image.open(io.BytesIO(qr_png)).convert("RGB")
    qr_img = qr_img.resize((540, 540))

    qr_x = (page_w - 540) // 2
    qr_y = 760

    draw.rectangle(
        (qr_x - 36, qr_y - 36, qr_x + 540 + 36, qr_y + 540 + 36),
        outline=ink,
        width=3,
    )

    page.paste(qr_img, (qr_x, qr_y))

    _draw_centered_text(
        draw,
        "SCAN FOR ATTENDANCE",
        qr_y + 585,
        page_w,
        small_font,
        purple,
    )

    # Fallback URL
    link_y = qr_y + 660
    draw.text((90, link_y), "If QR scanning fails, open this link:", fill=ink, font=small_font)
    link_y += 40

    max_chars = 74
    for i in range(0, len(check_in_url), max_chars):
        draw.text((90, link_y), check_in_url[i:i + max_chars], fill=purple, font=small_font)
        link_y += 34

    # Footer
    draw.line((90, page_h - 150, page_w - 90, page_h - 150), fill=ink, width=2)
    draw.text(
        (90, page_h - 105),
        "See you at the venue · SIGGRAPH BNMIT",
        fill=muted,
        font=small_font,
    )

    out = io.BytesIO()
    page.save(out, format="PDF", resolution=150.0)
    return out.getvalue()
# ---------------------------------------------------------------------------
# Certificate email background task
# ---------------------------------------------------------------------------
async def _send_emails_task(job_id: str, subject: str, body: str, sender_name: str):
    job = _get_job(job_id)

    if not brevo_ready():
        job["email_status"] = "failed"
        job["email_error"] = "Brevo not configured (BREVO_API_KEY / BREVO_SENDER_EMAIL)."
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
                _send_one_cert_email,
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
# QR invite background task
# ---------------------------------------------------------------------------
async def _send_invites_task(
    event_id: str,
    subject: str,
    body_paragraph: str,
    only_unsent: bool,
):
    event = events.get_event(event_id)

    if not event:
        logger.error(f"Event not found: {event_id}")
        return

    if not brevo_ready():
        logger.error("Brevo not configured; aborting invite send")
        return

    sender = os.getenv("BREVO_SENDER_NAME", "SIGGRAPH BNMIT")
    invite_batch_id = uuid.uuid4().hex[:8]

    for reg in event["registrants"]:
        if only_unsent and reg.get("invite_sent"):
            continue

        try:
            check_in_url = (
                f"{PUBLIC_FRONTEND_URL}/scan"
                f"?event={event_id}"
                f"&token={reg['token']}"
            )

            # Optional: generate/store QR URL for backend record/debugging.
            # The email itself will use PDF attachment, not inline QR image.
            qr_public_id = f"{event_id}_{reg['id']}_{invite_batch_id}"

            try:
                qr_url = await asyncio.to_thread(
                    upload_qr,
                    check_in_url,
                    qr_public_id,
                    QR_LOCAL_DIR,
                    PUBLIC_BASE_URL,
                )
            except Exception as qr_error:
                logger.warning(
                    f"[{event_id}] QR URL store failed for {reg['email']}: {qr_error}"
                )
                qr_url = ""

            # Build PDF attachment containing QR code
            pdf_bytes = await asyncio.to_thread(
                _build_attendance_qr_pdf,
                reg["name"],
                event["name"],
                event.get("date", ""),
                check_in_url,
            )

            encoded_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
            safe_name = _safe_filename(reg["name"])

            html = _render_invite_html(
                name=reg["name"],
                event_name=event["name"],
                event_date=event.get("date", ""),
                qr_url=qr_url,
                check_in_url=check_in_url,
                body_paragraph=body_paragraph,
                sender_name=sender,
            )

            await asyncio.to_thread(
                _send_brevo,
                reg["email"],
                reg["name"],
                subject,
                html,
                None,
                sender,
                {
                    "content": encoded_pdf,
                    "name": f"{safe_name}_attendance_qr.pdf",
                },
            )

            events.mark_invite_sent(event_id, reg["id"], qr_url)

            logger.info(
                f"[{event_id}] invite sent with QR PDF → {reg['email']}"
            )

        except Exception as e:
            logger.exception(
                f"[{event_id}] invite failed → {reg.get('email')}: {e}"
            )
# ===========================================================================
# Routes — Health / Diagnostics
# ===========================================================================
@app.get("/")
async def root():
    return {
        "service": "Certificate Generator",
        "version": "2.0.0",
        "status": "ok",
    }


@app.get("/api/health")
async def health():
    brevo = _brevo_config()

    return {
        "status": "ok",
        "brevo_configured": brevo_ready(),
        "brevo_sender": brevo["sender_email"] if brevo["sender_email"] else None,
        "cloudinary_configured": bool(os.getenv("CLOUDINARY_URL")),
        "public_base_url": PUBLIC_BASE_URL,
        "public_frontend_url": PUBLIC_FRONTEND_URL,
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


# ===========================================================================
# Routes — Events
# ===========================================================================
@app.get("/api/events")
async def list_events():
    return events.list_events()


@app.post("/api/events")
async def create_event(payload: EventCreate):
    return events.create_event(payload.name, payload.date)


@app.get("/api/events/{event_id}")
async def get_event(event_id: str):
    e = events.get_event(event_id)

    if not e:
        raise HTTPException(404, "Event not found")

    return e


@app.delete("/api/events/{event_id}")
async def delete_event(event_id: str):
    if not events.delete_event(event_id):
        raise HTTPException(404, "Event not found")

    return {"ok": True}


@app.post("/api/events/{event_id}/registrants")
async def upload_registrants(
    event_id: str,
    file: UploadFile = File(..., description="CSV/XLSX from Google Forms"),
):
    if not events.get_event(event_id):
        raise HTTPException(404, "Event not found")

    name_lower = (file.filename or "").lower()
    content = await file.read()

    try:
        if name_lower.endswith(".csv"):
            rows = _parse_recipients_from_csv(content)
        elif name_lower.endswith((".xlsx", ".xls")):
            rows = _parse_recipients_from_xlsx(content)
        else:
            raise HTTPException(400, "Upload .csv or .xlsx")
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Parse failed")
        raise HTTPException(400, f"Could not parse file: {e}")

    if not rows:
        raise HTTPException(400, "No valid rows.")

    return events.add_registrants(event_id, rows)


@app.post("/api/events/{event_id}/send-invites")
async def send_invites(
    event_id: str,
    req: SendInvitesRequest,
    bg: BackgroundTasks,
):
    event = events.get_event(event_id)

    if not event:
        raise HTTPException(404, "Event not found")

    if not event["registrants"]:
        raise HTTPException(400, "No registrants in this event.")

    if not brevo_ready():
        raise HTTPException(
            400,
            "Brevo not configured. Set BREVO_API_KEY and BREVO_SENDER_EMAIL in env.",
        )

    pending = [
        r for r in event["registrants"]
        if not (req.only_unsent and r.get("invite_sent"))
    ]

    bg.add_task(_send_invites_task, event_id, req.subject, req.body_paragraph, req.only_unsent)

    return {"queued": len(pending), "total": len(event["registrants"])}


@app.post("/api/events/{event_id}/checkin")
async def checkin(event_id: str, req: CheckInRequest):
    result = events.check_in(event_id, req.token)

    if not result["ok"]:
        raise HTTPException(404, result.get("error", "not found"))

    return result


@app.get("/api/events/{event_id}/scan/{token}", response_class=HTMLResponse)
async def scan_get(event_id: str, token: str):
    """Fallback for raw QR scanners that just open the URL directly."""
    result = events.check_in(event_id, token)

    if not result["ok"]:
        return HTMLResponse(
            "<html><body style='font-family:sans-serif;padding:40px;text-align:center;background:#f4f1ea;'>"
            "<h2 style='color:#c8312e'>Invalid QR code</h2>"
            "<p>This code is not recognized for this event.</p></body></html>",
            status_code=404,
        )

    name = result["registrant"]["name"]
    when = result.get("checked_in_at", "")
    already = result.get("already")

    badge = "ALREADY CHECKED IN" if already else "CHECKED IN"
    color = "#c8a838" if already else "#1f7a4d"

    return HTMLResponse(
        f"<html><body style='font-family:sans-serif;padding:40px;text-align:center;background:#f4f1ea;'>"
        f"<div style='font-size:11px;letter-spacing:0.2em;text-transform:uppercase;color:{color};font-weight:700;'>{badge}</div>"
        f"<h1 style='font-size:32px;margin:8px 0 4px;color:#0f0e1a;'>{name}</h1>"
        f"<p style='color:#2a283d;font-family:monospace;font-size:12px;'>{when}</p>"
        f"</body></html>"
    )


@app.get("/api/events/{event_id}/attendance.csv")
async def attendance_csv(event_id: str):
    event = events.get_event(event_id)

    if not event:
        raise HTTPException(404, "Event not found")

    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(["name", "email"])

    for r in event["registrants"]:
        if r.get("checked_in"):
            w.writerow([r["name"], r["email"]])

    return Response(
        content=out.getvalue(),
        media_type="text/csv",
        headers={
            "Content-Disposition": f'attachment; filename="attendance_{event_id}.csv"'
        },
    )


@app.post("/api/events/{event_id}/use-for-certificates")
async def use_event_for_certificates(
    event_id: str,
    template: UploadFile = File(..., description="Certificate template JPG/PNG"),
):
    """Skip the CSV upload — uses the event's attended list as recipients."""
    event = events.get_event(event_id)

    if not event:
        raise HTTPException(404, "Event not found")

    attended = events.get_attended_rows(event_id)

    if not attended:
        raise HTTPException(400, "No one has checked in yet.")

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
        up = cloudinary.uploader.upload(
            io.BytesIO(tpl_bytes),
            folder="siggraph_certificate_templates",
            public_id=f"{job_id}_template",
            resource_type="image",
            overwrite=True,
        )
        template_url = up["secure_url"]
        template_public_id = up["public_id"]
    except Exception as e:
        logger.exception(f"Cloudinary upload failed: {e}")
        raise HTTPException(500, f"Cloudinary upload failed: {str(e)}")

    with Image.open(tpl_path) as im:
        width, height = im.size

    JOBS[job_id] = {
        "job_id": job_id,
        "template_path": str(tpl_path),
        "template_url": template_url,
        "template_public_id": template_public_id,
        "template_width": width,
        "template_height": height,
        "recipients": attended,
        "certificates": [],
        "email_status": "idle",
        "sent": 0,
        "failed": 0,
        "email_log": [],
        "source_event_id": event_id,
    }

    _save_job(job_id)

    logger.info(
        f"Created job {job_id} from event {event_id} with "
        f"{len(attended)} attendees ({width}x{height})"
    )

    return {
        "job_id": job_id,
        "recipient_count": len(attended),
        "recipients": attended[:5],
        "template_width": width,
        "template_height": height,
        "template_url": template_url,
    }

@app.get("/api/events/{event_id}/qr/{registrant_id}.png")
async def get_attendance_qr_image(event_id: str, registrant_id: str):
    """
    Public QR image endpoint for Gmail.
    Gmail can directly load this HTTPS image.
    """

    event = events.get_event(event_id)

    if not event:
        raise HTTPException(404, "Event not found")

    reg = None

    for r in event.get("registrants", []):
        if str(r.get("id")) == str(registrant_id):
            reg = r
            break

    if not reg:
        raise HTTPException(404, "Registrant not found")

    check_in_url = (
        f"{PUBLIC_FRONTEND_URL}/scan"
        f"?event={event_id}"
        f"&token={reg['token']}"
    )

    png = build_qr_png(check_in_url)

    return Response(
        content=png,
        media_type="image/png",
        headers={
            "Cache-Control": "public, max-age=31536000",
        },
    )

@app.post("/api/events/{event_id}/reset-invites")
async def reset_invites(event_id: str):
    event = events.get_event(event_id)

    if not event:
        raise HTTPException(404, "Event not found")

    reset_count = 0

    for reg in event.get("registrants", []):
        reg["invite_sent"] = False
        reg["invite_sent_at"] = None
        reg["qr_url"] = None
        reset_count += 1

    # Save updated event JSON file
    event_path = EVENTS_DIR / f"{event_id}.json"

    with open(event_path, "w", encoding="utf-8") as f:
        json.dump(event, f, ensure_ascii=False, indent=2)

    return {
        "ok": True,
        "event_id": event_id,
        "reset_count": reset_count,
        "message": "Invite records cleared. You can resend invites now.",
    }
# ===========================================================================
# Routes — Certificate flow
# ===========================================================================
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
        raise HTTPException(500, f"Cloudinary upload failed: {str(e)}")

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

    sample_name = (
        job["recipients"][0]["name"]
        if job.get("recipients") else "Sample Name"
    )

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
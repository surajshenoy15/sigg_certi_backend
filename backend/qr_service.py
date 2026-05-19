"""QR generation + Cloudinary upload."""

from __future__ import annotations

import io
import os
import re
import logging
from pathlib import Path

from dotenv import load_dotenv
import qrcode
from qrcode.constants import ERROR_CORRECT_M

logger = logging.getLogger("qr")

load_dotenv()


def safe_public_id(value: str) -> str:
    value = str(value or "qr")
    value = re.sub(r"[^a-zA-Z0-9_-]+", "_", value)
    return value.strip("_")[:120] or "qr"


def build_qr_png(payload: str, box_size: int = 12, border: int = 2) -> bytes:
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_M,
        box_size=box_size,
        border=border,
    )

    qr.add_data(payload)
    qr.make(fit=True)

    img = qr.make_image(
        fill_color="#0f0e1a",
        back_color="#ffffff",
    )

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def upload_qr(
    payload: str,
    public_id: str,
    fallback_dir: Path,
    public_base_url: str = "",
) -> str:
    """
    Generate QR and return a public image URL.

    Priority:
    1. Cloudinary image URL
    2. Backend static file URL: /files-qr/name.png
    """

    clean_id = safe_public_id(public_id)
    png = build_qr_png(payload)

    cloudinary_url = os.getenv("CLOUDINARY_URL", "")

    if cloudinary_url:
        try:
            import cloudinary
            import cloudinary.uploader

            cloudinary.config(secure=True)

            res = cloudinary.uploader.upload(
                io.BytesIO(png),
                public_id=clean_id,
                folder="siggraph_bnmit_qr",
                overwrite=True,
                resource_type="image",
                format="png",
            )

            url = res.get("secure_url") or res.get("url")

            if url:
                logger.info(f"QR uploaded to Cloudinary: {url}")
                return url

        except Exception as e:
            logger.exception(f"Cloudinary QR upload failed: {e}")

    fallback_dir.mkdir(parents=True, exist_ok=True)

    out = fallback_dir / f"{clean_id}.png"
    out.write_bytes(png)

    rel = f"/files-qr/{clean_id}.png"

    if public_base_url:
        return f"{public_base_url.rstrip()}{rel}"

    return rel
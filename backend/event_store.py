"""JSON-backed event store. No DB."""
from __future__ import annotations
import json
import secrets
import threading
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any

_LOCK = threading.RLock()


def _make_id(prefix: str, length: int = 10) -> str:
    return f"{prefix}_{secrets.token_urlsafe(length)[:length]}"


def _make_token() -> str:
    return secrets.token_urlsafe(16)


class EventStore:
    def __init__(self, root: Path):
        self.root = root
        self.root.mkdir(exist_ok=True)

    def _path(self, event_id: str) -> Path:
        return self.root / f"{event_id}.json"

    def _load(self, event_id: str) -> Optional[dict]:
        p = self._path(event_id)
        if not p.exists():
            return None
        with _LOCK:
            return json.loads(p.read_text(encoding="utf-8"))

    def _save(self, event: dict) -> None:
        with _LOCK:
            self._path(event["id"]).write_text(
                json.dumps(event, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

    def list_events(self) -> List[dict]:
        events = []
        for p in sorted(self.root.glob("event_*.json")):
            try:
                e = json.loads(p.read_text(encoding="utf-8"))
                events.append({
                    "id": e["id"],
                    "name": e["name"],
                    "date": e.get("date"),
                    "created_at": e.get("created_at"),
                    "registrants_count": len(e.get("registrants", [])),
                    "attended_count": sum(1 for r in e.get("registrants", []) if r.get("checked_in")),
                    "invites_sent": sum(1 for r in e.get("registrants", []) if r.get("invite_sent")),
                })
            except Exception:
                continue
        events.sort(key=lambda e: e.get("created_at", ""), reverse=True)
        return events

    def create_event(self, name: str, date: str = "") -> dict:
        eid = _make_id("event", 10)
        event = {
            "id": eid,
            "name": name.strip(),
            "date": date.strip(),
            "created_at": datetime.utcnow().isoformat() + "Z",
            "registrants": [],
        }
        self._save(event)
        return event

    def get_event(self, event_id: str) -> Optional[dict]:
        return self._load(event_id)

    def delete_event(self, event_id: str) -> bool:
        p = self._path(event_id)
        if not p.exists():
            return False
        p.unlink()
        return True

    def add_registrants(self, event_id: str, rows: List[Dict[str, str]]) -> Dict[str, Any]:
        event = self._load(event_id)
        if not event:
            raise KeyError(event_id)
        existing = {r["email"].lower() for r in event["registrants"]}
        added, skipped = [], []
        for row in rows:
            email = (row.get("email") or "").strip().lower()
            name = (row.get("name") or "").strip()
            if not email or not name:
                skipped.append({"row": row, "reason": "missing name/email"})
                continue
            if email in existing:
                skipped.append({"row": row, "reason": "duplicate email"})
                continue
            reg = {
                "id": _make_id("reg", 8),
                "name": name,
                "email": email,
                "phone": (row.get("phone") or "").strip(),
                "usn": (row.get("usn") or "").strip(),
                "token": _make_token(),
                "invite_sent": False,
                "invite_sent_at": None,
                "qr_url": None,
                "checked_in": False,
                "checked_in_at": None,
                "created_at": datetime.utcnow().isoformat() + "Z",
            }
            event["registrants"].append(reg)
            existing.add(email)
            added.append(reg)
        self._save(event)
        return {
            "added": len(added),
            "skipped": len(skipped),
            "skipped_details": skipped[:20],
            "total": len(event["registrants"]),
        }

    def check_in(self, event_id: str, token: str) -> Dict[str, Any]:
        event = self._load(event_id)
        if not event:
            return {"ok": False, "error": "event_not_found"}
        for r in event["registrants"]:
            if r["token"] == token:
                if r["checked_in"]:
                    return {
                        "ok": True,
                        "already": True,
                        "registrant": {"name": r["name"], "email": r["email"]},
                        "checked_in_at": r["checked_in_at"],
                    }
                r["checked_in"] = True
                r["checked_in_at"] = datetime.utcnow().isoformat() + "Z"
                self._save(event)
                return {
                    "ok": True,
                    "already": False,
                    "registrant": {"name": r["name"], "email": r["email"]},
                    "checked_in_at": r["checked_in_at"],
                }
        return {"ok": False, "error": "token_not_found"}

    def mark_invite_sent(self, event_id: str, registrant_id: str, qr_url: str) -> None:
        event = self._load(event_id)
        if not event:
            return
        for r in event["registrants"]:
            if r["id"] == registrant_id:
                r["invite_sent"] = True
                r["invite_sent_at"] = datetime.utcnow().isoformat() + "Z"
                r["qr_url"] = qr_url
                break
        self._save(event)

    def get_attended_rows(self, event_id: str) -> List[Dict[str, str]]:
        event = self._load(event_id)
        if not event:
            return []
        return [{"name": r["name"], "email": r["email"]} for r in event["registrants"] if r.get("checked_in")]
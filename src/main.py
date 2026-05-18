import json
import base64
import re
from datetime import datetime

from js import Uint8Array
from workers import Response, WorkerEntrypoint

from workbook_parser import parse_workbook

ALL_MODULES = ("pricing", "calendario", "biancheria", "pulizie", "rendiconto")


def as_json(payload, status=200):
    return Response(
        json.dumps(payload),
        status=status,
        headers={"content-type": "application/json; charset=utf-8"},
    )


def form_number(form, name, default):
    value = form.get(name)
    try:
        return int(value) if value not in (None, "") else default
    except (TypeError, ValueError):
        return default


def jwt_payload(request):
    token = request.headers.get("Cf-Access-Jwt-Assertion")
    if not token:
        return {}
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload.encode()).decode())
    except Exception:
        return {}


async def current_user(env, request):
    payload = jwt_payload(request)
    email = payload.get("email")
    if not email:
        return None
    return await env.DB.prepare(
        "SELECT id, email, display_name, role, is_active FROM users WHERE email = ?1"
    ).bind(email).first()


async def user_permissions(env, user):
    if user is None or not user.is_active:
        return None
    if user.role == "admin":
        return {"role": "admin", "apartments": None, "modules": None}
    apartment_rows = await env.DB.prepare(
        "SELECT apartment_id FROM apartment_permissions WHERE user_id = ?1"
    ).bind(user.id).all()
    module_rows = await env.DB.prepare(
        "SELECT module_key FROM module_permissions WHERE user_id = ?1"
    ).bind(user.id).all()
    return {
        "role": user.role,
        "apartments": {row.apartment_id for row in apartment_rows.results},
        "modules": {row.module_key for row in module_rows.results},
    }


def filter_state(state, permissions):
    if state is None or permissions is None:
        return None
    apartments = permissions["apartments"]
    if apartments is None:
        return state

    def keep_apartment(row):
        return row.get("apartment") in apartments or row.get("name") in apartments

    filtered = dict(state)
    for key in ("apartments", "bookings", "reservations", "deliverySchedule", "pricing", "cleaningSchedule", "critical"):
        filtered[key] = [row for row in state.get(key, []) if keep_apartment(row)]
    filtered["workbookSheets"] = [
        row for row in state.get("workbookSheets", []) if row.get("name") in apartments
    ]
    return filtered


def split_csv(value):
    return [item.strip() for item in str(value or "").split(",") if item.strip()]


async def sync_staff_users(env, rows, parsed):
    apartment_ids = [row.get("name") for row in parsed.get("workbookSheets", []) if row.get("name")]
    for row in rows:
        email = str(row.get("email") or "").strip().lower()
        if not email:
            continue
        display_name = str(row.get("code") or email).strip()
        # Cloudflare D1 accepts NULL, but the Python Worker bridge may turn
        # Python None into JavaScript undefined before binding. Keep optional
        # text fields as empty strings so uploads do not fail on blank cells.
        phone = str(row.get("phone") or "").strip()
        role = str(row.get("role") or "pulizie").strip()
        if role not in {"admin", "pulizie", "proprietario"}:
            role = "pulizie"
        is_active = 1 if str(row.get("active") or "").lower() in {"1", "true", "on", "si", "sì"} else 0

        await env.DB.prepare(
            """
            INSERT INTO users (email, display_name, phone, role, is_active)
            VALUES (?1, ?2, ?3, ?4, ?5)
            ON CONFLICT(email) DO UPDATE SET
              display_name = excluded.display_name,
              phone = excluded.phone,
              role = excluded.role,
              is_active = excluded.is_active
            """
        ).bind(email, display_name, phone, role, is_active).run()
        user = await env.DB.prepare("SELECT id FROM users WHERE email = ?1").bind(email).first()
        if user is None:
            continue

        await env.DB.prepare("DELETE FROM apartment_permissions WHERE user_id = ?1").bind(user.id).run()
        await env.DB.prepare("DELETE FROM module_permissions WHERE user_id = ?1").bind(user.id).run()

        if role == "admin":
            continue

        allowed_apartments = split_csv(row.get("apartments")) or apartment_ids
        allowed_modules = split_csv(row.get("modules")) or list(ALL_MODULES)
        for apartment_id in allowed_apartments:
            await env.DB.prepare(
                "INSERT OR IGNORE INTO apartment_permissions (user_id, apartment_id) VALUES (?1, ?2)"
            ).bind(user.id, apartment_id).run()
        for module_key in allowed_modules:
            if module_key not in ALL_MODULES:
                continue
            await env.DB.prepare(
                "INSERT OR IGNORE INTO module_permissions (user_id, module_key) VALUES (?1, ?2)"
            ).bind(user.id, module_key).run()


class Default(WorkerEntrypoint):
    async def fetch(self, request):
        try:
            url = request.url
            method = request.method

            if "/api/health" in url:
                return as_json({"ok": True, "service": "gestione-appartamenti"})

            if method == "GET" and "/api/current" in url:
                user = await current_user(self.env, request)
                permissions = await user_permissions(self.env, user)
                row = await self.env.DB.prepare(
                    "SELECT value FROM app_settings WHERE key = ?1"
                ).bind("current_state").first()
                if row is None:
                    return as_json({"state": None})
                state = json.loads(str(row.value))
                if not state.get("uploadedAt"):
                    upload = await self.env.DB.prepare(
                        "SELECT uploaded_at FROM workbook_uploads WHERE is_current = 1 ORDER BY uploaded_at DESC LIMIT 1"
                    ).first()
                    if upload is not None:
                        state["uploadedAt"] = upload.uploaded_at
                return as_json({"state": filter_state(state, permissions)})

            if method == "GET" and "/api/me" in url:
                user = await current_user(self.env, request)
                permissions = await user_permissions(self.env, user)
                if user is None or permissions is None:
                    return as_json({"authenticated": False}, 401)
                return as_json({
                    "authenticated": True,
                    "email": user.email,
                    "displayName": user.display_name,
                    "role": user.role,
                    "apartments": None if permissions["apartments"] is None else sorted(permissions["apartments"]),
                    "modules": None if permissions["modules"] is None else sorted(permissions["modules"]),
                })

            if method == "POST" and "/api/upload" in url:
                user = await current_user(self.env, request)
                permissions = await user_permissions(self.env, user)
                if permissions is None or permissions["role"] != "admin":
                    return as_json({"error": "Solo admin puo caricare il file XLS."}, 403)
                payload = await request.json()
                if hasattr(payload, "to_py"):
                    payload = payload.to_py()
                filename = str(payload.get("filename") or "")
                if not filename:
                    return as_json({"error": "Nome file mancante."}, 400)
                if not filename.lower().endswith(".xlsx"):
                    return as_json({"error": "Formato file non supportato: carica un file Excel .xlsx."}, 400)

                raw = base64.b64decode(str(payload.get("contentBase64") or ""))
                if not raw:
                    return as_json({"error": "Il file caricato e vuoto."}, 400)

                try:
                    settings = payload.get("settings") or {}
                except Exception:
                    settings = {}

                try:
                    deliveries = payload.get("deliveries") or []
                except Exception:
                    deliveries = []
                try:
                    staff_users = payload.get("staffUsers") or []
                except Exception:
                    staff_users = []

                parsed = parse_workbook(raw, settings=settings, deliveries=deliveries)
                await sync_staff_users(self.env, staff_users, parsed)
                now = datetime.utcnow()
                uploaded_at = now.isoformat(timespec="seconds")
                parsed["uploadedAt"] = uploaded_at
                safe_stamp = now.strftime("%Y%m%dT%H%M%SZ")
                safe_filename = re.sub(r"[^A-Za-z0-9._-]+", "_", filename).strip("_") or "workbook.xlsx"
                object_key = f"workbooks/{safe_stamp}-{safe_filename}"

                await self.env.UPLOADS.put(object_key, Uint8Array.new(raw))
                await self.env.DB.prepare(
                    "UPDATE workbook_uploads SET is_current = 0 WHERE is_current = 1"
                ).run()
                await self.env.DB.prepare(
                    "INSERT INTO workbook_uploads (object_key, original_name, uploaded_at, is_current) VALUES (?1, ?2, ?3, 1)"
                ).bind(object_key, filename, uploaded_at).run()
                await self.env.DB.prepare(
                    """
                    INSERT INTO app_settings (key, value, updated_at)
                    VALUES (?1, ?2, ?3)
                    ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
                    """
                ).bind("current_state", json.dumps(parsed), uploaded_at).run()
                return as_json(parsed)

            if method == "POST" and "/api/staff-users" in url:
                user = await current_user(self.env, request)
                permissions = await user_permissions(self.env, user)
                if permissions is None or permissions["role"] != "admin":
                    return as_json({"error": "Solo admin puo salvare gli utenti."}, 403)
                payload = await request.json()
                if hasattr(payload, "to_py"):
                    payload = payload.to_py()
                rows = payload.get("staffUsers") or []
                row = await self.env.DB.prepare(
                    "SELECT value FROM app_settings WHERE key = ?1"
                ).bind("current_state").first()
                parsed = json.loads(str(row.value)) if row is not None else {"workbookSheets": []}
                await sync_staff_users(self.env, rows, parsed)
                return as_json({"ok": True, "saved": len(rows)})

            return await self.env.ASSETS.fetch(request)
        except Exception as error:
            return as_json({"error": f"Errore worker: {error}"}, 500)

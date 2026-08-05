"""Authenticated Supabase session and trusted tenant context for PCS Beta."""

from __future__ import annotations

from dataclasses import dataclass
import json
import secrets
import threading
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from flask import g, session

from pcs_local_settings import supabase_configuration


class TenantAuthenticationError(RuntimeError):
    pass


@dataclass(frozen=True)
class TenantContext:
    tenant_id: str
    tenant_name: str
    tenant_slug: str
    role: str
    user_id: str
    email: str
    api_key: str
    access_token: str


_AUTH_SESSIONS: dict[str, dict] = {}
_AUTH_LOCK = threading.RLock()


def _json_request(url: str, *, headers: dict[str, str], payload=None, method: str = "GET"):
    body = json.dumps(payload).encode("utf-8") if payload is not None else None
    request = Request(url, data=body, headers=headers, method=method)
    try:
        with urlopen(request, timeout=30) as response:
            raw = response.read()
    except HTTPError as exc:
        try:
            detail = json.loads(exc.read().decode("utf-8"))
            message = detail.get("msg") or detail.get("message") or detail.get("error_description")
        except Exception:
            message = ""
        raise TenantAuthenticationError(message or "Supabase rejected the authentication request.") from exc
    except (URLError, TimeoutError, OSError) as exc:
        raise TenantAuthenticationError("PCS could not connect to the configured Supabase project.") from exc
    return json.loads(raw.decode("utf-8")) if raw else {}


def _auth_headers(api_key: str, access_token: str | None = None) -> dict[str, str]:
    headers = {"apikey": api_key, "Content-Type": "application/json", "Accept": "application/json"}
    if access_token:
        headers["Authorization"] = f"Bearer {access_token}"
    return headers


def _memberships(project_url: str, api_key: str, access_token: str) -> list[dict]:
    query = urlencode({
        "select": "tenant_id,role,tenant:tenant(id,name,slug,is_active)",
        "is_active": "eq.true",
    })
    result = _json_request(
        f"{project_url.rstrip('/')}/rest/v1/tenant_membership?{query}",
        headers=_auth_headers(api_key, access_token),
    )
    return result if isinstance(result, list) else []


def _refresh(record: dict) -> None:
    if float(record.get("expires_at") or 0) > time.time() + 60:
        return
    result = _json_request(
        f"{record['project_url']}/auth/v1/token?grant_type=refresh_token",
        headers=_auth_headers(record["api_key"]),
        payload={"refresh_token": record["refresh_token"]},
        method="POST",
    )
    record["access_token"] = result["access_token"]
    record["refresh_token"] = result.get("refresh_token") or record["refresh_token"]
    record["expires_at"] = time.time() + int(result.get("expires_in") or 3600)


def sign_in(email: object, password: object) -> TenantContext:
    project_url, api_key = supabase_configuration()
    clean_email = str(email or "").strip().casefold()
    clean_password = str(password or "")
    if not project_url or not api_key:
        raise TenantAuthenticationError("Configure the Supabase URL and publishable key in Settings.")
    if not clean_email or not clean_password:
        raise TenantAuthenticationError("Enter your email address and password.")
    result = _json_request(
        f"{project_url}/auth/v1/token?grant_type=password",
        headers=_auth_headers(api_key),
        payload={"email": clean_email, "password": clean_password},
        method="POST",
    )
    access_token = str(result.get("access_token") or "")
    user = result.get("user") or {}
    memberships = _memberships(project_url, api_key, access_token)
    if not memberships:
        raise TenantAuthenticationError("Your account does not have an active company membership.")
    membership = memberships[0]
    tenant = membership.get("tenant") or {}
    opaque_id = secrets.token_urlsafe(32)
    record = {
        "project_url": project_url,
        "api_key": api_key,
        "access_token": access_token,
        "refresh_token": result.get("refresh_token") or "",
        "expires_at": time.time() + int(result.get("expires_in") or 3600),
        "user_id": str(user.get("id") or ""),
        "email": str(user.get("email") or clean_email),
        "tenant_id": str(membership["tenant_id"]),
    }
    with _AUTH_LOCK:
        _AUTH_SESSIONS[opaque_id] = record
    session.clear()
    session["auth_session_id"] = opaque_id
    session["tenant_id"] = record["tenant_id"]
    return _context_from_record(record, membership)


def _context_from_record(record: dict, membership: dict) -> TenantContext:
    tenant = membership.get("tenant") or {}
    return TenantContext(
        tenant_id=str(membership["tenant_id"]),
        tenant_name=str(tenant.get("name") or "Company"),
        tenant_slug=str(tenant.get("slug") or ""),
        role=str(membership.get("role") or "viewer"),
        user_id=record["user_id"],
        email=record["email"],
        api_key=record["api_key"],
        access_token=record["access_token"],
    )


def current_tenant_context() -> TenantContext:
    cached = getattr(g, "tenant_context", None)
    if cached is not None:
        return cached
    opaque_id = str(session.get("auth_session_id") or "")
    with _AUTH_LOCK:
        record = _AUTH_SESSIONS.get(opaque_id)
        if record:
            _refresh(record)
    if not record:
        raise TenantAuthenticationError("Sign in to continue.")
    memberships = _memberships(record["project_url"], record["api_key"], record["access_token"])
    selected_id = str(session.get("tenant_id") or record["tenant_id"])
    membership = next(
        (item for item in memberships if str(item.get("tenant_id")) == selected_id),
        None,
    )
    if membership is None:
        sign_out()
        raise TenantAuthenticationError("Your company membership is no longer active.")
    context = _context_from_record(record, membership)
    g.tenant_context = context
    return context


def sign_out() -> None:
    opaque_id = str(session.get("auth_session_id") or "")
    with _AUTH_LOCK:
        _AUTH_SESSIONS.pop(opaque_id, None)
    session.clear()


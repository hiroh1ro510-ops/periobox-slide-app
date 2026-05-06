import base64
import hashlib
import os
import secrets
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urlencode, urlsplit, urlunsplit, parse_qsl

import requests
import streamlit as st


ALLOWED_ROLES = {"admin", "moderator"}
AUTH_STATE_DIR = Path(".auth_state")
AUTH_STATE_TTL_SECONDS = 600


def load_local_env(path: str = ".env") -> None:
    """Small .env loader so local setup does not need an extra dependency."""

    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            os.environ.setdefault(key, value)


def get_secret_value(key: str, default: str = "") -> str:
    """Read config from environment variables or Streamlit secrets."""

    value = os.getenv(key, "").strip()
    if value:
        return value
    try:
        value = st.secrets.get(key, default)
    except Exception:
        value = default
    return str(value).strip() if value is not None else default


@dataclass
class AuthUser:
    email: str
    role: str
    access_token: str
    refresh_token: str = ""

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"


class SupabaseAuth:
    """Supabase Auth + allowed_users gate kept separate from the Streamlit app."""

    def __init__(self, supabase_url: str, anon_key: str, redirect_url: str):
        self.supabase_url = supabase_url.rstrip("/")
        self.anon_key = anon_key
        self.redirect_url = redirect_url

    @classmethod
    def from_env(cls) -> "SupabaseAuth":
        load_local_env()
        supabase_url = get_secret_value("SUPABASE_URL")
        anon_key = get_secret_value("SUPABASE_ANON_KEY")
        redirect_url = get_secret_value("APP_BASE_URL", "http://localhost:8501/")
        if not supabase_url or not anon_key:
            raise RuntimeError(
                "SUPABASE_URL と SUPABASE_ANON_KEY が設定されていません。"
                "Netlifyまたはローカル環境変数に設定してください。"
            )
        return cls(supabase_url=supabase_url, anon_key=anon_key, redirect_url=redirect_url)

    def sign_in_url(self) -> str:
        verifier = secrets.token_urlsafe(64)
        challenge = _pkce_challenge(verifier)
        st.session_state["supabase_code_verifier"] = verifier
        _save_code_verifier("latest", verifier)
        redirect_url = _url_with_query(self.redirect_url, {"pb_verifier": verifier})
        params = {
            "provider": "google",
            "redirect_to": redirect_url,
            "code_challenge": challenge,
            "code_challenge_method": "s256",
        }
        return f"{self.supabase_url}/auth/v1/authorize?{urlencode(params)}"

    def exchange_code(self, code: str) -> Dict[str, Any]:
        verifier = (
            str(st.query_params.get("pb_verifier", "")).strip()
            or st.session_state.get("supabase_code_verifier", "")
            or _load_code_verifier("latest")
        )
        if not verifier:
            raise RuntimeError("ログイン検証情報が見つかりません。もう一度Googleログインを押してください。")

        response = requests.post(
            f"{self.supabase_url}/auth/v1/token?grant_type=pkce",
            headers=self._headers(),
            json={"auth_code": code, "code_verifier": verifier},
            timeout=20,
        )
        _raise_for_supabase(response, "Googleログインの確認に失敗しました")
        _delete_code_verifier("latest")
        return response.json()

    def refresh_session(self, refresh_token: str) -> Dict[str, Any]:
        response = requests.post(
            f"{self.supabase_url}/auth/v1/token?grant_type=refresh_token",
            headers=self._headers(),
            json={"refresh_token": refresh_token},
            timeout=20,
        )
        _raise_for_supabase(response, "ログイン状態の更新に失敗しました")
        return response.json()

    def allowed_user_for_email(self, email: str, access_token: str) -> Optional[Dict[str, Any]]:
        response = requests.get(
            f"{self.supabase_url}/rest/v1/allowed_users",
            headers=self._headers(access_token),
            params={"email": f"eq.{email}", "select": "email,role,created_at"},
            timeout=20,
        )
        _raise_for_supabase(response, "利用権限の確認に失敗しました")
        rows = response.json()
        return rows[0] if rows else None

    def list_allowed_users(self, access_token: str) -> List[Dict[str, Any]]:
        response = requests.get(
            f"{self.supabase_url}/rest/v1/allowed_users",
            headers=self._headers(access_token),
            params={"select": "email,role,created_at", "order": "created_at.desc"},
            timeout=20,
        )
        _raise_for_supabase(response, "許可ユーザー一覧の取得に失敗しました")
        return response.json()

    def add_allowed_user(self, email: str, role: str, access_token: str) -> None:
        role = role.strip()
        if role not in ALLOWED_ROLES:
            raise RuntimeError("role は admin または moderator を指定してください。")
        response = requests.post(
            f"{self.supabase_url}/rest/v1/allowed_users",
            headers={**self._headers(access_token), "Prefer": "return=minimal,resolution=merge-duplicates"},
            params={"on_conflict": "email"},
            json={"email": email.strip().lower(), "role": role},
            timeout=20,
        )
        _raise_for_supabase(response, "許可ユーザーの追加に失敗しました")

    def delete_allowed_user(self, email: str, access_token: str) -> None:
        response = requests.delete(
            f"{self.supabase_url}/rest/v1/allowed_users",
            headers=self._headers(access_token),
            params={"email": f"eq.{email.strip().lower()}"},
            timeout=20,
        )
        _raise_for_supabase(response, "許可ユーザーの削除に失敗しました")

    def _headers(self, access_token: Optional[str] = None) -> Dict[str, str]:
        headers = {
            "apikey": self.anon_key,
            "Authorization": f"Bearer {access_token or self.anon_key}",
            "Content-Type": "application/json",
        }
        return headers


def require_authorized_user() -> Optional[AuthUser]:
    """Render login / authorization screens. Return a user only when app access is allowed."""

    try:
        auth = SupabaseAuth.from_env()
    except RuntimeError as error:
        st.error(str(error))
        st.info("ローカルでは `.env` やターミナル、Netlifyでは Site settings > Environment variables に設定します。")
        return None

    with st.spinner("認証状態を確認しています..."):
        try:
            _capture_oauth_callback(auth)
            _refresh_if_needed(auth)
            session = st.session_state.get("supabase_session")
            if not session:
                _render_login(auth)
                return None

            token = session["access_token"]
            user = session.get("user", {})
            email = (user.get("email") or "").lower()
            if not email:
                st.error("Googleアカウントのメールアドレスを確認できませんでした。")
                _render_logout()
                return None

            allowed = auth.allowed_user_for_email(email, token)
            if not allowed:
                st.error("このアプリを利用する権限がありません")
                st.caption(f"ログイン中: {email}")
                _render_logout()
                return None

            role = allowed.get("role", "moderator")
            user_obj = AuthUser(email=email, role=role, access_token=token, refresh_token=session.get("refresh_token", ""))
            _render_authenticated_sidebar(user_obj, auth)
            return user_obj
        except Exception as error:
            st.error(str(error))
            _render_logout()
            return None


def render_admin_panel(user: AuthUser) -> None:
    if not user.is_admin:
        return

    auth = SupabaseAuth.from_env()
    with st.sidebar.expander("管理画面: 許可ユーザー", expanded=False):
        with st.form("add_allowed_user"):
            email = st.text_input("追加するメールアドレス")
            role = st.selectbox("role", ["moderator", "admin"])
            submitted = st.form_submit_button("追加 / 更新")
            if submitted:
                auth.add_allowed_user(email, role, user.access_token)
                st.success("許可ユーザーを追加しました。")
                st.rerun()

        st.divider()
        rows = auth.list_allowed_users(user.access_token)
        for row in rows:
            cols = st.columns([3, 1, 1])
            cols[0].caption(row["email"])
            cols[1].caption(row["role"])
            if cols[2].button("削除", key=f"delete-{row['email']}"):
                auth.delete_allowed_user(row["email"], user.access_token)
                st.rerun()


def _capture_oauth_callback(auth: SupabaseAuth) -> None:
    auth_error = st.query_params.get("error")
    if auth_error:
        description = st.query_params.get("error_description", "")
        raise RuntimeError(f"Googleログインに失敗しました: {auth_error} {description}")

    code = st.query_params.get("code")
    if not code or st.session_state.get("supabase_session"):
        return
    session = auth.exchange_code(str(code))
    session["expires_at"] = int(time.time()) + int(session.get("expires_in", 3600))
    st.session_state["supabase_session"] = session
    st.session_state.pop("supabase_code_verifier", None)
    st.query_params.clear()
    st.rerun()


def _refresh_if_needed(auth: SupabaseAuth) -> None:
    session = st.session_state.get("supabase_session")
    if not session:
        return
    expires_at = int(session.get("expires_at", 0))
    refresh_token = session.get("refresh_token", "")
    if not refresh_token or time.time() < expires_at - 120:
        return
    refreshed = auth.refresh_session(refresh_token)
    refreshed["expires_at"] = int(time.time()) + int(refreshed.get("expires_in", 3600))
    st.session_state["supabase_session"] = refreshed


def _render_login(auth: SupabaseAuth) -> None:
    st.info("Googleログインしてください。許可されたメールアドレスのユーザーだけがアプリを利用できます。")
    st.caption(f"接続先Supabase: {auth.supabase_url}")
    st.link_button("Googleでログイン", auth.sign_in_url(), type="primary")
    st.stop()


def _render_authenticated_sidebar(user: AuthUser, auth: SupabaseAuth) -> None:
    with st.sidebar:
        st.caption(f"ログイン中: {user.email}")
        st.caption(f"role: {user.role}")
        if st.button("ログアウト"):
            _clear_session()
            st.rerun()
    render_admin_panel(user)


def _render_logout() -> None:
    if st.button("ログアウト"):
        _clear_session()
        st.rerun()


def _clear_session() -> None:
    st.session_state.pop("supabase_session", None)
    st.session_state.pop("supabase_code_verifier", None)


def _pkce_challenge(verifier: str) -> str:
    digest = hashlib.sha256(verifier.encode("utf-8")).digest()
    return base64.urlsafe_b64encode(digest).decode("utf-8").rstrip("=")


def _url_with_query(url: str, extra_params: Dict[str, str]) -> str:
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query.update(extra_params)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))


def _save_code_verifier(state: str, verifier: str) -> None:
    if not state:
        return
    AUTH_STATE_DIR.mkdir(exist_ok=True)
    _cleanup_old_code_verifiers()
    (AUTH_STATE_DIR / f"{state}.txt").write_text(f"{int(time.time())}\n{verifier}", encoding="utf-8")


def _load_code_verifier(state: str) -> str:
    if not state:
        return ""
    path = AUTH_STATE_DIR / f"{state}.txt"
    if not path.exists():
        return ""
    try:
        created_at_text, verifier = path.read_text(encoding="utf-8").split("\n", 1)
        if time.time() - int(created_at_text) > AUTH_STATE_TTL_SECONDS:
            path.unlink(missing_ok=True)
            return ""
        return verifier.strip()
    except Exception:
        return ""


def _delete_code_verifier(state: str) -> None:
    if state:
        (AUTH_STATE_DIR / f"{state}.txt").unlink(missing_ok=True)


def _cleanup_old_code_verifiers() -> None:
    if not AUTH_STATE_DIR.exists():
        return
    for path in AUTH_STATE_DIR.glob("*.txt"):
        try:
            created_at_text = path.read_text(encoding="utf-8").split("\n", 1)[0]
            if time.time() - int(created_at_text) > AUTH_STATE_TTL_SECONDS:
                path.unlink(missing_ok=True)
        except Exception:
            path.unlink(missing_ok=True)


def _raise_for_supabase(response: requests.Response, message: str) -> None:
    if response.ok:
        return
    detail = response.text
    raise RuntimeError(f"{message}: {response.status_code} {detail}")

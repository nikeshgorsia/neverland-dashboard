"""
Password + email one-time-code (2FA) gate for restricted dashboard sections.

Usage:
    from auth_gate import require_auth

    with some_tab:
        if require_auth("budget_2027"):
            ... render restricted content ...
"""
import random
import time

import requests
import streamlit as st

from sharepoint_sync import get_app_token

ALLOWED_DOMAIN = "imagineneverland.com"
CODE_TTL_SECONDS = 10 * 60
MAX_ATTEMPTS = 5
RESEND_COOLDOWN_SECONDS = 30


def _normalize_email(email: str) -> str:
    return email.strip().lower()


def _allowed_emails(key_prefix: str) -> set:
    secret_key = f"{key_prefix.upper()}_ALLOWED_EMAILS"
    return {_normalize_email(e) for e in st.secrets.get(secret_key, [])}


def _is_allowed(email: str, key_prefix: str) -> bool:
    email = _normalize_email(email)
    if not email.endswith(f"@{ALLOWED_DOMAIN}"):
        return False
    return email in _allowed_emails(key_prefix)


def _send_code(email: str, code: str) -> None:
    """Send the verification code via Microsoft Graph (app-only), not SMTP.

    Many Microsoft 365 tenants disable legacy SMTP AUTH outright, so this uses
    the same app-only Graph token already used for SharePoint sync instead.
    Requires the Mail.Send application permission (admin-consented) on that
    Azure AD app registration.
    """
    sender = st.secrets.get("SMTP_FROM") or st.secrets["SMTP_USER"]
    token = get_app_token()

    body = {
        "message": {
            "subject": "Your Neverland Finance verification code",
            "body": {
                "contentType": "Text",
                "content": (
                    f"Your verification code is: {code}\n\n"
                    "This code expires in 10 minutes. If you didn't request this, "
                    "you can ignore this email."
                ),
            },
            "toRecipients": [{"emailAddress": {"address": email}}],
        },
        "saveToSentItems": "false",
    }
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
        timeout=15,
    )
    if resp.status_code >= 300:
        raise ValueError(f"Graph sendMail failed ({resp.status_code}): {resp.text[:300]}")


def _reset(key_prefix: str, stage_key: str) -> None:
    st.session_state[stage_key] = "credentials"
    for k in ("pending_email", "pending_code", "code_sent_at", "attempts"):
        st.session_state.pop(f"{key_prefix}_{k}", None)


def require_auth(key_prefix: str) -> bool:
    """Render a password + email-2FA gate in the current container.

    Returns True once the viewer has completed both steps this session.
    """
    session_key = f"{key_prefix}_authenticated"
    if st.session_state.get(session_key):
        return True

    stage_key = f"{key_prefix}_gate_stage"
    stage = st.session_state.get(stage_key, "credentials")

    st.markdown("#### 🔒 Restricted — sign in to view")

    if stage == "credentials":
        with st.form(f"{key_prefix}_credentials_form"):
            email = st.text_input("Outlook email (@imagineneverland.com)")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Continue")

        if submitted:
            if not _is_allowed(email, key_prefix):
                st.error("That email isn't authorized to view this page.")
            elif password != st.secrets.get(f"{key_prefix.upper()}_PASSWORD"):
                st.error("Incorrect password.")
            else:
                code = f"{random.randint(0, 999999):06d}"
                try:
                    _send_code(_normalize_email(email), code)
                except Exception as exc:
                    st.error(f"Couldn't send verification email: {exc}")
                else:
                    st.session_state[f"{key_prefix}_pending_email"] = _normalize_email(email)
                    st.session_state[f"{key_prefix}_pending_code"] = code
                    st.session_state[f"{key_prefix}_code_sent_at"] = time.time()
                    st.session_state[f"{key_prefix}_attempts"] = 0
                    st.session_state[stage_key] = "code"
                    st.rerun()
        return False

    # stage == "code"
    pending_email = st.session_state.get(f"{key_prefix}_pending_email", "")
    sent_at = st.session_state.get(f"{key_prefix}_code_sent_at", 0)

    if time.time() - sent_at > CODE_TTL_SECONDS:
        st.warning("That code expired. Please sign in again.")
        _reset(key_prefix, stage_key)
        st.rerun()

    st.caption(f"We emailed a 6-digit code to **{pending_email}**.")

    with st.form(f"{key_prefix}_code_form"):
        code_input = st.text_input("6-digit code", max_chars=6)
        col1, col2 = st.columns([1, 1])
        verify = col1.form_submit_button("Verify")
        resend = col2.form_submit_button("Resend code")

    if resend:
        if time.time() - sent_at < RESEND_COOLDOWN_SECONDS:
            st.info("Please wait a moment before requesting another code.")
        else:
            code = f"{random.randint(0, 999999):06d}"
            try:
                _send_code(pending_email, code)
            except Exception as exc:
                st.error(f"Couldn't resend verification email: {exc}")
            else:
                st.session_state[f"{key_prefix}_pending_code"] = code
                st.session_state[f"{key_prefix}_code_sent_at"] = time.time()
                st.session_state[f"{key_prefix}_attempts"] = 0
                st.success("New code sent.")

    if verify:
        attempts = st.session_state.get(f"{key_prefix}_attempts", 0)
        if attempts >= MAX_ATTEMPTS:
            st.error("Too many incorrect attempts. Please sign in again.")
            _reset(key_prefix, stage_key)
            st.rerun()
        elif code_input.strip() == st.session_state.get(f"{key_prefix}_pending_code"):
            st.session_state[session_key] = True
            st.session_state[f"{key_prefix}_email"] = pending_email
            for k in ("pending_code", "code_sent_at", "attempts"):
                st.session_state.pop(f"{key_prefix}_{k}", None)
            st.session_state.pop(stage_key, None)
            st.rerun()
        else:
            st.session_state[f"{key_prefix}_attempts"] = attempts + 1
            st.error("Incorrect code.")

    return False

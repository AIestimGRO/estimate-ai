"""User accounts and durable login sessions for the web application."""

from __future__ import annotations

import base64
import hashlib
import hmac
import secrets
import sqlite3
from dataclasses import dataclass


ROLE_ADMIN = "admin"
ROLE_SPECIALIST = "specialist"
SUPPORTED_ROLES = frozenset({ROLE_ADMIN, ROLE_SPECIALIST})
PASSWORD_ITERATIONS = 210_000


@dataclass(frozen=True)
class UserRecord:
    id: int
    full_name: str
    login: str
    role: str
    is_active: bool
    created_at: str
    updated_at: str


def count_users(connection: sqlite3.Connection) -> int:
    row = connection.execute("SELECT COUNT(*) AS n FROM app_users").fetchone()
    return int(row["n"])


def create_user(
    connection: sqlite3.Connection,
    *,
    full_name: str,
    login: str,
    password: str,
    role: str,
) -> int:
    normalized_name = _required_text(full_name, "Full name is required")
    normalized_login = normalize_login(login)
    normalized_role = _validate_role(role)
    password_hash = hash_password(password)
    try:
        with connection:
            cursor = connection.execute(
                """
                INSERT INTO app_users(full_name, login, password_hash, role)
                VALUES (?, ?, ?, ?)
                """,
                (
                    normalized_name,
                    normalized_login,
                    password_hash,
                    normalized_role,
                ),
            )
    except sqlite3.IntegrityError as error:
        raise ValueError("Login already exists") from error
    return int(cursor.lastrowid)


def list_users(connection: sqlite3.Connection) -> list[UserRecord]:
    rows = connection.execute(
        """
        SELECT id, full_name, login, role, is_active, created_at, updated_at
        FROM app_users
        ORDER BY CASE role WHEN 'admin' THEN 0 ELSE 1 END, full_name, id
        """
    ).fetchall()
    return [_user_record(row) for row in rows]


def get_user(connection: sqlite3.Connection, user_id: int) -> UserRecord | None:
    row = connection.execute(
        """
        SELECT id, full_name, login, role, is_active, created_at, updated_at
        FROM app_users
        WHERE id = ?
        """,
        (int(user_id),),
    ).fetchone()
    return None if row is None else _user_record(row)


def get_user_by_login(
    connection: sqlite3.Connection,
    login: str,
) -> UserRecord | None:
    normalized_login = normalize_login(login)
    row = connection.execute(
        """
        SELECT id, full_name, login, role, is_active, created_at, updated_at
        FROM app_users
        WHERE login = ?
        """,
        (normalized_login,),
    ).fetchone()
    return None if row is None else _user_record(row)


def authenticate_user(
    connection: sqlite3.Connection,
    login: str,
    password: str,
) -> UserRecord | None:
    try:
        normalized_login = normalize_login(login)
    except ValueError:
        return None
    row = connection.execute(
        """
        SELECT id, full_name, login, password_hash, role, is_active,
               created_at, updated_at
        FROM app_users
        WHERE login = ?
        """,
        (normalized_login,),
    ).fetchone()
    if row is None or not bool(row["is_active"]):
        return None
    if not verify_password(password, str(row["password_hash"])):
        return None
    return _user_record(row)


def set_user_active(
    connection: sqlite3.Connection,
    user_id: int,
    *,
    is_active: bool,
) -> bool:
    with connection:
        cursor = connection.execute(
            """
            UPDATE app_users
            SET is_active = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (1 if is_active else 0, int(user_id)),
        )
        if not is_active:
            connection.execute(
                "DELETE FROM user_sessions WHERE user_id = ?",
                (int(user_id),),
            )
    return cursor.rowcount > 0


def reset_user_password(
    connection: sqlite3.Connection,
    user_id: int,
    *,
    password: str,
) -> bool:
    password_hash = hash_password(password)
    with connection:
        cursor = connection.execute(
            """
            UPDATE app_users
            SET password_hash = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (password_hash, int(user_id)),
        )
        connection.execute(
            "DELETE FROM user_sessions WHERE user_id = ?",
            (int(user_id),),
        )
    return cursor.rowcount > 0


def create_session(
    connection: sqlite3.Connection,
    user_id: int,
    *,
    lifetime_days: int = 7,
) -> str:
    token = secrets.token_urlsafe(32)
    token_hash = _token_hash(token)
    days = max(1, min(90, int(lifetime_days)))
    with connection:
        connection.execute(
            "DELETE FROM user_sessions WHERE expires_at <= datetime('now')"
        )
        connection.execute(
            """
            INSERT INTO user_sessions(user_id, token_hash, expires_at)
            VALUES (?, ?, datetime('now', ?))
            """,
            (int(user_id), token_hash, f"+{days} days"),
        )
    return token


def user_from_session(
    connection: sqlite3.Connection,
    token: str,
) -> UserRecord | None:
    text = str(token or "").strip()
    if not text:
        return None
    row = connection.execute(
        """
        SELECT u.id, u.full_name, u.login, u.role, u.is_active,
               u.created_at, u.updated_at
        FROM user_sessions s
        JOIN app_users u ON u.id = s.user_id
        WHERE s.token_hash = ?
          AND s.expires_at > datetime('now')
          AND u.is_active = 1
        """,
        (_token_hash(text),),
    ).fetchone()
    if row is None:
        return None
    connection.execute(
        """
        UPDATE user_sessions
        SET last_seen_at = datetime('now')
        WHERE token_hash = ?
        """,
        (_token_hash(text),),
    )
    connection.commit()
    return _user_record(row)


def revoke_session(connection: sqlite3.Connection, token: str) -> None:
    text = str(token or "").strip()
    if not text:
        return
    with connection:
        connection.execute(
            "DELETE FROM user_sessions WHERE token_hash = ?",
            (_token_hash(text),),
        )


def normalize_login(value: object) -> str:
    text = "" if value is None else str(value).strip().casefold()
    if len(text) < 3 or len(text) > 100:
        raise ValueError("Login must contain 3 to 100 characters")
    if any(character.isspace() for character in text):
        raise ValueError("Login must not contain whitespace")
    return text


def hash_password(password: str) -> str:
    text = str(password or "")
    if len(text) < 8:
        raise ValueError("Password must contain at least 8 characters")
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        text.encode("utf-8"),
        salt,
        PASSWORD_ITERATIONS,
    )
    return "$".join(
        (
            "pbkdf2_sha256",
            str(PASSWORD_ITERATIONS),
            _b64encode(salt),
            _b64encode(digest),
        )
    )


def verify_password(password: str, encoded: str) -> bool:
    try:
        algorithm, raw_iterations, raw_salt, raw_digest = encoded.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        iterations = int(raw_iterations)
        salt = _b64decode(raw_salt)
        expected = _b64decode(raw_digest)
    except (TypeError, ValueError):
        return False
    actual = hashlib.pbkdf2_hmac(
        "sha256",
        str(password or "").encode("utf-8"),
        salt,
        iterations,
    )
    return hmac.compare_digest(actual, expected)


def _token_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def _b64encode(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def _b64decode(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding)


def _validate_role(value: object) -> str:
    role = str(value or "").strip().lower()
    if role not in SUPPORTED_ROLES:
        raise ValueError("Unsupported role")
    return role


def _required_text(value: object, message: str) -> str:
    text = "" if value is None else str(value).strip()
    if not text:
        raise ValueError(message)
    return text


def _user_record(row: sqlite3.Row) -> UserRecord:
    return UserRecord(
        id=int(row["id"]),
        full_name=str(row["full_name"]),
        login=str(row["login"]),
        role=str(row["role"]),
        is_active=bool(row["is_active"]),
        created_at=str(row["created_at"]),
        updated_at=str(row["updated_at"]),
    )

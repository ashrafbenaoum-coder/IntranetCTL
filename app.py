import csv
import json
import os
import re
import secrets
import shutil
import sqlite3
import subprocess
import sys
import time
import unicodedatacls

import zipfile
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path
from threading import Lock
from xml.etree import ElementTree as ET

BASE_DIR = Path(__file__).resolve().parent
DEPS_DIR = BASE_DIR / ".deps"
if DEPS_DIR.exists():
    sys.path.insert(0, str(DEPS_DIR))

from flask import (  # noqa: E402
    Flask,
    Response,
    abort,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash  # noqa: E402

DB_PATH = BASE_DIR / "data" / "intranet.db"
FAILED_LOGINS = {}
MAX_LOGIN_ATTEMPTS = 5
LOCK_SECONDS = 5 * 60
WINDOWS_DIALOG_LOCK = Lock()

SUPPORT_RE = re.compile(r"^\d{8}$")
EAN_RE = re.compile(r"^\d{8,14}$")
PRODUCT_RE = re.compile(r"^\d{1,6}$")
USERNAME_RE = re.compile(r"^[A-Za-z0-9_.-]{3,50}$")
XLSX_NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
PKG_REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}

DATE_COLUMN_ALIASES = {
    "date",
    "jour",
    "day",
    "movement_date",
    "date_mouvement",
    "control_date",
    "date_controle",
    "operation_date",
}
PREPARED_COLUMN_ALIASES = {
    "total_colis_prepare",
    "total_colis_prepares",
    "colis_prepares",
    "nb_colis_prepare",
    "nb_colis_prepares",
    "prepared_total",
    "total_prepared",
    "prepared",
    "total_uvc_prepare",
    "total_uvc_prepares",
    "uvc_prepares",
}
FIABILITE_COLUMN_ALIASES = {
    "taux_fiabilite",
    "fiabilite",
    "reliability_rate",
    "quality_rate",
    "taux_corrige",
    "taux_corrigee",
    "corrected_rate",
}
UVC_CONTROLE_COLUMN_ALIASES = {
    "uvc_controle",
    "uvc_control",
    "total_uvc_controle",
    "total_uvc_controles",
    "uvc_controles",
    "controle_uvc",
}
UVC_ECART_COLUMN_ALIASES = {
    "uvc_ecart",
    "uvc_ecarts",
    "total_uvc_ecart",
    "total_uvc_ecarts",
    "ecart_uvc",
}
DEMARQUE_COLUMN_ALIASES = {
    "type_de_demarque",
    "demarque_type",
    "type_demarque",
    "motif_demarque",
}
ANALYSE_DEMARQUE_TYPES = {"manquant", "surplus"}
ANALYSE_FIABILITE_OBJECTIVE = 99.55

PERMISSION_CATALOG = [
    ("home", "Home"),
    ("library", "Library"),
    ("settings", "Settings"),
    ("analyse", "Analyse"),
    ("donnes", "Donnes"),
    ("user_management", "Gestion Users"),
]
ALL_PERMISSION_IDS = {perm_id for perm_id, _label in PERMISSION_CATALOG}
DEFAULT_PERMISSIONS_BY_ROLE = {
    "admin": [perm_id for perm_id, _label in PERMISSION_CATALOG],
    "controller": ["home", "library", "settings"],
}

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("INTRANET_SECRET_KEY", secrets.token_urlsafe(64))
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Strict"
app.config["SESSION_COOKIE_SECURE"] = os.getenv("INTRANET_COOKIE_SECURE", "0") == "1"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=8)
app.config["MAX_CONTENT_LENGTH"] = 1 * 1024 * 1024


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db() -> None:
    db = get_db()
    db.executescript(
        """
        PRAGMA foreign_keys = ON;

        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK (role IN ('admin', 'controller')),
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            library_id INTEGER,
            support_number TEXT NOT NULL,
            ean_code TEXT NOT NULL,
            product_code TEXT NOT NULL,
            diff_plus INTEGER NOT NULL CHECK (diff_plus >= 0),
            diff_minus INTEGER NOT NULL CHECK (diff_minus >= 0),
            movement_date TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id),
            FOREIGN KEY (library_id) REFERENCES libraries(id)
        );

        CREATE TABLE IF NOT EXISTS libraries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS library_users (
            library_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            PRIMARY KEY (library_id, user_id),
            FOREIGN KEY (library_id) REFERENCES libraries(id) ON DELETE CASCADE,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS donnes_connections (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            connection_type TEXT NOT NULL CHECK (connection_type IN ('excel', 'odbc', 'access')),
            connection_value TEXT NOT NULL,
            config_json TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_movements_date ON movements(movement_date DESC);
        CREATE INDEX IF NOT EXISTS idx_movements_user ON movements(user_id);
        CREATE INDEX IF NOT EXISTS idx_library_users_user ON library_users(user_id);
        """
    )
    db.commit()

    columns = {row["name"] for row in db.execute("PRAGMA table_info(movements)").fetchall()}
    if "library_id" not in columns:
        db.execute("ALTER TABLE movements ADD COLUMN library_id INTEGER")
    db.execute("CREATE INDEX IF NOT EXISTS idx_movements_library ON movements(library_id)")

    user_columns = {row["name"] for row in db.execute("PRAGMA table_info(users)").fetchall()}
    if "is_blocked" not in user_columns:
        db.execute("ALTER TABLE users ADD COLUMN is_blocked INTEGER NOT NULL DEFAULT 0")
    if "permissions_json" not in user_columns:
        db.execute("ALTER TABLE users ADD COLUMN permissions_json TEXT NOT NULL DEFAULT '[]'")

    connection_columns = {row["name"] for row in db.execute("PRAGMA table_info(donnes_connections)").fetchall()}
    if connection_columns and "config_json" not in connection_columns:
        db.execute("ALTER TABLE donnes_connections ADD COLUMN config_json TEXT NOT NULL DEFAULT '{}'")
    rows = db.execute("SELECT id, role, permissions_json FROM users").fetchall()
    for row in rows:
        permissions = parse_permissions_json(row["permissions_json"], row["role"])
        db.execute(
            "UPDATE users SET permissions_json = ? WHERE id = ?",
            (json.dumps(permissions), row["id"]),
        )
    db.commit()


def seed_users() -> None:
    db = get_db()
    count_row = db.execute("SELECT COUNT(*) AS total FROM users").fetchone()
    if count_row and int(count_row["total"]) > 0:
        return
    now = datetime.now().isoformat(timespec="seconds")
    defaults = [
        ("admin", "test", "admin"),
        ("test1", "test", "controller"),
        ("test2", "test", "controller"),
        ("test3", "test", "controller"),
    ]
    for username, plain_password, role in defaults:
        row = db.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
        if row:
            continue
        db.execute(
            """
            INSERT INTO users (username, password_hash, role, created_at, is_blocked, permissions_json)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                username,
                generate_password_hash(plain_password),
                role,
                now,
                0,
                json.dumps(default_permissions_for_role(role)),
            ),
        )
    db.commit()


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def get_client_ip() -> str:
    forwarded = request.headers.get("X-Forwarded-For")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return request.remote_addr or "unknown"


def login_key(username: str) -> str:
    return f"{get_client_ip()}::{username.lower()}"


def is_locked(key: str) -> bool:
    entry = FAILED_LOGINS.get(key)
    if not entry:
        return False
    if entry["locked_until"] <= time.time():
        FAILED_LOGINS.pop(key, None)
        return False
    return True


def register_failed_login(key: str) -> None:
    entry = FAILED_LOGINS.get(key, {"count": 0, "locked_until": 0.0})
    entry["count"] += 1
    if entry["count"] >= MAX_LOGIN_ATTEMPTS:
        entry["locked_until"] = time.time() + LOCK_SECONDS
        entry["count"] = 0
    FAILED_LOGINS[key] = entry


def clear_failed_login(key: str) -> None:
    FAILED_LOGINS.pop(key, None)


def ensure_csrf() -> None:
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)


def require_csrf() -> None:
    token = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token")
    if token and secrets.compare_digest(token, session.get("csrf_token", "")):
        return
    abort(400, description="Invalid CSRF token.")


def default_permissions_for_role(role: str) -> list[str]:
    return list(DEFAULT_PERMISSIONS_BY_ROLE.get(role, ["home"]))


def parse_permissions_json(raw_permissions: str | None, role: str) -> list[str]:
    defaults = default_permissions_for_role(role)
    if not raw_permissions:
        return defaults
    try:
        payload = json.loads(raw_permissions)
    except (TypeError, json.JSONDecodeError):
        return defaults
    if not isinstance(payload, list):
        return defaults
    normalized: list[str] = []
    seen: set[str] = set()
    for raw in payload:
        permission = str(raw).strip().lower()
        if permission not in ALL_PERMISSION_IDS or permission in seen:
            continue
        seen.add(permission)
        normalized.append(permission)
    return normalized or defaults


def normalize_permissions_payload(raw_permissions, role: str) -> list[str]:
    if raw_permissions is None:
        return default_permissions_for_role(role)
    if not isinstance(raw_permissions, list):
        raise ValueError("permissions must be a list of permission ids.")

    normalized: list[str] = []
    seen: set[str] = set()
    for raw in raw_permissions:
        permission = str(raw).strip().lower()
        if permission not in ALL_PERMISSION_IDS or permission in seen:
            continue
        seen.add(permission)
        normalized.append(permission)

    if not normalized:
        raise ValueError("At least one valid permission is required.")
    return normalized


def permissions_catalog_payload() -> list[dict]:
    return [{"id": perm_id, "label": label} for perm_id, label in PERMISSION_CATALOG]


def validate_username(raw_username: str) -> str:
    username = (raw_username or "").strip()
    if not USERNAME_RE.fullmatch(username):
        raise ValueError(
            "Username must be 3-50 chars and contain only letters, digits, _, - or ."
        )
    return username


def validate_password(raw_password: str) -> str:
    password = str(raw_password or "")
    if len(password) < 6:
        raise ValueError("Password must contain at least 6 characters.")
    if len(password) > 120:
        raise ValueError("Password is too long.")
    return password


def parse_user_role(raw_role: str) -> str:
    role = str(raw_role or "").strip().lower()
    if role not in {"admin", "controller"}:
        raise ValueError("role must be admin or controller.")
    return role


def parse_bool(value, field_name: str) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value != 0
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"1", "true", "yes", "on"}:
            return True
        if normalized in {"0", "false", "no", "off"}:
            return False
    raise ValueError(f"{field_name} must be a boolean value.")


def count_admins(
    db: sqlite3.Connection,
    *,
    exclude_user_id: int | None = None,
    only_unblocked: bool = False,
) -> int:
    clauses = ["role = 'admin'"]
    params: list[int] = []
    if only_unblocked:
        clauses.append("is_blocked = 0")
    if exclude_user_id is not None:
        clauses.append("id <> ?")
        params.append(exclude_user_id)
    where_clause = " AND ".join(clauses)
    row = db.execute(f"SELECT COUNT(*) AS total FROM users WHERE {where_clause}", params).fetchone()
    return int(row["total"]) if row else 0


def serialize_user_row(row: sqlite3.Row | dict) -> dict:
    payload = dict(row)
    payload["is_blocked"] = bool(payload.get("is_blocked"))
    payload["permissions"] = parse_permissions_json(
        payload.get("permissions_json"),
        payload.get("role", "controller"),
    )
    payload.pop("permissions_json", None)
    return payload


def current_user() -> dict | None:
    user_id = session.get("user_id")
    if not user_id:
        return None
    db = get_db()
    row = db.execute(
        """
        SELECT id, username, role, is_blocked, permissions_json
        FROM users
        WHERE id = ?
        """,
        (user_id,),
    ).fetchone()
    if not row:
        return None

    user = dict(row)
    if int(user.get("is_blocked") or 0):
        session.clear()
        return None
    user["permissions"] = parse_permissions_json(user.get("permissions_json"), user["role"])
    user["is_blocked"] = bool(user.get("is_blocked"))
    user.pop("permissions_json", None)
    return user


def login_required(view_func):
    @wraps(view_func)
    def wrapped(*args, **kwargs):
        if g.user is None:
            return redirect(url_for("login"))
        return view_func(*args, **kwargs)

    return wrapped


def role_required(*roles):
    def decorator(view_func):
        @wraps(view_func)
        def wrapped(*args, **kwargs):
            if g.user is None:
                return redirect(url_for("login"))
            if g.user["role"] not in roles:
                abort(403)
            return view_func(*args, **kwargs)

        return wrapped

    return decorator


def parse_non_negative_int(raw: str, field_name: str) -> int:
    if raw is None or raw == "":
        return 0
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(f"{field_name} must be a valid integer.") from exc
    if value < 0:
        raise ValueError(f"{field_name} cannot be negative.")
    return value


def validate_movement_payload(payload: dict) -> dict:
    support_number = (payload.get("support_number") or "").strip()
    ean_code = (payload.get("ean_code") or "").strip()
    product_code = (payload.get("product_code") or "").strip()

    if not SUPPORT_RE.fullmatch(support_number):
        raise ValueError("N Support must contain exactly 8 digits.")
    if not EAN_RE.fullmatch(ean_code):
        raise ValueError("Code EAN must contain only digits (8 to 14).")
    if not PRODUCT_RE.fullmatch(product_code):
        raise ValueError("Code produit must contain 1 to 6 digits.")

    diff_plus = parse_non_negative_int(str(payload.get("diff_plus", "0")), "Nb colis ecart +")
    diff_minus = parse_non_negative_int(str(payload.get("diff_minus", "0")), "Nb colis ecart -")

    return {
        "support_number": support_number,
        "ean_code": ean_code,
        "product_code": product_code,
        "diff_plus": diff_plus,
        "diff_minus": diff_minus,
    }


def validate_library_name(raw_name: str) -> str:
    name = (raw_name or "").strip()
    if len(name) < 2:
        raise ValueError("Library name must contain at least 2 characters.")
    if len(name) > 80:
        raise ValueError("Library name is too long.")
    return name


def parse_month_filter(raw_month: str) -> str:
    month = (raw_month or "").strip()
    if not month:
        return ""
    try:
        return datetime.strptime(month, "%Y-%m").strftime("%Y-%m")
    except ValueError as exc:
        raise ValueError("Month must use YYYY-MM format.") from exc


def parse_date_filter(raw_date: str, field_name: str) -> str:
    date_value = (raw_date or "").strip()
    if not date_value:
        return ""
    try:
        return datetime.strptime(date_value, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"{field_name} must use YYYY-MM-DD format.") from exc


def normalize_column_name(raw_name) -> str:
    ascii_text = (
        unicodedata.normalize("NFKD", str(raw_name or ""))
        .encode("ascii", "ignore")
        .decode("ascii")
    )
    return re.sub(r"[^a-z0-9]+", "_", ascii_text.lower()).strip("_")


def find_source_column_name(rows: list[dict], aliases: set[str], token_hints: tuple[str, ...]) -> str | None:
    if not rows:
        return None

    seen: list[str] = []
    seen_set: set[str] = set()
    for row in rows[:500]:
        for key in row.keys():
            if key in seen_set:
                continue
            seen_set.add(key)
            seen.append(str(key))

    normalized = {key: normalize_column_name(key) for key in seen}
    for key, normalized_name in normalized.items():
        if normalized_name in aliases:
            return key
    for key, normalized_name in normalized.items():
        if all(token in normalized_name for token in token_hints):
            return key
    return None


def parse_numeric_value(raw_value) -> float | None:
    if raw_value is None:
        return None
    if isinstance(raw_value, bool):
        return None
    if isinstance(raw_value, (int, float)):
        return float(raw_value)

    text = str(raw_value).strip()
    if not text:
        return None
    text = text.replace("%", "").replace(" ", "")
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    elif "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def parse_percent_value(raw_value) -> float | None:
    parsed = parse_numeric_value(raw_value)
    if parsed is None:
        return None
    if 0 <= parsed <= 1.2:
        parsed *= 100
    if parsed < 0:
        return None
    return max(0.0, min(100.0, parsed))


def parse_int_value(raw_value) -> int | None:
    parsed = parse_numeric_value(raw_value)
    if parsed is None or parsed < 0:
        return None
    return int(round(parsed))


def normalize_source_day(raw_value) -> str:
    if raw_value is None:
        return ""
    if isinstance(raw_value, datetime):
        return raw_value.date().isoformat()

    if isinstance(raw_value, (int, float)):
        # Excel serial date support (1900 date system).
        parsed = float(raw_value)
        if 20000 <= parsed <= 65000:
            excel_origin = datetime(1899, 12, 30)
            return (excel_origin + timedelta(days=parsed)).date().isoformat()

    text = str(raw_value).strip()
    if not text:
        return ""
    if re.fullmatch(r"\d+(?:[.,]\d+)?", text):
        parsed = parse_numeric_value(text)
        if parsed is not None and 20000 <= parsed <= 65000:
            excel_origin = datetime(1899, 12, 30)
            return (excel_origin + timedelta(days=float(parsed))).date().isoformat()
    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    if len(text) >= 10:
        candidate = text[:10].replace("/", "-")
        try:
            return datetime.strptime(candidate, "%Y-%m-%d").date().isoformat()
        except ValueError:
            pass
    return ""


def xlsx_column_index(cell_ref: str) -> int:
    letters = "".join(char for char in str(cell_ref or "") if char.isalpha()).upper()
    if not letters:
        return 0
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def read_xlsx_rows(file_path: Path) -> list[dict]:
    with zipfile.ZipFile(file_path) as workbook_zip:
        file_names = set(workbook_zip.namelist())
        if "xl/workbook.xml" not in file_names:
            raise ValueError("Invalid Excel .xlsx file.")

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in file_names:
            shared_root = ET.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
            for si in shared_root.findall("x:si", XLSX_NS):
                chunks = [node.text or "" for node in si.findall(".//x:t", XLSX_NS)]
                shared_strings.append("".join(chunks))

        workbook_root = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
        sheets = workbook_root.findall("x:sheets/x:sheet", XLSX_NS)
        if not sheets:
            return []
        first_sheet = sheets[0]
        rel_id = first_sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if not rel_id:
            raise ValueError("Excel sheet relationship not found.")

        rels_root = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
        target_path = ""
        for rel in rels_root.findall("pr:Relationship", PKG_REL_NS):
            if rel.attrib.get("Id") == rel_id:
                target_path = rel.attrib.get("Target", "")
                break
        if not target_path:
            raise ValueError("Excel sheet target not found.")
        sheet_path = target_path.lstrip("/")
        if not sheet_path.startswith("xl/"):
            sheet_path = f"xl/{sheet_path}"
        if sheet_path not in file_names:
            raise ValueError("Excel sheet file not found.")

        sheet_root = ET.fromstring(workbook_zip.read(sheet_path))
        matrix: list[dict[int, str]] = []
        for row_node in sheet_root.findall(".//x:sheetData/x:row", XLSX_NS):
            row_values: dict[int, str] = {}
            for cell in row_node.findall("x:c", XLSX_NS):
                col_index = xlsx_column_index(cell.attrib.get("r", ""))
                if col_index <= 0:
                    continue
                cell_type = cell.attrib.get("t", "")
                value = ""
                if cell_type == "inlineStr":
                    chunks = [node.text or "" for node in cell.findall(".//x:is/x:t", XLSX_NS)]
                    value = "".join(chunks)
                else:
                    value_node = cell.find("x:v", XLSX_NS)
                    value = (value_node.text or "").strip() if value_node is not None else ""
                    if cell_type == "s":
                        try:
                            value = shared_strings[int(value)]
                        except (ValueError, IndexError):
                            pass
                    elif cell_type == "b":
                        value = "1" if value == "1" else "0"
                if value != "":
                    row_values[col_index] = value
            if row_values:
                matrix.append(row_values)

        if not matrix:
            return []

        max_col = max(max(row_values.keys()) for row_values in matrix)
        header_values = matrix[0]
        headers: list[str] = []
        used_headers: set[str] = set()
        for col in range(1, max_col + 1):
            base = str(header_values.get(col, "")).strip() or f"col_{col}"
            header = base
            suffix = 2
            while header in used_headers:
                header = f"{base}_{suffix}"
                suffix += 1
            used_headers.add(header)
            headers.append(header)

        rows: list[dict] = []
        for row_values in matrix[1:]:
            row_payload: dict[str, str] = {}
            has_data = False
            for col, header in enumerate(headers, start=1):
                value = row_values.get(col, "")
                if isinstance(value, str):
                    value = value.strip()
                row_payload[header] = value
                if str(value).strip():
                    has_data = True
            if has_data:
                rows.append(row_payload)
        return rows


def read_csv_rows(file_path: Path) -> list[dict]:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with file_path.open("r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                if not reader.fieldnames:
                    return []
                rows = []
                for row in reader:
                    rows.append({str(key): value for key, value in row.items()})
                return rows
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    if last_error:
        raise ValueError(f"CSV cannot be decoded: {last_error}") from last_error
    return []


def read_excel_source_rows(raw_path: str) -> list[dict]:
    file_path = Path(raw_path).expanduser()
    if not file_path.exists() or not file_path.is_file():
        raise ValueError("Excel source file not found.")
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return read_csv_rows(file_path)
    if suffix == ".xlsx":
        return read_xlsx_rows(file_path)
    if suffix == ".xls":
        raise ValueError("Legacy .xls is not supported for Analyse. Use .xlsx/.csv or ODBC.")
    raise ValueError("Excel source must be .xlsx, .xls or .csv.")


def quote_sql_identifier(raw_name: str) -> str:
    return "[" + str(raw_name).replace("]", "]]") + "]"


def import_pyodbc():
    try:
        import pyodbc  # type: ignore
    except ImportError as exc:  # pragma: no cover
        raise ValueError("pyodbc is not installed on server.") from exc
    return pyodbc


def pick_source_table(cursor) -> tuple[str, str]:
    table_rows = cursor.tables(tableType="TABLE").fetchall()
    if not table_rows:
        raise ValueError("No table found in source connection.")

    candidates: list[dict] = []
    for row in table_rows:
        table_name = str(getattr(row, "table_name", "") or "").strip()
        table_schema = str(getattr(row, "table_schem", "") or "").strip()
        if not table_name or table_name.lower().startswith("msys"):
            continue
        if table_schema:
            label = f"{table_schema}.{table_name}"
            ref = f"{quote_sql_identifier(table_schema)}.{quote_sql_identifier(table_name)}"
        else:
            label = table_name
            ref = quote_sql_identifier(table_name)
        candidates.append({"label": label, "ref": ref, "schema": table_schema, "table": table_name})
    if not candidates:
        raise ValueError("No usable table found in source connection.")

    best = candidates[0]
    best_score = -1
    for candidate in candidates:
        score = 0
        col_rows = cursor.columns(
            table=candidate["table"],
            schema=candidate["schema"] or None,
        ).fetchall()
        normalized_columns = {
            normalize_column_name(str(getattr(col, "column_name", "") or ""))
            for col in col_rows
        }
        if normalized_columns & DATE_COLUMN_ALIASES:
            score += 2
        if normalized_columns & PREPARED_COLUMN_ALIASES:
            score += 2
        if normalized_columns & FIABILITE_COLUMN_ALIASES:
            score += 1
        if score > best_score:
            best = candidate
            best_score = score
    return best["ref"], best["label"]


def odbc_rows_from_connection(connection_string: str, table_hint: str = "") -> tuple[list[dict], str]:
    pyodbc = import_pyodbc()
    try:
        conn = pyodbc.connect(connection_string, timeout=8)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"ODBC connection failed: {exc}") from exc

    try:
        cursor = conn.cursor()
        table_hint = (table_hint or "").strip()
        if table_hint:
            parts = [part.strip(" []") for part in table_hint.split(".") if part.strip(" []")]
            if not parts:
                raise ValueError("Invalid source table name.")
            table_label = ".".join(parts)
            table_ref = ".".join(quote_sql_identifier(part) for part in parts)
        else:
            table_ref, table_label = pick_source_table(cursor)
        cursor.execute(f"SELECT * FROM {table_ref}")
        columns = [col[0] for col in (cursor.description or [])]
        rows = []
        for fetched in cursor.fetchall():
            rows.append({columns[index]: fetched[index] for index in range(len(columns))})
        return rows, table_label
    finally:
        conn.close()


def load_saved_donnes_connection(db: sqlite3.Connection) -> dict | None:
    row = db.execute(
        """
        SELECT id, connection_type, connection_value, config_json, updated_at
        FROM donnes_connections
        WHERE id = 1
        """
    ).fetchone()
    if not row:
        return None
    raw_config = row["config_json"]
    config = {}
    if raw_config:
        try:
            parsed = json.loads(raw_config)
            if isinstance(parsed, dict):
                config = parsed
        except json.JSONDecodeError:
            config = {}
    return {
        "type": row["connection_type"],
        "value": row["connection_value"],
        "config": config,
        "updated_at": row["updated_at"],
    }


def save_donnes_connection(
    db: sqlite3.Connection, connection_type: str, connection_value: str, *, config: dict | None = None
) -> None:
    payload = config if isinstance(config, dict) else {}
    db.execute(
        """
        INSERT INTO donnes_connections (id, connection_type, connection_value, config_json, updated_at)
        VALUES (1, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            connection_type = excluded.connection_type,
            connection_value = excluded.connection_value,
            config_json = excluded.config_json,
            updated_at = excluded.updated_at
        """,
        (connection_type, connection_value, json.dumps(payload), now_iso()),
    )
    db.commit()


def load_source_rows_from_connection(saved_connection: dict) -> tuple[list[dict], dict]:
    connection_type = str(saved_connection.get("type") or "").strip().lower()
    connection_value = str(saved_connection.get("value") or "").strip()
    config = saved_connection.get("config")
    config = config if isinstance(config, dict) else {}

    if connection_type == "excel":
        rows = read_excel_source_rows(connection_value)
        file_path = Path(connection_value).expanduser()
        return rows, {"connection_type": "excel", "source_label": str(file_path.name)}

    if connection_type == "odbc":
        table_hint = str(config.get("table") or "").strip()
        rows, table_label = odbc_rows_from_connection(connection_value, table_hint)
        return rows, {"connection_type": "odbc", "source_label": table_label}

    if connection_type == "access":
        file_path = Path(connection_value).expanduser()
        if not file_path.exists() or not file_path.is_file():
            raise ValueError("Access file not found.")
        driver = str(config.get("driver") or "{Microsoft Access Driver (*.mdb, *.accdb)}").strip()
        access_conn = f"Driver={driver};DBQ={file_path};"
        table_hint = str(config.get("table") or "").strip()
        rows, table_label = odbc_rows_from_connection(access_conn, table_hint)
        return rows, {"connection_type": "access", "source_label": table_label}

    raise ValueError("Invalid Donnes connection type.")


def normalized_cell_token(raw_value) -> str:
    return normalize_column_name(str(raw_value or "").strip())


def extract_interception_metrics(rows: list[dict], date_from_filter: str, date_to_filter: str) -> dict:
    date_column = find_source_column_name(rows, DATE_COLUMN_ALIASES, ("date", "controle"))
    controle_column = find_source_column_name(rows, UVC_CONTROLE_COLUMN_ALIASES, ("uvc", "controle"))
    ecart_column = find_source_column_name(rows, UVC_ECART_COLUMN_ALIASES, ("uvc", "ecart"))
    demarque_column = find_source_column_name(rows, DEMARQUE_COLUMN_ALIASES, ("demarque",))

    missing = []
    if not date_column:
        missing.append("DATE CONTROLE")
    if not controle_column:
        missing.append("UVC CONTROLE")
    if not ecart_column:
        missing.append("UVC ECART")
    if not demarque_column:
        missing.append("Type de demarque")
    if missing:
        raise ValueError(
            "Source columns not detected: "
            + ", ".join(missing)
            + ". Expected headers like DATE CONTROLE, UVC CONTROLE, UVC ECART and Type de demarque."
        )

    start_date = datetime.strptime(date_from_filter, "%Y-%m-%d").date() if date_from_filter else None
    end_date = datetime.strptime(date_to_filter, "%Y-%m-%d").date() if date_to_filter else None

    daily: dict[str, dict] = {}
    for row in rows:
        day = normalize_source_day(row.get(date_column))
        if not day:
            continue

        day_obj = datetime.strptime(day, "%Y-%m-%d").date()
        if start_date and day_obj < start_date:
            continue
        if end_date and day_obj > end_date:
            continue

        controle = parse_int_value(row.get(controle_column)) or 0
        ecart = parse_int_value(row.get(ecart_column)) or 0
        demarque = normalized_cell_token(row.get(demarque_column))

        bucket = daily.setdefault(day, {"uvc_controle": 0, "uvc_ecart": 0, "uvc_ecart_all": 0, "rows_count": 0})
        bucket["uvc_controle"] += controle
        bucket["uvc_ecart_all"] += ecart
        if demarque in ANALYSE_DEMARQUE_TYPES:
            bucket["uvc_ecart"] += ecart
        bucket["rows_count"] += 1

    if not daily:
        raise ValueError("No source rows found in selected period. Check Donnes source and date filters.")

    total_uvc_controle = 0
    total_uvc_ecart = 0
    total_uvc_ecart_all = 0
    daily_fiabilite = []
    daily_ecarts = []
    daily_details = []
    monthly_buckets: dict[str, dict] = {}

    for day in sorted(daily):
        entry = daily[day]
        uvc_controle = int(entry["uvc_controle"])
        uvc_ecart = int(entry["uvc_ecart"])
        uvc_ecart_all = int(entry["uvc_ecart_all"])
        uvc_livre = uvc_controle - uvc_ecart
        uvc_livre_dem = uvc_controle - uvc_ecart_all
        taux_avec_dem = ((uvc_livre_dem / uvc_controle) * 100.0) if uvc_controle > 0 else 0.0
        taux_avec_dem = max(0.0, min(100.0, taux_avec_dem))
        taux_corrige = ((uvc_livre / uvc_controle) * 100.0) if uvc_controle > 0 else 0.0
        taux_corrige = max(0.0, min(100.0, taux_corrige))
        taux_ecarts = max(0.0, min(100.0, 100.0 - taux_corrige))
        ecart_objectif = ANALYSE_FIABILITE_OBJECTIVE - taux_corrige

        total_uvc_controle += uvc_controle
        total_uvc_ecart += uvc_ecart
        total_uvc_ecart_all += uvc_ecart_all
        daily_fiabilite.append({"day": day, "value": round(taux_corrige, 4)})
        daily_ecarts.append({"day": day, "value": round(taux_ecarts, 4)})
        daily_details.append(
            {
                "day": day,
                "month": day[:7],
                "uvc_controle": uvc_controle,
                "uvc_ecart": uvc_ecart,
                "uvc_ecart_all": uvc_ecart_all,
                "uvc_livre": uvc_livre,
                "uvc_livre_dem": uvc_livre_dem,
                "rows_count": int(entry["rows_count"]),
                "taux_avec_dem": round(taux_avec_dem, 4),
                "taux_corrige": round(taux_corrige, 4),
                "taux_ecarts": round(taux_ecarts, 4),
                "ecart_objectif": round(ecart_objectif, 4),
            }
        )

        month_key = day[:7]
        month_entry = monthly_buckets.setdefault(
            month_key, {"uvc_controle": 0, "uvc_ecart": 0, "uvc_ecart_all": 0, "rows_count": 0}
        )
        month_entry["uvc_controle"] += uvc_controle
        month_entry["uvc_ecart"] += uvc_ecart
        month_entry["uvc_ecart_all"] += uvc_ecart_all
        month_entry["rows_count"] += int(entry["rows_count"])

    if total_uvc_controle <= 0:
        raise ValueError("UVC CONTROLE is empty in source. Interception.xlsx must contain daily totals.")

    total_uvc_livre = total_uvc_controle - total_uvc_ecart
    total_uvc_livre_dem = total_uvc_controle - total_uvc_ecart_all
    taux_avec_dem_global = (total_uvc_livre_dem / total_uvc_controle) * 100.0
    taux_avec_dem_global = max(0.0, min(100.0, taux_avec_dem_global))
    taux_corrige_global = (total_uvc_livre / total_uvc_controle) * 100.0
    taux_corrige_global = max(0.0, min(100.0, taux_corrige_global))
    taux_ecarts_global = max(0.0, min(100.0, 100.0 - taux_corrige_global))

    monthly_fiabilite = []
    monthly_ecarts = []
    monthly_details = []
    for month_key in sorted(monthly_buckets):
        entry = monthly_buckets[month_key]
        uvc_controle = int(entry["uvc_controle"])
        uvc_ecart = int(entry["uvc_ecart"])
        uvc_ecart_all = int(entry["uvc_ecart_all"])
        uvc_livre = uvc_controle - uvc_ecart
        uvc_livre_dem = uvc_controle - uvc_ecart_all
        taux_avec_dem = ((uvc_livre_dem / uvc_controle) * 100.0) if uvc_controle > 0 else 0.0
        taux_avec_dem = max(0.0, min(100.0, taux_avec_dem))
        taux_corrige = ((uvc_livre / uvc_controle) * 100.0) if uvc_controle > 0 else 0.0
        taux_corrige = max(0.0, min(100.0, taux_corrige))
        taux_ecarts = max(0.0, min(100.0, 100.0 - taux_corrige))
        ecart_objectif = ANALYSE_FIABILITE_OBJECTIVE - taux_corrige

        monthly_fiabilite.append({"month": month_key, "value": round(taux_corrige, 4)})
        monthly_ecarts.append({"month": month_key, "value": round(taux_ecarts, 4)})
        monthly_details.append(
            {
                "month": month_key,
                "uvc_controle": uvc_controle,
                "uvc_ecart": uvc_ecart,
                "uvc_ecart_all": uvc_ecart_all,
                "uvc_livre": uvc_livre,
                "uvc_livre_dem": uvc_livre_dem,
                "rows_count": int(entry["rows_count"]),
                "taux_avec_dem": round(taux_avec_dem, 4),
                "taux_corrige": round(taux_corrige, 4),
                "taux_ecarts": round(taux_ecarts, 4),
                "ecart_objectif": round(ecart_objectif, 4),
            }
        )

    return {
        "daily_fiabilite": daily_fiabilite,
        "daily_ecarts": daily_ecarts,
        "daily_details": daily_details,
        "monthly_fiabilite": monthly_fiabilite,
        "monthly_ecarts": monthly_ecarts,
        "monthly_details": monthly_details,
        "available_months": [entry["month"] for entry in monthly_details],
        "totals": {
            "total_uvc_controle": total_uvc_controle,
            "total_uvc_ecart": total_uvc_ecart,
            "total_uvc_ecart_all": total_uvc_ecart_all,
            "total_uvc_livre": total_uvc_livre,
            "total_uvc_livre_dem": total_uvc_livre_dem,
            "rows_count": sum(int(entry["rows_count"]) for entry in daily.values()),
            "taux_avec_dem": round(taux_avec_dem_global, 4),
            "taux_corrige": round(taux_corrige_global, 4),
            "taux_ecarts": round(taux_ecarts_global, 4),
            "ecart_objectif": round(ANALYSE_FIABILITE_OBJECTIVE - taux_corrige_global, 4),
        },
        "columns": {
            "date": date_column,
            "uvc_controle": controle_column,
            "uvc_ecart": ecart_column,
            "demarque": demarque_column,
        },
        "filters": {
            "demarque_types": ["Manquant", "Surplus"],
        },
    }


def extract_source_daily_metrics(rows: list[dict], date_from_filter: str, date_to_filter: str) -> dict:
    date_column = find_source_column_name(rows, DATE_COLUMN_ALIASES, ("date",))
    prepared_column = find_source_column_name(rows, PREPARED_COLUMN_ALIASES, ("prepare",))
    fiabilite_column = find_source_column_name(
        rows,
        FIABILITE_COLUMN_ALIASES,
        ("fiabil",),
    )
    if not fiabilite_column:
        fiabilite_column = find_source_column_name(rows, FIABILITE_COLUMN_ALIASES, ("reliab",))

    missing = []
    if not date_column:
        missing.append("date")
    if not prepared_column:
        missing.append("total colis prepares")
    if not fiabilite_column:
        missing.append("taux de fiabilite")
    if missing:
        raise ValueError(
            "Source columns not detected: "
            + ", ".join(missing)
            + ". Expected headers like date/jour, total_colis_prepare, taux_fiabilite."
        )

    start_date = datetime.strptime(date_from_filter, "%Y-%m-%d").date() if date_from_filter else None
    end_date = datetime.strptime(date_to_filter, "%Y-%m-%d").date() if date_to_filter else None

    daily: dict[str, dict] = {}
    total_prepared = 0
    fiabilite_weighted_sum = 0.0
    fiabilite_weight_total = 0.0

    for row in rows:
        day = normalize_source_day(row.get(date_column))
        if not day:
            continue
        day_obj = datetime.strptime(day, "%Y-%m-%d").date()
        if start_date and day_obj < start_date:
            continue
        if end_date and day_obj > end_date:
            continue

        prepared = parse_int_value(row.get(prepared_column))
        fiabilite = parse_percent_value(row.get(fiabilite_column))
        if prepared is None and fiabilite is None:
            continue

        bucket = daily.setdefault(
            day,
            {
                "prepared_total": 0,
                "fiabilite_weighted_sum": 0.0,
                "fiabilite_weight": 0.0,
            },
        )
        if prepared is not None:
            bucket["prepared_total"] += prepared
            total_prepared += prepared
        if fiabilite is not None:
            weight = prepared if prepared and prepared > 0 else 1
            bucket["fiabilite_weighted_sum"] += fiabilite * weight
            bucket["fiabilite_weight"] += weight
            fiabilite_weighted_sum += fiabilite * weight
            fiabilite_weight_total += weight

    daily_values = {}
    for day, bucket in daily.items():
        fiabilite = None
        if bucket["fiabilite_weight"] > 0:
            fiabilite = bucket["fiabilite_weighted_sum"] / bucket["fiabilite_weight"]
        daily_values[day] = {
            "prepared_total": int(bucket["prepared_total"]),
            "fiabilite": fiabilite,
        }

    global_fiabilite = (
        (fiabilite_weighted_sum / fiabilite_weight_total) if fiabilite_weight_total > 0 else 0.0
    )
    return {
        "daily": daily_values,
        "total_prepared": int(total_prepared),
        "global_fiabilite": max(0.0, min(100.0, global_fiabilite)),
        "columns": {
            "date": date_column,
            "prepared_total": prepared_column,
            "fiabilite": fiabilite_column,
        },
    }


def query_local_ecarts_daily(
    db: sqlite3.Connection, username_filter: str, date_from_filter: str, date_to_filter: str
) -> dict:
    filters = []
    params: list[str] = []
    if username_filter:
        filters.append("u.username = ?")
        params.append(username_filter)
    if date_from_filter:
        filters.append("date(m.movement_date) >= date(?)")
        params.append(date_from_filter)
    if date_to_filter:
        filters.append("date(m.movement_date) <= date(?)")
        params.append(date_to_filter)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    rows = db.execute(
        f"""
        SELECT
            date(m.movement_date) AS day,
            COUNT(*) AS uvc_controlled,
            COALESCE(SUM(m.diff_plus + m.diff_minus), 0) AS ecarts_total
        FROM movements m
        JOIN users u ON u.id = m.user_id
        {where_clause}
        GROUP BY date(m.movement_date)
        ORDER BY date(m.movement_date) ASC
        """,
        params,
    ).fetchall()

    daily = {}
    total_uvc_controlled = 0
    total_ecarts = 0
    for row in rows:
        day = str(row["day"])
        uvc = int(row["uvc_controlled"] or 0)
        ecarts = int(row["ecarts_total"] or 0)
        daily[day] = {"uvc_controlled": uvc, "ecarts_total": ecarts}
        total_uvc_controlled += uvc
        total_ecarts += ecarts

    return {
        "daily": daily,
        "total_uvc_controlled": total_uvc_controlled,
        "total_ecarts": total_ecarts,
    }


def clean_export_cell(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()


def rtf_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")


def pdf_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def export_headers() -> list[str]:
    return [
        "ID",
        "Controleur",
        "Library",
        "Support",
        "EAN",
        "Produit",
        "Ecart +",
        "Ecart -",
        "Date mouvement",
        "Date creation",
    ]


def export_row_values(row: dict) -> list[str]:
    return [
        clean_export_cell(row.get("id")),
        clean_export_cell(row.get("username")),
        clean_export_cell(row.get("library_name") or "-"),
        clean_export_cell(row.get("support_number")),
        clean_export_cell(row.get("ean_code")),
        clean_export_cell(row.get("product_code")),
        clean_export_cell(row.get("diff_plus")),
        clean_export_cell(row.get("diff_minus")),
        clean_export_cell(row.get("movement_date")),
        clean_export_cell(row.get("created_at")),
    ]


def build_excel_export(rows: list[dict]) -> bytes:
    lines = ["\t".join(export_headers())]
    for row in rows:
        lines.append("\t".join(export_row_values(row)))
    return ("\ufeff" + "\n".join(lines)).encode("utf-8")


def build_text_export(rows: list[dict]) -> bytes:
    lines = [" | ".join(export_headers())]
    lines.append("-" * 160)
    for row in rows:
        lines.append(" | ".join(export_row_values(row)))
    return "\n".join(lines).encode("utf-8")


def build_word_export(rows: list[dict]) -> bytes:
    lines = ["Intranet Controle - Extraction"]
    lines.append(" | ".join(export_headers()))
    for row in rows:
        lines.append(" | ".join(export_row_values(row)))
    body = "\\line ".join(rtf_escape(line) for line in lines)
    rtf = r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\f0\fs20 " + body + "}"
    return rtf.encode("utf-8")


def build_simple_pdf(lines: list[str]) -> bytes:
    text_ops = ["BT", "/F1 10 Tf", "40 800 Td"]
    for line in lines:
        safe = line.encode("latin-1", "replace").decode("latin-1")
        text_ops.append(f"({pdf_escape(safe)}) Tj")
        text_ops.append("0 -13 Td")
    text_ops.append("ET")
    stream = "\n".join(text_ops).encode("latin-1")

    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets: list[int] = []
    for index, content in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{index} 0 obj\n".encode("ascii"))
        pdf.extend(content)
        pdf.extend(b"\nendobj\n")

    xref_offset = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    pdf.extend(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_offset}\n%%EOF"
        ).encode("ascii")
    )
    return bytes(pdf)


def build_pdf_export(rows: list[dict]) -> bytes:
    lines = [
        "Intranet Controle - Extraction PDF",
        f"Generated at {now_iso()}",
        "",
        " | ".join(export_headers()),
    ]
    for row in rows:
        line = " | ".join(export_row_values(row))
        if len(line) > 170:
            line = line[:167] + "..."
        lines.append(line)
    return build_simple_pdf(lines)


def build_donnes_export_package(export_format: str, rows: list[dict]) -> tuple[bytes, str, str]:
    normalized = (export_format or "").strip().lower()
    if normalized == "excel":
        return build_excel_export(rows), "application/vnd.ms-excel; charset=utf-8", "xls"
    if normalized == "text":
        return build_text_export(rows), "text/plain; charset=utf-8", "txt"
    if normalized == "word":
        return build_word_export(rows), "application/rtf; charset=utf-8", "rtf"
    if normalized == "pdf":
        return build_pdf_export(rows), "application/pdf", "pdf"
    raise ValueError("Invalid format. Use excel, text, word or pdf.")


def resolve_powershell_path() -> str:
    for candidate in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    raise ValueError("PowerShell is not available on this machine.")


def run_windows_powershell(script: str, extra_env: dict[str, str] | None = None, timeout: int = 180) -> str:
    if os.name != "nt":
        raise ValueError("Windows desktop actions are only available on Windows.")

    env = os.environ.copy()
    if extra_env:
        env.update({key: str(value) for key, value in extra_env.items()})

    completed = subprocess.run(
        [resolve_powershell_path(), "-NoProfile", "-STA", "-Command", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        env=env,
        timeout=timeout,
        check=False,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "Unknown PowerShell error.").strip()
        raise ValueError(detail)
    return (completed.stdout or "").strip()


def open_windows_save_dialog(*, title: str, default_filename: str, filters: str) -> str | None:
    script = r"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.SaveFileDialog
$dialog.Title = $env:CQ_DIALOG_TITLE
$dialog.Filter = $env:CQ_DIALOG_FILTER
$dialog.FileName = $env:CQ_DIALOG_FILE_NAME
$dialog.AddExtension = $true
$dialog.CheckPathExists = $true
$dialog.OverwritePrompt = $true
$dialog.RestoreDirectory = $true
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    [Console]::Out.Write($dialog.FileName)
}
"""
    with WINDOWS_DIALOG_LOCK:
        chosen = run_windows_powershell(
            script,
            extra_env={
                "CQ_DIALOG_TITLE": title,
                "CQ_DIALOG_FILTER": filters,
                "CQ_DIALOG_FILE_NAME": default_filename,
            },
        )
    return chosen or None


def open_windows_file_dialog(*, title: str, filters: str) -> str | None:
    script = r"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = $env:CQ_DIALOG_TITLE
$dialog.Filter = $env:CQ_DIALOG_FILTER
$dialog.Multiselect = $false
$dialog.CheckFileExists = $true
$dialog.RestoreDirectory = $true
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    [Console]::Out.Write($dialog.FileName)
}
"""
    with WINDOWS_DIALOG_LOCK:
        chosen = run_windows_powershell(
            script,
            extra_env={
                "CQ_DIALOG_TITLE": title,
                "CQ_DIALOG_FILTER": filters,
            },
        )
    return chosen or None


def export_dialog_config(export_format: str) -> tuple[str, str]:
    normalized = (export_format or "").strip().lower()
    if normalized == "excel":
        return "Extraction Excel", "Excel (*.xls)|*.xls|All files (*.*)|*.*"
    if normalized == "text":
        return "Extraction Text", "Text (*.txt)|*.txt|All files (*.*)|*.*"
    if normalized == "word":
        return "Extraction Word", "Word RTF (*.rtf)|*.rtf|All files (*.*)|*.*"
    if normalized == "pdf":
        return "Extraction PDF", "PDF (*.pdf)|*.pdf|All files (*.*)|*.*"
    raise ValueError("Invalid format. Use excel, text, word or pdf.")


def source_file_dialog_config(connection_type: str) -> tuple[str, str]:
    normalized = (connection_type or "").strip().lower()
    if normalized == "excel":
        return (
            "Choisir une source Excel",
            "Excel/CSV (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|All files (*.*)|*.*",
        )
    if normalized == "access":
        return (
            "Choisir une base Access",
            "Access (*.accdb;*.mdb)|*.accdb;*.mdb|Access (*.accdb)|*.accdb|Access 2003 (*.mdb)|*.mdb|All files (*.*)|*.*",
        )
    raise ValueError("Invalid source picker type.")


def resolve_odbc_administrator_path() -> Path:
    if os.name != "nt":
        raise ValueError("ODBC Administrator is only available on Windows.")

    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    candidates = [
        system_root / "System32" / "odbcad32.exe",
        system_root / "SysWOW64" / "odbcad32.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise ValueError("ODBC Administrator was not found on this machine.")


def open_odbc_administrator() -> Path:
    target = resolve_odbc_administrator_path()
    subprocess.Popen([str(target)])
    return target


def list_windows_odbc_drivers() -> list[str]:
    if os.name != "nt":
        return []

    try:
        import winreg
    except ImportError:
        return []

    registry_paths = [
        r"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers",
        r"SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Drivers",
    ]
    detected: set[str] = set()
    for registry_path in registry_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path) as key:
                index = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, index)
                    except OSError:
                        break
                    if str(value).strip().lower() == "installed":
                        detected.add(str(name).strip())
                    index += 1
        except OSError:
            continue
    return sorted((driver for driver in detected if driver), key=str.lower)


def is_pyodbc_installed() -> bool:
    try:
        import pyodbc  # type: ignore # noqa: F401
    except ImportError:
        return False
    return True


def query_export_rows(
    db: sqlite3.Connection,
    username_filter: str,
    date_from_filter: str,
    date_to_filter: str,
) -> list[dict]:
    filters = []
    params: list[str] = []
    if username_filter:
        filters.append("u.username = ?")
        params.append(username_filter)
    if date_from_filter:
        filters.append("date(m.movement_date) >= date(?)")
        params.append(date_from_filter)
    if date_to_filter:
        filters.append("date(m.movement_date) <= date(?)")
        params.append(date_to_filter)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    rows = db.execute(
        f"""
        SELECT
            m.id,
            u.username,
            l.name AS library_name,
            m.support_number,
            m.ean_code,
            m.product_code,
            m.diff_plus,
            m.diff_minus,
            m.movement_date,
            m.created_at
        FROM movements m
        JOIN users u ON u.id = m.user_id
        LEFT JOIN libraries l ON l.id = m.library_id
        {where_clause}
        ORDER BY m.movement_date DESC
        LIMIT 20000
        """,
        params,
    ).fetchall()
    return [dict(row) for row in rows]


def parse_user_ids(raw_user_ids) -> list[int]:
    if raw_user_ids is None:
        return []
    if not isinstance(raw_user_ids, list):
        raise ValueError("user_ids must be a list of integers.")

    parsed: list[int] = []
    seen: set[int] = set()
    for raw in raw_user_ids:
        try:
            user_id = int(raw)
        except (TypeError, ValueError) as exc:
            raise ValueError("user_ids must contain valid integers.") from exc
        if user_id <= 0 or user_id in seen:
            continue
        seen.add(user_id)
        parsed.append(user_id)
    return parsed


def get_library_or_404(db: sqlite3.Connection, library_id: int) -> sqlite3.Row:
    row = db.execute(
        "SELECT id, name, created_at FROM libraries WHERE id = ?",
        (library_id,),
    ).fetchone()
    if not row:
        abort(404, description="Library not found.")
    return row


def user_has_library_access(db: sqlite3.Connection, library_id: int, user_id: int) -> bool:
    row = db.execute(
        "SELECT 1 FROM library_users WHERE library_id = ? AND user_id = ?",
        (library_id, user_id),
    ).fetchone()
    return row is not None


@app.before_request
def before_request():
    ensure_csrf()
    g.user = current_user()


@app.context_processor
def inject_globals():
    return {"current_user": g.user, "csrf_token": session.get("csrf_token", "")}


@app.after_request
def set_security_headers(response):
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "camera=(self)"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self'; "
        "style-src 'self' https://fonts.googleapis.com 'unsafe-inline'; "
        "font-src 'self' https://fonts.gstatic.com data:; "
        "img-src 'self' data: blob:; "
        "connect-src 'self'; "
        "media-src 'self' blob:; "
        "worker-src 'self' blob:; "
        "base-uri 'self'; "
        "form-action 'self'; "
        "frame-ancestors 'none'"
    )
    return response


@app.get("/")
def index():
    if g.user is None:
        return redirect(url_for("login"))
    return redirect(url_for("admin_dashboard"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        require_csrf()
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        key = login_key(username)
        if is_locked(key):
            flash("Too many failed attempts. Try again in 5 minutes.", "error")
            return render_template("login.html"), 429

        db = get_db()
        row = db.execute(
            "SELECT id, username, password_hash, role, is_blocked FROM users WHERE username = ?",
            (username,),
        ).fetchone()
        if row and int(row["is_blocked"] or 0):
            flash("Account is blocked. Contact your administrator.", "error")
            return render_template("login.html"), 403
        if not row or not check_password_hash(row["password_hash"], password):
            register_failed_login(key)
            return render_template("login.html", login_failed=True), 401

        clear_failed_login(key)
        session.clear()
        session["user_id"] = row["id"]
        session["csrf_token"] = secrets.token_urlsafe(32)
        session.permanent = True
        return redirect(url_for("admin_dashboard"))

    return render_template("login.html")


@app.post("/logout")
@login_required
def logout():
    require_csrf()
    session.clear()
    return redirect(url_for("login"))


@app.get("/controller")
@role_required("controller")
def controller_dashboard():
    return redirect(url_for("admin_dashboard"))


@app.post("/api/movements")
@role_required("controller")
def create_movement():
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400
    try:
        payload = validate_movement_payload(request.get_json() or {})
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    movement_date = now_iso()
    cur = db.execute(
        """
        INSERT INTO movements (
            user_id, library_id, support_number, ean_code, product_code, diff_plus, diff_minus, movement_date, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            g.user["id"],
            None,
            payload["support_number"],
            payload["ean_code"],
            payload["product_code"],
            payload["diff_plus"],
            payload["diff_minus"],
            movement_date,
            now_iso(),
        ),
    )
    db.commit()
    return jsonify(
        {
            "ok": True,
            "message": "Mouvement saved.",
            "movement_id": cur.lastrowid,
            "movement_date": movement_date,
        }
    )


@app.get("/admin")
@role_required("admin", "controller")
def admin_dashboard():
    return render_template("admin.html")


@app.get("/api/admin/users")
@role_required("admin")
def admin_users():
    db = get_db()
    rows = db.execute(
        """
        SELECT id, username, role
        FROM users
        WHERE role = 'controller'
        ORDER BY username ASC
        """
    ).fetchall()
    return jsonify([dict(row) for row in rows])


@app.get("/api/admin/user-management/users")
@role_required("admin")
def admin_user_management_users():
    db = get_db()
    rows = db.execute(
        """
        SELECT id, username, role, is_blocked, permissions_json, created_at
        FROM users
        ORDER BY
            CASE role WHEN 'admin' THEN 0 ELSE 1 END,
            username COLLATE NOCASE ASC
        """
    ).fetchall()
    return jsonify(
        {
            "ok": True,
            "permissions_catalog": permissions_catalog_payload(),
            "users": [serialize_user_row(row) for row in rows],
        }
    )


@app.post("/api/admin/user-management/users")
@role_required("admin")
def admin_user_management_create_user():
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    try:
        username = validate_username(str(payload.get("username", "")))
        password = validate_password(payload.get("password") or "")
        role = parse_user_role(str(payload.get("role", "controller")))
        permissions = normalize_permissions_payload(payload.get("permissions"), role)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    try:
        cur = db.execute(
            """
            INSERT INTO users (username, password_hash, role, created_at, is_blocked, permissions_json)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                username,
                generate_password_hash(password),
                role,
                now_iso(),
                0,
                json.dumps(permissions),
            ),
        )
        db.commit()
    except sqlite3.IntegrityError:
        return jsonify({"ok": False, "error": "Username already exists."}), 409

    row = db.execute(
        """
        SELECT id, username, role, is_blocked, permissions_json, created_at
        FROM users
        WHERE id = ?
        """,
        (cur.lastrowid,),
    ).fetchone()
    return jsonify({"ok": True, "user": serialize_user_row(row)})


@app.put("/api/admin/user-management/users/<int:user_id>")
@role_required("admin")
def admin_user_management_update_user(user_id: int):
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    db = get_db()
    row = db.execute(
        """
        SELECT id, username, role, is_blocked, permissions_json, created_at
        FROM users
        WHERE id = ?
        """,
        (user_id,),
    ).fetchone()
    if not row:
        return jsonify({"ok": False, "error": "User not found."}), 404

    current = serialize_user_row(row)
    try:
        new_role = parse_user_role(payload.get("role")) if "role" in payload else current["role"]
        if "is_blocked" in payload:
            new_is_blocked = parse_bool(payload.get("is_blocked"), "is_blocked")
        else:
            new_is_blocked = bool(current["is_blocked"])

        if "permissions" in payload:
            new_permissions = normalize_permissions_payload(payload.get("permissions"), new_role)
        elif new_role != current["role"]:
            new_permissions = default_permissions_for_role(new_role)
        else:
            new_permissions = list(current["permissions"])
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    if user_id == g.user["id"]:
        if new_is_blocked:
            return jsonify({"ok": False, "error": "You cannot block your own account."}), 400
        if new_role != "admin":
            return jsonify({"ok": False, "error": "You cannot remove your own admin role."}), 400

    if current["role"] == "admin" and new_role != "admin":
        if count_admins(db, exclude_user_id=user_id) == 0:
            return jsonify({"ok": False, "error": "At least one admin account must remain."}), 400
    if current["role"] == "admin" and not current["is_blocked"] and new_is_blocked:
        if count_admins(db, exclude_user_id=user_id, only_unblocked=True) == 0:
            return jsonify({"ok": False, "error": "At least one unblocked admin account must remain."}), 400

    db.execute(
        """
        UPDATE users
        SET role = ?, is_blocked = ?, permissions_json = ?
        WHERE id = ?
        """,
        (new_role, 1 if new_is_blocked else 0, json.dumps(new_permissions), user_id),
    )
    db.commit()

    updated = db.execute(
        """
        SELECT id, username, role, is_blocked, permissions_json, created_at
        FROM users
        WHERE id = ?
        """,
        (user_id,),
    ).fetchone()
    return jsonify({"ok": True, "user": serialize_user_row(updated)})


@app.delete("/api/admin/user-management/users/<int:user_id>")
@role_required("admin")
def admin_user_management_delete_user(user_id: int):
    require_csrf()
    if user_id == g.user["id"]:
        return jsonify({"ok": False, "error": "You cannot delete your own account."}), 400

    db = get_db()
    row = db.execute(
        """
        SELECT id, username, role, is_blocked
        FROM users
        WHERE id = ?
        """,
        (user_id,),
    ).fetchone()
    if not row:
        return jsonify({"ok": False, "error": "User not found."}), 404

    if row["role"] == "admin" and count_admins(db, exclude_user_id=user_id) == 0:
        return jsonify({"ok": False, "error": "At least one admin account must remain."}), 400

    movements_row = db.execute(
        "SELECT COUNT(*) AS total FROM movements WHERE user_id = ?",
        (user_id,),
    ).fetchone()
    movements_count = int(movements_row["total"]) if movements_row else 0
    library_links_row = db.execute(
        "SELECT COUNT(*) AS total FROM library_users WHERE user_id = ?",
        (user_id,),
    ).fetchone()
    library_links_count = int(library_links_row["total"]) if library_links_row else 0
    if movements_count > 0:
        db.execute("DELETE FROM movements WHERE user_id = ?", (user_id,))
    if library_links_count > 0:
        db.execute("DELETE FROM library_users WHERE user_id = ?", (user_id,))
    db.execute("DELETE FROM users WHERE id = ?", (user_id,))
    db.commit()
    return jsonify(
        {
            "ok": True,
            "message": f'User "{row["username"]}" deleted.',
            "deleted_movements": movements_count,
            "deleted_library_links": library_links_count,
        }
    )


@app.post("/api/admin/user-management/users/<int:user_id>/reset-password")
@role_required("admin")
def admin_user_management_reset_password(user_id: int):
    require_csrf()
    db = get_db()
    row = db.execute(
        """
        SELECT id, username
        FROM users
        WHERE id = ?
        """,
        (user_id,),
    ).fetchone()
    if not row:
        return jsonify({"ok": False, "error": "User not found."}), 404

    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789"
    temporary_password = "".join(secrets.choice(alphabet) for _ in range(10))
    db.execute(
        "UPDATE users SET password_hash = ? WHERE id = ?",
        (generate_password_hash(temporary_password), user_id),
    )
    db.commit()
    return jsonify(
        {
            "ok": True,
            "username": row["username"],
            "temporary_password": temporary_password,
        }
    )


@app.get("/api/admin/settings/library-access")
@role_required("admin")
def admin_settings_library_access():
    db = get_db()
    libraries = db.execute(
        """
        SELECT
            l.id,
            l.name,
            COUNT(lu.user_id) AS users_count
        FROM libraries l
        LEFT JOIN library_users lu ON lu.library_id = l.id
        GROUP BY l.id, l.name
        ORDER BY l.name COLLATE NOCASE ASC
        """
    ).fetchall()
    users = db.execute(
        """
        SELECT id, username, role
        FROM users
        ORDER BY
            CASE role WHEN 'admin' THEN 0 ELSE 1 END,
            username COLLATE NOCASE ASC
        """
    ).fetchall()
    assignments = db.execute(
        """
        SELECT library_id, user_id
        FROM library_users
        ORDER BY library_id ASC, user_id ASC
        """
    ).fetchall()
    return jsonify(
        {
            "ok": True,
            "libraries": [dict(row) for row in libraries],
            "users": [dict(row) for row in users],
            "assignments": [dict(row) for row in assignments],
        }
    )


@app.put("/api/admin/settings/libraries/<int:library_id>/users")
@role_required("admin")
def admin_settings_set_library_users(library_id: int):
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    try:
        user_ids = parse_user_ids(payload.get("user_ids"))
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    library = dict(get_library_or_404(db, library_id))
    if user_ids:
        placeholders = ",".join("?" for _ in user_ids)
        params: list[int] = [*user_ids]
        rows = db.execute(
            f"""
            SELECT id
            FROM users
            WHERE role = 'controller' AND id IN ({placeholders})
            """,
            params,
        ).fetchall()
        valid_user_ids = {row["id"] for row in rows}
        unknown = [str(uid) for uid in user_ids if uid not in valid_user_ids]
        if unknown:
            return jsonify({"ok": False, "error": f"Invalid controller user id(s): {', '.join(unknown)}"}), 400

    db.execute("DELETE FROM library_users WHERE library_id = ?", (library_id,))
    if user_ids:
        db.executemany(
            """
            INSERT INTO library_users (library_id, user_id, created_at)
            VALUES (?, ?, ?)
            """,
            [(library_id, user_id, now_iso()) for user_id in user_ids],
        )
    db.commit()

    assigned_rows = db.execute(
        """
        SELECT u.id, u.username, u.role
        FROM users u
        JOIN library_users lu ON lu.user_id = u.id
        WHERE lu.library_id = ?
        ORDER BY u.username COLLATE NOCASE ASC
        """,
        (library_id,),
    ).fetchall()
    return jsonify(
        {
            "ok": True,
            "library": library,
            "assigned_users": [dict(row) for row in assigned_rows],
        }
    )


@app.get("/api/admin/donnes/extract")
@role_required("admin")
def admin_donnes_extract():
    export_format = (request.args.get("format") or "").strip().lower()
    username = (request.args.get("username") or "").strip()
    date_from_raw = request.args.get("date_from") or ""
    date_to_raw = request.args.get("date_to") or ""

    try:
        date_from_filter = parse_date_filter(date_from_raw, "date_from")
        date_to_filter = parse_date_filter(date_to_raw, "date_to")
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    rows = query_export_rows(db, username, date_from_filter, date_to_filter)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        payload, content_type, ext = build_donnes_export_package(export_format, rows)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    response = Response(payload)
    response.headers["Content-Type"] = content_type
    response.headers["Content-Disposition"] = f'attachment; filename="donnes_{export_format}_{stamp}.{ext}"'
    return response


@app.post("/api/admin/donnes/extract/save")
@role_required("admin")
def admin_donnes_extract_save():
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    export_format = str(payload.get("format") or "").strip().lower()
    username = str(payload.get("username") or "").strip()
    date_from_raw = str(payload.get("date_from") or "").strip()
    date_to_raw = str(payload.get("date_to") or "").strip()

    try:
        date_from_filter = parse_date_filter(date_from_raw, "date_from")
        date_to_filter = parse_date_filter(date_to_raw, "date_to")
        dialog_title, dialog_filter = export_dialog_config(export_format)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    rows = query_export_rows(db, username, date_from_filter, date_to_filter)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    try:
        export_bytes, _content_type, ext = build_donnes_export_package(export_format, rows)
        default_name = f"donnes_{export_format}_{stamp}.{ext}"
        target_path = open_windows_save_dialog(
            title=dialog_title,
            default_filename=default_name,
            filters=dialog_filter,
        )
    except (OSError, ValueError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    if not target_path:
        return jsonify({"ok": True, "cancelled": True, "message": "Export cancelled."})

    try:
        final_path = Path(target_path).expanduser()
        final_path.write_bytes(export_bytes)
    except OSError as exc:
        return jsonify({"ok": False, "error": f"Unable to save file: {exc}"}), 400

    return jsonify(
        {
            "ok": True,
            "message": f"Extraction {export_format} saved: {final_path.name}.",
            "format": export_format,
            "saved_path": str(final_path),
            "rows_count": len(rows),
        }
    )


@app.post("/api/admin/donnes/connections/pick-file")
@role_required("admin")
def admin_donnes_connection_pick_file():
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    connection_type = str(payload.get("type") or "").strip().lower()
    try:
        dialog_title, dialog_filter = source_file_dialog_config(connection_type)
        selected_path = open_windows_file_dialog(title=dialog_title, filters=dialog_filter)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    if not selected_path:
        return jsonify({"ok": True, "cancelled": True, "message": "Selection cancelled."})

    return jsonify(
        {
            "ok": True,
            "path": str(Path(selected_path).expanduser()),
            "type": connection_type,
        }
    )


@app.post("/api/admin/donnes/connections/open-odbc-admin")
@role_required("admin")
def admin_donnes_connection_open_odbc_admin():
    require_csrf()
    try:
        launched_path = open_odbc_administrator()
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    return jsonify(
        {
            "ok": True,
            "message": "ODBC Administrator opened. Choose your driver, then paste the connection string below.",
            "launched_path": str(launched_path),
            "drivers": list_windows_odbc_drivers(),
            "pyodbc_installed": is_pyodbc_installed(),
        }
    )


@app.post("/api/admin/donnes/connections/test")
@role_required("admin")
def admin_donnes_connection_test():
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    connection_type = str(payload.get("type") or "").strip().lower()
    db = get_db()

    if connection_type == "excel":
        raw_path = (payload.get("path") or "").strip()
        if not raw_path:
            return jsonify({"ok": False, "error": "Excel path is required."}), 400
        file_path = Path(raw_path).expanduser()
        if not file_path.exists() or not file_path.is_file():
            return jsonify({"ok": False, "error": "Excel file not found."}), 400
        if file_path.suffix.lower() not in {".xlsx", ".xls", ".csv"}:
            return jsonify({"ok": False, "error": "Excel extension must be .xlsx, .xls or .csv."}), 400
        if file_path.suffix.lower() == ".xls":
            return jsonify(
                {
                    "ok": False,
                    "error": "Legacy .xls is not supported for Analyse. Use .xlsx/.csv or ODBC.",
                }
            ), 400
        try:
            source_rows = read_excel_source_rows(str(file_path))
            metrics = extract_interception_metrics(source_rows, "", "")
            size = file_path.stat().st_size
        except (OSError, ValueError) as exc:
            return jsonify({"ok": False, "error": f"Excel file cannot be read: {exc}"}), 400
        save_donnes_connection(db, "excel", str(file_path.resolve()))
        return jsonify(
            {
                "ok": True,
                "message": (
                    f"Excel source saved: {file_path.name} ({size} bytes, {len(source_rows)} row(s) read, {len(metrics['available_months'])} month(s) detected)."
                ),
                "saved_connection": {"type": "excel", "value": str(file_path.resolve())},
            }
        )

    if connection_type in {"odbc", "access"}:
        try:
            pyodbc = import_pyodbc()
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400

        if connection_type == "odbc":
            connection_string = (payload.get("connection_string") or "").strip()
            if not connection_string:
                return jsonify({"ok": False, "error": "ODBC connection string is required."}), 400
            try:
                conn = pyodbc.connect(connection_string, timeout=5)
                conn.close()
            except Exception as exc:  # noqa: BLE001
                return jsonify({"ok": False, "error": f"ODBC connection failed: {exc}"}), 400
            config = {}
            table_hint = str(payload.get("table") or "").strip()
            if table_hint:
                config["table"] = table_hint
            try:
                source_rows, _table_label = odbc_rows_from_connection(connection_string, table_hint)
                extract_interception_metrics(source_rows, "", "")
            except ValueError as exc:
                return jsonify({"ok": False, "error": str(exc)}), 400
            save_donnes_connection(db, "odbc", connection_string, config=config)
            return jsonify(
                {
                    "ok": True,
                    "message": "ODBC connection successful and saved for Analyse.",
                    "saved_connection": {"type": "odbc", "value": connection_string},
                }
            )

        raw_path = (payload.get("path") or "").strip()
        if not raw_path:
            return jsonify({"ok": False, "error": "Access path is required."}), 400
        file_path = Path(raw_path).expanduser()
        if not file_path.exists() or not file_path.is_file():
            return jsonify({"ok": False, "error": "Access file not found."}), 400
        if file_path.suffix.lower() not in {".mdb", ".accdb"}:
            return jsonify({"ok": False, "error": "Access extension must be .mdb or .accdb."}), 400

        driver = str(payload.get("driver") or "{Microsoft Access Driver (*.mdb, *.accdb)}").strip()
        access_conn = f"Driver={driver};DBQ={file_path};"
        try:
            conn = pyodbc.connect(access_conn, timeout=5)
            conn.close()
        except Exception as exc:  # noqa: BLE001
            return jsonify({"ok": False, "error": f"Access connection failed: {exc}"}), 400
        config = {"driver": driver}
        table_hint = str(payload.get("table") or "").strip()
        if table_hint:
            config["table"] = table_hint
        try:
            source_rows, _table_label = odbc_rows_from_connection(access_conn, table_hint)
            extract_interception_metrics(source_rows, "", "")
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        save_donnes_connection(db, "access", str(file_path.resolve()), config=config)
        return jsonify(
            {
                "ok": True,
                "message": f"Access connection successful and saved: {file_path.name}.",
                "saved_connection": {"type": "access", "value": str(file_path.resolve())},
            }
        )

    return jsonify({"ok": False, "error": "Invalid connection type."}), 400


@app.get("/api/admin/analyse")
@role_required("admin")
def admin_analyse():
    date_from_raw = request.args.get("date_from") or ""
    date_to_raw = request.args.get("date_to") or ""

    try:
        date_from_filter = parse_date_filter(date_from_raw, "date_from")
        date_to_filter = parse_date_filter(date_to_raw, "date_to")
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    saved_connection = load_saved_donnes_connection(db)
    if not saved_connection:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "No Donnes connection configured. Configure it in Donnes > Connection first.",
                }
            ),
            400,
        )

    try:
        source_rows, source_meta = load_source_rows_from_connection(saved_connection)
        source_metrics = extract_interception_metrics(source_rows, date_from_filter, date_to_filter)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    return jsonify(
        {
            "ok": True,
            "source": {
                "type": source_meta.get("connection_type") or saved_connection.get("type"),
                "label": source_meta.get("source_label") or "source",
                "updated_at": saved_connection.get("updated_at"),
                "columns": source_metrics.get("columns") or {},
            },
            "filters": source_metrics.get("filters") or {},
            "available_months": source_metrics.get("available_months") or [],
            "totals": source_metrics.get("totals") or {},
            "daily_fiabilite": source_metrics.get("daily_fiabilite") or [],
            "daily_ecarts": source_metrics.get("daily_ecarts") or [],
            "daily_details": source_metrics.get("daily_details") or [],
            "monthly_fiabilite": source_metrics.get("monthly_fiabilite") or [],
            "monthly_ecarts": source_metrics.get("monthly_ecarts") or [],
            "monthly_details": source_metrics.get("monthly_details") or [],
        }
    )


@app.get("/api/admin/libraries")
@role_required("admin", "controller")
def admin_libraries():
    db = get_db()
    if g.user["role"] == "admin":
        rows = db.execute(
            """
            SELECT
                l.id,
                l.name,
                l.created_at,
                COUNT(lu.user_id) AS users_count
            FROM libraries l
            LEFT JOIN library_users lu ON lu.library_id = l.id
            GROUP BY l.id, l.name, l.created_at
            ORDER BY l.name COLLATE NOCASE ASC
            """
        ).fetchall()
    else:
        rows = db.execute(
            """
            SELECT
                l.id,
                l.name,
                l.created_at,
                COUNT(lu_all.user_id) AS users_count
            FROM libraries l
            JOIN library_users lu_me
              ON lu_me.library_id = l.id
             AND lu_me.user_id = ?
            LEFT JOIN library_users lu_all ON lu_all.library_id = l.id
            GROUP BY l.id, l.name, l.created_at
            ORDER BY l.name COLLATE NOCASE ASC
            """,
            (g.user["id"],),
        ).fetchall()
    return jsonify([dict(row) for row in rows])


@app.post("/api/admin/libraries")
@role_required("admin")
def admin_create_library():
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    try:
        name = validate_library_name(str(payload.get("name", "")))
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    try:
        cur = db.execute(
            """
            INSERT INTO libraries (name, created_at)
            VALUES (?, ?)
            """,
            (name, now_iso()),
        )
        db.commit()
    except sqlite3.IntegrityError:
        return jsonify({"ok": False, "error": "Library already exists."}), 409

    return jsonify(
        {
            "ok": True,
            "library": {"id": cur.lastrowid, "name": name},
        }
    )


@app.delete("/api/admin/libraries/<int:library_id>")
@role_required("admin")
def admin_delete_library(library_id: int):
    require_csrf()
    db = get_db()
    library = dict(get_library_or_404(db, library_id))

    # Keep historical rows and unlink them from the removed library.
    db.execute("UPDATE movements SET library_id = NULL WHERE library_id = ?", (library_id,))
    db.execute("DELETE FROM libraries WHERE id = ?", (library_id,))
    db.commit()
    return jsonify({"ok": True, "message": f'Library "{library["name"]}" deleted.'})


@app.get("/api/admin/libraries/<int:library_id>/users")
@role_required("admin", "controller")
def admin_library_users(library_id: int):
    db = get_db()
    library = dict(get_library_or_404(db, library_id))
    if g.user["role"] != "admin" and not user_has_library_access(db, library_id, g.user["id"]):
        abort(403)

    assigned = db.execute(
        """
        SELECT u.id, u.username, u.role
        FROM library_users lu
        JOIN users u ON u.id = lu.user_id
        WHERE lu.library_id = ?
        ORDER BY u.username COLLATE NOCASE ASC
        """,
        (library_id,),
    ).fetchall()
    available = []
    if g.user["role"] == "admin":
        available = db.execute(
            """
            SELECT u.id, u.username, u.role
            FROM users u
            WHERE u.role = 'controller'
              AND NOT EXISTS (
                SELECT 1
                FROM library_users lu
                WHERE lu.library_id = ? AND lu.user_id = u.id
              )
            ORDER BY u.username COLLATE NOCASE ASC
            """,
            (library_id,),
        ).fetchall()
    return jsonify(
        {
            "library": library,
            "assigned_users": [dict(row) for row in assigned],
            "available_users": [dict(row) for row in available],
            "can_manage_users": g.user["role"] == "admin",
        }
    )


@app.post("/api/admin/libraries/<int:library_id>/users")
@role_required("admin")
def admin_library_add_user(library_id: int):
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    payload = request.get_json() or {}
    raw_user_id = payload.get("user_id")
    try:
        user_id = int(raw_user_id)
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "user_id must be a valid integer."}), 400

    db = get_db()
    get_library_or_404(db, library_id)
    user_row = db.execute(
        """
        SELECT id, username
        FROM users
        WHERE id = ? AND role = 'controller'
        """,
        (user_id,),
    ).fetchone()
    if not user_row:
        return jsonify({"ok": False, "error": "Controller user not found."}), 404

    db.execute(
        """
        INSERT OR IGNORE INTO library_users (library_id, user_id, created_at)
        VALUES (?, ?, ?)
        """,
        (library_id, user_id, now_iso()),
    )
    db.commit()
    return jsonify({"ok": True, "user": dict(user_row)})


@app.delete("/api/admin/libraries/<int:library_id>/users/<int:user_id>")
@role_required("admin")
def admin_library_remove_user(library_id: int, user_id: int):
    require_csrf()
    db = get_db()
    get_library_or_404(db, library_id)
    cur = db.execute(
        """
        DELETE FROM library_users
        WHERE library_id = ? AND user_id = ?
        """,
        (library_id, user_id),
    )
    db.commit()
    if cur.rowcount == 0:
        return jsonify({"ok": False, "error": "User not assigned to this library."}), 404
    return jsonify({"ok": True})


@app.post("/api/admin/libraries/<int:library_id>/users/<int:user_id>/movements")
@role_required("admin", "controller")
def admin_library_user_create_movement(library_id: int, user_id: int):
    require_csrf()
    if not request.is_json:
        return jsonify({"ok": False, "error": "JSON payload required."}), 400

    db = get_db()
    get_library_or_404(db, library_id)
    if g.user["role"] != "admin":
        if user_id != g.user["id"]:
            abort(403)
        if not user_has_library_access(db, library_id, g.user["id"]):
            abort(403)

    target_user = db.execute(
        """
        SELECT u.id, u.username
        FROM users u
        JOIN library_users lu ON lu.user_id = u.id
        WHERE lu.library_id = ? AND u.id = ? AND u.role = 'controller'
        """,
        (library_id, user_id),
    ).fetchone()
    if not target_user:
        return jsonify({"ok": False, "error": "User not found in this library."}), 404

    try:
        payload = validate_movement_payload(request.get_json() or {})
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    movement_date = now_iso()
    cur = db.execute(
        """
        INSERT INTO movements (
            user_id, library_id, support_number, ean_code, product_code, diff_plus, diff_minus, movement_date, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            user_id,
            library_id,
            payload["support_number"],
            payload["ean_code"],
            payload["product_code"],
            payload["diff_plus"],
            payload["diff_minus"],
            movement_date,
            now_iso(),
        ),
    )
    db.commit()
    return jsonify(
        {
            "ok": True,
            "message": "Data saved.",
            "movement_id": cur.lastrowid,
            "movement_date": movement_date,
            "library_id": library_id,
            "username": target_user["username"],
        }
    )


@app.get("/api/admin/libraries/<int:library_id>/users/<int:user_id>/movements")
@role_required("admin", "controller")
def admin_library_user_movements(library_id: int, user_id: int):
    month = request.args.get("month")
    try:
        month_filter = parse_month_filter(month or "")
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    db = get_db()
    library = dict(get_library_or_404(db, library_id))
    if g.user["role"] != "admin":
        if not user_has_library_access(db, library_id, g.user["id"]):
            abort(403)
        if user_id != g.user["id"]:
            abort(403)

    user_row = db.execute(
        """
        SELECT u.id, u.username
        FROM users u
        JOIN library_users lu ON lu.user_id = u.id
        WHERE lu.library_id = ? AND u.id = ?
        """,
        (library_id, user_id),
    ).fetchone()
    if not user_row:
        return jsonify({"ok": False, "error": "User not found in this library."}), 404

    filters = ["m.user_id = ?", "m.library_id = ?"]
    params: list[str | int] = [user_id, library_id]
    if month_filter:
        filters.append("strftime('%Y-%m', m.movement_date) = ?")
        params.append(month_filter)

    where_clause = "WHERE " + " AND ".join(filters)
    rows = db.execute(
        f"""
        SELECT
            m.id,
            m.support_number,
            m.ean_code,
            m.product_code,
            m.diff_plus,
            m.diff_minus,
            m.movement_date,
            m.created_at
        FROM movements m
        {where_clause}
        ORDER BY m.movement_date DESC
        """,
        params,
    ).fetchall()

    grouped: dict[str, list[dict]] = {}
    for row in rows:
        row_dict = dict(row)
        day_key = row_dict["movement_date"][:10]
        grouped.setdefault(day_key, []).append(row_dict)

    days = []
    for day in sorted(grouped.keys(), reverse=True):
        items = grouped[day]
        days.append({"day": day, "count": len(items), "items": items})

    return jsonify(
        {
            "ok": True,
            "library": library,
            "user": dict(user_row),
            "month": month_filter,
            "days": days,
            "count": len(rows),
        }
    )


@app.get("/api/admin/movements")
@role_required("admin", "controller")
def admin_movements():
    username = (request.args.get("username") or "").strip()
    date_from = (request.args.get("date_from") or "").strip()
    date_to = (request.args.get("date_to") or "").strip()

    filters = []
    params: list[str | int] = []

    if g.user["role"] == "admin" and username:
        filters.append("u.username = ?")
        params.append(username)
    if g.user["role"] != "admin":
        filters.append("u.id = ?")
        params.append(g.user["id"])
        filters.append(
            """
            EXISTS (
                SELECT 1
                FROM library_users lu
                WHERE lu.library_id = m.library_id
                  AND lu.user_id = ?
            )
            """
        )
        params.append(g.user["id"])
    if date_from:
        filters.append("date(m.movement_date) >= date(?)")
        params.append(date_from)
    if date_to:
        filters.append("date(m.movement_date) <= date(?)")
        params.append(date_to)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    query = f"""
        SELECT
            m.id,
            u.username,
            m.support_number,
            m.ean_code,
            m.product_code,
            m.diff_plus,
            m.diff_minus,
            m.movement_date
        FROM movements m
        JOIN users u ON u.id = m.user_id
        {where_clause}
        ORDER BY m.movement_date DESC
        LIMIT 1000
    """
    db = get_db()
    rows = db.execute(query, params).fetchall()
    return jsonify([dict(row) for row in rows])


@app.get("/api/admin/movements/<int:movement_id>")
@role_required("admin", "controller")
def admin_movement_detail(movement_id: int):
    db = get_db()
    row = db.execute(
        """
        SELECT
            m.id,
            m.user_id,
            m.library_id,
            u.username,
            m.support_number,
            m.ean_code,
            m.product_code,
            m.diff_plus,
            m.diff_minus,
            m.movement_date,
            m.created_at
        FROM movements m
        JOIN users u ON u.id = m.user_id
        WHERE m.id = ?
        """,
        (movement_id,),
    ).fetchone()
    if not row:
        abort(404, description="Movement not found.")
    row_dict = dict(row)
    if g.user["role"] != "admin":
        if row_dict["user_id"] != g.user["id"]:
            abort(403)
        library_id = row_dict.get("library_id")
        if library_id is None or not user_has_library_access(db, int(library_id), g.user["id"]):
            abort(403)
    row_dict.pop("library_id", None)
    row_dict.pop("user_id", None)
    return jsonify(row_dict)


@app.delete("/api/admin/movements/<int:movement_id>")
@role_required("admin", "controller")
def admin_delete_movement(movement_id: int):
    require_csrf()
    db = get_db()
    if g.user["role"] == "admin":
        cur = db.execute("DELETE FROM movements WHERE id = ?", (movement_id,))
    else:
        cur = db.execute(
            """
            DELETE FROM movements
            WHERE id = ?
              AND user_id = ?
              AND EXISTS (
                SELECT 1
                FROM library_users lu
                WHERE lu.library_id = movements.library_id
                  AND lu.user_id = ?
              )
            """,
            (movement_id, g.user["id"], g.user["id"]),
        )
    db.commit()
    if cur.rowcount == 0:
        return jsonify({"ok": False, "error": "Movement not found."}), 404
    return jsonify({"ok": True, "message": "Movement deleted."})


@app.errorhandler(400)
def bad_request(error):
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "error": getattr(error, "description", "Bad request.")}), 400
    return render_template("error.html", status=400, message=str(error)), 400


@app.errorhandler(403)
def forbidden(error):
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "error": "Forbidden"}), 403
    return render_template("error.html", status=403, message=str(error)), 403


@app.errorhandler(404)
def not_found(error):
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "error": "Not found"}), 404
    return render_template("error.html", status=404, message=str(error)), 404


def bootstrap():
    with app.app_context():
        init_db()
        seed_users()


bootstrap()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")), debug=False)

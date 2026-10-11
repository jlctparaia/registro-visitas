# -*- coding: utf-8 -*-
"""
RegistroVisitas2.py
===================
Registro de visitas de trabajadores a centros sanitarios, con generación
automática de un Excel por Distrito/Hospital y año (una pestaña por mes).

Requisitos:   Python 3.8+   y   pip install openpyxl
Ejecución:    python RegistroVisitas2.py

Archivos que crea junto al programa:
  * registro_visitas.db                      -> base de datos (fuente de los datos)
  * REGISTRO_VISITAS_<DISTRITO>_<AÑO>.xlsx   -> un Excel por distrito y año

Importar la estructura desde Excel (menú «Distritos, zonas y centros»):
    columna A = Distrito/Hospital, columna B = Zona, columna C = Centro
    (D opcional = centro al que pertenece una unidad, p. ej. UCCU).
    Las celdas vacías en A y B repiten la fila anterior. Hay un botón para descargar una plantilla.

Perfiles: «Trabajador» y «Administrador». Un administrador accede a las mismas opciones que la cuenta
principal «Administrador» (que no se puede borrar ni desactivar). Los datos de los usuarios se exportan a
Excel desde el menú «Trabajadores».

Registro de actividad (solo administradores): guarda la fecha y hora de cada acceso y todos los cambios que
cada usuario realiza en él (visitas, vacaciones, contraseñas, usuarios, estructura, exportaciones) y se
exporta a Excel desde el menú «Registro de actividad».

Usuarios: cada trabajador puede tener asignados varios Distritos/Hospitales (se editan cuando se quiera).
Las iniciales (hasta 3 caracteres) no pueden repetirse entre usuarios.
Las visitas se pueden exportar a Excel desde «Visitas registradas» / «Mis visitas».

Vacaciones: se anotan en el menú «Vacaciones» y aparecen como «V» (con el color del trabajador)
en su fila del Excel, solo de lunes a viernes. No se pueden anotar visitas en días de vacaciones.

Acceso inicial del administrador:  Administrador / Finidi.2026
Contraseña inicial de los trabajadores:  Prueba.2026

Si se olvida la contraseña del administrador:
    python RegistroVisitas2.py --reset-admin
"""
import calendar
import contextlib
import datetime as dt
import hashlib
import hmac
import os
import re
import secrets
import sqlite3
import subprocess
import sys
import unicodedata
from collections import Counter, defaultdict

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, colorchooser
except ImportError:
    sys.stderr.write("Este equipo no tiene tkinter instalado.\n"
                     "En Windows/macOS reinstale Python marcando la opción «tcl/tk»; "
                     "en Linux:  sudo apt install python3-tk\n")
    sys.exit(1)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write("Falta la librería openpyxl. Instálela con:  pip install openpyxl\n")
    sys.exit(1)

# --------------------------------------------------------------------------
#  CONFIGURACIÓN
# --------------------------------------------------------------------------
APP_NAME = "Registro de Visitas"
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "registro_visitas.db")

ADMIN_USER = "Administrador"
ADMIN_PASS = "Finidi.2026"
DEFAULT_USER_PASS = "Prueba.2026"
MIN_PASS_LEN = 8
MAX_INITIALS = 3          # máximo de caracteres de las iniciales de un trabajador

# True  -> los trabajadores ven las estadísticas de todo su distrito
# False -> los trabajadores solo ven sus propias estadísticas
WORKERS_SEE_ALL_STATS = False

MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
DIAS = ["L", "M", "X", "J", "V", "S", "D"]          # lunes ... domingo

# Colores para los trabajadores (sin rojo: el rojo está reservado a los fines de semana)
PALETTE = ["00B050", "FFFF00", "EE2EBC", "5B9BD5", "FFC000", "ED7D31", "B4A7D6",
           "70CCD6", "A9D18E", "F4B183", "C9A0DC", "9DC3E6", "FFD966", "8FD14F",
           "D9A441", "7FB5B5"]
PARTICLES = {"de", "del", "la", "las", "los", "el", "y", "san", "santa", "da", "do", "dos", "di",
             "van", "von"}

# Registro de actividad (log): categorías y cuáles cuentan como «cambios» dentro de cada acceso
AUDIT_CATEGORIES = ["Acceso", "Visitas", "Vacaciones", "Contraseñas", "Usuarios", "Estructura", "Exportaciones"]
AUDIT_CHANGE_CATEGORIES = ("Visitas", "Vacaciones", "Contraseñas", "Usuarios", "Estructura")

COLOR_WEEKEND = "FF0000"
COLOR_TITLE = "A8D08D"
COLOR_ZONE = "BFBFBF"
COLOR_HEAD = "0F4A2B"      # cabeceras de las exportaciones (verde del logo)


class DataError(Exception):
    """Error de validación con un mensaje apto para mostrar al usuario."""


class DataWarning(DataError):
    """Igual que DataError, pero se muestra como advertencia (p. ej. iniciales repetidas)."""


def in_clause(column, value):
    """Condición SQL para un id o una lista de ids. None/0 = sin filtro; lista vacía = sin resultados."""
    if value is None or (not isinstance(value, (list, tuple, set, frozenset)) and not value):
        return "", []
    if isinstance(value, (list, tuple, set, frozenset)):
        vals = list(value)
        if not vals:
            return "0=1", []
        return f"{column} IN ({','.join('?' * len(vals))})", vals
    return f"{column}=?", [value]


# --------------------------------------------------------------------------
#  UTILIDADES
# --------------------------------------------------------------------------
def strip_accents(s):
    return "".join(ch for ch in unicodedata.normalize("NFD", s or "") if unicodedata.category(ch) != "Mn")


def clean(s):
    return re.sub(r"\s+", " ", (s or "").strip())


def slug(s):
    s = re.sub(r"[^A-Z0-9]+", "_", strip_accents(s).upper()).strip("_")
    return s or "DISTRITO"


def fmt_date(value):
    """'2026-09-03' (o date) -> '03/09/2026'"""
    if isinstance(value, str):
        value = dt.date.fromisoformat(value)
    return value.strftime("%d/%m/%Y")


def parse_date(text):
    """Admite dd/mm/aaaa, dd-mm-aaaa o dd.mm.aaaa. Devuelve date o None."""
    m = re.fullmatch(r"\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})\s*", text or "")
    if not m:
        return None
    try:
        return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def text_color_for(hex_color):
    """Negro o blanco según el brillo del color de fondo (RRGGBB)."""
    h = (hex_color or "FFFFFF").lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "000000" if (0.299 * r + 0.587 * g + 0.114 * b) > 140 else "FFFFFF"


def alpha_key(text):
    """Clave de ordenación alfabética sin distinguir mayúsculas ni tildes."""
    return strip_accents(text or "").casefold()


def sort_centers_alpha(rows):
    """Centros por zona y nombre; cada unidad (UCCU...) queda justo debajo de su centro."""
    def key(r):
        top_name = r["parent_name"] if r["parent_id"] else r["name"]
        return (alpha_key(r["district_name"]), alpha_key(r["zone_name"]), alpha_key(top_name),
                r["parent_id"] or r["id"], 1 if r["parent_id"] else 0, alpha_key(r["name"]))
    return sorted(rows, key=key)


def is_primary_admin(row):
    """La cuenta principal «Administrador» no se puede borrar, desactivar ni cambiar de perfil."""
    return bool(row) and (row["username"] or "").lower() == ADMIN_USER.lower()


def default_password_for(row):
    return ADMIN_PASS if is_primary_admin(row) else DEFAULT_USER_PASS


def center_label(name, parent_name=None):
    return f"{parent_name} › {name}" if parent_name else name


def hash_password(pw, iterations=200_000):
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, iterations)
    return f"pbkdf2${iterations}${salt.hex()}${digest.hex()}"


def check_password(pw, stored):
    try:
        _, iterations, salt, digest = stored.split("$")
        calc = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), int(iterations))
        return hmac.compare_digest(calc.hex(), digest)
    except Exception:
        return False


def validate_password(pw):
    if len(pw or "") < MIN_PASS_LEN:
        raise DataError(f"La contraseña debe tener al menos {MIN_PASS_LEN} caracteres.")


def validate_dni(text):
    """Valida DNI (8 cifras + letra) o NIE (X/Y/Z + 7 cifras + letra) con su letra de control."""
    v = re.sub(r"[\s\-]", "", (text or "")).upper()
    letters = "TRWAGMYFPDXBNJZSQVHLCKE"
    m = re.fullmatch(r"(\d{8})([A-Z])", v)
    if m:
        number = int(m.group(1))
    else:
        m = re.fullmatch(r"([XYZ])(\d{7})([A-Z])", v)
        if not m:
            raise DataError("El DNI/NIE no tiene un formato válido (ejemplo: 12345678Z).")
        number = int("XYZ".index(m.group(1)) * 10 ** 7 + int(m.group(2)))
    if letters[number % 23] != v[-1]:
        raise DataError("La letra del DNI/NIE no es correcta.")
    return v


def make_username(nombre, apellidos):
    """Formato Nombre.Primerapellido (sin tildes ni espacios)."""
    def cap(t):
        return t[:1].upper() + t[1:].lower()

    first = re.sub(r"[^A-Za-z0-9]", "", strip_accents(clean(nombre)).split(" ")[0])
    toks = strip_accents(clean(apellidos)).split(" ")
    parts = []
    for tok in toks:
        parts.append(tok)
        if tok.lower() not in PARTICLES:
            break
    ap = "".join(cap(re.sub(r"[^A-Za-z0-9]", "", t)) for t in parts)
    return f"{cap(first)}.{ap}"


def suggest_initials(nombre, taken):
    """Propone iniciales (máx. MAX_INITIALS caracteres) que no estén en 'taken'."""
    words = [w for w in strip_accents(clean(nombre)).upper().split(" ") if w]
    taken = {strip_accents(t).upper() for t in taken}
    cands = []
    if words:
        if len(words) > 1:
            cands.append("".join(w[0] for w in words)[:MAX_INITIALS])
        cands.append(words[0][0])
        for k in range(2, MAX_INITIALS + 1):
            cands.append(words[0][:k])
    for c in cands:
        if c and c not in taken:
            return c
    base = words[0] if words else "X"
    for stem in (base[:2], base[:1]):
        for i in range(2, 100):
            c = f"{stem}{i}"
            if len(c) <= MAX_INITIALS and c not in taken:
                return c
    raise DataError("No se pueden generar unas iniciales automáticas; indíquelas manualmente.")


def workdays(d1, d2):
    """Días laborables (lunes a viernes) entre dos fechas, ambas incluidas."""
    return sum(1 for i in range((d2 - d1).days + 1) if (d1 + dt.timedelta(days=i)).weekday() < 5)


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                                    # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:                                      # pragma: no cover
        raise DataError(f"No se pudo abrir:\n{path}\n\n{exc}")


# --------------------------------------------------------------------------
#  BASE DE DATOS
# --------------------------------------------------------------------------
class Store:
    """Acceso a datos (SQLite). Toda la lógica de negocio y validaciones está aquí."""

    def __init__(self, path=None):
        self.path = path or DB_PATH
        self.con = sqlite3.connect(self.path, timeout=15)
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA foreign_keys=ON")
        self.actor = None          # (id, usuario, nombre) de quien ha iniciado sesión
        self.session = None        # nº de acceso (id de la fila «Inicio de sesión»)
        self._mute = 0
        self._create()

    # ---- infraestructura ------------------------------------------------
    def _q(self, sql, params=()):
        return self.con.execute(sql, params).fetchall()

    def _one(self, sql, params=()):
        return self.con.execute(sql, params).fetchone()

    def _x(self, sql, params=()):
        cur = self.con.execute(sql, params)
        self.con.commit()
        return cur

    def _create(self):
        self.con.executescript("""
        CREATE TABLE IF NOT EXISTS districts(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE COLLATE NOCASE);
        CREATE TABLE IF NOT EXISTS zones(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            district_id INTEGER NOT NULL REFERENCES districts(id),
            name TEXT NOT NULL COLLATE NOCASE,
            sort_order INTEGER NOT NULL DEFAULT 0,
            UNIQUE(district_id, name));
        CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE COLLATE NOCASE,
            first_name TEXT NOT NULL DEFAULT '',
            last_name TEXT NOT NULL DEFAULT '',
            full_name TEXT NOT NULL,
            dni TEXT,
            pw_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'worker',
            initials TEXT,
            color TEXT,
            active INTEGER NOT NULL DEFAULT 1);
        CREATE TABLE IF NOT EXISTS audit_log(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            session INTEGER,
            user_id INTEGER,
            username TEXT NOT NULL DEFAULT '',
            full_name TEXT NOT NULL DEFAULT '',
            category TEXT NOT NULL,
            action TEXT NOT NULL,
            details TEXT NOT NULL DEFAULT '');
        CREATE INDEX IF NOT EXISTS ix_audit_ts ON audit_log(ts);
        CREATE INDEX IF NOT EXISTS ix_audit_user ON audit_log(username);
        CREATE TABLE IF NOT EXISTS user_districts(
            user_id INTEGER NOT NULL REFERENCES users(id),
            district_id INTEGER NOT NULL REFERENCES districts(id),
            PRIMARY KEY(user_id, district_id));
        CREATE TABLE IF NOT EXISTS centers(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            district_id INTEGER NOT NULL REFERENCES districts(id),
            zone_id INTEGER NOT NULL REFERENCES zones(id),
            parent_id INTEGER REFERENCES centers(id),
            name TEXT NOT NULL,
            sort_order INTEGER NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS visits(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            district_id INTEGER NOT NULL REFERENCES districts(id),
            user_id INTEGER NOT NULL REFERENCES users(id),
            center_id INTEGER NOT NULL REFERENCES centers(id),
            visit_date TEXT NOT NULL,
            observations TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            created_by INTEGER);
        CREATE TABLE IF NOT EXISTS absences(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            district_id INTEGER NOT NULL REFERENCES districts(id),
            user_id INTEGER NOT NULL REFERENCES users(id),
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            created_at TEXT NOT NULL,
            created_by INTEGER);
        CREATE INDEX IF NOT EXISTS ix_abs_user ON absences(user_id, start_date);
        CREATE INDEX IF NOT EXISTS ix_visits_date ON visits(district_id, visit_date);
        CREATE INDEX IF NOT EXISTS ix_visits_center ON visits(center_id);
        CREATE INDEX IF NOT EXISTS ix_visits_user ON visits(user_id);
        """)
        try:      # bases de datos antiguas: un único distrito por usuario -> tabla de asignaciones
            self.con.execute("INSERT OR IGNORE INTO user_districts(user_id, district_id) "
                             "SELECT id, district_id FROM users WHERE district_id IS NOT NULL")
            self.con.execute("UPDATE users SET district_id=NULL WHERE district_id IS NOT NULL")
        except sqlite3.OperationalError:
            pass
        if not self._one("SELECT 1 FROM users WHERE username=? COLLATE NOCASE", (ADMIN_USER,)):
            self.con.execute(
                "INSERT INTO users(username, full_name, pw_hash, role) VALUES(?,?,?, 'admin')",
                (ADMIN_USER, "Administrador", hash_password(ADMIN_PASS)))
        self.con.commit()

    def close(self):
        self.con.close()

    # ---- registro de actividad (log) ------------------------------------------------------
    @contextlib.contextmanager
    def muted(self):
        """Dentro de este bloque no se anotan eventos individuales (p. ej. durante una importación)."""
        self._mute += 1
        try:
            yield
        finally:
            self._mute -= 1

    def log(self, category, action, details="", actor=None):
        """Anota un evento en el log. Sin sesión iniciada no se anota y nunca interrumpe la operación."""
        if self._mute:
            return None
        try:
            if actor is None:
                if self.actor is None:
                    return None
                (uid, uname, fname), session = self.actor, self.session
            else:
                (uid, uname, fname), session = actor, None
            cur = self.con.execute(
                "INSERT INTO audit_log(ts, session, user_id, username, full_name, category, action, details) "
                "VALUES(?,?,?,?,?,?,?,?)",
                (dt.datetime.now().isoformat(timespec="seconds"), session, uid, uname or "", fname or "",
                 category, action, details or ""))
            self.con.commit()
            return cur.lastrowid
        except Exception:
            return None

    def start_session(self, user):
        """Registra el acceso de un usuario; los cambios posteriores quedan asociados a este acceso."""
        self.actor = (user["id"], user["username"], user["full_name"])
        self.session = None
        rid = self.log("Acceso", "Inicio de sesión",
                       "Perfil: " + ("Administrador" if user["role"] == "admin" else "Trabajador"))
        if rid:
            self.con.execute("UPDATE audit_log SET session=? WHERE id=?", (rid, rid))
            self.con.commit()
            self.session = rid
        return rid

    def end_session(self, reason="Cierre de sesión"):
        if self.actor:
            self.log("Acceso", "Cierre de sesión", reason)
        self.actor, self.session = None, None

    @staticmethod
    def _require_admin(requester):
        if not requester or requester["role"] != "admin":
            raise DataError("Solo el administrador puede consultar el registro de actividad.")

    def _audit_where(self, d_from, d_to, username, category, text, prefix=""):
        where, params = [], []
        if d_from:
            where.append(f"{prefix}ts >= ?")
            params.append(d_from + "T00:00:00")
        if d_to:
            where.append(f"{prefix}ts <= ?")
            params.append(d_to + "T23:59:59")
        if username:
            where.append(f"{prefix}username = ? COLLATE NOCASE")
            params.append(username)
        if category:
            where.append(f"{prefix}category = ?")
            params.append(category)
        if text:
            like = f"%{text}%"
            where.append(f"({prefix}details LIKE ? OR {prefix}action LIKE ? OR {prefix}username LIKE ? "
                         f"OR {prefix}full_name LIKE ?)")
            params += [like] * 4
        return ("WHERE " + " AND ".join(where)) if where else "", params

    def audit_entries(self, requester, d_from=None, d_to=None, username=None, category=None, text=None,
                      limit=None, order="desc"):
        self._require_admin(requester)
        where, params = self._audit_where(d_from, d_to, username, category, text)
        direction = "ASC" if order == "asc" else "DESC"
        sql = f"SELECT * FROM audit_log {where} ORDER BY id {direction}"
        if limit:
            sql += f" LIMIT {int(limit)}"
        return self._q(sql, params)

    def session_numbers(self, requester):
        """Numeración correlativa de los accesos (1, 2, 3...): {id interno del acceso: número}."""
        self._require_admin(requester)
        rows = self._q("SELECT id FROM audit_log WHERE action='Inicio de sesión' ORDER BY id")
        return {r["id"]: i for i, r in enumerate(rows, start=1)}

    def audit_users(self, requester):
        self._require_admin(requester)
        rows = self._q("SELECT username, MAX(full_name) AS full_name FROM audit_log WHERE username<>'' "
                       "GROUP BY username COLLATE NOCASE")
        return sorted(rows, key=lambda r: alpha_key(r["full_name"] or r["username"]))

    def audit_sessions(self, requester, d_from=None, d_to=None, username=None):
        """Un registro por cada acceso: usuario, inicio, cierre y número de cambios realizados en él."""
        self._require_admin(requester)
        where, params = self._audit_where(d_from, d_to, username, None, None, prefix="l.")
        marks = ",".join("?" * len(AUDIT_CHANGE_CATEGORIES))
        extra = "l.action='Inicio de sesión'" + (" AND " + where[6:] if where else "")
        sql = (f"SELECT l.id AS session, l.ts AS start_ts, l.username, l.full_name, "
               f"(SELECT MAX(o.ts) FROM audit_log o WHERE o.session=l.id AND o.action='Cierre de sesión') AS end_ts, "
               f"(SELECT COUNT(*) FROM audit_log c WHERE c.session=l.id AND c.category IN ({marks})) AS changes "
               f"FROM audit_log l WHERE {extra} ORDER BY l.id")
        return self._q(sql, list(AUDIT_CHANGE_CATEGORIES) + params)

    # ---- autenticación ----------------------------------------------------
    def authenticate(self, username, password):
        uname = strip_accents(clean(username)).replace(" ", "")
        row = self._one("SELECT * FROM users WHERE username = ? COLLATE NOCASE", (uname,))
        if not row or not check_password(password, row["pw_hash"]):
            self.log("Acceso", "Acceso fallido", "Usuario o contraseña incorrectos",
                     actor=(row["id"] if row else None, row["username"] if row else uname,
                            row["full_name"] if row else ""))
            return None, "Usuario o contraseña incorrectos."
        if not row["active"]:
            self.log("Acceso", "Acceso denegado", "Usuario dado de baja",
                     actor=(row["id"], row["username"], row["full_name"]))
            return None, "Este usuario está dado de baja. Contacte con el administrador."
        return row, ""

    def uses_default_password(self, user):
        return check_password(default_password_for(user), user["pw_hash"])

    def change_own_password(self, user_id, current, new, confirm):
        row = self.user(user_id)
        if not row or not check_password(current, row["pw_hash"]):
            raise DataError("La contraseña actual no es correcta.")
        if new != confirm:
            raise DataError("La nueva contraseña y su confirmación no coinciden.")
        validate_password(new)
        if check_password(new, row["pw_hash"]):
            raise DataError("La nueva contraseña debe ser distinta de la actual.")
        self._x("UPDATE users SET pw_hash=? WHERE id=?", (hash_password(new), user_id))
        self.log("Contraseñas", "Cambio de contraseña propia", f"Usuario: {row['username']}")

    def set_password(self, user_id, new):
        validate_password(new)
        self._x("UPDATE users SET pw_hash=? WHERE id=?", (hash_password(new), user_id))
        row = self.user(user_id)
        self.log("Contraseñas", "Contraseña cambiada por el administrador", f"Usuario: {row['username']}")

    def reset_password(self, user_id):
        row = self.user(user_id)
        self._x("UPDATE users SET pw_hash=? WHERE id=?", (hash_password(default_password_for(row)), user_id))
        self.log("Contraseñas", "Contraseña restablecida a la inicial", f"Usuario: {row['username']}")

    # ---- distritos --------------------------------------------------------
    def districts(self):
        return sorted(self._q("SELECT * FROM districts"), key=lambda r: alpha_key(r["name"]))

    def district(self, district_id):
        return self._one("SELECT * FROM districts WHERE id=?", (district_id,))

    def add_district(self, name):
        name = clean(name)
        if not name:
            raise DataError("Escriba el nombre del Distrito/Hospital.")
        if self._one("SELECT 1 FROM districts WHERE name=?", (name,)):
            raise DataError("Ya existe un Distrito/Hospital con ese nombre.")
        rid = self._x("INSERT INTO districts(name) VALUES(?)", (name,)).lastrowid
        self.log("Estructura", "Distrito/Hospital creado", f"«{name}»")
        return rid

    def rename_district(self, district_id, name):
        name = clean(name)
        old = self.district(district_id)
        if not name:
            raise DataError("Escriba el nombre del Distrito/Hospital.")
        if self._one("SELECT 1 FROM districts WHERE name=? AND id<>?", (name, district_id)):
            raise DataError("Ya existe un Distrito/Hospital con ese nombre.")
        self._x("UPDATE districts SET name=? WHERE id=?", (name, district_id))
        self.log("Estructura", "Distrito/Hospital renombrado", f"«{old['name']}» → «{name}»")
        # los Excel existentes pasan a llamarse con el nuevo nombre
        prefix_old = f"REGISTRO_VISITAS_{slug(old['name'])}_"
        prefix_new = f"REGISTRO_VISITAS_{slug(name)}_"
        if prefix_old != prefix_new:
            for fn in os.listdir(BASE_DIR):
                if fn.startswith(prefix_old) and fn.lower().endswith(".xlsx"):
                    try:
                        os.replace(os.path.join(BASE_DIR, fn),
                                   os.path.join(BASE_DIR, prefix_new + fn[len(prefix_old):]))
                    except OSError:
                        pass

    def delete_district(self, district_id):
        if self._one("SELECT 1 FROM zones WHERE district_id=?", (district_id,)) or \
           self._one("SELECT 1 FROM user_districts ud JOIN users u ON u.id = ud.user_id "
                     "WHERE ud.district_id=? AND u.role='worker'", (district_id,)):
            raise DataError("El Distrito/Hospital tiene zonas, centros o trabajadores.\n"
                            "Elimínelos primero para poder borrarlo.")
        d = self.district(district_id)
        self._x("DELETE FROM districts WHERE id=?", (district_id,))
        if d:
            self.log("Estructura", "Distrito/Hospital eliminado", f"«{d['name']}»")

    # ---- zonas ------------------------------------------------------------
    def zones(self, district_id, alpha=False):
        """Zonas del distrito. alpha=True: orden alfabético; si no, el orden de alta/importación (el del Excel)."""
        rows = self._q("SELECT * FROM zones WHERE district_id=? ORDER BY sort_order, id", (district_id,))
        return sorted(rows, key=lambda r: alpha_key(r["name"])) if alpha else rows

    def zone(self, zone_id):
        return self._one("SELECT * FROM zones WHERE id=?", (zone_id,))

    def add_zone(self, district_id, name):
        name = clean(name)
        if not name:
            raise DataError("Escriba el nombre de la zona.")
        if self._one("SELECT 1 FROM zones WHERE district_id=? AND name=?", (district_id, name)):
            raise DataError("Ya existe una zona con ese nombre en este distrito.")
        order = self._one("SELECT COALESCE(MAX(sort_order),0)+1 FROM zones WHERE district_id=?",
                          (district_id,))[0]
        rid = self._x("INSERT INTO zones(district_id, name, sort_order) VALUES(?,?,?)",
                      (district_id, name, order)).lastrowid
        self.log("Estructura", "Zona creada", f"Distrito: {self.district(district_id)['name']} | Zona: «{name}»")
        return rid

    def rename_zone(self, zone_id, name):
        name = clean(name)
        z = self.zone(zone_id)
        if not name:
            raise DataError("Escriba el nombre de la zona.")
        if self._one("SELECT 1 FROM zones WHERE district_id=? AND name=? AND id<>?",
                     (z["district_id"], name, zone_id)):
            raise DataError("Ya existe una zona con ese nombre en este distrito.")
        self._x("UPDATE zones SET name=? WHERE id=?", (name, zone_id))
        self.log("Estructura", "Zona renombrada",
                 f"Distrito: {self.district(z['district_id'])['name']} | «{z['name']}» → «{name}»")

    def delete_zone(self, zone_id):
        if self._one("SELECT 1 FROM centers WHERE zone_id=?", (zone_id,)):
            raise DataError("La zona tiene centros.\nElimínelos primero para poder borrar la zona.")
        try:                       # bases de datos antiguas guardaban la zona del trabajador
            self.con.execute("UPDATE users SET zone_id=NULL WHERE zone_id=?", (zone_id,))
        except sqlite3.OperationalError:
            pass
        z = self.zone(zone_id)
        self._x("DELETE FROM zones WHERE id=?", (zone_id,))
        if z:
            self.log("Estructura", "Zona eliminada",
                     f"Distrito: {self.district(z['district_id'])['name']} | Zona: «{z['name']}»")

    # ---- centros ----------------------------------------------------------
    _CENTER_SQL = """
        SELECT c.id, c.district_id, c.zone_id, c.parent_id, c.name, c.sort_order,
               z.name AS zone_name, p.name AS parent_name, d.name AS district_name
        FROM centers c
        JOIN zones z ON z.id = c.zone_id
        JOIN districts d ON d.id = c.district_id
        LEFT JOIN centers p ON p.id = c.parent_id
        {where}
        ORDER BY d.name COLLATE NOCASE, z.sort_order, z.id,
                 COALESCE(p.sort_order, c.sort_order), COALESCE(p.id, c.id),
                 (c.parent_id IS NOT NULL), c.sort_order, c.id"""

    def centers(self, district_id=None, zone_id=None, alpha=False):
        where, params = [], []
        cond, p = in_clause("c.district_id", district_id)
        if cond:
            where.append(cond)
            params += p
        if zone_id:
            where.append("c.zone_id=?")
            params.append(zone_id)
        sql = self._CENTER_SQL.format(where=("WHERE " + " AND ".join(where)) if where else "")
        rows = self._q(sql, params)
        return sort_centers_alpha(rows) if alpha else rows

    def center(self, center_id):
        rows = self._q(self._CENTER_SQL.format(where="WHERE c.id=?"), (center_id,))
        return rows[0] if rows else None

    def _check_center(self, district_id, zone_id, parent_id, name, exclude_id=None):
        name = clean(name)
        if not name:
            raise DataError("Escriba el nombre del centro.")
        z = self.zone(zone_id) if zone_id else None
        if not z or z["district_id"] != district_id:
            raise DataError("Seleccione una zona válida del distrito.")
        if parent_id:
            p = self.center(parent_id)
            if not p or p["parent_id"] is not None or p["id"] == exclude_id:
                raise DataError("El centro principal seleccionado no es válido.")
            zone_id = p["zone_id"]
        dup = self._one("SELECT 1 FROM centers WHERE zone_id=? AND COALESCE(parent_id,0)=? "
                        "AND name=? COLLATE NOCASE AND id<>?",
                        (zone_id, parent_id or 0, name, exclude_id or 0))
        if dup:
            raise DataError("Ya existe un centro con ese nombre en esa zona.")
        return name, zone_id

    def add_center(self, district_id, zone_id, name, parent_id=None):
        name, zone_id = self._check_center(district_id, zone_id, parent_id, name)
        order = self._one("SELECT COALESCE(MAX(sort_order),0)+1 FROM centers WHERE district_id=?",
                          (district_id,))[0]
        rid = self._x("INSERT INTO centers(district_id, zone_id, parent_id, name, sort_order) "
                      "VALUES(?,?,?,?,?)", (district_id, zone_id, parent_id, name, order)).lastrowid
        c = self.center(rid)
        self.log("Estructura", "Unidad creada" if parent_id else "Centro creado",
                 f"Distrito: {c['district_name']} | Zona: {c['zone_name']} | «{center_label(c['name'], c['parent_name'])}»")
        return rid

    def update_center(self, center_id, name, zone_id, parent_id=None):
        c = self.center(center_id)
        if parent_id and self._one("SELECT 1 FROM centers WHERE parent_id=?", (center_id,)):
            raise DataError("Este centro tiene unidades dentro; no puede pasar a ser unidad de otro.")
        name, zone_id = self._check_center(c["district_id"], zone_id, parent_id, name, center_id)
        self._x("UPDATE centers SET name=?, zone_id=?, parent_id=? WHERE id=?",
                (name, zone_id, parent_id, center_id))
        self._x("UPDATE centers SET zone_id=? WHERE parent_id=?", (zone_id, center_id))
        a = self.center(center_id)
        self.log("Estructura", "Centro modificado",
                 f"Distrito: {c['district_name']} | ANTES: «{center_label(c['name'], c['parent_name'])}» "
                 f"(zona {c['zone_name']}) | DESPUÉS: «{center_label(a['name'], a['parent_name'])}» (zona {a['zone_name']})")

    def center_visit_count(self, center_id):
        return self._one("SELECT COUNT(*) FROM visits WHERE center_id=? OR center_id IN "
                         "(SELECT id FROM centers WHERE parent_id=?)", (center_id, center_id))[0]

    def delete_center(self, center_id):
        c = self.center(center_id)
        n = self.center_visit_count(center_id)
        ids = [center_id] + [r[0] for r in self._q("SELECT id FROM centers WHERE parent_id=?", (center_id,))]
        marks = ",".join("?" * len(ids))
        self.con.execute(f"DELETE FROM visits WHERE center_id IN ({marks})", ids)
        self.con.execute(f"DELETE FROM centers WHERE id IN ({marks})", ids)
        self.con.commit()
        if c:
            self.log("Estructura", "Centro eliminado",
                     f"Distrito: {c['district_name']} | Zona: {c['zone_name']} | "
                     f"«{center_label(c['name'], c['parent_name'])}» | Unidades eliminadas: {len(ids) - 1} | "
                     f"Visitas borradas: {n}")
        return n

    def visit_counts_by_center(self, district_id):
        rows = self._q("SELECT center_id, COUNT(*) AS n FROM visits WHERE district_id=? GROUP BY center_id",
                       (district_id,))
        return {r["center_id"]: r["n"] for r in rows}

    # ---- trabajadores -------------------------------------------------------
    _USER_SQL = """
        SELECT u.*, (SELECT GROUP_CONCAT(d.name, ' | ')
                     FROM user_districts ud JOIN districts d ON d.id = ud.district_id
                     WHERE ud.user_id = u.id) AS district_names
        FROM users u
        {where}
        ORDER BY u.full_name COLLATE NOCASE"""

    def user(self, user_id):
        rows = self._q(self._USER_SQL.format(where="WHERE u.id=?"), (user_id,))
        return rows[0] if rows else None

    def users(self, role=None, district_id=None, only_active=False):
        """Usuarios ordenados alfabéticamente. role: 'worker', 'admin' o None (todos)."""
        where, params = ["1=1"], []
        if role:
            where.append("u.role=?")
            params.append(role)
        cond, p = in_clause("ud.district_id", district_id)
        if cond:
            where.append(f"u.id IN (SELECT ud.user_id FROM user_districts ud WHERE {cond})")
            params += p
        if only_active:
            where.append("u.active=1")
        rows = self._q(self._USER_SQL.format(where="WHERE " + " AND ".join(where)), params)
        return sorted(rows, key=lambda r: alpha_key(r["full_name"]))

    def workers(self, district_id=None, only_active=False):
        """Trabajadores asignados a un distrito (o a cualquiera de una lista de distritos)."""
        return self.users("worker", district_id, only_active)

    def visit_counts_by_user(self):
        return {r[0]: r[1] for r in self._q("SELECT user_id, COUNT(*) FROM visits GROUP BY user_id")}

    def users_by_ids(self, ids):
        ids = list(ids)
        if not ids:
            return []
        rows = self._q(self._USER_SQL.format(where=f"WHERE u.id IN ({','.join('?' * len(ids))})"), ids)
        return sorted(rows, key=lambda r: alpha_key(r["full_name"]))

    def user_districts(self, user_id):
        return self._q("SELECT d.* FROM user_districts ud JOIN districts d ON d.id = ud.district_id "
                       "WHERE ud.user_id=? ORDER BY d.name COLLATE NOCASE", (user_id,))

    def user_district_ids(self, user_id):
        return [r["id"] for r in self.user_districts(user_id)]

    def set_user_districts(self, user_id, district_ids):
        ids = sorted({int(i) for i in district_ids})
        self.con.execute("DELETE FROM user_districts WHERE user_id=?", (user_id,))
        self.con.executemany("INSERT INTO user_districts(user_id, district_id) VALUES(?,?)",
                             [(user_id, i) for i in ids])
        self.con.commit()

    # --- iniciales: únicas en toda la aplicación, de 1 a MAX_INITIALS caracteres ---
    @staticmethod
    def _ini_key(ini):
        return strip_accents(ini or "").upper().replace(" ", "")

    def initials_owner(self, ini, exclude_id=None):
        """Nombre del trabajador que ya usa esas iniciales (o None si están libres)."""
        key = self._ini_key(ini)
        if not key:
            return None
        for r in self._q("SELECT full_name, initials FROM users WHERE role='worker' AND id<>?", (exclude_id or 0,)):
            if self._ini_key(r["initials"]) == key:
                return r["full_name"]
        return None

    def check_initials(self, ini, exclude_id=None):
        ini = clean(ini).upper().replace(" ", "")
        if not ini:
            raise DataError("Indique las iniciales del trabajador.")
        if len(ini) > MAX_INITIALS:
            raise DataError(f"Las iniciales pueden tener como máximo {MAX_INITIALS} caracteres.")
        if not re.fullmatch(r"[^\W_]+", ini):
            raise DataError("Las iniciales solo pueden contener letras o números.")
        owner = self.initials_owner(ini, exclude_id)
        if owner:
            raise DataWarning(f"Las iniciales «{ini}» ya las utiliza {owner}.\n\n"
                              "Cada usuario debe tener unas iniciales distintas. Elija otras.")
        return ini

    def _taken_initials(self):
        return {self._ini_key(r[0]) for r in self._q("SELECT initials FROM users WHERE role='worker'")}

    def _free_color(self, district_ids, exclude_id=None):
        cond, p = in_clause("ud.district_id", list(district_ids))
        used = {(r[0] or "").upper() for r in self._q(
            f"SELECT DISTINCT u.color FROM users u JOIN user_districts ud ON ud.user_id = u.id "
            f"WHERE u.role='worker' AND u.id<>? AND {cond}", [exclude_id or 0] + p)}
        for c in PALETTE:
            if c not in used:
                return c
        return PALETTE[len(used) % len(PALETTE)]

    def _check_worker_common(self, nombre, apellidos, dni, district_ids, exclude_id=None, require_districts=True):
        nombre, apellidos = clean(nombre), clean(apellidos)
        if not nombre or not apellidos:
            raise DataError("Indique el nombre y los apellidos.")
        dni = validate_dni(dni)
        if self._one("SELECT 1 FROM users WHERE dni=? AND id<>?", (dni, exclude_id or 0)):
            raise DataError("Ya existe un usuario con ese DNI/NIE.")
        ids = []
        if require_districts:
            ids = sorted({int(i) for i in (district_ids or [])})
            if not ids:
                raise DataError("Seleccione al menos un Distrito/Hospital.")
            if any(not self.district(i) for i in ids):
                raise DataError("Alguno de los distritos seleccionados ya no existe.")
        return nombre, apellidos, dni, ids

    @staticmethod
    def _valid_color(color):
        c = (color or "").lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", c):
            raise DataError("El color no es válido.")
        return c

    def _other_active_admins(self, user_id):
        return self._one("SELECT COUNT(*) FROM users WHERE role='admin' AND active=1 AND id<>?", (user_id,))[0]

    def add_user(self, role, nombre, apellidos, dni, district_ids=(), initials="", color=""):
        """Crea un usuario con perfil 'worker' (Trabajador) o 'admin' (Administrador)."""
        if role not in ("worker", "admin"):
            raise DataError("Seleccione el perfil del usuario.")
        is_worker = role == "worker"
        nombre, apellidos, dni, ids = self._check_worker_common(nombre, apellidos, dni, district_ids,
                                                                require_districts=is_worker)
        base = make_username(nombre, apellidos)
        username, n = base, 2
        while self._one("SELECT 1 FROM users WHERE username=? COLLATE NOCASE", (username,)):
            username = f"{base}{n}"
            n += 1
        ini = col = None
        if is_worker:
            ini = self.check_initials(initials) if clean(initials) else \
                suggest_initials(nombre, self._taken_initials())
            col = self._valid_color(color) if color else self._free_color(ids)
        cur = self._x(
            "INSERT INTO users(username, first_name, last_name, full_name, dni, pw_hash, role, "
            "initials, color, active) VALUES(?,?,?,?,?,?,?,?,?,1)",
            (username, nombre, apellidos, f"{nombre} {apellidos}", dni, hash_password(DEFAULT_USER_PASS),
             role, ini, col))
        if is_worker:
            self.set_user_districts(cur.lastrowid, ids)
        new = self.user(cur.lastrowid)
        self.log("Usuarios", "Usuario creado",
                 f"Usuario: {username} | Nombre: {nombre} {apellidos} | Perfil: {'Trabajador' if is_worker else 'Administrador'}"
                 + (f" | Iniciales: {ini} | Distritos: {new['district_names']}" if is_worker else ""))
        return cur.lastrowid, username

    def add_worker(self, nombre, apellidos, dni, district_ids, initials="", color=""):
        return self.add_user("worker", nombre, apellidos, dni, district_ids, initials, color)

    def update_user(self, user_id, role, nombre, apellidos, dni, district_ids,
                    initials, color, username, active, acting_user_id=None):
        u = self.user(user_id)
        if not u:
            raise DataError("Usuario no encontrado.")
        if is_primary_admin(u):
            raise DataError("La cuenta principal «Administrador» no se puede modificar "
                            "(solo se puede cambiar su contraseña).")
        if role not in ("worker", "admin"):
            raise DataError("Seleccione el perfil del usuario.")
        if user_id == acting_user_id and (role != "admin" or not active):
            raise DataError("No puede quitarse a sí mismo el perfil de administrador ni darse de baja.")
        if u["role"] == "admin" and (role != "admin" or not active) and self._other_active_admins(user_id) == 0:
            raise DataError("Debe quedar al menos un administrador activo.")
        is_worker = role == "worker"
        nombre, apellidos, dni, ids = self._check_worker_common(nombre, apellidos, dni, district_ids, user_id,
                                                                require_districts=is_worker)
        uname = strip_accents(clean(username)).replace(" ", "")
        if not re.fullmatch(r"[A-Za-z0-9._\-]+", uname):
            raise DataError("El nombre de usuario solo puede contener letras, números, punto y guion.")
        if self._one("SELECT 1 FROM users WHERE username=? COLLATE NOCASE AND id<>?", (uname, user_id)):
            raise DataError("Ese nombre de usuario ya existe.")
        ini, col = u["initials"], u["color"]       # un administrador conserva sus datos de trabajador
        if is_worker:
            ini = self.check_initials(initials, user_id) if clean(initials) else \
                suggest_initials(nombre, self._taken_initials())
            col = self._valid_color(color) if color else (u["color"] or self._free_color(ids, user_id))
        self._x("UPDATE users SET username=?, first_name=?, last_name=?, full_name=?, dni=?, role=?, "
                "initials=?, color=?, active=? WHERE id=?",
                (uname, nombre, apellidos, f"{nombre} {apellidos}", dni, role, ini, col,
                 1 if active else 0, user_id))
        if is_worker:
            self.set_user_districts(user_id, ids)
        new = self.user(user_id)
        changes = []
        for label, a, b in (("Usuario", u["username"], new["username"]), ("Nombre", u["full_name"], new["full_name"]),
                            ("Perfil", "Administrador" if u["role"] == "admin" else "Trabajador",
                             "Administrador" if new["role"] == "admin" else "Trabajador"),
                            ("Estado", "Activo" if u["active"] else "Baja", "Activo" if new["active"] else "Baja")):
            if a != b:
                changes.append(f"{label}: {a} → {b}")
        if (u["dni"] or "") != (new["dni"] or ""):
            changes.append("DNI/NIE modificado")
        if is_worker:
            for label, a, b in (("Iniciales", u["initials"] or "-", new["initials"] or "-"),
                                ("Color", u["color"] or "-", new["color"] or "-"),
                                ("Distritos", u["district_names"] or "-", new["district_names"] or "-")):
                if a != b:
                    changes.append(f"{label}: {a} → {b}")
        self.log("Usuarios", "Usuario modificado",
                 f"Usuario: {u['username']} | " + ("; ".join(changes) if changes else "Sin cambios en los datos"))

    def update_worker(self, user_id, nombre, apellidos, dni, district_ids,
                      initials, color, username, active):
        return self.update_user(user_id, "worker", nombre, apellidos, dni, district_ids,
                                initials, color, username, active)

    def worker_visit_count(self, user_id):
        return self._one("SELECT COUNT(*) FROM visits WHERE user_id=?", (user_id,))[0]

    def delete_user(self, user_id, acting_user_id=None):
        u = self.user(user_id)
        if not u:
            raise DataError("Usuario no encontrado.")
        if is_primary_admin(u):
            raise DataError("La cuenta principal «Administrador» no se puede eliminar.")
        if user_id == acting_user_id:
            raise DataError("No puede eliminar su propia cuenta.")
        if u["role"] == "admin" and u["active"] and self._other_active_admins(user_id) == 0:
            raise DataError("Debe quedar al menos un administrador activo.")
        n = self.worker_visit_count(user_id)
        self.con.execute("DELETE FROM visits WHERE user_id=?", (user_id,))
        self.con.execute("DELETE FROM absences WHERE user_id=?", (user_id,))
        self.con.execute("DELETE FROM user_districts WHERE user_id=?", (user_id,))
        self.con.execute("DELETE FROM users WHERE id=?", (user_id,))
        self.con.commit()
        self.log("Usuarios", "Usuario eliminado",
                 f"Usuario: {u['username']} ({u['full_name']}) | Perfil: "
                 f"{'Administrador' if u['role'] == 'admin' else 'Trabajador'} | Visitas borradas: {n}")
        return n

    def delete_worker(self, user_id):
        u = self.user(user_id)
        if not u or u["role"] != "worker":
            raise DataError("Usuario no encontrado.")
        return self.delete_user(user_id)

    # ---- visitas ------------------------------------------------------------
    def add_visit(self, user_id, center_id, date_iso, observations, created_by=None):
        u = self.user(user_id)
        c = self.center(center_id)
        if not u or u["role"] != "worker":
            raise DataError("Seleccione un trabajador válido.")
        if not c:
            raise DataError("Seleccione un centro válido.")
        if c["district_id"] not in self.user_district_ids(user_id):
            raise DataError("El trabajador no tiene asignado el Distrito/Hospital de ese centro.")
        d = dt.date.fromisoformat(date_iso)
        if d.weekday() >= 5:
            raise DataError("Los sábados y domingos son días NO laborables.\nElija un día de lunes a viernes.")
        if self.on_absence(user_id, date_iso):
            raise DataError("El trabajador tiene vacaciones («V») ese día.\n"
                            "Elimine antes el periodo de vacaciones si la visita es correcta.")
        if self._one("SELECT 1 FROM visits WHERE user_id=? AND center_id=? AND visit_date=?",
                     (user_id, center_id, date_iso)):
            raise DataError("Esa visita ya está registrada (mismo trabajador, centro y fecha).")
        rid = self._x(
            "INSERT INTO visits(district_id, user_id, center_id, visit_date, observations, created_at, created_by) "
            "VALUES(?,?,?,?,?,?,?)",
            (c["district_id"], user_id, center_id, date_iso, (observations or "").strip(),
             dt.datetime.now().isoformat(timespec="seconds"), created_by)).lastrowid
        self.log("Visitas", "Visita registrada", self._visit_text(rid))
        return rid

    def _visit_text(self, visit_id):
        r = self._one("""SELECT v.visit_date, v.observations, u.full_name, d.name AS dist, c.name AS cname,
                                p.name AS pname
                         FROM visits v JOIN users u ON u.id = v.user_id JOIN centers c ON c.id = v.center_id
                         JOIN districts d ON d.id = v.district_id LEFT JOIN centers p ON p.id = c.parent_id
                         WHERE v.id=?""", (visit_id,))
        if not r:
            return ""
        t = (f"Trabajador: {r['full_name']} | Centro: {center_label(r['cname'], r['pname'])} | "
             f"Distrito: {r['dist']} | Fecha: {fmt_date(r['visit_date'])}")
        return t + (f" | Observaciones: {r['observations']}" if r["observations"] else "")

    def update_visit(self, visit_id, center_id, date_iso, observations, by_user):
        """Modifica centro, fecha u observaciones de una visita. Devuelve [(distrito, año)] a actualizar en Excel."""
        v = self.visit(visit_id)
        if not v:
            raise DataError("La visita ya no existe.")
        if by_user["role"] != "admin" and v["user_id"] != by_user["id"]:
            raise DataError("Solo puede modificar sus propias visitas.")
        c = self.center(center_id)
        if not c:
            raise DataError("Seleccione un centro válido.")
        if c["district_id"] not in self.user_district_ids(v["user_id"]):
            raise DataError("El trabajador no tiene asignado el Distrito/Hospital de ese centro.")
        d = dt.date.fromisoformat(date_iso)
        if d.weekday() >= 5:
            raise DataError("Los sábados y domingos son días NO laborables.\nElija un día de lunes a viernes.")
        if self.on_absence(v["user_id"], date_iso):
            raise DataError("El trabajador tiene vacaciones («V») ese día.")
        if self._one("SELECT 1 FROM visits WHERE user_id=? AND center_id=? AND visit_date=? AND id<>?",
                     (v["user_id"], center_id, date_iso, visit_id)):
            raise DataError("Ya existe otra visita con el mismo trabajador, centro y fecha.")
        obs = (observations or "").strip()
        if (v["center_id"], v["visit_date"], v["observations"] or "") == (center_id, date_iso, obs):
            return []
        before = self._visit_text(visit_id)
        self._x("UPDATE visits SET center_id=?, district_id=?, visit_date=?, observations=? WHERE id=?",
                (center_id, c["district_id"], date_iso, obs, visit_id))
        self.log("Visitas", "Visita modificada", f"ANTES: {before}  ||  DESPUÉS: {self._visit_text(visit_id)}")
        return [(v["district_id"], int(v["visit_date"][:4])), (c["district_id"], d.year)]

    def visit(self, visit_id):
        return self._one("SELECT * FROM visits WHERE id=?", (visit_id,))

    def delete_visit(self, visit_id, by_user):
        v = self.visit(visit_id)
        if not v:
            raise DataError("La visita ya no existe.")
        if by_user["role"] != "admin" and v["user_id"] != by_user["id"]:
            raise DataError("Solo puede eliminar sus propias visitas.")
        text = self._visit_text(visit_id)
        self._x("DELETE FROM visits WHERE id=?", (visit_id,))
        self.log("Visitas", "Visita eliminada", text)
        return v["district_id"], int(v["visit_date"][:4])

    def visits(self, district_id=None, user_id=None, d_from=None, d_to=None, center_id=None,
               order="desc", limit=None):
        sql = """
            SELECT v.id, v.visit_date, v.observations, v.created_at, v.district_id, v.user_id, v.center_id,
                   u.full_name, u.initials, u.color, d.name AS district_name,
                   z.id AS zone_id, z.name AS zone_name, c.name AS center_name, p.name AS parent_name
            FROM visits v
            JOIN users u ON u.id = v.user_id
            JOIN centers c ON c.id = v.center_id
            JOIN zones z ON z.id = c.zone_id
            JOIN districts d ON d.id = v.district_id
            LEFT JOIN centers p ON p.id = c.parent_id
            WHERE 1=1"""
        cond, params = in_clause("v.district_id", district_id)
        if cond:
            sql += " AND " + cond
        for cond, val in (("v.user_id=?", user_id),
                          ("v.visit_date>=?", d_from), ("v.visit_date<=?", d_to),
                          ("v.center_id=?", center_id)):
            if val:
                sql += " AND " + cond
                params.append(val)
        direction = "ASC" if order == "asc" else "DESC"
        sql += f" ORDER BY v.visit_date {direction}, v.id {direction}"
        if limit:
            sql += f" LIMIT {int(limit)}"
        return self._q(sql, params)

    # ---- vacaciones (la «V» del Excel) ------------------------------------------
    def on_absence(self, user_id, date_iso):
        return bool(self._one("SELECT 1 FROM absences WHERE user_id=? AND start_date<=? AND end_date>=?",
                              (user_id, date_iso, date_iso)))

    def add_absence(self, user_id, start_iso, end_iso, created_by=None):
        u = self.user(user_id)
        if not u or u["role"] != "worker":
            raise DataError("Seleccione un trabajador válido.")
        dids = self.user_district_ids(user_id)
        if not dids:
            raise DataError("El trabajador no tiene ningún Distrito/Hospital asignado.")
        d1, d2 = dt.date.fromisoformat(start_iso), dt.date.fromisoformat(end_iso)
        if d1 > d2:
            raise DataError("La fecha «Desde» no puede ser posterior a «Hasta».")
        if workdays(d1, d2) == 0:
            raise DataError("El periodo solo incluye sábados y domingos (días no laborables).")
        if self._one("SELECT 1 FROM absences WHERE user_id=? AND start_date<=? AND end_date>=?",
                     (user_id, end_iso, start_iso)):
            raise DataError("Ese periodo se solapa con otras vacaciones ya registradas del trabajador.")
        n = self._one("SELECT COUNT(*) FROM visits WHERE user_id=? AND visit_date>=? AND visit_date<=?",
                      (user_id, start_iso, end_iso))[0]
        if n:
            raise DataError(f"El trabajador tiene {n} visita(s) registradas en esas fechas.\n"
                            "Elimínelas antes de registrar las vacaciones.")
        rid = self._x("INSERT INTO absences(district_id, user_id, start_date, end_date, created_at, created_by) "
                      "VALUES(?,?,?,?,?,?)",
                      (dids[0], user_id, start_iso, end_iso,
                       dt.datetime.now().isoformat(timespec="seconds"), created_by)).lastrowid
        self.log("Vacaciones", "Vacaciones registradas",
                 f"Trabajador: {u['full_name']} | Del {fmt_date(d1)} al {fmt_date(d2)} | Días laborables: {workdays(d1, d2)}")
        return rid

    def absences(self, district_id=None, user_id=None, d_from=None, d_to=None):
        sql = """SELECT a.id, a.user_id, a.start_date, a.end_date, u.full_name, u.initials, u.color
                   FROM absences a JOIN users u ON u.id = a.user_id
                   WHERE 1=1"""
        cond, params = in_clause("ud.district_id", district_id)
        if cond:
            sql += f" AND a.user_id IN (SELECT ud.user_id FROM user_districts ud WHERE {cond})"
        for cond, val in (("a.user_id=?", user_id),
                          ("a.end_date>=?", d_from), ("a.start_date<=?", d_to)):
            if val:
                sql += " AND " + cond
                params.append(val)
        return self._q(sql + " ORDER BY a.start_date DESC, a.id DESC", params)

    def delete_absence(self, absence_id, by_user):
        a = self._one("SELECT * FROM absences WHERE id=?", (absence_id,))
        if not a:
            raise DataError("El periodo ya no existe.")
        if by_user["role"] != "admin" and a["user_id"] != by_user["id"]:
            raise DataError("Solo puede eliminar sus propias vacaciones.")
        self._x("DELETE FROM absences WHERE id=?", (absence_id,))
        who = self.user(a["user_id"])
        d1, d2 = dt.date.fromisoformat(a["start_date"]), dt.date.fromisoformat(a["end_date"])
        self.log("Vacaciones", "Vacaciones eliminadas",
                 f"Trabajador: {who['full_name'] if who else '?'} | Del {fmt_date(d1)} al {fmt_date(d2)}")
        return self.user_district_ids(a["user_id"]), list(range(int(a["start_date"][:4]), int(a["end_date"][:4]) + 1))

    def years_for(self, district_id):
        years = {int(r[0]) for r in self._q(
            "SELECT DISTINCT substr(visit_date,1,4) FROM visits WHERE district_id=?", (district_id,))}
        for r in self._q("SELECT start_date, end_date FROM absences WHERE user_id IN "
                         "(SELECT user_id FROM user_districts WHERE district_id=?)", (district_id,)):
            years.update(range(int(r[0][:4]), int(r[1][:4]) + 1))
        years.add(dt.date.today().year)
        return sorted(years)


# --------------------------------------------------------------------------
#  EXCEL
# --------------------------------------------------------------------------
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FONT_NAME = "Arial Narrow"


def _fill(hex_color):
    return PatternFill("solid", fgColor="FF" + hex_color)


def excel_path(district_name, year):
    return os.path.join(BASE_DIR, f"REGISTRO_VISITAS_{slug(district_name)}_{year}.xlsx")


def build_layout(zones, centers):
    """Filas lógicas de la columna A (zonas, centros, unidades y separadores)."""
    by_zone = defaultdict(list)
    for c in centers:
        by_zone[c["zone_id"]].append(c)
    rows, first = [], True
    for z in zones:
        zc = by_zone.get(z["id"], [])
        tops = [c for c in zc if c["parent_id"] is None]
        kids = defaultdict(list)
        for c in zc:
            if c["parent_id"] is not None:
                kids[c["parent_id"]].append(c)
        if not first:
            rows.append({"type": "blank", "name": "", "center_id": None})
        first = False
        # zona con un único centro del mismo nombre (p. ej. un edificio): una sola fila
        if len(tops) == 1 and not kids[tops[0]["id"]] and \
                clean(tops[0]["name"]).upper() == clean(z["name"]).upper():
            rows.append({"type": "zone", "name": z["name"], "center_id": tops[0]["id"]})
            continue
        rows.append({"type": "zone", "name": z["name"], "center_id": None})
        for t in tops:
            rows.append({"type": "center", "name": t["name"], "center_id": t["id"]})
            ks = kids.get(t["id"], [])
            for k in ks:
                rows.append({"type": "unit", "name": k["name"], "center_id": k["id"]})
            if ks and t is not tops[-1]:
                rows.append({"type": "blank", "name": "", "center_id": None})
    return rows


def write_month_sheet(ws, district_name, year, month, workers, layout, visits, absences=()):
    ndays = calendar.monthrange(year, month)[1]
    last_col = 1 + ndays
    first_row = 3 + len(workers)
    row_of = {}
    r = first_row
    for item in layout:
        item["row"] = r
        if item["center_id"]:
            row_of[item["center_id"]] = r
        r += 1
    last_row = max(r - 1, first_row - 1, 2)

    weekend_cols = {1 + d for d in range(1, ndays + 1) if dt.date(year, month, d).weekday() >= 5}

    f_norm = Font(name=FONT_NAME, size=11)
    f_bold = Font(name=FONT_NAME, size=11, bold=True)
    center = Alignment(horizontal="center", vertical="center")

    # --- base: bordes, fuente, alineación y fin de semana en rojo ---------------
    for row in range(1, last_row + 1):
        for col in range(1, last_col + 1):
            cell = ws.cell(row, col)
            cell.font = f_norm
            if row >= 3:
                cell.border = BORDER
            if col > 1:
                cell.alignment = center
            if col in weekend_cols:
                cell.fill = _fill(COLOR_WEEKEND)

    # --- título (A1:A2) ------------------------------------------------------------
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    t = ws.cell(1, 1, f"VISITAS {district_name.upper()}\n{MESES[month - 1]} {year}")
    t.font = Font(name=FONT_NAME, size=12, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    t.fill = _fill(COLOR_TITLE)
    for rr in (1, 2):
        ws.cell(rr, 1).border = BORDER

    # --- días del mes y letra del día de la semana -----------------------------------
    for d in range(1, ndays + 1):
        c1 = ws.cell(1, 1 + d, d)
        c1.font = f_norm
        c2 = ws.cell(2, 1 + d, DIAS[dt.date(year, month, d).weekday()])
        c2.font = Font(name=FONT_NAME, size=8, color="FF595959" if (1 + d) not in weekend_cols else "FFFFFFFF")
        if (1 + d) in weekend_cols:
            c1.font = Font(name=FONT_NAME, size=11, color="FFFFFFFF", bold=True)

    # --- leyenda de trabajadores (filas 3, 4, 5 ...) -----------------------------------
    for i, w in enumerate(workers):
        cell = ws.cell(3 + i, 1, f"{w['full_name'].upper()} ({w['initials']})")
        cell.fill = _fill(w["color"])
        cell.font = Font(name=FONT_NAME, size=11, bold=True, color="FF" + text_color_for(w["color"]))

    # --- vacaciones: «V» en la fila del trabajador, con su color (solo lunes a viernes) --------------
    row_of_worker = {w["id"]: 3 + i for i, w in enumerate(workers)}
    color_of_worker = {w["id"]: w["color"] for w in workers}
    for a in absences:
        if a["user_id"] not in row_of_worker:
            continue
        a1, a2 = dt.date.fromisoformat(a["start_date"]), dt.date.fromisoformat(a["end_date"])
        for d in range(1, ndays + 1):
            day = dt.date(year, month, d)
            if a1 <= day <= a2 and day.weekday() < 5:
                cell = ws.cell(row_of_worker[a["user_id"]], 1 + d, "V")
                col = color_of_worker[a["user_id"]]
                cell.fill = _fill(col)
                cell.font = Font(name=FONT_NAME, size=11, bold=True, color="FF" + text_color_for(col))

    # --- zonas, centros y unidades ----------------------------------------------------------
    for item in layout:
        cell = ws.cell(item["row"], 1)
        if item["type"] == "zone":
            cell.value = item["name"]
            cell.font = Font(name=FONT_NAME, size=12, bold=True)
            cell.fill = _fill(COLOR_ZONE)
        elif item["type"] == "center":
            cell.value = item["name"]
            cell.font = f_bold
        elif item["type"] == "unit":
            cell.value = item["name"]
            cell.font = f_norm

    # --- visitas -------------------------------------------------------------------------------
    cells = defaultdict(list)
    for v in visits:
        d = dt.date.fromisoformat(v["visit_date"])
        if d.year == year and d.month == month and v["center_id"] in row_of:
            cells[(row_of[v["center_id"]], d.day)].append(v)
    for (row, day), vs in cells.items():
        vs = sorted(vs, key=lambda x: x["full_name"].lower())
        cell = ws.cell(row, 1 + day)
        cell.value = "/".join(v["initials"] for v in vs)
        cell.fill = _fill(vs[0]["color"])
        cell.font = Font(name=FONT_NAME, size=11, color="FF" + text_color_for(vs[0]["color"]))
        cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        notes = [f"{v['full_name']}: {v['observations']}" for v in vs if v["observations"]]
        if notes:
            cm = Comment("\n".join(notes), "Registro de Visitas")
            cm.width, cm.height = 280, 20 + 18 * sum(1 + len(n) // 40 for n in notes)
            cell.comment = cm

    # --- formato de hoja ---------------------------------------------------------------------------
    ws.column_dimensions["A"].width = 48.44
    for col in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 4.66
    for row in range(1, last_row + 1):
        ws.row_dimensions[row].height = 16 if row <= 2 else 14.25
    ws.freeze_panes = "B3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"


def write_log_sheet(ws, visits):
    heads = ["Fecha", "Trabajador", "Iniciales", "Zona", "Centro", "Observaciones"]
    widths = [12, 30, 10, 34, 46, 70]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(1, i, h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        c.fill = _fill(COLOR_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, v in enumerate(visits, start=2):
        vals = [dt.date.fromisoformat(v["visit_date"]), v["full_name"], v["initials"], v["zone_name"],
                center_label(v["center_name"], v["parent_name"]), v["observations"]]
        for i, val in enumerate(vals, start=1):
            c = ws.cell(r, i, val)
            c.font = Font(name="Arial", size=10)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(i == 6),
                                    horizontal="center" if i in (1, 3) else None)
            if i == 1:
                c.number_format = "DD/MM/YYYY"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{max(2, len(visits) + 1)}"


def write_absence_sheet(ws, absences):
    heads = ["Trabajador", "Iniciales", "Desde", "Hasta", "Días laborables"]
    widths = [34, 10, 13, 13, 17]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(1, i, h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        c.fill = _fill(COLOR_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, a in enumerate(sorted(absences, key=lambda x: (x["full_name"].lower(), x["start_date"])), start=2):
        d1, d2 = dt.date.fromisoformat(a["start_date"]), dt.date.fromisoformat(a["end_date"])
        for i, val in enumerate([a["full_name"], a["initials"], d1, d2, workdays(d1, d2)], start=1):
            c = ws.cell(r, i, val)
            c.font = Font(name="Arial", size=10)
            c.border = BORDER
            if i in (3, 4):
                c.number_format = "DD/MM/YYYY"
            if i > 1:
                c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def write_district_excel(store, district_id, year):
    d = store.district(district_id)
    layout = build_layout(store.zones(district_id), store.centers(district_id))
    visits = [dict(v) for v in store.visits(district_id=district_id, d_from=f"{year}-01-01",
                                            d_to=f"{year}-12-31", order="asc")]
    absences = [dict(a) for a in store.absences(district_id=district_id, d_from=f"{year}-01-01",
                                                 d_to=f"{year}-12-31")]
    used = {v["user_id"] for v in visits} | {a["user_id"] for a in absences}
    workers = [w for w in store.workers(district_id) if w["active"] or w["id"] in used]
    known = {w["id"] for w in workers}
    workers += [u for u in store.users_by_ids(used - known)]
    workers = sorted(workers, key=lambda w: w["full_name"].lower())

    wb = Workbook()
    wb.remove(wb.active)
    for m in range(1, 13):
        ws = wb.create_sheet(MESES[m - 1])
        write_month_sheet(ws, d["name"], year, m, workers, [dict(x) for x in layout], visits, absences)
    write_log_sheet(wb.create_sheet("REGISTRO"), visits)
    write_absence_sheet(wb.create_sheet("VACACIONES"), absences)

    path = excel_path(d["name"], year)
    tmp = path + ".tmp"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise
    return path


def sync_excel(store, district_id, years=None):
    """Regenera el Excel del distrito. Devuelve (ok, mensaje)."""
    d = store.district(district_id)
    if d is None:
        return True, ""
    errors = []
    for y in sorted(set(years or store.years_for(district_id))):
        try:
            write_district_excel(store, district_id, y)
        except PermissionError:
            errors.append(f"«{os.path.basename(excel_path(d['name'], y))}» está abierto en otro programa.\n"
                          "Ciérrelo: los datos están guardados y el Excel se actualizará en el próximo cambio "
                          "(o desde «Archivos Excel > Regenerar»).")
        except OSError as exc:
            errors.append(f"No se pudo escribir el Excel de {y}: {exc}")
    return (not errors), "\n\n".join(errors)


# --------------------------------------------------------------------------
#  IMPORTAR ESTRUCTURA DESDE UN EXCEL CON EL FORMATO ORIGINAL
# --------------------------------------------------------------------------
def _is_grey(cell):
    f = cell.fill
    if not f or f.fill_type != "solid" or f.fgColor is None or f.fgColor.type != "rgb":
        return False
    rgb = f.fgColor.rgb
    if not isinstance(rgb, str) or len(rgb) < 6:
        return False
    r, g, b = int(rgb[-6:-4], 16), int(rgb[-4:-2], 16), int(rgb[-2:], 16)
    return abs(r - g) < 12 and abs(g - b) < 12 and 0x90 <= r <= 0xE0


def _has_color_fill(cell):
    f = cell.fill
    return bool(f and f.fill_type == "solid" and f.fgColor is not None and f.fgColor.type == "rgb"
                and isinstance(f.fgColor.rgb, str) and f.fgColor.rgb[-6:] not in ("FFFFFF", "000000"))


def _import_structure(store, district_id, path):
    """
    Lee la columna A de un Excel con el formato original:
      fondo gris = zona; resto de filas con texto = centros; una fila sin negrita justo debajo de
      un centro en negrita (o 'UCCU') = unidad dentro de ese centro.
    Las filas de leyenda de trabajadores (con color) anteriores a la primera zona se ignoran.
    """
    wb = load_workbook(path)
    ws = next((s for s in wb.worksheets if any(_is_grey(s.cell(r, 1)) for r in range(1, s.max_row + 1))),
              wb.worksheets[0])
    stats = {"zones": 0, "centers": 0, "units": 0, "skipped": 0, "sheet": ws.title}
    zone_id, new_zones = None, []
    prev_kind, last_top, last_top_bold = None, None, False
    seen_zone = False
    for r in range(3, ws.max_row + 1):
        cell = ws.cell(r, 1)
        name = clean(str(cell.value)) if cell.value not in (None, "") else ""
        if not name:
            prev_kind = None
            continue
        if _is_grey(cell):
            seen_zone = True
            existing = store._one("SELECT id FROM zones WHERE district_id=? AND name=?", (district_id, name))
            if existing:
                zone_id = existing["id"]
            else:
                zone_id = store.add_zone(district_id, name)
                new_zones.append(zone_id)
                stats["zones"] += 1
            prev_kind, last_top = "zone", None
            continue
        if not seen_zone and _has_color_fill(cell):
            continue                                            # leyenda de trabajadores
        if zone_id is None:
            existing = store._one("SELECT id FROM zones WHERE district_id=? AND name=?", (district_id, "GENERAL"))
            zone_id = existing["id"] if existing else store.add_zone(district_id, "GENERAL")
            if not existing:
                new_zones.append(zone_id)
                stats["zones"] += 1
        bold = bool(cell.font and cell.font.b)
        is_unit = prev_kind in ("top", "unit") and last_top is not None and \
            (name.upper() == "UCCU" or (not bold and prev_kind == "top" and last_top_bold))
        try:
            if is_unit:
                store.add_center(district_id, zone_id, name, parent_id=last_top)
                stats["units"] += 1
                prev_kind = "unit"
            else:
                last_top = store.add_center(district_id, zone_id, name)
                last_top_bold = bold
                stats["centers"] += 1
                prev_kind = "top"
        except DataError:
            stats["skipped"] += 1
    # zona sin centros (p. ej. un edificio visitable): se crea un centro con el mismo nombre
    for zid in new_zones:
        if not store._one("SELECT 1 FROM centers WHERE zone_id=?", (zid,)):
            store.add_center(district_id, zid, store.zone(zid)["name"])
            stats["centers"] += 1
    return stats


def import_structure(store, district_id, path):
    with store.muted():
        st = _import_structure(store, district_id, path)
    store.log("Estructura", "Importación de estructura (Excel de registro)",
              f"Distrito: {store.district(district_id)['name']} | Archivo: {os.path.basename(path)} | "
              f"Zonas nuevas: {st['zones']} | Centros: {st['centers']} | Unidades: {st['units']} | Omitidos: {st['skipped']}")
    return st


def _norm_key(text):
    return re.sub(r"[^a-z]", "", strip_accents(text or "").lower())


def _import_hierarchy(store, path):
    """
    Importa desde un Excel con:  columna A = Distrito/Hospital, B = Zona, C = Centro
    (columna D opcional = centro al que pertenece, para unidades como UCCU).
    Las celdas vacías en A o B repiten el valor de la fila anterior, de modo que sirve tanto una
    tabla completa (todo repetido en cada fila) como un listado agrupado. La primera fila puede ser
    una cabecera. Lo que ya existe se respeta (no se duplica).
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    st = {"districts": 0, "zones": 0, "centers": 0, "units": 0, "existing": 0, "skipped": 0,
          "district_ids": set(), "notes": []}
    cur_d = cur_z = None
    dcache, zcache = {}, {}

    def get_district(name):
        key = name.lower()
        if key not in dcache:
            row = store._one("SELECT id FROM districts WHERE name=?", (name,))
            if row:
                dcache[key] = row["id"]
            else:
                dcache[key] = store.add_district(name)
                st["districts"] += 1
        st["district_ids"].add(dcache[key])
        return dcache[key]

    def get_zone(did, name):
        key = (did, name.lower())
        if key not in zcache:
            row = store._one("SELECT id FROM zones WHERE district_id=? AND name=?", (did, name))
            if row:
                zcache[key] = row["id"]
            else:
                zcache[key] = store.add_zone(did, name)
                st["zones"] += 1
        return zcache[key]

    for n, row in enumerate(ws.iter_rows(values_only=True), start=1):
        vals = [clean(str(v)) if v not in (None, "") else "" for v in (list(row) + [None] * 4)[:4]]
        a, b, c, d = vals
        if n == 1 and _norm_key(a) in ("distrito", "hospital", "distritohospital", "hospitaldistrito",
                                       "distritos", "hospitales", "distritoshospitales") \
                and (_norm_key(b) in ("zona", "zonas", "zonadeaccion", "zonasdeaccion")
                     or _norm_key(c) in ("centro", "centros")):
            continue
        if not (a or b or c):
            continue
        if a and a.lower() != (cur_d or "").lower():
            cur_d, cur_z = a, None
        if b:
            cur_z = b
        if not cur_d:
            st["skipped"] += 1
            st["notes"].append(f"Fila {n}: falta el Distrito/Hospital.")
            continue
        try:
            did = get_district(cur_d)
            if not cur_z:
                if c:
                    raise DataError("falta la zona")
                continue
            zid = get_zone(did, cur_z)
            if not c:
                continue
            parent = None
            if d:
                prow = store._one("SELECT id FROM centers WHERE zone_id=? AND parent_id IS NULL AND name=? "
                                  "COLLATE NOCASE", (zid, d))
                if not prow:
                    raise DataError(f"no existe el centro «{d}» en esa zona")
                parent = prow["id"]
            if store._one("SELECT 1 FROM centers WHERE zone_id=? AND COALESCE(parent_id,0)=? AND name=? "
                          "COLLATE NOCASE", (zid, parent or 0, c)):
                st["existing"] += 1
                continue
            store.add_center(did, zid, c, parent)
            st["units" if parent else "centers"] += 1
        except DataError as exc:
            st["skipped"] += 1
            st["notes"].append(f"Fila {n}: {exc}.")
    return st


def import_hierarchy(store, path):
    with store.muted():
        st = _import_hierarchy(store, path)
    store.log("Estructura", "Importación de estructura (Distrito | Zona | Centro)",
              f"Archivo: {os.path.basename(path)} | Distritos nuevos: {st['districts']} | Zonas nuevas: {st['zones']} | "
              f"Centros nuevos: {st['centers']} | Unidades nuevas: {st['units']} | Ya existían: {st['existing']} | "
              f"Omitidas: {st['skipped']}")
    return st


def create_import_template(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Estructura"
    heads = ["Distrito/Hospital", "Zona", "Centro", "Unidad de (opcional)"]
    for i, (h, w) in enumerate(zip(heads, (38, 34, 46, 30)), start=1):
        c = ws.cell(1, i, h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        c.fill = _fill(COLOR_HEAD)
        ws.column_dimensions[get_column_letter(i)].width = w
    example = [("DISTRITO EJEMPLO", "Z.B. NORTE", "CENTRO DE SALUD NORTE", ""),
               ("", "", "UCCU", "CENTRO DE SALUD NORTE"),
               ("", "", "CONSULTORIO LOS OLIVOS", ""),
               ("", "Z.B. SUR", "CENTRO DE SALUD SUR", ""),
               ("HOSPITAL EJEMPLO", "PLANTA 1", "CONSULTA 1", "")]
    for r, row in enumerate(example, start=2):
        for i, v in enumerate(row, start=1):
            ws.cell(r, i, v or None).font = Font(name="Arial", size=10, color="FF7F7F7F")
    ws.freeze_panes = "A2"
    info = wb.create_sheet("Instrucciones")
    info.column_dimensions["A"].width = 110
    lines = ["Escriba en la hoja «Estructura» (la primera hoja del libro):",
             "  Columna A: Distrito/Hospital      Columna B: Zona      Columna C: Centro",
             "  Columna D (opcional): centro al que pertenece una unidad, p. ej. UCCU dentro de un centro de salud.",
             "Las celdas vacías de A y B repiten el valor de la fila anterior.",
             "Borre las filas de ejemplo (en gris) antes de importar. Lo que ya exista no se duplica."]
    for i, text in enumerate(lines, start=1):
        info.cell(i, 1, text).font = Font(name="Arial", size=10)
    wb.save(path)


# --------------------------------------------------------------------------
#  ESTADÍSTICAS
# --------------------------------------------------------------------------
def stats_by_center(visits, centers=None):
    """-> lista de (zona, centro, visitas, % ) ordenada de más a menos visitas."""
    cnt, info = Counter(), {}
    for v in visits:
        cnt[v["center_id"]] += 1
        info[v["center_id"]] = (v["zone_name"], center_label(v["center_name"], v["parent_name"]))
    for c in centers or []:
        info.setdefault(c["id"], (c["zone_name"], center_label(c["name"], c["parent_name"])))
    total = sum(cnt.values()) or 1
    rows = [(info[k][0], info[k][1], cnt.get(k, 0), 100.0 * cnt.get(k, 0) / total) for k in info]
    rows.sort(key=lambda r: (-r[2], r[0].lower(), r[1].lower()))
    return rows


def stats_by_worker(visits, workers=None):
    """-> lista de (trabajador, iniciales, visitas, centros distintos, %)."""
    cnt, centers, info = Counter(), defaultdict(set), {}
    for v in visits:
        cnt[v["user_id"]] += 1
        centers[v["user_id"]].add(v["center_id"])
        info[v["user_id"]] = (v["full_name"], v["initials"])
    for w in workers or []:
        info.setdefault(w["id"], (w["full_name"], w["initials"]))
    total = sum(cnt.values()) or 1
    rows = [(info[k][0], info[k][1], cnt.get(k, 0), len(centers.get(k, ())), 100.0 * cnt.get(k, 0) / total)
            for k in info]
    rows.sort(key=lambda r: (-r[2], r[0].lower()))
    return rows


def stats_by_date(visits, group="Día"):
    """-> lista de (clave, etiqueta, etiqueta corta, visitas, centros distintos, trabajadores distintos)."""
    cnt, cs, ws_, labels = Counter(), defaultdict(set), defaultdict(set), {}
    for v in visits:
        d = dt.date.fromisoformat(v["visit_date"])
        if group == "Semana":
            iso = d.isocalendar()
            key = (iso[0], iso[1])
            labels[key] = (f"Semana {iso[1]:02d} de {iso[0]}", f"S{iso[1]:02d}")
        elif group == "Mes":
            key = (d.year, d.month)
            labels[key] = (f"{MESES[d.month - 1].title()} {d.year}", MESES[d.month - 1][:3].title())
        else:
            key = (d.year, d.month, d.day)
            labels[key] = (f"{DIAS[d.weekday()]} {d.strftime('%d/%m/%Y')}", d.strftime("%d/%m"))
        cnt[key] += 1
        cs[key].add(v["center_id"])
        ws_[key].add(v["user_id"])
    return [(k, labels[k][0], labels[k][1], cnt[k], len(cs[k]), len(ws_[k])) for k in sorted(cnt)]


def export_audit(path, title, filters_text, entries, sessions, numbers=None):
    """Exporta el registro de actividad: hoja con todos los eventos y hoja resumen de accesos."""
    numbers = numbers or {}
    wb = Workbook()
    ws = wb.active
    ws.title = "Registro de actividad"
    head_font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")

    def header(sheet, heads, widths):
        sheet.cell(1, 1, title).font = Font(name="Arial", size=13, bold=True)
        sheet.cell(2, 1, filters_text).font = Font(name="Arial", size=10, italic=True, color="FF595959")
        for i, (h, w) in enumerate(zip(heads, widths), start=1):
            c = sheet.cell(4, i, h)
            c.font, c.fill = head_font, _fill(COLOR_HEAD)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            sheet.column_dimensions[get_column_letter(i)].width = w

    header(ws, ["Fecha", "Hora", "Acceso nº", "Usuario", "Nombre", "Categoría", "Acción", "Detalle"],
           [12, 10, 10, 22, 30, 15, 38, 110])
    for r, e in enumerate(entries, start=5):
        t = dt.datetime.fromisoformat(e["ts"])
        vals = [t.date(), t.time(), numbers.get(e["session"]) if e["session"] else None, e["username"], e["full_name"],
                e["category"], e["action"], e["details"]]
        for i, val in enumerate(vals, start=1):
            c = ws.cell(r, i, val)
            c.font = Font(name="Arial", size=10, bold=(e["action"] == "Inicio de sesión"))
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(i == 8), horizontal="center" if i <= 3 else None)
            if i == 1:
                c.number_format = "DD/MM/YYYY"
            elif i == 2:
                c.number_format = "HH:MM:SS"
    last = 4 + len(entries)
    ws.cell(last + 2, 1, "TOTAL REGISTROS").font = Font(name="Arial", size=10, bold=True)
    c = ws.cell(last + 2, 3, f"=COUNTA(A5:A{max(5, last)})" if entries else 0)
    c.font = Font(name="Arial", size=10, bold=True)
    c.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A4:H{max(5, last)}"
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    w2 = wb.create_sheet("Accesos")
    header(w2, ["Acceso nº", "Usuario", "Nombre", "Inicio de sesión", "Cierre de sesión", "Duración", "Nº de cambios"],
           [11, 22, 30, 20, 20, 12, 15])
    for r, s_ in enumerate(sessions, start=5):
        start = dt.datetime.fromisoformat(s_["start_ts"])
        end = dt.datetime.fromisoformat(s_["end_ts"]) if s_["end_ts"] else None
        vals = [numbers.get(s_["session"], s_["session"]), s_["username"], s_["full_name"], start, end,
                f'=IF(E{r}="","",E{r}-D{r})', s_["changes"]]
        for i, val in enumerate(vals, start=1):
            c = w2.cell(r, i, val)
            c.font = Font(name="Arial", size=10)
            c.border = BORDER
            c.alignment = Alignment(horizontal="center" if i in (1, 4, 5, 6, 7) else None)
            if i in (4, 5):
                c.number_format = "DD/MM/YYYY HH:MM:SS"
            elif i == 6:
                c.number_format = "[h]:mm:ss"
    last2 = 4 + len(sessions)
    w2.cell(last2 + 2, 1, "TOTAL ACCESOS").font = Font(name="Arial", size=10, bold=True)
    c = w2.cell(last2 + 2, 3, f"=COUNTA(A5:A{max(5, last2)})" if sessions else 0)
    c.font = Font(name="Arial", size=10, bold=True)
    c.alignment = Alignment(horizontal="center")
    c = w2.cell(last2 + 2, 7, f"=SUM(G5:G{max(5, last2)})" if sessions else 0)
    c.font = Font(name="Arial", size=10, bold=True)
    c.alignment = Alignment(horizontal="center")
    w2.auto_filter.ref = f"A4:G{max(5, last2)}"
    w2.freeze_panes = "A5"
    wb.save(path)


def export_users(path, title, filters_text, users, visit_counts):
    """Exporta usuarios (filas de Store.users) a un Excel. Nunca incluye contraseñas."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Usuarios"
    ws.cell(1, 1, title).font = Font(name="Arial", size=13, bold=True)
    ws.cell(2, 1, filters_text).font = Font(name="Arial", size=10, italic=True, color="FF595959")
    heads = ["Usuario", "Nombre", "Apellidos", "DNI / NIE", "Perfil", "Distritos / Hospitales", "Iniciales",
             "Color", "Estado", "Visitas registradas"]
    widths = [22, 22, 30, 14, 15, 52, 10, 11, 10, 18]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(4, i, h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        c.fill = _fill(COLOR_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, u in enumerate(users, start=5):
        admin = u["role"] == "admin"
        vals = [u["username"], u["first_name"] or u["full_name"], u["last_name"] or "", u["dni"] or "",
                "Administrador" if admin else "Trabajador",
                "Todos (administrador)" if admin else (u["district_names"] or ""),
                "" if admin else (u["initials"] or ""), "" if admin else ("#" + u["color"] if u["color"] else ""),
                "Activo" if u["active"] else "Baja", visit_counts.get(u["id"], 0)]
        for i, val in enumerate(vals, start=1):
            c = ws.cell(r, i, val)
            c.font = Font(name="Arial", size=10)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(i == 6), horizontal="center" if i >= 7 else None)
        if not admin and u["color"]:
            ws.cell(r, 8).fill = _fill(u["color"])
            ws.cell(r, 8).font = Font(name="Arial", size=10, color="FF" + text_color_for(u["color"]))
    last = 4 + len(users)
    ws.cell(last + 2, 1, "TOTAL USUARIOS").font = Font(name="Arial", size=10, bold=True)
    c = ws.cell(last + 2, 3, f"=COUNTA(A5:A{max(5, last)})" if users else 0)
    c.font = Font(name="Arial", size=10, bold=True)
    c.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A4:J{max(5, last)}"
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


def export_visits(path, title, filters_text, visits):
    """Exporta una lista de visitas (filas de Store.visits) a un Excel."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Visitas"
    ws.cell(1, 1, title).font = Font(name="Arial", size=13, bold=True)
    ws.cell(2, 1, filters_text).font = Font(name="Arial", size=10, italic=True, color="FF595959")
    heads = ["Fecha", "Trabajador", "Iniciales", "Distrito / Hospital", "Zona", "Centro", "Observaciones"]
    widths = [12, 30, 10, 34, 34, 46, 70]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(4, i, h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        c.fill = _fill(COLOR_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    rows = sorted(visits, key=lambda v: (v["visit_date"], v["full_name"].lower(), v["id"]))
    for r, v in enumerate(rows, start=5):
        vals = [dt.date.fromisoformat(v["visit_date"]), v["full_name"], v["initials"], v["district_name"],
                v["zone_name"], center_label(v["center_name"], v["parent_name"]), v["observations"]]
        for i, val in enumerate(vals, start=1):
            c = ws.cell(r, i, val)
            c.font = Font(name="Arial", size=10)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(i == 7),
                                    horizontal="center" if i in (1, 3) else None)
            if i == 1:
                c.number_format = "DD/MM/YYYY"
    last = 4 + len(rows)
    ws.cell(last + 2, 1, "TOTAL VISITAS").font = Font(name="Arial", size=10, bold=True)
    c = ws.cell(last + 2, 3, f"=COUNTA(A5:A{max(5, last)})" if rows else 0)
    c.font = Font(name="Arial", size=10, bold=True)
    c.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A4:G{max(5, last)}"
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


def export_stats(path, title, filters_text, by_center, by_worker, by_date):
    wb = Workbook()
    wb.remove(wb.active)
    head_fill, head_font = _fill(COLOR_HEAD), Font(name="Arial", size=10, bold=True, color="FFFFFFFF")

    def sheet(name, heads, widths, rows, pct_col=None):
        ws = wb.create_sheet(name)
        ws.cell(1, 1, title).font = Font(name="Arial", size=13, bold=True)
        ws.cell(2, 1, filters_text).font = Font(name="Arial", size=10, italic=True, color="FF595959")
        for i, (h, w) in enumerate(zip(heads, widths), start=1):
            c = ws.cell(4, i, h)
            c.font, c.fill = head_font, head_fill
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(i)].width = w
        for r, row in enumerate(rows, start=5):
            for i, val in enumerate(row, start=1):
                c = ws.cell(r, i, val)
                c.font = Font(name="Arial", size=10)
                c.border = BORDER
                if pct_col and i == pct_col:
                    c.value = val / 100.0
                    c.number_format = "0.0%"
        if rows:
            tot = len(rows) + 5
            ws.cell(tot, 1, "TOTAL").font = Font(name="Arial", size=10, bold=True)
            visit_col = {"Por centro": 3, "Por trabajador": 3, "Por fecha": 3}[name]
            col = get_column_letter(visit_col)
            c = ws.cell(tot, visit_col, f"=SUM({col}5:{col}{tot - 1})")
            c.font = Font(name="Arial", size=10, bold=True)
        ws.freeze_panes = "A5"

    sheet("Por centro", ["Zona", "Centro", "Visitas", "% del total"], [36, 52, 10, 12],
          [(z, c, n, p) for z, c, n, p in by_center], pct_col=4)
    sheet("Por trabajador", ["Trabajador", "Iniciales", "Visitas", "Centros distintos", "% del total"],
          [34, 10, 10, 18, 12], [(a, b, n, c, p) for a, b, n, c, p in by_worker], pct_col=5)
    sheet("Por fecha", ["Periodo", "Clave", "Visitas", "Centros distintos", "Trabajadores"],
          [30, 12, 10, 18, 14], [(lab, short, n, c, w) for _, lab, short, n, c, w in by_date])
    wb.save(path)


# ==========================================================================
#  INTERFAZ GRÁFICA
# ==========================================================================
def _font_family():
    if sys.platform.startswith("win"):
        return "Segoe UI"
    if sys.platform == "darwin":
        return "Helvetica Neue"
    return "DejaVu Sans"


FAMILY = _font_family()
F_BASE = (FAMILY, 10)
F_BOLD = (FAMILY, 10, "bold")
F_SMALL = (FAMILY, 9)
F_SMALL_B = (FAMILY, 8, "bold")
F_H1 = (FAMILY, 20, "bold")
F_H2 = (FAMILY, 12, "bold")
F_KPI = (FAMILY, 22, "bold")
F_LOGIN = (FAMILY, 12)

C_GREEN = "#0F4A2B"          # verde del botón del logo
C_GREEN2 = "#1B7345"         # verde del aro de la «C» (hover / elemento activo)
C_GREEN_DARK = "#08301C"
C_GOLD = "#B8975A"
C_GOLD_LIGHT = "#E6C679"     # dorado claro para textos sobre fondo verde
C_BG = "#F1F5F2"
C_TEXT = "#1B2B22"
C_MUTED = "#5F7268"
C_LINE = "#D5DFD8"
C_OK = "#2E7D32"
C_ERR = "#B3261E"
C_ROW_ALT = "#F5F9F6"
C_SEL = "#D6EADD"
C_SIDE_TXT = "#BFD6C8"

ALL_ZONES = "Todas las zonas"
ALL = "Todos"
ALL_DISTRICTS = "Todos los distritos"
NO_PARENT = "(ninguno: centro independiente)"
ROLE_WORKER = "Trabajador"
ROLE_ADMIN = "Administrador"
ALL_USERS = "Todos los usuarios"
ADMINS_ONLY = "Administradores"


def setup_styles(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    root.option_add("*TCombobox*Listbox.font", F_BASE)
    root.option_add("*TCombobox*Listbox.selectBackground", C_GREEN)
    style.configure(".", font=F_BASE, background=C_BG, foreground=C_TEXT)
    style.configure("TFrame", background=C_BG)
    style.configure("TLabel", background=C_BG, foreground=C_TEXT)
    style.configure("TEntry", fieldbackground="white", padding=6, bordercolor=C_LINE,
                    lightcolor=C_LINE, darkcolor=C_LINE)
    style.map("TEntry", bordercolor=[("focus", C_GOLD)], lightcolor=[("focus", C_GOLD)],
              darkcolor=[("focus", C_GOLD)])
    style.configure("TCombobox", padding=5, bordercolor=C_LINE, lightcolor=C_LINE, darkcolor=C_LINE,
                    arrowcolor=C_GREEN)
    style.map("TCombobox", fieldbackground=[("readonly", "white")],
              selectbackground=[("readonly", "white")], selectforeground=[("readonly", C_TEXT)])
    for name, bg, fg, active in (("Primary", C_GREEN, "white", C_GREEN2),
                                 ("Secondary", "#E0EBE4", C_GREEN, "#CDE0D4"),
                                 ("Danger", "#F6E1DF", C_ERR, "#EFCBC7"),
                                 ("Gold", C_GOLD, "white", "#A38548")):
        style.configure(f"{name}.TButton", background=bg, foreground=fg, padding=(16, 8),
                        borderwidth=0, focusthickness=0, font=F_BOLD)
        style.map(f"{name}.TButton", background=[("active", active), ("disabled", "#E8EEEA")],
                  foreground=[("disabled", "#9DB1A5")])
    style.configure("Login.TEntry", padding=11)
    style.configure("Login.TButton", background=C_GREEN, foreground="white", padding=(20, 14), borderwidth=0,
                    focusthickness=0, font=(FAMILY, 12, "bold"))
    style.map("Login.TButton", background=[("active", C_GREEN2)])
    for name, bg, fg, active in (("PrimarySm", C_GREEN, "white", C_GREEN2), ("SecondarySm", "#E0EBE4", C_GREEN, "#CDE0D4"),
                                 ("DangerSm", "#F6E1DF", C_ERR, "#EFCBC7")):
        style.configure(f"{name}.TButton", background=bg, foreground=fg, padding=(10, 7), borderwidth=0,
                        focusthickness=0, font=F_BOLD)
        style.map(f"{name}.TButton", background=[("active", active), ("disabled", "#E8EEEA")],
                  foreground=[("disabled", "#9DB1A5")])
    style.configure("Treeview", rowheight=28, background="white", fieldbackground="white",
                    foreground=C_TEXT, borderwidth=0, font=F_BASE)
    style.configure("Treeview.Heading", background=C_GREEN, foreground="white", font=F_BOLD,
                    padding=(8, 7), relief="flat")
    style.map("Treeview.Heading", background=[("active", C_GREEN2)])
    style.map("Treeview", background=[("selected", C_SEL)], foreground=[("selected", C_GREEN)])
    style.configure("TNotebook", background=C_BG, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(20, 9), font=F_BOLD, background="#E0EBE4", foreground=C_MUTED)
    style.map("TNotebook.Tab", background=[("selected", "white")], foreground=[("selected", C_GREEN)])
    style.configure("TCheckbutton", background="white", font=F_BASE)
    style.configure("Vertical.TScrollbar", background="#E0EBE4", troughcolor="white", bordercolor="white",
                    arrowcolor=C_GREEN)


# ---- pequeños constructores de widgets -----------------------------------------------------
def card(parent):
    return tk.Frame(parent, bg="white", highlightbackground=C_LINE, highlightthickness=1)


def page_header(parent, title, subtitle=""):
    f = tk.Frame(parent, bg=C_BG)
    f.pack(fill="x", padx=28, pady=(22, 12))
    tk.Label(f, text=title, font=F_H1, bg=C_BG, fg=C_GREEN).pack(anchor="w")
    if subtitle:
        tk.Label(f, text=subtitle, font=F_BASE, bg=C_BG, fg=C_MUTED, justify="left").pack(anchor="w", pady=(2, 0))
    tk.Frame(f, bg=C_GOLD, height=2, width=64).pack(anchor="w", pady=(10, 0))
    return f


def form_field(parent, row, text, widget, col=0, colspan=1, padx=(0, 0)):
    tk.Label(parent, text=text.upper(), font=F_SMALL_B, fg=C_MUTED, bg="white").grid(
        row=row, column=col, columnspan=colspan, sticky="w", pady=(10, 3), padx=padx)
    widget.grid(row=row + 1, column=col, columnspan=colspan, sticky="ew", padx=padx)


def make_entry(parent, var=None, width=24, **kw):
    kw.setdefault("font", F_BASE)
    return ttk.Entry(parent, textvariable=var, width=width, **kw)


def make_combo(parent, var=None, values=(), width=24, readonly=True):
    cb = ttk.Combobox(parent, textvariable=var, values=list(values), width=width, font=F_BASE,
                      state="readonly" if readonly else "normal")
    return cb


def make_tree(parent, columns, height=12):
    """columns: lista de (id, cabecera, ancho, anclaje)."""
    wrap = tk.Frame(parent, bg="white", highlightbackground=C_LINE, highlightthickness=1)
    tree = ttk.Treeview(wrap, columns=[c[0] for c in columns], show="headings", height=height,
                        selectmode="browse")
    for cid, head, width, anchor in columns:
        tree.heading(cid, text=head)
        tree.column(cid, width=width, anchor=anchor, stretch=True)
    sb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    tree.tag_configure("odd", background=C_ROW_ALT)
    tree.tag_configure("admin", font=F_BOLD)
    return wrap, tree


def fill_tree(tree, rows):
    """rows: lista de (iid, valores)."""
    tree.delete(*tree.get_children())
    for i, row in enumerate(rows):
        iid, values = row[0], row[1]
        tags = (("odd",) if i % 2 else ()) + ((row[2],) if len(row) > 2 and row[2] else ())
        tree.insert("", "end", iid=str(iid), values=values, tags=tags)


class ScrollFrame(tk.Frame):
    """Contenedor con barra de desplazamiento vertical: añada los widgets a .inner"""

    def __init__(self, parent, bg="white"):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.bind_all("<MouseWheel>", self._on_wheel, add="+")
        self.bind_all("<Button-4>", self._on_wheel, add="+")
        self.bind_all("<Button-5>", self._on_wheel, add="+")

    def _on_wheel(self, e):
        try:
            if not self.winfo_exists():
                return
            w = self.winfo_containing(e.x_root, e.y_root)
            while w is not None and w is not self:
                w = w.master
            if w is None:
                return                                   # el ratón no está sobre este contenedor
            if self.inner.winfo_reqheight() <= self.canvas.winfo_height():
                return                                   # cabe entero: nada que desplazar
            step = -1 if (getattr(e, "num", 0) == 4 or getattr(e, "delta", 0) > 0) else 1
            self.canvas.yview_scroll(step, "units")
        except Exception:
            pass


def sel_id(tree):
    s = tree.selection()
    return int(s[0]) if s else None


def show_error(parent, exc):
    if isinstance(exc, DataWarning):
        messagebox.showwarning(APP_NAME, str(exc), parent=parent)
    else:
        messagebox.showerror(APP_NAME, str(exc), parent=parent)


def show_info(parent, msg):
    messagebox.showinfo(APP_NAME, msg, parent=parent)


def ask_yes(parent, msg):
    return messagebox.askyesno(APP_NAME, msg, parent=parent)


def fmt_pct(p):
    return f"{p:.1f} %".replace(".", ",")


def month_bounds(year, month):
    return dt.date(year, month, 1), dt.date(year, month, calendar.monthrange(year, month)[1])


# ---- selector de fecha --------------------------------------------------------------------------
class CalendarPopup(tk.Toplevel):
    def __init__(self, anchor, current, on_pick):
        super().__init__(anchor)
        self.title("Seleccionar fecha")
        self.configure(bg="white", padx=12, pady=12)
        self.resizable(False, False)
        self.transient(anchor.winfo_toplevel())
        self.on_pick = on_pick
        self.selected = current
        base = current or dt.date.today()
        self.year, self.month = base.year, base.month
        head = tk.Frame(self, bg="white")
        head.pack(fill="x")
        self._nav(head, "‹", -1).pack(side="left")
        self.lbl = tk.Label(head, bg="white", fg=C_GREEN, font=F_BOLD, width=18)
        self.lbl.pack(side="left", expand=True)
        self._nav(head, "›", 1).pack(side="right")
        self.grid_frame = tk.Frame(self, bg="white")
        self.grid_frame.pack(pady=(8, 0))
        self._draw()
        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height() + 4
        self.geometry(f"+{x}+{y}")
        self.after(60, self._grab)

    def _grab(self):
        try:
            self.grab_set()
            self.focus_set()
        except Exception:
            pass

    def _nav(self, parent, text, delta):
        lb = tk.Label(parent, text=text, bg="white", fg=C_GREEN, font=(FAMILY, 14, "bold"), cursor="hand2", padx=8)
        lb.bind("<Button-1>", lambda e: self._move(delta))
        return lb

    def _move(self, delta):
        self.month += delta
        if self.month < 1:
            self.month, self.year = 12, self.year - 1
        elif self.month > 12:
            self.month, self.year = 1, self.year + 1
        self._draw()

    def _draw(self):
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self.lbl.config(text=f"{MESES[self.month - 1].title()} {self.year}")
        for c, name in enumerate(DIAS):
            tk.Label(self.grid_frame, text=name, bg="white", width=4, font=F_SMALL_B,
                     fg=C_ERR if c >= 5 else C_MUTED).grid(row=0, column=c, pady=(0, 4))
        today = dt.date.today()
        for r, week in enumerate(calendar.Calendar(0).monthdayscalendar(self.year, self.month), start=1):
            for c, day in enumerate(week):
                if day == 0:
                    continue
                d = dt.date(self.year, self.month, day)
                is_sel = self.selected == d
                bg = C_GREEN if is_sel else ("#FBE9E7" if c >= 5 else "white")
                fg = "white" if is_sel else (C_ERR if c >= 5 else C_TEXT)
                lb = tk.Label(self.grid_frame, text=str(day), width=4, bg=bg, fg=fg, cursor="hand2",
                              font=F_BOLD if d == today else F_BASE, pady=3,
                              highlightthickness=1 if d == today else 0, highlightbackground=C_GOLD)
                lb.grid(row=r, column=c, padx=1, pady=1)
                lb.bind("<Button-1>", lambda e, dd=d: self._pick(dd))

    def _pick(self, d):
        self.on_pick(d)
        self.destroy()


class DateEntry(tk.Frame):
    """Campo de fecha dd/mm/aaaa con calendario desplegable."""

    def __init__(self, parent, value=None, bg="white"):
        super().__init__(parent, bg=bg)
        self.var = tk.StringVar(value=fmt_date(value) if value else "")
        self.entry = ttk.Entry(self, textvariable=self.var, width=12, font=F_BASE)
        self.entry.pack(side="left")
        ttk.Button(self, text="▾", width=3, style="Secondary.TButton", command=self.open_calendar).pack(
            side="left", padx=(4, 0))

    def open_calendar(self):
        CalendarPopup(self.entry, parse_date(self.var.get()), self.set_date)

    def set_date(self, d):
        self.var.set(fmt_date(d) if d else "")

    def get_date(self, blank_ok=False):
        text = self.var.get().strip()
        if not text and blank_ok:
            return None
        d = parse_date(text)
        if d is None:
            raise DataError("La fecha debe tener el formato dd/mm/aaaa (por ejemplo 15/09/2026).")
        return d


class PasswordDialog(tk.Toplevel):
    """Pide una contraseña nueva dos veces. Resultado en .result (None si se cancela)."""

    def __init__(self, parent, title, prompt):
        super().__init__(parent)
        self.title(title)
        self.configure(bg="white", padx=26, pady=22)
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.result = None
        tk.Label(self, text=prompt, bg="white", fg=C_GREEN, font=F_H2, justify="left").grid(
            row=0, column=0, sticky="w")
        self.v1, self.v2 = tk.StringVar(), tk.StringVar()
        self.e1 = make_entry(self, self.v1, 30, show="•")
        self.e2 = make_entry(self, self.v2, 30, show="•")
        form_field(self, 1, "Nueva contraseña", self.e1)
        form_field(self, 3, "Repetir contraseña", self.e2)
        tk.Label(self, text=f"Mínimo {MIN_PASS_LEN} caracteres.", bg="white", fg=C_MUTED, font=F_SMALL).grid(
            row=5, column=0, sticky="w", pady=(6, 0))
        btns = tk.Frame(self, bg="white")
        btns.grid(row=6, column=0, sticky="e", pady=(16, 0))
        ttk.Button(btns, text="Cancelar", style="Secondary.TButton", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(btns, text="Guardar", style="Primary.TButton", command=self._ok).pack(side="left")
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.e1.focus_set()
        self.after(60, self._grab)
        self.wait_window(self)

    def _grab(self):
        try:
            self.grab_set()
        except Exception:
            pass

    def _ok(self):
        try:
            if self.v1.get() != self.v2.get():
                raise DataError("Las dos contraseñas no coinciden.")
            validate_password(self.v1.get())
        except DataError as exc:
            show_error(self, exc)
            return
        self.result = self.v1.get()
        self.destroy()


# ---- gráfico de barras ---------------------------------------------------------------------------
class BarChart(tk.Canvas):
    def __init__(self, parent, horizontal=True, **kw):
        super().__init__(parent, bg="white", highlightthickness=1, highlightbackground=C_LINE, **kw)
        self.horizontal = horizontal
        self.labels, self.values = [], []
        self.bind("<Configure>", lambda e: self.redraw())

    def set_data(self, labels, values):
        self.labels, self.values = list(labels), list(values)
        self.redraw()

    def redraw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 60 or h < 60:
            return
        if not self.values or max(self.values) <= 0:
            self.create_text(w / 2, h / 2, text="Sin datos en el periodo seleccionado", fill=C_MUTED, font=F_BASE)
            return
        if self.horizontal:
            self._draw_horizontal(w, h)
        else:
            self._draw_vertical(w, h)

    def _draw_horizontal(self, w, h):
        data = list(zip(self.labels, self.values))[:15]
        n, top = len(data), 14
        row_h = min(32.0, (h - 2 * top) / n)
        bar_h = row_h * 0.62
        label_w = int(w * 0.44)
        bar_max = max(20, w - label_w - 64)
        vmax = max(v for _, v in data)
        for i, (lab, val) in enumerate(data):
            y = top + i * row_h + row_h / 2
            text = lab if len(lab) <= 36 else lab[:35] + "…"
            self.create_text(label_w - 8, y, text=text, anchor="e", font=F_SMALL, fill=C_TEXT)
            x1 = label_w + max(2, bar_max * val / vmax)
            self.create_rectangle(label_w, y - bar_h / 2, x1, y + bar_h / 2, fill=C_GOLD if val == vmax else C_GREEN,
                                  outline="")
            self.create_text(x1 + 6, y, text=str(val), anchor="w", font=F_SMALL_B, fill=C_GREEN)

    def _draw_vertical(self, w, h):
        data = list(zip(self.labels, self.values))[-40:]
        n = len(data)
        pad_l, pad_r, top, bottom = 18, 14, 26, 34
        plot_w, plot_h = w - pad_l - pad_r, h - top - bottom
        slot = plot_w / n
        bar_w = max(3.0, slot * 0.7)
        vmax = max(v for _, v in data)
        base_y = top + plot_h
        self.create_line(pad_l, base_y, w - pad_r, base_y, fill=C_LINE)
        step = max(1, int(26 / slot) + (1 if 26 % slot else 0)) if slot < 26 else 1
        for i, (lab, val) in enumerate(data):
            x = pad_l + i * slot + slot / 2
            bh = plot_h * val / vmax
            self.create_rectangle(x - bar_w / 2, base_y - bh, x + bar_w / 2, base_y, fill=C_GOLD if val == vmax else C_GREEN,
                                  outline="")
            if slot >= 18:
                self.create_text(x, base_y - bh - 8, text=str(val), font=F_SMALL_B, fill=C_GREEN)
            if i % step == 0:
                self.create_text(x, base_y + 14, text=lab, font=F_SMALL, fill=C_MUTED)


# ---- ventana principal ---------------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1220x760")
        self.minsize(1040, 660)
        self.configure(bg=C_GREEN_DARK)
        self.store = Store()
        self.user = None
        self._logo = None
        setup_styles(self)
        self.container = tk.Frame(self, bg=C_BG)
        self.container.pack(fill="both", expand=True)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.show_login()

    def get_logo(self):
        """Imagen del logo (o None si no se pudiera cargar). Se conserva en la aplicación para que Tk no la libere."""
        if self._logo is None:
            try:
                self._logo = tk.PhotoImage(data=LOGO_PNG_B64)
            except Exception:
                self._logo = False
            try:
                self._icon = tk.PhotoImage(data=LOGO_ICON_B64)
                self.iconphoto(True, self._icon)
            except Exception:
                pass
        return self._logo or None

    def _clear(self):
        for w in self.container.winfo_children():
            w.destroy()

    def on_close(self):
        """Al cerrar la ventana se anota el cierre de la sesión que estuviera abierta."""
        try:
            if self.user:
                self.store.end_session("Cierre de la aplicación")
        finally:
            self.destroy()

    def show_login(self):
        if self.user:
            self.store.end_session("Cierre de sesión")
        self.user = None
        self._clear()
        LoginView(self.container, self).pack(fill="both", expand=True)

    def on_login(self, user):
        self.user = user
        self.store.start_session(user)
        self._clear()
        MainView(self.container, self).pack(fill="both", expand=True)
        if self.store.uses_default_password(user):
            self.after(400, self._offer_password_change)

    def _offer_password_change(self):
        if ask_yes(self, "Está utilizando la contraseña inicial.\n\n¿Desea cambiarla ahora?"):
            self.main.show("cuenta")

    def sync(self, district_id, years=None):
        """Actualiza el Excel del distrito; avisa si no es posible (p. ej. está abierto)."""
        ok, msg = sync_excel(self.store, district_id, years)
        if not ok:
            messagebox.showwarning(APP_NAME, "Los datos se han guardado, pero el Excel no se pudo actualizar:\n\n" + msg,
                                   parent=self)
        return ok

    def sync_many(self, district_ids, years=None):
        """Actualiza el Excel de varios distritos y avisa una sola vez si alguno falla."""
        problems = []
        for did in sorted(set(d for d in district_ids if d)):
            ok, msg = sync_excel(self.store, did, years)
            if not ok:
                problems.append(msg)
        if problems:
            messagebox.showwarning(APP_NAME, "Los datos se han guardado, pero algún Excel no se pudo actualizar:\n\n"
                                   + "\n\n".join(problems), parent=self)
        return not problems


class LoginView(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_GREEN_DARK)
        self.app = app
        box = tk.Frame(self, bg="white")
        box.place(relx=0.5, rely=0.5, anchor="center")

        # --- panel izquierdo: logo y nombre de la aplicación
        left = tk.Frame(box, bg=C_GREEN, width=440, height=610)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        logo = app.get_logo()
        if logo is not None:
            tk.Label(left, image=logo, bg=C_GREEN, bd=0).pack(pady=(24, 4))
        else:                                              # respaldo si no se pudiera cargar el logo
            cv = tk.Canvas(left, width=120, height=120, bg=C_GREEN, highlightthickness=0)
            cv.pack(pady=(80, 30))
            cv.create_oval(4, 4, 116, 116, outline=C_GOLD_LIGHT, width=2)
            cv.create_text(60, 60, text="RV", fill=C_GOLD_LIGHT, font=(FAMILY, 32, "bold"))
        tk.Label(left, text="Registro de Visitas", bg=C_GREEN, fg=C_GOLD_LIGHT, font=(FAMILY, 24, "bold")).pack()
        tk.Frame(left, bg=C_GOLD_LIGHT, height=2, width=64).pack(pady=14)
        tk.Label(left, text="Control de Visitas a Centros Sanitarios por Distrito", bg=C_GREEN, fg=C_GOLD_LIGHT,
                 font=(FAMILY, 12), justify="center", wraplength=340).pack()

        # --- panel derecho: formulario de acceso
        right = tk.Frame(box, bg="white")
        right.pack(side="left", fill="both", padx=66, pady=60)
        tk.Label(right, text="Iniciar sesión", bg="white", fg=C_GREEN, font=(FAMILY, 24, "bold")).pack(anchor="w")
        tk.Label(right, text="Introduzca sus credenciales para continuar", bg="white", fg=C_MUTED,
                 font=(FAMILY, 11)).pack(anchor="w", pady=(4, 28))
        self.v_user, self.v_pass = tk.StringVar(), tk.StringVar()
        tk.Label(right, text="USUARIO", font=(FAMILY, 9, "bold"), fg=C_MUTED, bg="white").pack(anchor="w", pady=(4, 5))
        e1 = make_entry(right, self.v_user, 34, font=F_LOGIN, style="Login.TEntry")
        e1.pack(fill="x")
        tk.Label(right, text="CONTRASEÑA", font=(FAMILY, 9, "bold"), fg=C_MUTED, bg="white").pack(anchor="w", pady=(20, 5))
        self.e2 = make_entry(right, self.v_pass, 34, show="•", font=F_LOGIN, style="Login.TEntry")
        self.e2.pack(fill="x")
        self.v_show = tk.BooleanVar(value=False)
        ttk.Checkbutton(right, text="Mostrar contraseña", variable=self.v_show, command=self._toggle).pack(
            anchor="w", pady=(12, 0))
        self.lbl_err = tk.Label(right, text="", bg="white", fg=C_ERR, font=(FAMILY, 10), wraplength=340, justify="left")
        self.lbl_err.pack(anchor="w", pady=(14, 6))
        ttk.Button(right, text="Entrar", style="Login.TButton", command=self.login).pack(fill="x", pady=(6, 0))
        tk.Label(right, text="Formato de usuario: Nombre.Primerapellido", bg="white", fg=C_MUTED,
                 font=(FAMILY, 10)).pack(anchor="w", pady=(22, 0))
        for e in (e1, self.e2):
            e.bind("<Return>", lambda ev: self.login())
        e1.focus_set()

    def _toggle(self):
        self.e2.config(show="" if self.v_show.get() else "•")

    def login(self):
        user, msg = self.app.store.authenticate(self.v_user.get(), self.v_pass.get())
        if not user:
            self.lbl_err.config(text=msg)
            self.v_pass.set("")
            return
        self.app.on_login(user)


class MainView(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        app.main = self
        self.user = app.user
        self.is_admin = self.user["role"] == "admin"
        self.page = None
        self.nav = {}
        self.current = None

        header = tk.Frame(self, bg=C_GREEN, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="REGISTRO DE VISITAS", bg=C_GREEN, fg=C_GOLD_LIGHT, font=(FAMILY, 13, "bold")).pack(
            side="left", padx=(24, 0))
        tk.Frame(header, bg=C_GOLD, width=2, height=24).pack(side="left", padx=14)
        tk.Label(header, text="Administrador" if self.is_admin else "Panel del trabajador", bg=C_GREEN,
                 fg=C_SIDE_TXT, font=F_BASE).pack(side="left")
        out = tk.Label(header, text="Cerrar sesión", bg=C_GREEN, fg=C_GOLD, font=F_BOLD, cursor="hand2")
        out.pack(side="right", padx=24)
        out.bind("<Button-1>", lambda e: app.show_login())
        who = self.user["full_name"]
        tk.Label(header, text=who, bg=C_GREEN, fg="white", font=F_BASE).pack(side="right", padx=8)
        tk.Frame(self, bg=C_GOLD, height=2).pack(fill="x")

        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True)
        self.side = tk.Frame(body, bg=C_GREEN_DARK, width=240)
        self.side.pack(side="left", fill="y")
        self.side.pack_propagate(False)
        self.content = tk.Frame(body, bg=C_BG)
        self.content.pack(side="left", fill="both", expand=True)

        items = [("registro", "Registrar visita"),
                 ("visitas", "Visitas registradas" if self.is_admin else "Mis visitas"),
                 ("vacaciones", "Vacaciones"),
                 ("stats", "Estadísticas")]
        if self.is_admin:
            items += [("estructura", "Distritos, zonas y centros"),
                      ("trabajadores", "Trabajadores"), ("excel", "Archivos Excel"),
                      ("log", "Registro de actividad")]
        items.append(("cuenta", "Mi contraseña"))
        tk.Label(self.side, text="MENÚ", bg=C_GREEN_DARK, fg="#7FA38D", font=F_SMALL_B, anchor="w").pack(
            fill="x", padx=24, pady=(22, 8))
        for key, text in items:
            self._nav_item(key, text)
        self.show("registro")

    def _nav_item(self, key, text):
        lb = tk.Label(self.side, text=text, anchor="w", bg=C_GREEN_DARK, fg=C_SIDE_TXT, font=F_BASE,
                      padx=24, pady=12, cursor="hand2")
        lb.pack(fill="x")
        lb.bind("<Button-1>", lambda e, k=key: self.show(k))
        lb.bind("<Enter>", lambda e, l=lb: l.config(bg=C_GREEN2))
        lb.bind("<Leave>", lambda e, l=lb, k=key: l.config(bg=C_GREEN2 if self.current == k else C_GREEN_DARK))
        self.nav[key] = lb

    def show(self, key):
        self.current = key
        for k, lb in self.nav.items():
            active = k == key
            lb.config(bg=C_GREEN2 if active else C_GREEN_DARK, fg=C_GOLD if active else C_SIDE_TXT,
                      font=F_BOLD if active else F_BASE)
        if self.page is not None:
            self.page.destroy()
        self.page = PAGES[key](self.content, self.app)
        self.page.pack(fill="both", expand=True)


# ---- página: registrar visita -----------------------------------------------------------------------------
class RegistroPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store, self.user = app, app.store, app.user
        self.is_admin = self.user["role"] == "admin"
        self.zone_map, self.center_map, self.center_labels = {}, {}, []
        self.worker_map, self.district_map = {}, {}
        self.district_id = None
        page_header(self, "Registrar visita",
                    "Anote la visita realizada a un centro. Se reflejará automáticamente en el Excel del distrito.")
        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=28, pady=(0, 24))
        body.columnconfigure(0, weight=4, uniform="r")
        body.columnconfigure(1, weight=5, uniform="r")
        body.rowconfigure(0, weight=1)

        form = card(body)
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        inner = tk.Frame(form, bg="white")
        inner.pack(fill="both", expand=True, padx=26, pady=16)
        inner.columnconfigure(0, weight=1)
        r = 0
        self.v_district, self.v_worker = tk.StringVar(), tk.StringVar()
        self.v_zone, self.v_center = tk.StringVar(), tk.StringVar()
        self.cb_district = make_combo(inner, self.v_district, [], 30)
        form_field(inner, r, "Distrito / Hospital", self.cb_district)
        self.cb_district.bind("<<ComboboxSelected>>", lambda e: self._on_district())
        r += 2
        if self.is_admin:
            self.cb_worker = make_combo(inner, self.v_worker, [], 30)
            form_field(inner, r, "Nombre (trabajador)", self.cb_worker)
            self.cb_worker.bind("<<ComboboxSelected>>", lambda e: self._on_worker())
        else:
            self.v_worker.set(self.user["full_name"])
            form_field(inner, r, "Nombre", make_entry(inner, self.v_worker, 30, state="readonly"))
        r += 2
        self.date = DateEntry(inner, dt.date.today())
        form_field(inner, r, "Fecha", self.date)
        r += 2
        self.cb_zone = make_combo(inner, self.v_zone, [], 30)
        form_field(inner, r, "Zona", self.cb_zone)
        self.cb_zone.bind("<<ComboboxSelected>>", lambda e: self._load_centers())
        r += 2
        self.cb_center = make_combo(inner, self.v_center, [], 30, readonly=False)
        form_field(inner, r, "Centro que visita (escriba para filtrar)", self.cb_center)
        self.cb_center.bind("<KeyRelease>", self._filter_centers)
        r += 2
        self.txt = tk.Text(inner, height=4, font=F_BASE, relief="solid", bd=1, wrap="word",
                           highlightthickness=1, highlightbackground=C_LINE, highlightcolor=C_GOLD)
        form_field(inner, r, "Observaciones", self.txt)
        r += 2
        btns = tk.Frame(inner, bg="white")
        btns.grid(row=r, column=0, sticky="ew", pady=(16, 0))
        ttk.Button(btns, text="Registrar visita", style="Primary.TButton", command=self.save).pack(side="left")
        ttk.Button(btns, text="Limpiar", style="Secondary.TButton", command=self._clear).pack(side="left", padx=8)
        self.lbl_status = tk.Label(inner, text="", bg="white", fg=C_OK, font=F_BOLD, wraplength=360, justify="left")
        self.lbl_status.grid(row=r + 1, column=0, sticky="w", pady=(10, 0))

        side = card(body)
        side.grid(row=0, column=1, sticky="nsew")
        tk.Label(side, text="Últimas visitas registradas", bg="white", fg=C_GREEN, font=F_H2).pack(
            anchor="w", padx=20, pady=(18, 8))
        wrap, self.tree = make_tree(side, [("fecha", "Fecha", 90, "center"), ("quien", "Trabajador", 80, "center"),
                                           ("centro", "Centro", 300, "w")], height=12)
        wrap.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        ds = list(self.store.districts()) if self.is_admin else list(self.store.user_districts(self.user["id"]))
        self.district_map = {d["name"]: d["id"] for d in ds}
        self.cb_district["values"] = [d["name"] for d in ds]
        if ds:
            self.v_district.set(ds[0]["name"])
            self._on_district()
        elif self.is_admin:
            self.lbl_status.config(text="Cree primero un Distrito/Hospital, sus zonas, centros y trabajadores.",
                                   fg=C_ERR)
        else:
            self.lbl_status.config(text="No tiene ningún Distrito/Hospital asignado. Contacte con el administrador.",
                                   fg=C_ERR)

    # --- datos dependientes -----------------------------------------------------------------
    def _on_district(self):
        self.district_id = self.district_map.get(self.v_district.get())
        if self.is_admin:
            ws = self.store.workers(self.district_id, only_active=True)
            self.worker_map = {w["full_name"]: w for w in ws}
            self.cb_worker["values"] = list(self.worker_map)
            self.v_worker.set(ws[0]["full_name"] if ws else "")
        self._load_zones(None)
        self._on_worker()

    def _on_worker(self):
        self._reload_recent()

    def _load_zones(self, default_zone_id):
        zones = self.store.zones(self.district_id, alpha=True) if self.district_id else []
        self.zone_map = {z["name"]: z["id"] for z in zones}
        self.cb_zone["values"] = [ALL_ZONES] + list(self.zone_map)
        default = next((z["name"] for z in zones if z["id"] == default_zone_id), ALL_ZONES)
        self.v_zone.set(default)
        self._load_centers()

    def _load_centers(self):
        zid = self.zone_map.get(self.v_zone.get())
        centers = self.store.centers(self.district_id, zid, alpha=True) if self.district_id else []
        self.center_map = {}
        for c in centers:
            label = center_label(c["name"], c["parent_name"])
            if not zid:
                label = f"{c['zone_name']}  |  {label}"
            self.center_map[label] = c["id"]
        self.center_labels = list(self.center_map)
        self.cb_center["values"] = self.center_labels
        self.v_center.set("")

    def _filter_centers(self, event):
        if event.keysym in ("Up", "Down", "Left", "Right", "Return", "Tab", "Escape"):
            return
        txt = strip_accents(self.v_center.get()).lower()
        vals = [l for l in self.center_labels if txt in strip_accents(l).lower()]
        self.cb_center["values"] = vals if vals else self.center_labels

    def _current_worker(self):
        if not self.is_admin:
            return self.user
        w = self.worker_map.get(self.v_worker.get())
        if not w:
            raise DataError("Seleccione el trabajador que realiza la visita.")
        return w

    def _current_center_id(self):
        text = self.v_center.get().strip()
        if not text:
            raise DataError("Seleccione el centro visitado.")
        if text in self.center_map:
            return self.center_map[text]
        key = strip_accents(text).lower()
        exact = [l for l in self.center_labels if strip_accents(l).lower() == key]
        if len(exact) == 1:
            return self.center_map[exact[0]]
        matches = [l for l in self.center_labels if key in strip_accents(l).lower()]
        if len(matches) == 1:
            return self.center_map[matches[0]]
        raise DataError("Seleccione un centro de la lista desplegable.")

    # --- acciones --------------------------------------------------------------------------------
    def save(self):
        self.lbl_status.config(text="")
        try:
            worker = self._current_worker()
            d = self.date.get_date()
            cid = self._current_center_id()
            if d > dt.date.today() and not ask_yes(self, "La fecha es posterior a hoy.\n\n¿Registrar la visita de todos modos?"):
                return
            obs = self.txt.get("1.0", "end").strip()
            self.store.add_visit(worker["id"], cid, d.isoformat(), obs, self.user["id"])
        except DataError as exc:
            show_error(self, exc)
            return
        ok = self.app.sync(self.district_id, [d.year])
        self.txt.delete("1.0", "end")
        self.v_center.set("")
        self.cb_center["values"] = self.center_labels
        self.lbl_status.config(text="✔ Visita registrada" + (" y Excel actualizado." if ok else "."),
                               fg=C_OK if ok else C_ERR)
        self._reload_recent()

    def _clear(self):
        self.txt.delete("1.0", "end")
        self.v_center.set("")
        self.cb_center["values"] = self.center_labels
        self.date.set_date(dt.date.today())
        self.lbl_status.config(text="")

    def _reload_recent(self):
        try:
            uid = None
            if self.is_admin:
                w = self.worker_map.get(self.v_worker.get())
                uid = w["id"] if w else -1
            else:
                uid = self.user["id"]
            rows = [] if uid == -1 else self.store.visits(user_id=uid, limit=14)
        except Exception:
            rows = []
        fill_tree(self.tree, [(v["id"], (fmt_date(v["visit_date"]), v["initials"],
                                         center_label(v["center_name"], v["parent_name"]))) for v in rows])


# ---- diálogo: modificar visita ---------------------------------------------------------------------------------------------------------------
class EditVisitDialog(tk.Toplevel):
    """Permite cambiar fecha, centro y observaciones de una visita. Resultado en .result (None si se cancela)."""

    def __init__(self, parent, store, visit):
        super().__init__(parent)
        self.title("Modificar visita")
        self.configure(bg="white", padx=28, pady=22)
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.result = None
        self.center_map = {}
        tk.Label(self, text="Modificar visita", bg="white", fg=C_GREEN, font=F_H2).grid(row=0, column=0, sticky="w")
        tk.Label(self, text=f"Trabajador: {visit['full_name']}    |    Distrito/Hospital: {visit['district_name']}",
                 bg="white", fg=C_MUTED, font=F_SMALL).grid(row=1, column=0, sticky="w", pady=(2, 0))
        self.date = DateEntry(self, dt.date.fromisoformat(visit["visit_date"]))
        form_field(self, 2, "Fecha", self.date)
        current = ""
        for c in store.centers(visit["district_id"], alpha=True):
            label = f"{c['zone_name']}  |  {center_label(c['name'], c['parent_name'])}"
            self.center_map[label] = c["id"]
            if c["id"] == visit["center_id"]:
                current = label
        self.v_center = tk.StringVar(value=current)
        form_field(self, 4, "Centro", make_combo(self, self.v_center, list(self.center_map), 58))
        self.txt = tk.Text(self, height=4, width=58, font=F_BASE, relief="solid", bd=1, wrap="word",
                           highlightthickness=1, highlightbackground=C_LINE, highlightcolor=C_GOLD)
        self.txt.insert("1.0", visit["observations"] or "")
        form_field(self, 6, "Observaciones", self.txt)
        btns = tk.Frame(self, bg="white")
        btns.grid(row=8, column=0, sticky="e", pady=(18, 0))
        ttk.Button(btns, text="Cancelar", style="Secondary.TButton", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(btns, text="Guardar cambios", style="Primary.TButton", command=self._ok).pack(side="left")
        self.bind("<Escape>", lambda e: self.destroy())
        self.after(60, self._grab)
        self.wait_window(self)

    def _grab(self):
        try:
            self.grab_set()
        except Exception:
            pass

    def _ok(self):
        try:
            d = self.date.get_date()
            cid = self.center_map.get(self.v_center.get())
            if not cid:
                raise DataError("Seleccione el centro de la lista.")
        except DataError as exc:
            show_error(self, exc)
            return
        self.result = (cid, d, self.txt.get("1.0", "end").strip())
        self.destroy()


# ---- página: visitas registradas --------------------------------------------------------------------------
class VisitasPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store, self.user = app, app.store, app.user
        self.is_admin = self.user["role"] == "admin"
        self.rows = {}
        self.visits_list, self.period = [], (None, None)
        self.district_map, self.worker_map = {}, {}
        page_header(self, "Visitas registradas" if self.is_admin else "Mis visitas",
                    "Consulte las visitas por periodo. Deje las fechas en blanco para ver todas.")
        filt = card(self)
        filt.pack(fill="x", padx=28, pady=(0, 12))
        fi = tk.Frame(filt, bg="white")
        fi.pack(fill="x", padx=18, pady=(2, 14))
        first, last = month_bounds(dt.date.today().year, dt.date.today().month)
        self.d_from, self.d_to = DateEntry(fi, first), DateEntry(fi, last)
        form_field(fi, 0, "Desde", self.d_from, col=0, padx=(0, 14))
        form_field(fi, 0, "Hasta", self.d_to, col=1, padx=(0, 14))
        c = 2
        self.v_district, self.v_worker = tk.StringVar(value=ALL_DISTRICTS), tk.StringVar(value=ALL)
        if self.is_admin:
            ds = list(self.store.districts())
            self.district_map = {d["name"]: d["id"] for d in ds}
            self.cb_district = make_combo(fi, self.v_district, [ALL_DISTRICTS] + list(self.district_map), 28)
            self.cb_district.bind("<<ComboboxSelected>>", lambda e: self._load_workers())
            form_field(fi, 0, "Distrito / Hospital", self.cb_district, col=c, padx=(0, 14))
            self.cb_worker = make_combo(fi, self.v_worker, [ALL], 26)
            form_field(fi, 0, "Trabajador", self.cb_worker, col=c + 1, padx=(0, 14))
            c += 2
            self._load_workers()
        bt = tk.Frame(fi, bg="white")
        bt.grid(row=1, column=c, sticky="e")
        ttk.Button(bt, text="Filtrar", style="Primary.TButton", command=self.load).pack(side="left")

        cols = [("fecha", "Fecha", 90, "center")]
        if self.is_admin:
            cols += [("quien", "Trabajador", 170, "w"), ("distrito", "Distrito", 170, "w")]
        cols += [("zona", "Zona", 190, "w"), ("centro", "Centro", 250, "w"), ("obs", "Observaciones", 280, "w")]
        wrap, self.tree = make_tree(self, cols, height=12)
        wrap.pack(fill="both", expand=True, padx=28)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        bottom = tk.Frame(self, bg=C_BG)
        bottom.pack(fill="x", padx=28, pady=(10, 20))
        self.lbl_detail = tk.Label(bottom, text="", bg=C_BG, fg=C_TEXT, font=F_BASE, justify="left", wraplength=800)
        self.lbl_detail.pack(side="left", anchor="w")
        ttk.Button(bottom, text="Eliminar visita seleccionada", style="Danger.TButton", command=self.delete).pack(
            side="right")
        ttk.Button(bottom, text="Modificar visita seleccionada", style="Secondary.TButton", command=self.edit).pack(
            side="right", padx=(0, 10))
        ttk.Button(bottom, text="Exportar a Excel", style="Gold.TButton", command=self.export).pack(side="right", padx=(0, 10))
        self.lbl_count = tk.Label(bottom, text="", bg=C_BG, fg=C_MUTED, font=F_BASE)
        self.lbl_count.pack(side="right", padx=16)
        self.load()

    def _load_workers(self):
        did = self.district_map.get(self.v_district.get())
        ws = self.store.workers(did)
        self.worker_map = {w["full_name"]: w["id"] for w in ws}
        self.cb_worker["values"] = [ALL] + list(self.worker_map)
        self.v_worker.set(ALL)

    def load(self):
        try:
            d1, d2 = self.d_from.get_date(True), self.d_to.get_date(True)
            if d1 and d2 and d1 > d2:
                raise DataError("La fecha «Desde» no puede ser posterior a «Hasta».")
        except DataError as exc:
            show_error(self, exc)
            return
        if self.is_admin:
            did = self.district_map.get(self.v_district.get())
            uid = self.worker_map.get(self.v_worker.get())
        else:
            did, uid = None, self.user["id"]
        visits = self.store.visits(district_id=did, user_id=uid, d_from=d1.isoformat() if d1 else None,
                                   d_to=d2.isoformat() if d2 else None)
        self.rows = {v["id"]: v for v in visits}
        self.visits_list = [dict(v) for v in visits]
        self.period = (d1, d2)
        data = []
        for v in visits:
            obs = (v["observations"] or "").replace("\n", " ")
            vals = [fmt_date(v["visit_date"])]
            if self.is_admin:
                vals += [v["full_name"], v["district_name"]]
            vals += [v["zone_name"], center_label(v["center_name"], v["parent_name"]),
                     obs if len(obs) <= 60 else obs[:59] + "…"]
            data.append((v["id"], vals))
        fill_tree(self.tree, data)
        self.lbl_count.config(text=f"{len(visits)} visita(s)")
        self.lbl_detail.config(text="")

    def _on_select(self, _event=None):
        vid = sel_id(self.tree)
        v = self.rows.get(vid)
        if not v:
            return
        obs = v["observations"] or "(sin observaciones)"
        self.lbl_detail.config(text=f"Observaciones: {obs}")

    def edit(self):
        vid = sel_id(self.tree)
        v = self.rows.get(vid)
        if v is None:
            show_info(self, "Seleccione primero una visita de la lista.")
            return
        dlg = EditVisitDialog(self, self.store, v)
        if not dlg.result:
            return
        cid, d, obs = dlg.result
        try:
            pairs = self.store.update_visit(vid, cid, d.isoformat(), obs, self.user)
        except DataError as exc:
            show_error(self, exc)
            return
        if not pairs:
            show_info(self, "No se ha cambiado ningún dato de la visita.")
            return
        for did, year in set(pairs):
            self.app.sync(did, [year])
        self.load()

    def export(self):
        if not self.visits_list:
            show_info(self, "No hay visitas en la lista para exportar.\nAjuste los filtros y pulse «Filtrar».")
            return
        d1, d2 = self.period
        periodo = f"Periodo: {fmt_date(d1) if d1 else 'sin límite'} a {fmt_date(d2) if d2 else 'sin límite'}"
        if self.is_admin:
            title = "Visitas registradas"
            filt = f"{periodo}  |  Distrito: {self.v_district.get()}  |  Trabajador: {self.v_worker.get()}"
            scope = self.v_district.get() if self.v_district.get() != ALL_DISTRICTS else "TODOS"
        else:
            title = f"Mis visitas - {self.user['full_name']}"
            filt = periodo
            scope = self.user["username"]
        default = f"VISITAS_{slug(scope)}_{dt.date.today().strftime('%Y%m%d')}.xlsx"
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx", initialdir=BASE_DIR,
                                            initialfile=default, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            export_visits(path, title, filt, self.visits_list)
        except OSError as exc:
            show_error(self, DataError(f"No se pudo guardar el archivo (¿está abierto en Excel?):\n{exc}"))
            return
        self.store.log("Exportaciones", "Exportación de visitas",
                       f"{len(self.visits_list)} visita(s) → {os.path.basename(path)}")
        show_info(self, f"Se han exportado {len(self.visits_list)} visita(s) a:\n{path}")

    def delete(self):
        vid = sel_id(self.tree)
        if vid is None:
            show_info(self, "Seleccione primero una visita de la lista.")
            return
        if not ask_yes(self, "¿Eliminar la visita seleccionada?\nSe quitará también del Excel."):
            return
        try:
            district_id, year = self.store.delete_visit(vid, self.user)
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(district_id, [year])
        self.load()


# ---- página: estadísticas --------------------------------------------------------------------------------------
class StatsPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store, self.user = app, app.store, app.user
        self.is_admin = self.user["role"] == "admin"
        self.can_pick_worker = self.is_admin or WORKERS_SEE_ALL_STATS
        self.visits, self.cur = [], {"center": [], "worker": [], "date": []}
        self.district_map, self.worker_map = {}, {}
        sub = "Visitas por centro, por trabajador y por fecha." + \
              ("" if self.can_pick_worker else " Se muestran sus propias visitas.")
        page_header(self, "Estadísticas", sub)

        filt = card(self)
        filt.pack(fill="x", padx=28, pady=(0, 10))
        fi = tk.Frame(filt, bg="white")
        fi.pack(fill="x", padx=18, pady=(2, 14))
        today = dt.date.today()
        self.d_from, self.d_to = DateEntry(fi, dt.date(today.year, 1, 1)), DateEntry(fi, today)
        form_field(fi, 0, "Desde", self.d_from, col=0, padx=(0, 12))
        form_field(fi, 0, "Hasta", self.d_to, col=1, padx=(0, 12))
        self.all_label = ALL_DISTRICTS if self.is_admin else "Todos mis distritos"
        self.v_district, self.v_worker = tk.StringVar(value=self.all_label), tk.StringVar(value=ALL)
        c = 2
        ds = list(self.store.districts()) if self.is_admin else list(self.store.user_districts(self.user["id"]))
        self.district_map = {d["name"]: d["id"] for d in ds}
        cb = make_combo(fi, self.v_district, [self.all_label] + list(self.district_map), 26)
        if self.can_pick_worker:
            cb.bind("<<ComboboxSelected>>", lambda e: self._load_workers())
        form_field(fi, 0, "Distrito / Hospital", cb, col=c, padx=(0, 12))
        c += 1
        if self.can_pick_worker:
            self.cb_worker = make_combo(fi, self.v_worker, [ALL], 24)
            form_field(fi, 0, "Trabajador", self.cb_worker, col=c, padx=(0, 12))
            c += 1
            self._load_workers()
        tk.Label(fi, text="PERIODO RÁPIDO", font=F_SMALL_B, fg=C_MUTED, bg="white").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(12, 3))
        quick = tk.Frame(fi, bg="white")
        quick.grid(row=3, column=0, columnspan=c, sticky="w")
        for text, kind in (("Mes actual", "mes"), ("Mes anterior", "ant"), ("Año actual", "anio"), ("Todo", "todo")):
            ttk.Button(quick, text=text, style="Secondary.TButton", command=lambda k=kind: self._quick(k)).pack(
                side="left", padx=(0, 8))
        ttk.Button(quick, text="Aplicar", style="Primary.TButton", command=self.refresh).pack(side="left", padx=(14, 8))
        ttk.Button(quick, text="Exportar a Excel", style="Gold.TButton", command=self.export).pack(side="left")

        kp = tk.Frame(self, bg=C_BG)
        kp.pack(fill="x", padx=28, pady=(0, 10))
        self.kpi = {}
        for i, (key, title) in enumerate((("total", "VISITAS"), ("centros", "CENTROS VISITADOS"),
                                          ("trab", "TRABAJADORES ACTIVOS"), ("top", "CENTRO MÁS VISITADO"))):
            kp.columnconfigure(i, weight=1, uniform="k")
            f = card(kp)
            f.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0))
            tk.Label(f, text=title, bg="white", fg=C_MUTED, font=F_SMALL_B).pack(anchor="w", padx=16, pady=(12, 0))
            lb = tk.Label(f, text="0", bg="white", fg=C_GREEN, font=F_KPI if key != "top" else F_H2,
                          anchor="w", wraplength=230, justify="left")
            lb.pack(anchor="w", padx=16, pady=(0, 12), fill="x")
            self.kpi[key] = lb

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        # por centro
        t1 = tk.Frame(nb, bg="white")
        nb.add(t1, text="  Por centro  ")
        self.v_empty = tk.BooleanVar(value=False)
        ttk.Checkbutton(t1, text="Incluir centros sin visitas", variable=self.v_empty, command=self.refresh).pack(
            anchor="w", padx=12, pady=(8, 0))
        self.tree_c, self.chart_c = self._tab(t1, [("zona", "Zona", 150, "w"), ("centro", "Centro", 230, "w"),
                                                   ("n", "Visitas", 60, "center"), ("p", "%", 60, "center")], True)
        # por trabajador
        t2 = tk.Frame(nb, bg="white")
        nb.add(t2, text="  Por trabajador  ")
        self.tree_w, self.chart_w = self._tab(t2, [("quien", "Trabajador", 190, "w"), ("ini", "Inic.", 50, "center"),
                                                   ("n", "Visitas", 60, "center"), ("c", "Centros", 65, "center"),
                                                   ("p", "%", 60, "center")], True)
        # por fecha
        t3 = tk.Frame(nb, bg="white")
        nb.add(t3, text="  Por fecha  ")
        gl = tk.Frame(t3, bg="white")
        gl.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(gl, text="Agrupar por:", bg="white", fg=C_MUTED, font=F_BASE).pack(side="left")
        self.v_group = tk.StringVar(value="Día")
        cbg = make_combo(gl, self.v_group, ["Día", "Semana", "Mes"], 10)
        cbg.pack(side="left", padx=8)
        cbg.bind("<<ComboboxSelected>>", lambda e: self._fill_dates())
        self.tree_d, self.chart_d = self._tab(t3, [("per", "Periodo", 170, "w"), ("n", "Visitas", 60, "center"),
                                                   ("c", "Centros", 65, "center"), ("t", "Trabaj.", 65, "center")],
                                              False)
        self.refresh()

    def _tab(self, parent, cols, horizontal):
        f = tk.Frame(parent, bg="white")
        f.pack(fill="both", expand=True, padx=12, pady=10)
        f.columnconfigure(0, weight=5, uniform="s")
        f.columnconfigure(1, weight=6, uniform="s")
        f.rowconfigure(0, weight=1)
        wrap, tree = make_tree(f, cols, height=8)
        wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        chart = BarChart(f, horizontal=horizontal, height=200)
        chart.grid(row=0, column=1, sticky="nsew")
        return tree, chart

    def _load_workers(self):
        did = self._district_filter()
        ws = self.store.workers(did)
        self.worker_map = {w["full_name"]: w["id"] for w in ws}
        self.cb_worker["values"] = [ALL] + list(self.worker_map)
        self.v_worker.set(ALL)

    def _quick(self, kind):
        t = dt.date.today()
        if kind == "mes":
            a, b = month_bounds(t.year, t.month)
        elif kind == "ant":
            y, m = (t.year, t.month - 1) if t.month > 1 else (t.year - 1, 12)
            a, b = month_bounds(y, m)
        elif kind == "anio":
            a, b = dt.date(t.year, 1, 1), dt.date(t.year, 12, 31)
        else:
            a, b = None, None
        self.d_from.set_date(a)
        self.d_to.set_date(b)
        self.refresh()

    def _district_filter(self):
        """Id del distrito elegido; para un trabajador, 'todos' = la lista de sus distritos."""
        did = self.district_map.get(self.v_district.get())
        if did is None and not self.is_admin:
            return list(self.district_map.values()) or [-1]
        return did

    def _scope(self):
        did = self._district_filter()
        if self.can_pick_worker:
            uid = self.worker_map.get(self.v_worker.get())
        else:
            uid = self.user["id"]
        return did, uid

    def refresh(self):
        try:
            d1, d2 = self.d_from.get_date(True), self.d_to.get_date(True)
            if d1 and d2 and d1 > d2:
                raise DataError("La fecha «Desde» no puede ser posterior a «Hasta».")
        except DataError as exc:
            show_error(self, exc)
            return
        did, uid = self._scope()
        self.visits = [dict(v) for v in self.store.visits(
            district_id=did, user_id=uid, d_from=d1.isoformat() if d1 else None,
            d_to=d2.isoformat() if d2 else None)]
        centers = self.store.centers(did) if self.v_empty.get() else None
        if self.is_admin or self.can_pick_worker:
            workers = [w for w in self.store.workers(did, only_active=True) if not uid or w["id"] == uid]
        else:
            workers = [self.user]
        self.cur["center"] = stats_by_center(self.visits, centers)
        self.cur["worker"] = stats_by_worker(self.visits, workers)
        self.cur["period"] = (d1, d2)

        fill_tree(self.tree_c, [(i, (z, c, n, fmt_pct(p))) for i, (z, c, n, p) in enumerate(self.cur["center"])])
        shown = [r for r in self.cur["center"] if r[2] > 0]
        self.chart_c.set_data([r[1] for r in shown], [r[2] for r in shown])
        fill_tree(self.tree_w, [(i, (a, b, n, c, fmt_pct(p))) for i, (a, b, n, c, p) in enumerate(self.cur["worker"])])
        self.chart_w.set_data([r[0] for r in self.cur["worker"]], [r[2] for r in self.cur["worker"]])
        self._fill_dates()

        self.kpi["total"].config(text=str(len(self.visits)))
        self.kpi["centros"].config(text=str(len({v["center_id"] for v in self.visits})))
        self.kpi["trab"].config(text=str(len({v["user_id"] for v in self.visits})))
        top = next((r for r in self.cur["center"] if r[2] > 0), None)
        self.kpi["top"].config(text=f"{top[1]} ({top[2]})" if top else "—")

    def _fill_dates(self):
        rows = stats_by_date(self.visits, self.v_group.get())
        self.cur["date"] = rows
        fill_tree(self.tree_d, [(i, (lab, n, c, w)) for i, (_, lab, short, n, c, w) in enumerate(rows)])
        self.chart_d.set_data([r[2] for r in rows], [r[3] for r in rows])

    def export(self):
        d1, d2 = self.cur.get("period", (None, None))
        did, uid = self._scope()
        scope = self.v_district.get()
        title = f"Estadísticas de visitas - {scope}"
        filt = (f"Periodo: {fmt_date(d1) if d1 else 'sin límite'} a {fmt_date(d2) if d2 else 'sin límite'}"
                f"  |  Trabajador: {self.v_worker.get() if self.can_pick_worker else self.user['full_name']}")
        default = f"ESTADISTICAS_{slug(scope)}_{dt.date.today().strftime('%Y%m%d')}.xlsx"
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx", initialdir=BASE_DIR,
                                            initialfile=default, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            export_stats(path, title, filt, self.cur["center"], self.cur["worker"], self.cur["date"])
        except OSError as exc:
            show_error(self, DataError(f"No se pudo guardar el archivo:\n{exc}"))
            return
        self.store.log("Exportaciones", "Exportación de estadísticas", f"{title} → {os.path.basename(path)}")
        show_info(self, f"Estadísticas exportadas a:\n{path}")


# ---- página: mi contraseña -----------------------------------------------------------------------------------------
class CuentaPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store, self.user = app, app.store, app.user
        page_header(self, "Mi contraseña", "Cambie su contraseña de acceso.")
        row = tk.Frame(self, bg=C_BG)
        row.pack(fill="x", padx=28)
        c = card(row)
        c.pack(side="left", anchor="n")
        inner = tk.Frame(c, bg="white")
        inner.pack(padx=30, pady=22)
        u = self.store.user(self.user["id"])
        info = f"Usuario:  {u['username']}\nPerfil:  {ROLE_ADMIN if u['role'] == 'admin' else ROLE_WORKER}"
        if u["role"] == "worker":
            info += f"\nDistritos/Hospitales:  {u['district_names'] or '-'}"
        tk.Label(inner, text=info, bg="white", fg=C_MUTED, font=F_BASE, justify="left").grid(row=0, column=0, sticky="w")
        self.v = [tk.StringVar(), tk.StringVar(), tk.StringVar()]
        self.entries = [make_entry(inner, v, 34, show="•") for v in self.v]
        form_field(inner, 1, "Contraseña actual", self.entries[0])
        form_field(inner, 3, "Nueva contraseña", self.entries[1])
        form_field(inner, 5, "Repetir nueva contraseña", self.entries[2])
        self.v_show = tk.BooleanVar(value=False)
        ttk.Checkbutton(inner, text="Mostrar contraseñas", variable=self.v_show, command=self._toggle).grid(
            row=7, column=0, sticky="w", pady=(10, 0))
        tk.Label(inner, text=f"Mínimo {MIN_PASS_LEN} caracteres.", bg="white", fg=C_MUTED, font=F_SMALL).grid(
            row=8, column=0, sticky="w", pady=(4, 0))
        ttk.Button(inner, text="Cambiar contraseña", style="Primary.TButton", command=self.save).grid(
            row=9, column=0, sticky="w", pady=(16, 0))

    def _toggle(self):
        for e in self.entries:
            e.config(show="" if self.v_show.get() else "•")

    def save(self):
        try:
            self.store.change_own_password(self.user["id"], *[v.get() for v in self.v])
        except DataError as exc:
            show_error(self, exc)
            return
        for v in self.v:
            v.set("")
        show_info(self, "Contraseña cambiada correctamente.")


# ---- página: distritos, zonas y centros (unificada) -----------------------------------------------------------------------
class EstructuraPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store = app, app.store
        self.cur_d, self.cur_z = None, None
        self.parent_map = {}
        page_header(self, "Distritos, zonas y centros",
                    "Seleccione un distrito para ver sus zonas y una zona para ver sus centros. "
                    "Todo lo que cree aparecerá en los desplegables.")
        bar = tk.Frame(self, bg=C_BG)
        bar.pack(fill="x", padx=28, pady=(0, 10))
        ttk.Button(bar, text="Importar desde Excel…", style="Gold.TButton", command=self.import_hier).pack(side="left")
        ttk.Button(bar, text="Descargar plantilla", style="Secondary.TButton", command=self.template).pack(
            side="left", padx=8)
        ttk.Button(bar, text="Importar un Excel de registro…", style="Secondary.TButton", command=self.import_old).pack(
            side="left")
        tk.Label(bar, text="Excel: columna A = Distrito/Hospital, B = Zona, C = Centro", bg=C_BG, fg=C_MUTED,
                 font=F_SMALL).pack(side="left", padx=14)

        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=28, pady=(0, 22))
        for i, w in enumerate((10, 10, 14)):
            body.columnconfigure(i, weight=w, uniform="e")
        body.rowconfigure(0, weight=1)
        self.v_d, self.v_z, self.v_c, self.v_parent = (tk.StringVar() for _ in range(4))

        # --- columna 1: distritos
        c1 = card(body)
        c1.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        tk.Label(c1, text="Distritos / Hospitales", bg="white", fg=C_GREEN, font=F_H2).pack(anchor="w", padx=14, pady=(14, 8))
        wrap, self.tree_d = make_tree(c1, [("n", "Nombre", 150, "w"), ("z", "Zonas", 50, "center")], height=8)
        wrap.pack(fill="both", expand=True, padx=14)
        make_entry(c1, self.v_d, 20).pack(fill="x", padx=14, pady=(10, 6))
        b = tk.Frame(c1, bg="white")
        b.pack(fill="x", padx=14, pady=(0, 14))
        ttk.Button(b, text="Añadir", style="PrimarySm.TButton", command=self.add_d).pack(side="left")
        ttk.Button(b, text="Renombrar", style="SecondarySm.TButton", command=self.ren_d).pack(side="left", padx=4)
        ttk.Button(b, text="Eliminar", style="DangerSm.TButton", command=self.del_d).pack(side="left")

        # --- columna 2: zonas
        c2 = card(body)
        c2.grid(row=0, column=1, sticky="nsew", padx=5)
        self.lbl_z = tk.Label(c2, text="Zonas", bg="white", fg=C_GREEN, font=F_H2)
        self.lbl_z.pack(anchor="w", padx=14, pady=(14, 8))
        wrap, self.tree_z = make_tree(c2, [("n", "Zona", 150, "w"), ("c", "Centros", 55, "center")], height=8)
        wrap.pack(fill="both", expand=True, padx=14)
        make_entry(c2, self.v_z, 20).pack(fill="x", padx=14, pady=(10, 6))
        b = tk.Frame(c2, bg="white")
        b.pack(fill="x", padx=14, pady=(0, 14))
        ttk.Button(b, text="Añadir", style="PrimarySm.TButton", command=self.add_z).pack(side="left")
        ttk.Button(b, text="Renombrar", style="SecondarySm.TButton", command=self.ren_z).pack(side="left", padx=4)
        ttk.Button(b, text="Eliminar", style="DangerSm.TButton", command=self.del_z).pack(side="left")

        # --- columna 3: centros
        c3 = card(body)
        c3.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        self.lbl_c = tk.Label(c3, text="Centros", bg="white", fg=C_GREEN, font=F_H2)
        self.lbl_c.pack(anchor="w", padx=14, pady=(14, 8))
        wrap, self.tree_c = make_tree(c3, [("n", "Centro / unidad", 190, "w"), ("z", "Zona", 100, "w"),
                                           ("v", "Visitas", 55, "center")], height=8)
        wrap.pack(fill="both", expand=True, padx=14)
        make_entry(c3, self.v_c, 20).pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(c3, text="PERTENECE AL CENTRO (SOLO UNIDADES, P. EJ. UCCU)", font=F_SMALL_B, fg=C_MUTED,
                 bg="white").pack(anchor="w", padx=14, pady=(4, 2))
        self.cb_parent = make_combo(c3, self.v_parent, [NO_PARENT], 20)
        self.cb_parent.pack(fill="x", padx=14)
        b = tk.Frame(c3, bg="white")
        b.pack(fill="x", padx=14, pady=(8, 14))
        ttk.Button(b, text="Añadir", style="PrimarySm.TButton", command=self.add_c).pack(side="left")
        ttk.Button(b, text="Guardar cambios", style="SecondarySm.TButton", command=self.save_c).pack(side="left", padx=4)
        ttk.Button(b, text="Eliminar", style="DangerSm.TButton", command=self.del_c).pack(side="left")

        self.tree_d.bind("<<TreeviewSelect>>", self._on_d)
        self.tree_z.bind("<<TreeviewSelect>>", self._on_z)
        self.tree_c.bind("<<TreeviewSelect>>", self._on_c)
        self.refresh()

    # --- carga de datos ----------------------------------------------------------------------------------
    def refresh(self):
        self._fill_d()
        self._fill_z()
        self._fill_c()

    def _fill_d(self):
        rows = [(d["id"], (d["name"], len(self.store.zones(d["id"])))) for d in self.store.districts()]
        fill_tree(self.tree_d, rows)
        if self.cur_d and self.store.district(self.cur_d):
            self.tree_d.selection_set(str(self.cur_d))
        else:
            self.cur_d = self.cur_z = None

    def _fill_z(self):
        if not self.cur_d:
            fill_tree(self.tree_z, [])
            self.lbl_z.config(text="Zonas")
            return
        self.lbl_z.config(text="Zonas de: " + self.store.district(self.cur_d)["name"])
        counts = {r[0]: r[1] for r in self.store._q(
            "SELECT zone_id, COUNT(*) FROM centers WHERE district_id=? GROUP BY zone_id", (self.cur_d,))}
        fill_tree(self.tree_z, [(z["id"], (z["name"], counts.get(z["id"], 0)))
                                for z in self.store.zones(self.cur_d, alpha=True)])
        z = self.store.zone(self.cur_z) if self.cur_z else None
        if z and z["district_id"] == self.cur_d:
            self.tree_z.selection_set(str(self.cur_z))
        else:
            self.cur_z = None

    def _fill_c(self):
        if not self.cur_d:
            fill_tree(self.tree_c, [])
            self.lbl_c.config(text="Centros")
            self._load_parents(None)
            return
        if self.cur_z:
            self.lbl_c.config(text="Centros de: " + self.store.zone(self.cur_z)["name"])
        else:
            self.lbl_c.config(text="Centros (todas las zonas)")
        counts = self.store.visit_counts_by_center(self.cur_d)
        rows = []
        for c in self.store.centers(self.cur_d, self.cur_z, alpha=True):
            name = ("      └  " + c["name"]) if c["parent_id"] else c["name"]
            rows.append((c["id"], (name, c["zone_name"], counts.get(c["id"], 0))))
        fill_tree(self.tree_c, rows)
        self._load_parents(self.cur_z)

    def _load_parents(self, zone_id, exclude=None):
        tops = []
        if self.cur_d and zone_id:
            tops = [c for c in self.store.centers(self.cur_d, zone_id, alpha=True)
                    if c["parent_id"] is None and c["id"] != exclude]
        self.parent_map = {c["name"]: c["id"] for c in tops}
        self.cb_parent["values"] = [NO_PARENT] + list(self.parent_map)
        if self.v_parent.get() not in self.parent_map:
            self.v_parent.set(NO_PARENT)

    # --- selección ---------------------------------------------------------------------------------------------
    def _on_d(self, _e=None):
        did = sel_id(self.tree_d)
        if did is None or did == self.cur_d:
            return
        self.cur_d, self.cur_z = did, None
        self.v_d.set(self.store.district(did)["name"])
        self.v_z.set("")
        self.v_c.set("")
        self._fill_z()
        self._fill_c()

    def _on_z(self, _e=None):
        zid = sel_id(self.tree_z)
        if zid is None or zid == self.cur_z:
            return
        self.cur_z = zid
        self.v_z.set(self.store.zone(zid)["name"])
        self.v_c.set("")
        self._fill_c()

    def _on_c(self, _e=None):
        cid = sel_id(self.tree_c)
        c = self.store.center(cid) if cid else None
        if not c:
            return
        self.v_c.set(c["name"])
        self._load_parents(c["zone_id"], exclude=cid)
        self.v_parent.set(c["parent_name"] or NO_PARENT)

    def _need(self, what):
        if what == "d" and not self.cur_d:
            show_info(self, "Seleccione primero un Distrito/Hospital de la lista.")
        elif what == "z" and not self.cur_z:
            show_info(self, "Seleccione primero una zona de la lista.")
        else:
            return True
        return False

    # --- distritos ---------------------------------------------------------------------------------------------------
    def add_d(self):
        try:
            did = self.store.add_district(self.v_d.get())
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(did)
        self.cur_d, self.cur_z = did, None
        self.v_z.set("")
        self.refresh()

    def ren_d(self):
        if not self._need("d"):
            return
        try:
            self.store.rename_district(self.cur_d, self.v_d.get())
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(self.cur_d)
        self.refresh()

    def del_d(self):
        if not self._need("d"):
            return
        if not ask_yes(self, "¿Eliminar el Distrito/Hospital seleccionado?\n(Sus archivos Excel no se borran.)"):
            return
        try:
            self.store.delete_district(self.cur_d)
        except DataError as exc:
            show_error(self, exc)
            return
        self.cur_d = self.cur_z = None
        self.v_d.set("")
        self.v_z.set("")
        self.refresh()

    # --- zonas -----------------------------------------------------------------------------------------------------------
    def add_z(self):
        if not self._need("d"):
            return
        try:
            zid = self.store.add_zone(self.cur_d, self.v_z.get())
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(self.cur_d)
        self.cur_z = zid
        self.refresh()

    def ren_z(self):
        if not self._need("z"):
            return
        try:
            self.store.rename_zone(self.cur_z, self.v_z.get())
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(self.cur_d)
        self.refresh()

    def del_z(self):
        if not self._need("z"):
            return
        if not ask_yes(self, "¿Eliminar la zona seleccionada?"):
            return
        try:
            self.store.delete_zone(self.cur_z)
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(self.cur_d)
        self.cur_z = None
        self.v_z.set("")
        self.refresh()

    # --- centros -------------------------------------------------------------------------------------------------------------
    def _parent_id(self):
        return self.parent_map.get(self.v_parent.get())

    def add_c(self):
        if not self._need("d") or not self._need("z"):
            return
        try:
            self.store.add_center(self.cur_d, self.cur_z, self.v_c.get(), self._parent_id())
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(self.cur_d)
        self.v_c.set("")
        self.refresh()

    def save_c(self):
        cid = sel_id(self.tree_c)
        if not cid:
            show_info(self, "Seleccione en la lista el centro que desea modificar.")
            return
        c = self.store.center(cid)
        try:
            self.store.update_center(cid, self.v_c.get(), c["zone_id"], self._parent_id())
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync(self.cur_d)
        self.refresh()

    def del_c(self):
        cid = sel_id(self.tree_c)
        if not cid:
            show_info(self, "Seleccione en la lista el centro que desea eliminar.")
            return
        n = self.store.center_visit_count(cid)
        msg = "¿Eliminar el centro seleccionado (y sus unidades, si las tiene)?"
        if n:
            msg += f"\n\nATENCIÓN: tiene {n} visita(s) registradas que también se borrarán."
        if not ask_yes(self, msg):
            return
        self.store.delete_center(cid)
        self.app.sync(self.cur_d)
        self.v_c.set("")
        self.refresh()

    # --- importación / plantilla ------------------------------------------------------------------------------------------------
    def import_hier(self):
        path = filedialog.askopenfilename(parent=self, title="Excel con Distrito | Zona | Centro",
                                          filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            st = import_hierarchy(self.store, path)
        except Exception as exc:
            show_error(self, DataError(f"No se pudo leer el archivo:\n{exc}"))
            return
        for did in st["district_ids"]:
            self.app.sync(did)
        self.refresh()
        msg = (f"Importación terminada.\n\n  Distritos/Hospitales nuevos: {st['districts']}\n"
               f"  Zonas nuevas: {st['zones']}\n  Centros nuevos: {st['centers']}\n"
               f"  Unidades nuevas: {st['units']}\n  Ya existían: {st['existing']}\n  Omitidas: {st['skipped']}")
        if st["notes"]:
            msg += "\n\n" + "\n".join(st["notes"][:8])
        show_info(self, msg)

    def template(self):
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx", initialdir=BASE_DIR,
                                            initialfile="PLANTILLA_IMPORTAR_ESTRUCTURA.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            create_import_template(path)
        except OSError as exc:
            show_error(self, DataError(f"No se pudo guardar el archivo:\n{exc}"))
            return
        show_info(self, f"Plantilla creada en:\n{path}")

    def import_old(self):
        if not self._need("d"):
            return
        path = filedialog.askopenfilename(parent=self, title="Excel de registro de visitas (formato original)",
                                          filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            st = import_structure(self.store, self.cur_d, path)
        except Exception as exc:
            show_error(self, DataError(f"No se pudo leer el archivo:\n{exc}"))
            return
        self.app.sync(self.cur_d)
        self.refresh()
        show_info(self, f"Importación terminada (hoja «{st['sheet']}»):\n\n"
                        f"  Zonas nuevas: {st['zones']}\n  Centros: {st['centers']}\n"
                        f"  Unidades dentro de centros: {st['units']}\n  Omitidos (repetidos): {st['skipped']}")


# ---- página: trabajadores y administradores ---------------------------------------------------------------------------------------------
class TrabajadoresPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store = app, app.store
        self.district_map, self.district_names = {}, []
        self.color = None          # None = automático
        page_header(self, "Trabajadores y administradores",
                    f"Alta, modificación y baja de usuarios. Contraseña inicial de los nuevos usuarios: {DEFAULT_USER_PASS}")
        top = tk.Frame(self, bg=C_BG)
        top.pack(fill="x", padx=28, pady=(0, 10))
        tk.Label(top, text="Mostrar:", bg=C_BG, fg=C_MUTED, font=F_BASE).pack(side="left")
        ds = list(self.store.districts())
        self.district_map = {d["name"]: d["id"] for d in ds}
        self.district_names = list(self.district_map)
        self.v_filter = tk.StringVar(value=ALL_USERS)
        cb = make_combo(top, self.v_filter, [ALL_USERS, ADMINS_ONLY] + self.district_names, 34)
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: self.load_tree())
        ttk.Button(top, text="Exportar a Excel", style="Gold.TButton", command=self.export).pack(side="right")

        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        body.columnconfigure(0, weight=5, uniform="t")
        body.columnconfigure(1, weight=4, uniform="t")
        body.rowconfigure(0, weight=1)
        wrap, self.tree = make_tree(body, [("u", "Usuario", 125, "w"), ("n", "Nombre y apellidos", 165, "w"),
                                           ("d", "Perfil / Distritos", 160, "w"), ("i", "Inic.", 45, "center")],
                                    height=14)
        wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = card(body)
        form.grid(row=0, column=1, sticky="nsew")
        sf = ScrollFrame(form)
        sf.pack(fill="both", expand=True)
        inner = tk.Frame(sf.inner, bg="white")
        inner.pack(fill="x", padx=20, pady=8)
        inner.columnconfigure(0, weight=1, uniform="f")
        inner.columnconfigure(1, weight=1, uniform="f")
        self.v = {k: tk.StringVar() for k in ("nombre", "apellidos", "dni", "ini", "user")}
        self.v_role = tk.StringVar(value=ROLE_WORKER)
        self.v_active = tk.BooleanVar(value=True)
        self.cb_role = make_combo(inner, self.v_role, [ROLE_WORKER, ROLE_ADMIN], 16)
        self.cb_role.bind("<<ComboboxSelected>>", lambda e: self._on_role_change())
        form_field(inner, 0, "Perfil", self.cb_role, col=0, padx=(0, 8))
        self.e_user = make_entry(inner, self.v["user"], 16)
        form_field(inner, 0, "Usuario (se genera al crear)", self.e_user, col=1)
        form_field(inner, 2, "Nombre", make_entry(inner, self.v["nombre"], 16), col=0, padx=(0, 8))
        form_field(inner, 2, "Apellidos", make_entry(inner, self.v["apellidos"], 16), col=1)
        form_field(inner, 4, "DNI / NIE", make_entry(inner, self.v["dni"], 16), col=0, padx=(0, 8))
        vcmd = (self.register(lambda proposed: len(proposed) <= MAX_INITIALS), "%P")
        self.e_ini = make_entry(inner, self.v["ini"], 16, validate="key", validatecommand=vcmd)
        form_field(inner, 4, f"Iniciales (máx. {MAX_INITIALS})", self.e_ini, col=1)
        ttk.Checkbutton(inner, text="Usuario activo", variable=self.v_active).grid(row=6, column=0, sticky="w", pady=(4, 0))
        self.lbl_ini = tk.Label(inner, text="", bg="white", fg=C_MUTED, font=F_SMALL, wraplength=190, justify="left")
        self.lbl_ini.grid(row=6, column=1, sticky="w", pady=(4, 0))

        lb_wrap = tk.Frame(inner, bg="white")
        form_field(inner, 7, "Distritos / Hospitales (pulse para marcar varios)", lb_wrap, colspan=2)
        self.lb = tk.Listbox(lb_wrap, selectmode="multiple", exportselection=False, height=3, font=F_BASE,
                             relief="solid", bd=1, activestyle="none", highlightthickness=1,
                             highlightbackground=C_LINE, selectbackground=C_GREEN, selectforeground="white")
        sb = ttk.Scrollbar(lb_wrap, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")
        for name in self.district_names:
            self.lb.insert("end", name)

        cf = tk.Frame(inner, bg="white")
        cf.grid(row=9, column=0, columnspan=2, sticky="w", pady=(10, 0))
        tk.Label(cf, text="COLOR EN EL EXCEL", font=F_SMALL_B, fg=C_MUTED, bg="white").pack(side="left")
        self.swatch = tk.Label(cf, text="auto", width=7, bg="#EEEEEE", fg=C_MUTED, font=F_SMALL, relief="solid", bd=1)
        self.swatch.pack(side="left", padx=10)
        self.btn_color = ttk.Button(cf, text="Elegir…", style="SecondarySm.TButton", command=self.pick_color)
        self.btn_color.pack(side="left")

        bt = tk.Frame(inner, bg="white")
        bt.grid(row=10, column=0, columnspan=2, sticky="w", pady=(14, 0))
        self.btn_create = ttk.Button(bt, text="Crear usuario", style="Primary.TButton", command=self.create)
        self.btn_create.pack(side="left")
        self.btn_save = ttk.Button(bt, text="Guardar cambios", style="Secondary.TButton", command=self.save)
        self.btn_save.pack(side="left", padx=6)
        ttk.Button(bt, text="Nuevo", style="Secondary.TButton", command=self.clear).pack(side="left")
        bt2 = tk.Frame(inner, bg="white")
        bt2.grid(row=11, column=0, columnspan=2, sticky="w", pady=(8, 4))
        ttk.Button(bt2, text="Nueva contraseña…", style="Secondary.TButton", command=self.set_pw).pack(side="left")
        ttk.Button(bt2, text="Restablecer", style="Secondary.TButton", command=self.reset_pw).pack(side="left", padx=6)
        ttk.Button(bt2, text="Eliminar", style="Danger.TButton", command=self.delete).pack(side="left")
        self.lbl_note = tk.Label(inner, text="", bg="white", fg=C_MUTED, font=F_SMALL, wraplength=380, justify="left")
        self.lbl_note.grid(row=12, column=0, columnspan=2, sticky="w", pady=(2, 8))

        self.e_ini.bind("<KeyRelease>", self._upper_ini)
        self.v["ini"].trace_add("write", self._on_ini_change)
        self.clear()
        self.load_tree()

    # --- perfil -----------------------------------------------------------------------------------------------
    def _role(self):
        return "admin" if self.v_role.get() == ROLE_ADMIN else "worker"

    def _on_role_change(self, _e=None):
        admin = self._role() == "admin"
        state = "disabled" if admin else "normal"
        self.lb.config(state=state)
        self.e_ini.config(state=state)
        self.btn_color.state(["disabled"] if admin else ["!disabled"])
        if admin:
            self.lbl_ini.config(text="Los administradores acceden a todo; no necesitan distrito, iniciales ni color.",
                                fg=C_MUTED)
        else:
            self._on_ini_change()

    # --- iniciales: comprobación inmediata ------------------------------------------------------------
    def _upper_ini(self, _event=None):
        val = self.v["ini"].get()
        up = val.upper().replace(" ", "")
        if up != val:
            self.v["ini"].set(up)

    def _on_ini_change(self, *_):
        if self._role() == "admin":
            return
        up = self.v["ini"].get().upper().replace(" ", "")
        if not up:
            self.lbl_ini.config(text="" if sel_id(self.tree) else "Vacío: se asignan automáticamente.", fg=C_MUTED)
            return
        owner = self.store.initials_owner(up, sel_id(self.tree))
        if owner:
            self.lbl_ini.config(text=f"Ya las usa {owner}", fg=C_ERR)
        else:
            self.lbl_ini.config(text="✔ Disponibles", fg=C_OK)

    # --- distritos del usuario (selección múltiple) ----------------------------------------------------
    def _selected_districts(self):
        return [self.district_map[self.district_names[i]] for i in self.lb.curselection()]

    def _set_selected(self, ids):
        self.lb.config(state="normal")
        self.lb.selection_clear(0, "end")
        for i, name in enumerate(self.district_names):
            if self.district_map[name] in ids:
                self.lb.selection_set(i)

    # --- tabla ------------------------------------------------------------------------------------------------
    def _list_users(self):
        f = self.v_filter.get()
        if f == ADMINS_ONLY:
            return self.store.users(role="admin")
        did = self.district_map.get(f)
        if did:
            return self.store.users(role="worker", district_id=did)
        return self.store.users()

    def load_tree(self, select=None):
        rows = []
        for u in self._list_users():
            admin = u["role"] == "admin"
            rows.append((u["id"], (u["username"] + ("" if u["active"] else "  (baja)"), u["full_name"],
                                   "Administrador" if admin else (u["district_names"] or ""),
                                   "-" if admin else (u["initials"] or "")), "admin" if admin else ""))
        fill_tree(self.tree, rows)
        if select:
            self.tree.selection_set(str(select))

    def _set_color(self, hex_color):
        self.color = hex_color
        if hex_color:
            self.swatch.config(text="", bg="#" + hex_color)
        else:
            self.swatch.config(text="auto", bg="#EEEEEE")

    def pick_color(self):
        initial = "#" + self.color if self.color else "#FFFFFF"
        _, chosen = colorchooser.askcolor(color=initial, parent=self, title="Color del trabajador")
        if chosen:
            self._set_color(chosen.lstrip("#").upper())

    def clear(self):
        if self.tree.selection():
            self.tree.selection_remove(*self.tree.selection())
        for var in self.v.values():
            var.set("")
        self.v_role.set(ROLE_WORKER)
        self.v_active.set(True)
        self._set_color(None)
        self.e_user.config(state="disabled")
        self.btn_create.state(["!disabled"])
        self.btn_save.state(["disabled"])
        self.lbl_note.config(text="")
        pre = self.district_map.get(self.v_filter.get())
        self._set_selected([pre] if pre else [])
        self._on_role_change()

    def _on_select(self, _e=None):
        uid = sel_id(self.tree)
        u = self.store.user(uid) if uid else None
        if not u:
            return
        self.v_role.set(ROLE_ADMIN if u["role"] == "admin" else ROLE_WORKER)
        self.v["nombre"].set(u["first_name"] or u["full_name"])
        self.v["apellidos"].set(u["last_name"] or "")
        self.v["dni"].set(u["dni"] or "")
        self.v["ini"].set(u["initials"] or "" if u["role"] == "worker" else "")
        self._set_selected(self.store.user_district_ids(uid) if u["role"] == "worker" else [])
        self.e_user.config(state="normal")
        self.v["user"].set(u["username"])
        self.v_active.set(bool(u["active"]))
        self._set_color(u["color"] if u["role"] == "worker" else None)
        self.btn_create.state(["disabled"])
        primary = is_primary_admin(u)
        self.btn_save.state(["disabled"] if primary else ["!disabled"])
        self.lbl_note.config(text="Cuenta principal del sistema: solo se puede cambiar su contraseña "
                                  "(no se puede desactivar, renombrar ni eliminar)." if primary else "")
        self._on_role_change()

    def create(self):
        role = self._role()
        ids = self._selected_districts() if role == "worker" else []
        try:
            uid, username = self.store.add_user(role, self.v["nombre"].get(), self.v["apellidos"].get(),
                                                self.v["dni"].get(), ids, self.v["ini"].get(), self.color or "")
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync_many(ids)
        u = self.store.user(uid)
        extra = f"  Iniciales:  {u['initials']}\n" if role == "worker" else ""
        self.load_tree()
        self.clear()
        show_info(self, f"Usuario creado correctamente.\n\n  Perfil:  {ROLE_ADMIN if role == 'admin' else ROLE_WORKER}\n"
                        f"  Usuario:  {username}\n{extra}  Contraseña inicial:  {DEFAULT_USER_PASS}")

    def save(self):
        uid = sel_id(self.tree)
        if not uid:
            return
        old_ids = self.store.user_district_ids(uid)
        role = self._role()
        ids = self._selected_districts() if role == "worker" else []
        try:
            self.store.update_user(uid, role, self.v["nombre"].get(), self.v["apellidos"].get(), self.v["dni"].get(),
                                   ids, self.v["ini"].get(), self.color or "", self.v["user"].get(),
                                   self.v_active.get(), acting_user_id=self.app.user["id"])
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync_many(set(old_ids) | set(ids))
        self.load_tree(uid)
        show_info(self, "Datos del usuario actualizados.")

    def set_pw(self):
        uid = sel_id(self.tree)
        if not uid:
            show_info(self, "Seleccione un usuario de la lista.")
            return
        dlg = PasswordDialog(self, "Nueva contraseña", "Nueva contraseña para " + self.store.user(uid)["username"])
        if dlg.result:
            try:
                self.store.set_password(uid, dlg.result)
            except DataError as exc:
                show_error(self, exc)
                return
            show_info(self, "Contraseña actualizada.")

    def reset_pw(self):
        uid = sel_id(self.tree)
        if not uid:
            show_info(self, "Seleccione un usuario de la lista.")
            return
        default = default_password_for(self.store.user(uid))
        if ask_yes(self, f"¿Restablecer la contraseña a «{default}»?"):
            self.store.reset_password(uid)
            show_info(self, "Contraseña restablecida.")

    def delete(self):
        uid = sel_id(self.tree)
        if not uid:
            show_info(self, "Seleccione un usuario de la lista.")
            return
        u = self.store.user(uid)
        n = self.store.worker_visit_count(uid)
        msg = f"¿Eliminar al usuario {u['username']}?"
        if n:
            msg += (f"\n\nATENCIÓN: tiene {n} visita(s) registradas que también se borrarán.\n"
                    "Si solo quiere impedir su acceso, desmarque «Usuario activo» y guarde los cambios.")
        if not ask_yes(self, msg):
            return
        dids = self.store.user_district_ids(uid)
        try:
            self.store.delete_user(uid, acting_user_id=self.app.user["id"])
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync_many(dids)
        self.clear()
        self.load_tree()

    def export(self):
        users = [dict(u) for u in self._list_users()]
        if not users:
            show_info(self, "No hay usuarios en la lista para exportar.")
            return
        scope = self.v_filter.get()
        default = f"USUARIOS_{slug(scope)}_{dt.date.today().strftime('%Y%m%d')}.xlsx"
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx", initialdir=BASE_DIR,
                                            initialfile=default, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            export_users(path, "Trabajadores y administradores",
                         f"Mostrando: {scope}  |  Exportado el {fmt_date(dt.date.today())}", users,
                         self.store.visit_counts_by_user())
        except OSError as exc:
            show_error(self, DataError(f"No se pudo guardar el archivo (¿está abierto en Excel?):\n{exc}"))
            return
        self.store.log("Exportaciones", "Exportación de usuarios",
                       f"{len(users)} usuario(s) (filtro: {scope}) → {os.path.basename(path)}")
        show_info(self, f"Se han exportado {len(users)} usuario(s) a:\n{path}")


# ---- página: vacaciones ("V" en el Excel) ---------------------------------------------------------------------------------------------------
class VacacionesPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store, self.user = app, app.store, app.user
        self.is_admin = self.user["role"] == "admin"
        self.district_map, self.worker_map = {}, {}
        page_header(self, "Vacaciones",
                    "Registre los periodos de vacaciones. En el Excel aparecen con una «V» en la fila del trabajador "
                    "(solo de lunes a viernes) y no se pueden anotar visitas esos días.")
        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=28, pady=(0, 24))
        body.columnconfigure(0, weight=4, uniform="v")
        body.columnconfigure(1, weight=5, uniform="v")
        body.rowconfigure(0, weight=1)

        form = card(body)
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        inner = tk.Frame(form, bg="white")
        inner.pack(fill="both", expand=True, padx=26, pady=16)
        inner.columnconfigure(0, weight=1)
        inner.columnconfigure(1, weight=1)
        self.v_district, self.v_worker = tk.StringVar(), tk.StringVar()
        r = 0
        if self.is_admin:
            ds = list(self.store.districts())
            self.district_map = {d["name"]: d["id"] for d in ds}
            self.cb_district = make_combo(inner, self.v_district, list(self.district_map), 28)
            self.cb_district.bind("<<ComboboxSelected>>", lambda e: self._on_district())
            form_field(inner, r, "Distrito / Hospital", self.cb_district, colspan=2)
            r += 2
            self.cb_worker = make_combo(inner, self.v_worker, [], 28)
            form_field(inner, r, "Nombre (trabajador)", self.cb_worker, colspan=2)
        else:
            self.v_worker.set(self.user["full_name"])
            form_field(inner, r, "Nombre", make_entry(inner, self.v_worker, 28, state="readonly"), colspan=2)
        r += 2
        today = dt.date.today()
        self.d_from, self.d_to = DateEntry(inner, today), DateEntry(inner, today)
        form_field(inner, r, "Desde", self.d_from, col=0, padx=(0, 8))
        form_field(inner, r, "Hasta", self.d_to, col=1)
        r += 2
        ttk.Button(inner, text="Registrar vacaciones", style="Primary.TButton", command=self.save).grid(
            row=r, column=0, columnspan=2, sticky="w", pady=(18, 0))
        self.lbl_status = tk.Label(inner, text="", bg="white", fg=C_OK, font=F_BOLD, wraplength=340, justify="left")
        self.lbl_status.grid(row=r + 1, column=0, columnspan=2, sticky="w", pady=(10, 0))

        side = card(body)
        side.grid(row=0, column=1, sticky="nsew")
        tk.Label(side, text="Periodos registrados", bg="white", fg=C_GREEN, font=F_H2).pack(anchor="w", padx=20, pady=(18, 8))
        cols = ([("q", "Trabajador", 170, "w")] if self.is_admin else []) + \
            [("a", "Desde", 90, "center"), ("b", "Hasta", 90, "center"), ("n", "Días lab.", 70, "center")]
        wrap, self.tree = make_tree(side, cols, height=12)
        wrap.pack(fill="both", expand=True, padx=20)
        ttk.Button(side, text="Eliminar periodo seleccionado", style="Danger.TButton", command=self.delete).pack(
            anchor="e", padx=20, pady=14)

        if self.is_admin:
            if ds:
                self.v_district.set(ds[0]["name"])
                self._on_district()
            else:
                self.lbl_status.config(text="Cree primero un Distrito/Hospital y sus trabajadores.", fg=C_ERR)
        else:
            self.load()

    def _on_district(self):
        did = self.district_map.get(self.v_district.get())
        ws = self.store.workers(did, only_active=True)
        self.worker_map = {w["full_name"]: w for w in ws}
        self.cb_worker["values"] = list(self.worker_map)
        self.v_worker.set(ws[0]["full_name"] if ws else "")
        self.load()

    def load(self):
        if self.is_admin:
            did = self.district_map.get(self.v_district.get())
            rows = self.store.absences(district_id=did) if did else []
        else:
            rows = self.store.absences(user_id=self.user["id"])
        data = []
        for a in rows:
            d1, d2 = dt.date.fromisoformat(a["start_date"]), dt.date.fromisoformat(a["end_date"])
            vals = ([a["full_name"]] if self.is_admin else []) + [fmt_date(d1), fmt_date(d2), workdays(d1, d2)]
            data.append((a["id"], vals))
        fill_tree(self.tree, data)

    def save(self):
        self.lbl_status.config(text="")
        try:
            if self.is_admin:
                worker = self.worker_map.get(self.v_worker.get())
                if not worker:
                    raise DataError("Seleccione el trabajador.")
            else:
                worker = self.user
            d1, d2 = self.d_from.get_date(), self.d_to.get_date()
            self.store.add_absence(worker["id"], d1.isoformat(), d2.isoformat(), self.user["id"])
        except DataError as exc:
            show_error(self, exc)
            return
        ok = self.app.sync_many(self.store.user_district_ids(worker["id"]), list(range(d1.year, d2.year + 1)))
        self.lbl_status.config(text="✔ Vacaciones registradas" + (" y Excel actualizado." if ok else "."),
                               fg=C_OK if ok else C_ERR)
        self.load()

    def delete(self):
        aid = sel_id(self.tree)
        if aid is None:
            show_info(self, "Seleccione un periodo de la lista.")
            return
        if not ask_yes(self, "¿Eliminar el periodo de vacaciones seleccionado?"):
            return
        try:
            district_ids, years = self.store.delete_absence(aid, self.user)
        except DataError as exc:
            show_error(self, exc)
            return
        self.app.sync_many(district_ids, years)
        self.load()


# ---- página: registro de actividad (solo administrador) ----------------------------------------------------------------------------------------
class LogPage(tk.Frame):
    LIMIT = 3000

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store, self.user = app, app.store, app.user
        self.rows, self.user_map, self.nums = {}, {}, {}
        page_header(self, "Registro de actividad",
                    "Accesos de los usuarios y cambios que realizan en cada acceso. Solo visible para administradores.")
        filt = card(self)
        filt.pack(fill="x", padx=28, pady=(0, 10))
        fi = tk.Frame(filt, bg="white")
        fi.pack(fill="x", padx=18, pady=(2, 14))
        today = dt.date.today()
        self.d_from, self.d_to = DateEntry(fi, today - dt.timedelta(days=30)), DateEntry(fi, today)
        form_field(fi, 0, "Desde", self.d_from, col=0, padx=(0, 12))
        form_field(fi, 0, "Hasta", self.d_to, col=1, padx=(0, 12))
        users = list(self.store.audit_users(self.user))
        self.user_map = {f"{u['full_name'] or u['username']} ({u['username']})": u["username"] for u in users}
        self.v_user, self.v_cat, self.v_text = tk.StringVar(value=ALL), tk.StringVar(value="Todas"), tk.StringVar()
        form_field(fi, 0, "Usuario", make_combo(fi, self.v_user, [ALL] + list(self.user_map), 32), col=2, padx=(0, 12))
        form_field(fi, 0, "Categoría", make_combo(fi, self.v_cat, ["Todas"] + AUDIT_CATEGORIES, 16), col=3)
        form_field(fi, 2, "Buscar en el detalle", make_entry(fi, self.v_text, 36), col=0, colspan=3, padx=(0, 12))
        bt = tk.Frame(fi, bg="white")
        bt.grid(row=3, column=3, sticky="w")
        ttk.Button(bt, text="Filtrar", style="Primary.TButton", command=self.load).pack(side="left")
        ttk.Button(bt, text="Exportar a Excel", style="Gold.TButton", command=self.export).pack(side="left", padx=8)

        wrap, self.tree = make_tree(self, [("ts", "Fecha y hora", 135, "center"), ("ses", "Acceso", 55, "center"),
                                           ("user", "Usuario", 140, "w"), ("cat", "Categoría", 95, "w"),
                                           ("act", "Acción", 200, "w"), ("det", "Detalle", 420, "w")], height=12)
        wrap.pack(fill="both", expand=True, padx=28)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        bottom = tk.Frame(self, bg=C_BG)
        bottom.pack(fill="x", padx=28, pady=(10, 18))
        self.lbl_count = tk.Label(bottom, text="", bg=C_BG, fg=C_MUTED, font=F_BASE)
        self.lbl_count.pack(side="right")
        self.lbl_detail = tk.Label(bottom, text="", bg=C_BG, fg=C_TEXT, font=F_BASE, justify="left", wraplength=760)
        self.lbl_detail.pack(side="left", anchor="w")
        self.load()

    def _filters(self):
        d1, d2 = self.d_from.get_date(True), self.d_to.get_date(True)
        if d1 and d2 and d1 > d2:
            raise DataError("La fecha «Desde» no puede ser posterior a «Hasta».")
        cat = self.v_cat.get()
        return (d1.isoformat() if d1 else None, d2.isoformat() if d2 else None,
                self.user_map.get(self.v_user.get()), None if cat == "Todas" else cat, self.v_text.get().strip() or None)

    def load(self):
        try:
            d1, d2, uname, cat, text = self._filters()
            entries = self.store.audit_entries(self.user, d1, d2, uname, cat, text, limit=self.LIMIT + 1)
        except DataError as exc:
            show_error(self, exc)
            return
        self.nums = self.store.session_numbers(self.user)
        truncated = len(entries) > self.LIMIT
        entries = entries[:self.LIMIT]
        self.rows = {e["id"]: e for e in entries}
        data = []
        for e in entries:
            det = (e["details"] or "").replace("\n", " ")
            data.append((e["id"], (dt.datetime.fromisoformat(e["ts"]).strftime("%d/%m/%Y %H:%M:%S"),
                                   self.nums.get(e["session"], ""), e["username"] or "-", e["category"], e["action"],
                                   det if len(det) <= 90 else det[:89] + "…"),
                         "admin" if e["action"] == "Inicio de sesión" else ""))
        fill_tree(self.tree, data)
        self.lbl_count.config(text=f"{len(entries)} registro(s)" +
                              (f" (los {self.LIMIT} más recientes; el Excel incluye todos)" if truncated else ""))
        self.lbl_detail.config(text="")

    def _on_select(self, _e=None):
        e = self.rows.get(sel_id(self.tree))
        if e:
            self.lbl_detail.config(text=f"{e['full_name'] or e['username']}: {e['details'] or '(sin detalle)'}")

    def export(self):
        try:
            d1, d2, uname, cat, text = self._filters()
            entries = self.store.audit_entries(self.user, d1, d2, uname, cat, text, order="asc")
            sessions = self.store.audit_sessions(self.user, d1, d2, uname)
        except DataError as exc:
            show_error(self, exc)
            return
        if not entries:
            show_info(self, "No hay registros para exportar con los filtros actuales.")
            return
        filt = (f"Periodo: {fmt_date(dt.date.fromisoformat(d1)) if d1 else 'sin límite'} a "
                f"{fmt_date(dt.date.fromisoformat(d2)) if d2 else 'sin límite'}  |  Usuario: {self.v_user.get()}  |  "
                f"Categoría: {self.v_cat.get()}" + (f"  |  Texto: {text}" if text else "") +
                f"  |  Exportado el {fmt_date(dt.date.today())} por {self.user['username']}")
        default = f"REGISTRO_ACTIVIDAD_{dt.date.today().strftime('%Y%m%d')}.xlsx"
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx", initialdir=BASE_DIR,
                                            initialfile=default, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            export_audit(path, "Registro de actividad", filt, entries, sessions, self.store.session_numbers(self.user))
        except OSError as exc:
            show_error(self, DataError(f"No se pudo guardar el archivo (¿está abierto en Excel?):\n{exc}"))
            return
        self.store.log("Exportaciones", "Exportación del registro de actividad",
                       f"{len(entries)} registro(s), {len(sessions)} acceso(s) → {os.path.basename(path)}")
        self.load()
        show_info(self, f"Registro exportado ({len(entries)} registros, {len(sessions)} accesos) a:\n{path}")


# ---- página: archivos Excel ----------------------------------------------------------------------------------------------------------------
class ExcelPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app, self.store = app, app.store
        page_header(self, "Archivos Excel",
                    "Un archivo por Distrito/Hospital y año, con una pestaña por mes. Se guardan junto al programa.")
        c = card(self)
        c.pack(fill="both", expand=True, padx=28, pady=(0, 24))
        tk.Label(c, text=BASE_DIR, bg="white", fg=C_MUTED, font=F_SMALL).pack(anchor="w", padx=20, pady=(16, 8))
        wrap, self.tree = make_tree(c, [("f", "Archivo", 460, "w"), ("m", "Última modificación", 150, "center"),
                                        ("s", "Tamaño", 80, "center")], height=10)
        wrap.pack(fill="both", expand=True, padx=20)
        bt = tk.Frame(c, bg="white")
        bt.pack(fill="x", padx=20, pady=16)
        ttk.Button(bt, text="Abrir archivo", style="Primary.TButton", command=self.open_file).pack(side="left")
        ttk.Button(bt, text="Abrir carpeta", style="Secondary.TButton", command=self.open_folder).pack(side="left", padx=6)
        ttk.Button(bt, text="Regenerar todos", style="Gold.TButton", command=self.regen).pack(side="left", padx=(6, 24))
        tk.Label(bt, text="Crear año:", bg="white", fg=C_MUTED, font=F_BASE).pack(side="left")
        y = dt.date.today().year
        self.v_year = tk.StringVar(value=str(y))
        make_combo(bt, self.v_year, [str(i) for i in range(y - 1, y + 4)], 8).pack(side="left", padx=8)
        ttk.Button(bt, text="Generar", style="Secondary.TButton", command=self.gen_year).pack(side="left")
        self.load()

    def load(self):
        rows = []
        for fn in sorted(os.listdir(BASE_DIR)):
            if fn.startswith("REGISTRO_VISITAS_") and fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                p = os.path.join(BASE_DIR, fn)
                rows.append((fn, (fn, dt.datetime.fromtimestamp(os.path.getmtime(p)).strftime("%d/%m/%Y %H:%M"),
                                  f"{os.path.getsize(p) // 1024} KB")))
        fill_tree(self.tree, rows)

    def open_file(self):
        s = self.tree.selection()
        if not s:
            show_info(self, "Seleccione un archivo de la lista.")
            return
        try:
            open_path(os.path.join(BASE_DIR, s[0]))
        except DataError as exc:
            show_error(self, exc)

    def open_folder(self):
        try:
            open_path(BASE_DIR)
        except DataError as exc:
            show_error(self, exc)

    def regen(self):
        for d in self.store.districts():
            self.app.sync(d["id"])
        self.load()
        show_info(self, "Archivos Excel regenerados.")

    def gen_year(self):
        year = int(self.v_year.get())
        for d in self.store.districts():
            self.app.sync(d["id"], [year])
        self.load()
        show_info(self, f"Archivos del año {year} generados.")


PAGES = {"registro": RegistroPage, "visitas": VisitasPage, "vacaciones": VacacionesPage, "stats": StatsPage,
         "cuenta": CuentaPage, "estructura": EstructuraPage, "trabajadores": TrabajadoresPage, "excel": ExcelPage,
         "log": LogPage}


# --------------------------------------------------------------------------
#  LOGO CSIF SEVILLA SANIDAD con halo dorado (PNG de 360 px, y un icono de 64 px, incrustados en base64;
#  así el programa sigue
#  siendo un único archivo y no depende de librerías de imagen)
# --------------------------------------------------------------------------
LOGO_PNG_B64 = (
    "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"
    "7FG/3bT1tmt2ptl3bWWafdfu27qq6ma9M+u9Xe/MujJB0Wi46lxdm6btWU1nd986fuXt2c07ZTEpCz0py8l0Mp1MJtNZWUxUobQutFL+5eXYGj"
    "atMW3XNX3b9OFt2/e9Mb3pemvCnKywdT7vxguW2WIEEUF0HBc3okTV3sfwStBCQ/kTtLyIGCBIGO8bLAyEJH4dK6ZUFIUKAfyEPShUqOM+mpJQ"
    "IUrQbkUESZChIN/WFEJkJoXs1yjFpMBOq07d2kxuzM3+yGyW7d72tjO67qWa8KJz1VRVU7WcqsVMLaZqPuFpaYtCa1eQLpRypJywDfARe7ktYh"
    "/KsGEshTiYzTFjH0OMS847gNEFEECMN2wgGIDkuQU4hjBfF7HDd7Yy7JCrsCN2ai9iR35efivY8dsoVa5BDRznawwzr8N+kBjqBZmRPAaRR9QI"
    "eeIpRCPZtwLX0FgqVeoRdmSoQYUmnc2qpPKk8EQD6M5bf3nvnT9Vuob9j+161Vabtt71iWs0+7bZV1Wz2fWrnVntzWZv13uz9a3Wlqu6b1oD5X"
    "z+5rvLV96cHd2cFMVEq9lsPpvNJuW0nJRFWSqlUZE3DFjnemMb09Zts2urqmubtmm6rnOmt31njbXWOuucNeJC/jCDA46eLU6yPPoV7SICLoCC"
    "AIi/B4HIsFVWPMeWGHaa3kPpFRtTiyWueUrLuQHJ+kw9Ir8gwcf7KEW9CrKyCntcMropfmVe2C+FjhBZPPf3kcSOZFss93o5c92RXR13W2e63u"
    "imV3XPdcdV55adWs48+1DziZuWrigs+0UQSqNywk5IMTsiNSNdkGrIf0fCSD289oGOCQEJiNJGG457eTi5vxCBPNEAQN+JdRyfRwLHqJ6NHXmf"
    "JWHHdXO0v/UGrf7N042Lzdenoka+yTV1Xoee6yXUKGI2RIHJ7ol5h9XDRKlo4l1eg8ARDWDKlyc+lccbyUkrZFcvlq+8+t7fHdy+C+ZjPnnUbb"
    "dtvembXd/su2bfN7u22ddVtdl3q61Z7cx6Z1Z7s63stnFVy/u6bxuD88PFO984ePmN6eHxtChmZTmfz+eTWTmdTqaToiiU0oLAIFZsZ0xr+qqt"
    "d029a/dVXdV925m+7U3vesfsnDM+WxREWLxsAYwi7FMx/dMsof4QBGJhH5QrwJiv1Q17E9JyEP+LEU6e0rE/DLytNP0CRYTDsgb0I6IgAIQoRO"
    "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"
)

LOGO_ICON_B64 = (
    "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"
)


# --------------------------------------------------------------------------
def main():
    if "--reset-admin" in sys.argv:
        s = Store()
        admin = s._one("SELECT id FROM users WHERE username=? COLLATE NOCASE", (ADMIN_USER,))
        s.reset_password(admin["id"])
        print(f"Contraseña del administrador restablecida a: {ADMIN_PASS}")
        return
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()

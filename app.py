"""Local trial server. Run with Python 3.10 or later."""
import base64
import binascii
import hmac
import io
import json
import os
import re
import secrets
import sqlite3
import time
from contextlib import closing
from datetime import datetime, timedelta, timezone
from html import escape
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parent
DB = Path(os.environ.get("JINOU_DB", str(ROOT / "registrations.sqlite3")))
TAIPEI = timezone(timedelta(hours=8))
PASSWORD = os.environ.get("JINOU_ADMIN_PASSWORD") or secrets.token_urlsafe(12)
ADMIN = {}
DEPARTMENTS = dict(zip(("101", "401", "407", "430", "433", "434"),
    ("普通科", "商業經營科", "觀光事業科", "多媒體設計科", "應用英語科", "應用日語科")))
FIELDS = {
    "sno":"報名編號", "registerDate":"登錄日期", "department":"報考科別",
    "applicationGrade":"報考年級", "studentName":"姓名", "originalSchoolName":"原學校名稱",
    "originalDepartment":"原學校科組", "originalStudentId":"原學校學號",
    "originalSemester":"原學校學期", "studentPhone":"學生手機", "studentEmail":"學生信箱",
    "householdZip":"戶籍區號", "phone":"電話", "householdAddress":"戶籍地址",
    "mailZip":"郵遞區號", "mailAddress":"通訊地址", "contactPerson":"連絡人", "contactPhone":"連絡電話",
}
REQUIRED = ("studentName", "originalSchoolName", "originalDepartment", "studentPhone",
            "studentEmail", "householdAddress", "mailAddress")


def connect():
    db = sqlite3.connect(DB, timeout=10)
    db.row_factory = sqlite3.Row
    return db


def initialise():
    DB.parent.mkdir(parents=True, exist_ok=True)
    with closing(connect()) as db, db:
        db.execute("""CREATE TABLE IF NOT EXISTS applications (
            sno TEXT PRIMARY KEY, token TEXT UNIQUE NOT NULL, register_date TEXT NOT NULL,
            data TEXT, deleted INTEGER NOT NULL DEFAULT 0)""")


def image(value, kinds, limit):
    if not isinstance(value, str):
        raise ValueError("請上傳相片並完成簽名。")
    match = re.fullmatch(r"data:image/(png|jpeg);base64,([A-Za-z0-9+/=]+)", value)
    if not match or match[1] not in kinds:
        raise ValueError("請使用 JPG 或 PNG 圖檔。")
    raw = base64.b64decode(match[2], validate=True)
    magic = b"\x89PNG\r\n\x1a\n" if match[1] == "png" else b"\xff\xd8\xff"
    if len(raw) > limit or not raw.startswith(magic):
        raise ValueError("圖檔格式錯誤或超過大小限制。")
    return value


def clean(payload, row):
    data = {}
    for field in FIELDS:
        value = payload.get(field, "")
        if not isinstance(value, str) or len(value) > 250:
            raise ValueError("欄位格式不符或超過 250 字。")
        data[field] = value.strip()
    data["sno"] = row["sno"]
    data["registerDate"] = row["register_date"]
    if any(not data[key] for key in REQUIRED):
        raise ValueError("請填寫所有必填欄位。")
    if data["department"] not in DEPARTMENTS or data["applicationGrade"] not in ("高一", "高二"):
        raise ValueError("請選擇一個科別及一個年級。")
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", data["studentEmail"]):
        raise ValueError("請輸入有效的電子郵件。")
    data["photo"] = image(payload.get("photo"), {"jpeg", "png"}, 3 * 1024**2)
    data["signature"] = image(payload.get("signature"), {"png"}, 1024**2)
    return data


def records(db):
    return [json.loads(row["data"]) for row in db.execute(
        "SELECT data FROM applications WHERE data IS NOT NULL AND deleted=0 ORDER BY sno")]


def cell_text(value):
    return xml_escape("".join(character for character in value
                              if character in "\t\n\r" or ord(character) >= 32))


def excel(items):
    # Inline string cells retain leading zeros and prevent spreadsheet formula execution.
    columns = list(FIELDS)
    rows = [list(FIELDS.values()) + ["相片", "簽名"]]
    for item in items:
        row = [DEPARTMENTS[item[k]] if k == "department" else item[k] for k in columns]
        rows.append(row + ["已上傳", "已簽名"])
    sheet_rows = "".join(f'<row r="{i}">' + "".join(
        f'<c r="{chr(65+j)}{i}" t="inlineStr"><is><t xml:space="preserve">{cell_text(value)}</t></is></c>'
        for j, value in enumerate(row)) + "</row>" for i, row in enumerate(rows, 1))
    output = io.BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as book:
        book.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        package = "http://schemas.openxmlformats.org/package/2006/relationships"
        book.writestr("_rels/.rels", f'<Relationships xmlns="{package}"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        book.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{package}"><Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        book.writestr("xl/workbook.xml", f'<workbook xmlns="{main}" xmlns:r="{rel}"><sheets><sheet name="報名總表" sheetId="1" r:id="rId1"/></sheets></workbook>')
        book.writestr("xl/worksheets/sheet1.xml", f'<worksheet xmlns="{main}"><sheetData>{sheet_rows}</sheetData></worksheet>')
    return output.getvalue()


class Handler(BaseHTTPRequestHandler):
    def send_body(self, status, body, content="application/json; charset=utf-8", headers=()):
        if not isinstance(body, bytes):
            body = json.dumps(body, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", content)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("X-Frame-Options", "DENY")
        for key, value in headers:
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def error(self, status, message):
        return self.send_body(status, {"error": message})

    def cookie(self, name):
        try:
            jar = SimpleCookie(self.headers.get("Cookie", ""))
            return jar[name].value if name in jar else ""
        except (ValueError, KeyError):
            return ""

    def authenticated(self):
        return ADMIN.get(self.cookie("jinou_admin"), 0) > time.monotonic()

    def page(self, name):
        return self.send_body(200, (ROOT / "web" / name).read_bytes(), "text/html; charset=utf-8")

    def do_GET(self):
        if self.path == "/":
            return self.page("apply.html")
        if self.path == "/admin":
            return self.page("admin.html" if self.authenticated() else "login.html")
        if self.path == "/api/admin/records" or self.path == "/api/admin/export":
            if not self.authenticated():
                return self.error(401, "請重新登入管理後台。")
            with closing(connect()) as db:
                items = records(db)
            if self.path.endswith("export"):
                return self.send_body(200, excel(items), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      (("Content-Disposition", 'attachment; filename="registrations.xlsx"'),))
            return self.send_body(200, {"records": items})
        return self.error(404, "找不到頁面。")

    def do_POST(self):
        self.mutate()
    def do_PUT(self):
        self.mutate()
    def do_DELETE(self):
        self.mutate()

    def mutate(self):
        if self.headers.get("Content-Type", "").split(";")[0] != "application/json":
            return self.error(415, "請使用 JSON 格式。")
        if self.headers.get("Sec-Fetch-Site", "") == "cross-site":
            return self.error(403, "請從本網站操作。")
        try:
            size = int(self.headers.get("Content-Length", "0"))
            if not 0 < size <= 6 * 1024**2:
                return self.error(413, "資料大小超出限制。")
            payload = json.loads(self.rfile.read(size))
            if not isinstance(payload, dict):
                raise ValueError("資料格式有誤。")
            with closing(connect()) as db, db:
                self.action(db, payload)
        except (ValueError, binascii.Error) as exc:
            return self.error(400, str(exc))
        except sqlite3.Error:
            return self.error(503, "資料庫暫時無法作業，請稍後再試。")

    def action(self, db, payload):
        if self.command == "POST" and self.path == "/api/admin/login":
            supplied = payload.get("password")
            if not isinstance(supplied, str) or not hmac.compare_digest(supplied.encode(), PASSWORD.encode()):
                return self.error(401, "密碼錯誤。")
            token = secrets.token_urlsafe(32)
            expiry = time.monotonic()
            for key, end in list(ADMIN.items()):
                if end <= expiry:
                    ADMIN.pop(key, None)
            ADMIN[token] = expiry + 8 * 3600
            return self.send_body(200, {}, headers=(("Set-Cookie",
                f"jinou_admin={token}; Path=/; HttpOnly; SameSite=Strict; Max-Age=28800"),))
        if self.command == "POST" and self.path == "/api/admin/logout":
            ADMIN.pop(self.cookie("jinou_admin"), None)
            return self.send_body(200, {}, headers=(("Set-Cookie",
                "jinou_admin=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0"),))

        if self.command == "POST" and self.path in ("/api/application/start", "/api/application/submit"):
            token = self.cookie("jinou_candidate")
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM applications WHERE token=?", (token,)).fetchone()
            if self.path.endswith("start"):
                if row is None:
                    now = datetime.now(TAIPEI)
                    sno = now.strftime("%y%m%d%H%M%S")
                    token = secrets.token_urlsafe(32)
                    inserted = db.execute("INSERT OR IGNORE INTO applications(sno,token,register_date) VALUES(?,?,?)",
                                          (sno, token, now.date().isoformat()))
                    if inserted.rowcount == 0:
                        return self.error(409, "正在配號，請稍候。")
                    row = db.execute("SELECT * FROM applications WHERE token=?", (token,)).fetchone()
                if row["deleted"]:
                    return self.error(410, "這筆報名已由校方註銷。")
                db.commit()
                return self.send_body(200, {"sno": row["sno"], "date": row["register_date"],
                    "record": json.loads(row["data"]) if row["data"] else None},
                    headers=(("Set-Cookie", f"jinou_candidate={token}; Path=/; HttpOnly; SameSite=Strict; Max-Age=2592000"),))
            if row is None or row["deleted"]:
                return self.error(403, "找不到有效報名編號，請重新開啟報名頁。")
            data = json.loads(row["data"]) if row["data"] else clean(payload, row)
            if not row["data"]:
                db.execute("UPDATE applications SET data=? WHERE sno=?", (json.dumps(data, ensure_ascii=False), row["sno"]))
            db.commit()
            return self.send_body(200, {"record": data})

        match = re.fullmatch(r"/api/admin/records/(\d{12})", self.path)
        if match and self.command in ("PUT", "DELETE"):
            if not self.authenticated():
                return self.error(401, "請重新登入管理後台。")
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM applications WHERE sno=? AND data IS NOT NULL AND deleted=0",
                             (match[1],)).fetchone()
            if not row:
                return self.error(404, "找不到報名資料。")
            if self.command == "DELETE":
                db.execute("UPDATE applications SET deleted=1, data=NULL WHERE sno=?", (row["sno"],))
                db.commit()
                return self.send_body(200, {})
            data = clean(payload, row)
            db.execute("UPDATE applications SET data=? WHERE sno=?", (json.dumps(data, ensure_ascii=False), row["sno"]))
            db.commit()
            return self.send_body(200, {"record": data})
        return self.error(404, "找不到頁面。")


if __name__ == "__main__":
    initialise()
    port = int(os.environ.get("JINOU_PORT", "8000"))
    print(f"考生報名 http://127.0.0.1:{port}/\n管理後台 http://127.0.0.1:{port}/admin")
    if not os.environ.get("JINOU_ADMIN_PASSWORD"):
        print(f"本次管理密碼：{PASSWORD}")
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()

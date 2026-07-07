import os
import uuid
import json
import smtplib
import threading
import time
import base64
import tempfile
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from werkzeug.utils import secure_filename
from flask import (
    Flask, render_template, request, jsonify, redirect, url_for,
    session as flask_session, Response, stream_with_context, send_from_directory
)
import pandas as pd
import db

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "cert-app-secret-2024")
app.permanent_session_lifetime = 86400  # 24 hours

# Temp folder for generated PDFs — only used briefly for email attachments
GENERATED_FOLDER = tempfile.mkdtemp(prefix="certgen_")

ALLOWED_EXCEL = {"xlsx", "xls", "csv"}
ALLOWED_PDF = {"pdf"}

# Per-session progress dicts (in-memory, keyed by session_id)
# These don't need to be in DB — they're only relevant while the operation is running
_generate_progress: dict = {}   # session_id -> {total, done_count, done, running}
_email_progress: dict = {}      # session_id -> {total, sent, failed, current, done, running}
_progress_lock = threading.Lock()


# ── session helpers ───────────────────────────────────────────────────────────

def get_sid() -> str:
    """Get or create a unique session ID for this browser."""
    if "sid" not in flask_session:
        flask_session["sid"] = str(uuid.uuid4())
        flask_session.permanent = True
    return flask_session["sid"]


def get_default_progress():
    return {"total": 0, "done_count": 0, "done": False, "running": False}


def get_default_email_progress():
    return {"total": 0, "sent": 0, "failed": 0, "current": "", "done": False, "running": False}


# ── certificate generation (works entirely from bytes) ────────────────────────

def allowed_file(filename, allowed):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed


def pdf_bytes_to_image_base64(pdf_bytes: bytes) -> str:
    import fitz
    doc = fitz.open("pdf", pdf_bytes)
    page = doc[0]
    mat = fitz.Matrix(2.0, 2.0)
    pix = page.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    doc.close()
    return base64.b64encode(img_bytes).decode("utf-8")


def generate_certificate_pdf_bytes(template_bytes: bytes, student: dict, editor_settings: dict) -> bytes:
    import fitz
    from reportlab.pdfgen import canvas as rl_canvas

    doc = fitz.open("pdf", template_bytes)
    page = doc[0]
    page_width = page.rect.width
    page_height = page.rect.height

    font_family = editor_settings.get("font_family", "Helvetica")
    font_size = int(editor_settings.get("font_size", 24))
    text_color_hex = editor_settings.get("text_color", "#000000")
    bold = editor_settings.get("bold", False)
    italic = editor_settings.get("italic", False)

    r = int(text_color_hex[1:3], 16) / 255.0
    g = int(text_color_hex[3:5], 16) / 255.0
    b = int(text_color_hex[5:7], 16) / 255.0

    rl_font_map = {
        "Helvetica": {
            "normal": "Helvetica", "bold": "Helvetica-Bold",
            "italic": "Helvetica-Oblique", "bolditalic": "Helvetica-BoldOblique"
        },
        "Times New Roman": {
            "normal": "Times-Roman", "bold": "Times-Bold",
            "italic": "Times-Italic", "bolditalic": "Times-BoldItalic"
        },
        "Courier": {
            "normal": "Courier", "bold": "Courier-Bold",
            "italic": "Courier-Oblique", "bolditalic": "Courier-BoldOblique"
        },
        "Arial": {
            "normal": "Helvetica", "bold": "Helvetica-Bold",
            "italic": "Helvetica-Oblique", "bolditalic": "Helvetica-BoldOblique"
        },
    }

    style_key = "bolditalic" if bold and italic else ("bold" if bold else ("italic" if italic else "normal"))
    font_variants = rl_font_map.get(font_family, rl_font_map["Helvetica"])
    rl_font = font_variants.get(style_key, font_variants["normal"])

    overlay_buf = BytesIO()
    c = rl_canvas.Canvas(overlay_buf, pagesize=(page_width, page_height))
    c.setFillColorRGB(r, g, b)
    c.setFont(rl_font, font_size)

    fields = {
        "name": student.get("name", ""),
        "roll": student.get("roll", ""),
        "college": student.get("college", ""),
    }

    positions = editor_settings.get("positions", {})
    for field_key, text in fields.items():
        pos = positions.get(field_key)
        if not pos or not text:
            continue
        x_pct = float(pos["x"])
        y_pct = float(pos["y"])
        x = (x_pct / 100.0) * page_width
        y_from_top = (y_pct / 100.0) * page_height
        y_rl = page_height - y_from_top
        c.drawCentredString(x, y_rl, str(text))

    c.save()
    overlay_buf.seek(0)

    overlay_doc = fitz.open("pdf", overlay_buf.read())
    page.show_pdf_page(page.rect, overlay_doc, 0)

    result = doc.tobytes(garbage=4, deflate=True)
    doc.close()
    overlay_doc.close()
    return result


# ── routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    sid = get_sid()
    db.touch_session(sid) if _session_exists(sid) else None
    state = db.get_state(sid)
    students = state["students"]
    total = len(students)
    generated = sum(1 for s in students if s.get("cert_path"))
    sent = sum(1 for s in students if s.get("email_status") == "sent")
    failed = sum(1 for s in students if s.get("email_status") == "failed")
    pending = sum(1 for s in students if s.get("email_status") == "pending")
    template_ready = db.get_template_bytes(sid) is not None
    return render_template("index.html",
        total=total, generated=generated, sent=sent, failed=failed, pending=pending,
        template_ready=template_ready, session_data=state["editor_settings"])


def _session_exists(sid: str) -> bool:
    try:
        with db._Conn() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1 FROM sessions WHERE session_id = %s", (sid,))
                return cur.fetchone() is not None
    except Exception:
        return False


@app.route("/upload")
def upload_page():
    return render_template("upload.html")


@app.route("/_/upload-excel", methods=["POST"])
def upload_excel():
    sid = get_sid()
    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400
    file = request.files["file"]
    if not file.filename or not allowed_file(file.filename, ALLOWED_EXCEL):
        return jsonify({"error": "Invalid file type. Use .xlsx, .xls or .csv"}), 400

    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    file_bytes = file.read()

    try:
        bio = BytesIO(file_bytes)
        if ext == ".csv":
            df = pd.read_csv(bio)
        elif ext == ".xls":
            df = pd.read_excel(bio, engine="xlrd")
        else:
            df = pd.read_excel(bio, engine="openpyxl")

        columns = list(df.columns)
        total_rows = len(df)
        records = df.fillna("").to_dict(orient="records")
        sample = records[:3]

        # Ensure session exists before saving
        db.ensure_session(sid)
        db.save_keys(sid, {
            "excel_records": records,
            "excel_columns": columns,
        })

        return jsonify({"columns": columns, "total_rows": total_rows, "sample": sample})
    except Exception as e:
        return jsonify({"error": f"Failed to read file: {str(e)}"}), 500


@app.route("/_/save-mapping", methods=["POST"])
def save_mapping():
    sid = get_sid()
    data = request.json or {}
    mapping = data.get("mapping", {})
    required = ["name", "roll", "college", "email"]
    for k in required:
        if k not in mapping or not mapping[k]:
            return jsonify({"error": f"Missing mapping for {k}"}), 400

    state = db.get_state(sid)
    records = state.get("excel_records") or []
    if not records:
        return jsonify({"error": "No data found, please upload the sheet again"}), 400

    students = []
    next_id = 1
    for row in records:
        name = str(row.get(mapping["name"], "")).strip()
        roll = str(row.get(mapping["roll"], "")).strip()
        college = str(row.get(mapping["college"], "")).strip()
        email = str(row.get(mapping["email"], "")).strip()
        if name and name.lower() != "nan":
            students.append({
                "id": next_id,
                "name": name,
                "roll": roll,
                "college": college,
                "email": email,
                "cert_path": "",
                "email_status": "pending",
                "error_msg": None,
            })
            next_id += 1

    db.save_keys(sid, {
        "students": students,
        "next_id": next_id,
        "column_mapping": mapping,
    })

    return jsonify({"success": True, "students_loaded": len(students)})


@app.route("/template")
def template_page():
    return render_template("template_upload.html")


@app.route("/_/upload-template", methods=["POST"])
def upload_template():
    sid = get_sid()
    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400
    file = request.files["file"]
    if not file.filename or not allowed_file(file.filename, ALLOWED_PDF):
        return jsonify({"error": "Only PDF files are allowed"}), 400

    import fitz
    file_bytes = file.read()
    try:
        doc = fitz.open("pdf", file_bytes)
        if len(doc) == 0:
            return jsonify({"error": "PDF has no pages"}), 400
        doc.close()
    except Exception as e:
        return jsonify({"error": f"Invalid PDF: {str(e)}"}), 400

    try:
        img_b64 = pdf_bytes_to_image_base64(file_bytes)
    except Exception as e:
        return jsonify({"error": f"Could not render preview: {str(e)}"}), 500

    db.ensure_session(sid)
    db.save_template_bytes(sid, file_bytes)

    return jsonify({"success": True, "preview": img_b64})


@app.route("/editor")
def editor_page():
    sid = get_sid()
    template_bytes = db.get_template_bytes(sid)
    if template_bytes is None:
        return redirect(url_for("template_page"))

    img_b64 = pdf_bytes_to_image_base64(template_bytes)
    state = db.get_state(sid)
    return render_template("editor.html", preview_b64=img_b64, editor_settings=state["editor_settings"])


@app.route("/_/save-editor-settings", methods=["POST"])
def save_editor_settings():
    sid = get_sid()
    data = request.json or {}
    editor_settings = {
        "positions": data.get("positions", {}),
        "font_family": data.get("font_family", "Helvetica"),
        "font_size": data.get("font_size", 24),
        "text_color": data.get("text_color", "#000000"),
        "bold": data.get("bold", False),
        "italic": data.get("italic", False),
    }
    db.ensure_session(sid)
    db.save_key(sid, "editor_settings", editor_settings)
    return jsonify({"success": True})


@app.route("/preview")
def preview_page():
    sid = get_sid()
    state = db.get_state(sid)
    students = state["students"][:5]
    editor_settings = state["editor_settings"]
    has_template = db.get_template_bytes(sid) is not None
    has_students = len(students) > 0
    has_positions = bool(editor_settings.get("positions"))
    return render_template("preview.html",
        students=students,
        has_template=has_template,
        has_students=has_students,
        has_positions=has_positions)


@app.route("/_/preview-certificate/<int:student_id>")
def preview_certificate(student_id):
    sid = get_sid()
    state = db.get_state(sid)
    student = next((s for s in state["students"] if s["id"] == student_id), None)
    if not student:
        return jsonify({"error": "Student not found"}), 404

    template_bytes = db.get_template_bytes(sid)
    if template_bytes is None:
        return jsonify({"error": "Template not found"}), 400

    editor_settings = state["editor_settings"]
    if not editor_settings.get("positions"):
        return jsonify({"error": "Editor settings not saved"}), 400

    try:
        pdf_bytes = generate_certificate_pdf_bytes(template_bytes, student, editor_settings)
        import fitz
        doc = fitz.open("pdf", pdf_bytes)
        page = doc[0]
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat)
        img_b64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")
        doc.close()
        return jsonify({"success": True, "preview": img_b64, "name": student["name"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/generate")
def generate_page():
    sid = get_sid()
    state = db.get_state(sid)
    total = len(state["students"])
    return render_template("generate.html", total=total)


def _run_generate_all(sid: str, students: list, template_bytes: bytes, editor_settings: dict):
    with _progress_lock:
        _generate_progress[sid] = {
            "total": len(students),
            "done_count": 0,
            "done": False,
            "running": True,
        }

    updated_students = list(students)

    for i, student in enumerate(updated_students):
        safe_name = secure_filename(student["name"].replace(" ", "_")) or "certificate"
        out_path = os.path.join(GENERATED_FOLDER, f"{sid}_{safe_name}_{student['id']}.pdf")
        try:
            pdf_bytes = generate_certificate_pdf_bytes(template_bytes, student, editor_settings)
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)
            updated_students[i] = {**student, "cert_path": out_path}
        except Exception:
            pass

        with _progress_lock:
            _generate_progress[sid]["done_count"] = i + 1

    # Save updated cert_paths back to DB
    db.save_key(sid, "students", updated_students)

    with _progress_lock:
        _generate_progress[sid]["done"] = True
        _generate_progress[sid]["running"] = False


@app.route("/_/generate-all", methods=["POST"])
def generate_all():
    sid = get_sid()
    prog = _generate_progress.get(sid, {})
    if prog.get("running"):
        return jsonify({"error": "Generation already in progress"}), 400

    template_bytes = db.get_template_bytes(sid)
    if template_bytes is None:
        return jsonify({"error": "No template uploaded"}), 400

    state = db.get_state(sid)
    editor_settings = state["editor_settings"]
    if not editor_settings.get("positions"):
        return jsonify({"error": "No editor settings saved"}), 400

    students = state["students"]
    if not students:
        return jsonify({"error": "No student data loaded"}), 400

    t = threading.Thread(
        target=_run_generate_all,
        args=(sid, students, template_bytes, editor_settings),
        daemon=True,
    )
    t.start()
    return jsonify({"success": True})


@app.route("/_/generate-progress")
def generate_progress_stream():
    sid = get_sid()

    def event_stream():
        while True:
            with _progress_lock:
                p = _generate_progress.get(sid, get_default_progress()).copy()
            yield f"data: {json.dumps(p)}\n\n"
            if p.get("done"):
                break
            time.sleep(0.5)

    return Response(stream_with_context(event_stream()),
                    mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/email-settings")
def email_settings_page():
    sid = get_sid()
    state = db.get_state(sid)
    return render_template("email_settings.html",
        settings=state["email_settings"],
        template=state["email_template"])


@app.route("/_/save-email-settings", methods=["POST"])
def save_email_settings():
    sid = get_sid()
    data = request.json or {}
    settings = {
        "smtp_host": data.get("smtp_host", "smtp.gmail.com"),
        "smtp_port": int(data.get("smtp_port", 587)),
        "sender_email": data.get("sender_email", ""),
        "app_password": data.get("app_password", ""),
    }
    db.ensure_session(sid)
    db.save_key(sid, "email_settings", settings)
    return jsonify({"success": True})


@app.route("/_/save-email-template", methods=["POST"])
def save_email_template():
    sid = get_sid()
    data = request.json or {}
    template = {
        "subject": data.get("subject", ""),
        "body": data.get("body", ""),
    }
    db.ensure_session(sid)
    db.save_key(sid, "email_template", template)
    return jsonify({"success": True})


@app.route("/_/test-smtp", methods=["POST"])
def test_smtp():
    sid = get_sid()
    state = db.get_state(sid)
    settings = state["email_settings"]
    if not settings.get("sender_email") or not settings.get("app_password"):
        return jsonify({"error": "Email settings not configured"}), 400
    try:
        server = smtplib.SMTP(settings["smtp_host"], settings["smtp_port"], timeout=10)
        server.starttls()
        server.login(settings["sender_email"], settings["app_password"])
        server.quit()
        return jsonify({"success": True, "message": "SMTP connection successful"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/send-emails")
def send_emails_page():
    sid = get_sid()
    state = db.get_state(sid)
    students = state["students"]
    total = len(students)
    pending = sum(1 for s in students if s["email_status"] == "pending")
    sent = sum(1 for s in students if s["email_status"] == "sent")
    failed = sum(1 for s in students if s["email_status"] == "failed")
    return render_template("email_send.html",
        total=total, pending=pending, sent=sent, failed=failed,
        email_settings=state["email_settings"],
        email_template=state["email_template"])


def _send_single_email(smtp_settings: dict, to_email: str, subject: str, body: str, attachment_path: str):
    msg = MIMEMultipart()
    msg["From"] = smtp_settings["sender_email"]
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        fname = os.path.basename(attachment_path)
        part.add_header("Content-Disposition", f"attachment; filename={fname}")
        msg.attach(part)

    server = smtplib.SMTP(smtp_settings["smtp_host"], smtp_settings["smtp_port"], timeout=15)
    server.starttls()
    server.login(smtp_settings["sender_email"], smtp_settings["app_password"])
    server.sendmail(smtp_settings["sender_email"], to_email, msg.as_string())
    server.quit()


def _run_send_emails(sid: str, student_ids=None):
    state = db.get_state(sid)
    smtp_settings = state["email_settings"]
    email_template_data = state["email_template"]
    students = state["students"]

    if student_ids:
        targets = [s for s in students if s["id"] in student_ids]
    else:
        targets = [s for s in students if s.get("email_status") != "sent"]

    with _progress_lock:
        _email_progress[sid] = {
            "total": len(targets),
            "sent": 0,
            "failed": 0,
            "current": "",
            "done": False,
            "running": True,
        }

    for s in targets:
        with _progress_lock:
            _email_progress[sid]["current"] = s.get("name", "")

        subject = email_template_data.get("subject", "").format(
            name=s.get("name", ""), roll=s.get("roll", ""), college=s.get("college", "")
        )
        body = email_template_data.get("body", "").format(
            name=s.get("name", ""), roll=s.get("roll", ""), college=s.get("college", "")
        )

        try:
            _send_single_email(smtp_settings, s["email"], subject, body, s.get("cert_path", ""))
            s["email_status"] = "sent"
            s["error_msg"] = None
            with _progress_lock:
                _email_progress[sid]["sent"] += 1
        except Exception as e:
            s["email_status"] = "failed"
            s["error_msg"] = str(e)
            with _progress_lock:
                _email_progress[sid]["failed"] += 1

        time.sleep(0.1)

    # Save updated email statuses back to DB
    db.save_key(sid, "students", students)

    with _progress_lock:
        _email_progress[sid]["done"] = True
        _email_progress[sid]["running"] = False
        _email_progress[sid]["current"] = ""


@app.route("/_/send-emails", methods=["POST"])
def start_send_emails():
    sid = get_sid()
    prog = _email_progress.get(sid, {})
    if prog.get("running"):
        return jsonify({"error": "Email sending already in progress"}), 400

    state = db.get_state(sid)
    settings = state["email_settings"]
    if not settings.get("sender_email") or not settings.get("app_password"):
        return jsonify({"error": "Email settings not configured"}), 400

    t = threading.Thread(target=_run_send_emails, args=(sid,), daemon=True)
    t.start()
    return jsonify({"success": True})


@app.route("/_/retry-failed", methods=["POST"])
def retry_failed():
    sid = get_sid()
    prog = _email_progress.get(sid, {})
    if prog.get("running"):
        return jsonify({"error": "Email operation in progress"}), 400

    state = db.get_state(sid)
    students = state["students"]
    failed_ids = [s["id"] for s in students if s.get("email_status") == "failed"]
    if not failed_ids:
        return jsonify({"error": "No failed emails to retry"}), 400

    for s in students:
        if s.get("email_status") == "failed":
            s["email_status"] = "pending"
    db.save_key(sid, "students", students)

    t = threading.Thread(target=_run_send_emails, args=(sid, failed_ids), daemon=True)
    t.start()
    return jsonify({"success": True, "retrying": len(failed_ids)})


@app.route("/_/email-progress")
def email_progress_stream():
    sid = get_sid()

    def event_stream():
        while True:
            with _progress_lock:
                p = _email_progress.get(sid, get_default_email_progress()).copy()
            yield f"data: {json.dumps(p)}\n\n"
            if p.get("done"):
                break
            time.sleep(0.5)

    return Response(stream_with_context(event_stream()),
                    mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/_/email-status")
def email_status():
    sid = get_sid()
    with _progress_lock:
        p = _email_progress.get(sid, get_default_email_progress()).copy()
    return jsonify(p)


@app.route("/_/stats")
def stats():
    sid = get_sid()
    state = db.get_state(sid)
    students = state["students"]
    total = len(students)
    generated = sum(1 for s in students if s.get("cert_path"))
    sent = sum(1 for s in students if s.get("email_status") == "sent")
    failed = sum(1 for s in students if s.get("email_status") == "failed")
    pending = sum(1 for s in students if s.get("email_status") == "pending")
    return jsonify({"total": total, "generated": generated, "sent": sent, "failed": failed, "pending": pending})


@app.route("/_/students")
def list_students():
    sid = get_sid()
    state = db.get_state(sid)
    return jsonify([
        {
            "id": s["id"], "name": s["name"], "roll": s["roll"], "college": s["college"],
            "email": s["email"], "email_status": s["email_status"], "error_msg": s["error_msg"],
        }
        for s in state["students"]
    ])


@app.route("/_/reset", methods=["POST"])
def reset_all():
    sid = get_sid()
    db.delete_session(sid)

    # Clean up any generated PDFs for this session
    for fn in os.listdir(GENERATED_FOLDER):
        if fn.startswith(sid):
            try:
                os.remove(os.path.join(GENERATED_FOLDER, fn))
            except Exception:
                pass

    with _progress_lock:
        _generate_progress.pop(sid, None)
        _email_progress.pop(sid, None)

    return jsonify({"success": True})


@app.route("/generated/<path:filename>")
def download_generated(filename):
    return send_from_directory(GENERATED_FOLDER, filename)


# ── startup ───────────────────────────────────────────────────────────────────

def create_app():
    db.init_db()
    return app


if __name__ == "__main__":
    db.init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

import os
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
    Response, stream_with_context, send_from_directory
)
import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "cert-app-secret-2024")

# Temp folder only used to hold generated PDFs long enough to email them.
# Nothing here needs to survive a restart/redeploy.
GENERATED_FOLDER = tempfile.mkdtemp(prefix="certgen_")

ALLOWED_EXCEL = {"xlsx", "xls", "csv"}
ALLOWED_PDF = {"pdf"}

# ---------------------------------------------------------------------------
# In-memory application state. No database, nothing written to disk except
# the temporary generated-certificate PDFs used right before emailing.
# This state resets whenever the app restarts.
# ---------------------------------------------------------------------------
STATE = {
    "students": [],          # list of dicts: id, name, roll, college, email, cert_path, email_status, error_msg
    "next_id": 1,
    "excel_records": [],     # parsed rows from the uploaded sheet, used for column mapping
    "excel_columns": [],
    "column_mapping": {},
    "template_bytes": None,  # uploaded certificate template PDF (raw bytes)
    "editor_settings": {},   # positions, font_family, font_size, text_color, bold, italic
    "email_settings": {
        "smtp_host": os.environ.get("SMTP_HOST", "smtp.gmail.com"),
        "smtp_port": int(os.environ.get("SMTP_PORT", "587")),
        "sender_email": os.environ.get("SENDER_EMAIL", ""),
        "app_password": os.environ.get("APP_PASSWORD", ""),
    },
    "email_template": {
        "subject": "Certificate of Participation",
        "body": "Dear {name},\n\nThank you for participating.\n\nYour certificate is attached.\n\nRegards,\nEvent Team",
    },
}
state_lock = threading.Lock()

email_progress = {"total": 0, "sent": 0, "failed": 0, "current": "", "done": False, "running": False}
generate_progress = {"total": 0, "done_count": 0, "done": False, "running": False}


def allowed_file(filename, allowed):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed


def pdf_bytes_to_image_base64(pdf_bytes):
    import fitz
    doc = fitz.open("pdf", pdf_bytes)
    page = doc[0]
    mat = fitz.Matrix(2.0, 2.0)
    pix = page.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    doc.close()
    return base64.b64encode(img_bytes).decode("utf-8")


def generate_certificate_pdf_bytes(template_bytes, student, editor_settings):
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


@app.route("/")
def dashboard():
    students = STATE["students"]
    total = len(students)
    generated = sum(1 for s in students if s.get("cert_path"))
    sent = sum(1 for s in students if s.get("email_status") == "sent")
    failed = sum(1 for s in students if s.get("email_status") == "failed")
    pending = sum(1 for s in students if s.get("email_status") == "pending")

    template_ready = STATE["template_bytes"] is not None
    return render_template("index.html",
        total=total, generated=generated, sent=sent, failed=failed, pending=pending,
        template_ready=template_ready, session_data=STATE["editor_settings"])


@app.route("/upload")
def upload_page():
    return render_template("upload.html")


@app.route("/_/upload-excel", methods=["POST"])
def upload_excel():
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

        with state_lock:
            STATE["excel_records"] = records
            STATE["excel_columns"] = columns

        return jsonify({"columns": columns, "total_rows": total_rows, "sample": sample})
    except Exception as e:
        return jsonify({"error": f"Failed to read file: {str(e)}"}), 500


@app.route("/_/save-mapping", methods=["POST"])
def save_mapping():
    data = request.json or {}
    mapping = data.get("mapping", {})
    required = ["name", "roll", "college", "email"]
    for k in required:
        if k not in mapping or not mapping[k]:
            return jsonify({"error": f"Missing mapping for {k}"}), 400

    records = STATE.get("excel_records") or []
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

    with state_lock:
        STATE["students"] = students
        STATE["next_id"] = next_id
        STATE["column_mapping"] = mapping

    return jsonify({"success": True, "students_loaded": len(students)})


@app.route("/template")
def template_page():
    return render_template("template_upload.html")


@app.route("/_/upload-template", methods=["POST"])
def upload_template():
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

    with state_lock:
        STATE["template_bytes"] = file_bytes

    return jsonify({"success": True, "preview": img_b64})


@app.route("/editor")
def editor_page():
    if STATE["template_bytes"] is None:
        return redirect(url_for("template_page"))

    img_b64 = pdf_bytes_to_image_base64(STATE["template_bytes"])
    return render_template("editor.html", preview_b64=img_b64, editor_settings=STATE["editor_settings"])


@app.route("/_/save-editor-settings", methods=["POST"])
def save_editor_settings():
    data = request.json or {}
    with state_lock:
        STATE["editor_settings"] = {
            "positions": data.get("positions", {}),
            "font_family": data.get("font_family", "Helvetica"),
            "font_size": data.get("font_size", 24),
            "text_color": data.get("text_color", "#000000"),
            "bold": data.get("bold", False),
            "italic": data.get("italic", False),
        }
    return jsonify({"success": True})


@app.route("/preview")
def preview_page():
    students = STATE["students"][:5]
    editor_settings = STATE["editor_settings"]
    has_template = STATE["template_bytes"] is not None
    has_students = len(students) > 0
    has_positions = bool(editor_settings.get("positions"))
    return render_template("preview.html",
        students=students,
        has_template=has_template,
        has_students=has_students,
        has_positions=has_positions)


@app.route("/_/preview-certificate/<int:student_id>")
def preview_certificate(student_id):
    student = next((s for s in STATE["students"] if s["id"] == student_id), None)
    if not student:
        return jsonify({"error": "Student not found"}), 404

    if STATE["template_bytes"] is None:
        return jsonify({"error": "Template not found"}), 400

    editor_settings = STATE["editor_settings"]
    if not editor_settings.get("positions"):
        return jsonify({"error": "Editor settings not saved"}), 400

    try:
        pdf_bytes = generate_certificate_pdf_bytes(STATE["template_bytes"], student, editor_settings)
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
    total = len(STATE["students"])
    return render_template("generate.html", total=total)


def run_generate_all():
    global generate_progress
    students = STATE["students"]
    template_bytes = STATE["template_bytes"]
    editor_settings = STATE["editor_settings"]

    generate_progress["total"] = len(students)
    generate_progress["done_count"] = 0
    generate_progress["done"] = False
    generate_progress["running"] = True

    for student in students:
        safe_name = secure_filename(student["name"].replace(" ", "_")) or "certificate"
        out_path = os.path.join(GENERATED_FOLDER, f"Certificate_{safe_name}_{student['id']}.pdf")
        try:
            pdf_bytes = generate_certificate_pdf_bytes(template_bytes, student, editor_settings)
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)
            student["cert_path"] = out_path
        except Exception:
            pass
        generate_progress["done_count"] += 1

    generate_progress["done"] = True
    generate_progress["running"] = False


@app.route("/_/generate-all", methods=["POST"])
def generate_all():
    global generate_progress
    if generate_progress.get("running"):
        return jsonify({"error": "Generation already in progress"}), 400

    if STATE["template_bytes"] is None:
        return jsonify({"error": "No template uploaded"}), 400

    if not STATE["editor_settings"].get("positions"):
        return jsonify({"error": "No editor settings saved"}), 400

    if len(STATE["students"]) == 0:
        return jsonify({"error": "No student data loaded"}), 400

    t = threading.Thread(target=run_generate_all)
    t.daemon = True
    t.start()
    return jsonify({"success": True})


@app.route("/_/generate-progress")
def generate_progress_stream():
    def event_stream():
        while True:
            p = generate_progress.copy()
            data = json.dumps(p)
            yield f"data: {data}\n\n"
            if p.get("done"):
                break
            time.sleep(0.5)
    return Response(stream_with_context(event_stream()),
                    mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/email-settings")
def email_settings_page():
    return render_template("email_settings.html", settings=STATE["email_settings"], template=STATE["email_template"])


@app.route("/_/save-email-settings", methods=["POST"])
def save_email_settings():
    data = request.json or {}
    with state_lock:
        STATE["email_settings"] = {
            "smtp_host": data.get("smtp_host", "smtp.gmail.com"),
            "smtp_port": int(data.get("smtp_port", 587)),
            "sender_email": data.get("sender_email", ""),
            "app_password": data.get("app_password", ""),
        }
    return jsonify({"success": True})


@app.route("/_/save-email-template", methods=["POST"])
def save_email_template():
    data = request.json or {}
    with state_lock:
        STATE["email_template"] = {
            "subject": data.get("subject", ""),
            "body": data.get("body", ""),
        }
    return jsonify({"success": True})


@app.route("/_/test-smtp", methods=["POST"])
def test_smtp():
    settings = STATE["email_settings"]
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
    students = STATE["students"]
    total = len(students)
    pending = sum(1 for s in students if s["email_status"] == "pending")
    sent = sum(1 for s in students if s["email_status"] == "sent")
    failed = sum(1 for s in students if s["email_status"] == "failed")
    return render_template("email_send.html",
        total=total, pending=pending, sent=sent, failed=failed,
        email_settings=STATE["email_settings"], email_template=STATE["email_template"])


def send_single_email(smtp_settings, to_email, subject, body, attachment_path):
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


def run_send_emails(student_ids=None):
    global email_progress

    smtp_settings = STATE["email_settings"]
    email_template_data = STATE["email_template"]

    if student_ids:
        students = [s for s in STATE["students"] if s["id"] in student_ids]
    else:
        students = [s for s in STATE["students"] if s["email_status"] != "sent"]

    email_progress["total"] = len(students)
    email_progress["sent"] = 0
    email_progress["failed"] = 0
    email_progress["done"] = False
    email_progress["running"] = True
    email_progress["current"] = ""

    for s in students:
        email_progress["current"] = s.get("name", "")

        subject = email_template_data.get("subject", "").format(
            name=s.get("name", ""), roll=s.get("roll", ""), college=s.get("college", "")
        )
        body = email_template_data.get("body", "").format(
            name=s.get("name", ""), roll=s.get("roll", ""), college=s.get("college", "")
        )

        try:
            send_single_email(smtp_settings, s["email"], subject, body, s.get("cert_path", ""))
            s["email_status"] = "sent"
            s["error_msg"] = None
            email_progress["sent"] += 1
        except Exception as e:
            s["email_status"] = "failed"
            s["error_msg"] = str(e)
            email_progress["failed"] += 1

        time.sleep(0.1)

    email_progress["done"] = True
    email_progress["running"] = False
    email_progress["current"] = ""


@app.route("/_/send-emails", methods=["POST"])
def start_send_emails():
    global email_progress
    if email_progress.get("running"):
        return jsonify({"error": "Email sending already in progress"}), 400

    smtp_settings = STATE["email_settings"]
    if not smtp_settings.get("sender_email") or not smtp_settings.get("app_password"):
        return jsonify({"error": "Email settings not configured"}), 400

    t = threading.Thread(target=run_send_emails)
    t.daemon = True
    t.start()
    return jsonify({"success": True})


@app.route("/_/retry-failed", methods=["POST"])
def retry_failed():
    global email_progress
    if email_progress.get("running"):
        return jsonify({"error": "Email operation in progress"}), 400

    failed_ids = [s["id"] for s in STATE["students"] if s["email_status"] == "failed"]
    if not failed_ids:
        return jsonify({"error": "No failed emails to retry"}), 400

    for s in STATE["students"]:
        if s["email_status"] == "failed":
            s["email_status"] = "pending"

    t = threading.Thread(target=run_send_emails, args=(failed_ids,))
    t.daemon = True
    t.start()
    return jsonify({"success": True, "retrying": len(failed_ids)})


@app.route("/_/email-progress")
def email_progress_stream():
    def event_stream():
        while True:
            p = email_progress.copy()
            data = json.dumps(p)
            yield f"data: {data}\n\n"
            if p.get("done"):
                break
            time.sleep(0.5)
    return Response(stream_with_context(event_stream()),
                    mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/_/email-status")
def email_status():
    return jsonify(email_progress.copy())


@app.route("/_/stats")
def stats():
    students = STATE["students"]
    total = len(students)
    generated = sum(1 for s in students if s.get("cert_path"))
    sent = sum(1 for s in students if s["email_status"] == "sent")
    failed = sum(1 for s in students if s["email_status"] == "failed")
    pending = sum(1 for s in students if s["email_status"] == "pending")
    return jsonify({"total": total, "generated": generated, "sent": sent, "failed": failed, "pending": pending})


@app.route("/_/students")
def list_students():
    return jsonify([
        {
            "id": s["id"], "name": s["name"], "roll": s["roll"], "college": s["college"],
            "email": s["email"], "email_status": s["email_status"], "error_msg": s["error_msg"],
        }
        for s in STATE["students"]
    ])


@app.route("/_/reset", methods=["POST"])
def reset_all():
    with state_lock:
        STATE["students"] = []
        STATE["next_id"] = 1
        STATE["excel_records"] = []
        STATE["excel_columns"] = []
        STATE["column_mapping"] = {}
        STATE["template_bytes"] = None
        STATE["editor_settings"] = {}

    for fn in os.listdir(GENERATED_FOLDER):
        fp = os.path.join(GENERATED_FOLDER, fn)
        try:
            os.remove(fp)
        except Exception:
            pass

    return jsonify({"success": True})


@app.route("/generated/<path:filename>")
def download_generated(filename):
    return send_from_directory(GENERATED_FOLDER, filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

import io
import json
import os
import tempfile
from typing import List

from flask import Flask, request, jsonify
from flask_cors import CORS

# Outlook COM
try:
    import win32com.client  # type: ignore
    win32_client = win32com.client
except Exception as e:
    win32_client = None
try:
    import pythoncom  # type: ignore
except Exception:
    pythoncom = None

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": ["http://127.0.0.1", "http://localhost", "*"]}})


def open_outlook_drafts(subject: str, recipients: List[str], body: str, attachment_path: str) -> int:
    if win32_client is None:
        raise RuntimeError("pywin32 is not available. Install requirements on Windows.")

    # Initialize COM for this thread (Flask handles requests in threads)
    initialized = False
    try:
        if pythoncom is not None:
            pythoncom.CoInitialize()
            initialized = True

        outlook = win32_client.Dispatch("Outlook.Application")
        created = 0
        for rcpt in recipients:
            rcpt = rcpt.strip()
            if not rcpt:
                continue
            mail = outlook.CreateItem(0)  # olMailItem
            mail.To = rcpt
            mail.Subject = subject
            # Prefer HTML body to preserve formatting; fallback to plain text if needed
            try:
                mail.HTMLBody = body
            except Exception:
                mail.Body = body
            if attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(Source=attachment_path)
            # Display opens as draft window (does not send)
            mail.Display(False)
            created += 1
        return created
    finally:
        if initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})


@app.route('/draft', methods=['POST'])
def draft():
    try:
        subject = request.form.get('subject', '').strip()
        body = request.form.get('body', '')
        recipients_raw = request.form.get('recipients', '[]')
        if not subject or not body:
            return jsonify({"message": "subject and body are required"}), 400
        try:
            recipients = json.loads(recipients_raw)
            if not isinstance(recipients, list):
                raise ValueError()
        except Exception:
            # Also try to split a raw string fallback
            recipients = [s.strip() for s in recipients_raw.split(',') if s.strip()]
        recipients = [r for r in recipients if r]
        if not recipients:
            return jsonify({"message": "no recipients provided"}), 400

        # Handle file upload
        file = request.files.get('attachment')
        if not file:
            return jsonify({"message": "attachment is required"}), 400
        # Save to a secure temp file
        suffix = os.path.splitext(file.filename or '')[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            file.save(tmp)
            tmp_path = tmp.name

        try:
            created = open_outlook_drafts(subject, recipients, body, tmp_path)
            return jsonify({"created": created, "recipients": recipients})
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
    except Exception as e:
        return jsonify({"message": str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '5005'))
    app.run(host='127.0.0.1', port=port)

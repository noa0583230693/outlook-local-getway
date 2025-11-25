import base64
import os
import tempfile
import pythoncom
from flask import Flask, request, jsonify
from flask_cors import CORS

LISTEN_HOST = "127.0.0.1"
LISTEN_PORT = 5001
SHARED_SECRET = "secure_secret_here"

app = Flask(__name__)
CORS(app)


@app.get("/ping")
def ping():
    return {"status": "ok"}


def save_file(filename, b64data):
    suffix = ""
    if "." in filename:
        suffix = "." + filename.split(".")[-1]

    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)

    with open(path, "wb") as f:
        f.write(base64.b64decode(b64data))

    return path


@app.post("/open-mail")
def open_mail():
    try:
        data = request.get_json(force=True)

        # ----------------------
        # Validation Layer
        # ----------------------
        if data.get("secret") != SHARED_SECRET:
            return jsonify({"ok": False, "error": "Unauthorized"}), 401

        subject = data.get("subject", "").strip()
        body = data.get("body", "").strip()
        recipients = data.get("recipients", [])

        if not recipients:
            return jsonify({"ok": False, "error": "Missing recipients"}), 400

        if not subject:
            return jsonify({"ok": False, "error": "Missing subject"}), 400

        # ----------------------
        # Save attachment if exists
        # ----------------------
        attachment_path = None
        if data.get("attachment_base64") and data.get("attachment_name"):
            attachment_path = save_file(
                data["attachment_name"],
                data["attachment_base64"]
            )

        # ----------------------
        # Outlook Initialization
        # ----------------------
        pythoncom.CoInitialize()
        import win32com.client

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception:
            pythoncom.CoUninitialize()
            return jsonify({"ok": False, "error": "Outlook is not available"}), 500

        # ----------------------
        # Create mails
        # ----------------------
        for r in recipients:
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.HTMLBody = body
            mail.To = r

            if attachment_path:
                mail.Attachments.Add(attachment_path)

            mail.Display(False)

        pythoncom.CoUninitialize()

        return jsonify({"ok": True})

    except Exception as ex:
        return jsonify({"ok": False, "error": str(ex)}), 500


if __name__ == "__main__":
    print(f"Agent running at http://{LISTEN_HOST}:{LISTEN_PORT}")
    app.run(host=LISTEN_HOST, port=LISTEN_PORT)

import os, json
import requests
from flask import Flask, request, jsonify, send_from_directory

# Load .env in local dev (Railway injects env vars directly in prod)
env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(env_path):
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

ONYX_API_KEY  = os.environ.get("ONYX_API_KEY", "")
ONYX_ENDPOINT = os.environ.get("ONYX_ENDPOINT", "")

app = Flask(__name__, static_folder=".", static_url_path="")

# ── Static files ──────────────────────────────────────────────────────────────

ALLOWED_STATIC = {"index.html", "styles.css", "app.js"}

@app.route("/")
def index():
    return send_from_directory(".", "index.html")

@app.route("/<path:filename>")
def static_files(filename):
    if filename not in ALLOWED_STATIC:
        return "", 404
    return send_from_directory(".", filename)


# ── Lead proxy ────────────────────────────────────────────────────────────────

def to_e164(raw: str) -> str | None:
    digits = "".join(c for c in raw if c.isdigit())
    if len(digits) == 10:
        digits = "1" + digits
    if len(digits) == 11 and digits.startswith("1"):
        return "+" + digits
    return None


@app.route("/submit-lead", methods=["POST"])
def submit_lead():
    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify(error="Invalid JSON"), 400

    for field in ("zip", "first", "last", "phone"):
        if not str(data.get(field, "")).strip():
            return jsonify(error=f"Missing field: {field}"), 400

    phone = to_e164(data["phone"])
    if not phone:
        return jsonify(error="Invalid phone number"), 400

    payload = {
        "leads": [{
            "phone":      phone,
            "first_name": data["first"].strip(),
            "last_name":  data["last"].strip(),
            "zip_code":   data["zip"].strip(),
            "additional_context": f"Source: {data.get('source', 'medicareplancheck.org')}",
        }]
    }

    try:
        resp = requests.post(
            ONYX_ENDPOINT,
            json=payload,
            headers={"X-API-Key": ONYX_API_KEY, "Content-Type": "application/json"},
            timeout=10,
        )
        result = resp.json()
        accepted = result.get("total_accepted", "?")
        rejected = result.get("total_rejected", 0)
        app.logger.info("Onyx accepted=%s rejected=%s phone=%s", accepted, rejected, phone)

        if rejected:
            reason = (result.get("results") or [{}])[0].get("reason", "unknown")
            app.logger.warning("Onyx rejected lead: %s", reason)

        return jsonify(ok=True, accepted=accepted, rejected=rejected), 200

    except requests.exceptions.Timeout:
        app.logger.error("Onyx request timed out for %s", phone)
        return jsonify(ok=True, note="CRM timeout — lead queued for retry"), 200
    except Exception as e:
        app.logger.error("Onyx submission error: %s", e)
        return jsonify(error=str(e)), 502


# ── Entry point (local dev only) ──────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    app.run(host="0.0.0.0", port=port, debug=False)

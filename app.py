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

# ── State advertorial network (Host-based routing) ──────────────────────────────
# Every state domain points (DNS) to this same app; we dispatch on the Host header.
# Add a state = one line here + drop its HTML in state-sites/.
STATE_SITES = {
    "texasplancheck.org":         "tx-extra-benefits-advertorial.html",
    "northcarolinaplancheck.org": "nc-extra-benefits-advertorial.html",
    "floridaplancheck.org":       "fl-extra-benefits-advertorial.html",
}

def _advertorial_for_host():
    """Return the advertorial filename if the request Host is a state site, else None."""
    host = request.host.split(":")[0].lower()
    if host.startswith("www."):
        host = host[4:]
    return STATE_SITES.get(host)

# ── Static files ──────────────────────────────────────────────────────────────

ALLOWED_STATIC = {"index.html", "thank-you.html", "privacy-policy.html", "styles.css", "app.js", "robots.txt", "favicon.ico"}

@app.route("/")
def index():
    advertorial = _advertorial_for_host()
    if advertorial:
        return send_from_directory("state-sites", advertorial)
    return send_from_directory(".", "index.html")

@app.route("/hero-benefits-review.jpg")
def advertorial_hero():
    return send_from_directory("state-sites", "hero-benefits-review.jpg", mimetype="image/jpeg")

@app.route("/thank-you")
def thank_you():
    return send_from_directory(".", "thank-you.html")

@app.route("/privacy-policy")
def privacy_policy():
    return send_from_directory(".", "privacy-policy.html")

@app.route("/robots.txt")
def robots():
    return send_from_directory(".", "robots.txt", mimetype="text/plain")

@app.route("/favicon.ico")
def favicon():
    return send_from_directory(".", "favicon.ico", mimetype="image/x-icon")

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
            "cost_dollars": 0,
            "additional_context": " | ".join(filter(None, [
                f"Source: {data.get('source', 'medicareplancheck.org')}",
                f"Medium: {data['utm_medium']}"   if data.get('utm_medium')   else "",
                f"Campaign: {data['utm_campaign']}" if data.get('utm_campaign') else "",
                f"Creative: {data['utm_content']}"  if data.get('utm_content')  else "",
            ])),
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
        accepted = result.get("total_accepted", 0)
        rejected = result.get("total_rejected", 0)
        lead_result = (result.get("results") or [{}])[0]
        lead_id  = lead_result.get("lead_id")
        status   = lead_result.get("status", "unknown")
        reason   = lead_result.get("reason")

        app.logger.info("Onyx status=%s accepted=%s rejected=%s lead_id=%s phone=%s",
                        status, accepted, rejected, lead_id, phone)

        if rejected or status == "rejected":
            app.logger.warning("Onyx rejected lead: %s", reason or "no reason given")

        return jsonify(ok=True, accepted=accepted, rejected=rejected, lead_id=lead_id), 200

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

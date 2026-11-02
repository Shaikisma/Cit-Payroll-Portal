import hmac
import json
import os
import secrets
import shutil
import uuid
from datetime import datetime
from pathlib import Path

from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_file, session, url_for)

import processor
from processor import PortalError

BASE = Path(__file__).parent
DATA = Path(os.environ.get("DATA_DIR", BASE / "data"))
RUNS = DATA / "runs"
TMP = DATA / "tmp"
for d in (RUNS, TMP):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-change-me")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
HR_PASSWORD = os.environ.get("HR_PASSWORD", "hr-change-me")  # set a real one on Render
SLOTS = {"payroll": "Previous month payroll", "invoice": "Invoices workbook", "health": "Health insurance (carrier)"}


def csrf():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return session["csrf"]


app.jinja_env.globals["csrf"] = csrf


@app.before_request
def guard():
    if request.endpoint in ("login", "static", "healthz"):
        return
    if not session.get("ok"):
        return redirect(url_for("login"))
    if request.method == "POST" and not hmac.compare_digest(request.form.get("csrf", ""), session.get("csrf", "x")):
        abort(400)


@app.get("/healthz")
def healthz():
    return "ok"


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if hmac.compare_digest(request.form.get("password", ""), HR_PASSWORD):
            session.clear()
            session["ok"] = True
            return redirect(url_for("index"))
        flash("Wrong password.", "error")
    return render_template("login.html")


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.get("/")
def index():
    return render_template("index.html", slots=SLOTS)


@app.post("/generate")
def generate():
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:4]
    work = TMP / run_id
    work.mkdir()
    paths = {}
    out_dir = RUNS / run_id
    ok = False
    try:
        for slot, label in SLOTS.items():
            f = request.files.get(slot)
            if not f or not f.filename:
                raise PortalError(f"Please choose the file for: {label}.")
            if not f.filename.lower().endswith(".xlsx"):
                raise PortalError(f"'{f.filename}' is not an .xlsx file ({label}).")
            p = work / f"{slot}.xlsx"
            f.save(p)
            kind = processor.classify(p)
            if kind != slot:
                raise PortalError(f"'{f.filename}' doesn't look like the {label} file"
                                  + (f" (it looks like the {SLOTS[kind]} file)." if kind else "."))
            paths[slot] = p
        out_dir = RUNS / run_id
        out_dir.mkdir()
        out = out_dir / "payroll.xlsx"
        cfg = json.loads((BASE / "config.json").read_text())
        result = processor.generate(paths["payroll"], paths["invoice"], paths["health"], out,
                                    request.form.get("pay_month") or None,
                                    request.form.get("run_date") or None, cfg)
        result["id"] = run_id
        result["created"] = datetime.now().strftime("%d %b %Y %H:%M")
        result["filename"] = f"CIT_payroll_{result['sheet_name'].replace(' payroll', '')}-{result['run_date'][:4]}.xlsx"
        (out_dir / "meta.json").write_text(json.dumps(result))
        ok = True
        return redirect(url_for("run", run_id=run_id))
    except PortalError as e:
        flash(str(e), "error")
    except Exception as e:  # noqa
        app.logger.exception("generate failed")
        flash(f"Something unexpected went wrong while reading the files ({type(e).__name__}). "
              "Check that the files are the normal monthly formats.", "error")
    finally:
        if not ok:
            shutil.rmtree(out_dir, ignore_errors=True)
        for p in work.glob("*"):   # uploaded inputs are never kept
            p.unlink(missing_ok=True)
        work.rmdir()
    return redirect(url_for("index"))


def load_meta(run_id):
    if not run_id.replace("-", "").isalnum():
        abort(404)
    f = RUNS / run_id / "meta.json"
    if not f.exists():
        abort(404)
    return json.loads(f.read_text())


@app.get("/run/<run_id>")
def run(run_id):
    return render_template("result.html", r=load_meta(run_id))


@app.get("/download/<run_id>")
def download(run_id):
    m = load_meta(run_id)
    return send_file(RUNS / run_id / "payroll.xlsx", as_attachment=True, download_name=m["filename"])


@app.get("/history")
def history():
    items = []
    for d in sorted(RUNS.iterdir(), reverse=True):
        if (d / "meta.json").exists():
            items.append(json.loads((d / "meta.json").read_text()))
    return render_template("history.html", items=items)


if __name__ == "__main__":
    app.run(debug=True, port=5000)

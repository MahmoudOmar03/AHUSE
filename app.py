"""
Flask web application for the AUHSE HSE inspection platform.
Complete redesign with auth, dashboard, history, and multi-page support.
"""

import json
import os
import secrets
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from dotenv import load_dotenv

# Load .env before importing pipeline (pipeline → LLM_VLM reads env at import).
load_dotenv(Path(__file__).resolve().parent / ".env")

import requests
from authlib.integrations.flask_client import OAuth
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_from_directory,
    url_for,
)
from flask_login import (
    LoginManager,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

from db_schema import ensure_schema
from models import ContactMessage, HazardDetection, Inspection, Project, Report, User, db
from pipeline import (
    VIDEO_YOLO_JSON,
    append_video_summary_to_folder,
    process_hse_request,
    process_video_only_request,
)


# ── Upload limits (enforced in addition to MAX_CONTENT_LENGTH) ─────────
MAX_INSPECTION_IMAGES = 20
MAX_IMAGE_BATCH_BYTES = 100 * 1024 * 1024
MAX_VIDEO_BYTES = 250 * 1024 * 1024
ALLOWED_IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp"}
ALLOWED_VIDEO_EXT = {".mp4", ".webm", ".mov", ".avi", ".mkv"}
ALLOWED_IMAGE_MIMES = frozenset(
    {"image/jpeg", "image/png", "image/webp", "image/jpg"},
)


def _google_oauth_configured() -> bool:
    return all(
        os.environ.get(k, "").strip()
        for k in (
            "GOOGLE_OAUTH_CLIENT_ID",
            "GOOGLE_OAUTH_CLIENT_SECRET",
            "GOOGLE_OAUTH_REDIRECT_URI",
        )
    )


def _ext_ok(filename: str, allowed: set) -> bool:
    return Path(filename or "").suffix.lower() in allowed


def _validate_image_files(files: List[Any]) -> Tuple[bool, str, int]:
    """Returns (ok, error_message, total_bytes)."""
    total = 0
    for f in files:
        if not f or not f.filename:
            continue
        if not _ext_ok(f.filename, ALLOWED_IMAGE_EXT):
            return False, f"Unsupported image type: {f.filename}. Use JPEG, PNG, or WebP.", 0
        mime = (f.mimetype or "").split(";")[0].strip().lower()
        if mime and mime not in ALLOWED_IMAGE_MIMES and not mime.startswith("image/"):
            return False, f"Invalid file type for {f.filename}.", 0
        f.seek(0, 2)
        size = f.tell()
        f.seek(0)
        total += size
    if total > MAX_IMAGE_BATCH_BYTES:
        return (
            False,
            f"Total image size exceeds {MAX_IMAGE_BATCH_BYTES // (1024 * 1024)} MB limit.",
            total,
        )
    return True, "", total


def _validate_video_file(f: Any) -> Tuple[bool, str, int]:
    if not f or not f.filename:
        return False, "", 0
    if not _ext_ok(f.filename, ALLOWED_VIDEO_EXT):
        return False, f"Unsupported video format: {f.filename}. Use MP4, WebM, MOV, AVI, or MKV.", 0
    f.seek(0, 2)
    size = f.tell()
    f.seek(0)
    if size > MAX_VIDEO_BYTES:
        return False, f"Video exceeds {MAX_VIDEO_BYTES // (1024 * 1024)} MB limit.", size
    return True, "", size


def create_app(config: Optional[Dict[str, str]] = None) -> Flask:
    base_dir = Path(__file__).resolve().parent
    load_dotenv(base_dir / ".env")

    app = Flask(
        __name__,
        root_path=str(base_dir),
        template_folder="templates",
        static_folder="static",
    )

    default_max_body = 350 * 1024 * 1024
    app.config.update(
        SECRET_KEY=os.environ.get("FLASK_SECRET_KEY", "dev-secret-key-change-in-prod"),
        UPLOAD_FOLDER=str(base_dir / "uploads"),
        OUTPUT_FOLDER=str(base_dir / "outputs"),
        YOLO_MODEL_PATH=os.environ.get("YOLO_MODEL_PATH", str(base_dir / "best.pt")),
        YOLO_CONFIDENCE=float(os.environ.get("YOLO_CONFIDENCE", "0.25")),
        YOLO_USE_LOCAL=os.environ.get("YOLO_USE_LOCAL", ""),
        MAX_RECORDS=int(os.environ.get("HSE_MAX_RECORDS", "5")),
        MAX_CONTENT_LENGTH=int(os.environ.get("MAX_UPLOAD_BYTES", str(default_max_body))),
        PROJECT_NAME_DEFAULT="",
        SQLALCHEMY_DATABASE_URI=f"sqlite:///{base_dir / 'auhse.db'}",
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
    )

    if config:
        app.config.update(config)

    Path(app.config["UPLOAD_FOLDER"]).mkdir(parents=True, exist_ok=True)
    Path(app.config["OUTPUT_FOLDER"]).mkdir(parents=True, exist_ok=True)

    oauth = OAuth(app)
    if _google_oauth_configured():
        oauth.register(
            name="google",
            client_id=os.environ["GOOGLE_OAUTH_CLIENT_ID"].strip(),
            client_secret=os.environ["GOOGLE_OAUTH_CLIENT_SECRET"].strip(),
            server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
            client_kwargs={"scope": "openid email profile"},
        )

    db.init_app(app)
    login_manager = LoginManager()
    login_manager.init_app(app)
    login_manager.login_view = "login"
    login_manager.login_message = "Please sign in to access this page."
    login_manager.login_message_category = "info"

    @login_manager.user_loader
    def load_user(user_id):
        return db.session.get(User, int(user_id))

    with app.app_context():
        db.create_all()
        ensure_schema()

    @app.context_processor
    def inject_google_oauth_flag():
        return {"google_oauth_enabled": _google_oauth_configured()}

    register_routes(app, oauth)
    register_error_handlers(app)
    return app


def register_error_handlers(app: Flask) -> None:
    msg = (
        "Upload too large for server limits. "
        "Images: up to 20 files, 100 MB combined. Video: up to 250 MB."
    )

    @app.errorhandler(RequestEntityTooLarge)
    def request_entity_too_large(_e):
        flash(msg, "error")
        if request.path.startswith("/inspection/new") or (
            request.referrer and "inspection" in (request.referrer or "")
        ):
            return redirect(url_for("new_inspection"))
        return redirect(url_for("index"))


def register_routes(app: Flask, oauth: OAuth) -> None:
    # ── Landing Page ──────────────────────────────────────────────
    @app.route("/")
    def index():
        return render_template("index.html")

    # ── Google OAuth ──────────────────────────────────────────────
    @app.route("/auth/google")
    def google_login():
        if not _google_oauth_configured():
            flash("Google sign-in is not configured on this server.", "error")
            return redirect(url_for("login"))
        redirect_uri = os.environ["GOOGLE_OAUTH_REDIRECT_URI"].strip()
        return oauth.google.authorize_redirect(redirect_uri)

    @app.route("/auth/google/callback")
    def google_callback():
        if not _google_oauth_configured():
            flash("Google sign-in is not configured.", "error")
            return redirect(url_for("login"))
        try:
            token = oauth.google.authorize_access_token()
        except Exception:
            flash("Google sign-in was cancelled or failed.", "error")
            return redirect(url_for("login"))

        userinfo = token.get("userinfo")
        if not userinfo:
            try:
                r = requests.get(
                    "https://openidconnect.googleapis.com/v1/userinfo",
                    headers={"Authorization": f"Bearer {token['access_token']}"},
                    timeout=30,
                )
                r.raise_for_status()
                userinfo = r.json()
            except Exception:
                flash("Could not load your Google profile.", "error")
                return redirect(url_for("login"))

        sub = userinfo.get("sub")
        email = (userinfo.get("email") or "").strip().lower()
        name = (userinfo.get("name") or "").strip() or (email.split("@")[0] if email else "User")

        if not sub or not email:
            flash("Google did not return a usable email address.", "error")
            return redirect(url_for("login"))

        user = User.query.filter_by(google_sub=sub).first()
        if not user:
            user = User.query.filter_by(email=email).first()
            if user:
                user.google_sub = sub
                if name and (not user.full_name or user.full_name == user.email.split("@")[0]):
                    user.full_name = name
            else:
                user = User(full_name=name, email=email, google_sub=sub)
                user.set_password(secrets.token_urlsafe(32))
                db.session.add(user)
        db.session.commit()
        login_user(user, remember=True)
        flash("Signed in with Google.", "success")
        next_page = request.args.get("next")
        return redirect(next_page or url_for("dashboard"))

    # ── Auth Routes ───────────────────────────────────────────────
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if current_user.is_authenticated:
            return redirect(url_for("dashboard"))

        if request.method == "POST":
            email = request.form.get("email", "").strip().lower()
            password = request.form.get("password", "")
            remember = request.form.get("remember") == "on"

            user = User.query.filter_by(email=email).first()
            if user and user.check_password(password):
                login_user(user, remember=remember)
                next_page = request.args.get("next")
                return redirect(next_page or url_for("dashboard"))
            flash("Invalid email or password.", "error")

        return render_template("auth/login.html")

    @app.route("/register", methods=["GET", "POST"])
    def register():
        if current_user.is_authenticated:
            return redirect(url_for("dashboard"))

        if request.method == "POST":
            full_name = request.form.get("full_name", "").strip()
            email = request.form.get("email", "").strip().lower()
            password = request.form.get("password", "")
            confirm_password = request.form.get("confirm_password", "")

            if not all([full_name, email, password]):
                flash("All fields are required.", "error")
            elif password != confirm_password:
                flash("Passwords do not match.", "error")
            elif len(password) < 6:
                flash("Password must be at least 6 characters.", "error")
            elif User.query.filter_by(email=email).first():
                flash("An account with this email already exists.", "error")
            else:
                user = User(full_name=full_name, email=email)
                user.set_password(password)
                db.session.add(user)
                db.session.commit()
                login_user(user)
                flash("Account created successfully.", "success")
                return redirect(url_for("dashboard"))

        return render_template("auth/register.html")

    @app.route("/forgot-password", methods=["GET", "POST"])
    def forgot_password():
        if request.method == "POST":
            _ = request.form.get("email", "").strip().lower()
            flash("If an account with that email exists, a reset link has been sent.", "success")
        return render_template("auth/forgot_password.html")

    @app.route("/logout")
    @login_required
    def logout():
        logout_user()
        flash("You have been signed out.", "success")
        return redirect(url_for("index"))

    # ── Dashboard ─────────────────────────────────────────────────
    @app.route("/dashboard")
    @login_required
    def dashboard():
        total_inspections = Inspection.query.join(Project).filter(
            Project.user_id == current_user.id
        ).count()
        completed = Inspection.query.join(Project).filter(
            Project.user_id == current_user.id,
            Inspection.status == "completed",
        ).count()
        pending = Inspection.query.join(Project).filter(
            Project.user_id == current_user.id,
            Inspection.status.in_(["pending", "processing"]),
        ).count()
        recent_inspections = (
            Inspection.query.join(Project)
            .filter(Project.user_id == current_user.id)
            .order_by(Inspection.created_at.desc())
            .limit(10)
            .all()
        )
        return render_template(
            "dashboard.html",
            total_inspections=total_inspections,
            completed=completed,
            pending=pending,
            recent_inspections=recent_inspections,
        )

    def _save_uploads(
        upload_dir: Path,
        image_files: List[Any],
        video_file: Optional[Any],
    ) -> Tuple[List[Path], Optional[Path]]:
        """Save images and optional video; returns (image_paths, video_path)."""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        image_paths: List[Path] = []
        for i, f in enumerate(image_files):
            if not f or not f.filename:
                continue
            name = secure_filename(f.filename)
            p = upload_dir / f"{ts}_{i}_{name}"
            f.save(p)
            image_paths.append(p)
        video_path: Optional[Path] = None
        if video_file and video_file.filename:
            name = secure_filename(video_file.filename)
            video_path = upload_dir / f"{ts}_v_{name}"
            video_file.save(video_path)
        return image_paths, video_path

    # ── New Inspection ────────────────────────────────────────────
    @app.route("/inspection/new", methods=["GET", "POST"])
    @login_required
    def new_inspection():
        upload_ctx = {
            "max_images": MAX_INSPECTION_IMAGES,
            "max_image_mb": MAX_IMAGE_BATCH_BYTES // (1024 * 1024),
            "max_video_mb": MAX_VIDEO_BYTES // (1024 * 1024),
        }
        if request.method == "POST":
            form = request.form
            project_name = form.get("project_name", "").strip()
            site_location = form.get("site_location", "").strip()
            inspection_by = form.get("inspection_by", "").strip()
            verified_by = form.get("verified_by", "").strip()
            notes = form.get("notes", "").strip()
            upload_mode = (form.get("upload_mode") or "image").strip().lower()
            if upload_mode not in ("image", "video", "both"):
                upload_mode = "image"

            missing = [
                f
                for f, v in {
                    "Project Name": project_name,
                    "Site Location": site_location,
                    "Inspection By": inspection_by,
                    "Verified By": verified_by,
                }.items()
                if not v
            ]

            raw_images = request.files.getlist("site_photos")
            image_inputs = [f for f in raw_images if f and f.filename]
            video_f = request.files.get("site_video")

            if missing:
                flash(f"Missing required fields: {', '.join(missing)}.", "error")
                return redirect(url_for("new_inspection"))

            if upload_mode == "image":
                if len(image_inputs) < 1:
                    flash("Please upload at least one image.", "error")
                    return redirect(url_for("new_inspection"))
                if len(image_inputs) > MAX_INSPECTION_IMAGES:
                    flash(f"Too many images (max {MAX_INSPECTION_IMAGES}).", "error")
                    return redirect(url_for("new_inspection"))
                ok, err, _ = _validate_image_files(image_inputs)
                if not ok:
                    flash(err, "error")
                    return redirect(url_for("new_inspection"))
            elif upload_mode == "video":
                if not video_f or not video_f.filename:
                    flash("Please upload a video.", "error")
                    return redirect(url_for("new_inspection"))
                ok, err, _ = _validate_video_file(video_f)
                if not ok:
                    flash(err, "error")
                    return redirect(url_for("new_inspection"))
            else:  # both
                if len(image_inputs) < 1:
                    flash("Please upload at least one image.", "error")
                    return redirect(url_for("new_inspection"))
                if len(image_inputs) > MAX_INSPECTION_IMAGES:
                    flash(f"Too many images (max {MAX_INSPECTION_IMAGES}).", "error")
                    return redirect(url_for("new_inspection"))
                if not video_f or not video_f.filename:
                    flash("Please upload a video as well (mode: both).", "error")
                    return redirect(url_for("new_inspection"))
                ok, err, _ = _validate_image_files(image_inputs)
                if not ok:
                    flash(err, "error")
                    return redirect(url_for("new_inspection"))
                ok, err, _ = _validate_video_file(video_f)
                if not ok:
                    flash(err, "error")
                    return redirect(url_for("new_inspection"))

            upload_root = Path(app.config["UPLOAD_FOLDER"])
            if upload_mode == "image":
                img_paths, vid_path = _save_uploads(upload_root, image_inputs, None)
            elif upload_mode == "video":
                img_paths, vid_path = _save_uploads(upload_root, [], video_f)
            else:
                img_paths, vid_path = _save_uploads(upload_root, image_inputs, video_f)

            project = Project.query.filter_by(
                user_id=current_user.id,
                project_name=project_name,
                site_location=site_location,
            ).first()
            if not project:
                project = Project(
                    user_id=current_user.id,
                    project_name=project_name,
                    site_location=site_location,
                )
                db.session.add(project)
                db.session.commit()

            primary_image = str(img_paths[0]) if img_paths else ""
            all_image_strs = [str(p) for p in img_paths]
            inspection = Inspection(
                project_id=project.id,
                inspected_by=inspection_by,
                verified_by=verified_by,
                notes=notes,
                image_url=primary_image,
                image_paths_json=json.dumps(all_image_strs) if all_image_strs else None,
                video_path=str(vid_path) if vid_path else None,
                status="processing",
            )
            db.session.add(inspection)
            db.session.commit()

            if upload_mode == "video":
                result = process_video_only_request(app.config, str(vid_path))
                if result["status"] == "ok":
                    inspection.status = "completed"
                    inspection.output_folder = result["output_folder"]
                    inspection.risk_level = ""
                    inspection.risk_score = 0
                    db.session.commit()
                    flash("Video analysis completed.", "success")
                    return redirect(url_for("inspection_detail", inspection_id=inspection.id))
                inspection.status = "failed"
                db.session.commit()
                flash(result.get("message", "Video processing failed."), "error")
                return redirect(url_for("new_inspection"))

            result = process_hse_request(
                config=app.config,
                image_path=primary_image,
                project_name=project_name,
                site_location=site_location,
                inspection_by=inspection_by,
                verified_by=verified_by,
                max_records=app.config.get("MAX_RECORDS"),
                image_paths_for_yolo=all_image_strs,
            )

            def _persist_hse_success(ins: Inspection, res: Dict[str, Any]) -> None:
                ins.status = "completed"
                ins.output_folder = res["output_folder"]
                hse = res.get("hse", {})
                risk = hse.get("risk_analysis", {})
                ins.risk_level = risk.get("risk_level", "").title()
                ins.risk_score = risk.get("risk_rating_lxS", 0)
                for det in res.get("detections", []):
                    hd = HazardDetection(
                        inspection_id=ins.id,
                        hazard_type=det.get("class_name", ""),
                        confidence=det.get("confidence", 0),
                        bounding_box=json.dumps(det.get("bbox", [])),
                    )
                    db.session.add(hd)
                report = Report(
                    inspection_id=ins.id,
                    report_url=res.get("docx_path", ""),
                    json_url=res.get("json_path", ""),
                    raw_url=res.get("raw_text_path", ""),
                    report_status="generated",
                    generated_at=(
                        datetime.fromisoformat(res["generated_at"])
                        if res.get("generated_at")
                        else datetime.utcnow()
                    ),
                )
                db.session.add(report)

            if result["status"] == "ok":
                _persist_hse_success(inspection, result)
                if upload_mode == "both" and vid_path and result.get("output_folder"):
                    try:
                        append_video_summary_to_folder(
                            result["output_folder"],
                            app.config,
                            str(vid_path),
                        )
                    except Exception as exc:
                        flash(f"Report saved; video sidecar failed: {exc}", "warning")
                db.session.commit()
                flash("Inspection completed and report generated.", "success")
                return redirect(url_for("inspection_detail", inspection_id=inspection.id))

            if result["status"] == "not_relevant":
                inspection.status = "completed"
                inspection.output_folder = result.get("output_folder", "")
                db.session.commit()
                if upload_mode == "both" and vid_path and result.get("output_folder"):
                    try:
                        append_video_summary_to_folder(
                            result["output_folder"],
                            app.config,
                            str(vid_path),
                        )
                        db.session.commit()
                    except Exception:
                        pass
                flash(result["message"], "warning")
                return redirect(url_for("inspection_detail", inspection_id=inspection.id))

            inspection.status = "failed"
            db.session.commit()
            flash(result["message"], "error")
            return redirect(url_for("new_inspection"))

        return render_template("inspection.html", **upload_ctx)

    # ── Inspection Detail ─────────────────────────────────────────
    @app.route("/inspection/<int:inspection_id>")
    @login_required
    def inspection_detail(inspection_id):
        inspection = Inspection.query.get_or_404(inspection_id)
        if inspection.project.user_id != current_user.id:
            flash("Access denied.", "error")
            return redirect(url_for("dashboard"))

        report = inspection.reports.first()
        hse_data = None
        result_context: Dict[str, Any] = {}
        video_summary = None
        video_json_url = ""

        if report and report.json_url and Path(report.json_url).exists():
            try:
                hse_data = json.loads(Path(report.json_url).read_text(encoding="utf-8-sig"))
            except Exception:
                pass

        folder = inspection.output_folder
        if folder:
            out = Path(app.config["OUTPUT_FOLDER"]) / folder
            vj = out / VIDEO_YOLO_JSON
            if vj.exists():
                try:
                    video_summary = json.loads(vj.read_text(encoding="utf-8"))
                except Exception:
                    video_summary = None
                video_json_url = url_for(
                    "download_report",
                    folder=folder,
                    asset=VIDEO_YOLO_JSON,
                )

        if report:
            if folder:
                result_context["docx_url"] = (
                    url_for("download_report", folder=folder, asset=Path(report.report_url).name)
                    if report.report_url
                    else ""
                )
                result_context["json_url"] = (
                    url_for("download_report", folder=folder, asset=Path(report.json_url).name)
                    if report.json_url
                    else ""
                )
                result_context["raw_url"] = (
                    url_for("download_report", folder=folder, asset=Path(report.raw_url).name)
                    if report.raw_url
                    else ""
                )

        detections = inspection.hazards.all()

        return render_template(
            "inspection_detail.html",
            inspection=inspection,
            report=report,
            hse_data=hse_data,
            result_context=result_context,
            detections=detections,
            video_summary=video_summary,
            video_json_url=video_json_url,
        )

    # ── History / Archive ─────────────────────────────────────────
    @app.route("/history")
    @login_required
    def history():
        search = request.args.get("search", "").strip()
        status_filter = request.args.get("status", "").strip()
        risk_filter = request.args.get("risk", "").strip()

        query = Inspection.query.join(Project).filter(Project.user_id == current_user.id)

        if search:
            query = query.filter(
                db.or_(
                    Project.project_name.ilike(f"%{search}%"),
                    Project.site_location.ilike(f"%{search}%"),
                    Inspection.inspected_by.ilike(f"%{search}%"),
                )
            )
        if status_filter:
            query = query.filter(Inspection.status == status_filter)
        if risk_filter:
            query = query.filter(Inspection.risk_level.ilike(f"%{risk_filter}%"))

        inspections = query.order_by(Inspection.created_at.desc()).all()

        return render_template(
            "history.html",
            inspections=inspections,
            search=search,
            status_filter=status_filter,
            risk_filter=risk_filter,
        )

    # ── Contact ───────────────────────────────────────────────────
    @app.route("/contact", methods=["GET", "POST"])
    def contact():
        if request.method == "POST":
            name = request.form.get("name", "").strip()
            email = request.form.get("email", "").strip()
            message = request.form.get("message", "").strip()

            if not all([name, email, message]):
                flash("All fields are required.", "error")
            else:
                msg = ContactMessage(name=name, email=email, message=message)
                db.session.add(msg)
                db.session.commit()
                flash("Message sent successfully. We'll get back to you soon.", "success")
                return redirect(url_for("contact"))

        return render_template("contact.html")

    # ── Profile ───────────────────────────────────────────────────
    @app.route("/profile", methods=["GET", "POST"])
    @login_required
    def profile():
        if request.method == "POST":
            action = request.form.get("action")

            if action == "update_profile":
                full_name = request.form.get("full_name", "").strip()
                email = request.form.get("email", "").strip().lower()

                if not full_name or not email:
                    flash("Name and email are required.", "error")
                else:
                    existing = User.query.filter_by(email=email).first()
                    if existing and existing.id != current_user.id:
                        flash("That email is already in use.", "error")
                    else:
                        current_user.full_name = full_name
                        current_user.email = email
                        db.session.commit()
                        flash("Profile updated.", "success")

            elif action == "change_password":
                current_pw = request.form.get("current_password", "")
                new_pw = request.form.get("new_password", "")
                confirm_pw = request.form.get("confirm_password", "")

                if not current_user.check_password(current_pw):
                    flash("Current password is incorrect.", "error")
                elif len(new_pw) < 6:
                    flash("New password must be at least 6 characters.", "error")
                elif new_pw != confirm_pw:
                    flash("New passwords do not match.", "error")
                else:
                    current_user.set_password(new_pw)
                    db.session.commit()
                    flash("Password changed successfully.", "success")

        return render_template("profile.html")

    # ── Report Downloads ──────────────────────────────────────────
    @app.route("/reports/<folder>/<path:asset>")
    @login_required
    def download_report(folder: str, asset: str):
        reports_dir = Path(app.config["OUTPUT_FOLDER"]) / folder
        if not reports_dir.exists():
            flash("Requested report folder does not exist.", "error")
            return redirect(url_for("dashboard"))
        allowed = (
            Inspection.query.join(Project)
            .filter(
                Inspection.output_folder == folder,
                Project.user_id == current_user.id,
            )
            .first()
        )
        if not allowed:
            flash("Access denied.", "error")
            return redirect(url_for("dashboard"))
        return send_from_directory(reports_dir, asset, as_attachment=True, download_name=asset)

    # ── Legacy analyze route (for backward compatibility) ─────────
    @app.post("/analyze")
    def analyze():
        if not current_user.is_authenticated:
            flash("Please sign in to run inspections.", "info")
            return redirect(url_for("login"))
        return redirect(url_for("new_inspection"))


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)

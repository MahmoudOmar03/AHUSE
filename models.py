"""
SQLAlchemy models for the AUHSE platform.
"""

from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    google_sub = db.Column(db.String(255), unique=True, nullable=True, index=True)
    role = db.Column(db.String(32), default="user")  # user, admin
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    projects = db.relationship("Project", backref="owner", lazy="dynamic")

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f"<User {self.email}>"


class Project(db.Model):
    __tablename__ = "projects"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    project_name = db.Column(db.String(200), nullable=False)
    site_location = db.Column(db.String(300), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    inspections = db.relationship("Inspection", backref="project", lazy="dynamic")

    def __repr__(self):
        return f"<Project {self.project_name}>"


class Inspection(db.Model):
    __tablename__ = "inspections"

    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)
    inspected_by = db.Column(db.String(120), nullable=False)
    verified_by = db.Column(db.String(120), nullable=False)
    inspection_date = db.Column(db.DateTime, default=datetime.utcnow)
    notes = db.Column(db.Text, default="")
    image_url = db.Column(db.String(500))
    image_paths_json = db.Column(db.Text, nullable=True)
    video_path = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(32), default="pending")  # pending, processing, completed, failed
    risk_level = db.Column(db.String(32), default="")
    risk_score = db.Column(db.Integer, default=0)
    output_folder = db.Column(db.String(300), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    hazards = db.relationship("HazardDetection", backref="inspection", lazy="dynamic")
    reports = db.relationship("Report", backref="inspection", lazy="dynamic")

    def __repr__(self):
        return f"<Inspection {self.id} - {self.status}>"


class HazardDetection(db.Model):
    __tablename__ = "hazard_detections"

    id = db.Column(db.Integer, primary_key=True)
    inspection_id = db.Column(db.Integer, db.ForeignKey("inspections.id"), nullable=False)
    hazard_type = db.Column(db.String(200), nullable=False)
    confidence = db.Column(db.Float, default=0.0)
    bounding_box = db.Column(db.Text, default="")  # JSON string
    severity = db.Column(db.String(32), default="")
    recommendation = db.Column(db.Text, default="")

    def __repr__(self):
        return f"<HazardDetection {self.hazard_type}>"


class Report(db.Model):
    __tablename__ = "reports"

    id = db.Column(db.Integer, primary_key=True)
    inspection_id = db.Column(db.Integer, db.ForeignKey("inspections.id"), nullable=False)
    report_url = db.Column(db.String(500), default="")
    json_url = db.Column(db.String(500), default="")
    raw_url = db.Column(db.String(500), default="")
    report_status = db.Column(db.String(32), default="generated")
    generated_at = db.Column(db.DateTime, default=datetime.utcnow)
    version = db.Column(db.Integer, default=1)

    def __repr__(self):
        return f"<Report {self.id} v{self.version}>"


class ContactMessage(db.Model):
    __tablename__ = "contact_messages"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), nullable=False)
    message = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f"<ContactMessage from {self.email}>"

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, abort
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from flask_sqlalchemy import SQLAlchemy
from flask_wtf import FlaskForm
from flask_wtf.file import MultipleFileField, FileAllowed
from wtforms import StringField, IntegerField, DateField, TextAreaField, SelectField, SubmitField, PasswordField
from wtforms.validators import DataRequired, Email, Optional, NumberRange, Length, EqualTo, ValidationError
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import re
import uuid
import shutil
import sqlite3 as sqlite3_stdlib
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired

app = Flask(__name__, instance_relative_config=True)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'corp-crm-super-secret-key-change-me')
os.makedirs(app.instance_path, exist_ok=True)

# Prefer Render persistent disk (/var/data) when mounted, so data survives deploys
_instance_db = os.path.join(app.instance_path, 'corp_crm.db')
if os.environ.get('DATABASE_URL'):
    _default_db_uri = os.environ['DATABASE_URL']
elif os.path.isdir('/var/data'):
    _default_db_uri = 'sqlite:////var/data/corp_crm.db'
else:
    _default_db_uri = 'sqlite:///' + _instance_db
app.config['SQLALCHEMY_DATABASE_URI'] = _default_db_uri
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB total per request

def get_sqlite_filesystem_path():
    """Return the on-disk path for the configured SQLite database, if any."""
    uri = app.config.get('SQLALCHEMY_DATABASE_URI') or ''
    if not uri.startswith('sqlite:///'):
        return None
    path = uri[len('sqlite:///'):]
    if path.startswith('/'):
        return path
    if not os.path.isabs(path):
        path = os.path.join(app.root_path, path)
    return path

# Keep uploads on the same volume as the DB when possible
_db_path = get_sqlite_filesystem_path()
if _db_path and _db_path.startswith('/var/data'):
    app.config['UPLOAD_FOLDER'] = '/var/data/uploads'
else:
    app.config['UPLOAD_FOLDER'] = os.path.join(app.instance_path, 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Optional emergency restore token (set RESTORE_TOKEN in host env to enable)
_ONE_TIME_RESTORE_TOKEN = os.environ.get('RESTORE_TOKEN', '')

ALLOWED_NOTE_EXTENSIONS = {
    'pdf', 'png', 'jpg', 'jpeg', 'gif', 'webp',
    'doc', 'docx', 'xls', 'xlsx', 'csv', 'txt',
    'zip', 'heic',
}

db = SQLAlchemy(app)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Please log in to access this page.'
login_manager.login_message_category = 'info'

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# ==================== MODELS ====================

class Client(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    uen = db.Column(db.String(20), unique=True, nullable=False)
    contact_person = db.Column(db.String(100))
    email = db.Column(db.String(500))  # one or more addresses, comma/semicolon separated
    phone = db.Column(db.String(30))
    fye_month = db.Column(db.Integer)  # 1-12
    incorporation_date = db.Column(db.Date)
    archived = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    # Last composed per-client reminder email (reused next time)
    saved_reminder_subject = db.Column(db.Text)
    saved_reminder_body = db.Column(db.Text)

    deadlines = db.relationship('Deadline', backref='client', lazy=True, cascade="all, delete-orphan")
    invoices = db.relationship('Invoice', backref='client', lazy=True, cascade="all, delete-orphan")
    notes = db.relationship('Note', backref='client', lazy=True, cascade="all, delete-orphan")

    def email_list(self):
        """Return all client email addresses as a list."""
        return parse_emails(self.email)

    def emails_display(self):
        """Comma-separated display of all client emails."""
        emails = self.email_list()
        return ', '.join(emails) if emails else ''

    def __repr__(self):
        return f'<Client {self.name}>'

class Deadline(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey('client.id'), nullable=False)
    deadline_type = db.Column(db.String(50), nullable=False)  # FinancialYearEnd, AGM, AnnualReturn, ECI, TaxFiling
    due_date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(20), default='Pending')  # Pending, Completed, Overdue
    completed_date = db.Column(db.Date)
    reminder_days = db.Column(db.Integer, default=7)  # days before to remind
    last_reminded = db.Column(db.Date)
    comments = db.Column(db.Text)

    @property
    def is_overdue(self):
        # Include both Pending (not yet flipped by the daily job) and Overdue status
        return self.status in ('Pending', 'Overdue') and self.due_date < date.today()

    def get_days_left(self):
        return (self.due_date - date.today()).days

class Invoice(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey('client.id'), nullable=False)
    invoice_type = db.Column(db.String(100), nullable=False)  # e.g. Secretarial Fee, Registered Address
    description = db.Column(db.String(255))
    amount = db.Column(db.Float)
    due_date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(20), default='Pending')  # Pending, Paid, Overdue
    paid_date = db.Column(db.Date)
    reminder_days = db.Column(db.Integer, default=1)
    last_reminded = db.Column(db.Date)

    @property
    def is_overdue(self):
        return self.status in ('Pending', 'Overdue') and self.due_date < date.today()

    def get_days_left(self):
        return (self.due_date - date.today()).days

class Note(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey('client.id'), nullable=False)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)
    title = db.Column(db.String(200), nullable=False)
    content = db.Column(db.Text)
    category = db.Column(db.String(50), default='General')  # ACRA, IRAS, Client, General
    attachments = db.relationship(
        'NoteAttachment',
        backref='note',
        lazy=True,
        cascade='all, delete-orphan',
    )

class NoteAttachment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    note_id = db.Column(db.Integer, db.ForeignKey('note.id'), nullable=False)
    original_filename = db.Column(db.String(255), nullable=False)
    stored_filename = db.Column(db.String(255), nullable=False)
    file_size = db.Column(db.Integer, default=0)

class Setting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    my_email = db.Column(db.String(120), default='yourname@gmail.com')
    smtp_server = db.Column(db.String(100), default='smtp.gmail.com')
    smtp_port = db.Column(db.Integer, default=587)
    smtp_username = db.Column(db.String(120))
    smtp_password = db.Column(db.String(200))
    use_tls = db.Column(db.Boolean, default=True)
    email_client = db.Column(db.Boolean, default=False)
    reminder_subject_template = db.Column(db.Text)
    reminder_body_template = db.Column(db.Text)
    internal_reminder_subject_template = db.Column(db.Text)
    internal_reminder_body_template = db.Column(db.Text)

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password, method='pbkdf2:sha256')

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

# ==================== FORMS ====================

MONTH_NAMES = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
]

FYE_MONTH_CHOICES = [('', '— Select month —')] + [
    (str(i), name) for i, name in enumerate(MONTH_NAMES, start=1)
]

def optional_int_coerce(value):
    if value == '' or value is None:
        return None
    return int(value)


# Split on commas, semicolons, or newlines (not spaces inside addresses)
_EMAIL_SPLIT_RE = re.compile(r'[,;\n\r]+')
_EMAIL_SIMPLE_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


def parse_emails(value):
    """Parse one or more email addresses from a free-text field."""
    if not value:
        return []
    parts = _EMAIL_SPLIT_RE.split(str(value).strip())
    seen = set()
    result = []
    for part in parts:
        email = part.strip()
        if not email:
            continue
        key = email.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(email)
    return result


def validate_multi_email(form, field):
    """Allow one or more emails separated by commas or semicolons."""
    raw = (field.data or '').strip()
    if not raw:
        return
    emails = parse_emails(raw)
    if not emails:
        raise ValidationError('Enter one or more valid email addresses separated by commas.')
    for email in emails:
        if not _EMAIL_SIMPLE_RE.match(email):
            raise ValidationError(f'Invalid email address: {email}')
    # Normalise stored value to comma-separated
    field.data = ', '.join(emails)


class ClientForm(FlaskForm):
    name = StringField('Client Name', validators=[DataRequired()])
    uen = StringField('UEN', validators=[DataRequired()])
    contact_person = StringField('Contact Person')
    email = StringField(
        'Client Email(s)',
        validators=[Optional(), validate_multi_email],
        description='Separate multiple addresses with commas, e.g. a@company.com, b@company.com',
    )
    phone = StringField('Phone')
    fye_month = SelectField(
        'Financial Year End Month',
        choices=FYE_MONTH_CHOICES,
        coerce=optional_int_coerce,
        validators=[Optional()]
    )
    incorporation_date = DateField('Incorporation Date', validators=[Optional()])
    submit = SubmitField('Save Client')

class DeadlineForm(FlaskForm):
    deadline_type = SelectField('Deadline Type', choices=[
        ('FinancialYearEnd', 'Financial Year End'),
        ('AGM', 'AGM Deadline'),
        ('AnnualReturn', 'Annual Return Filing'),
        ('ECI', 'ECI Filing'),
        ('TaxFiling', 'Income Tax Filing'),
        ('Others', 'Others'),
    ], validators=[DataRequired()])
    due_date = DateField('Due Date', validators=[DataRequired()])
    reminder_days = IntegerField('Email me this many days before', default=7, validators=[NumberRange(min=1)])
    comments = TextAreaField('Comments (optional)')
    submit = SubmitField('Add Deadline')

class InvoiceForm(FlaskForm):
    invoice_type = SelectField('Invoice Type', choices=[
        ('Secretarial Fee', 'Secretarial Fee'),
        ('Registered Address', 'Registered Address'),
        ('Director Fee', 'Director / Nominee Fee'),
        ('Other', 'Other Service Fee')
    ], validators=[DataRequired()])
    description = StringField('Description (optional)')
    amount = StringField('Amount (SGD, optional)')
    due_date = DateField('Payment Due Date', validators=[DataRequired()])
    submit = SubmitField('Add Invoice')

class NoteForm(FlaskForm):
    title = StringField('Title / Action', validators=[DataRequired()])
    content = TextAreaField('Details')
    category = SelectField('Category', choices=[
        ('General', 'General'),
        ('ACRA', 'ACRA Update'),
        ('IRAS', 'IRAS / Tax'),
        ('Client', 'Client Communication'),
        ('Other', 'Other')
    ], default='General')
    attachments = MultipleFileField(
        'Attachments',
        validators=[
            FileAllowed(
                list(ALLOWED_NOTE_EXTENSIONS),
                'Allowed types: PDF, images, Word, Excel, CSV, TXT, ZIP',
            )
        ],
    )
    submit = SubmitField('Add Note')


def allowed_note_file(filename):
    if not filename or '.' not in filename:
        return False
    return filename.rsplit('.', 1)[1].lower() in ALLOWED_NOTE_EXTENSIONS


def save_note_attachments(note, file_list):
    """Save uploaded files for a note. Returns (saved_count, error_message)."""
    if not file_list:
        return 0, None
    upload_root = app.config['UPLOAD_FOLDER']
    note_dir = os.path.join(upload_root, 'notes', str(note.client_id), str(note.id))
    os.makedirs(note_dir, exist_ok=True)
    saved = 0
    for storage in file_list:
        if not storage or not storage.filename:
            continue
        original = storage.filename
        if not allowed_note_file(original):
            return saved, f'File type not allowed: {original}'
        safe = secure_filename(original) or 'file'
        stored = f'{uuid.uuid4().hex}_{safe}'
        path = os.path.join(note_dir, stored)
        storage.save(path)
        size = os.path.getsize(path)
        db.session.add(NoteAttachment(
            note_id=note.id,
            original_filename=original[:255],
            stored_filename=stored,
            file_size=size,
        ))
        saved += 1
    return saved, None


def note_attachment_path(attachment):
    note = attachment.note
    return os.path.join(
        app.config['UPLOAD_FOLDER'],
        'notes',
        str(note.client_id),
        str(note.id),
        attachment.stored_filename,
    )

class SettingsForm(FlaskForm):
    my_email = StringField('Your Email (for reminders)', validators=[DataRequired(), Email()])
    smtp_server = StringField('SMTP Server')
    smtp_port = IntegerField('SMTP Port')
    smtp_username = StringField('SMTP Username')
    smtp_password = PasswordField('SMTP Password')
    use_tls = SelectField('Use TLS', choices=[('True', 'Yes'), ('False', 'No')])
    email_client = SelectField(
        'Also email the client automatically',
        choices=[('False', 'No — notify me only (I will remind the client)'), ('True', 'Yes — email client and me')]
    )
    internal_reminder_subject_template = StringField(
        'Your Reminder Email Subject',
        validators=[DataRequired(), Length(max=500)]
    )
    internal_reminder_body_template = TextAreaField(
        'Your Reminder Email Body',
        validators=[DataRequired(), Length(max=5000)]
    )
    reminder_subject_template = StringField(
        'Client Email Subject (if enabled)',
        validators=[Optional(), Length(max=500)]
    )
    reminder_body_template = TextAreaField(
        'Client Email Body (if enabled)',
        validators=[Optional(), Length(max=5000)]
    )
    submit = SubmitField('Save Settings')

class LoginForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=3, max=80)])
    password = PasswordField('Password', validators=[DataRequired()])
    submit = SubmitField('Login')

class AddUserForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=3, max=80)])
    password = PasswordField('Password', validators=[DataRequired(), Length(min=6)])
    confirm_password = PasswordField(
        'Confirm Password',
        validators=[DataRequired(), EqualTo('password', message='Passwords must match.')]
    )
    submit = SubmitField('Add User')

class ChangePasswordForm(FlaskForm):
    current_password = PasswordField('Current Password', validators=[DataRequired()])
    new_password = PasswordField('New Password', validators=[DataRequired(), Length(min=6)])
    confirm_password = PasswordField(
        'Confirm New Password',
        validators=[DataRequired(), EqualTo('new_password', message='Passwords must match.')]
    )
    submit = SubmitField('Change Password')

class ForgotPasswordForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=3, max=80)])
    submit = SubmitField('Send Reset Link')

class ResetPasswordForm(FlaskForm):
    new_password = PasswordField('New Password', validators=[DataRequired(), Length(min=6)])
    confirm_password = PasswordField(
        'Confirm New Password',
        validators=[DataRequired(), EqualTo('new_password', message='Passwords must match.')]
    )
    submit = SubmitField('Set New Password')

class AdminSetPasswordForm(FlaskForm):
    new_password = PasswordField('New Password', validators=[DataRequired(), Length(min=6)])
    confirm_password = PasswordField(
        'Confirm New Password',
        validators=[DataRequired(), EqualTo('new_password', message='Passwords must match.')]
    )
    submit = SubmitField('Set Password')

class ClientReminderForm(FlaskForm):
    subject = StringField('Email subject', validators=[DataRequired(), Length(max=300)])
    body = TextAreaField('Email message', validators=[DataRequired(), Length(max=10000)])
    send_to_me = SelectField(
        'Email me a copy',
        choices=[('True', 'Yes'), ('False', 'No')],
        default='True',
    )
    send_to_client = SelectField(
        'Email the client',
        choices=[('True', 'Yes'), ('False', 'No')],
        default='True',
    )
    submit = SubmitField('Send reminder email')


class ClientLettersForm(FlaskForm):
    """Compose email notifying client that mail/letters are ready for collection."""
    subject = StringField('Email subject', validators=[DataRequired(), Length(max=300)])
    body = TextAreaField('Email message', validators=[DataRequired(), Length(max=10000)])
    send_to_me = SelectField(
        'Email me a copy',
        choices=[('True', 'Yes'), ('False', 'No')],
        default='True',
    )
    send_to_client = SelectField(
        'Email the client',
        choices=[('True', 'Yes'), ('False', 'No')],
        default='True',
    )
    submit = SubmitField('Send letters notice')

# ==================== HELPER FUNCTIONS ====================

PASSWORD_RESET_SALT = 'astutecorp-crm-password-reset'
PASSWORD_RESET_MAX_AGE = 3600  # 1 hour


def _password_reset_serializer():
    return URLSafeTimedSerializer(app.config['SECRET_KEY'], salt=PASSWORD_RESET_SALT)


def generate_password_reset_token(user_id):
    return _password_reset_serializer().dumps({'user_id': user_id})


def verify_password_reset_token(token, max_age=PASSWORD_RESET_MAX_AGE):
    try:
        data = _password_reset_serializer().loads(token, max_age=max_age)
        return data.get('user_id')
    except SignatureExpired:
        return None
    except BadSignature:
        return None

# Types used by Calculate from FYE (Others is manual-only and never auto-added)
DEADLINE_TYPES = ['FinancialYearEnd', 'ECI', 'AGM', 'AnnualReturn', 'TaxFiling']

DEADLINE_TYPE_LABELS = {
    'FinancialYearEnd': 'Financial Year End',
    'AGM': 'AGM Deadline',
    'AnnualReturn': 'Annual Return Filing',
    'ECI': 'ECI Filing',
    'TaxFiling': 'Income Tax Filing',
    'Others': 'Others',
}

DEFAULT_LETTERS_SUBJECT = 'Mail / letters awaiting collection — {client_name}'
DEFAULT_LETTERS_BODY = (
    'Dear {contact_person},\n\n'
    'We are writing to inform you that we have received letter(s) and/or mail for '
    '{client_name} (UEN: {uen}) at our office under our registered address service.\n\n'
    'Please arrange to collect them at your earliest convenience during our office hours.\n\n'
    'If you have any questions, please reply to this email.\n\n'
    'Best regards,\n'
    'Astute Solutions'
)

DEFAULT_INTERNAL_REMINDER_SUBJECT = 'Upcoming deadline: {client_name} - {deadline_type} in {days_remaining} days'

DEFAULT_INTERNAL_REMINDER_BODY = """Hi,

A client deadline is coming up. Please follow up and remind the client.

Client: {client_name}
UEN: {uen}
Contact: {contact_person}
Email: {client_email}
Phone: {client_phone}

Deadline: {deadline_type}
Due Date: {due_date}
Days Remaining: {days_remaining}
Status: {deadline_status}
Notes: {comments}

---
AstuteCorp CRM"""

DEFAULT_OVERDUE_INTERNAL_SUBJECT = 'OVERDUE: {client_name} - {deadline_type} was due {due_date}'

DEFAULT_OVERDUE_INTERNAL_BODY = """Hi,

A client deadline is OVERDUE. Please follow up with the client.

Client: {client_name}
UEN: {uen}
Contact: {contact_person}
Email: {client_email}
Phone: {client_phone}

Deadline: {deadline_type}
Due Date: {due_date}
Days Overdue: {days_overdue}
Status: {deadline_status}
Notes: {comments}

---
AstuteCorp CRM"""

DEFAULT_REMINDER_SUBJECT = 'Deadline Reminder: {client_name} - {deadline_type} due {due_date}'

DEFAULT_REMINDER_BODY = """Dear {contact_person},

This is a reminder for {client_name} (UEN: {uen}).

Deadline: {deadline_type}
Due Date: {due_date}
Days Remaining: {days_remaining}

Please ensure this is actioned on time.

Best regards,
Corporate Service Provider"""

REMINDER_TEMPLATE_PLACEHOLDERS = [
    ('{contact_person}', 'Client contact person (or "Client" if blank)'),
    ('{client_name}', 'Client company name'),
    ('{uen}', 'Client UEN'),
    ('{client_email}', 'Client email address(es)'),
    ('{client_phone}', 'Client phone number'),
    ('{deadline_type}', 'Deadline type, e.g. AGM Deadline'),
    ('{due_date}', 'Due date in dd-mm-yyyy format'),
    ('{days_remaining}', 'Number of days until due date'),
    ('{comments}', 'Deadline comments/notes'),
]

DEADLINE_TYPE_BADGE_CLASSES = {
    'FinancialYearEnd': 'badge-deadline-fye',
    'AGM': 'badge-deadline-agm',
    'AnnualReturn': 'badge-deadline-ar',
    'ECI': 'badge-deadline-eci',
    'TaxFiling': 'badge-deadline-tax',
    'Others': 'badge-deadline-others',
}

@app.template_filter('deadline_label')
def deadline_label(deadline_type):
    return DEADLINE_TYPE_LABELS.get(deadline_type, deadline_type)

@app.template_filter('deadline_badge_class')
def deadline_badge_class(deadline_type):
    return DEADLINE_TYPE_BADGE_CLASSES.get(deadline_type, 'badge-deadline-default')

@app.template_filter('month_name')
def month_name(month_number):
    if not month_number or month_number < 1 or month_number > 12:
        return '-'
    return MONTH_NAMES[month_number - 1]

def render_reminder_template(template, **values):
    """Fill reminder template placeholders with deadline/client values."""
    result = template
    for key, value in values.items():
        result = result.replace(f'{{{key}}}', str(value))
    return result

def get_internal_reminder_templates(settings):
    subject = (settings.internal_reminder_subject_template or '').strip() or DEFAULT_INTERNAL_REMINDER_SUBJECT
    body = (settings.internal_reminder_body_template or '').strip() or DEFAULT_INTERNAL_REMINDER_BODY
    return subject, body

def get_client_reminder_templates(settings):
    subject = (settings.reminder_subject_template or '').strip() or DEFAULT_REMINDER_SUBJECT
    body = (settings.reminder_body_template or '').strip() or DEFAULT_REMINDER_BODY
    return subject, body


def get_client_open_deadlines(client_id):
    """Pending or overdue deadlines for one client (not completed)."""
    return (
        Deadline.query.filter(
            Deadline.client_id == client_id,
            Deadline.status.in_(['Pending', 'Overdue']),
        )
        .order_by(Deadline.due_date.asc())
        .all()
    )


def format_deadline_summary_lines(deadlines):
    lines = []
    for dl in deadlines:
        label = DEADLINE_TYPE_LABELS.get(dl.deadline_type, dl.deadline_type)
        days_left = dl.get_days_left()
        if days_left < 0:
            timing = f'{abs(days_left)} day(s) overdue'
        elif days_left == 0:
            timing = 'due today'
        else:
            timing = f'{days_left} day(s) remaining'
        line = f'- {label}: due {format_date_display(dl.due_date)} ({timing})'
        if dl.comments:
            line += f' — {dl.comments.strip()}'
        lines.append(line)
    return '\n'.join(lines)


def build_client_reminder_defaults(client, deadlines, settings):
    """Pre-fill subject/body for a one-client reminder (editable before send)."""
    settings = settings or Setting.query.first()
    subject_tpl, body_tpl = get_client_reminder_templates(settings)
    summary = format_deadline_summary_lines(deadlines)
    if deadlines:
        primary = deadlines[0]
        values = build_deadline_template_values(primary)
        subject = render_reminder_template(subject_tpl, **values)
        if len(deadlines) > 1:
            subject = f'Deadline reminder: {client.name} ({len(deadlines)} outstanding items)'
        body = render_reminder_template(body_tpl, **values)
        if len(deadlines) > 1:
            body = (
                f'Dear {client.contact_person or "Client"},\n\n'
                f'This is a reminder for {client.name} (UEN: {client.uen or "-"}).\n\n'
                f'Outstanding deadlines:\n{summary}\n\n'
                f'Please ensure these are actioned on time.\n\n'
                f'Best regards,\nCorporate Service Provider'
            )
    else:
        subject = f'Reminder: {client.name}'
        body = (
            f'Dear {client.contact_person or "Client"},\n\n'
            f'This is a follow-up regarding {client.name} (UEN: {client.uen or "-"}).\n\n'
            f'Best regards,\nCorporate Service Provider'
        )
    return subject, body, summary


def build_deadline_template_values(deadline):
    client = deadline.client
    deadline_name = DEADLINE_TYPE_LABELS.get(deadline.deadline_type, deadline.deadline_type)
    days_left = deadline.get_days_left()
    if days_left < 0:
        deadline_status = 'OVERDUE'
        days_remaining_display = f'{abs(days_left)} day(s) overdue'
        days_overdue = abs(days_left)
    elif days_left == 0:
        deadline_status = 'DUE TODAY'
        days_remaining_display = 'Due today'
        days_overdue = 0
    else:
        deadline_status = 'Upcoming'
        days_remaining_display = days_left
        days_overdue = 0
    return {
        'contact_person': client.contact_person or 'Client',
        'client_name': client.name,
        'uen': client.uen,
        'client_email': client.emails_display() or 'Not set',
        'client_phone': client.phone or 'Not set',
        'deadline_type': deadline_name,
        'due_date': format_date_display(deadline.due_date),
        'days_remaining': days_remaining_display,
        'days_overdue': days_overdue,
        'deadline_status': deadline_status,
        'comments': (deadline.comments or '').strip() or 'None',
    }

def format_date_display(value, include_time=False):
    if not value:
        return '-'
    if include_time and isinstance(value, datetime):
        return value.strftime('%d-%m-%Y %H:%M')
    if isinstance(value, datetime):
        value = value.date()
    return value.strftime('%d-%m-%Y')

@app.template_filter('format_date')
def format_date(value):
    return format_date_display(value)

@app.template_filter('format_datetime')
def format_datetime(value):
    return format_date_display(value, include_time=True)

def add_months_keep_month_end(start, months):
    """Add months, keeping month-end dates at month-end (30 Sep + 3 months = 31 Dec)."""
    result = start + relativedelta(months=months)
    if (start + timedelta(days=1)).month != start.month:
        result = result + relativedelta(day=31)
    return result

def calculate_default_deadlines(fye_date):
    """Suggest default deadlines based on an existing Financial Year End date."""
    if not fye_date:
        return []

    deadlines = []

    # ECI: 3 months after financial year end
    eci_due = add_months_keep_month_end(fye_date, 3)
    deadlines.append(('ECI', eci_due))

    # AGM: 6 months after financial year end
    agm_due = add_months_keep_month_end(fye_date, 6)
    deadlines.append(('AGM', agm_due))

    # Annual Return: assume 1 month after AGM (common practice)
    ar_due = add_months_keep_month_end(fye_date, 7)
    deadlines.append(('AnnualReturn', ar_due))

    # Income Tax Filing: 30 November of the year following FYE
    tax_due = date(fye_date.year + 1, 11, 30)
    deadlines.append(('TaxFiling', tax_due))

    return deadlines

def group_deadlines_by_client(deadlines):
    """Group deadline records into one row per client, sorted by earliest due date."""
    grouped = {}
    for deadline in deadlines:
        entry = grouped.get(deadline.client_id)
        if not entry:
            entry = {
                'client': deadline.client,
                'deadlines': [],
                'earliest_due': deadline.due_date,
                'soonest_days_left': deadline.get_days_left(),
            }
            grouped[deadline.client_id] = entry
        entry['deadlines'].append(deadline)
        if deadline.due_date < entry['earliest_due']:
            entry['earliest_due'] = deadline.due_date
            entry['soonest_days_left'] = deadline.get_days_left()

    for entry in grouped.values():
        entry['deadlines'].sort(key=lambda dl: dl.due_date)

    return sorted(grouped.values(), key=lambda entry: entry['earliest_due'])

def get_fye_reference_deadlines(client_id):
    """Return Financial Year End deadlines to use when calculating defaults."""
    pending = Deadline.query.filter_by(
        client_id=client_id,
        deadline_type='FinancialYearEnd',
        status='Pending'
    ).order_by(Deadline.due_date).all()

    if pending:
        return pending

    latest = Deadline.query.filter_by(
        client_id=client_id,
        deadline_type='FinancialYearEnd'
    ).order_by(Deadline.due_date.desc()).first()

    return [latest] if latest else []

_last_email_error = None

def send_email(to_email, subject, body, is_html=False):
    """Send email using settings from DB. to_email may be one address or a list."""
    global _last_email_error
    settings = Setting.query.first()
    if not settings or not settings.my_email:
        _last_email_error = 'Your email address is not configured in Settings.'
        print("⚠️ Email settings not configured!")
        return False

    if not settings.smtp_server or not settings.smtp_port:
        _last_email_error = 'SMTP server and port are required.'
        return False

    if not settings.smtp_username or not settings.smtp_password:
        _last_email_error = 'SMTP username and password are required.'
        return False

    if isinstance(to_email, (list, tuple)):
        recipients = [e.strip() for e in to_email if e and str(e).strip()]
    else:
        recipients = [str(to_email).strip()] if to_email and str(to_email).strip() else []

    if not recipients:
        _last_email_error = 'No recipient email address provided.'
        return False

    try:
        msg = MIMEMultipart()
        msg['From'] = settings.my_email
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = subject

        if is_html:
            msg.attach(MIMEText(body, 'html'))
        else:
            msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP(settings.smtp_server, settings.smtp_port, timeout=30)
        if settings.use_tls:
            server.starttls()
        server.login(settings.smtp_username, settings.smtp_password)
        server.send_message(msg)
        server.quit()
        _last_email_error = None
        print(f"✅ Email sent to {', '.join(recipients)}")
        return True
    except Exception as e:
        _last_email_error = str(e)
        print(f"❌ Email failed to {', '.join(recipients)}: {e}")
        return False


def send_email_to_each(emails, subject, body, is_html=False):
    """Send the same email to each address individually. Returns (sent_list, failed_list)."""
    sent = []
    failed = []
    for email in emails:
        if send_email(email, subject, body, is_html=is_html):
            sent.append(email)
        else:
            failed.append(email)
    return sent, failed


def send_deadline_reminder(deadline, overdue=False):
    """Email you about an upcoming or overdue deadline; optionally email the client too."""
    client = deadline.client
    settings = Setting.query.first()
    if not client or not settings or not settings.my_email:
        return False

    template_values = build_deadline_template_values(deadline)

    if overdue:
        # Clear overdue wording for you — still use custom templates if they work with placeholders
        internal_subject = render_reminder_template(DEFAULT_OVERDUE_INTERNAL_SUBJECT, **template_values)
        internal_body = render_reminder_template(DEFAULT_OVERDUE_INTERNAL_BODY, **template_values)
    else:
        internal_subject_template, internal_body_template = get_internal_reminder_templates(settings)
        internal_subject = render_reminder_template(internal_subject_template, **template_values)
        internal_body = render_reminder_template(internal_body_template, **template_values)

    sent_to_self = send_email(settings.my_email, f"[AstuteCorp CRM] {internal_subject}", internal_body)

    sent_to_client = True
    # Only auto-email the client for upcoming deadlines (not overdue), when enabled
    client_emails = client.email_list()
    if not overdue and settings.email_client and client_emails:
        client_subject_template, client_body_template = get_client_reminder_templates(settings)
        client_subject = render_reminder_template(client_subject_template, **template_values)
        client_body = render_reminder_template(client_body_template, **template_values)
        sent_ok, _failed = send_email_to_each(client_emails, client_subject, client_body)
        sent_to_client = bool(sent_ok)

    if sent_to_self or (not overdue and sent_to_client):
        deadline.last_reminded = date.today()
        db.session.commit()
        return bool(sent_to_self)
    return False


def build_letters_notice_defaults(client):
    """Default subject/body for registered-address letters collection notice."""
    values = {
        'contact_person': client.contact_person or 'Client',
        'client_name': client.name,
        'uen': client.uen or '-',
    }
    subject = render_reminder_template(DEFAULT_LETTERS_SUBJECT, **values)
    body = render_reminder_template(DEFAULT_LETTERS_BODY, **values)
    return subject, body

def send_test_reminder_email():
    """Send a sample reminder email to verify SMTP settings."""
    settings = Setting.query.first()
    if not settings or not settings.my_email:
        return False, 'Set your email address in Settings first.'

    sample_values = {
        'contact_person': 'Ms. Tan',
        'client_name': 'Demo Pte. Ltd.',
        'uen': '202612345K',
        'client_email': 'client@example.com',
        'client_phone': '+65 9123 4567',
        'deadline_type': 'AGM Deadline',
        'due_date': format_date_display(date.today() + timedelta(days=7)),
        'days_remaining': 7,
        'comments': 'Waiting for client to confirm AGM date',
    }
    subject_template, body_template = get_internal_reminder_templates(settings)
    subject = render_reminder_template(subject_template, **sample_values)
    body = render_reminder_template(body_template, **sample_values)
    if send_email(settings.my_email, f"[AstuteCorp CRM Test] {subject}", body):
        return True, f'Test email sent to {settings.my_email}.'
    error = _last_email_error or 'Unknown error'
    return False, f'Failed to send test email: {error}'

def check_and_send_reminders():
    """Check pending deadlines and send reminders; mark overdue invoices without emailing."""
    with app.app_context():
        return _check_and_send_reminders()

def _check_and_send_reminders():
    """Send approaching + overdue deadline emails (once per day per deadline)."""
    today = date.today()
    sent_count = 0

    # Pending and already-Overdue deadlines (until marked Completed)
    deadlines = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status.in_(['Pending', 'Overdue'])
    ).all()
    for dl in deadlines:
        days_left = dl.get_days_left()
        already_reminded_today = dl.last_reminded and dl.last_reminded >= today

        # Past due → mark Overdue and email you daily until completed
        if days_left < 0:
            if dl.status != 'Overdue':
                dl.status = 'Overdue'
                db.session.commit()
            # An FYE is a reference date, not a filing: no daily overdue emails once it passes
            if dl.deadline_type == 'FinancialYearEnd':
                continue
            if not already_reminded_today:
                if send_deadline_reminder(dl, overdue=True):
                    sent_count += 1
            continue

        # Due date was moved out after being flagged Overdue → back to Pending
        if dl.status == 'Overdue':
            dl.status = 'Pending'
            db.session.commit()

        # Due today → email you (treated as urgent / overdue-style notice)
        if days_left == 0:
            if not already_reminded_today:
                if send_deadline_reminder(dl, overdue=True):
                    sent_count += 1
            continue

        # Upcoming within reminder window (e.g. 7 days before)
        if dl.reminder_days >= days_left and not already_reminded_today:
            if send_deadline_reminder(dl, overdue=False):
                sent_count += 1

    invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status.in_(['Pending', 'Overdue'])
    ).all()
    for inv in invoices:
        # Overdue only once past due, matching the dashboard's overdue list
        new_status = 'Overdue' if inv.get_days_left() < 0 else 'Pending'
        if inv.status != new_status:
            inv.status = new_status
            db.session.commit()

    print(f"Reminder check complete. Sent {sent_count} deadline reminders.")
    return sent_count

# ==================== ROUTES ====================

@app.route('/')
@login_required
def dashboard():
    today = date.today()
    sixty_days_out = today + timedelta(days=60)
    # Overdue status with a future date means the due date was pushed out after the daily job flagged it
    upcoming = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status.in_(['Pending', 'Overdue']),
        Deadline.due_date >= today,
        Deadline.due_date <= sixty_days_out
    ).order_by(Deadline.due_date).all()
    upcoming_by_client = group_deadlines_by_client(upcoming)

    # Include status Overdue (set by the daily reminder job) as well as
    # Pending past-due items that have not been flipped yet.
    overdue = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status.in_(['Pending', 'Overdue']),
        Deadline.due_date < today
    ).order_by(Deadline.due_date).all()

    upcoming_invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status.in_(['Pending', 'Overdue']),
        Invoice.due_date >= today
    ).order_by(Invoice.due_date).limit(10).all()

    overdue_invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status.in_(['Pending', 'Overdue']),
        Invoice.due_date < today
    ).order_by(Invoice.due_date).all()

    clients_count = Client.query.filter_by(archived=False).count()
    pending_deadlines = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status.in_(['Pending', 'Overdue'])
    ).count()
    pending_invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status.in_(['Pending', 'Overdue'])
    ).count()

    return render_template('dashboard.html', 
                           upcoming=upcoming,
                           upcoming_by_client=upcoming_by_client,
                           overdue=overdue,
                           upcoming_invoices=upcoming_invoices,
                           overdue_invoices=overdue_invoices,
                           clients_count=clients_count,
                           pending_count=pending_deadlines + pending_invoices)

CLIENT_EXPORT_HEADERS = [
    'Status', 'Client Name', 'UEN', 'Contact Person', 'Email', 'Phone',
    'FYE Month', 'Incorporation Date', 'Created Date',
    'Pending Deadlines', 'Pending Invoices', 'Notes',
]

CLIENT_EXPORT_WRAP_COLUMNS = {10, 11, 12}

def format_pending_deadlines_export(client):
    open_deadlines = sorted(
        (dl for dl in client.deadlines if dl.status in ('Pending', 'Overdue')),
        key=lambda dl: dl.due_date,
    )
    if not open_deadlines:
        return ''
    lines = []
    for dl in open_deadlines:
        label = DEADLINE_TYPE_LABELS.get(dl.deadline_type, dl.deadline_type)
        line = f'{label} - Due: {format_date_display(dl.due_date)} ({dl.status})'
        if dl.comments:
            line += f' | {dl.comments}'
        lines.append(line)
    return '\n'.join(lines)

def format_pending_invoices_export(client):
    open_invoices = sorted(
        (inv for inv in client.invoices if inv.status in ('Pending', 'Overdue')),
        key=lambda inv: inv.due_date,
    )
    if not open_invoices:
        return ''
    lines = []
    for inv in open_invoices:
        amount = f'${inv.amount:,.2f}' if inv.amount is not None else 'No amount'
        description = inv.description or 'No description'
        lines.append(
            f'{inv.invoice_type} - {description} - {amount} - Due: {format_date_display(inv.due_date)} ({inv.status})'
        )
    return '\n'.join(lines)

def format_notes_export(client):
    notes = sorted(client.notes, key=lambda note: note.timestamp, reverse=True)
    if not notes:
        return ''
    lines = []
    for note in notes:
        line = f'[{note.category}] {note.title} ({format_date_display(note.timestamp, include_time=True)})'
        if note.content:
            line += f': {note.content}'
        if note.attachments:
            names = ', '.join(a.original_filename for a in note.attachments)
            line += f' [files: {names}]'
        lines.append(line)
    return '\n'.join(lines)

def client_export_row(client):
    return [
        'Archived' if client.archived else 'Active',
        client.name,
        client.uen,
        client.contact_person or '',
        client.email or '',
        client.phone or '',
        MONTH_NAMES[client.fye_month - 1] if client.fye_month else '',
        format_date_display(client.incorporation_date),
        format_date_display(client.created_at, include_time=True),
        format_pending_deadlines_export(client),
        format_pending_invoices_export(client),
        format_notes_export(client),
    ]

def build_clients_workbook(clients):
    wb = Workbook()
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill('solid', start_color='0D3B66')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    active_clients = [c for c in clients if not c.archived]
    archived_clients = [c for c in clients if c.archived]

    sheets = [
        ('Active Clients', active_clients),
        ('Archived Clients', archived_clients),
        ('All Clients', clients),
    ]

    first_sheet = True
    for sheet_name, sheet_clients in sheets:
        if first_sheet:
            ws = wb.active
            ws.title = sheet_name
            first_sheet = False
        else:
            ws = wb.create_sheet(sheet_name)

        ws.append(CLIENT_EXPORT_HEADERS)
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        wrap_alignment = Alignment(wrap_text=True, vertical='top')
        for client in sheet_clients:
            ws.append(client_export_row(client))

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if cell.column in CLIENT_EXPORT_WRAP_COLUMNS:
                    cell.alignment = wrap_alignment

        for column_cells in ws.columns:
            column_letter = column_cells[0].column_letter
            column_index = column_cells[0].column
            if column_index in CLIENT_EXPORT_WRAP_COLUMNS:
                ws.column_dimensions[column_letter].width = 50
                continue
            max_length = max((len(str(cell.value)) for cell in column_cells if cell.value is not None), default=0)
            ws.column_dimensions[column_letter].width = min(max_length + 2, 40)

        ws.freeze_panes = 'A2'

    return wb

@app.route('/clients/export')
@login_required
def clients_export():
    clients = Client.query.order_by(Client.archived, Client.name).all()
    wb = build_clients_workbook(clients)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f'astutecorp-clients-{date.today().strftime("%Y%m%d")}.xlsx'
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )

@app.route('/clients')
@login_required
def clients_list():
    show_archived = request.args.get('view') == 'archived'
    query = Client.query.filter_by(archived=show_archived)
    search = request.args.get('q', '').strip()
    if search:
        pattern = f'%{search}%'
        query = query.filter(
            db.or_(
                Client.name.ilike(pattern),
                Client.uen.ilike(pattern),
                Client.contact_person.ilike(pattern),
                Client.email.ilike(pattern),
                Client.phone.ilike(pattern),
            )
        )
    clients = query.order_by(Client.name).all()
    return render_template(
        'clients/list.html',
        clients=clients,
        show_archived=show_archived,
        search=search,
    )

@app.route('/clients/new', methods=['GET', 'POST'])
@login_required
def client_new():
    form = ClientForm()
    if form.validate_on_submit():
        client = Client(
            name=form.name.data,
            uen=form.uen.data,
            contact_person=form.contact_person.data,
            email=form.email.data,
            phone=form.phone.data,
            fye_month=form.fye_month.data,
            incorporation_date=form.incorporation_date.data
        )
        db.session.add(client)
        db.session.commit()
        flash('Client created successfully!', 'success')
        return redirect(url_for('client_detail', client_id=client.id))
    return render_template('clients/form.html', form=form, title='New Client')

@app.route('/clients/<int:client_id>')
@login_required
def client_detail(client_id):
    client = Client.query.get_or_404(client_id)
    deadlines = Deadline.query.filter_by(client_id=client_id).order_by(Deadline.due_date).all()
    notes = Note.query.filter_by(client_id=client_id).order_by(Note.timestamp.desc()).all()
    note_form = NoteForm()
    invoice_form = InvoiceForm()
    return render_template('clients/detail.html', client=client, deadlines=deadlines, invoices=client.invoices, notes=notes, note_form=note_form, invoice_form=invoice_form)

@app.route('/clients/<int:client_id>/edit', methods=['GET', 'POST'])
@login_required
def client_edit(client_id):
    client = Client.query.get_or_404(client_id)
    form = ClientForm(obj=client)
    if form.validate_on_submit():
        form.populate_obj(client)
        db.session.commit()
        flash('Client updated.', 'success')
        return redirect(url_for('client_detail', client_id=client.id))
    return render_template('clients/form.html', form=form, title='Edit Client')

@app.route('/clients/<int:client_id>/send-reminder', methods=['GET', 'POST'])
@login_required
def client_send_reminder(client_id):
    """Compose and send a reminder email for this client only (not all clients)."""
    client = Client.query.get_or_404(client_id)
    settings = Setting.query.first()
    deadlines = get_client_open_deadlines(client_id)
    form = ClientReminderForm()
    client_emails = client.email_list()

    if request.method == 'GET':
        # Reuse last saved message for this client if available
        if (client.saved_reminder_subject or '').strip() and (client.saved_reminder_body or '').strip():
            form.subject.data = client.saved_reminder_subject
            form.body.data = client.saved_reminder_body
        else:
            subject, body, _summary = build_client_reminder_defaults(client, deadlines, settings)
            form.subject.data = subject
            form.body.data = body
        form.send_to_client.data = 'True' if client_emails else 'False'
        form.send_to_me.data = 'True'

    if form.validate_on_submit():
        to_me = form.send_to_me.data == 'True'
        to_client = form.send_to_client.data == 'True'
        if not to_me and not to_client:
            flash('Choose at least one recipient: you and/or the client.', 'warning')
            return render_template(
                'clients/send_reminder.html',
                form=form,
                client=client,
                deadlines=deadlines,
            )
        if not settings or not settings.my_email:
            flash('Configure Your Email and SMTP in Settings before sending.', 'danger')
            return redirect(url_for('settings'))

        subject = form.subject.data.strip()
        body = form.body.data.strip()

        # Always keep this client's composed message for next time
        client.saved_reminder_subject = subject
        client.saved_reminder_body = body
        db.session.commit()

        sent = []
        failed = []

        if to_me:
            if send_email(settings.my_email, f'[AstuteCorp CRM] {subject}', body):
                sent.append(f'you ({settings.my_email})')
            else:
                failed.append(f'you ({_last_email_error or "error"})')

        if to_client:
            if not client_emails:
                failed.append('client (no email on file)')
            else:
                ok, bad = send_email_to_each(client_emails, subject, body)
                for e in ok:
                    sent.append(f'client ({e})')
                for e in bad:
                    failed.append(f'client ({e}: {_last_email_error or "error"})')

        if sent:
            today = date.today()
            for dl in deadlines:
                dl.last_reminded = today

            # Title + timestamp only in Activity history (no full email body)
            note_title = f'Email reminder sent: {subject[:160]}'
            if len(subject) > 160:
                note_title = f'Email reminder sent: {subject[:157]}...'
            db.session.add(Note(
                client_id=client.id,
                title=note_title,
                content='',  # keep Activity list clean; full text stays in saved draft
                category='Client',
            ))
            db.session.commit()
            flash(
                f'Reminder sent to: {", ".join(sent)}. Message saved for this client and logged in Activity & Notes.',
                'success',
            )
        if failed:
            flash(f'Could not send to: {", ".join(failed)}.', 'danger')
        if sent:
            return redirect(url_for('client_detail', client_id=client_id))

    return render_template(
        'clients/send_reminder.html',
        form=form,
        client=client,
        deadlines=deadlines,
        has_saved_message=bool(
            (client.saved_reminder_subject or '').strip()
            and (client.saved_reminder_body or '').strip()
        ),
    )


@app.route('/clients/<int:client_id>/send-letters', methods=['GET', 'POST'])
@login_required
def client_send_letters(client_id):
    """Notify client that letters/mail are ready for collection (registered address service)."""
    client = Client.query.get_or_404(client_id)
    settings = Setting.query.first()
    form = ClientLettersForm()
    client_emails = client.email_list()

    if request.method == 'GET':
        subject, body = build_letters_notice_defaults(client)
        form.subject.data = subject
        form.body.data = body
        form.send_to_client.data = 'True' if client_emails else 'False'
        form.send_to_me.data = 'True'

    if form.validate_on_submit():
        to_me = form.send_to_me.data == 'True'
        to_client = form.send_to_client.data == 'True'
        if not to_me and not to_client:
            flash('Choose at least one recipient: you and/or the client.', 'warning')
            return render_template('clients/send_letters.html', form=form, client=client)
        if not settings or not settings.my_email:
            flash('Configure Your Email and SMTP in Settings before sending.', 'danger')
            return redirect(url_for('settings'))

        subject = form.subject.data.strip()
        body = form.body.data.strip()
        sent = []
        failed = []

        if to_me:
            if send_email(settings.my_email, f'[AstuteCorp CRM] {subject}', body):
                sent.append(f'you ({settings.my_email})')
            else:
                failed.append(f'you ({_last_email_error or "error"})')

        if to_client:
            if not client_emails:
                failed.append('client (no email on file)')
            else:
                ok, bad = send_email_to_each(client_emails, subject, body)
                for e in ok:
                    sent.append(f'client ({e})')
                for e in bad:
                    failed.append(f'client ({e}: {_last_email_error or "error"})')

        if sent:
            note_title = f'Letters collection notice sent: {subject[:140]}'
            if len(subject) > 140:
                note_title = f'Letters collection notice sent: {subject[:137]}...'
            db.session.add(Note(
                client_id=client.id,
                title=note_title,
                content='',
                category='Client',
            ))
            db.session.commit()
            flash(
                f'Letters notice sent to: {", ".join(sent)}. Logged in Activity & Notes.',
                'success',
            )
        if failed:
            flash(f'Could not send to: {", ".join(failed)}.', 'danger')
        if sent:
            return redirect(url_for('client_detail', client_id=client_id))

    return render_template('clients/send_letters.html', form=form, client=client)


@app.route('/clients/<int:client_id>/delete', methods=['POST'])
@login_required
def client_delete(client_id):
    client = Client.query.get_or_404(client_id)
    name = client.name
    db.session.delete(client)
    db.session.commit()
    flash(f'Client "{name}" and all related records have been deleted.', 'success')
    return redirect(url_for('clients_list'))

@app.route('/clients/<int:client_id>/archive', methods=['POST'])
@login_required
def client_archive(client_id):
    client = Client.query.get_or_404(client_id)
    client.archived = True
    db.session.commit()
    flash(f'Client "{client.name}" has been archived.', 'success')
    return redirect(request.referrer or url_for('clients_list'))

@app.route('/clients/<int:client_id>/unarchive', methods=['POST'])
@login_required
def client_unarchive(client_id):
    client = Client.query.get_or_404(client_id)
    client.archived = False
    db.session.commit()
    flash(f'Client "{client.name}" has been restored to active clients.', 'success')
    return redirect(request.referrer or url_for('clients_list', view='archived'))

@app.route('/clients/<int:client_id>/deadlines/new', methods=['GET', 'POST'])
@login_required
def deadline_new(client_id):
    client = Client.query.get_or_404(client_id)
    form = DeadlineForm()
    if form.validate_on_submit():
        dl = Deadline(
            client_id=client_id,
            deadline_type=form.deadline_type.data,
            due_date=form.due_date.data,
            reminder_days=form.reminder_days.data,
            comments=form.comments.data
        )
        db.session.add(dl)
        db.session.commit()
        flash('Deadline added.', 'success')
        return redirect(url_for('client_detail', client_id=client_id))
    return render_template('deadlines/form.html', form=form, client=client, title='Add Deadline')

@app.route('/clients/<int:client_id>/deadlines/<int:dl_id>/edit', methods=['GET', 'POST'])
@login_required
def deadline_edit(client_id, dl_id):
    client = Client.query.get_or_404(client_id)
    dl = Deadline.query.filter_by(id=dl_id, client_id=client_id).first_or_404()
    form = DeadlineForm(obj=dl)
    form.submit.label.text = 'Save Changes'

    if form.validate_on_submit():
        dl.deadline_type = form.deadline_type.data
        dl.due_date = form.due_date.data
        dl.reminder_days = form.reminder_days.data
        dl.comments = form.comments.data
        # Moving an Overdue deadline to a future date makes it Pending again
        if dl.status == 'Overdue' and dl.due_date >= date.today():
            dl.status = 'Pending'
            dl.last_reminded = None
        db.session.commit()
        flash('Deadline updated successfully.', 'success')
        return redirect(url_for('client_detail', client_id=client_id))

    return render_template('deadlines/form.html', form=form, client=client, title='Edit Deadline')

@app.route('/clients/<int:client_id>/deadlines/<int:dl_id>/complete', methods=['POST'])
@login_required
def deadline_complete(client_id, dl_id):
    dl = Deadline.query.get_or_404(dl_id)
    dl.status = 'Completed'
    dl.completed_date = date.today()
    db.session.commit()
    flash('Deadline marked as completed.', 'success')
    return redirect(url_for('client_detail', client_id=client_id))

@app.route('/clients/<int:client_id>/deadlines/<int:dl_id>/delete', methods=['POST'])
@login_required
def deadline_delete(client_id, dl_id):
    dl = Deadline.query.filter_by(id=dl_id, client_id=client_id).first_or_404()
    deadline_name = DEADLINE_TYPE_LABELS.get(dl.deadline_type, dl.deadline_type)
    db.session.delete(dl)
    db.session.commit()
    flash(f'Deadline "{deadline_name}" deleted.', 'success')
    return redirect(url_for('client_detail', client_id=client_id))

@app.route('/clients/<int:client_id>/notes/new', methods=['POST'])
@login_required
def note_new(client_id):
    Client.query.get_or_404(client_id)
    form = NoteForm()
    if form.validate_on_submit():
        note = Note(
            client_id=client_id,
            title=form.title.data,
            content=form.content.data,
            category=form.category.data
        )
        db.session.add(note)
        db.session.flush()  # get note.id for upload folder

        files = request.files.getlist('attachments')
        saved, err = save_note_attachments(note, files)
        if err:
            db.session.rollback()
            flash(err, 'danger')
            return redirect(url_for('client_detail', client_id=client_id))

        db.session.commit()
        if saved:
            flash(f'Note added with {saved} attachment{"s" if saved != 1 else ""}.', 'success')
        else:
            flash('Note added successfully.', 'success')
    else:
        for field, errors in form.errors.items():
            for error in errors:
                flash(f'{error}', 'danger')
                break
    return redirect(url_for('client_detail', client_id=client_id))


@app.route('/clients/<int:client_id>/notes/attachments/<int:attachment_id>/download')
@login_required
def note_attachment_download(client_id, attachment_id):
    Client.query.get_or_404(client_id)
    attachment = NoteAttachment.query.get_or_404(attachment_id)
    note = attachment.note
    if not note or note.client_id != client_id:
        abort(404)
    path = note_attachment_path(attachment)
    if not os.path.isfile(path):
        flash('Attachment file is missing on the server.', 'danger')
        return redirect(url_for('client_detail', client_id=client_id))
    return send_file(
        path,
        as_attachment=True,
        download_name=attachment.original_filename,
    )

@app.route('/clients/<int:client_id>/calculate-deadlines', methods=['POST'])
@login_required
def calculate_deadlines(client_id):
    Client.query.get_or_404(client_id)
    fye_deadlines = get_fye_reference_deadlines(client_id)
    if not fye_deadlines:
        flash('Please add a Financial Year End deadline first.', 'warning')
        return redirect(url_for('client_detail', client_id=client_id))

    added = 0
    for fye_deadline in fye_deadlines:
        suggestions = calculate_default_deadlines(fye_deadline.due_date)
        for dtype, ddate in suggestions:
            # Within a week counts as existing (older calculations landed 1-3 days before month-end)
            existing = Deadline.query.filter(
                Deadline.client_id == client_id,
                Deadline.deadline_type == dtype,
                Deadline.due_date.between(ddate - timedelta(days=7), ddate + timedelta(days=7))
            ).first()
            if not existing:
                dl = Deadline(client_id=client_id, deadline_type=dtype, due_date=ddate)
                db.session.add(dl)
                added += 1

    db.session.commit()
    if added:
        flash('Suggested deadlines added from Financial Year End date(s) (you can edit them).', 'success')
    else:
        flash('All suggested deadlines already exist for the Financial Year End date(s).', 'info')
    return redirect(url_for('client_detail', client_id=client_id))

@app.route('/clients/<int:client_id>/invoices/new', methods=['GET', 'POST'])
@login_required
def invoice_new(client_id):
    client = Client.query.get_or_404(client_id)
    form = InvoiceForm()
    if form.validate_on_submit():
        try:
            amount = float(form.amount.data) if form.amount.data else None
        except:
            amount = None

        inv = Invoice(
            client_id=client_id,
            invoice_type=form.invoice_type.data,
            description=form.description.data,
            amount=amount,
            due_date=form.due_date.data
        )
        db.session.add(inv)
        db.session.commit()
        flash('Invoice added successfully.', 'success')
        return redirect(url_for('client_detail', client_id=client_id))
    return render_template('invoices/form.html', form=form, client=client, title='Add Invoice')

@app.route('/clients/<int:client_id>/invoices/<int:inv_id>/edit', methods=['GET', 'POST'])
@login_required
def invoice_edit(client_id, inv_id):
    client = Client.query.get_or_404(client_id)
    inv = Invoice.query.filter_by(id=inv_id, client_id=client_id).first_or_404()
    form = InvoiceForm(obj=inv)
    if inv.amount is not None:
        form.amount.data = f'{inv.amount:g}'
    form.submit.label.text = 'Save Changes'

    if form.validate_on_submit():
        try:
            amount = float(form.amount.data) if form.amount.data else None
        except ValueError:
            amount = None

        inv.invoice_type = form.invoice_type.data
        inv.description = form.description.data
        inv.amount = amount
        inv.due_date = form.due_date.data
        # Moving an Overdue invoice to today or later makes it Pending again
        if inv.status == 'Overdue' and inv.due_date >= date.today():
            inv.status = 'Pending'
        db.session.commit()
        flash('Invoice updated successfully.', 'success')
        return redirect(url_for('client_detail', client_id=client_id))

    return render_template('invoices/form.html', form=form, client=client, title='Edit Invoice')

@app.route('/clients/<int:client_id>/invoices/<int:inv_id>/paid', methods=['POST'])
@login_required
def invoice_paid(client_id, inv_id):
    inv = Invoice.query.get_or_404(inv_id)
    inv.status = 'Paid'
    inv.paid_date = date.today()
    db.session.commit()
    flash('Invoice marked as paid.', 'success')
    return redirect(url_for('client_detail', client_id=client_id))

@app.route('/send-reminders', methods=['POST'])
@login_required
def send_reminders():
    count = check_and_send_reminders()
    if count:
        flash(f'{count} deadline reminder email(s) sent to your inbox.', 'success')
    else:
        flash(
            'No new reminders to send right now. '
            'Emails go out when a deadline is within its reminder window, due today, or overdue '
            '(once per day). Automatic check runs daily at 9:00 AM SGT.',
            'info'
        )
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    setting = Setting.query.first()
    if not setting:
        setting = Setting()
        db.session.add(setting)
        db.session.commit()

    if not setting.internal_reminder_subject_template:
        setting.internal_reminder_subject_template = DEFAULT_INTERNAL_REMINDER_SUBJECT
    if not setting.internal_reminder_body_template:
        setting.internal_reminder_body_template = DEFAULT_INTERNAL_REMINDER_BODY
    if not setting.reminder_subject_template:
        setting.reminder_subject_template = DEFAULT_REMINDER_SUBJECT
    if not setting.reminder_body_template:
        setting.reminder_body_template = DEFAULT_REMINDER_BODY

    form = SettingsForm(obj=setting)
    if form.validate_on_submit():
        setting.my_email = form.my_email.data
        setting.smtp_server = form.smtp_server.data
        setting.smtp_port = form.smtp_port.data
        setting.smtp_username = form.smtp_username.data
        if form.smtp_password.data:
            setting.smtp_password = form.smtp_password.data
        setting.use_tls = (form.use_tls.data == 'True')
        setting.email_client = (form.email_client.data == 'True')
        setting.internal_reminder_subject_template = form.internal_reminder_subject_template.data
        setting.internal_reminder_body_template = form.internal_reminder_body_template.data
        setting.reminder_subject_template = form.reminder_subject_template.data
        setting.reminder_body_template = form.reminder_body_template.data
        db.session.commit()
        flash('Settings saved.', 'success')
        return redirect(url_for('settings'))
    form.use_tls.data = 'True' if setting.use_tls else 'False'
    form.email_client.data = 'True' if setting.email_client else 'False'
    return render_template(
        'settings.html',
        form=form,
        placeholders=REMINDER_TEMPLATE_PLACEHOLDERS,
        default_internal_subject=DEFAULT_INTERNAL_REMINDER_SUBJECT,
        default_internal_body=DEFAULT_INTERNAL_REMINDER_BODY,
        default_subject=DEFAULT_REMINDER_SUBJECT,
        default_body=DEFAULT_REMINDER_BODY
    )

@app.route('/settings/test-email', methods=['POST'])
@login_required
def test_email():
    success, message = send_test_reminder_email()
    flash(message, 'success' if success else 'danger')
    return redirect(url_for('settings'))


def _validate_sqlite_upload(file_storage):
    """Save upload to a temp path and verify it is a readable CRM SQLite DB."""
    if not file_storage or not file_storage.filename:
        return None, 'No database file uploaded.'
    tmp_path = os.path.join(app.instance_path, f'restore-{uuid.uuid4().hex}.db')
    file_storage.save(tmp_path)
    try:
        conn = sqlite3_stdlib.connect(tmp_path)
        try:
            tables = {
                row[0]
                for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
            required = {'client', 'deadline', 'user'}
            if not required.issubset(tables):
                return None, 'File is not a valid AstuteCorp CRM database (missing tables).'
            client_count = conn.execute('SELECT COUNT(*) FROM client').fetchone()[0]
        finally:
            conn.close()
    except Exception as exc:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        return None, f'Invalid SQLite file: {exc}'
    return tmp_path, f'OK ({client_count} clients)'


def replace_sqlite_database(source_path):
    """Replace the live SQLite file with source_path. Returns (ok, message)."""
    dest = get_sqlite_filesystem_path()
    if not dest:
        return False, 'Database path is not SQLite; cannot import.'
    targets = [dest]
    # Also write to persistent disk when available (Render)
    if os.path.isdir('/var/data'):
        persistent = '/var/data/corp_crm.db'
        if persistent not in targets:
            targets.append(persistent)
    # Drop active connections before swapping the file
    db.session.remove()
    db.engine.dispose()
    written = []
    for target in targets:
        os.makedirs(os.path.dirname(target) or '.', exist_ok=True)
        if os.path.isfile(target):
            backup = target + f'.bak-{datetime.utcnow().strftime("%Y%m%d%H%M%S")}'
            try:
                shutil.copy2(target, backup)
            except OSError:
                pass
        shutil.copy2(source_path, target)
        written.append(target)
    # Re-bind and ensure schema
    db.engine.dispose()
    with app.app_context():
        db.create_all()
        migrate_database()
    return True, written[0]


@app.route('/settings/export-database')
@login_required
def export_database():
    path = get_sqlite_filesystem_path()
    if not path or not os.path.isfile(path):
        flash('Database file not found on this server.', 'danger')
        return redirect(url_for('settings'))
    return send_file(
        path,
        as_attachment=True,
        download_name=f'corp_crm-backup-{date.today().isoformat()}.db',
    )


@app.route('/settings/import-database', methods=['POST'])
@login_required
def import_database():
    tmp_path, msg = _validate_sqlite_upload(request.files.get('database'))
    if not tmp_path:
        flash(msg, 'danger')
        return redirect(url_for('settings'))
    try:
        ok, result = replace_sqlite_database(tmp_path)
        if not ok:
            flash(result, 'danger')
        else:
            flash(
                f'Database imported successfully ({msg}). '
                'If anything looks stale, restart the service on Render.',
                'success',
            )
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
    return redirect(url_for('settings'))


@app.route('/internal/restore-database', methods=['POST'])
def internal_restore_database():
    """One-time / emergency DB restore via token (no login). Disable after use."""
    token = request.headers.get('X-Restore-Token') or request.form.get('token') or ''
    if not _ONE_TIME_RESTORE_TOKEN or token != _ONE_TIME_RESTORE_TOKEN:
        abort(403)
    tmp_path, msg = _validate_sqlite_upload(request.files.get('database'))
    if not tmp_path:
        return {'ok': False, 'error': msg}, 400
    try:
        ok, result = replace_sqlite_database(tmp_path)
        if not ok:
            return {'ok': False, 'error': result}, 500
        # Report counts from new DB
        conn = sqlite3_stdlib.connect(result)
        try:
            clients = conn.execute('SELECT COUNT(*) FROM client').fetchone()[0]
            deadlines = conn.execute('SELECT COUNT(*) FROM deadline').fetchone()[0]
        finally:
            conn.close()
        return {
            'ok': True,
            'path': result,
            'clients': clients,
            'deadlines': deadlines,
            'message': msg,
        }
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()
        if user and user.check_password(form.password.data):
            login_user(user)
            flash('Logged in successfully.', 'success')
            next_page = request.args.get('next')
            return redirect(next_page or url_for('dashboard'))
        else:
            flash('Invalid username or password.', 'danger')
    return render_template('login.html', form=form)


@app.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    form = ForgotPasswordForm()
    if form.validate_on_submit():
        username = (form.username.data or '').strip()
        user = User.query.filter_by(username=username).first()
        settings = Setting.query.first()
        # Always show the same message (do not reveal whether username exists)
        generic_msg = (
            'If that username exists and email is configured, a reset link has been sent '
            'to the CRM notification email in Settings. Check that inbox (and spam).'
        )
        if not user:
            flash(generic_msg, 'info')
            return redirect(url_for('login'))
        if not settings or not settings.my_email:
            flash(
                'Password reset email cannot be sent: configure Your Email and SMTP in Settings first '
                '(ask a logged-in teammate, or restore access via database backup).',
                'danger',
            )
            return redirect(url_for('login'))
        token = generate_password_reset_token(user.id)
        reset_url = url_for('reset_password', token=token, _external=True)
        subject = '[AstuteCorp CRM] Password reset'
        body = (
            f'A password reset was requested for username: {user.username}\n\n'
            f'Open this link within 1 hour to set a new password:\n{reset_url}\n\n'
            f'If you did not request this, ignore this email.\n\n'
            f'— AstuteCorp CRM'
        )
        if send_email(settings.my_email, subject, body):
            flash(generic_msg, 'info')
        else:
            err = _last_email_error or 'Unknown email error'
            flash(f'Could not send reset email: {err}', 'danger')
        return redirect(url_for('login'))
    return render_template('forgot_password.html', form=form)


@app.route('/reset-password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    user_id = verify_password_reset_token(token)
    if not user_id:
        flash('This reset link is invalid or has expired. Please request a new one.', 'danger')
        return redirect(url_for('forgot_password'))
    user = User.query.get(user_id)
    if not user:
        flash('User account not found.', 'danger')
        return redirect(url_for('login'))
    form = ResetPasswordForm()
    if form.validate_on_submit():
        user.set_password(form.new_password.data)
        db.session.commit()
        flash('Password updated. You can log in with your new password.', 'success')
        return redirect(url_for('login'))
    return render_template('reset_password.html', form=form, username=user.username)


@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))

@app.route('/users', methods=['GET', 'POST'])
@login_required
def users():
    form = AddUserForm()
    if form.validate_on_submit():
        if User.query.filter_by(username=form.username.data).first():
            flash('That username is already taken.', 'danger')
        else:
            user = User(username=form.username.data)
            user.set_password(form.password.data)
            db.session.add(user)
            db.session.commit()
            flash(f'User "{form.username.data}" created.', 'success')
            return redirect(url_for('users'))
    return render_template('users.html', users=User.query.order_by(User.username).all(), form=form)


@app.route('/users/<int:user_id>/set-password', methods=['GET', 'POST'])
@login_required
def set_user_password(user_id):
    """Set a user's password without knowing the old one (for admins / recovery)."""
    user = User.query.get_or_404(user_id)
    form = AdminSetPasswordForm()
    if form.validate_on_submit():
        user.set_password(form.new_password.data)
        db.session.commit()
        flash(f'Password updated for "{user.username}".', 'success')
        return redirect(url_for('users'))
    return render_template('set_user_password.html', form=form, user=user)

@app.route('/users/<int:user_id>/delete', methods=['POST'])
@login_required
def delete_user(user_id):
    user = User.query.get_or_404(user_id)
    if user.id == current_user.id:
        flash('You cannot delete your own account while logged in.', 'danger')
        return redirect(url_for('users'))
    if User.query.count() <= 1:
        flash('You must keep at least one user account.', 'danger')
        return redirect(url_for('users'))
    username = user.username
    db.session.delete(user)
    db.session.commit()
    flash(f'User "{username}" deleted.', 'success')
    return redirect(url_for('users'))

@app.route('/change-password', methods=['GET', 'POST'])
@login_required
def change_password():
    form = ChangePasswordForm()
    if form.validate_on_submit():
        if not current_user.check_password(form.current_password.data):
            flash('Current password is incorrect.', 'danger')
        else:
            current_user.set_password(form.new_password.data)
            db.session.commit()
            flash('Password changed successfully.', 'success')
            return redirect(url_for('dashboard'))
    return render_template('change_password.html', form=form)

# ==================== INITIAL SETUP ====================

def migrate_database():
    """Add new columns to existing SQLite databases."""
    from sqlalchemy import inspect, text
    inspector = inspect(db.engine)
    table_columns = {
        'setting': {
            'reminder_subject_template': 'TEXT',
            'reminder_body_template': 'TEXT',
            'email_client': 'BOOLEAN DEFAULT 0',
            'internal_reminder_subject_template': 'TEXT',
            'internal_reminder_body_template': 'TEXT',
        },
        'deadline': {
            'comments': 'TEXT',
        },
        'client': {
            'archived': 'BOOLEAN DEFAULT 0',
            'saved_reminder_subject': 'TEXT',
            'saved_reminder_body': 'TEXT',
        },
    }
    with db.engine.begin() as conn:
        for table_name, new_columns in table_columns.items():
            if table_name not in inspector.get_table_names():
                continue
            columns = {col['name'] for col in inspector.get_columns(table_name)}
            for column_name, column_type in new_columns.items():
                if column_name not in columns:
                    conn.execute(text(f'ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}'))

        marker_path = os.path.join(app.instance_path, '.reminder_days_defaulted_to_1')
        if not os.path.exists(marker_path):
            table_names = inspector.get_table_names()
            if 'deadline' in table_names:
                conn.execute(text('UPDATE deadline SET reminder_days = 1'))
            if 'invoice' in table_names:
                conn.execute(text('UPDATE invoice SET reminder_days = 1'))
            os.makedirs(app.instance_path, exist_ok=True)
            with open(marker_path, 'w') as marker_file:
                marker_file.write('done')

def create_default_settings():
    if Setting.query.count() == 0:
        s = Setting(
            my_email='kelvin@astutecorp-s.com',
            smtp_server='smtppro.zoho.com',
            smtp_port=587,
            smtp_username='kelvin@astutecorp-s.com',
            smtp_password='',
            use_tls=True,
            internal_reminder_subject_template=DEFAULT_INTERNAL_REMINDER_SUBJECT,
            internal_reminder_body_template=DEFAULT_INTERNAL_REMINDER_BODY,
            reminder_subject_template=DEFAULT_REMINDER_SUBJECT,
            reminder_body_template=DEFAULT_REMINDER_BODY,
        )
        db.session.add(s)
        db.session.commit()
        print("Default settings created. Please update in /settings")
    else:
        setting = Setting.query.first()
        updated = False
        if not setting.internal_reminder_subject_template:
            setting.internal_reminder_subject_template = DEFAULT_INTERNAL_REMINDER_SUBJECT
            updated = True
        if not setting.internal_reminder_body_template:
            setting.internal_reminder_body_template = DEFAULT_INTERNAL_REMINDER_BODY
            updated = True
        if not setting.reminder_subject_template:
            setting.reminder_subject_template = DEFAULT_REMINDER_SUBJECT
            updated = True
        if not setting.reminder_body_template:
            setting.reminder_body_template = DEFAULT_REMINDER_BODY
            updated = True
        if updated:
            db.session.commit()

    # Create a demo client if none exist
    if Client.query.count() == 0:
        demo = Client(
            name="Demo Pte. Ltd.",
            uen="202612345K",
            contact_person="Ms. Tan",
            email="client@example.com",
            phone="+65 9123 4567",
            fye_month=12
        )
        db.session.add(demo)
        db.session.commit()
        print("Demo client created.")

    # Create default admin user if none exist
    if User.query.count() == 0:
        admin = User(username='admin')
        admin.set_password('admin123')
        db.session.add(admin)
        db.session.commit()
        print("Default admin user created: username=admin, password=admin123")

# Background scheduler for reminders (9:00 AM Singapore time daily)
SGT = ZoneInfo('Asia/Singapore')
scheduler = BackgroundScheduler(timezone=SGT)
scheduler.add_job(func=check_and_send_reminders, trigger="cron", hour=9, minute=0)

def start_scheduler():
    """Start the reminder scheduler once (skip Flask debug reloader parent process)."""
    if scheduler.running or os.environ.get('DISABLE_SCHEDULER') == '1':
        return
    is_reloader_parent = app.debug and os.environ.get('WERKZEUG_RUN_MAIN') is None
    if not is_reloader_parent:
        scheduler.start()

start_scheduler()

@app.before_request
def before_first_request():
    db.create_all()
    migrate_database()
    create_default_settings()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('FLASK_DEBUG', '1') == '1'
    print(f"Starting AstuteCorp CRM on port {port}...")
    app.run(debug=debug, host='0.0.0.0', port=port)
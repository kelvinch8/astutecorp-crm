from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from flask_sqlalchemy import SQLAlchemy
from flask_wtf import FlaskForm
from wtforms import StringField, IntegerField, DateField, TextAreaField, SelectField, SubmitField, PasswordField
from wtforms.validators import DataRequired, Email, Optional, NumberRange, Length, EqualTo
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__, instance_relative_config=True)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'corp-crm-super-secret-key-change-me')
os.makedirs(app.instance_path, exist_ok=True)
_default_db = 'sqlite:///' + os.path.join(app.instance_path, 'corp_crm.db')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', _default_db)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

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
    email = db.Column(db.String(120))
    phone = db.Column(db.String(30))
    fye_month = db.Column(db.Integer)  # 1-12
    incorporation_date = db.Column(db.Date)
    archived = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    deadlines = db.relationship('Deadline', backref='client', lazy=True, cascade="all, delete-orphan")
    invoices = db.relationship('Invoice', backref='client', lazy=True, cascade="all, delete-orphan")
    notes = db.relationship('Note', backref='client', lazy=True, cascade="all, delete-orphan")

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
        return self.status == 'Pending' and self.due_date < date.today()

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
        return self.status == 'Pending' and self.due_date < date.today()

    def get_days_left(self):
        return (self.due_date - date.today()).days

class Note(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey('client.id'), nullable=False)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)
    title = db.Column(db.String(200), nullable=False)
    content = db.Column(db.Text)
    category = db.Column(db.String(50), default='General')  # ACRA, IRAS, Client, General

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

class ClientForm(FlaskForm):
    name = StringField('Client Name', validators=[DataRequired()])
    uen = StringField('UEN', validators=[DataRequired()])
    contact_person = StringField('Contact Person')
    email = StringField('Client Email', validators=[Optional(), Email()])
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
        ('TaxFiling', 'Income Tax Filing')
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
    submit = SubmitField('Add Note')

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

# ==================== HELPER FUNCTIONS ====================

DEADLINE_TYPES = ['FinancialYearEnd', 'ECI', 'AGM', 'AnnualReturn', 'TaxFiling']

DEADLINE_TYPE_LABELS = {
    'FinancialYearEnd': 'Financial Year End',
    'AGM': 'AGM Deadline',
    'AnnualReturn': 'Annual Return Filing',
    'ECI': 'ECI Filing',
    'TaxFiling': 'Income Tax Filing',
}

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
    ('{client_email}', 'Client email address'),
    ('{client_phone}', 'Client phone number'),
    ('{deadline_type}', 'Deadline type, e.g. AGM Deadline'),
    ('{due_date}', 'Due date in dd-mm-yyyy format'),
    ('{days_remaining}', 'Number of days until due date'),
    ('{comments}', 'Deadline comments/notes'),
]

@app.template_filter('deadline_label')
def deadline_label(deadline_type):
    return DEADLINE_TYPE_LABELS.get(deadline_type, deadline_type)

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

def build_deadline_template_values(deadline):
    client = deadline.client
    deadline_name = DEADLINE_TYPE_LABELS.get(deadline.deadline_type, deadline.deadline_type)
    return {
        'contact_person': client.contact_person or 'Client',
        'client_name': client.name,
        'uen': client.uen,
        'client_email': client.email or 'Not set',
        'client_phone': client.phone or 'Not set',
        'deadline_type': deadline_name,
        'due_date': format_date_display(deadline.due_date),
        'days_remaining': deadline.get_days_left(),
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

def calculate_default_deadlines(fye_date):
    """Suggest default deadlines based on an existing Financial Year End date."""
    if not fye_date:
        return []

    deadlines = []

    # ECI: 3 months after financial year end
    eci_due = fye_date + relativedelta(months=3)
    deadlines.append(('ECI', eci_due))

    # AGM: 6 months after financial year end
    agm_due = fye_date + relativedelta(months=6)
    deadlines.append(('AGM', agm_due))

    # Annual Return: assume 1 month after AGM (common practice)
    ar_due = agm_due + relativedelta(months=1)
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
    """Send email using settings from DB."""
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

    try:
        msg = MIMEMultipart()
        msg['From'] = settings.my_email
        msg['To'] = to_email
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
        print(f"✅ Email sent to {to_email}")
        return True
    except Exception as e:
        _last_email_error = str(e)
        print(f"❌ Email failed to {to_email}: {e}")
        return False

def send_deadline_reminder(deadline):
    """Email you about an upcoming deadline; optionally email the client too."""
    client = deadline.client
    settings = Setting.query.first()
    if not client or not settings:
        return False

    template_values = build_deadline_template_values(deadline)

    internal_subject_template, internal_body_template = get_internal_reminder_templates(settings)
    internal_subject = render_reminder_template(internal_subject_template, **template_values)
    internal_body = render_reminder_template(internal_body_template, **template_values)
    sent_to_self = send_email(settings.my_email, f"[AstuteCorp CRM] {internal_subject}", internal_body)

    sent_to_client = True
    if settings.email_client and client.email:
        client_subject_template, client_body_template = get_client_reminder_templates(settings)
        client_subject = render_reminder_template(client_subject_template, **template_values)
        client_body = render_reminder_template(client_body_template, **template_values)
        sent_to_client = send_email(client.email, client_subject, client_body)

    if sent_to_self or sent_to_client:
        deadline.last_reminded = date.today()
        db.session.commit()
        return True
    return False

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
    today = date.today()
    sent_count = 0

    deadlines = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status == 'Pending'
    ).all()
    for dl in deadlines:
        days_left = dl.get_days_left()
        if days_left <= 0:
            dl.status = 'Overdue'
            db.session.commit()
            continue

        if dl.reminder_days >= days_left and (not dl.last_reminded or dl.last_reminded < today):
            if send_deadline_reminder(dl):
                sent_count += 1

    invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status == 'Pending'
    ).all()
    for inv in invoices:
        if inv.get_days_left() <= 0:
            inv.status = 'Overdue'
            db.session.commit()

    print(f"Reminder check complete. Sent {sent_count} deadline reminders.")
    return sent_count

# ==================== ROUTES ====================

@app.route('/')
@login_required
def dashboard():
    today = date.today()
    sixty_days_out = today + timedelta(days=60)
    upcoming = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status == 'Pending',
        Deadline.due_date >= today,
        Deadline.due_date <= sixty_days_out
    ).order_by(Deadline.due_date).all()
    upcoming_by_client = group_deadlines_by_client(upcoming)

    overdue = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status == 'Pending',
        Deadline.due_date < today
    ).all()

    upcoming_invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status == 'Pending',
        Invoice.due_date >= today
    ).order_by(Invoice.due_date).limit(10).all()

    overdue_invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status == 'Pending',
        Invoice.due_date < today
    ).all()

    clients_count = Client.query.filter_by(archived=False).count()
    pending_deadlines = Deadline.query.join(Client).filter(
        Client.archived == False,
        Deadline.status == 'Pending'
    ).count()
    pending_invoices = Invoice.query.join(Client).filter(
        Client.archived == False,
        Invoice.status == 'Pending'
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
    form = NoteForm()
    if form.validate_on_submit():
        note = Note(
            client_id=client_id,
            title=form.title.data,
            content=form.content.data,
            category=form.category.data
        )
        db.session.add(note)
        db.session.commit()
        flash('Note added successfully.', 'success')
    return redirect(url_for('client_detail', client_id=client_id))

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
            existing = Deadline.query.filter_by(
                client_id=client_id,
                deadline_type=dtype,
                due_date=ddate
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
        flash('No reminders due right now. Emails are sent when a deadline is within its reminder window.', 'info')
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
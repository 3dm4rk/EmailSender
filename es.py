from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify
import getpass, os, sys, re, json
import smtplib
import gspread, time, datetime
from datetime import *
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import secrets
import threading

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Configuration
EMAIL_TEMPLATE_FILE = "EmailTemplate/template1.txt"
CREDENTIALS_FILE = "credentials.txt"
EXTRACTED_DATA_FILE = "extracted_data.json"
FILES_DIRECTORY = "Files"

# Initialize directories
os.makedirs("EmailTemplate", exist_ok=True)
os.makedirs(FILES_DIRECTORY, exist_ok=True)
os.makedirs("templates", exist_ok=True)

# Global variable to track email sending progress
email_progress = {
    'current': 0,
    'total': 0,
    'status': 'idle',
    'results': []
}

def count_emails(file_path):
    """Counts the total number of emails in a JSON file."""
    try:
        with open(file_path, 'r') as f:
            data = json.load(f)

        email_count = 0
        for entry in data:
            if 'Email' in entry:
                email_count += 1

        return email_count
    except:
        return 0

def get_current_user():
    """Get the current logged-in user."""
    try:
        with open(CREDENTIALS_FILE, 'r') as file:
            for line in file:
                if '@' in line and '.' in line.split('@')[-1]:
                    email = line.strip().split(':')[0]
                    return email
    except FileNotFoundError:
        pass
    return None

def load_credentials(filename=CREDENTIALS_FILE):
    """Loads the username and password from a file."""
    try:
        with open(filename, "r") as file:
            line = file.readline()
            username, password = line.strip().split(":")
            return username, password
    except (FileNotFoundError, IOError):
        return None

def send_personalized_email(receiver, last_name, template_file=EMAIL_TEMPLATE_FILE, attachment_paths=[]):
    """Sends a personalized email with an optional attachment."""
    credentials = load_credentials()
    if not credentials:
        return False, "No credentials found. Please add your account first."

    sender, password = credentials

    try:
        with open(template_file, "r", encoding='utf-8') as f:
            lines = f.readlines()
            if len(lines) > 1:
                lines[1] = lines[1].replace("name", last_name)  # Replace "name" on the second line
            subject = lines[0].strip() if lines else "No Subject"
            message_body = "".join(lines[1:]) if len(lines) > 1 else "".join(lines)
    except Exception as e:
        return False, f"Error reading template file: {e}"

    # Create MIMEMultipart message
    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = receiver

    # Attach message body
    message.attach(MIMEText(message_body, _charset="utf-8"))

    # Attach files if provided
    if attachment_paths:
        for attachment_path in attachment_paths:
            try:
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())

                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    message.attach(part)
            except Exception as e:
                return False, f"Error attaching file {attachment_path}: {e}"

    # Send email with error handling and log successful sends
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, receiver, message.as_string())
            
            # Log the successful send to logs.txt
            with open("logs.txt", "a", encoding='utf-8') as log_file:
                log_file.write(f"{receiver} > done > {datetime.now().strftime('%Y-%m-%d %H:%M:%S %p')}\n")
            
            return True, f"Message sent successfully to: {receiver}"
    except Exception as e:
        return False, f"Error sending email to {receiver}: {e}"

def send_emails_thread():
    """Send emails in a separate thread to avoid blocking the web interface."""
    global email_progress
    
    try:
        with open(EXTRACTED_DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        email_progress['status'] = 'error'
        email_progress['results'].append({"success": False, "message": f"No data found: {e}"})
        return
    
    attachment_paths = [os.path.join(FILES_DIRECTORY, file) for file in get_attachment_files()]
    
    email_progress['total'] = len(data)
    email_progress['current'] = 0
    email_progress['status'] = 'sending'
    email_progress['results'] = []
    
    for i, entry in enumerate(data):
        if 'Email' not in entry:
            email_progress['results'].append({"success": False, "message": f"Missing email address for entry: {entry}"})
            email_progress['current'] = i + 1
            continue
            
        receiver = entry['Email']
        last_name = entry.get("Last_Name", "Valued Customer")
        
        success, message = send_personalized_email(receiver, last_name, EMAIL_TEMPLATE_FILE, attachment_paths)
        email_progress['results'].append({"success": success, "message": message})
        email_progress['current'] = i + 1
        
        # Small delay to avoid hitting rate limits
        time.sleep(0.5)
    
    email_progress['status'] = 'completed'

def scan_data(floc, gid):
    """Scan data from Google Sheets and save to JSON."""
    try:
        # Replace with the path to your service account key JSON file 
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = ServiceAccountCredentials.from_json_keyfile_name(floc, SCOPES)

        # Authorize
        client = gspread.authorize(creds)

        # Open the sheet by ID
        sheet = client.open_by_key(gid).sheet1

        # Get all data from the sheet
        data = sheet.get_all_values()

        # Assuming Last Name and Email are in columns with known positions
        last_name_column = 2  # Adjust this based on your actual column position
        email_column = 5  # Adjust this based on your actual column position

        # Create an empty list to store extracted data
        extracted_data = []

        # Skip the header row (assuming the first row contains column names)
        for row in data[1:]:
            if len(row) >= max(last_name_column, email_column):
                last_name = row[last_name_column - 1]
                email = row[email_column - 1]
                # Only add entries with valid emails
                if email and '@' in email:
                    entry = {"Last_Name": last_name, "Email": email}
                    extracted_data.append(entry)

        # Write the extracted data to a JSON file
        with open(EXTRACTED_DATA_FILE, 'w', encoding='utf-8') as outfile:
            json.dump(extracted_data, outfile, indent=4)

        return True, f"Scan Completed! {len(extracted_data)} entries extracted."
    except Exception as e:
        return False, f"Error during scan: {e}"

def get_attachment_files():
    """Get list of files in the Files directory."""
    try:
        files = []
        for file in os.listdir(FILES_DIRECTORY):
            if os.path.isfile(os.path.join(FILES_DIRECTORY, file)):
                files.append(file)
        return files
    except:
        return []

def create_default_template():
    """Create a default email template if it doesn't exist."""
    default_template = """Welcome to Our Service!

Dear name,

Thank you for your interest in our services. We're excited to have you on board!

We believe our solution will help you achieve your goals more efficiently. If you have any questions, please don't hesitate to reach out.

Best regards,
The Support Team"""
    
    try:
        with open(EMAIL_TEMPLATE_FILE, 'w', encoding='utf-8') as f:
            f.write(default_template)
    except:
        pass

# Create default template on startup
create_default_template()

# HTML Templates as string variables that will be written to files
BASE_HTML = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Sender</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css">
    <style>
        .navbar-brand {
            font-weight: bold;
        }
        .card {
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .stats-card {
            text-align: center;
            padding: 20px;
        }
        .action-btn {
            width: 100%;
            margin-bottom: 10px;
        }
        .progress {
            height: 25px;
        }
        .progress-bar {
            line-height: 25px;
        }
    </style>
</head>
<body>
    <nav class="navbar navbar-expand-lg navbar-dark bg-dark">
        <div class="container">
            <a class="navbar-brand" href="{{ url_for('index') }}">
                <i class="bi bi-envelope"></i> Email Sender
            </a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item">
                        <a class="nav-link" href="{{ url_for('index') }}">Dashboard</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="{{ url_for('scan') }}">Scan</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="{{ url_for('template') }}">Template</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="{{ url_for('attachments') }}">Attachments</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="{{ url_for('account') }}">Account</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="{{ url_for('logs') }}">Logs</a>
                    </li>
                </ul>
            </div>
        </div>
    </nav>

    <div class="container mt-4">
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ 'danger' if category == 'error' else 'success' if category == 'success' else 'warning' }} alert-dismissible fade show">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        {% block content %}{% endblock %}
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    {% block scripts %}{% endblock %}
</body>
</html>'''

INDEX_HTML = '''{% extends "base.html" %}

{% block content %}
<div class="row">
    <div class="col-md-8">
        <div class="card">
            <div class="card-header">
                <h4>Dashboard</h4>
            </div>
            <div class="card-body">
                <div class="row mb-4">
                    <div class="col-md-6">
                        <div class="card stats-card bg-primary text-white">
                            <h5>Logged In As</h5>
                            <h3>{{ current_user if current_user else "No Account" }}</h3>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card stats-card bg-success text-white">
                            <h5>Emails Loaded</h5>
                            <h3>{{ total_emails }}</h3>
                        </div>
                    </div>
                </div>
                
                <div class="d-grid gap-2">
                    <button id="sendEmailsBtn" class="btn btn-primary btn-lg action-btn" {{ "disabled" if total_emails == 0 else "" }}>
                        <i class="bi bi-send"></i> Start Sending Emails
                    </button>
                    <a href="{{ url_for('scan') }}" class="btn btn-info btn-lg action-btn">
                        <i class="bi bi-search"></i> Scan Data
                    </a>
                    <a href="{{ url_for('template') }}" class="btn btn-warning btn-lg action-btn">
                        <i class="bi bi-file-earmark-text"></i> Edit Template
                    </a>
                    <a href="{{ url_for('attachments') }}" class="btn btn-secondary btn-lg action-btn">
                        <i class="bi bi-paperclip"></i> Manage Attachments
                    </a>
                </div>
                
                {% if total_emails == 0 %}
                <div class="alert alert-warning mt-3">
                    <i class="bi bi-exclamation-triangle"></i> No emails loaded. Please scan data first.
                </div>
                {% endif %}
            </div>
        </div>
    </div>
    
    <div class="col-md-4">
        <div class="card">
            <div class="card-header">
                <h5>Current Attachments</h5>
            </div>
            <div class="card-body">
                {% if attachment_files %}
                    <ul class="list-group">
                        {% for file in attachment_files %}
                            <li class="list-group-item d-flex justify-content-between align-items-center">
                                {{ file }}
                                <span class="badge bg-primary rounded-pill">
                                    <i class="bi bi-paperclip"></i>
                                </span>
                            </li>
                        {% endfor %}
                    </ul>
                {% else %}
                    <p class="text-muted">No attachments added yet.</p>
                {% endif %}
            </div>
        </div>
        
        <div class="card mt-4">
            <div class="card-header">
                <h5>Quick Actions</h5>
            </div>
            <div class="card-body">
                <div class="d-grid gap-2">
                    <a href="{{ url_for('account') }}" class="btn btn-outline-primary">
                        <i class="bi bi-person"></i> Account Settings
                    </a>
                    <a href="{{ url_for('logs') }}" class="btn btn-outline-info">
                        <i class="bi bi-clock-history"></i> View Logs
                    </a>
                </div>
            </div>
        </div>
    </div>
</div>

<div id="progressModal" class="modal fade" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Sending Emails</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <div class="progress mb-3">
                    <div id="progressBar" class="progress-bar" role="progressbar" style="width: 0%">0%</div>
                </div>
                <div id="progressDetails" style="max-height: 300px; overflow-y: auto;">
                    <!-- Progress details will be shown here -->
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
            </div>
        </div>
    </div>
</div>
{% endblock %}

{% block scripts %}
<script>
document.getElementById('sendEmailsBtn').addEventListener('click', function() {
    const modal = new bootstrap.Modal(document.getElementById('progressModal'));
    modal.show();
    
    const progressBar = document.getElementById('progressBar');
    const progressDetails = document.getElementById('progressDetails');
    
    progressDetails.innerHTML = '<p>Starting email sending process...</p>';
    
    // Start the email sending process
    fetch('{{ url_for("start_sending") }}', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        }
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // Start polling for progress
            const progressInterval = setInterval(() => {
                fetch('{{ url_for("get_progress") }}')
                .then(response => response.json())
                .then(progressData => {
                    const percent = progressData.total > 0 ? Math.round((progressData.current / progressData.total) * 100) : 0;
                    progressBar.style.width = percent + '%';
                    progressBar.textContent = percent + '%';
                    
                    // Update progress details
                    let detailsHtml = '';
                    if (progressData.status === 'sending') {
                        detailsHtml = `<p>Sending ${progressData.current} of ${progressData.total} emails...</p>`;
                    } else if (progressData.status === 'completed') {
                        detailsHtml = `<div class="alert alert-success">Email sending completed!</div>`;
                        clearInterval(progressInterval);
                    } else if (progressData.status === 'error') {
                        detailsHtml = `<div class="alert alert-danger">Error occurred during sending.</div>`;
                        clearInterval(progressInterval);
                    }
                    
                    // Add recent results
                    if (progressData.results && progressData.results.length > 0) {
                        const recentResults = progressData.results.slice(-10); // Show last 10 results
                        recentResults.forEach(result => {
                            const alertClass = result.success ? 'alert-success' : 'alert-danger';
                            detailsHtml += `<div class="alert ${alertClass} py-2">${result.message}</div>`;
                        });
                    }
                    
                    progressDetails.innerHTML = detailsHtml;
                    
                    // Scroll to bottom
                    progressDetails.scrollTop = progressDetails.scrollHeight;
                    
                    if (progressData.status === 'completed' || progressData.status === 'error') {
                        clearInterval(progressInterval);
                    }
                });
            }, 1000);
        } else {
            progressBar.style.width = '100%';
            progressBar.textContent = '100%';
            progressBar.classList.add('bg-danger');
            progressDetails.innerHTML = `<div class="alert alert-danger">${data.message}</div>`;
        }
    })
    .catch(error => {
        progressBar.style.width = '100%';
        progressBar.textContent = '100%';
        progressBar.classList.add('bg-danger');
        progressDetails.innerHTML = `<div class="alert alert-danger">Error: ${error}</div>`;
    });
});
</script>
{% endblock %}
'''

SCAN_HTML = '''{% extends "base.html" %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-8">
        <div class="card">
            <div class="card-header">
                <h4>Scan Data from Google Sheets</h4>
            </div>
            <div class="card-body">
                <form method="POST">
                    <div class="mb-3">
                        <label for="file_location" class="form-label">Service Account Key File Location</label>
                        <input type="text" class="form-control" id="file_location" name="file_location" 
                               placeholder="Path to your service account key JSON file" required>
                        <div class="form-text">The path to your Google Service Account key file.</div>
                    </div>
                    
                    <div class="mb-3">
                        <label for="sheet_id" class="form-label">Google Sheet ID</label>
                        <input type="text" class="form-control" id="sheet_id" name="sheet_id" 
                               placeholder="Your Google Sheet ID" required>
                        <div class="form-text">
                            The ID can be found in your Google Sheet URL: 
                            https://docs.google.com/spreadsheets/d/<strong>SHEET_ID</strong>/edit
                        </div>
                    </div>
                    
                    <div class="alert alert-info">
                        <h6>Note:</h6>
                        <ul>
                            <li>Last Name is expected in column 2</li>
                            <li>Email is expected in column 5</li>
                            <li>The first row is assumed to be headers and will be skipped</li>
                        </ul>
                    </div>
                    
                    <div class="d-grid">
                        <button type="submit" class="btn btn-primary btn-lg">
                            <i class="bi bi-search"></i> Start Scan
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>
{% endblock %}
'''

ACCOUNT_HTML = '''{% extends "base.html" %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-8">
        <div class="card">
            <div class="card-header">
                <h4>Account Settings</h4>
            </div>
            <div class="card-body">
                {% if current_account %}
                <div class="alert alert-info mb-4">
                    <h5>Current Account</h5>
                    <p class="mb-0">Logged in as: <strong>{{ current_account }}</strong></p>
                </div>
                {% else %}
                <div class="alert alert-warning mb-4">
                    <h5>No Account Configured</h5>
                    <p class="mb-0">Please add your email account credentials to send emails.</p>
                </div>
                {% endif %}
                
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h5>Add/Update Account</h5>
                            </div>
                            <div class="card-body">
                                <form method="POST">
                                    <input type="hidden" name="action" value="add">
                                    <div class="mb-3">
                                        <label for="username" class="form-label">Email Address</label>
                                        <input type="email" class="form-control" id="username" name="username" 
                                               placeholder="your.email@gmail.com" required>
                                    </div>
                                    <div class="mb-3">
                                        <label for="password" class="form-label">Password/App Password</label>
                                        <input type="password" class="form-control" id="password" name="password" 
                                               placeholder="Your password or app password" required>
                                        <div class="form-text">
                                            For Gmail, you may need to use an App Password if 2FA is enabled.
                                        </div>
                                    </div>
                                    <div class="d-grid">
                                        <button type="submit" class="btn btn-primary">
                                            <i class="bi bi-save"></i> Save Credentials
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>
                    </div>
                    
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h5>Remove Account</h5>
                            </div>
                            <div class="card-body">
                                <p>This will remove your saved account credentials.</p>
                                <form method="POST" onsubmit="return confirm('Are you sure you want to remove your account?');">
                                    <input type="hidden" name="action" value="remove">
                                    <div class="d-grid">
                                        <button type="submit" class="btn btn-danger">
                                            <i class="bi bi-trash"></i> Remove Account
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
{% endblock %}
'''

TEMPLATE_HTML = '''{% extends "base.html" %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-10">
        <div class="card">
            <div class="card-header">
                <h4>Email Template</h4>
            </div>
            <div class="card-body">
                <div class="alert alert-info mb-4">
                    <h6>Template Instructions:</h6>
                    <ul class="mb-0">
                        <li>The first line will be used as the email subject</li>
                        <li>Use <code>name</code> as a placeholder for the recipient's last name</li>
                        <li>The rest of the template will be used as the email body</li>
                    </ul>
                </div>
                
                <form method="POST">
                    <div class="mb-3">
                        <label for="template_content" class="form-label">Template Content</label>
                        <textarea class="form-control" id="template_content" name="template_content" 
                                  rows="15" required>{{ template_content }}</textarea>
                    </div>
                    
                    <div class="d-grid">
                        <button type="submit" class="btn btn-primary btn-lg">
                            <i class="bi bi-save"></i> Save Template
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>
{% endblock %}
'''

ATTACHMENTS_HTML = '''{% extends "base.html" %}

{% block content %}
<div class="row">
    <div class="col-md-8">
        <div class="card">
            <div class="card-header">
                <h4>Manage Attachments</h4>
            </div>
            <div class="card-body">
                <form method="POST" enctype="multipart/form-data">
                    <div class="mb-3">
                        <label for="file" class="form-label">Upload Attachment</label>
                        <input class="form-control" type="file" id="file" name="file" required>
                    </div>
                    <div class="d-grid">
                        <button type="submit" class="btn btn-primary">
                            <i class="bi bi-upload"></i> Upload File
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    
    <div class="col-md-4">
        <div class="card">
            <div class="card-header">
                <h5>Current Attachments</h5>
            </div>
            <div class="card-body">
                {% if attachment_files %}
                    <div class="list-group">
                        {% for file in attachment_files %}
                            <div class="list-group-item d-flex justify-content-between align-items-center">
                                <span>{{ file }}</span>
                                <a href="{{ url_for('delete_attachment', filename=file) }}" 
                                   class="btn btn-sm btn-outline-danger"
                                   onclick="return confirm('Are you sure you want to delete {{ file }}?')">
                                    <i class="bi bi-trash"></i>
                                </a>
                            </div>
                        {% endfor %}
                    </div>
                {% else %}
                    <p class="text-muted">No attachments added yet.</p>
                {% endif %}
            </div>
        </div>
    </div>
</div>
{% endblock %}
'''

LOGS_HTML = '''{% extends "base.html" %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-10">
        <div class="card">
            <div class="card-header d-flex justify-content-between align-items-center">
                <h4>Email Logs</h4>
                <a href="{{ url_for('logs') }}" class="btn btn-outline-primary btn-sm">
                    <i class="bi bi-arrow-clockwise"></i> Refresh
                </a>
            </div>
            <div class="card-body">
                {% if logs_content %}
                    <pre style="max-height: 500px; overflow-y: auto; background: #f8f9fa; padding: 15px; border-radius: 5px;">{{ logs_content }}</pre>
                {% else %}
                    <p class="text-muted">No logs available.</p>
                {% endif %}
            </div>
        </div>
    </div>
</div>
{% endblock %}
'''

# Write HTML templates to files
def create_template_files():
    templates = {
        'base.html': BASE_HTML,
        'index.html': INDEX_HTML,
        'scan.html': SCAN_HTML,
        'account.html': ACCOUNT_HTML,
        'template.html': TEMPLATE_HTML,
        'attachments.html': ATTACHMENTS_HTML,
        'logs.html': LOGS_HTML
    }
    
    for filename, content in templates.items():
        with open(f'templates/{filename}', 'w', encoding='utf-8') as f:
            f.write(content)

# Create template files on startup
create_template_files()

# Flask Routes
@app.route('/')
def index():
    """Main dashboard page."""
    current_user = get_current_user()
    total_emails = count_emails(EXTRACTED_DATA_FILE)
    attachment_files = get_attachment_files()
    
    return render_template('index.html', 
                         current_user=current_user, 
                         total_emails=total_emails,
                         attachment_files=attachment_files)

@app.route('/start_sending', methods=['POST'])
def start_sending():
    """Start the email sending process in a separate thread."""
    global email_progress
    
    # Check if already sending
    if email_progress['status'] == 'sending':
        return jsonify({"success": False, "message": "Email sending is already in progress."})
    
    # Check if we have data
    if count_emails(EXTRACTED_DATA_FILE) == 0:
        return jsonify({"success": False, "message": "No email data found. Please scan first."})
    
    # Check if we have credentials
    if not load_credentials():
        return jsonify({"success": False, "message": "No account credentials found. Please add your account first."})
    
    # Start sending in a separate thread
    thread = threading.Thread(target=send_emails_thread)
    thread.daemon = True
    thread.start()
    
    return jsonify({"success": True, "message": "Email sending started."})

@app.route('/get_progress')
def get_progress():
    """Get the current progress of email sending."""
    global email_progress
    return jsonify(email_progress)

@app.route('/scan', methods=['GET', 'POST'])
def scan():
    """Scan data from Google Sheets."""
    if request.method == 'POST':
        floc = request.form.get('file_location')
        gid = request.form.get('sheet_id')
        
        if not floc or not gid:
            flash('Please provide both file location and sheet ID.', 'error')
            return redirect(url_for('scan'))
        
        success, message = scan_data(floc, gid)
        
        if success:
            flash(message, 'success')
        else:
            flash(message, 'error')
        
        return redirect(url_for('index'))
    
    return render_template('scan.html')

@app.route('/account', methods=['GET', 'POST'])
def account():
    """Manage account settings."""
    if request.method == 'POST':
        action = request.form.get('action')
        
        if action == 'add':
            username = request.form.get('username')
            password = request.form.get('password')
            
            if not username or not password:
                flash('Please provide both username and password.', 'error')
                return redirect(url_for('account'))
            
            try:
                with open(CREDENTIALS_FILE, "w", encoding='utf-8') as file:
                    file.write(f"{username}:{password}\n")
                flash('Credentials saved successfully!', 'success')
            except Exception as e:
                flash(f'Error saving credentials: {e}', 'error')
                
        elif action == 'remove':
            try:
                if os.path.exists(CREDENTIALS_FILE):
                    os.remove(CREDENTIALS_FILE)
                    flash('Account removed successfully!', 'success')
                else:
                    flash('No account found to remove.', 'warning')
            except Exception as e:
                flash(f'Error removing account: {e}', 'error')
    
    # Check if account exists
    credentials = load_credentials()
    current_account = None
    if credentials:
        current_account = credentials[0]  # Just the username
    
    return render_template('account.html', current_account=current_account)

@app.route('/template', methods=['GET', 'POST'])
def template():
    """Manage email template."""
    if request.method == 'POST':
        content = request.form.get('template_content')
        
        try:
            with open(EMAIL_TEMPLATE_FILE, 'w', encoding='utf-8') as f:
                f.write(content)
            flash('Template saved successfully!', 'success')
        except Exception as e:
            flash(f'Error saving template: {e}', 'error')
        
        return redirect(url_for('template'))
    
    # Read current template
    template_content = ""
    try:
        with open(EMAIL_TEMPLATE_FILE, 'r', encoding='utf-8') as f:
            template_content = f.read()
    except:
        create_default_template()
        with open(EMAIL_TEMPLATE_FILE, 'r', encoding='utf-8') as f:
            template_content = f.read()
    
    return render_template('template.html', template_content=template_content)

@app.route('/attachments', methods=['GET', 'POST'])
def attachments():
    """Manage attachment files."""
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file selected.', 'error')
            return redirect(url_for('attachments'))
        
        file = request.files['file']
        if file.filename == '':
            flash('No file selected.', 'error')
            return redirect(url_for('attachments'))
        
        if file:
            filename = file.filename
            file.save(os.path.join(FILES_DIRECTORY, filename))
            flash(f'File "{filename}" uploaded successfully!', 'success')
            return redirect(url_for('attachments'))
    
    attachment_files = get_attachment_files()
    return render_template('attachments.html', attachment_files=attachment_files)

@app.route('/delete_attachment/<filename>')
def delete_attachment(filename):
    """Delete an attachment file."""
    try:
        file_path = os.path.join(FILES_DIRECTORY, filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            flash(f'File "{filename}" deleted successfully!', 'success')
        else:
            flash(f'File "{filename}" not found.', 'error')
    except Exception as e:
        flash(f'Error deleting file: {e}', 'error')
    
    return redirect(url_for('attachments'))

@app.route('/logs')
def logs():
    """View email sending logs."""
    try:
        with open("logs.txt", "r", encoding='utf-8') as log_file:
            logs_content = log_file.read()
    except:
        logs_content = "No logs available."
    
    return render_template('logs.html', logs_content=logs_content)

if __name__ == '__main__':
    print("Starting Email Sender Web Application...")
    print("Access the application at: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)

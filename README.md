# 📧 Outlook Email Gateway for Recruitment

A lightweight local HTTP gateway that enables recruitment and staffing agencies to programmatically create Microsoft Outlook draft emails. This system bridges a web application with Microsoft Outlook, allowing recruiting teams to compose personalized emails to candidates with rich formatting, CV attachments, and multiple recipients—each receiving their own separate draft for independent review and sending.

---

## 🎯 Overview

**Problem:** Recruitment specialists need to send customized emails with CV attachments to multiple candidate recipients (companies, hiring managers, etc.). Manual email composition is time-consuming, and simple mailto: links cannot handle file attachments or create multiple personalized drafts efficiently.

**Solution:** The Outlook Email Gateway provides a web-based form that feeds into a local REST API, automatically creating individual Outlook email drafts on the user's machine. This allows recruiting teams to:
- ✅ Compose emails in a web form
- ✅ Attach CV files directly
- ✅ Create separate drafts for each recipient (not group emails)
- ✅ Review, edit, and send emails personally from their Outlook account

**Who It's For:**
- 👔 Recruitment specialists and headhunters
- 💼 Staffing agency teams
- 📧 HR professionals managing candidate outreach
- 🎯 Any team handling bulk personalized email with attachments

---

## 🏗️ Architecture

```
Recruiting Specialist
         ↓
   Web Form (index.html)
   - Email subject
   - Recipient(s) email address(es)
   - Email body (HTML)
   - CV attachment (file)
         ↓
    [Send Button Click]
         ↓
   HTTP POST Request
   (w/ authentication token)
         ↓
    [Outlook Email Gateway Server] ← Flask REST API (localhost:5001)
         ↓
  Validation & File Handling
         ↓
  win32com.client (COM Bridge)
         ↓
   Microsoft Outlook
         ↓
   Create Individual Draft(s)
   (One per recipient)
         ↓
   Outlook Opens Locally
   with Draft(s) Ready for Review
         ↓
   Specialist Reviews, Edits & Sends
   (from personal email account)
```

**Gateway Role:**
- **Web Form Handler:** Captures email composition data from the web UI
- **Validation Layer:** Enforces required fields and validates recipient addresses
- **File Management:** Converts base64-encoded attachments to temporary files
- **COM Orchestration:** Handles Windows COM objects for reliable Outlook interaction
- **Multi-Draft Creation:** Creates one Outlook draft per recipient (not group emails)
- **Security:** Validates requests with shared secret authentication
- **Error Handling:** Returns structured JSON responses with meaningful error messages
- **Local Operation:** Runs on localhost (127.0.0.1) for security and direct system access

---

## ✨ Features

✅ **Web Form Interface** - User-friendly form for composing recruitment emails  
✅ **Multiple Recipients with Separate Drafts** - Each recipient gets their own individual draft (not group emails)  
✅ **CV File Attachments** - Upload and attach CV files directly (not as links)  
✅ **HTML Rich Text Support** - Full HTML formatting support in email body  
✅ **Shared Secret Authentication** - Token-based request validation for security  
✅ **CORS Enabled** - Allow cross-origin requests from local web applications  
✅ **Local Operation** - Runs on localhost (127.0.0.1:5001) - no internet communication needed  
✅ **Automatic Draft Opening** - Drafts open immediately in Outlook after form submission  
✅ **Health Checks** - Ping endpoint to verify gateway availability  
✅ **Structured Error Responses** - JSON-formatted error messages for debugging  
✅ **Temporary File Management** - Safely handles file uploads with automatic cleanup  
✅ **COM Error Handling** - Graceful handling when Outlook is unavailable  
✅ **Independent Email Drafting** - Allows users to review, edit, and send emails personally from their own Outlook account  

---

## 🧠 Why This Solution?

### Why not use `mailto:` links?
- **No file attachments:** `mailto:` cannot attach files—recruiters must manually add CVs
- **No bulk handling:** Creating 50+ separate emails requires 50+ clicks
- **Limited formatting:** Complex HTML formatting not reliably supported
- **Poor UX:** No form validation or user feedback

### Why can't a web app directly open Outlook?
- **Browser security sandbox:** Web browsers cannot access local applications due to CORS and same-origin policies
- **Operating system protection:** Browsers intentionally cannot execute local programs or access file systems
- **No COM access:** JavaScript running in browsers cannot interact with Windows COM objects (Outlook)

### Why can't a web app access local CV files?
- **File system access:** Browsers cannot read files from a user's hard drive without explicit user interaction through a file picker
- **Security design:** This is by design to prevent malicious websites from stealing personal data
- **Uploaded files only:** Web apps can only read files that users explicitly upload through form inputs

### The Solution:
A **local HTTP gateway** acts as a bridge:
- ✅ Runs on the user's machine with full system access
- ✅ Handles file uploads securely
- ✅ Directly interfaces with Outlook via COM
- ✅ Maintains security through local-only operation and authentication tokens
- ✅ Allows the web app to leverage the browser's UI while delegating to a trusted local service

---

| Component | Technology |
|-----------|-----------|
| **Backend** | Python 3.x |
| **Framework** | Flask |
| **Outlook Integration** | win32com.client (pywin32) |
| **Middleware** | Flask-CORS |
| **Packaging** | PyInstaller |
| **Frontend** | HTML5 / JavaScript (vanilla) |
| **Styling** | CSS3 |

---

## 📦 Installation

### Prerequisites
- **Windows OS** (Outlook integration requires Windows COM)
- **Microsoft Outlook** installed and configured locally
- **Python 3.7+** (or use the pre-compiled executable)

### Quick Start for End Users (Using Pre-Compiled Executable)

If you have the compiled `agent.exe` executable:

1. **Download and place the executable** in a folder on your machine (e.g., `C:\OutlookGateway\`)
2. **Run the executable:**
   - Double-click `agent.exe`, or
   - Run from command line: `agent.exe`
3. **You should see:**
   ```
   Agent running at http://127.0.0.1:5001
   ```
4. **Open the web interface** by going to `http://127.0.0.1:5001` in your browser
5. **Fill the form and click "Open Drafts"** to create email drafts in Outlook

### Setup for Developers (From Source)

1. **Clone or download the project:**
   ```bash
   git clone <repository-url>
   cd outlook-email-gateway
   ```

2. **Install Python dependencies:**
   ```bash
   pip install flask flask-cors pywin32
   ```

3. **Run the gateway:**
   ```bash
   python agent.py
   ```
   
   You should see:
   ```
   Agent running at http://127.0.0.1:5001
   ```

4. **Access the web interface:**
   - Open `http://127.0.0.1:5001` in your browser
   - Or open `index.html` locally (requires gateway to be running)

### Building Your Own Standalone Executable

To create a standalone `.exe` file from source:

1. **Install PyInstaller:**
   ```bash
   pip install pyinstaller
   ```

2. **Build the executable:**
   ```bash
   pyinstaller agent.spec
   ```

3. **Find the compiled exe in:**
   ```
   dist/agent/agent.exe
   ```

4. **Distribute `agent.exe`** to users without requiring Python installation

### Firewall and Security Notes

⚠️ **Important:**
- The gateway listens on `127.0.0.1:5001` (localhost only, not accessible from network)
- Windows Firewall may prompt when first running—this is normal
- For deployment on shared networks, consider adding firewall exceptions

---

## ⚙️ Configuration

Edit the configuration in `agent.py`:

```python
LISTEN_HOST = "127.0.0.1"    # Gateway listening address
LISTEN_PORT = 5001            # Gateway listening port
SHARED_SECRET = "secure_secret_here"  # Authentication token
```

For frontend configuration, update `config.js`:

```javascript
window.AGENT_SECRET = "secure_secret_here";  // Must match backend secret
```

---

## 🔐 Environment Variables & Secrets

| Variable | Default | Purpose |
|----------|---------|---------|
| `LISTEN_HOST` | `127.0.0.1` | IP address for gateway binding (localhost only for security) |
| `LISTEN_PORT` | `5001` | HTTP port for API server |
| `SHARED_SECRET` | `secure_secret_here` | Authentication token required for all POST requests |

⚠️ **Security Note:** Change `SHARED_SECRET` in both `agent.py` and `config.js` to a strong, unique value before deployment.

---

## 📡 API Documentation

### 1. Health Check

**Endpoint:** `GET /ping`

**Purpose:** Verify gateway availability

**Request:**
```bash
curl http://127.0.0.1:5001/ping
```

**Response (200 OK):**
```json
{
  "status": "ok"
}
```

---

### 2. Create Email Draft

**Endpoint:** `POST /open-mail`

**Purpose:** Create a new Outlook email draft with recipients, subject, body, and optional attachment

**Request Headers:**
```
Content-Type: application/json
```

**Request Body:**
```json
{
  "secret": "secure_secret_here",
  "subject": "Project Status Update",
  "body": "<h2>Hello!</h2><p>This is an HTML email.</p>",
  "recipients": ["user1@example.com", "user2@example.com"],
  "attachment_name": "report.pdf",
  "attachment_base64": "JVBERi0xLjQKJeLj..."
}
```

**Request Parameters:**

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `secret` | string | Yes | Shared authentication token |
| `subject` | string | Yes | Email subject line |
| `body` | string | No | Email body (supports HTML) |
| `recipients` | array | Yes | List of email addresses |
| `attachment_name` | string | No | Filename for attachment |
| `attachment_base64` | string | No | Base64-encoded file content |

**Response (200 OK):**
```json
{
  "ok": true
}
```

**Error Responses:**

```json
// Unauthorized (401)
{
  "ok": false,
  "error": "Unauthorized"
}

// Missing fields (400)
{
  "ok": false,
  "error": "Missing recipients"
}

// Server error (500)
{
  "ok": false,
  "error": "Outlook is not available"
}
```

---

## 📂 Project Structure

```
outlook-email-gateway/
├── agent.py                 # Main Flask API server & Outlook integration
├── agent.spec              # PyInstaller configuration for building .exe
├── index.html              # Web form UI for recruiters
├── index.css               # Styling for web interface
├── config.js               # Frontend configuration & authentication
├── build/                  # PyInstaller build artifacts
│   └── agent/              # Compiled executable resources
└── README.md               # This documentation
```

### Key Files Explained

**agent.py** — Main Application
- Flask REST API server on localhost:5001
- Two endpoints:
  - `GET /ping` — Health check
  - `POST /open-mail` — Create email draft(s)
- Outlook COM integration using `win32com.client`
- Authentication & validation logic
- File handling (base64 to temporary files)

**index.html** — Recruiter Web Interface
- Form with fields:
  - Recipients (comma-separated email addresses)
  - Subject
  - Message body (HTML support)
  - CV file attachment
- Real-time status feedback
- Form validation
- Automatic file-to-base64 conversion
- Error handling and user notifications

**config.js** — Shared Secret Configuration
- Stores authentication token used by frontend
- Must match the `SHARED_SECRET` in agent.py
- Loaded by index.html

**agent.spec** — Build Configuration
- PyInstaller specification file
- Contains build options for creating `agent.exe`
- Defines entry point and dependencies

---

## 🔑 Authentication

The gateway uses **token-based authentication** via a shared secret:

1. **Server Configuration:** Set `SHARED_SECRET` in `agent.py`
2. **Client Configuration:** Set `window.AGENT_SECRET` in `config.js` to match server secret
3. **Request Validation:** Every POST request to `/open-mail` must include the `secret` field in JSON payload
4. **Validation Logic:** If the provided secret doesn't match `SHARED_SECRET`, the request is rejected with a **401 Unauthorized** response

⚠️ **Best Practice:** Use environment variables or secure configuration management for the shared secret instead of hardcoding.

---

## 💡 Usage Examples

### Typical Workflow for Recruiters

1. **Start the Gateway**
   - Launch `agent.exe` (or `python agent.py`)
   - Ensure Outlook is running in the background

2. **Open the Web Interface**
   - Navigate to `http://127.0.0.1:5001`
   - (Or open `index.html` if hosting locally)

3. **Fill in the recruitment email details:**
   - **Subject:** "Hiring Opportunity - Senior Developer Position"
   - **Recipients:** `hiring@company1.com, hiring@company2.com`
   - **Message:** Your recruitment pitch (supports HTML formatting)
   - **Attachment:** Upload the candidate's CV file

4. **Click "Open Drafts"**
   - Outlook automatically opens with **2 separate drafts** (one per recipient)
   - Each draft contains:
     - The same subject line
     - The same message body
     - The CV attachment
     - Ready for personal review and editing

5. **In Outlook:**
   - Review each draft individually
   - Edit the message if needed (personalize for each company)
   - Add CC/BCC or other fields as needed
   - Click "Send" from your personal email account

---

### Example: Sending to Multiple Companies

**Form Input:**
```
Subject: Recruiting Talent - Full-Stack Developer
Recipients: hr@amazon.com, talent@google.com, careers@microsoft.com
Body: <p>We have an exciting opportunity for a Full-Stack Developer...</p>
Attachment: candidate_cv.pdf
```

**Result:**
✅ 3 separate Outlook drafts open, each with:
- To: one of the companies
- Subject: "Recruiting Talent - Full-Stack Developer"
- Body: The HTML message
- Attachment: candidate_cv.pdf file

---

### Example: Via API (cURL)

If you're integrating with another system:

```bash
curl -X POST http://127.0.0.1:5001/open-mail \
  -H "Content-Type: application/json" \
  -d '{
    "secret": "secure_secret_here",
    "subject": "Job Opportunity",
    "body": "<h3>Hello!</h3><p>We have a great role for you...</p>",
    "recipients": ["company@example.com"]
  }'
```

**Response:**
```json
{
  "ok": true
}
```

Outlook draft(s) will open immediately.

---

### Example: With File Attachment (Python)

For programmatic integration:

```python
import requests
import base64

# Read CV file and encode as base64
with open("candidate_cv.pdf", "rb") as f:
    cv_base64 = base64.b64encode(f.read()).decode()

# Create email drafts
payload = {
    "secret": "secure_secret_here",
    "subject": "Exciting Career Opportunity",
    "body": "<p>We'd like to introduce a great role...</p>",
    "recipients": ["hr@company.com", "hiring@anothercompany.com"],
    "attachment_name": "candidate_cv.pdf",
    "attachment_base64": cv_base64
}

response = requests.post(
    "http://127.0.0.1:5001/open-mail",
    json=payload
)

if response.json()["ok"]:
    print("✅ Email drafts opened in Outlook!")
else:
    print("❌ Error:", response.json()["error"])
```

---

## ⚠️ Error Handling

The gateway implements structured error handling:

**Validation Errors:**
- Missing `secret` parameter → **401 Unauthorized**
- Missing `recipients` array → **400 Bad Request**
- Missing `subject` → **400 Bad Request**

**Runtime Errors:**
- Outlook not installed/running → **500 Server Error** with message "Outlook is not available"
- File handling issues → **500 Server Error** with detailed exception message

**Error Response Format:**
```json
{
  "ok": false,
  "error": "Descriptive error message"
}
```

---

## 🚀 Future Improvements

**Recruitment-Specific Features:**
- [ ] CC/BCC field support in form
- [ ] Email template library (common recruitment pitches)
- [ ] Recipient email validation before sending
- [ ] Candidate profile auto-fill integration
- [ ] Send email tracking (beyond drafts)
- [ ] Email history/log viewer
- [ ] Resume database integration
- [ ] Bulk import of recipients from CSV
- [ ] Email preview before opening Outlook

**Infrastructure & Reliability:**
- [ ] Persistent configuration file instead of hardcoding secrets
- [ ] Comprehensive logging and audit trail
- [ ] Automatic gateway restart on failure
- [ ] Multiple attachment support (not just CV)
- [ ] Outlook availability check before processing
- [ ] Batch email creation (queue management)
- [ ] Rate limiting to prevent abuse
- [ ] Support for different email body templates

**Security Enhancements:**
- [ ] Environment variables for configuration (avoid secrets in code)
- [ ] API key authentication (rotating keys)
- [ ] HTTPS/TLS support for secure communication
- [ ] HTML sanitization to prevent injection attacks
- [ ] Input validation and sanitization for recipient addresses
- [ ] IP whitelisting for network deployments
- [ ] Encrypted configuration storage

**User Experience:**
- [ ] Dark mode support for web interface
- [ ] Rich text editor for email body composition
- [ ] Drag-and-drop file upload
- [ ] Multi-language support (including Hebrew)
- [ ] Improved error messages and tooltips
- [ ] Recipients validation (email format check)
- [ ] Draft preview before opening Outlook
- [ ] Keyboard shortcuts for power users

**Integration & Extensibility:**
- [ ] Webhook support for external systems
- [ ] REST API documentation and OpenAPI/Swagger spec
- [ ] Integration with recruitment platforms (LinkedIn, Indeed, etc.)
- [ ] CRM integration (Salesforce, HubSpot)
- [ ] Database backend for storing email history
- [ ] Support for other email clients (Gmail, etc.) via web APIs

---

## 📄 License

[Your License Here]

---

## 📞 Support

For issues or questions:
1. Check the error messages returned by the API
2. Verify Outlook is installed and running
3. Ensure the shared secret matches between `agent.py` and `config.js`
4. Check that the gateway is running on port 5001

---

**Built with ❤️ for Windows automation and integration excellence**

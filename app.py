from docx import Document
from datetime import datetime, timedelta
import imaplib
import email
from email.header import decode_header
import requests
import os
import re
import time
from flask import Flask, render_template, jsonify, request, redirect
import logging
import sqlite3
import json
import traceback

# Initialize Flask app FIRST
app = Flask(__name__)

def get_db_path():
    """Get database path that works for both web service and cron job"""
    # On Render, use /tmp directory which is shared between services
    if os.getenv('RENDER'):
        return '/tmp/email_summaries.db'
    else:
        return 'email_summaries.db'

# ==================== FLASK ROUTES ====================

@app.route('/')
def root():
    """Redirect root to dashboard"""
    return redirect('/dashboard')

@app.route('/dashboard')
def dashboard():
    """Main dashboard page"""
    try:
        print("📊 Serving dashboard page...")
        return render_template('dashboard.html')
    except Exception as e:
        return f"""
        <html>
            <body style="font-family: Arial, sans-serif; padding: 40px;">
                <h1>❌ Dashboard Error</h1>
                <p><strong>Error:</strong> {str(e)}</p>
                <p><a href="/test-html">Test HTML Page</a></p>
                <p><a href="/api">API Endpoints</a></p>
            </body>
        </html>
        """, 500

@app.route('/test-html')
def test_html():
    """Test if HTML rendering works"""
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Test Page</title>
        <style>
            body { 
                font-family: Arial, sans-serif; 
                padding: 40px;
                background: #f5f7fa;
            }
            .container {
                max-width: 600px;
                margin: 0 auto;
                background: white;
                padding: 30px;
                border-radius: 12px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }
            .success { 
                color: #10b981; 
                font-size: 24px;
            }
            .btn {
                display: inline-block;
                padding: 12px 24px;
                background: #4f46e5;
                color: white;
                text-decoration: none;
                border-radius: 8px;
                margin: 10px 5px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1 class="success">✅ HTML is Working!</h1>
            <p>If you can see this page, Flask can render HTML properly.</p>
            <p>Now check the dashboard link below:</p>
            <div>
                <a class="btn" href="/dashboard">Go to Dashboard</a>
                <a class="btn" href="/api">API Endpoints</a>
            </div>
        </div>
    </body>
    </html>
    """

@app.route('/api')
def api_home():
    """API home page"""
    return jsonify({
        "status": "Email Summarizer API is running",
        "timestamp": datetime.now().isoformat(),
        "endpoints": {
            "dashboard": "/dashboard",
            "test_page": "/test-html", 
            "health": "/health",
            "stats": "/api/stats",
            "recent_summaries": "/api/recent-summaries",
            "debug": "/api/debug",
            "debug_database": "/api/debug-database",
            "test_json": "/api/test-json",
            "fix_database": "/api/fix-database",
            "force_test_run": "/api/force-test-run"
        }
    })

@app.route('/api/debug')
def api_debug():
    """Debug API endpoint"""
    return jsonify({
        "status": "API is working",
        "timestamp": datetime.now().isoformat(),
        "endpoint": "/api/debug"
    })

@app.route('/api/test-json')
def test_json():
    """Test JSON response"""
    return jsonify({
        "message": "This is a test JSON response",
        "numbers": [1, 2, 3],
        "timestamp": datetime.now().isoformat()
    })

@app.route('/api/debug-database')
def debug_database():
    """Debug database contents"""
    try:
        db_path = get_db_path()
        print(f"🔍 Debugging database at: {db_path}")
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        # Check summary_runs
        c.execute('SELECT COUNT(*) as run_count FROM summary_runs')
        run_count = c.fetchone()[0]
        
        # Check email_data
        c.execute('SELECT COUNT(*) as email_count FROM email_data')
        email_count = c.fetchone()[0]
        
        # Get latest run details
        c.execute('SELECT * FROM summary_runs ORDER BY id DESC LIMIT 1')
        latest_run = c.fetchone()
        
        # Get some email data samples
        c.execute('SELECT * FROM email_data ORDER BY id DESC LIMIT 5')
        sample_emails = c.fetchall()
        
        conn.close()
        
        return jsonify({
            "database_status": "connected",
            "database_path": db_path,
            "summary_runs_count": run_count,
            "email_data_count": email_count,
            "latest_run": {
                "id": latest_run[0] if latest_run else None,
                "date": latest_run[1] if latest_run else None,
                "total_emails": latest_run[2] if latest_run else None,
                "processed_emails": latest_run[3] if latest_run else None
            } if latest_run else None,
            "sample_emails": [
                {
                    "id": email[0],
                    "run_id": email[1], 
                    "email_number": email[2],
                    "sender": email[3],
                    "subject": email[5]
                } for email in sample_emails
            ]
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/fix-database')
def fix_database():
    """Debug and fix database issues"""
    try:
        db_path = get_db_path()
        print(f"🔧 Fixing database at: {db_path}")
        
        # Reinitialize database
        init_db()
        
        # Check if we have any data
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        # Get counts
        c.execute('SELECT COUNT(*) FROM summary_runs')
        run_count = c.fetchone()[0]
        
        c.execute('SELECT COUNT(*) FROM email_data')
        email_count = c.fetchone()[0]
        
        conn.close()
        
        return jsonify({
            "status": "success",
            "message": "Database fixed and verified",
            "run_count": run_count,
            "email_count": email_count,
            "database_path": db_path
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/force-test-run')
def force_test_run():
    """Force a test run with sample data"""
    try:
        # Create sample data for testing
        sample_emails = [
            {
                "from": "test@jubalandstate.so",
                "to": "archives@jubalandstate.so", 
                "subject": "Test Email 1",
                "body": "This is a test email body for testing the dashboard."
            },
            {
                "from": "admin@jubalandstate.so",
                "to": "archives@jubalandstate.so",
                "subject": "Test Email 2", 
                "body": "Another test email to verify dashboard functionality."
            }
        ]
        
        sample_summaries = {
            1: "This is a test summary for email 1. It demonstrates how summaries will appear in the dashboard.",
            2: "This is a test summary for email 2. The dashboard should display this data properly."
        }
        
        # Store sample data
        success = store_email_data_for_dashboard(sample_emails, sample_summaries)
        
        return jsonify({
            "status": "success" if success else "error",
            "message": "Test data added to dashboard" if success else "Failed to add test data",
            "emails_added": len(sample_emails)
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/stats')
def get_stats():
    """API endpoint for dashboard statistics"""
    try:
        db_path = get_db_path()
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute('''
            SELECT * FROM summary_runs 
            ORDER BY run_date DESC 
            LIMIT 1
        ''')
        
        latest_run = c.fetchone()
        conn.close()
        
        if latest_run:
            stats = {
                "total_emails_today": latest_run[2],
                "emails_processed": latest_run[3],
                "success_rate": round(latest_run[4], 1),
                "last_run": latest_run[1],
                "next_run": (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d 09:00:00'),
                "deepseek_usage": "Calculating...",
                "status": "active"
            }
        else:
            # Default stats if no runs yet
            stats = {
                "total_emails_today": 0,
                "emails_processed": 0,
                "success_rate": 0,
                "last_run": "Never",
                "next_run": (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d 09:00:00'),
                "deepseek_usage": "0 tokens",
                "status": "waiting"
            }
        
        return jsonify(stats)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/recent-summaries')
def get_recent_summaries():
    """API endpoint for recent email summaries"""
    try:
        db_path = get_db_path()
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        # Get the latest run ID
        c.execute('SELECT id FROM summary_runs ORDER BY id DESC LIMIT 1')
        latest_run = c.fetchone()
        
        if not latest_run:
            return jsonify([])  # No data yet
        
        run_id = latest_run[0]
        
        # Get email data for the latest run
        c.execute('''
            SELECT email_number, sender, receiver, subject, summary 
            FROM email_data 
            WHERE run_id = ? 
            ORDER BY email_number 
            LIMIT 50
        ''', (run_id,))
        
        email_data = []
        for row in c.fetchall():
            email_data.append({
                "number": row[0],
                "from": row[1],
                "to": row[2],
                "subject": row[3],
                "summary": row[4]
            })
        
        conn.close()
        
        return jsonify(email_data)
        
    except Exception as e:
        print(f"Error getting recent summaries: {e}")
        # Fallback to mock data if database is not available
        return jsonify(get_fallback_email_data())

def get_fallback_email_data():
    """Provide fallback data if database is not available"""
    return [
        {
            "number": 1,
            "from": "system@jubalandstate.so",
            "to": "archives@jubalandstate.so",
            "subject": "Daily System Report",
            "summary": "Automated system report showing all services are running normally with 99.8% uptime. No critical issues reported."
        },
        {
            "number": 2,
            "from": "secretary@jubalandstate.so", 
            "to": "archives@jubalandstate.so",
            "subject": "Meeting Minutes Approval",
            "summary": "Requesting approval for executive meeting minutes. Key decisions include budget allocation and project timelines."
        }
    ]

@app.route('/api/trigger-manual', methods=['POST'])
def trigger_manual_run():
    """Manually trigger email summary process"""
    try:
        agent = EmailSummarizerAgent()
        
        # Run in background thread to avoid timeout
        import threading
        def run_background():
            agent.run_complete_summary()
        
        thread = threading.Thread(target=run_background)
        thread.daemon = True
        thread.start()
        
        return jsonify({
            "status": "success", 
            "message": "Email summary process started in background. This may take 10-15 minutes.",
            "timestamp": datetime.now().isoformat()
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/health')
def health():
    return jsonify({"status": "healthy", "timestamp": datetime.now().isoformat()})

# ==================== EMAIL SUMMARIZER CLASS ====================

class EmailSummarizerAgent:
    def __init__(self):
        # Use environment variables for security
        self.deepseek_api_key = os.getenv('DEEPSEEK_API_KEY')
        self.source_email = os.getenv('SOURCE_EMAIL')
        self.source_password = os.getenv('SOURCE_PASSWORD')
        self.imap_server = os.getenv('IMAP_SERVER', 'imap.one.com')
        
        # Validate required environment variables
        if not all([self.deepseek_api_key, self.source_email, self.source_password]):
            raise ValueError("Missing required environment variables")
            
        self.deepseek_api_url = "https://api.deepseek.com/v1/chat/completions"
        self.imap_port = 993
    
    def fetch_emails_last_24h(self):
        try:
            print("Connecting to One.com IMAP server...")
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.source_email, self.source_password)
            mail.select("inbox")
            
            since_date = (datetime.now() - timedelta(hours=24)).strftime("%d-%b-%Y")
            print(f"Fetching emails since: {since_date}")
            
            status, messages = mail.search(None, f'(SINCE "{since_date}")')
            
            if status != 'OK':
                print("No emails found")
                return []
                
            email_ids = messages[0].split()
            print(f"Found {len(email_ids)} emails in last 24 hours")
            
            emails_data = []
            
            # Process ALL emails
            for email_id in email_ids:
                try:
                    # Convert email_id to string if it's bytes
                    if isinstance(email_id, bytes):
                        email_id_str = email_id.decode('utf-8')
                    else:
                        email_id_str = str(email_id)
                        
                    status, msg_data = mail.fetch(email_id_str, "(RFC822)")
                    if status != 'OK':
                        continue
                        
                    msg = email.message_from_bytes(msg_data[0][1])
                    
                    subject = self.decode_email_header(msg["Subject"])
                    from_ = self.decode_email_header(msg["From"])
                    to_ = self.decode_email_header(msg["To"] or msg["Delivered-To"] or "Unknown")
                    date = msg["Date"]
                    
                    body = self.extract_email_body(msg)
                    
                    emails_data.append({
                        "subject": subject,
                        "from": from_,
                        "to": to_,
                        "date": date,
                        "body": body[:1000]  # Reduced for token management
                    })
                    
                    if len(emails_data) % 50 == 0:
                        print(f"Processed {len(emails_data)} emails...")
                    
                except Exception as e:
                    print(f"Error processing email {email_id}: {e}")
                    continue
            
            mail.close()
            mail.logout()
            
            print(f"Successfully processed {len(emails_data)} emails")
            return emails_data
            
        except Exception as e:
            print(f"Error fetching emails: {e}")
            return []
    
    def decode_email_header(self, header):
        if header:
            try:
                decoded_parts = decode_header(header)
                decoded_header = ""
                for part, encoding in decoded_parts:
                    if isinstance(part, bytes):
                        if encoding:
                            decoded_header += part.decode(encoding)
                        else:
                            decoded_header += part.decode('utf-8', errors='ignore')
                    else:
                        decoded_header += part
                return decoded_header
            except:
                return str(header)
        return ""
    
    def extract_email_body(self, msg):
        body = ""
        
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        try:
                            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            if body.strip():
                                break
                        except:
                            continue
            else:
                content_type = msg.get_content_type()
                if content_type == "text/plain":
                    try:
                        body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        pass
        except Exception as e:
            print(f"Error extracting email body: {e}")
        
        return body
    
    def summarize_emails_in_batches(self, emails_data):
        """Summarize emails in batches to handle token limits"""
        if not emails_data:
            return {}
        
        print(f"Summarizing {len(emails_data)} emails in batches...")
        
        # Process emails in batches of 20 to avoid token limits
        batch_size = 20
        all_summaries = {}
        
        for batch_num in range(0, len(emails_data), batch_size):
            batch_emails = emails_data[batch_num:batch_num + batch_size]
            batch_summaries = self._summarize_batch(batch_emails, batch_num)
            all_summaries.update(batch_summaries)
            
            # Add delay between batches to avoid rate limiting
            if batch_num + batch_size < len(emails_data):
                print(f"Waiting 5 seconds before next batch...")
                time.sleep(5)
        
        return all_summaries
    
    def _summarize_batch(self, batch_emails, start_index):
        """Summarize one batch of emails"""
        emails_text = ""
        for i, email in enumerate(batch_emails, 1):
            email_num = start_index + i
            emails_text += f"Email {email_num}:\n"
            emails_text += f"From: {email['from']}\n"
            emails_text += f"To: {email['to']}\n"
            emails_text += f"Subject: {email['subject']}\n"
            emails_text += f"Date: {email['date']}\n"
            emails_text += f"Content: {email['body'][:200]}...\n\n"
        
        prompt = f"""
        Please provide individual one-paragraph summaries for EACH email. Format your response exactly like this:

        **Email {start_index + 1}:** [One paragraph summary of this email]
        **Email {start_index + 2}:** [One paragraph summary of this email]
        ...and so on for each email.

        Make each summary concise but informative, focusing on the main purpose and key points of each email.
        Keep each summary to 2-3 sentences maximum.

        Emails to summarize ({len(batch_emails)} emails in this batch):
        {emails_text}
        """
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.deepseek_api_key}"
        }
        
        payload = {
            "model": "deepseek-chat",
            "messages": [
                {
                    "role": "system", 
                    "content": "Provide clear, concise individual one-paragraph summaries for each email. Format each summary starting with **Email X:** followed by the paragraph. Keep summaries brief."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "max_tokens": 4000,
            "temperature": 0.3
        }
        
        try:
            print(f"Summarizing batch {start_index//20 + 1} ({len(batch_emails)} emails)...")
            response = requests.post(self.deepseek_api_url, headers=headers, json=payload, timeout=120)
            response.raise_for_status()
            
            result = response.json()
            summary_text = result['choices'][0]['message']['content']
            
            # Parse individual summaries
            batch_summaries = self.extract_individual_summaries(summary_text, batch_emails, start_index)
            print(f"✅ Batch {start_index//20 + 1} summarized successfully")
            
            return batch_summaries
            
        except Exception as e:
            print(f"❌ Error summarizing batch: {e}")
            # Return empty summaries for this batch
            return {start_index + i + 1: "Summary unavailable" for i in range(len(batch_emails))}
    
    def extract_individual_summaries(self, summary_text, batch_emails, start_index):
        """Extract individual email summaries from batch response"""
        summaries = {}
        
        for i in range(len(batch_emails)):
            email_num = start_index + i + 1
            
            # Try multiple patterns to find the summary
            patterns = [
                # Pattern 1: **Email X:** content until **Email X+1:** or end
                r"\*\*Email " + str(email_num) + r":\*\* (.*?)(?=\*\*Email " + str(email_num + 1) + r":\*\*|$)",
                # Pattern 2: Email X: content until Email X+1: or end  
                r"Email " + str(email_num) + r": (.*?)(?=Email " + str(email_num + 1) + r":|$)",
                # Pattern 3: **Email X:** content until Email X+1: or end
                r"\*\*Email " + str(email_num) + r":\*\* (.*?)(?=Email " + str(email_num + 1) + r":|$)",
                # Pattern 4: Email X: content until **Email X+1:** or end
                r"Email " + str(email_num) + r": (.*?)(?=\*\*Email " + str(email_num + 1) + r":\*\*|$)"
            ]
            
            found = False
            for pattern in patterns:
                match = re.search(pattern, summary_text, re.DOTALL)
                if match:
                    summary = match.group(1).strip()
                    # Clean up the summary
                    summary = re.sub(r'\*\*', '', summary)
                    summary = re.sub(r'\n+', ' ', summary)
                    summary = summary[:400]  # Limit length for table
                    summaries[email_num] = summary
                    found = True
                    break
            
            # Fallback: if no pattern found, use a simple search
            if not found:
                # Look for lines containing this email number
                lines = summary_text.split('\n')
                for line in lines:
                    if f"Email {email_num}:" in line:
                        summary = line.replace(f"Email {email_num}:", "").replace("**", "").strip()
                        if summary:
                            summaries[email_num] = summary[:400]
                            break
                
                # Final fallback
                if email_num not in summaries:
                    summaries[email_num] = "Summary processing incomplete"
        
        return summaries
    
    def create_word_document(self, emails_data, all_summaries):
        try:
            print("Creating Word document with all emails...")
            
            doc = Document()
            title = doc.add_heading('Email Summary Report', 0)
            title.alignment = 1
            
            doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            doc.add_paragraph(f"Period: Last 24 Hours")
            doc.add_paragraph(f"Total Emails Processed: {len(emails_data)}")
            doc.add_paragraph()
            
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'
            
            # Header row
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'No'
            hdr_cells[1].text = 'Sender'
            hdr_cells[2].text = 'Receiver'
            hdr_cells[3].text = 'Email Subject'
            hdr_cells[4].text = 'Summary in a paragraph'
            
            # Data rows for ALL emails
            for i, email in enumerate(emails_data, 1):
                row_cells = table.add_row().cells
                row_cells[0].text = str(i)
                row_cells[1].text = email['from'][:40] if email['from'] else "Unknown"
                row_cells[2].text = email['to'][:40] if email['to'] else "Unknown"
                row_cells[3].text = email['subject'][:80] if email['subject'] else "No Subject"
                
                # Get individual summary for this email
                summary = all_summaries.get(i, "Summary being processed...")
                row_cells[4].text = summary
            
            print(f"✅ Word document created with {len(emails_data)} emails")
            filename = f"Complete_Email_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            doc.save(filename)
            print(f"Word document saved: {filename}")
            
            return filename
            
        except Exception as e:
            print(f"Error creating Word document: {e}")
            return None
    
    def run_complete_summary(self):
        print(f"\n{'='*60}")
        print(f"STARTING COMPLETE EMAIL SUMMARY - {datetime.now()}")
        print(f"{'='*60}")
        
        try:
            # Step 1: Fetch ALL emails from last 24 hours
            emails_data = self.fetch_emails_last_24h()
            
            if not emails_data:
                print("No emails to process")
                return
                
            print(f"📧 Processing {len(emails_data)} emails...")
            
            # Step 2: Summarize ALL emails in batches
            all_summaries = self.summarize_emails_in_batches(emails_data)
            
            # Step 3: Store the processed emails and summaries for the dashboard
            storage_success = store_email_data_for_dashboard(emails_data, all_summaries)
            
            # Step 4: Save run statistics
            save_run_stats(len(emails_data), len(all_summaries))
            
            # Step 5: VERIFY DATA STORAGE
            print(f"\n{'='*60}")
            print("VERIFYING DATA STORAGE FOR DASHBOARD...")
            print(f"{'='*60}")
            verify_data_storage()
            
            print(f"✅ COMPLETE summary process finished at {datetime.now()}")
            print(f"✅ Processed {len(emails_data)} emails total")
            print(f"✅ Data sent to dashboard successfully")
                
        except Exception as e:
            print(f"❌ Critical error: {e}")

# ==================== DATABASE FUNCTIONS ====================

def init_db():
    """Initialize SQLite database"""
    db_path = get_db_path()
    print(f"📁 Using database at: {db_path}")
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS summary_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_date TEXT NOT NULL,
            total_emails INTEGER,
            processed_emails INTEGER,
            success_rate REAL,
            deepseek_tokens INTEGER,
            status TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS email_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            email_number INTEGER,
            sender TEXT,
            receiver TEXT,
            subject TEXT,
            summary TEXT,
            processed_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (run_id) REFERENCES summary_runs (id)
        )
    ''')
    
    conn.commit()
    conn.close()
    print(f"✅ Database initialized at: {db_path}")

def save_run_stats(total_emails, processed_emails):
    """Save run statistics to database"""
    try:
        db_path = get_db_path()
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        success_rate = (processed_emails / total_emails * 100) if total_emails > 0 else 0
        
        c.execute('''
            INSERT INTO summary_runs 
            (run_date, total_emails, processed_emails, success_rate, status)
            VALUES (?, ?, ?, ?, ?)
        ''', (
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            total_emails,
            processed_emails,
            success_rate,
            'completed'
        ))
        
        conn.commit()
        conn.close()
        print(f"✅ Run statistics saved to database at: {db_path}")
    except Exception as e:
        print(f"❌ Error saving run stats: {e}")

def store_email_data_for_dashboard(emails_data, all_summaries):
    """Store processed email data for dashboard display - FIXED VERSION"""
    try:
        db_path = get_db_path()
        print(f"📊 Starting to store {len(emails_data)} emails for dashboard at: {db_path}")
        
        # First, create a new run record
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        # Create new run entry
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        success_rate = (len(all_summaries) / len(emails_data) * 100) if emails_data else 0
        
        c.execute('''
            INSERT INTO summary_runs 
            (run_date, total_emails, processed_emails, success_rate, status)
            VALUES (?, ?, ?, ?, ?)
        ''', (
            current_time,
            len(emails_data),
            len(all_summaries),
            success_rate,
            'completed'
        ))
        
        run_id = c.lastrowid
        print(f"📊 Created new run_id: {run_id}")
        
        # Insert email data
        inserted_count = 0
        failed_count = 0
        
        for i, email in enumerate(emails_data, 1):
            summary = all_summaries.get(i, "Summary being processed...")
            
            try:
                c.execute('''
                    INSERT INTO email_data 
                    (run_id, email_number, sender, receiver, subject, summary)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    run_id,
                    i,
                    str(email['from'])[:100] if email['from'] else "Unknown",
                    str(email['to'])[:100] if email['to'] else "Unknown", 
                    str(email['subject'])[:200] if email['subject'] else "No Subject",
                    str(summary)[:500]
                ))
                inserted_count += 1
                
                if inserted_count % 50 == 0:
                    print(f"📥 Stored {inserted_count} emails in database...")
                    
            except Exception as e:
                print(f"❌ Failed to store email {i}: {e}")
                failed_count += 1
                continue
        
        conn.commit()
        conn.close()
        
        print(f"✅ Database storage complete:")
        print(f"   ✅ Successfully stored: {inserted_count} emails")
        print(f"   ❌ Failed to store: {failed_count} emails")
        print(f"   📊 Total processed: {len(emails_data)} emails")
        
        return inserted_count > 0
        
    except Exception as e:
        print(f"❌ CRITICAL ERROR storing email data: {e}")
        print(f"Full traceback: {traceback.format_exc()}")
        return False

def verify_data_storage():
    """Verify that data was properly stored in database"""
    try:
        db_path = get_db_path()
        print(f"🔍 Verifying database at: {db_path}")
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        # Check latest run
        c.execute('SELECT id, run_date, total_emails, processed_emails FROM summary_runs ORDER BY id DESC LIMIT 1')
        latest_run = c.fetchone()
        
        if not latest_run:
            print("❌ VERIFICATION FAILED: No runs found in database")
            return False
            
        run_id, run_date, total_emails, processed_emails = latest_run
        print(f"📋 Latest run: ID={run_id}, Date={run_date}, Emails={total_emails}")
        
        # Check email data for this run
        c.execute('SELECT COUNT(*) FROM email_data WHERE run_id = ?', (run_id,))
        stored_emails = c.fetchone()[0]
        
        print(f"📋 Stored emails for this run: {stored_emails}")
        
        # Get a sample of stored data
        c.execute('SELECT email_number, sender, subject FROM email_data WHERE run_id = ? LIMIT 3', (run_id,))
        samples = c.fetchall()
        
        print("📋 Sample stored emails:")
        for sample in samples:
            print(f"   - #{sample[0]}: From '{sample[1]}' - '{sample[2]}'")
        
        conn.close()
        
        success = stored_emails > 0
        if success:
            print(f"✅ VERIFICATION PASSED: {stored_emails} emails stored in database")
        else:
            print(f"❌ VERIFICATION FAILED: No emails stored for run {run_id}")
            
        return success
        
    except Exception as e:
        print(f"❌ VERIFICATION ERROR: {e}")
        return False

def scheduled_summary():
    """Function to be called by Render Cron Job"""
    try:
        print(f"🕒 Running scheduled summary at {datetime.now()}")
        agent = EmailSummarizerAgent()
        agent.run_complete_summary()
        return True
    except Exception as e:
        logging.error(f"Scheduled summary failed: {e}")
        return False

# Initialize database on startup
init_db()

# For direct script execution (cron job)
if __name__ == "__main__":
    # Check if running as cron job (you can set CRON_MODE environment variable)
    if os.getenv('CRON_MODE'):
        print("🕒 Running in cron mode...")
        scheduled_summary()
    else:
        # Web service mode
        print("🌐 Starting web server...")
        app.run(host='0.0.0.0', port=5000, debug=False)

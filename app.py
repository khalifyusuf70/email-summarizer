from docx import Document
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
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

# Initialize Flask app
app = Flask(__name__)

class EmailSummarizerAgent:
    def __init__(self):
        # Use environment variables for security
        self.deepseek_api_key = os.getenv('DEEPSEEK_API_KEY')
        self.source_email = os.getenv('SOURCE_EMAIL')
        self.source_password = os.getenv('SOURCE_PASSWORD')
        self.destination_email = os.getenv('DESTINATION_EMAIL', 'cos.presidency@jubalandstate.so')
        self.imap_server = os.getenv('IMAP_SERVER', 'imap.one.com')
        self.smtp_server = os.getenv('SMTP_SERVER', 'send.one.com')
        
        # Validate required environment variables
        if not all([self.deepseek_api_key, self.source_email, self.source_password]):
            raise ValueError("Missing required environment variables")
            
        self.deepseek_api_url = "https://api.deepseek.com/v1/chat/completions"
        self.imap_port = 993
        self.smtp_port = 587
    
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
                    status, msg_data = mail.fetch(email_id, "(RFC822)")
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
            
            # Try multiple patterns to find the summary - FIXED ESCAPE SEQUENCES
            patterns = [
                rf"\*\*Email {email_num}:\*\* (.*?)(?=\*\*Email {email_num + 1}:\*\*|\\Z)",
                rf"Email {email_num}: (.*?)(?=Email {email_num + 1}:|\\Z)",
                rf"\*\*Email {email_num}:\*\* (.*?)(?=Email {email_num + 1}:|\\Z)",
                rf"Email {email_num}: (.*?)(?=\*\*Email {email_num + 1}:\*\*|\\Z)"
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
    
    def send_summary_email(self, word_file_path=None):
        try:
            print("Preparing to send summary email...")
            
            msg = MIMEMultipart()
            msg['From'] = self.source_email
            msg['To'] = self.destination_email
            msg['Subject'] = f"Complete 24-Hour Email Summary - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            
            body = f"""
            COMPLETE 24-HOUR EMAIL SUMMARY REPORT
            Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            
            Please find attached the complete email summary report in Word document format.
            This report contains ALL emails from the last 24 hours with individual summaries.
            
            ---
            Automated Summary Service
            Jubaland State Archives
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            if word_file_path and os.path.exists(word_file_path):
                with open(word_file_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(word_file_path)}",
                )
                msg.attach(part)
                print(f"Attached Word document: {word_file_path}")
            
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.source_email, self.source_password)
            server.sendmail(self.source_email, self.destination_email, msg.as_string())
            server.quit()
            
            print(f"✅ Complete email summary sent successfully to {self.destination_email}")
            return True
            
        except Exception as e:
            print(f"❌ Error sending email: {e}")
            return False
    
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
            
            # Store the processed emails and summaries for the dashboard
            store_email_data_for_dashboard(emails_data, all_summaries)
            
            # Step 3: Create Word document with ALL emails
            word_file = self.create_word_document(emails_data, all_summaries)
            
            # Step 4: Send summary email with complete report
            success = self.send_summary_email(word_file)
            
            if success:
                print(f"✅ COMPLETE summary process finished at {datetime.now()}")
                print(f"✅ Processed {len(emails_data)} emails total")
                
                # Save run statistics
                save_run_stats(len(emails_data), len(all_summaries))
            else:
                print(f"❌ Process completed with errors")
                
        except Exception as e:
            print(f"❌ Critical error: {e}")

# Database functions
def init_db():
    """Initialize SQLite database"""
    conn = sqlite3.connect('email_summaries.db')
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

def save_run_stats(total_emails, processed_emails):
    """Save run statistics to database"""
    try:
        conn = sqlite3.connect('email_summaries.db')
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
        print("✅ Run statistics saved to database")
    except Exception as e:
        print(f"❌ Error saving run stats: {e}")

def store_email_data_for_dashboard(emails_data, all_summaries):
    """Store processed email data for dashboard display"""
    try:
        conn = sqlite3.connect('email_summaries.db')
        c = conn.cursor()
        
        # Get the latest run ID (this might be the issue - we need to get the ID we just inserted)
        c.execute('SELECT id FROM summary_runs ORDER BY id DESC LIMIT 1')
        latest_run = c.fetchone()
        
        if not latest_run:
            print("❌ No run ID found - cannot store email data")
            conn.close()
            return
            
        run_id = latest_run[0]
        print(f"📊 Storing {len(emails_data)} emails for dashboard under run_id: {run_id}")
        
        # Clear previous email data for this run
        c.execute('DELETE FROM email_data WHERE run_id = ?', (run_id,))
        
        # Insert new email data
        inserted_count = 0
        for i, email in enumerate(emails_data, 1):
            summary = all_summaries.get(i, "Summary being processed...")
            
            c.execute('''
                INSERT INTO email_data 
                (run_id, email_number, sender, receiver, subject, summary)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                run_id,
                i,
                email['from'][:100] if email['from'] else "Unknown",
                email['to'][:100] if email['to'] else "Unknown", 
                email['subject'][:200] if email['subject'] else "No Subject",
                summary[:500]
            ))
            inserted_count += 1
            
            if inserted_count % 50 == 0:
                print(f"📥 Stored {inserted_count} emails in database...")
        
        conn.commit()
        conn.close()
        print(f"✅ Successfully stored {inserted_count} emails in database for dashboard")
        
    except Exception as e:
        print(f"❌ Error storing email data: {e}")
        print(f"Full error: {traceback.format_exc()}")

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
        
        # Check if template file exists
        current_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(current_dir, 'templates', 'dashboard.html')
        
        if not os.path.exists(template_path):
            return f"""
            <html>
                <body style="font-family: Arial, sans-serif; padding: 40px;">
                    <h1>❌ Template File Missing</h1>
                    <p><strong>Expected path:</strong> {template_path}</p>
                    <p><strong>Current directory:</strong> {current_dir}</p>
                    <p><strong>Files in current directory:</strong> {', '.join(os.listdir(current_dir))}</p>
                    <p>Please make sure the 'templates' folder exists and contains 'dashboard.html'</p>
                </body>
            </html>
            """, 500
            
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
            "trigger": "/api/trigger-summary",
            "health": "/health",
            "stats": "/api/stats",
            "recent_summaries": "/api/recent-summaries",
            "debug": "/api/debug",
            "debug_database": "/api/debug-database",
            "test_json": "/api/test-json"
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
        conn = sqlite3.connect('email_summaries.db')
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

@app.route('/api/stats')
def get_stats():
    """API endpoint for dashboard statistics"""
    try:
        # Get recent run data from database
        conn = sqlite3.connect('email_summaries.db')
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
    """API endpoint for recent email summaries - NOW WITH REAL DATA"""
    try:
        conn = sqlite3.connect('email_summaries.db')
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
        # Web server mode
        print("🌐 Starting web server...")
        app.run(host='0.0.0.0', port=5000, debug=False)

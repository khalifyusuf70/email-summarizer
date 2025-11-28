import os
import sqlite3

def get_db_path():
    """Get database path that works for both web service and cron job"""
    # On Render, use /tmp directory which is shared between services
    if os.getenv('RENDER'):
        return '/tmp/email_summaries.db'
    else:
        return 'email_summaries.db'

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
    """Store processed email data for dashboard display"""
    try:
        db_path = get_db_path()
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        print(f"📊 Starting to store {len(emails_data)} emails for dashboard at: {db_path}")
        
        # Get the latest run ID 
        c.execute('SELECT id, run_date FROM summary_runs ORDER BY id DESC LIMIT 1')
        latest_run = c.fetchone()
        
        if not latest_run:
            print("❌ CRITICAL: No run ID found in summary_runs table!")
            print("   This means save_run_stats() didn't work properly")
            conn.close()
            return False
            
        run_id, run_date = latest_run
        print(f"📊 Found run_id: {run_id} from {run_date}")
        print(f"📊 We have {len(all_summaries)} summaries out of {len(emails_data)} emails")
        
        # Clear previous email data for this run
        delete_count = c.execute('DELETE FROM email_data WHERE run_id = ?', (run_id,)).rowcount
        if delete_count > 0:
            print(f"🗑️  Cleared {delete_count} previous email records for this run")
        
        # Insert new email data
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
                    email['from'][:100] if email['from'] else "Unknown",
                    email['to'][:100] if email['to'] else "Unknown", 
                    email['subject'][:200] if email['subject'] else "No Subject",
                    summary[:500]
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
        
        print(f"✅ Database storage complete at {db_path}:")
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

# Update the API routes to use the shared database
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
    """API endpoint for recent email summaries - NOW WITH REAL DATA"""
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

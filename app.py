from datetime import datetime, timedelta
from flask import Flask, jsonify, render_template, request, redirect
import imaplib, email, sqlite3, time, os, re, requests
from email.header import decode_header

app = Flask(__name__)

###################################################################################
# DATABASE PATH
###################################################################################
def get_db_path():
    return '/tmp/email_summaries.db' if os.getenv("RENDER") else "email_summaries.db"


###################################################################################
# DASHBOARD ROUTES
###################################################################################

@app.route('/')
def home():
    return redirect("/dashboard")

@app.route('/dashboard')
def dashboard():
    return render_template("dashboard.html")


@app.route("/api/stats")
def stats():
    try:
        conn=sqlite3.connect(get_db_path()); c=conn.cursor()
        c.execute("SELECT * FROM summary_runs ORDER BY id DESC LIMIT 1")
        row=c.fetchone(); conn.close()

        if row:
            return jsonify({
                "total_emails_today": row[2],
                "emails_processed": row[3],
                "success_rate": row[4],
                "last_run": row[1],
                "status": "active"
            })
        return jsonify({"status":"waiting","last_run":"Never"})
    except:
        return jsonify({"error":"stats failed"}),500


@app.route("/api/recent-summaries")
def recent():
    try:
        conn=sqlite3.connect(get_db_path()); c=conn.cursor()
        c.execute("SELECT id FROM summary_runs ORDER BY id DESC LIMIT 1")
        latest=c.fetchone()
        if not latest: return jsonify([])

        c.execute("""
            SELECT email_number, sender, receiver, subject, summary
            FROM email_data WHERE run_id=? ORDER BY email_number""",(latest[0],))
        rows=c.fetchall(); conn.close()

        return jsonify([{
            "number":r[0], "from":r[1], "to":r[2], "subject":r[3], "summary":r[4]} for r in rows
        ])

    except:
        return jsonify([])


###################################################################################
# MANUAL RUN
###################################################################################

@app.route("/api/trigger-manual", methods=["POST"])
def manual_trigger():
    try:
        agent = EmailSummarizer()
        import threading
        threading.Thread(target=agent.run,daemon=True).start()
        return jsonify({"status":"running"})
    except Exception as e:
        return jsonify({"error":str(e)}),500


###################################################################################
# EMAIL SUMMARIZER
###################################################################################

class EmailSummarizer:

    def __init__(self):
        self.api_key = os.getenv("DEEPSEEK_API_KEY")
        self.email = os.getenv("SOURCE_EMAIL")
        self.password = os.getenv("SOURCE_PASSWORD")
        self.server = os.getenv("IMAP_SERVER","imap.one.com")

        if not all([self.api_key,self.email,self.password]):
            raise ValueError("Missing credentials")

        self.url = "https://api.deepseek.com/v1/chat/completions"


    #############################################
    # FETCH EMAILS LAST 24H
    #############################################
    def fetch(self):
        try:
            mail=imaplib.IMAP4_SSL(self.server,993)
            mail.login(self.email,self.password)
            mail.select("inbox")

            since=(datetime.now()-timedelta(hours=24)).strftime("%d-%b-%Y")
            status,msgs=mail.search(None,f'(SINCE "{since}")')
            if status!="OK": return []

            emails=[]
            for i in msgs[0].split():
                _,data=mail.fetch(i,"(RFC822)")
                msg=email.message_from_bytes(data[0][1])
                emails.append({
                    "from":self.dec(msg["From"]),
                    "to":self.dec(msg["To"]),
                    "subject":self.dec(msg["Subject"]),
                    "body":self.body(msg)[:900]
                })
            mail.logout()
            return emails
        except:
            return []


    #############################################
    # UTILITIES
    #############################################
    def dec(self,h):
        if not h: return ""
        parts=decode_header(h); out=""
        for p,c in parts:
            out+=p.decode(c or "utf-8") if isinstance(p,bytes) else p
        return out

    def body(self,m):
        if m.is_multipart():
            for part in m.walk():
                if part.get_content_type()=="text/plain":
                    return part.get_payload(decode=True).decode("utf-8","ignore")
        return (m.get_payload(decode=True) or b"").decode("utf-8","ignore")


    #############################################
    # AI SUMMARIZATION
    #############################################
    def summarize_batch(self,emails,start):
        text=""
        for i,e in enumerate(emails,1):
            n=start+i
            text+=f"Email {n}:\nFrom:{e['from']}\nTo:{e['to']}\nSubject:{e['subject']}\n{e['body']}\n\n"

        r=requests.post(self.url,json={
            "model":"deepseek-chat",
            "messages":[{"role":"system","content":"summarize each email clearly"},
                        {"role":"user","content":text}],
            "temperature":0.3
        },headers={"Authorization":f"Bearer {self.api_key}"})
        
        out=r.json()['choices'][0]['message']['content']
        summaries={}
        for i in range(len(emails)):
            n=start+i+1
            m=re.search(rf"Email {n}:(.*?)(?=Email {n+1}:|$)",out,re.S)
            summaries[n]=(m.group(1).strip() if m else "No summary")
        return summaries


    #############################################
    # RUN FULL SUMMARY JOB
    #############################################
    def run(self):
        emails=self.fetch()
        if not emails: return

        summary={}
        for s in range(0,len(emails),20):
            summary.update(self.summarize_batch(emails[s:s+20],s))
            time.sleep(2)

        store_to_db(emails,summary)


###################################################################################
# DATABASE WRITE
###################################################################################

def store_to_db(emails,summaries):
    conn=sqlite3.connect(get_db_path()); c=conn.cursor()
    now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    c.execute("""
        INSERT INTO summary_runs(run_date,total_emails,processed_emails,success_rate,status)
        VALUES(?,?,?,?,?)""",(now,len(emails),len(summaries),
        (len(summaries)/len(emails))*100,"completed"))
    run_id=c.lastrowid

    for i,e in enumerate(emails,1):
        c.execute("""INSERT INTO email_data(run_id,email_number,sender,receiver,subject,summary)
                     VALUES(?,?,?,?,?,?)""",(run_id,i,e['from'],e['to'],e['subject'],summaries.get(i)))
    conn.commit(); conn.close()


###################################################################################
# CRON MODE ENABLED
###################################################################################

def cron_task():
    print("⏰ CRON AUTO RUN STARTED")
    agent=EmailSummarizer()
    agent.run()
    print("✔ CRON COMPLETE")

# Required when Render Scheduler calls `python app.py`
if __name__=="__main__":
    if os.getenv("CRON_MODE"):  # set CRON_MODE=1 in Render scheduler
        cron_task()
    else:
        app.run(host="0.0.0.0",port=5000)

import streamlit as st
import pandas as pd
from io import BytesIO
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import datetime
import time
import threading
import schedule

# --- User Authentication ---
USERS = {
    "Surekha Menon": "surekha123",
    "Harpreet Singh Dadial": "harpreet123",
    "Mayur Bharunde": "mayur123",
    "Prateek Thopte": "prateek123"  # New user added
}

# Email Configuration - Moved to backend
EMAIL_CONFIG = {
    "sender": "vishwa.s@mapledti.com",  # Replace with your Outlook email
    "password": "Paraj@1424",  # Replace with your email password
    "recipients": ["pratik.t@mapledti.com", "surekha.m@mapledti.com", "mayur.b@mapledti.com"],  # Replace with recipient emails
    "subject": "Daily Audit Report",
    "body": "Please find attached the daily audit report.",
    "smtp_server": "smtp.office365.com",
    "smtp_port": 587
}

# File paths
EXCEL_FILE_PATH = "audit_data.xlsx"

# Session State for storing entries and login
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "data" not in st.session_state:
    st.session_state.data = []
if "entry_submitted" not in st.session_state:
    st.session_state.entry_submitted = False
if "form_reset" not in st.session_state:
    st.session_state.form_reset = False

# --- Login Page ---
def login():
    st.title("Maple DTI Call Center Audit Form Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in USERS and USERS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success(f"Welcome, {username}!")
        else:
            st.error("Invalid username or password")

# --- Scoring Logic ---
def calculate_scores(parameter_responses):
    total_score = 0
    fatal_flag = False
    weights = {
        "Energetic Assumptive Opening": 5,
        "Call Opening with Smile": 5,
        "Acknowledgement": 3,
        "Purpose Of Call": 5,
        "Current Device Confirmation": -100,
        "Product Pitch Opening Script Adhered": 5,
        "Product USP Pitched": 5,
        "Upgrade Process Explained": 5,
        "Confirm Customer & Nearest Store  Location": 5,
        "Pricing Anchored": 3,
        "Pre Booking Pitched": -100,
        "Urgency Created to Pre-Book": 5,
        "Refund Pre Booking Script Adhered": 5,
        "Balance Payment  With Credit card informed": 5,
        "Assured Buy Back Script Adhered": -100,
        "Assured Buy Back Timeline Explained": -100,
        "Assured Buy-back T&C informed": -100,
        "Protection Accessories Pitched": -100,
        "Pitch Closure Script": 5,
        "Follow Up Date & Time Confirmed": -100,
        "Upgrade Store Visit Confirmation": 5,
        "Payment method Confirmation": 5,
        "WhatsApp Number Confirmed": 2,
        "Pre-Booking Template Sent on WhatsApp": 5,
        "Payment Confirmation": 5,
        "Post Payment Script": 5,
        "SPOC Details &  Data Transfer informed": 2,
        "Confirm Store Visit Date & Time": 5,
        "Call Closing Script Adhered": 5
    }
    
    for param, followed in parameter_responses.items():
        weight = weights.get(param, 0)
        score = 0
        if weight in [5, 3, 2]:
            if followed == "YES":
                score = weight
            elif followed == "NO":
                score = 0
            elif followed == "NA":
                score = 5
        elif weight == -100:
            if followed == "NO":
                fatal_flag = True
            score = 0
        total_score += score
        
    if fatal_flag:
        return 0, "Fail"
    else:
        max_possible_score = sum(w for w in weights.values() if w > 0)
        final_score = (total_score / max_possible_score) * 100
        call_status = "Pass" if final_score >= 85 else "Fail"
        return final_score, call_status

# --- Excel File Functions ---
def save_to_excel(data):
    """Save data to Excel file"""
    df = pd.DataFrame(data)
    
    # If file exists, try to merge with existing data
    if os.path.exists(EXCEL_FILE_PATH):
        try:
            existing_df = pd.read_excel(EXCEL_FILE_PATH)
            # Concatenate with existing data
            df = pd.concat([existing_df, df], ignore_index=True)
        except Exception as e:
            st.error(f"Error reading existing file: {e}")
    
    # Save to Excel
    try:
        df.to_excel(EXCEL_FILE_PATH, index=False, engine='xlsxwriter')
        return True
    except Exception as e:
        st.error(f"Error saving data: {e}")
        return False

# --- Email Functions ---
def send_email():
    """Send Excel file as email attachment"""
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"File {EXCEL_FILE_PATH} not found.")
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG["sender"]
        msg['To'] = ", ".join(EMAIL_CONFIG["recipients"])
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = f"{EMAIL_CONFIG['subject']} - {datetime.date.today()}"
        
        msg.attach(MIMEText(EMAIL_CONFIG["body"]))
        
        # Attach Excel file
        with open(EXCEL_FILE_PATH, "rb") as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(EXCEL_FILE_PATH)}"')
            msg.attach(part)
        
        # Connect to server and send email
        server = smtplib.SMTP(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"])
        server.starttls()
        server.login(EMAIL_CONFIG["sender"], EMAIL_CONFIG["password"])
        server.send_message(msg)
        server.quit()
        
        print(f"Email sent successfully to {', '.join(EMAIL_CONFIG['recipients'])} at {datetime.datetime.now()}")
        
        # Create a backup file with date in filename after successful email
        today = datetime.date.today().strftime("%Y-%m-%d")
        backup_file = f"audit_data_{today}.xlsx"
        if os.path.exists(EXCEL_FILE_PATH):
            df = pd.read_excel(EXCEL_FILE_PATH)
            df.to_excel(backup_file, index=False)
            print(f"Backup created: {backup_file}")
            
        return True
        
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False

def schedule_daily_email():
    """Schedule email to be sent at 7PM each day"""
    def job():
        print(f"Running scheduled email job at {datetime.datetime.now()}...")
        send_email()
    
    # Schedule to run daily at 7PM
    schedule.every().day.at("19:00").do(job)
    
    # Run the scheduler in a separate thread
    def run_scheduler():
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
            
    # Start the scheduler thread
    email_thread = threading.Thread(target=run_scheduler, daemon=True)
    email_thread.start()
    print("Email scheduler started - emails will be sent daily at 7:00 PM")

# --- Form Entry Page ---
def form_page():
    st.title("Audit Entry Form")
    
    # Handle clearing form
    if st.button("Clear Form"):
        # Clear all form-related values in session state
        for key in list(st.session_state.keys()):
            if key not in ["logged_in", "username", "data", "form_reset"]:
                if key in st.session_state:
                    del st.session_state[key]
        st.session_state.form_reset = True
        st.experimental_rerun()
    
    with st.form("entry_form", clear_on_submit=True):
        # Form fields with empty default values if form was reset
        mobile_number = st.text_input("Mobile Number", value="")
        call_date = st.date_input("Call Date")
        call_duration = st.text_input("Call Duration", value="")
        
        if st.session_state.username == "Surekha Menon":
            caller_name = st.selectbox("Caller Name", ["Diya", "Bhagyashree", "Harpreet", "Saud", "Allwyn", "Hussain", "Pooja", "Noor", "Jayesh"])
        else:
            caller_name = st.selectbox("Caller Name", ["Harpreet"])
            
        customer_name = st.text_input("Customer Name", value="")
        reporting_manager = st.selectbox("Reporting Manager", ["Prateek", "Surekha"])
        channel = st.selectbox("Channel", ["Call", "Walk-in", "Online"])
        department = st.selectbox("Department", ["Ecom INB", "EUP Outbound", "Meta Buy-Back","CSAT Campaign"])
        skill = st.selectbox("Skill", ["Voice", "Non-Voice"])
        auditor_name = st.selectbox("Auditor Name", ["Surekha", "Prateek", "Mayur"])
        type_of_audit = st.selectbox("Type of Audit", ["Inbound Audit", "Outbound Call", "WhatsApp Chat", "Email"])
        data_source = st.selectbox("Data Source", ["Maple", "Iplanet"])
        call_type = st.selectbox("Call Type", ["Lead", "Non-Lead", "Prospect", "Not-interested", "Not Eligible"])
        dialer_campaign_name = st.selectbox("Dialer Campaign Name", ["CSAT Campaign","Trade-in Campaign","Other","iPhone 11", "iPhone 12", "iPhone 13", "iPhone 14", "iPhone 15", "iPhone 16"])
        call_recording_link = st.text_input("Call Recording Link", value="")
        impact = st.text_input("Impact", value="")
        purpose_of_call = st.text_input("Purpose of Call", value="")
        callers_response = st.text_input("Caller's Response", value="")
        call_gist = st.text_area("Call Gist", value="")
        aoi = st.text_area("Area of Improvement (AOI)", value="")
        fatal = st.selectbox("Fatal", ["Yes", "No", "NA"])
        service_recovery = st.selectbox("Service Recovery", ["Yes", "No", "NA"])
        
        # Parameters of Audit section
        st.markdown("### Parameters of Audit")
        parameters = [
            "Energetic Assumptive Opening", "Call Opening with Smile", "Acknowledgement", "Purpose Of Call",
            "Current Device Confirmation", "Product Pitch Opening Script Adhered", "Product USP Pitched",
            "Upgrade Process Explained", "Confirm Customer & Nearest Store  Location", "Pricing Anchored",
            "Pre Booking Pitched", "Urgency Created to Pre-Book", "Refund Pre Booking Script Adhered",
            "Balance Payment  With Credit card informed", "Assured Buy Back Script Adhered",
            "Assured Buy Back Timeline Explained", "Assured Buy-back T&C informed", "Protection Accessories Pitched",
            "Pitch Closure Script", "Follow Up Date & Time Confirmed", "Upgrade Store Visit Confirmation",
            "Payment method Confirmation", "WhatsApp Number Confirmed", "Pre-Booking Template Sent on WhatsApp",
            "Payment Confirmation", "Post Payment Script", "SPOC Details &  Data Transfer informed",
            "Confirm Store Visit Date & Time", "Call Closing Script Adhered"
        ]
        
        parameter_responses = {}
        for param in parameters:
            parameter_responses[param] = st.selectbox(param, ["YES", "NO", "NA"])
        
        # Submit button
        submitted = st.form_submit_button("Submit Entry")
        
        if submitted:
            final_score, call_status = calculate_scores(parameter_responses)
            new_entry = {
                "Mobile Number": mobile_number,
                "Call Date": str(call_date),
                "Call Duration": call_duration,
                "Caller Name": caller_name,
                "Customer Name": customer_name,
                "Reporting Manager": reporting_manager,
                "Channel": channel,
                "Department": department,
                "Skill": skill,
                "Auditor Name": auditor_name,
                "Type of Audit": type_of_audit,
                "Data Source": data_source,
                "Call Type": call_type,
                "Fatal": fatal,
                "Dialer Campaign Name": dialer_campaign_name,
                "Call Recording Link": call_recording_link,
                "Impact": impact,
                "Purpose of Call": purpose_of_call,
                "Caller's Response": callers_response,
                "Call Gist": call_gist,
                "AOI": aoi,
                "Service Recovery": service_recovery,
                "Final Score %": final_score,
                "Call Status": call_status,
                "Entry Date": str(datetime.date.today()),
                "Entry Time": datetime.datetime.now().strftime("%H:%M:%S")
            }
            new_entry.update(parameter_responses)
            st.session_state.data.append(new_entry)
            
            # Save to Excel file
            if save_to_excel([new_entry]):
                st.success("Entry submitted successfully and saved to Excel!")
            else:
                st.warning("Entry submitted but failed to save to Excel.")
    
    # Display the audit data
    if st.session_state.data:
        df = pd.DataFrame(st.session_state.data)
        st.dataframe(df)
        filtered_df = df
        username = st.session_state.username
        if username == "Harpreet Singh Dadial":
            filtered_df = df[df["Reporting Manager"] == "Surekha"]
        elif username == "Mayur Bharunde":
            filtered_df = df[df["Auditor Name"] == "Mayur"]
        output = BytesIO()
        filtered_df.to_excel(output, index=False, engine='xlsxwriter')
        st.download_button("Download Filtered Data as Excel", data=output.getvalue(), file_name="audit_data_filtered.xlsx")

    # Show email schedule information
    with st.expander("Daily Email Information"):
        st.info("The system is configured to automatically send the audit data to all recipients daily at 7:00 PM.")
        st.write("Recipients: " + ", ".join(EMAIL_CONFIG["recipients"]))
        
        # Calculate time until next email
        now = datetime.datetime.now()
        next_run = datetime.datetime.combine(
            now.date() if now.hour < 19 else now.date() + datetime.timedelta(days=1),
            datetime.time(19, 0)
        )
        time_until = next_run - now
        hours, remainder = divmod(time_until.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        
        st.write(f"Next scheduled email: Today at 7:00 PM" if now.hour < 19 else f"Next scheduled email: Tomorrow at 7:00 PM")
        st.write(f"Time until next email: {hours} hours, {minutes} minutes")

# Create a separate python file for scheduled tasks
def create_scheduler_script():
    with open("email_scheduler.py", "w") as f:
        f.write("""
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import datetime
import time
import schedule

# Email Configuration
EMAIL_CONFIG = {
    "sender": "YOUR_OUTLOOK_EMAIL@outlook.com",  # Replace with your actual email
    "password": "YOUR_EMAIL_PASSWORD",  # Replace with your actual password
    "recipients": ["recipient1@example.com", "recipient2@example.com", "recipient3@example.com"],  # Replace with actual recipients
    "subject": "Daily Audit Report",
    "body": "Please find attached the daily audit report.",
    "smtp_server": "smtp.office365.com",
    "smtp_port": 587
}

# File path
EXCEL_FILE_PATH = "audit_data.xlsx"

def send_email():
    \"\"\"Send Excel file as email attachment\"\"\"
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"File {EXCEL_FILE_PATH} not found.")
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG["sender"]
        msg['To'] = ", ".join(EMAIL_CONFIG["recipients"])
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = f"{EMAIL_CONFIG['subject']} - {datetime.date.today()}"
        
        msg.attach(MIMEText(EMAIL_CONFIG["body"]))
        
        # Attach Excel file
        with open(EXCEL_FILE_PATH, "rb") as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(EXCEL_FILE_PATH)}"')
            msg.attach(part)
        
        # Connect to server and send email
        server = smtplib.SMTP(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"])
        server.starttls()
        server.login(EMAIL_CONFIG["sender"], EMAIL_CONFIG["password"])
        server.send_message(msg)
        server.quit()
        
        print(f"Email sent successfully to {', '.join(EMAIL_CONFIG['recipients'])} at {datetime.datetime.now()}")
        
        # Create a backup file with date in filename
        today = datetime.date.today().strftime("%Y-%m-%d")
        backup_file = f"audit_data_{today}.xlsx"
        if os.path.exists(EXCEL_FILE_PATH):
            df = pd.read_excel(EXCEL_FILE_PATH)
            df.to_excel(backup_file, index=False)
            print(f"Backup created: {backup_file}")
            
        return True
        
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False

def main():
    print("Daily Email Scheduler Started")
    print(f"Email will be sent at 7:00 PM to: {', '.join(EMAIL_CONFIG['recipients'])}")
    
    # Schedule to run daily at 7PM
    schedule.every().day.at("19:00").do(send_email)
    
    # Run once at startup if it's after work hours and no email was sent today
    current_time = datetime.datetime.now()
    if current_time.hour >= 19:
        today_backup = f"audit_data_{datetime.date.today().strftime('%Y-%m-%d')}.xlsx"
        if not os.path.exists(today_backup):
            print("Running initial email send...")
            send_email()
    
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute

if __name__ == "__main__":
    main()
""")

# Initialize scheduler when app starts
schedule_daily_email()

# Create the separate scheduler script (only happens once)
if not os.path.exists("email_scheduler.py"):
    create_scheduler_script()

# Main app logic
if not st.session_state.logged_in:
    login()
else:
    form_page()
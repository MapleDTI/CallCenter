import streamlit as st
import pandas as pd
import requests
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
    "Mayur Bhorunde": "mayur123",
    "Pratik Thopate": "pratik123"  # New user added
}

# --- Email Configuration - Backend Configuration ---
EMAIL_CONFIG = {
    "sender": "vishwa.s@mapledti.com",  # Replace with your actual Outlook email
    "password": "Paraj@1424",  # Replace with your actual email password
    "recipients": ["mayur.b@mapledti.com", "pratik.t@mapledti.com", "surekha.m@mapledti.com"],  # Replace with actual recipient emails
    "subject": "Daily Audit Report",
    "body": "Please find attached the daily audit report with all audits completed today.",
    "smtp_server": "smtp.office365.com",
    "smtp_port": 587
}

# --- File paths ---
EXCEL_FILE_PATH = "audit_data.xlsx"

# --- Flag to track if email has been sent today ---
EMAIL_SENT_TODAY_FLAG = f"email_sent_{datetime.date.today()}.flag"

# --- Session State for storing entries and login ---
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
if "email_status" not in st.session_state:
    st.session_state.email_status = "Not sent yet today"

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
                score = weight
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
def save_to_excel(data_to_save):
    """
    Saves the provided data (list of dictionaries) to the Excel file.
    Appends data if the file exists, otherwise creates a new file.
    Includes basic backup mechanism on read error.
    """
    # Convert the list of new entries (usually just one) to a DataFrame
    new_df = pd.DataFrame(data_to_save)

    # Create a combined DataFrame (either new or appended)
    combined_df = pd.DataFrame()

    # Check if the Excel file already exists
    if os.path.exists(EXCEL_FILE_PATH):
        try:
            # Read the existing data
            existing_df = pd.read_excel(EXCEL_FILE_PATH)
            # Concatenate existing data with new data
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        except Exception as e:
            print(f"Error reading existing file: {e}. Attempting to back up.")
            # If reading fails, try to back up the corrupted file
            backup_file = f"audit_data_backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            try:
                os.rename(EXCEL_FILE_PATH, backup_file)
                print(f"Created backup of potentially corrupted file: {backup_file}")
                # Since reading failed, start fresh with only the new data
                combined_df = new_df
            except Exception as backup_error:
                print(f"Failed to create backup: {backup_error}")
                # Could not back up, proceed with caution (might overwrite)
                # Or decide to raise an error here
                combined_df = new_df # Or handle error more strictly
    else:
        # If the file doesn't exist, the new data is the only data
        combined_df = new_df

    # Save the combined data back to the Excel file
    try:
        # Use xlsxwriter engine as specified in your original code
        # index=False prevents writing the pandas DataFrame index as a column
        combined_df.to_excel(EXCEL_FILE_PATH, index=False, engine='xlsxwriter')
        print(f"Data successfully saved to {EXCEL_FILE_PATH}")
        return True
    except Exception as e:
        print(f"Error saving data to Excel: {e}")
        return False

# --- Function to prepare daily audit data ---
def prepare_daily_audit_data():
    """Prepare Excel file with only today's audits"""
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"File {EXCEL_FILE_PATH} not found.")
        return None

    try:
        # Read the full Excel file
        full_df = pd.read_excel(EXCEL_FILE_PATH)

        # Filter for today's date only
        today_str = str(datetime.date.today())
        daily_df = full_df[full_df["Entry Date"] == today_str]

        if daily_df.empty:
            print("No audit entries found for today.")
            return None

        # Create a temporary file for today's data
        temp_file = f"daily_audit_{today_str}.xlsx"
        daily_df.to_excel(temp_file, index=False, engine='xlsxwriter')
        return temp_file

    except Exception as e:
        print(f"Error preparing daily audit data: {e}")
        return None

# --- Email Functions ---
def send_daily_email():
    """Send Excel file with today's audits as email attachment"""
    # Check if email has already been sent today
    if os.path.exists(EMAIL_SENT_TODAY_FLAG):
        print(f"Email already sent today. Skipping.")
        return False

    # Prepare daily audit data
    daily_file = prepare_daily_audit_data()
    if not daily_file:
        print("No data to send or error preparing data.")
        return False

    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG["sender"]
        msg['To'] = ", ".join(EMAIL_CONFIG["recipients"])
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = f"{EMAIL_CONFIG['subject']} - {datetime.date.today()}"

        # Add summary of today's audits
        try:
            df = pd.read_excel(daily_file)
            num_entries = len(df)
            num_pass = len(df[df["Call Status"] == "Pass"])
            num_fail = len(df[df["Call Status"] == "Fail"])
            avg_score = df["Final Score %"].mean()

            summary = (
                f"Daily Audit Summary for {datetime.date.today()}:\n\n"
                f"Total Audits: {num_entries}\n"
                f"Pass: {num_pass}\n"
                f"Fail: {num_fail}\n"
                f"Average Score: {avg_score:.2f}%\n\n"
                f"{EMAIL_CONFIG['body']}"
            )
            msg.attach(MIMEText(summary))
        except Exception as e:
            # Fall back to default message if summary creation fails
            msg.attach(MIMEText(EMAIL_CONFIG["body"]))
            print(f"Error creating email summary: {e}")

        # Attach Excel file
        with open(daily_file, "rb") as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="Daily_Audit_Report_{datetime.date.today()}.xlsx"')
            msg.attach(part)

        # Connect to server and send email
        server = smtplib.SMTP(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"])
        server.starttls()
        server.login(EMAIL_CONFIG["sender"], EMAIL_CONFIG["password"])
        server.send_message(msg)
        server.quit()

        # Create flag file to indicate email was sent today
        with open(EMAIL_SENT_TODAY_FLAG, 'w') as flag_file:
            flag_file.write(f"Email sent at {datetime.datetime.now()}")

        # Update status
        st.session_state.email_status = f"Sent at {datetime.datetime.now().strftime('%H:%M:%S')}"

        print(f"Daily email sent successfully to {', '.join(EMAIL_CONFIG['recipients'])} at {datetime.datetime.now()}")

        # Clean up temporary file
        try:
            os.remove(daily_file)
        except:
            pass

        return True

    except Exception as e:
        print(f"Failed to send email: {e}")
        return False

def reset_email_flag():
    """Reset the email sent flag at midnight"""
    if os.path.exists(EMAIL_SENT_TODAY_FLAG):
        try:
            os.remove(EMAIL_SENT_TODAY_FLAG)
            print(f"Email flag reset at {datetime.datetime.now()}")
        except Exception as e:
            print(f"Error resetting email flag: {e}")



def schedule_daily_email():
    """Schedule email to be sent at 7PM every day and reset flag at midnight"""
    def email_job():
        print(f"Running scheduled email job at {datetime.datetime.now()}...")
        send_daily_email()

    def reset_job():
        print(f"Running reset job at {datetime.datetime.now()}...")
        reset_email_flag()
        st.session_state.email_status = "Not sent yet today"

    # Schedule to run daily at 7PM
    schedule.every().day.at("19:00").do(email_job)

    # Schedule to reset email flag at midnight
    schedule.every().day.at("00:01").do(reset_job)

    # Run the scheduler in a separate thread
    def run_scheduler():
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute

    # Start the scheduler thread
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()
    print("Email scheduler started - will send daily report at 7:00 PM")

    # Check if we need to reset the flag (in case the app restarts on a new day)
    if os.path.exists(EMAIL_SENT_TODAY_FLAG):
        with open(EMAIL_SENT_TODAY_FLAG, 'r') as f:
            sent_date_str = f.read().strip()
            if datetime.date.today().strftime("%Y-%m-%d") not in sent_date_str:
                reset_email_flag()


# --- Form Entry Page ---
def form_page():
    st.title("Audit Entry Form")

    # Handle clearing form
    if st.button("Clear Form"):
        # Clear all form-related values in session state
        for key in list(st.session_state.keys()):
            if key not in ["logged_in", "username", "data", "form_reset", "email_status"]:
                if key in st.session_state:
                    del st.session_state[key]
        st.session_state.form_reset = True
        st.rerun()

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
        department = st.selectbox("Department", ["Ecom INB", "EUP Outbound", "Meta Buy-Back"])
        skill = st.selectbox("Skill", ["Voice", "Non-Voice"])
        auditor_name = st.selectbox("Auditor Name", ["Surekha", "Prateek", "Mayur"])
        type_of_audit = st.selectbox("Type of Audit", ["Inbound Audit", "Outbound Call", "WhatsApp Chat", "Email"])
        data_source = st.selectbox("Data Source", ["Maple", "Iplanet","CSAT"])
        call_type = st.selectbox("Call Type", ["Lead", "Non-Lead", "Prospect", "Not-interested", "Not Eligible"])
        dialer_campaign_name = st.selectbox("Dialer Campaign Name", ["iPhone 11", "iPhone 12", "iPhone 13", "iPhone 14", "iPhone 15", "iPhone 16","Others","Trade-in Campaign","CSAT Campaign"])
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
        elif username == "Mayur Bhorunde":
            filtered_df = df[df["Auditor Name"] == "Mayur"]
        output = BytesIO()
        filtered_df.to_excel(output, index=False, engine='xlsxwriter')
        st.download_button("Download Filtered Data as Excel", data=output.getvalue(), file_name="audit_data_filtered.xlsx")

    # Add email status and manual send option for admins
    col1, col2 = st.columns(2)

    with col1:
        st.info(f"Daily email status: {st.session_state.email_status}")
        st.info("Next scheduled email: Today at 7:00 PM")

    with col2:
        # Manual send option for admins
        if st.session_state.username in ["Surekha Menon", "Pratik Thopte"]:
            if st.button("Send Today's Report Now"):
                if os.path.exists(EMAIL_SENT_TODAY_FLAG):
                    st.warning("Email already sent today. Delete the flag file to send again.")
                else:
                    with st.spinner("Sending email..."):
                        if send_daily_email():
                            st.success("Email sent successfully!")
                        else:
                            st.error("Failed to send email. Check logs for details.")

# Initialize scheduler when app starts
schedule_daily_email()

# Main app logic
if not st.session_state.logged_in:
    login()
else:
    form_page()

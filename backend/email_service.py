import os
import tempfile
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def send_missing_files_email(recipient_email, missing_files, group_name):
    """
    Sends an email notification regarding missing files in a specific group.
    """
    if not recipient_email or not missing_files:
        return False, "No recipient email or missing files provided."
        
    sender_email = "iwk.certificate.system@example.com"
    subject = f"⚠️ [Action Required] Incomplete Folder Content for {group_name}"
    
    body = f"Hello,\n\nThe system detected that the following files are missing for the group '{group_name}':\n\n"
    for f in missing_files:
        body += f"- {f}\n"
    
    body += "\nPlease review the uploaded files and the BOM.\n\nThank you,\nIWK Certificate System"
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        print(f"--- MOCK EMAIL SENT ---\nTo: {recipient_email}\nSubject: {subject}\nBody: {body}\n-----------------------")
        return True, "Email sent successfully (Mocked)."
    except Exception as e:
        return False, str(e)

def generate_etk_eml(stats, excel_bytes, order_no="T/XXXXX", sender_name="Documentation Officer", recipient_email=""):
    """
    Generates a downloadable .eml email file containing subject, body, and attached Excel report.
    Opening this file launches Outlook/Thunderbird directly with pre-populated draft and attachment.
    """
    mode = stats.get("Mode", "CSP2 BOM")
    proj_rows = stats.get("Project Rows", 0)
    bom_rows = stats.get("BOM Rows", 0)
    matched = stats.get("Matched Rows", 0)
    
    if mode == "CSP2 BOM":
        subject = f"Project PLAN vs CSP2 Comparison Summary {order_no}"
        body = f"""Hello Team,

Please find attached the comparison result of PROJECT PLAN vs CSP2 BOM.

From Project plan: {proj_rows} Gr.
From CSP2 BOM (after filter): {bom_rows} Gr.
Matched rows: {matched} Gr.

- Kindly update the information accordingly. 

Notes:
- Not found in Project plan means exist in CSP2 but not in Project.

Best regards,
{sender_name}
Documentation Officer"""
    else:
        subject = f"Project PLAN vs CSPB Comparison ETK Summary: {order_no}"
        body = f"""Hello Team,

Please find below the ETK summary of the PROJECT PLAN vs CSPB BOM comparison.

From Project plan: {proj_rows} Gr.
From CSPB BOM (excluding F): {bom_rows} Gr.
Matched rows: {matched} Gr.

Kindly update the information accordingly.

Best regards,
{sender_name}
Documentation Officer"""

    msg = MIMEMultipart()
    msg['Subject'] = subject
    if recipient_email:
        msg['To'] = recipient_email
        
    msg.attach(MIMEText(body, 'plain'))
    
    filename = f"Project_vs_{mode.replace(' ', '_')}_Summary.xlsx"
    part = MIMEApplication(excel_bytes, Name=filename)
    part['Content-Disposition'] = f'attachment; filename="{filename}"'
    msg.attach(part)
    
    return subject, body, msg.as_bytes()

def open_outlook_draft(stats, excel_bytes, order_no="T/XXXXX", sender_name="Documentation Officer", recipient_email=""):
    """
    Directly launches Outlook desktop app with draft email and attachment (matching VBA .Display).
    """
    mode = stats.get("Mode", "CSP2 BOM")
    proj_rows = stats.get("Project Rows", 0)
    bom_rows = stats.get("BOM Rows", 0)
    matched = stats.get("Matched Rows", 0)
    
    if mode == "CSP2 BOM":
        subject = f"Project PLAN vs CSP2 Comparison Summary {order_no}"
        body = f"""Hello Team,

Please find attached the comparison result of PROJECT PLAN vs CSP2 BOM.

From Project plan: {proj_rows} Gr.
From CSP2 BOM (after filter): {bom_rows} Gr.
Matched rows: {matched} Gr.

- Kindly update the information accordingly. 

Notes:
- Not found in Project plan means exist in CSP2 but not in Project.

Best regards,
{sender_name}
Documentation Officer"""
    else:
        subject = f"Project PLAN vs CSPB Comparison ETK Summary: {order_no}"
        body = f"""Hello Team,

Please find below the ETK summary of the PROJECT PLAN vs CSPB BOM comparison.

From Project plan: {proj_rows} Gr.
From CSPB BOM (excluding F): {bom_rows} Gr.
Matched rows: {matched} Gr.

Kindly update the information accordingly.

Best regards,
{sender_name}
Documentation Officer"""

    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0) # 0 = olMailItem
        mail.Subject = subject
        mail.Body = body
        if recipient_email:
            mail.To = recipient_email
            
        temp_dir = tempfile.gettempdir()
        temp_filename = f"Project_vs_{mode.replace(' ', '_')}_Summary.xlsx"
        temp_filepath = os.path.join(temp_dir, temp_filename)
        
        with open(temp_filepath, "wb") as f:
            f.write(excel_bytes)
            
        mail.Attachments.Add(temp_filepath)
        mail.Display() # Opens Outlook draft window
        return True, "Outlook draft opened successfully!"
    except Exception as e:
        return False, f"Unable to launch Outlook directly: {str(e)}"

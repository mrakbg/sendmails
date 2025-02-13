import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Load the Excel file with no header, assuming data starts from B1 and C1
df = pd.read_excel('contacts2.xlsx', header=None)

# Manually set column names
df.columns = ['Email', 'Name']

# Email configuration
smtp_server = 'smtp.gmail.com'  # Gmail SMTP server
smtp_port = 587  # Port for TLS
sender_email = 'anujgupt869@gmail.com'  # Your email address
password = 'kifk zvfn wjsn ctcp'  # Your app password

# Create a SMTP session
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()  # Upgrade to secure connection
server.login(sender_email, password)  # Log in to your email account

# Path to your resume
resume_path = 'AnujResume.pdf'  # Update with the actual path to your resume

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    name = row['Name']  # Access the 'Name' column
    recipient_email = row['Email']  # Access the 'Email' column
    
    # Compose the email in HTML format
    subject = 'Let’s Connect: DevOps!'
    body = f"""
    <html>
    <body>
        <p>Hello {name},</p>

        <p>I hope this email finds you well.</p>

        <p>I’m Anuj Gupta, with over 3.5+ years of hands-on experience in DevOps and Cloud technologies like
        GCP, Kubernetes, Docker, Jenkins, Terraform, Bash scripting, and Linux Systems.</p>

        <p>I’ve attached my resume for your review. I appreciate your time and look forward to any potential openings!</p>

        <p>Thanks & Regards,<br>
        Anuj Gupta<br>
        Notice Period: LWD 20 Feb<br>
        Email: guptnuj@gmail.com<br>
        Phone: +91 9908992784</p>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))  # Attach the email body as HTML

    # Attach the resume
    with open(resume_path, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name='resume.pdf')
        part['Content-Disposition'] = 'attachment; filename="AnujDevOps.pdf"'
        msg.attach(part)

    # Send the email
    try:
        server.send_message(msg)
        print(f'Email sent to {name} at {recipient_email}')
    except Exception as e:
        print(f'Failed to send email to {name}: {e}')

# Close the SMTP session
server.quit()

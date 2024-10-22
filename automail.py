import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Load the Excel file
df = pd.read_excel('contacts.xlsx')  # Update the filename if necessary

# Email configuration
smtp_server = 'smtp.gmail.com'  # Gmail SMTP server
smtp_port = 587  # Port for TLS
sender_email = 'guptanujk@gmail.com'  # Your email address
password = 'password'  # Your app password

# Create a SMTP session
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()  # Upgrade to secure connection
server.login(sender_email, password)  # Log in to your email account

# Path to your resume
resume_path = 'AnujResume.pdf'  # Update with the actual path to your resume

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    name = row[df.columns[1]]  # Name column
    recipient_email = row[df.columns[0]]  # Email column
    
    # Compose the email
    subject = 'Let’s Connect: DevOps!'
    body = f"""
    Hello {name},

    I hope you’re doing well.

    I’m Anuj Gupta, with over 3+ years of hands-on experience in DevOps and Cloud technologies like
    GCP, Kubernetes, Docker, Jenkins, Terraform, Bash scripting and linux Systems.

    I’ve attached my resume for your review. I appreciate your time and look forward to any potential openings!

    Best Regards,
    Anuj Gupta
    """

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))  # Attach the email body

    # Attach the resume
    with open(resume_path, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name='resume.pdf')
        part['Content-Disposition'] = 'attachment; filename="resume.pdf"'
        msg.attach(part)

    # Send the email
    try:
        server.send_message(msg)
        print(f'Email sent to {name} at {recipient_email}')
    except Exception as e:
        print(f'Failed to send email to {name}: {e}')

# Close the SMTP session
server.quit()

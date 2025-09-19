import pandas as pd
from jinja2 import Template
import smtplib
from email.mime.text import MIMEText

# Load data
df = pd.read_excel("recipients.xlsx")
# Email template
template_str = """
Hi {{ name }},
Here's your weekly update: {{ report_link }}.
Best regards,
{{ sender_name }}
"""
template = Template(template_str)
# SMTP setup
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login("your_email@gmail.com", "your_password")
# Compile and send emails
for _, row in df.iterrows():
    message = template.render(
        name=row['name'],
        report_link=row['report_link'],
        sender_name="Amit"
    )
    msg = MIMEText(message, "plain")
    msg["Subject"] = "Weekly Update"
    msg["From"] = "your_email@gmail.com"
    msg["To"] = row['email']
    server.sendmail("your_email@gmail.com", row['email'], msg.as_string())
server.quit()
print("All emails sent successfully!")
# main.py
from fastapi import FastAPI
from flask import Flask
import openpyxl
import smtplib
import waitress
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

app = Flask(__name__)
api = FastAPI()


# Load Excel data
def load_excel_data(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)
    return data


# Send email
def send_email(subject, message, to_email):
    from_email = "ruksharkhann00@gmail.com"
    password = "fgrcaqaykcgucgkk"
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    server = smtplib.SMTP('smtp.example.com', 587)  # Update with your SMTP server details
    server.starttls()
    server.login(from_email, password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()


# Process absences and send emails
@api.post("/process-absences")
async def process_absences(payload: dict):
    file_path = payload['Employee.xlsx']
    data = load_excel_data(file_path)

    for row in data[1:]:
        employee_name = row[0]
        absent_days = row[1]

        if absent_days >= 3:
            # Send email to employer and candidate
            employer_email = "ruksharkhan00@gmail.com"  # Update with employer's email
            candidate_email = "ruksharparveen527@gmail.com"  # Update with candidate's email
            subject = f"Employee Absent Alert - {employee_name}"
            message = (f"This email is to inform you that according to attendance records, John did not attend to the "
                       f"duties for {absent_days} days. \n Thanks")

            send_email(subject, message, employer_email)
            send_email(subject, message, candidate_email)

        elif absent_days == 2:
            # Send warning email to employee
            employee_email = "ruksharparveen527@gmail.com"  # Update with employee's email
            subject = f"Reminder: Absence Warning - {employee_name}"
            message = (f"Hi {employee_name},\n This email is to inform you that according to our attendance records, "
                       f"you are absent from your duties for {absent_days} days. Please apply for leave. \n Thanks")

            send_email(subject, message, employee_email)

    return {"message": "Absence processing complete"}


if __name__ == "__main__":
    from waitress import serve

    serve(app, host="0.0.0.0", port=8000)
# fgrcaqaykcgucgkk

# fgrcaqaykcgucgkk

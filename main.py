from fastapi import FastAPI
from contextlib import asynccontextmanager
from apscheduler.schedulers.asyncio import AsyncIOScheduler
import imaplib
import email
import logging
import os
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PyPDF2 import PdfWriter, PdfReader

# app = FastAPI()
scheduler = AsyncIOScheduler()
logging.basicConfig(level=logging.INFO)


# Connect to Gmail
def connect_to_gmail(user, password):
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(user, password)
    mail.select('inbox')
    logging.info(f'Logged in as {user}')
    return mail


# Fetch today's unseen emails ids
def fetch_todays_unseen_emails(mail):
    today = datetime.now().strftime("%d-%b-%Y")  # Format: 10-Jul-2024
    search_criteria = f'(UNSEEN ON "{today}")'
    _, data = mail.search(None, search_criteria)
    email_ids = data[0].split()
    return email_ids


# Sanitize filename
def sanitize_filename(filename):
    return "".join(c for c in filename if c.isalnum() or c in (' ', '.', '_', '-')).rstrip()


# Fetch and merge invoice PDFs
def fetch_and_merge_invoice_pdfs(mail, email_id, output_folder):
    try:
        # Fetch the email by ID
        _, data = mail.fetch(email_id, '(RFC822)')

        # Check if data is valid
        if not data or not data[0] or not data[0][1]:
            logging.warning("No data found for email ID: %s", email_id)
            return

        # Parse the email content
        msg = email.message_from_bytes(data[0][1])

        # Extract subject, sender, and date
        subject = msg.get('subject', '')
        sender = msg.get('from', '')
        date = msg.get('date', '')

        logging.info("Processing email - Subject: %s, Sender: %s, Date: %s", subject, sender, date)

        # Initialize PDF writer
        pdf_writer = PdfWriter()

        # Function to append PDF to writer
        def append_pdf(pdf_bytes, writer):
            pdf_reader = PdfReader(BytesIO(pdf_bytes))
            for page_num in range(len(pdf_reader.pages)):
                writer.add_page(pdf_reader.pages[page_num])

        # Find and append PDF attachments
        attachment_found = False
        for part in msg.walk():
            if part.get_content_type() == 'application/pdf' and part.get_filename():
                filename = part.get_filename()
                if filename.lower().startswith('invoice'):
                    logging.info("Found invoice PDF: %s", filename)
                    pdf_bytes = part.get_payload(decode=True)
                    append_pdf(pdf_bytes, pdf_writer)
                    attachment_found = True

        if not attachment_found:
            logging.info("No invoice PDF attachments found in this email.")
            return

        # Create a PDF with email content
        email_content = f"Subject: {subject}\nFrom: {sender}\nDate: {date}\n\n"

        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == 'text/plain':
                    email_content += part.get_payload(decode=True).decode()
                    break
        else:
            email_content += msg.get_payload(decode=True).decode()

        email_pdf_buffer = BytesIO()
        c = canvas.Canvas(email_pdf_buffer, pagesize=letter)
        text_object = c.beginText(40, 750)
        for line in email_content.split('\n'):
            text_object.textLine(line)
        c.drawText(text_object)
        c.showPage()
        c.save()
        email_pdf_buffer.seek(0)

        # Append the email content PDF
        append_pdf(email_pdf_buffer.read(), pdf_writer)

        # Ensure the output directory exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            logging.info('Folder created: %s', output_folder)

        # Generate filename
        sanitized_subject = sanitize_filename(subject)
        sanitized_date = datetime.strptime(date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
        output_file_name = f"{sanitized_subject}_{sanitized_date}.pdf"
        output_file_path = os.path.join(output_folder, output_file_name)

        # Write out the final merged PDF
        with open(output_file_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

        logging.info("Merged PDF created successfully at %s", output_file_path)

        # Mark the email as seen
        mail.store(email_id, '+FLAGS', '\\Seen')
        logging.info("Email ID %s marked as seen", email_id)

    except Exception as e:
        logging.error("Failed to process email %s: %s", email_id, str(e))


# Task to run every day at 8 AM
def daily_task():
    email_address = 'xxx@gmail.com'
    email_password = 'xxx xxx xxx'
    output_folder = 'email_generate'

    # Calling email
    mail = connect_to_gmail(email_address, email_password)
    email_ids = fetch_todays_unseen_emails(mail)
    logging.info(email_ids)
    if not email_ids:
        logging.info("No unseen emails found for today.")
        return
    for email_id in email_ids:
        fetch_and_merge_invoice_pdfs(mail, email_id, output_folder)


# Schedule the task
scheduler.add_job(daily_task, 'interval', minutes=1)


@asynccontextmanager
async def lifespan(_: FastAPI):
    scheduler.start()
    logging.info("Scheduler started.")
    yield
    scheduler.shutdown()
    logging.info("Scheduler stopped.")


app = FastAPI(lifespan=lifespan)


@app.get("/alive")
async def read_root():
    return {"message": "Hello World"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)

import os
import smtplib
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from textwrap import wrap
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pandas as pd
from xlsx2csv import Xlsx2csv
import io


# Define page size and where to write on page
PAGE_WIDTH, PAGE_HEIGHT = letter
TEXT_WIDTH_LIMIT = 400  # Maximum width for text before wrapping
TEXT_START_Y = 550  # Initial Y position for text

# Function to center text and wrap if needed
def draw_centered_text(c, text, font_size, y_position, max_width):
    c.setFont("Times-Roman", font_size)

    # Wrap text if it exceeds max_width
    wrapped_text = wrap(text, width=int(max_width / (font_size * 0.45)))  # Adjust width for better wrapping
    for line in wrapped_text:
        text_width = c.stringWidth(line, "Times-Roman", font_size)
        x_position = (PAGE_WIDTH - text_width) / 2  # Center alignment
        c.drawString(x_position, y_position, line)
        y_position -= font_size + 5  # Move to next line with spacing

# Convert Excel to CSV and read with pandas
csv_buffer = io.StringIO()
Xlsx2csv("authors_data.xlsx", outputencoding="utf-8").convert(csv_buffer)
csv_buffer.seek(0)  # Reset buffer position
df = pd.read_csv(csv_buffer)

# Convert to a list of dictionaries
authors_data = df.to_dict(orient="records")

# Email Configuration
email_username = "youremail@gmail.com"
email_password = "yourpassword"
# Open connection for each email
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(email_username, email_password)

def get_unique_filename(folder, filename):
    """ Ensure filename is unique by appending a number if it already exists. """
    base, ext = os.path.splitext(filename)  # Separate filename and extension
    counter = 1
    unique_filename = filename

    # Keep incrementing the number until a unique filename is found
    while os.path.exists(os.path.join(folder, unique_filename)):
        unique_filename = f"{base}_{counter}{ext}"
        counter += 1

    return unique_filename

# select PDF Template and directory to save new
pdf_template = 'certificate_temp.pdf'
output_folder = "certificates"
os.makedirs(output_folder, exist_ok=True)  # Ensure the folder exists

# Process Each certificate and mail
for author_data in authors_data:
    author_name = author_data["name"].strip()
    filename = f"{author_name}_certificate.pdf"
    author_email = author_data['email']
    topic = author_data["topic"]

    # Get unique filename if needed
    unique_filename = get_unique_filename(output_folder, filename)
    output_pdf_path = os.path.join(output_folder, unique_filename)

    # Create a temporary PDF with centered text
    temp_pdf_path = "temp.pdf"
    c = canvas.Canvas(temp_pdf_path, pagesize=letter)

    # Draw centered name and topic on template
    draw_centered_text(c, author_name, 24, TEXT_START_Y, TEXT_WIDTH_LIMIT)
    draw_centered_text(c, topic, 14, TEXT_START_Y - 75, TEXT_WIDTH_LIMIT)

    c.save()

    # Read the template and merge text onto it
    template_reader = PdfReader(pdf_template)
    template_page = template_reader.pages[0]  # Get the first page of the template

    # Read the temporary text PDF
    text_reader = PdfReader(temp_pdf_path)
    text_page = text_reader.pages[0]

    # Merge the text onto the template
    template_page.merge_page(text_page)

    # Write the final certificate PDF
    writer = PdfWriter()
    writer.add_page(template_page)
    with open(output_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)

    # Remove the temporary text PDF
    os.remove(temp_pdf_path)

    # Attach the certificate to email
    msg = MIMEMultipart()
    msg['From'] = email_username
    msg['To'] = author_email
    msg['Subject'] = 'Your SEDT Conference Certificate'

    body = f"""
    Dear {author_name}, 

Please find your certificate attached.

Best regards,

    """
    msg.attach(MIMEText(body, 'plain'))

    # Attach the Certificate
    with open(output_pdf_path, 'rb') as attachment:
        part = MIMEBase('application', 'pdf')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{author_name}_certificate.pdf"')
        msg.attach(part)

    # Attach the Another file if necessary
    """attach2 = "Book_of_Proceedings_for_SEDT_4th_Biennial_Conference_2025.pdf"
    attachment2 = open(attach2, 'rb')
    part2 = MIMEBase('application', 'octet-stream')
    part2.set_payload(attachment2.read())
    encoders.encode_base64(part2)
    part2.add_header('Content-Disposition', 'attachment; filename=SEDT_Book_of_Proceedings.pdf')
    msg.attach(part2)
    attachment2.close()"""

    # Send Email
    server.sendmail(email_username, author_email, msg.as_string())

    print(f"Certificate sent to {author_email}")

    log_filename = "sent_log.txt"

    with open(log_filename, "a") as log_file:
        log_file.write(f"{author_name}'s certificate sent to {author_email}\n")

    print(f"Email log saved to {log_filename}")

# Close SMTP connection after all emails are sent
server.quit()
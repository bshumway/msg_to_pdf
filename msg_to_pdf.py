import os
import re
from extract_msg import Message
from fpdf import FPDF
from datetime import datetime

def sanitize_text(text):
    """
    Removes characters that are not supported by latin-1 encoding.
    """
    return text.encode('latin-1', 'ignore').decode('latin-1')

def format_email_date(date):
    """
    Formats the email date for the PDF title.
    """
    # If the date is already a datetime object, no conversion is needed.
    # Otherwise, parsing from string to datetime might be necessary.
    if isinstance(date, datetime):
        return date.strftime("%Y-%m-%d")
    return date

def sanitize_filename(text):
    """
    Sanitizes the filename by removing unsuitable characters, trimming extra spaces.
    The first letter of the resulting filename is not capitalized, but the rest of the filename is processed to title case.
    """
    # Remove unsuitable characters
    text = re.sub(r'[<>:"/\\|?*\'+$!]', '', text)
    # Remove extra spaces
    text = " ".join(text.split())
    # Process to title case without changing the first character
    parts = text.split(' ')
    parts = [parts[0].lower()] + [part.capitalize() for part in parts[1:]] if len(parts) > 1 else [parts[0].lower()]
    text = ' '.join(parts)
    return text

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, self.title, 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def convert_msg_to_pdf(msg_file_path, pdf_file_path):
    """
    Convert a single .msg file to a PDF, handling encoding issues and cleaning filenames.
    """
    # Load the .msg file
    msg = Message(msg_file_path)
    
    # Initialize PDF
    pdf = PDF()
    pdf.title = sanitize_text("Subject: " + msg.subject)
    pdf.add_page()
    
    # Set the font to Arial
    pdf.set_font('Arial', size=12)
    
    # Add sanitized body
    body = sanitize_text(msg.body) if msg.body else "No body content."
    pdf.multi_cell(0, 10, body)
    
    # Save the PDF
    pdf.output(pdf_file_path)

def convert_folder_msg_to_pdf(folder_path):
    """
    Convert all .msg files in a folder to PDFs, handling encoding issues and cleaning filenames.
    """
    for filename in os.listdir(folder_path):
        if filename.endswith(".msg"):
            msg_file_path = os.path.join(folder_path, filename)
            msg = Message(msg_file_path)
            email_date_formatted = format_email_date(msg.date)
            # Cleaning filename
            cleaned_filename = sanitize_filename(f"{email_date_formatted} - {msg.subject}.pdf")
            pdf_file_path = os.path.join(folder_path, cleaned_filename)
            convert_msg_to_pdf(msg_file_path, pdf_file_path)
            print(f"Converted {filename} to PDF as {cleaned_filename}.")

# Set the folder path containing your .msg files
folder_path = '/path/to/file'
convert_folder_msg_to_pdf(folder_path)
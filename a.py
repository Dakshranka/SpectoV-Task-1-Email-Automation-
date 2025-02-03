import os
import smtplib
from flask import Flask, request, jsonify
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageDraw, ImageFont
from apscheduler.schedulers.background import BackgroundScheduler
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow

# Flask App
app = Flask(__name__)

# Configuration
SMTP_SERVER = "smtp.gmail.com"
PORT = 587
EMAIL_LOGIN = "rankadaksh4@gmail.com"  # Replace with your email
EMAIL_PASSWORD = "cpna tpeu qajk pctk"  # Replace with your email password
TEMPLATE_PATH = "template.png"  # Path to your template image
OUTPUT_FOLDER = "output_images"
SPREADSHEET_ID = '1x6_Oo_IO_uRTc8JfzE_5FD9znekSFp2ZiQWv8oDM-EE'  # Replace with your Google Sheets ID
RANGE_NAME = 'Sheet1!A2:B'  # Assuming names are in column A and emails in column B

# Ensure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Google Sheets Authentication
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def authenticate_google_sheets():
    """Authenticate the Google Sheets API."""
    creds = None
    if os.path.exists('gen-lang-client-0033592652-00b17a4d57ff.json'):
        creds = Credentials.from_authorized_user_file('gen-lang-client-0033592652-00b17a4d57ff.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for future use
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def fetch_data_from_sheets():
    """Fetch name and email from Google Sheets."""
    creds = authenticate_google_sheets()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Get the data including the header row
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        raise Exception("No data found in the Google Sheets.")
    
    # Find the column indexes for "Name" and "Email Address"
    headers = values[0]  # First row contains headers
    name_index = headers.index("Name")
    email_index = headers.index("Email Adress")

    # Extract name and email from the data, skipping the header row
    intern_data = [(row[name_index], row[email_index]) for row in values[1:]]

    return intern_data

def create_welcome_image(template_path, output_path, intern_name):
    """Generates a personalized welcome image."""
    try:
        image = Image.open(template_path)
        draw = ImageDraw.Draw(image)
        
        # Ensure the font path is correct
        try:
            font = ImageFont.truetype("arial.ttf", 36)  # Update path if necessary
        except IOError:
            font = ImageFont.load_default()  # Fallback to default font
            
        # Adjust the position so the name is placed after "Dear: "
        text_position = (180, 150)  # Position after "Dear: " (adjust based on your template)
        
        # Add the intern's name (you may need to fine-tune the position)
        draw.text(text_position, intern_name, fill="black", font=font)
        
        image.save(output_path)
    except Exception as e:
        raise Exception(f"Error creating welcome image: {e}")

def send_email(recipient_email, name):
    """Sends an email with the given details."""
    subject = "Welcome to the SpectoV Team!"
    body = f"""\
    Dear {name},

    On behalf of the entire team, we would like to extend a warm welcome to you at SpectoV! We are thrilled to have you onboard as part of our talented team. Your skills and enthusiasm will certainly help us move forward as a company.

    Please take this opportunity to explore the team, get to know everyone, and feel free to reach out to us if you need any assistance. We are excited to work with you and are confident that you'll thrive here!

    We have attached a personalized welcome image for you to celebrate your first day. We're looking forward to achieving great things together.

    Best Regards,
    The SpectoV Team
    """

    msg = MIMEMultipart()
    msg['From'] = EMAIL_LOGIN
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Generate and attach the welcome image
    output_path = os.path.join(OUTPUT_FOLDER, f"{name}.png")
    try:
        create_welcome_image(TEMPLATE_PATH, output_path, name)
    except Exception as e:
        raise Exception(f"Error creating image for {name}: {e}")

    # Attach the image as a file
    try:
        with open(output_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(output_path)}')
        msg.attach(part)
    except Exception as e:
        raise Exception(f"Error attaching image: {e}")

    # Send the email
    try:
        with smtplib.SMTP(SMTP_SERVER, PORT) as server:
            server.starttls()
            server.login(EMAIL_LOGIN, EMAIL_PASSWORD)
            server.sendmail(EMAIL_LOGIN, recipient_email, msg.as_string())
    except Exception as e:
        raise Exception(f"Error sending email to {recipient_email}: {e}")

def batch_send_emails():
    """Fetches data from Google Sheets and sends batch emails."""
    try:
        intern_data = fetch_data_from_sheets()  # Fetch all data from Google Sheets
        
        for name, email in intern_data:
            try:
                send_email(email, name)
                print(f"Sent email to {name} ({email})")
            except Exception as e:
                print(f"Error sending email to {name} ({email}): {e}")
    except Exception as e:
        print(f"Error in batch process: {e}")

# Set up APScheduler to run the batch job periodically (e.g., every day at 9 AM)
scheduler = BackgroundScheduler()
scheduler.add_job(batch_send_emails, 'cron', hour=9, minute=0)  # Run every day at 9:00 AM
scheduler.start()

@app.route('/trigger-email', methods=['POST'])
def trigger_email():
    """Handles the incoming trigger and sends the email."""
    try:
        data = request.json

        # Log the received data for debugging purposes
        print(f"Received data: {data}")

        # Ensure the JSON contains both 'Name' and 'Email Adress'
        name = data.get('Name')
        email = data.get('Email Adress')

        if not name or not email:
            return jsonify({"error": "'Name' and 'Email Adress' are required fields."}), 400

        # Send email with the provided name and email
        try:
            send_email(email, name)
        except Exception as e:
            return jsonify({"error": f"Error sending email to {name}: {str(e)}"}), 500

        return jsonify({"message": f"Email sent to {name} ({email})"}), 200

    except Exception as e:
        # Log the error for debugging
        print(f"Unexpected error: {str(e)}")
        return jsonify({"error": f"Unexpected error: {str(e)}"}), 500

# Run the Flask application
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)

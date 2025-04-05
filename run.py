import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.message import EmailMessage
import os

# Load Excel file
try:
    df = pd.read_excel('data.xlsx')  # Make sure this is in the same directory or provide path
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Certificate template path
template_path = 'certificate.jpeg'
font_path = 'Chateau des Oliviers.ttf'  # Change this to any font path available on your system
font_size = 420

# Load font
try:
    font = ImageFont.truetype(font_path, font_size)
except IOError:
    print(f"Font file {font_path} not found. Using default font.")
    font = ImageFont.load_default()

# Email credentials
SENDER_EMAIL = 'info@techniknest.tech'
SENDER_PASSWORD = 'nopassword202#'

# Create folder to save generated certificates
os.makedirs("generated_certificates", exist_ok=True)


def generate_certificate(name):
    try:
        image = Image.open(template_path)
        draw = ImageDraw.Draw(image)

        # Define the max and min font sizes
        max_font_size = 420
        min_font_size = 50
        reference_width = 3000  # Set your custom reference width here (in pixels)

        # Start with the max font size and adjust as necessary
        font_size = max_font_size
        font = ImageFont.truetype(font_path, font_size)

        # Calculate the bounding box of the text to get its width and height
        bbox = draw.textbbox((0, 0), name, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        # While the text is too wide compared to the reference width, reduce the font size
        while text_width > reference_width:  # Adjust based on your reference width
            font_size -= 5
            font = ImageFont.truetype(font_path, font_size)

            # Recalculate the text dimensions with the new font size
            bbox = draw.textbbox((0, 0), name, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            # If the font size is too small, break the loop
            if font_size <= min_font_size:
                print(f"Name '{name}' is too long, using the smallest font size.")
                break

        # Calculate the position to center the text horizontally, using reference width
        name_position = ((image.width - text_width) // 2, 1700)  # Adjust vertical position if needed

        # Draw the text with the chosen font size and centered position
        draw.text(name_position, name, font=font, fill="black")

        # Save the certificate
        filename = f"generated_certificates/{name.replace(' ', '_')}.png"
        image.save(filename)
        return filename
    except Exception as e:
        print(f"Error generating certificate for {name}: {e}")
        return None


def send_email(receiver_email, subject, body, attachment_path):
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = SENDER_EMAIL
        msg['To'] = receiver_email
        msg.set_content(body)

        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)

        msg.add_attachment(file_data, maintype='image', subtype='png', filename=file_name)

        with smtplib.SMTP_SSL('smtp.titan.email', 465) as smtp:
            smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
            smtp.send_message(msg)
            print(f"Sent to {receiver_email}")
    except Exception as e:
        print(f"Error sending email to {receiver_email}: {e}")


def format_name(name):
    # Capitalize the first letter of the first word and after spaces
    words = name.split()
    
    # Capitalize first letter of each word and make the rest lowercase
    formatted_name = ' '.join([word.capitalize() for word in words])

    return formatted_name

# Loop through Excel rows
for index, row in df.iterrows():
    name = row['Name']
    email = row['Email']

    # Format the name
    formatted_name = format_name(name)
    
    cert_path = generate_certificate(formatted_name)
    if cert_path:
        send_email(
            receiver_email=email,
            subject='Your Certificate of Participation',
            body=f'Dear Participant,\n\nThank you for attending the Artificial Intelligence: Algorithms & Opportunities Seminar in Pak Austria Fachhochschule Institute of Applied Sciences And Technology. Your participation made it a great success. Please find your certificate of participation attached. We’d love to see your achievement—upload it to your social media and share the link with us. Don’t forget to tag us!\n\n📌 Connect with us:\n\n🔹 Facebook: https://www.facebook.com/techniknest/\n🔹 Instagram: https://www.instagram.com/techniknest/ \n🔹 LinkedIn: https://www.linkedin.com/company/techniknest/ \n🔹 WhatsApp: https://whatsapp.com/channel/0029VaEGUWDBlHpTjOtHxS2L \n\nLooking forward to staying connected!\n\nBest Regards,\n Muhammad Ali \nFounder & CEO, Technik Nest',
            attachment_path=cert_path
        )

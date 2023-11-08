#"hier --. einfügen" deklariert offene pfade

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import ssl
import os

# Read the CSV file
df = pd.read_csv("/stads_intra/datathon/participantsfinal.csv", sep=";", encoding='utf-8') #hier teilnehmerliste einfügen
pdf_folder = "/stads_intra/datathon/urkunden" #hier pfad zu Urkunden einfügen

# Set up your email server connection
smtp_server = 'smtp-mail.outlook.com'
port = 587
sender_email = '' #hier stadsmail einfügen
password = '' #hier pw einfügen

context = ssl.create_default_context()

# Loop through the rows of the CSV
for index, row in df.iterrows():
    fullname = row['fullname']
    surname = row["surname"]
    email = row['email']
    subject = 'Thank you for your participation! | STADS Datathon Fall 2023'

    # Check if the PDF file exists
    pdf_filename = f'Urkunde_{fullname}.pdf'  # Assuming the PDFs are in a subfolder named 'urkunden'
    if not os.path.exists(os.path.join(pdf_folder, pdf_filename)):
        print(f'Error: PDF file {pdf_filename} not found for {fullname}. Skipping email.')
        continue


    body = f'''
    <html>
        <body>
            <p>Dear {surname},</p>

            <p>The Datathon Fall 2023 is over and we would like to thank you again on behalf of the organizing team for your participation. Your team has made this event an absolute success with enthusiasm, creativity and hard work!</p>

            <p>We would like to congratulate each of you personally, because your project was absolutely impressive. The quality and variety of ideas presented during the Datathon left a lasting impression on our corporate partners. From innovative solutions to finished prototypes, each project was unique and was a testament to your expertise and dedication. For this, you and your team will now receive your certificate of participation attached.</p>

            <p>Since your opinion is invaluable to us, we would love to get your feedback on the Datathon and the event overall. Please take a moment to share your thoughts, suggestions and ideas with us anonymously in this <a href="https://forms.office.com/Pages/ResponsePage.aspx?id=RqmXPyiPoU6G399hkc8qF30rDZmA5vlDka1NRPVNxXZUNjUwVUhTNEI4VzJBR01XUktQRFlHVFhTRC4u">survey</a>. Your feedback will help us make our event even better in the future and meet the needs of participants.</p>

            <p>Finally, we would like to invite you to review the Datathon. We have captured the best moments and group pictures in this <a href="https://stadsinitiative-my.sharepoint.com/:f:/g/personal/elise_wolf_stads_de/EvqjxuNHg3VAkjmHXtcGVtUB8hkMxLI-FwB3rzJZPCijqA?e=hQ9cOL">folder</a>. The aftermovie will follow in a few weeks. Have fun watching! You are invited to share it on all social media channels.</p>

            <p>Last, you can fill out this <a href="https://forms.office.com/Pages/ResponsePage.aspx?id=RqmXPyiPoU6G399hkc8qF30rDZmA5vlDka1NRPVNxXZUMzRLQlNVODFCRExSUFcwQzJFTkhHVTdZVy4u">form</a> to be notified of upcoming events.</p>

            <p>Teaser: the next Datathon is expected to take place on May 02 and 03, 2024, so mark your calendars!</p>

            <p>To make sure you don't miss anything, feel free to follow us on <a href="https://www.instagram.com/stads_mannheim/?hl=de">Instagram</a> and <a href="https://www.linkedin.com/company/stadsmannheim/">LinkedIn</a>!</p>

            <p>We hope to see you again at the next Datathon at FSS 2024. Until then!</p>

            <p>Your Datathon organizing team</p>
            <p>by <a href="https://stads.uni-mannheim.de/">STADS</a></p>

            <!-- Footer -->
            <p>&mdash;&mdash;</p>
            <p><strong style="color: #203864;">Elise Wolf</strong></p>
            <p style="color: #203864;">Students&rsquo; Association for Data Analytics&nbsp;&amp; Statistics (STADS) Mannheim e.V.</p>
            <p style="color: #203864;">Schloss, 68131 Mannheim</p>
            <p><img src="https://stads.uni-mannheim.de/wp-content/uploads/2021/04/STADS_LOGO_2021-dark-blue.jpg" width="250" height="53" alt="" /></p>

            <p><a href="https://stads.uni-mannheim.de/" style="color: #203864;">stads.de</a>&nbsp;||&nbsp;<a href="https://www.linkedin.com/company/stadsmannheim" style="color: #203864;">linked.in/company/stads</a></p>
            <p style="color: #203864;">Vereinsregister Mannheim VR701996&nbsp;|&nbsp;Vertretungsberechtigter Vorstand: Luca Marohn, Elise Wolf, Simon Schumacher</p>

        </body>
    </html>'''

    # Create mail
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = subject

    # Add the text
    msg.attach(MIMEText(body, 'html'))

    # Add the PDF
    
    # Add the PDF attachment
    with open(os.path.join(pdf_folder, pdf_filename), 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype="pdf")
        attachment.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
        msg.attach(attachment)

    # Send the email
    try:
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls(context=context)
            server.login(sender_email, password)
            server.sendmail(sender_email, email, msg.as_string())

        print(f'E-Mail an {email} gesendet.')
    except Exception as e:
        print(f'Error sending email to {email}: {e}')
print('Alle E-Mails versendet.')

import streamlit as st
import pandas as pd
import os
import shutil
import zipfile
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import reportlab.rl_config
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



#Funtion for sending emails
def send_emails(participants, subject, body, sender_email, sender_password):
    try:
        # Set up SMTP server
        smtp_server = "smtp.gmail.com"
        port = 587
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()

        # Log in to email account
        server.login(sender_email, sender_password)

        # Iterate over each participant's email and send email
        for index, participant in participants.iterrows():
            receiver_email = participant['Email']  # Assuming 'Email' is the column name for email addresses
            message = MIMEMultipart()
            message['From'] = sender_email
            message['To'] = receiver_email
            message['Subject'] = subject

            # Add email body
            message.attach(MIMEText(body, 'plain'))

            # Send email
            server.sendmail(sender_email, receiver_email, message.as_string())

        # Close the SMTP server connection
        server.quit()
        
        return True  # Email sent successfully

    except Exception as e:
        print("Error sending email:", e)
        return False  # Email sending failed


# Function to generate a course certificate
def generate_course_certificate(name, course, id, score, month, year):
    # Create variables to be added to the certificate template
    data = "For completing " + course
    data2 = "on " + str(month) + " " + str(year) + ", with a passing score of " + str(score) +"%"+ "."
    certificate_id = "Certificate ID : " + str(id)

    # Set canvas conditions to add text to the template
    packet = io.BytesIO()
    width, height = letter
    c = canvas.Canvas(packet, pagesize=(width*2, height*2))

    # Get text font types
    reportlab.rl_config.warnOnMissingFontGlyphs = 0
    pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
    pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
    pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))

    # Set font features for each variable
    c.setFillColorRGB(0, 0, 0)
    c.setFont('VeraBd', 25)
    c.drawString(170, 260, name)

    c.setFillColorRGB(1/255, 37/255, 84/255)
    c.setFont('Helvetica-Bold', 16)
    c.drawString(172, 220, data)

    c.setFillColorRGB(1/255, 37/255, 84/255)
    c.setFont('Helvetica-Bold', 16)
    c.drawString(172, 195, data2)

    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont('Vera', 11)
    c.drawString(145, 25.5, str(certificate_id))

    # Save all Canvas settings
    c.save()

    # Get PDF template
    existing_pdf = PdfReader(open("course_completion_certificate.pdf", "rb"))
    page = existing_pdf.pages[0]

    # Add generated content to the template
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page.merge_page(new_pdf.pages[0])

    # Return the merged PDF page
    return page

# Function to generate an internship certificate
def generate_internship_certificate(name, course, id, month, year, duration):
    # Create variables to be added to the certificate template
    data1 = f"For outstanding completion of the {duration} internship program at"
    data2 = f"Clover Technologies in {course} in {month} {year}."
    certificate_id = f"Certificate ID: {id}"

    # Set canvas conditions to add text to the template
    packet = io.BytesIO()
    width, height = letter
    c = canvas.Canvas(packet, pagesize=(width*2, height*2))

    # Get text font types
    reportlab.rl_config.warnOnMissingFontGlyphs = 0
    pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
    pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
    pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))

    # Set font features for each variable
    c.setFillColorRGB(0, 0, 0)
    c.setFont('VeraBd', 32)
    c.drawString(170, 260, name)

    c.setFillColorRGB(1/255, 37/255, 84/255)
    c.setFont('Helvetica-Bold', 16)
    c.drawString(172, 220, data1)

    c.setFillColorRGB(1/255, 37/255, 84/255)
    c.setFont('Helvetica-Bold', 16)
    c.drawString(172, 195, data2)

    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont('Vera', 14)
    c.drawString(145, 25.5, certificate_id)

    # Save all Canvas settings
    c.save()

    # Get PDF template
    existing_pdf = PdfReader(open("internship_certificate.pdf", "rb"))
    page = existing_pdf.pages[0]

    # Add generated content to the template
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page.merge_page(new_pdf.pages[0])

    return page

# Streamlit UI
st.title("Certificate Generator")
st.sidebar.title("Menu")

# Static list of pages
pages = ["Home", "Internship Certificate Generator", "Course Certificate Generator", "Template Editor"]

# Page selection
selected_page = st.sidebar.radio("Select Page", pages)

# Clear uploaded file when switching pages
uploaded_file = None


if selected_page == "Home":
    st.title("Welcome")
    st.write("This  certificate generator is used to genrate bulk certificates from Excel.")

elif selected_page == "Internship Certificate Generator":
    st.title("Internship Certificate Generator")

    # Add file uploader for participants Excel file
    uploaded_file = st.file_uploader("Upload participants Excel file", type="xlsx")

    if uploaded_file:
        try:
            participants = pd.read_excel(uploaded_file)

            # Check if required columns are present
            required_columns = ['Name', 'Course', 'Id', 'Month', 'Year', 'Duration']
            if not set(required_columns).issubset(participants.columns):
                st.warning("The uploaded Excel file does not contain all required columns.")
            else:
                generated_certificates = []

                if st.button("Generate Internship Certificates"):
                    # Certificate generation
                    output_path = os.path.join(os.getcwd(), "certificates")
                    os.makedirs(output_path, exist_ok=True)

                    zip_file_path = os.path.join(os.getcwd(), "certificates.zip")
                    with zipfile.ZipFile(zip_file_path, 'w') as zipf:
                        for i in range(len(participants)):
                            student = participants.loc[i, 'Name']
                            course = participants.loc[i, 'Course']
                            id = participants.loc[i, 'Id']
                            month = participants.loc[i, 'Month']
                            year = participants.loc[i, 'Year']
                            duration = participants.loc[i, 'Duration']
                            certificate_page = generate_internship_certificate(student, course, id, month, year, duration)

                            generated_certificates.append(certificate_page)

                            file_name = student.replace(" ", "_")
                            certificate_path = os.path.join(output_path, f"{file_name}_certificate.pdf")

                            output = PdfWriter()
                            output.add_page(certificate_page)
                            with open(certificate_path, "wb") as f:
                                output.write(f)

                            zipf.write(certificate_path, arcname=f"{file_name}_certificate.pdf")

                    # Display buttons for download and send email in a horizontal layout
                    with open(zip_file_path, "rb") as f:
                        download_button = st.download_button(label="Download All Certificates", data=f.read(), file_name="certificates.zip", mime="application/zip")

                    # Add spacer to create space between buttons
                    st.write("")
                    
                    if st.button("Send Email"):
                        if send_emails(participants, "Your Subject", "Your Email Body", "your_email@gmail.com", "your_password"):
                              st.success("Emails sent successfully!")
                        else:
                              st.error("Failed to send emails. Please check your credentials and try again.")

                              
                    # Clean up generated certificate files and zip
                    shutil.rmtree(output_path)
                    os.remove(zip_file_path)

        except Exception as e:
            st.warning("An error occurred while processing the file. Please make sure the file format is correct and try again.")


elif selected_page == "Course Certificate Generator":
    st.title("Course Certificate Generator")

    # Add file uploader for participants Excel file
    uploaded_file = st.file_uploader("Upload participants Excel file", type="xlsx")

    if uploaded_file:
        try:
            participants = pd.read_excel(uploaded_file)

            # Check if required columns are present
            required_columns = ['Name', 'Course', 'Id', 'Score', 'Month', 'Year']
            if not set(required_columns).issubset(participants.columns):
                st.warning("The uploaded Excel file does not contain all required columns.")
            else:
                generated_certificates = []

                if st.button("Generate Course Certificates"):
                    # Certificate generation
                    output_path = os.path.join(os.getcwd(), "certificates")
                    os.makedirs(output_path, exist_ok=True)

                    zip_file_path = os.path.join(os.getcwd(), "certificates.zip")
                    with zipfile.ZipFile(zip_file_path, 'w') as zipf:
                        for i in range(len(participants)):
                            student = participants.loc[i, 'Name']
                            course = participants.loc[i, 'Course']
                            id = participants.loc[i, 'Id']
                            score = participants.loc[i, 'Score']
                            month = participants.loc[i, 'Month']
                            year = participants.loc[i, 'Year']
                            certificate_page = generate_course_certificate(student, course, id, score, month, year)

                            generated_certificates.append(certificate_page)

                            file_name = student.replace(" ", "_")
                            certificate_path = os.path.join(output_path, f"{file_name}_certificate.pdf")

                            output = PdfWriter()
                            output.add_page(certificate_page)
                            with open(certificate_path, "wb") as f:
                                output.write(f)

                            zipf.write(certificate_path, arcname=f"{file_name}_certificate.pdf")

                    with open(zip_file_path, "rb") as f:
                        st.download_button(label="Download All Certificates", data=f.read(), file_name="certificates.zip", mime="application/zip")

                    # Clean up generated certificate files and zip
                    shutil.rmtree(output_path)
                    os.remove(zip_file_path)
        except Exception as e:
            st.warning("An error occurred while processing the file. Please make sure the file format is correct and try again.")

elif selected_page == "Template Editor":
    st.title("Template Editor")
    st.write("This is the Template Editor page.")


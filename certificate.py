from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
import reportlab.rl_config
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import pandas as pd
import streamlit as st
import os
import shutil
import zipfile

def generate_course_certificate(student, course, id, score):
    packet = io.BytesIO()
    width, height = 960, 720
    c = canvas.Canvas(packet, pagesize=(width, height))

    reportlab.rl_config.warnOnMissingFontGlyphs = 0
    pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
    pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
    pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))

    # Student
    c.setFillColorRGB(0, 0, 0)
    c.setFont('VeraBd', 32)
    c.drawString(170, 255, student)

    # Course
    c.setFillColorRGB(1/255, 37/255, 84/255)
    c.setFont('Helvetica-Bold', 15)
    c.drawString(370, 214, course)

    # ID
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont('Vera', 14)
    c.drawString(240, 25.5, str(id))

    # Score
    c.setFillColorRGB(1/255, 37/255, 84/255)
    c.setFont('Helvetica-Bold', 15)
    c.drawString(498, 192.5, str(score))

    # Save changes
    c.save()

    existing_pdf = PdfReader(open("course_completion_certificate.pdf", "rb"))
    page = existing_pdf.pages[0]
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page.merge_page(new_pdf.pages[0])

    return page


def generate_internship_certificate(student, course, id, month, year):
    packet = io.BytesIO()
    width, height = 960, 720
    c = canvas.Canvas(packet, pagesize=(width, height))

    reportlab.rl_config.warnOnMissingFontGlyphs = 0
    pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
    pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
    pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))

    line = course + " in " + month + " " + str(year) + "."
    num = " Certificate ID : " + id

    # Student
    c.setFillColorRGB(0, 0, 0)
    c.setFont('VeraBd', 32)
    c.drawString(170, 255, student)

    # Course
    c.setFillColorRGB(1/255, 37/255, 84/255)
    c.setFont('Helvetica-Bold', 15)
    c.drawString(348.5, 192.5, line)

    # ID
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont('Vera', 12)
    c.drawString(135, 25.3, num)

    # Save changes
    c.save()

    existing_pdf = PdfReader(open("Internship_certificate.pdf", "rb"))
    page = existing_pdf.pages[0]
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page.merge_page(new_pdf.pages[0])

    return page

# Streamlit UI
st.title("Certificate Generator")

# Add file uploader for participants Excel file
uploaded_file = st.file_uploader("Upload participants Excel file", type="xlsx")
template_options = ["Course Completion Certificate", "Internship Certificate"]  # Add more template options here

if uploaded_file:
    participants = pd.read_excel(uploaded_file)
    generated_certificates = []

    # Add template selection dropdown
    selected_template = st.selectbox("Select Certificate Template", template_options)

    if st.button("Generate and Download All Certificates"):
        output_path = os.path.join(os.getcwd(), "certificates")
        os.makedirs(output_path, exist_ok=True)

        zip_file_path = os.path.join(os.getcwd(), "certificates.zip")
        with zipfile.ZipFile(zip_file_path, 'w') as zipf:
            for i in range(len(participants)):
                student = participants.loc[i, 'Name']  # Corrected column name
                course = participants.loc[i, 'Course']  # Corrected column name
                id = participants.loc[i, 'Id']
                if selected_template == "Course Completion Certificate":
                    score = participants.loc[i, 'Score']  # Corrected column name
                    certificate_page = generate_course_certificate(student, course, id, score)
                elif selected_template == "Internship Certificate":
                    month = participants.loc[i, 'Month']  # Corrected column name
                    year = participants.loc[i, 'Year']  # Corrected column name
                    certificate_page = generate_internship_certificate(student, course, id, month, year)
                else:
                    # Default to Course Completion Certificate
                    score = participants.loc[i, 'Score']  # Corrected column name
                    certificate_page = generate_course_certificate(student, course, id, score)

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

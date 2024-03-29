from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
import reportlab.rl_config
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import pandas as pd
import streamlit as st
import os  # Import the os module

def generate_certificate(student, course, id, output_path):
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
    c.setFillColorRGB(0, 0, 0)       
    c.setFont('Helvetica', 13)         
    c.drawString(135, 60, id)  

    # Save changes
    c.save()

    existing_pdf = PdfReader(open("certificate_template.pdf", "rb")) 
    page = existing_pdf.pages[0]             
    packet.seek(0)                           
    new_pdf = PdfReader(packet)               
    page.merge_page(new_pdf.pages[0])       

    file_name = student.replace(" ","_")
    certificate = os.path.join(output_path, f"{file_name}_certificate.pdf")  # Specify the output path
    outputStream = open(certificate, "wb")
    output = PdfWriter()
    output.add_page(page)
    output.write(outputStream)
    outputStream.close()

    return certificate

# Streamlit UI
st.title("Certificate Generator")

# Add file uploader for participants Excel file
uploaded_file = st.file_uploader("Upload participants Excel file", type="xlsx")

# Add text input for output path
output_path = st.text_input("Enter output path")

if uploaded_file and output_path:
    participants = pd.read_excel(uploaded_file)

    for i in range(len(participants)):
        student = participants.loc[i, 'Student']
        course = participants.loc[i, 'Course']
        id = participants.loc[i, 'Id']

        st.write(f"Generating certificate for: {student}")
        certificate = generate_certificate(student, course, id, output_path)
        st.write(f"Certificate generated: {certificate}")

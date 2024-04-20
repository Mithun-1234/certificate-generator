import streamlit as st
import pandas as pd
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io

# Function to generate certificates
def generate_certificates(participants):
    for i in range(len(participants)):
        name = participants.loc[i, 'Name']
        course = participants.loc[i, 'Course']
        id = participants.loc[i, 'Id']
        month = participants.loc[i, 'Month']
        year = participants.loc[i, 'Year']
        duration = participants.loc[i, 'Duration']

        data1 = "For outstanding completion of the " + str(duration) + " internship program at"
        data2 = "Clover Technologies in " + course + " in " + str(month) + " " + str(year) + "."
        certificate_id = "Certificate Id : " + str(id)

        packet = io.BytesIO()
        width, height = letter
        c = canvas.Canvas(packet, pagesize=(width * 2, height * 2))

        pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
        pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
        pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))

        c.setFillColorRGB(0, 0, 0)
        c.setFont('VeraBd', 32)
        c.drawString(170, 260, name)

        c.setFillColorRGB(1 / 255, 37 / 255, 84 / 255)
        c.setFont('Helvetica-Bold', 16)
        c.drawString(172, 220, data1)

        c.setFillColorRGB(1 / 255, 37 / 255, 84 / 255)
        c.setFont('Helvetica-Bold', 16)
        c.drawString(172, 195, data2)

        c.setFillColorRGB(0.2, 0.2, 0.2)
        c.setFont('Vera', 14)
        c.drawString(145, 25.5, str(certificate_id))

        c.save()

        packet.seek(0)
        new_pdf = PdfReader(packet)
        page = new_pdf.pages[0]

        certificado = f"{id}_certificate.pdf"
        outputStream = open(certificado, "wb")
        output = PdfWriter()
        output.add_page(page)
        output.write(outputStream)
        outputStream.close()

# Streamlit UI
def certificate_generator():
    st.title("Certificate Generator")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if uploaded_file is not None:
        participants = pd.read_excel(uploaded_file)
        st.write("Uploaded Data:")
        st.write(participants)

        if st.button("Generate Certificates"):
            generate_certificates(participants)
            st.success("Certificates generated successfully!")

# Main function to run the Streamlit app
def main():
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Go to", ["Home", "Certificate Generator", "Send Email", "Template Editor"])

    if page == "Home":
        st.title("Home Page")
        # Add content for home page as needed
    elif page == "Certificate Generator":
        certificate_generator()
    elif page == "Send Email":
        st.title("Send Email Page")
        # Add content for send email page as needed
    elif page == "Template Editor":
        st.title("Template Editor Page")
        # Add content for template editor page as needed

if __name__ == "__main__":
    main()

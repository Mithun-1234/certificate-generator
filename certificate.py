import streamlit as st

# Define page functions
def home():
    st.title("Home Page")
    st.write("Welcome to the Certificate Generator App!")
    # Add more content for the home page as needed

def certificate_generator():
    st.title("Certificate Generator")
    # Add content for certificate generator page

def send_email():
    st.title("Send Email")
    # Add content for sending email page

def template_editor():
    st.title("Template Editor")
    # Add content for template editor page

# Sidebar navigation
page = st.sidebar.selectbox("Navigate", ["Home", "Certificate Generator", "Send Email", "Template Editor"])

# Render selected page
if page == "Home":
    home()
elif page == "Certificate Generator":
    certificate_generator()
elif page == "Send Email":
    send_email()
elif page == "Template Editor":
    template_editor()

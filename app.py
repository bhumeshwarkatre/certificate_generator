# ⬇️ All necessary imports
import os
import re
import csv
import qrcode
import random
import string
import tempfile
import base64
import streamlit as st
from datetime import datetime
from smtplib import SMTP
from docx import Document  # ✅ Required for inserting QR before rendering
from docxtpl import DocxTemplate
from docx.shared import Inches
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
import pandas as pd

# ✅ Aspose Cloud Imports
import asposewordscloud
from asposewordscloud.apis.words_api import WordsApi
from asposewordscloud.models.requests import UploadFileRequest, SaveAsRequest, DownloadFileRequest
from asposewordscloud.models import PdfSaveOptionsData

# --- Streamlit Config ---
st.set_page_config("Completion Certificate Generator", layout="wide")

# --- Secrets ---
EMAIL = st.secrets["email"]["user"]
PASSWORD = st.secrets["email"]["password"]
ADMIN_KEY = st.secrets["admin"]["key"]
ASPOSE_ID = st.secrets["aspose"]["app_sid"]
ASPOSE_SECRET = st.secrets["aspose"]["app_key"]

# --- File Paths ---
CSV_FILE = "intern_data.csv"
TEMPLATE_FILE = os.path.join(tempfile.gettempdir(), "completion_template.docx")
LOGO = "logo.png"

# ✅ Load Word Template from base64
if not os.path.exists(TEMPLATE_FILE):
    encoded_template = st.secrets["template_base64"]["template_base64"]
    with open(TEMPLATE_FILE, "wb") as f:
        f.write(base64.b64decode(encoded_template))

# ✅ Aspose Setup
words_api = WordsApi(ASPOSE_ID, ASPOSE_SECRET)

def convert_to_pdf_asp(word_path, output_path):
    cloud_doc_name = os.path.basename(word_path)
    cloud_pdf_name = cloud_doc_name.replace(".docx", ".pdf")

    with open(word_path, "rb") as f:
        words_api.upload_file(UploadFileRequest(f, cloud_doc_name))

    save_opts = PdfSaveOptionsData(file_name=cloud_pdf_name)
    save_as_request = SaveAsRequest(name=cloud_doc_name, save_options_data=save_opts)
    words_api.save_as(save_as_request)

    pdf_stream = words_api.download_file(DownloadFileRequest(cloud_pdf_name))
    with open(output_path, "wb") as f:
        f.write(pdf_stream)

# --- Utility Functions ---
def format_date(date_obj):
    return date_obj.strftime("%d %B %Y")

def generate_certificate_key():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=9))

def generate_qr(data):
    try:
        qr = qrcode.QRCode(box_size=10, border=0)
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        filename = ''.join(random.choices(string.ascii_lowercase + string.digits, k=8)) + "_qr.png"
        path = os.path.join(tempfile.gettempdir(), filename)
        img.save(path)
        return path
    except Exception as e:
        st.error(f"❌ QR code generation failed: {e}")
        return ""

def insert_qr_into_table(doc_path, qr_path):
    try:
        doc = Document(doc_path)
        if not doc.tables or len(doc.tables[0].rows) < 1 or len(doc.tables[0].rows[0].cells) < 1:
            raise Exception("Template must have at least a 1x1 table to insert QR code.")

        cell = doc.tables[0].rows[0].cells[0]
        para = cell.paragraphs[0]
        para.clear()  # Clear any placeholder text
        para.add_run().add_picture(qr_path, width=Inches(1.4))

        modified_path = os.path.join(tempfile.gettempdir(), "template_with_qr.docx")
        doc.save(modified_path)
        return modified_path
    except Exception as e:
        st.warning(f"⚠️ QR code insert failed: {e}")
        return doc_path  # Fallback to original template

def send_email(receiver, pdf_path, data):
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = receiver
    msg['Subject'] = f"🎉 Completion Certificate - {data['name']}"

    html = f"""
    <html><body>
        <p>Dear <strong>{data['name']}</strong>,</p>
        <p>Congratulations on completing your <strong>{data['month']} month</strong> internship at <strong>SkyHighes Technology</strong>!</p>
        <p><b>Details:</b></p>
        <ul>
            <li><strong>Domain:</strong> {data['domain']}</li>
            <li><strong>Duration:</strong> {data['start_date']} to {data['end_date']}</li>
            <li><strong>Grade:</strong> {data['grade']}</li>
            <li><strong>Certificate ID:</strong> {data['c_id']}</li>
        </ul>
        <p>Your certificate is attached as a PDF.</p>
        <p>All the best for your future!</p>
        <p><strong>SkyHighes Technology Team</strong></p>
    </body></html>
    """
    msg.attach(MIMEText(html, 'html'))

    with open(pdf_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(pdf_path)}")
        msg.attach(part)

    with SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)

def save_to_csv(data, status="Sent"):
    exists = os.path.exists(CSV_FILE)
    with open(CSV_FILE, mode='a', newline='') as f:
        writer = csv.writer(f)
        if not exists:
            writer.writerow(["Name", "Domain", "Months", "Start Date", "End Date", "Grade", "Certificate ID", "Email", "send_mail"])
        writer.writerow([data['name'], data['domain'], data['month'], data['start_date'], data['end_date'], data['grade'], data['c_id'], data['email'], status])

# --- UI ---
st.markdown("""
<style>
.title-text {
    font-size: 2rem;
    font-weight: 700;
}
.stButton>button {
    background-color: #1E88E5;
    color: white;
    padding: 0.5rem 1.5rem;
    border-radius: 8px;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

with st.container():
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        if os.path.exists(LOGO):
            st.image(LOGO, width=80)
    with col_title:
        st.markdown('<div class="title-text">SkyHighes Technologies Completion Certificate Portal</div>', unsafe_allow_html=True)

st.divider()

# --- Form ---
with st.form("certificate_form"):
    st.subheader("🎓 Generate Completion Certificate")

    col1, col2, col3 = st.columns(3)
    with col1:
        name = st.text_input("Intern Name")
    with col2:
        domain = st.text_input("Domain")
    with col3:
        email = st.text_input("Recipient Email")

    col4, col5, col6 = st.columns(3)
    with col4:
        month = st.number_input("Internship Duration (Months)", min_value=1, max_value=12)
    with col5:
        start_date = st.date_input("Start Date", value=datetime.today())
    with col6:
        end_date = st.date_input("End Date", value=datetime.today())

    grade = st.selectbox("Grade", ["A+", "A", "B+", "B", "C"])
    submit = st.form_submit_button("🎯 Generate & Send Certificate")

# --- Generate Logic ---
if submit:
    if not all([name, domain, email]):
        st.error("❌ Please fill all fields.")
    elif end_date < start_date:
        st.warning("⚠️ End date cannot be before start date.")
    elif not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        st.warning("⚠️ Invalid email.")
    else:
        cert_id = generate_certificate_key()
        data = {
            "name": name.strip(),
            "domain": domain.strip(),
            "month": month,
            "start_date": format_date(start_date),
            "end_date": format_date(end_date),
            "grade": grade,
            "c_id": cert_id,
            "email": email.strip()
        }

        save_to_csv(data)
        qr_path = generate_qr(", ".join(str(v) for v in data.values()))
        modified_template = insert_qr_into_table(TEMPLATE_FILE, qr_path)

        doc = DocxTemplate(modified_template)
        doc.render(data)

        docx_path = os.path.join(tempfile.gettempdir(), f"Certificate_{name}.docx")
        pdf_path = os.path.join(tempfile.gettempdir(), f"Certificate_{name}.pdf")
        doc.save(docx_path)

        try:
            convert_to_pdf_asp(docx_path, pdf_path)
        except Exception as e:
            st.error(f"❌ Aspose conversion failed: {e}")
            pdf_path = docx_path

        try:
            send_email(email, pdf_path, data)
            st.success(f"✅ Certificate sent to {email}")
            with open(pdf_path, "rb") as f:
                st.download_button("📥 Download Certificate", f, file_name=os.path.basename(pdf_path))
        except Exception as e:
            st.error(f"❌ Email failed: {e}")

# --- Admin Panel ---
st.divider()
with st.expander("🔐 Admin Panel"):
    admin_key = st.text_input("Enter Admin Key", type="password")
    if admin_key == ADMIN_KEY:
        st.success("✅ Access granted.")
        if os.path.exists(CSV_FILE):
            df = pd.read_csv(CSV_FILE)
            if not df.empty:
                st.dataframe(df)
                with open(CSV_FILE, "rb") as f:
                    st.download_button("📥 Download CSV", f, file_name="completion_certificates.csv")
            else:
                st.info("CSV file is empty.")
        else:
            st.info("CSV log not found.")

        uploaded_csv = st.file_uploader("Upload Existing Intern CSV", type=["csv"])
        if uploaded_csv is not None:
            try:
                df_upload = pd.read_csv(uploaded_csv)
                df_upload.to_csv(CSV_FILE, index=False)
                st.success("✅ CSV uploaded and saved.")
            except Exception as e:
                st.error(f"❌ Failed to load CSV: {e}")
    elif admin_key:
        st.error("❌ Invalid key.")

# --- Footer ---
st.markdown("<hr><center><small>© 2025 SkyHighes Technologies. All Rights Reserved.</small></center>", unsafe_allow_html=True)

import streamlit as st
from fpdf import FPDF
import pandas as pd
import plotly.express as px
import base64
from datetime import datetime
from io import StringIO, BytesIO
from datetime import datetime
import os
import sqlite3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

 # Set page configuration
st.set_page_config(layout="wide")

def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)


def generate_html_download_link(fig):
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Download Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)


def get_risk(GPASem3, CGPA):
    if CGPA < 2.42:
        return 'High Risk'
    elif CGPA >= 2.42 and GPASem3 < 2.95:
        return 'Medium Risk'
    elif CGPA >= 2.96:
        return 'Low Risk'
    elif CGPA >= 2.42:
        return 'Medium Risk'


def home_page():

    st.image("images/banner.png", use_column_width=True)
    st.title("Welcome to Student Academic Performance Prediction System (SAPPS) 📈")
    st.write("SAPPS provides 3 menus:")
    st.write("1) Predict Risk Status")
    st.write("User can predict the risk status of students by uploading an excel file containing students' GPA and CGPA.")
    st.write("2) Predict Risk Status with Mitigation")
    st.write("User can predict the risk status of students and add mitigation for high-risk students.")
    st.write("3) Generate Graph")
    st.write("User can generate a graph from an uploaded excel file based on the gender, sponsorship and risk status of students.")
def generate_graph_page():
    
    
    st.title('GENERATE GRAPH 📈')
    st.write("In this page, you can:")
    st.write("1) Upload an excel file.")
    st.write("2) Choose suitable data that you want to analyze.")
    st.write("3) Download an image of the generated graph.")
    st.write("To generate a graph, please follow the excel template. Thank you 😊")
    image_path = 'images/generategraph.png'
    st.image(image_path, caption='\n\n')
    st.subheader('Import your excel file below to generate graph 👇')
    uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')
    if uploaded_file:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.dataframe(df)
        groupby_column = st.selectbox('What would you like to analyze?', ('Gender', 'Sponsorship', 'GPASem1', 'GPASem2','GPASem3','GPASem4','CGPA','Status Risk'))

        output_columns = ['Total Students', 'Student']
        df_grouped = df.groupby(by=[groupby_column], as_index=False)[output_columns].count()

        fig = px.bar(
            df_grouped,
            x=groupby_column,
            y='Total Students',
            color='Student',
            color_continuous_scale=['red', 'yellow', 'green'],
            template='plotly_white',
            title=f'<b>Total Students by {groupby_column}</b>'
        )
        st.plotly_chart(fig)

def display_info():
    st.balloons()     


def predict_risk_status_page():
    st.title('PREDICT RISK STATUS 🎯')
    st.write("In this page, you can:")
    st.write("1) Upload an excel file.")
    st.write("2) View risk status for each student predicted by the system automatically.")
    st.write("3) Download the excel file with predicted risk status.")
    st.subheader('Import your excel file below to predict status 👇')
    uploaded_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.subheader("Original Uploaded File")
            st.write(df)

            df['Status Risk'] = df.apply(lambda row: get_risk(row['GPASem3'], row['CGPA']), axis=1)

            st.subheader("Student Status Risk")
            st.write(df)
            
            
            if st.button("Download Risk Status"):
                st.write("Move your mouse cursor over the button to see the info.")
                timestamp = datetime.now().strftime("%d%m%H%M")
                filename = f"student_riskstatus_{timestamp}.xlsx"
                df.to_excel(filename, index=False)

                with open(filename, "rb") as file:
                    b64_data = base64.b64encode(file.read()).decode()
                    file.close()

                href = f'<a href="data:application/octet-stream;base64,{b64_data}" download="{filename}">Download Risk Status</a>'
                st.markdown(href, unsafe_allow_html=True)

            st.on_hover(display_info)

        except Exception as e:
            st.error("Error occurred while reading the file.")



            
# Connect to the SQLite database
conn = sqlite3.connect('students.db')
c = conn.cursor()

# Create the table if it doesn't exist
c.execute('''CREATE TABLE IF NOT EXISTS student_data
             (id INTEGER PRIMARY KEY AUTOINCREMENT,
              student_name TEXT,
              student_semester INTEGER,
              GPASem1 REAL,
              GPASem2 REAL,
              GPASem3 REAL,
              GPASem4 REAL,
              student_id TEXT,
              student_gender TEXT,
              student_sponsorship TEXT,
              CGPA REAL,
              risk_status TEXT,
              mitigation TEXT)''')
conn.commit()

def save_form_data(data):
    c.execute('''INSERT INTO student_data (student_name, student_semester, GPASem1, GPASem2, GPASem3, GPASem4,
                                           student_id, student_gender, student_sponsorship, CGPA, risk_status, mitigation)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (data['student_name'], data['student_semester'], data['GPASem1'], data['GPASem2'], data['GPASem3'],
               data['GPASem4'], data['student_id'], data['student_gender'], data['student_sponsorship'], data['CGPA'], data['risk_status'], data['mitigation']))
    conn.commit()

class PDF(FPDF):
    def header(self):
        # Set header image
        pass

    def footer(self):
        pass



def export_to_pdf(data):
    pdf = PDF()
    pdf.add_page()


 # Adjust margins for A4 paper
    pdf.set_left_margin(15)  # Adjust this value as needed
    pdf.set_right_margin(15)  # Adjust this value as needed
    pdf.set_top_margin(15)  # Adjust this value as needed
    pdf.set_auto_page_break(auto=True, margin=15)  # Adjust this value as needed

    pdf.set_font("Arial", "BU", size=14)

    # Set header image with adjusted size
    header_image_width = 100
    header_image_height = 35
    pdf.image("images/header_image.png", x=pdf.w / 2 - header_image_width / 2, y=10, w=header_image_width, h=header_image_height)
    pdf.ln(35)  # Adjust the spacing between header image and data

    # New heading with two lines
    pdf.cell(0, 10, "PREDICTION FORM REPORT", ln=True, align="C")
    pdf.cell(0, 10, "STUDENT ACADEMIC PERFORMANCE PREDICTION REPORT", ln=True, align="C")  # Add the second line here
    pdf.ln(10)

    # Add Student Information heading row
    pdf.set_fill_color(192, 192, 192)  # Light gray fill color for heading row
    pdf.set_font("Arial", "B", size=12)
    pdf.cell(0, 10, "STUDENT INFORMATION", border=1, ln=True, align="C", fill=True)
    
    # Adjust the layout by modifying the cell positioning and width
    left_margin = 20
    top_margin = pdf.get_y()
    left_column_width = 60  # Adjust the width of the left column
    right_column_width = 120  # Adjust the width of the right column
    cell_height = 10
    spacing = 5
    
    # Set the alignment and borders for the table
    pdf.set_fill_color(255, 255, 255)
    pdf.set_font("Arial", "B", size=12)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_text_color(0, 0, 0)

    for index, (key, value) in enumerate(data.items()):
        column = index % 2
        row = index // 2

        x = left_margin + column * (left_column_width + right_column_width + spacing)
        y = top_margin + row * (cell_height + spacing)

        # Customize the heading for each data
        if key == "student_name":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "Student Name", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "student_id":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "Student ID", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "student_semester":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "Semester", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "student_gender":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "Gender", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "student_sponsorship":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "Sponsorship", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "GPASem1":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "GPA Semester 1", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "GPASem2":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "GPA Semester 2", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "GPASem3":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "GPA Semester 3", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "GPASem4":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "GPA Semester 4", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)
        elif key == "CGPA":
            pdf.set_font("Arial", "B", 12)
            pdf.cell(left_column_width, cell_height, "CGPA", border=1, ln=False, fill=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(right_column_width, cell_height, str(value), border=1, ln=True)

        # Table 2: Risk Status and Mitigation
    pdf.set_font("Arial", "BU", size=14)  # Set font for the table heading
    pdf.cell(0, 15, "", ln=True, align="C")  # Centered heading for the table
    pdf.set_font("Arial", size=12)  # Set font back to normal for table content

    # Customize the content for risk status and mitigation
    risk_status = data['risk_status']
    mitigation = data['mitigation']

    # Centered table with one column and one row
    pdf.set_fill_color(255)  # White fill color for table cells
    pdf.set_font("Arial", "B", size=12)  # Set font for the table heading
    pdf.cell(0, 8, "Risk Status: ", border="LTR", ln=True, align="C", fill=True)
    pdf.set_font("Arial", size=12)  # Set font for the table heading
    pdf.cell(0, 8, risk_status, border="LR", ln=True, align="C", fill=True)
    
   
    pdf.set_font("Arial", "B", size=12)  # Set font for the mitigation heading (bold)
    pdf.cell(0, 8, "Mitigation:", border="LTR", ln=True, align="C", fill=True)  # Draw the cell with borders
    pdf.set_font("Arial", size=12)  # Set font for the mitigation content (unbold)
    pdf.multi_cell(0, 8, mitigation, border="LRB", align="C", fill=True)  # Draw the cell with borders and wrap the text
  
    pdf.output("current_form_data.pdf")

def send_email(file_path, recipient_email, student_name):
    # Email configuration
    sender_email = "farahfafar@gmail.com"  # Replace with your email
    sender_password = "yueifnnonyvkpwgk"  # Replace with your email password
    smtp_server = "smtp.gmail.com"  # Replace with your SMTP server if using a different provider
    smtp_port = 587  # Replace with your SMTP server port if using a different provider

    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "SAPPS - Prediction Form Report"

    # Email body
    body = f"Prof. / Dr. / Mr. / Mrs. / Ms.\n\n"
    body += f"STUDENT ACADEMIC PERFORMANCE PREDICTION REPORT,\n\n"
    body += f"This is a copy of the REPORT for student name {student_name}.\n\n"
    body += f"View attached report below\n\n"
    body += f"2023 - UNIVERSITI POLY-TECH MALAYSIA"
    msg.attach(MIMEText(body, 'plain'))

    # Attach the PDF report file
    with open(file_path, "rb") as file:
        part = MIMEApplication(file.read(), Name="form_data.pdf")
        part['Content-Disposition'] = f'attachment; filename="{file_path}"'
        msg.attach(part)

    # Send the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email. Error: {e}")

def new_form():
    st.title('📝 PREDICTION FORM')
    st.write("In this page, you can:")
    st.write("1) Fill the form below to generate risk status for a single student.")
    st.write("2) View mitigation for risk status generated.")
  

    # Form inputs
    col1, col2 = st.columns(2)

    # Form inputs
    with col1:
        student_name = st.text_input("Student Name")
        student_semester = st.number_input("Student Semester", min_value=0)
        GPASem1 = GPASem2 = GPASem3 = GPASem4 = None
        GPASem1 = st.number_input("GPA Semester 1")
        if student_semester >= 2:
            GPASem2 = st.number_input("GPA Semester 2")
        if student_semester >= 4:
            GPASem4 = st.number_input("GPA Semester 4")
    with col2:
        student_id = st.text_input("Student ID")
        student_gender = st.selectbox("Gender", ["Male", "Female"])
        student_sponsorship = st.selectbox("Sponsorship", ["Y", "N"])
        if student_semester >= 3:
            GPASem3 = st.number_input("GPA Semester 3")
        CGPA = st.number_input("Cumulative GPA")

    if st.button("PREDICT"):
        if student_name and student_semester and student_id and student_gender and student_sponsorship and CGPA or GPASem1 or GPASem2 or GPASem3 or GPASem4:
            risk_status = get_risk(GPASem3, CGPA)
            

            
            if risk_status == 'High Risk':
                mitigation = "1) Minimize total credit hours for next semester \n2) Advise meeting with mentor and counselor \n3) Schedule extra classes \n4) Review course materials"
            elif risk_status == 'Medium Risk':
                mitigation = "Advising meeting with mentor and counselor"
            else:
                mitigation = "No mitigation needed"
                
            st.success(f"Prediction status for {student_name} is: {risk_status}  \n Mitigation status: {mitigation}")

                # Save the form data to the database
            data = {
                    'student_name': student_name,
                    'student_id': student_id,
                    'student_gender': student_gender,
                    'student_sponsorship': student_sponsorship,
                    'student_semester': student_semester,
                    'GPASem1': GPASem1,
                    'GPASem2': GPASem2,
                    'GPASem3': GPASem3,
                    'GPASem4': GPASem4,
                    'CGPA': CGPA,
                    'risk_status': risk_status,
                    'mitigation' : mitigation
                }   
        
            save_form_data(data)
            export_to_pdf(data)
            with open("current_form_data.pdf", "rb") as file:
                b64_data = base64.b64encode(file.read()).decode()
                file.close()

            href = f'<a href="data:application/octet-stream;base64,{b64_data}" download="form_data.pdf">Download Current Form Data (PDF)</a>'
            st.markdown(href, unsafe_allow_html=True)

        else:
            st.warning("Please fill in all the required fields.")
            
        # Prompt user to enter recipient's email address
    recipient_email = st.text_input("Recipient's Email Address")
    
    if recipient_email:
        file_path = "current_form_data.pdf"  # Update this with the correct file path of the PDF report
        send_email(file_path, recipient_email, student_name)
        st.success("Prediction form report sent successfully via email.")
    else:
        st.warning("Please enter recipient's email address.")
        

# Configure sidebar navigation
st.sidebar.title('SAPPS Menu')

menu_options = [
    '🏠 Home',
    '🎯 Predict Risk Status',
    '📊 Generate Graph',
    '📝 Prediction Form'
]
choice = st.sidebar.radio('Go to', menu_options)


if choice == '🏠 Home':
        home_page()
elif choice == '📊 Generate Graph':
        generate_graph_page()
elif choice == '🎯 Predict Risk Status':
        predict_risk_status_page()
elif choice == '📝 Prediction Form':
        new_form()


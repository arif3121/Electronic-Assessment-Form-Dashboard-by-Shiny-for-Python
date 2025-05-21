from shiny import App, ui, reactive, render
import pandas as pd
from fpdf import FPDF
import os 

# Load data from Excel
print("Loading data from Excel...")
data = pd.read_excel("data.xlsx")
print("Data loaded successfully.")

# Clean column names to remove leading/trailing spaces
data.columns = data.columns.str.strip()

# Define the UI
app_ui = ui.page_fluid(
    ui.h2("LSBU EEE Form Submission Dashboard"),
    ui.input_select("student_name", "Select Student", choices=list(data["Student Name"])),
    ui.input_select("student_id", "Select Student ID", choices=list(data["Student ID"])),
    ui.input_select("grade", "Select Grade", choices=list(data["Grade"])),
    ui.input_text_area("comments", "Comments"),
    ui.input_action_button("generate", "Generate PDF"),
    ui.output_text_verbatim("status")  # Output text for status messages
)

# Define the server logic
def server(input, output, session):
    @reactive.event(input.generate)
    def generate_pdf():
        try:
            print("Button clicked")
            # Get the selected student's details
            selected_student = data[
                (data["Student Name"] == input.student_name()) &
                (data["Student ID"] == input.student_id())
            ]
            print(f"Selected student: {selected_student}")

            if selected_student.empty:
                output.status.set("Error: Student not found.")
                return

            # Fetch the details from input
            student_name = input.student_name()
            student_id = input.student_id()
            grade = input.grade()
            comments = input.comments()

            # More debug prints
            print(f"Student Name: {student_name}")
            print(f"Student ID: {student_id}")
            print(f"Grade: {grade}")
            print(f"Comments: {comments}")

            # Generate the PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt="Student Assessment Form", ln=True, align="C")
            pdf.ln(10)  # Line break
            pdf.cell(200, 10, txt=f"Name: {student_name}", ln=True)
            pdf.cell(200, 10, txt=f"Student ID: {student_id}", ln=True)
            pdf.cell(200, 10, txt=f"Grade: {grade}", ln=True)
            pdf.multi_cell(200, 10, txt=f"Comments: {comments}")

            # Ensure you're saving the file to the correct location
            pdf_file = f"{student_name}_{student_id}_assessment.pdf"
            print(f"Trying to save PDF to: {pdf_file}")

            # Save the PDF
            pdf.output(pdf_file)

            if os.path.exists(pdf_file):
                print(f"PDF saved successfully: {pdf_file}")
            else:
                print(f"Error: PDF not saved!")

            # Set status
            output.status.set(f"PDF generated: {pdf_file}")
        except Exception as e:
            print(f"Error: {e}")
            output.status.set(f"Error: {e}")

    @output
    @render.text
    def status():
        return "Ready to generate PDF."

# Create the Shiny app
app = App(app_ui, server)

if __name__ == "__main__":
    print("Starting the app...")
    app.run(port=8002)
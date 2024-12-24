import os
import pandas as pd
import fitz  # PyMuPDF

def add_text_and_line_to_pdf(input_pdf_path, output_pdf_path, text, position, font_size=12, font_color=(0, 0, 0), line_position=(100, 650), line_color=(0, 0, 0), status=""):
    # Open the PDF file
    pdf_document = fitz.open(input_pdf_path)

    lineWidth = 1
    lineWidth2 = 1
    line_position2=(55, 258)  # Default line position
    
    if status == "Mr":
        line_position=(55, 258)                   
        line_position2=(160, 258)                   
        lineWidth=72
        lineWidth2=55
    elif status == "Mrs":
        line_position=(55, 258)                   
        line_position2=(186, 258)                   
        lineWidth=94
        lineWidth2=30
    elif status == "Miss":
        line_position=(55, 258)                   
        line_position2=(137, 258)                   
        lineWidth=45
        lineWidth2=79
    elif status == "Family":
        line_position=(55, 258)                   
        line_position2=(160, 258)                   
        lineWidth=124
        lineWidth2=0
    elif status == "Couple":
        line_position=(108, 258)                   
        line_position2=(160, 258)                   
        lineWidth=105
        lineWidth2=0

    # Loop through each page
    for page_number in range(len(pdf_document)):
        page = pdf_document[page_number]
        x, y = position  # Unpack the position tuple
        line_start_x, line_start_y = line_position  # Unpack the line position
        line_start_x2, line_start_y2 = line_position2  # Unpack the line position

        # Insert text with the custom font (Cursive)
        page.insert_text(
            (x, y),                              # Position (x, y)
            text,                                 # Text to add
            fontsize=font_size,                   # Font size
            color=font_color                      # Font color (R, G, B)
        )
        
        # Draw a line (start_x, start_y) to (end_x, end_y)
        page.draw_line(
            (line_start_x, line_start_y),         # Start position
            (line_start_x + lineWidth, line_start_y),   # End position (change the X value for line length)
            color=line_color,                     # Line color (R, G, B)
            width=1                               # Line width
        )

        page.draw_line(
            (line_start_x2, line_start_y2),         # Start position
            (line_start_x2 + lineWidth2, line_start_y2),   # End position (change the X value for line length)
            color=line_color,                     # Line color (R, G, B)
            width=1                               # Line width
        )

    # Save the updated PDF in the 'Invitation' folder
    pdf_document.save(output_pdf_path)
    pdf_document.close()
    print(f"Text and line added successfully to '{output_pdf_path}'.")

# Function to create PDF for each row in the Excel file
def create_pdfs_from_excel(excel_file_path, input_pdf_path, output_folder="NangiFriends"):
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Read the Excel file with pandas
    df = pd.read_excel(excel_file_path)
    
    # Loop through the rows and create a PDF for each row
    for index, row in df.iterrows():
        status = row['Status']
        name = row['Name']
        
        # Define file name based on name and status, and save it to the 'Invitation' folder
        output_pdf_path = os.path.join(output_folder, f"{name}.pdf")
        
        # Call the add_text_and_line_to_pdf function to generate PDF for each row
        add_text_and_line_to_pdf(
            input_pdf_path=input_pdf_path,            # Input PDF path
            output_pdf_path=output_pdf_path,          # Output PDF path in 'Invitation' folder
            text=name,                                # Name to add
            position=(105, 284),                      # Coordinates (x, y) for text
            font_size=14,                             # Font size
            font_color=(0, 0, 0),                     # Text color
            status=status                             # Status text to determine line and font
        )

# Usage Example
create_pdfs_from_excel(
    excel_file_path="nangi.xlsx",    # Excel file path containing 'status' and 'name' columns
    input_pdf_path="input.pdf"          # Path to the input PDF template
)

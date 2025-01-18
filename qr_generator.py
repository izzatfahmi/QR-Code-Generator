import os
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
import pandas as pd

# Input Excel file path
excel_file = "QR_DataSet_2.xlsx"  # Replace with the path to your Excel file
image_path_left = "logo1.jpg"  # Path to the left image/logo
image_path_right = "logo2.jpg"  # Path to the right image/logo

# Read Excel file
df = pd.read_excel(excel_file)

# Create a folder with the same name as the Excel file (without extension)
folder_name = os.path.splitext(excel_file)[0]
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    # Get the data for the QR code
    data = str(row['QR DataSet'])  # Replace 'QR DataSet' with your column name
    caption = data  # Use the data as the caption

    # Generate QR code as a vector (SVG-like format)
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    # Save the QR code as a pure matrix (list of lists)
    qr_matrix = qr.modules

    # Create a PDF canvas
    pdf_filename = os.path.join(folder_name, f"{caption}.pdf")  # Save inside the new folder
    pdf = canvas.Canvas(pdf_filename, pagesize=letter)
    width, height = letter

    # Define parameters
    qr_size = 200  # Size of the QR code in the PDF
    border_padding = 20  # Padding inside the border
    x_center = width / 2
    y_center = height / 2

    # Define the border dimensions
    border_width = qr_size + border_padding * 2
    border_height = qr_size + border_padding * 2 + 70  # Additional space for caption
    border_x = x_center - border_width / 2
    border_y = y_center - border_height / 2

    # Draw a rounded rectangle for the border
    pdf.setStrokeColor(colors.black)
    pdf.setLineWidth(3)
    pdf.roundRect(border_x, border_y, border_width, border_height, radius=15, stroke=1, fill=0)

    # Draw the QR code as a vector inside the border
    box_size = qr_size / len(qr_matrix)
    x_offset = border_x + border_padding
    y_offset = border_y + border_padding + 30  # Leave space for the caption below
    for row_index, row in enumerate(qr_matrix):
        for col_index, value in enumerate(row):
            if value:  # If the QR code module is black
                x = x_offset + col_index * box_size
                y = y_offset + (len(qr_matrix) - 1 - row_index) * box_size
                pdf.rect(x, y, box_size, box_size, stroke=0, fill=1)

    # Add a caption below the QR code
    caption_font_size = 35
    pdf.setFont("Helvetica-Bold", caption_font_size)
    caption_width = pdf.stringWidth(caption, "Helvetica-Bold", caption_font_size)
    pdf.drawString((width - caption_width) / 2, border_y + 10, caption)

    # Insert two logos above the QR code with different sizes
    left_logo_width = 70
    left_logo_height = 40
    right_logo_width = 55
    right_logo_height = 40
    gap_between_logos = 50  # Gap between the two logos

    if image_path_left and image_path_right:
        # Calculate the positions for the logos
        total_logo_width = left_logo_width + right_logo_width + gap_between_logos
        left_logo_x = x_center - total_logo_width / 2
        right_logo_x = left_logo_x + left_logo_width + gap_between_logos
        logo_y = border_y + border_height - max(left_logo_height, right_logo_height) - 10  # Align above the QR code based on tallest logo
        
        # Draw the left logo
        pdf.drawImage(ImageReader(image_path_left), left_logo_x, logo_y, width=left_logo_width, height=left_logo_height)
        
        # Draw the right logo
        pdf.drawImage(ImageReader(image_path_right), right_logo_x, logo_y, width=right_logo_width, height=right_logo_height)

    # Save the PDF
    pdf.save()

    print(f"Generated QR code PDF: {pdf_filename}")

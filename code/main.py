import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def read_data(excel_file_path, sheet_name_data, sheet_name_student_profile):
    # Read the Excel file
    df_data = pd.read_excel(excel_file_path, sheet_name=sheet_name_data)
    df_sheet1 = pd.read_excel(excel_file_path, sheet_name=sheet_name_student_profile)

    # Get the enrollment number from the "data" sheet
    enrollment_no = df_data['Enrollment No.'].iloc[-1]

    print('Enrollment number:', enrollment_no)

    # Retrieve the corresponding row values from "sheet1"
    corresponding_row = df_sheet1[df_sheet1['Enrollment No.'] == enrollment_no]

    # Print the corresponding row values
    print(corresponding_row)

    return corresponding_row, enrollment_no


def create_receipt(source_image, destination_image, corresponding_row):
    # Open the image
    image = Image.open(source_image)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    # Define the font and font size
    font = ImageFont.truetype('arial.ttf', 15)

    # Define the text to be added
    name_text = str(corresponding_row['Student Name'].values[0])
    class_text = str(corresponding_row['Class'].values[0])
    section_text = str(corresponding_row['Section'].values[0])
    # Get the current date
    date = datetime.date.today().strftime('%d/%m/%Y')

    # Define the positions to add the text
    enrollment_position = (50, 50)
    name_position = (340, 162)
    class_position = (700, 162)
    section_position = (720, 162)
    date_position = (910, 160)

    # Add the text to the image
    draw.text(name_position, name_text, fill='white', font=font)
    draw.text(class_position, class_text, fill='white', font=font)
    draw.text(section_position, section_text, fill='white', font=font)
    draw.text(date_position, str(date), fill='white', font=font)

    # Save the modified image
    image.save(destination_image)

    return image


def convert_to_pdf(destination_image, destination_pdf, image):
    # Define the maximum width and height for the image
    max_width = A4[0]
    max_height = A4[1]

    # Calculate the new width and height while maintaining the aspect ratio
    image_width, image_height = image.size
    aspect_ratio = image_width / image_height

    if image_width > max_width or image_height > max_height:
        if aspect_ratio > 1:
            image_width = max_width
            image_height = int(max_width / aspect_ratio)
        else:
            image_height = max_height
            image_width = int(max_height * aspect_ratio)

    # Create a new PDF file
    pdf_file = destination_pdf
    c = canvas.Canvas(pdf_file, pagesize=A4)

    # Set the image position and size on the page
    image_position = (0, A4[1] - image_height)  # Position the image at the top

    # Draw the image on the PDF canvas
    c.drawImage(destination_image, *image_position, width=image_width, height=image_height)

    # Save and close the PDF file
    c.save()

    print(f"PDF file '{pdf_file}' created successfully.")    


def cleanup_files(image_file, pdf_file):
    import os

    # Delete the image file
    if os.path.exists(image_file):
        os.remove(image_file)
        print(f"Image file '{image_file}' deleted.")

    # Delete the PDF file
    # if os.path.exists(pdf_file):
    #     os.remove(pdf_file)
    #     print(f"PDF file '{pdf_file}' deleted.")


def main(pdf_file):
    import win32print
    import win32api

    # Get the default printer
    printer_name = win32print.GetDefaultPrinter()
    win32api.ShellExecute(
        0,
        "open",
        pdf_file,
        f'/p /h "{pdf_file}" "{printer_name}"',
        ".",
        1
    )

    print(f"PDF file '{pdf_file}' sent to the printer.")




if __name__ == '__main__':
    corresponding_row, enrollment_no = read_data('Student Data.xls', 'data', 'Student Profile Detail Report')
    image = create_receipt('code/Coupon GCS.png', 'code/bin/Modified_Coupon.png', corresponding_row)
    convert_to_pdf('code/bin/Modified_Coupon.png', 'code/bin/Modified_Coupon.pdf', image)
    main('code/bin\Modified_Coupon.pdf')
    cleanup_files('code/bin/Modified_Coupon.png', 'code/bin/Modified_Coupon.pdf')

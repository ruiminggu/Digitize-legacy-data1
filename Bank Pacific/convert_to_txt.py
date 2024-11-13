import openpyxl
from PyPDF2 import PdfWriter, PdfReader
from PIL import Image, ImageOps, ImageEnhance
import pytesseract

import os
import re

from pdf2image import convert_from_path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def preprocess_image(image):
    # Enhance the image contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)

    # Convert to grayscale
    image = ImageOps.grayscale(image)

    return image

def detect_and_rotate_image(image_path):
    with Image.open(image_path) as img:
        # Use pytesseract to detect the orientation and script detection
        osd = pytesseract.image_to_osd(img)

        # Extract the rotation angle from OSD output
        rotation_angle = int(re.search(r'Rotate: (\d+)', osd).group(1))

        # Correct the rotation based on the detected angle
        if rotation_angle == 180:
            img = img.rotate(180, expand=True)
        elif rotation_angle == 90:
            img = img.rotate(270, expand=True)
        elif rotation_angle == 270:
            img = img.rotate(90, expand=True)

        # Preprocess the image
        img = preprocess_image(img)

        # Save the correctly oriented and preprocessed image
        img.save(image_path)

def split_pdf(output_dir,inputpdf):
    # Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    inputpdf = PdfReader(open(inputpdf, "rb"))
    for i in range(len(inputpdf.pages)):
        output = PdfWriter()
        output.add_page(inputpdf.pages[i])
        # Create the output file path
        output_file_path = os.path.join(output_dir, f"document-page{i + 1}.pdf")
        with open(output_file_path, "wb") as outputStream:
            output.write(outputStream)

def convert_folder_pdf_to_png(folder_path):
    # Create a directory for the output PNG files
    output_folder = os.path.join(folder_path, 'png_files')
    os.makedirs(output_folder, exist_ok=True)

    # Iterate over PDF files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_file_path = os.path.join(folder_path, filename)
            output_prefix = os.path.splitext(filename)[0]
            output_prefix = os.path.join(output_folder, output_prefix)

            # Convert PDF to PNG using the provided function
            convert_pdf_to_png(pdf_file_path, output_prefix)
    return output_folder

# helper function that convert one page pdf to png
def convert_pdf_to_png(input_file, output_prefix):
    try:
        # Convert PDF to a list of PIL.Image objects
        images = convert_from_path(input_file)

        # Save each image as PNG
        for i, image in enumerate(images):
            output_file = f"{output_prefix}_{i + 1}.png"
            image.save(output_file, "PNG")
            detect_and_rotate_image(output_file)

    except Exception as e:
        print("Conversion failed:", str(e))

def convert_folder_images(folder_path):
    # Create a directory for the output text files
    output_folder = os.path.join(folder_path, 'txt_files')
    os.makedirs(output_folder, exist_ok=True)

    # Iterate over PNG files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.png'):
            png_file_path = os.path.join(folder_path, filename)
            txt_file_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.txt')
            png_to_txt(png_file_path, txt_file_path)

# helper function that convery png to txt
def png_to_txt(png_file_path, txt_file_path):
    # Open the PNG image using PIL
    image = Image.open(png_file_path)

    # Apply OCR using pytesseract with specific configuration
    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(image, config=custom_config)

    # Write the extracted text to a text file
    with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
        txt_file.write(text)

def update_CTR_titlename(folder_path):
    for filename in os.listdir(folder_path):
        if filename.startswith("CTR"):
            file_path = os.path.join(folder_path, filename)
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            # Assuming the header row is the first row, change the row number if needed
            header_row = sheet[1]

            for cell in header_row:
                if "cash_direction" == str(cell.value).lower():
                    cell.value = "cashDirection"
                if "date" == str(cell.value).lower():
                    cell.value = "dateOfTransaction"
                if "cash_amount" == str(cell.value).lower():
                    cell.value = "cashAmount"
                if "ctrid" == str(cell.value).lower():
                    cell.value = "CTRID"
            workbook.save(file_path)

def update_PIT_titlename(folder_path):
    for filename in os.listdir(folder_path):
        if filename.startswith("PIT"):
            file_path = os.path.join(folder_path, filename)
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            # Assuming the header row is the first row, change the row number if needed
            header_row = sheet[1]

            for cell in header_row:
                if "name" == str(cell.value).lower():
                    cell.value = "lastNameOrNameOfEntity"
                if "occupation" == str(cell.value).lower():
                    cell.value = "occupationOrTypeOfBusiness"
                if "ctrid" == str(cell.value).lower():
                    cell.value = "CTRID"
                if "city" == str(cell.value).lower():
                    cell.value = "addressCity"
                if "state" == str(cell.value).lower():
                    cell.value = "addressState"
                if "zipcode" == str(cell.value).lower():
                    cell.value = "zipCode"
                if "address_country" == str(cell.value).lower():
                    cell.value = "addressCountry"
                if "dob" == str(cell.value).lower():
                    cell.value = "dateOfBirth"
                if "contact_number" == str(cell.value).lower():
                    cell.value = "contactPhoneNumber"
                if "id_type" == str(cell.value).lower():
                    cell.value = "idType"
                if "id_number" == str(cell.value).lower():
                    cell.value = "idNumber"
                if "id_country" == str(cell.value).lower():
                    cell.value = "idCountry"
                if "account_number" == str(cell.value).lower():
                    cell.value = "accountNumbers"
                if "cash_direction" == str(cell.value).lower():
                    cell.value = "cashDirection"
                if "cash_amount" == str(cell.value).lower():
                    cell.value = "cashAmount"
                if "accountnumber" == str(cell.value).lower():
                    cell.value = "accountNumbers"
                if "relationship" == str(cell.value).lower():
                    cell.value = "relationshipToTransaction"

            workbook.save(file_path)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # the first should enter the name of the destination folder, the second is the pdf name
    split_pdf("bankofP","20230802194518.pdf")
    # enter the first again.
    convert_folder_images(convert_folder_pdf_to_png("bankofP"))


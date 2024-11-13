import os
import subprocess
import sys


def install_packages():
    # Install required Python packages
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])


def check_and_install_tesseract():
    # Check if Tesseract is installed
    try:
        subprocess.check_call(['tesseract', '-v'])
    except subprocess.CalledProcessError:
        print("Tesseract is not installed or not found in PATH.")
        print("Please install Tesseract OCR from https://github.com/tesseract-ocr/tesseract and add it to your PATH.")
        sys.exit(1)


def add_python_to_path():
    python_path = os.path.dirname(sys.executable)
    scripts_path = os.path.join(python_path, 'Scripts')
    current_path = os.environ.get('PATH', '')
    new_path = f"{python_path};{scripts_path};{current_path}"
    os.environ['PATH'] = new_path


def main():
    # Ensure we are in the correct directory
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Add Python and pip to PATH if not already there
    add_python_to_path()

    # Install necessary packages
    install_packages()

    # Check and install Tesseract OCR
    check_and_install_tesseract()

    # Define the name of the PDF file (assumed to be in the same directory as this script)
    pdf_name = 'Bank_of_Hawaii.pdf'

    if not os.path.isfile(pdf_name):
        print(f"PDF file '{pdf_name}' not found in the current directory.")
        sys.exit(1)

    # Run the provided code to process the PDF
    import convert_to_txt
    import bankofhawaii

    output_dir = 'output'

    # Split the PDF into single pages
    convert_to_txt.split_pdf(output_dir, pdf_name)

    # Convert the PDF pages to PNG images
    convert_to_txt.convert_folder_images(convert_to_txt.convert_folder_pdf_to_png(output_dir))

    # Process the images and extract information to an Excel file
    bankofhawaii.read_folder_PIT(output_dir, 'output_pit.xlsx')
    bankofhawaii.read_folder_CTR(output_dir, 'output_ctr.xlsx')


if __name__ == '__main__':
    main()

import os
import sys

import pytesseract
import PyPDF2
import io

poppler_path = r'C:\Program Files\poppler-22.04.0-hea5ffa9_2\Library\bin'
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    input_folder = "input"
    output_file = "output/searchable"

    mode = str(input("Type 'A' to enter Advanced Mode, and press any other key for Simple Mode."))
    if mode.lower() == 'a':
        print()
        print("What folder should we look for the input input in?")
        print("DEFAULT: \"" + input_folder + "\"")
        input_folder = str(input("Specify a folder name in the current directory or file path.\n"))

        print()
        print("What file should we output to?")
        print("DEFAULT: \"" + output_file + "\"")
        output_file = str(input("Specify a file name in the current directory; DO NOT include the extension.\n"))

    print()
    print("Scanning all files in " + input_folder + "...")

    '''
    Get all available files in the given directory
    and append them to the set object "all_files."
    '''
    all_files = []
    for (path, dirs, files) in os.walk(input_folder):
        for file in files:
            file = os.path.join(path, file)
            all_files.append(file)

    # Ensure that there is >1 file in the given input folder
    num_files = all_files.__len__()
    if num_files < 1:
        print(input_folder + " either has no files or doesn't exist.")
        print("Exiting program.")
        sys.exit()

    '''
    Use PyTesseract to OCR (Optical Character Recognition)
    every file, and add it to the PDF Writer object.
    '''
    pdf_writer = PyPDF2.PdfFileWriter()
    failed_files = 0
    for file in all_files:
        try:
            page = pytesseract.image_to_pdf_or_hocr(file, extension='pdf')
            pdf = PyPDF2.PdfFileReader(io.BytesIO(page))
            pdf_writer.addPage(pdf.getPage(0))
        except Exception as e:
            print("Skipping " + file + "; " + str(e))
            failed_files += 1

    '''
    Open the PDF file for  writing, truncating the file first
    and in binary mode. Write all recognized pages from the above
    region into a single PDF file.
    '''
    with open(output_file + ".pdf", "wb") as page_to_write:
        pdf_writer.write(page_to_write)

    num_pages = pdf_writer.pages.__len__()
    print("Scan complete! ")
    print(f"Successfully converted { num_files - failed_files } input file(s) into a single PDF with { num_pages } pages.")

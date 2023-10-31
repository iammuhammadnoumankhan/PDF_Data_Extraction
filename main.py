import PyPDF2
import pytesseract
from pdf2image import convert_from_path
import re
import os
import csv
import yaml

try:
    # Load configuration from config.yaml
    with open('config.yaml', 'r') as config_file:
        config = yaml.safe_load(config_file)
except FileNotFoundError:
    print("Error: config.yaml file not found.")
    exit(1)

# Function to extract text from an image
def extract_text_from_image(image_path):
    try:
        text = pytesseract.image_to_string(image_path)
        return text
    except Exception as e:
        print(f"Error extracting text from image: {e}")
        return ""

def extract_revenue_values(text):
    # Define regular expression to match revenue values
    pattern_1 = r'Revenue\s\d+\s([\d,]+)\s([\d,]+)'
    pattern_2 = r'Turnover\s+\d+\s+([\d,]+)\s+([\d,]+)'
    
    matches = re.search(pattern_1, text)
    if matches:
        year1 = matches.group(1).replace(',', '')  # Remove commas from numbers
        year2 = matches.group(2).replace(',', '')
        return year1, year2
    else:
        matches = re.search(pattern_2, text, re.IGNORECASE)
        if matches:
            value1 = matches.group(1).replace(',', '')  # Remove commas from the number
            value2 = matches.group(2).replace(',', '')
            return value1, value2
        else:
            return None

# Create or update a CSV file to store results
def create_or_update_csv(file_path, pdf_path, year1, year2):
    try:
        if not os.path.isfile(file_path):
            # Create the CSV file with headers
            with open(file_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['pdf_path', '2022', '2021'])

        # Append results to the CSV file
        with open(file_path, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([pdf_path, year1, year2])
    except Exception as e:
        print(f"Error writing to CSV file: {e}")

# Directory path containing PDF files
pdf_directory = config.get('pdf_directory', '')
output_csv_file = config.get('output_csv_file', '')

if not pdf_directory:
    print("Error: PDF directory path is not specified in config.yaml.")
    exit(1)

if not output_csv_file:
    print("Error: Output CSV file path is not specified in config.yaml.")
    exit(1)

try:
    for pdf_file in os.listdir(pdf_directory):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_directory, pdf_file)

            # Convert the PDF pages to images
            images = convert_from_path(pdf_path)
            page_text = ""
            for page_num, image in enumerate(images):
                # Extract text from the image using OCR
                page_text += extract_text_from_image(image)

            # Extract revenue values
            result = extract_revenue_values(page_text)
            if result:
                year1, year2 = result
                print(f"{pdf_path}: 2022: {year1}, 2021: {year2}")
                create_or_update_csv(output_csv_file, pdf_path, year1, year2)
            else:
                print(f"{pdf_path}: Revenue values not found.")

    print("Processing complete. Results saved to", output_csv_file)
except Exception as e:
    print(f"An error occurred: {e}")


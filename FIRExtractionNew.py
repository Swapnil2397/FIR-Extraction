import os
import csv
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import re
import pandas as pd
import multiprocessing
from concurrent.futures import ProcessPoolExecutor
from googletrans import Translator
import xlsxwriter

# Define the paths
pdf_folder_path = '/Users/sukanyasaha/Desktop/Pdfs'
output_folder_path = '/Users/sukanyasaha/Desktop/Pdfs/Output'
tesseract_cmd_path = r'/opt/homebrew/bin/tesseract'  # Updated path for Tesseract

# Ensure the output folder exists
os.makedirs(output_folder_path, exist_ok=True)

# Define the path to the Tesseract executable (if not in your PATH already)
pytesseract.pytesseract.tesseract_cmd = tesseract_cmd_path

# Get a list of all PDF files from the folder and its subfolders
pdf_files = []
for root, _, files in os.walk(pdf_folder_path):
    for file in files:
        if file.endswith('.pdf'):
            pdf_files.append(os.path.join(root, file))

translator = Translator()

# Extract first three words after detecting "P.S."
def extract_ps(text):
    lines = text.splitlines()
    capturing = False
    captured_text = []

    for line in lines:
        # Start capturing text after detecting "P.S."
        if "P.S." in line:
            capturing = True
            # Split the content after "P.S." and take the first three words
            words = line.split("P.S.")[1].strip().split()
            captured_text.extend(words[:3])  # Capture only the first three words

            # Stop capturing if three words are already captured
            if len(captured_text) >= 3:
                break
        elif capturing:
            # Additional checks to stop capturing on specific patterns like "Year"
            if "2024" in line or "Year" in line:
                break
            # Add more words if needed, stopping at a total of three
            words = line.strip().split()
            captured_text.extend(words[:max(3 - len(captured_text), 0)])  # Add words to reach a total of three

        # Ensure we stop if exactly three words are captured
        if len(captured_text) >= 3:
            break

    # Join and return exactly three words
    return " ".join(captured_text[:3]).strip()

# Line-by-line location extraction
def extract_location(text):
    capturing = False
    captured_text = []

    # Split the text into lines
    lines = text.splitlines()

    for line in lines:
        # Start capturing after detecting "Address (पत्ता):"
        if "Address (पत्ता):" in line:
            capturing = True
            captured_text.append(line.split("Address (पत्ता):")[1].strip())
        elif capturing:
            # Stop capturing when "N.C.R.B" or "In case, outside" is found
            if "N.C.R.B" in line or "In case, outside" in line:
                break
            captured_text.append(line.strip())

    return " ".join(captured_text).strip()

def extract_summary(text):
    capturing = False
    captured_text = []

    # Split the text into lines
    lines = text.splitlines()

    for line in lines:
        # Start capturing if "S.No. UIDB Number" or "First Information contents" is found
        if "S.No. UIDB Number" in line or "First Information contents" in line:
            capturing = True
            # Capture the part of the line after the marker if present
            if "S.No. UIDB Number" in line:
                captured_text.append(line.split("S.No. UIDB Number")[1].strip())
            elif "First Information contents" in line:
                captured_text.append(line.split("First Information contents")[1].strip())
        elif capturing:
            # Stop capturing when "Action taken" is found
            if "Action taken" in line:
                break
            captured_text.append(line.strip())

    return " ".join(captured_text).strip()


# Function to extract text between specific markers and take the first word only
def extract_segments(text):
    segments = {}

    # Helper function to return all matched content between the given phrases
    def get_all_text(match):
        if match:
            return match.group(1).strip()  # Return the full text between the match
        return ''

    # Extract section for date and time using a more flexible pattern
    section_pattern = r"Occurrence of offence.*?(?=Action taken:|$)"
    section_match = re.search(section_pattern, text, re.DOTALL)

    if section_match:
        section_text = section_match.group(0)
        # print(f"Extracted section:\n{section_text}\n")  # Debugging statement

        # Patterns to capture the first date and time
        date_pattern = r"([0-9]{2}/[0-9]{2}/[0-9]{4})"
        time_pattern = r"([0-9]{1,2}:[0-9]{2})"
        
        # Match patterns to extract the first date and time
        date_match = re.search(date_pattern, section_text)
        time_match = re.search(time_pattern, section_text)
        
        # Extract date and time or return "N/A" if not found
        date = date_match.group(1) if date_match else "N/A"
        time = time_match.group(1) if time_match else "N/A"
        
        # print(f"Extracted Date: {date}, Extracted Time: {time}")  # Debugging statement
        
        segments['Date'] = date
        segments['TIME'] = time
    else:
        # print("Section not found.")
        segments['Date'] = ""
        segments['TIME'] = ""

    # Example of other extractions, using PS extraction function
    segments['PS'] = extract_ps(text)
    segments['Location'] = extract_location(text)
    segments['Summary'] = extract_summary(text)

    # Other regex-based extractions (e.g., District)
    district_match = re.search(r'District(.*?)(P\.S\.|FIR No\.)', text, re.DOTALL)
    segments['District'] = get_all_text(district_match)

    # FIR No. extraction
    fir_no_match = re.search(r'FIR No\.(.*?)Date &Time of FIR', text, re.DOTALL)
    segments['FIR No.'] = get_all_text(fir_no_match)

    # Return the filled segments dictionary
    return segments

# Function to translate segments
def translate_segments(segments):
    translated_segments = {}
    for key, value in segments.items():
        if value:
            translated_text = translator.translate(value, src='hi', dest='en').text
            translated_segments[f'{key} (EN)'] = translated_text
        else:
            translated_segments[f'{key} (EN)'] = value
    return translated_segments

# Prepare the CSV file to store the extracted data
csv_filename = os.path.join(output_folder_path, 'extracted_data.csv')
csv_columns = [
    'Text Filename', 'District', 'District (EN)', 'PS', 'PS (EN)', 'FIR No.', 'FIR No. (EN)', 
    'Date', 'Date (EN)', 'TIME', 'TIME (EN)', 'Location', 'Location (EN)', 'Summary', 'Summary (EN)', 
    'Summary1', 'Summary1 (EN)', 'GPS Co-ordinate (Lat,Long)', 'Number of Fatalities', 
    'Number of Serious Injuries', 'Number of Minor Injuries', 'Total number of vehicle involved', 
    'Total number of pedestrians involved', 'Crash Configuration', 'Crash Contributing Factor', 
    'Injury Contributing Factor', 'Offending Vehicle', 'Victim Vehicle'
]

# Function to process a single PDF
def process_pdf(pdf_path):
    try:
        # Convert PDF to images
        images = convert_from_path(pdf_path)

        # Prepare to collect all extracted text in one file
        text_filename = os.path.join(output_folder_path, f"{os.path.splitext(os.path.basename(pdf_path))[0]}.txt")
        all_extracted_text = {key: '' for key in csv_columns}
        all_extracted_text['Text Filename'] = text_filename

        with open(text_filename, 'w', encoding='utf-8') as txt_file:
            # Extract text from each image
            for i, image in enumerate(images):
                image_filename = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page_{i+1}.png"
                image_path = os.path.join(output_folder_path, image_filename)
                image.save(image_path, 'PNG')

                text = pytesseract.image_to_string(image, lang='eng+hin')
                txt_file.write(text + "\n\n")

                segments = extract_segments(text)
                translated_segments = translate_segments(segments)

                # Store the extracted segments and translations
                for key, value in {**segments, **translated_segments}.items():
                    if key in all_extracted_text:
                        all_extracted_text[key] += f" {value}".strip()

        # Write the collected data to the CSV file
        with open(csv_filename, 'a', newline='', encoding='utf-8') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=csv_columns)
            writer.writerow(all_extracted_text)
        print(f"Processed {pdf_path}, extracted data written to CSV and text file")

    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")

if __name__ == '__main__':
    # Initialize the CSV file with headers
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=csv_columns)
        writer.writeheader()

    # Process each PDF using multiprocessing
    with ProcessPoolExecutor(max_workers=multiprocessing.cpu_count()) as executor:
        executor.map(process_pdf, pdf_files)

    print("All PDFs have been processed and data has been written to CSV and text files.")

    # Convert CSV to Excel with hyperlinks
    df = pd.read_csv(csv_filename)
    xlsx_filename = os.path.join(output_folder_path, 'extracted_data.xlsx')

    writer = pd.ExcelWriter(xlsx_filename, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    for i, text_filename in enumerate(df['Text Filename']):
        worksheet.write_url(i + 1, 0, f'file://{text_filename}', string=os.path.basename(text_filename))

    writer.close()

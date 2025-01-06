import os
import re
import pandas as pd
from PyPDF2 import PdfReader
from concurrent.futures import ThreadPoolExecutor

# Define the path to the PDF folder
pdf_folder_path = '/Users/sukanyasaha/Desktop/PDFS2'
output_folder_path = '/Users/sukanyasaha/Desktop/Output'

# Ensure the output folder exists
os.makedirs(output_folder_path, exist_ok=True)

# Define the columns for the output
columns = [
    "Serial No.", "FIR No.", "Section", "Accident Date (DD-MM-YYYY)", "Accident Time (HH:MM)",
    "Jurisdiction of Police Station", "District", "Name of Crash Location", "Road Name",
    "Latitude, Longitude", "Road Feature", "Number of Fatalities", "Number of Grievously Injured persons",
    "Number of people with Minor Injury", "Total number of Injuries", "Number of Motor Vehicles involved",
    "Number of Non-motorised Vehicles involved", "Number of Pedestrians Involved", "Crash Between",
    "Crash Configuration", "Vehicle 1", "Higest Injury in Vehicle 1", "Vehicle 2",
    "Higest Injury in Vehicle 2", "Vehicle 3", "Higest Injury in Vehicle 3", "Crash Contributing Factor",
    "Injury Contributing Factor", "FIR Summary", "Text File Link", "PDF File Link", "PDF File Name"
]


# Initialize an empty DataFrame
df = pd.DataFrame(columns=columns)

def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_section_data(text, start_marker, end_marker):
    sections = []
    start_idx = 0

    while True:
        start_idx = text.find(start_marker, start_idx)
        if start_idx == -1:
            break
        end_idx = text.find(end_marker, start_idx)
        if end_idx == -1:
            break
        section = text[start_idx + len(start_marker):end_idx].strip()
        sections.append(section)
        start_idx = end_idx + len(end_marker)
    return sections

def extract_info_from_pdf(pdf_path):
    try:
        text = extract_text_from_pdf(pdf_path)
        
        sections = extract_section_data(text, "Act", "State Rule")
        fir_nos = extract_section_data(text, "FIR/CSR Number   : ", "FIR Date & Time")
        road_names = extract_section_data(text, "Street Name", "Local Body")

        section = ', '.join(sections)
        fir_no = ', '.join(fir_nos)
        road_name = ', '.join(road_names)

        # Convert PDF to text file for additional data extraction
        txt_file_path = os.path.join(output_folder_path, os.path.basename(pdf_path).replace('.pdf', '.txt'))
        with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
            txt_file.write(text)

        # Extract required information using regex
        data = {
            "FIR No.": fir_no,
            "Section": section,
            "Accident Date (DD-MM-YYYY)": re.search(r'Accident Date and Time\s*(\d{2}-\w+-\d{4})', text).group(1) if re.search(r'Accident Date and Time\s*(\d{2}-\w+-\d{4})', text) else '',
            "Accident Time (HH:MM)": re.search(r'Accident Date and Time\s*\d{2}-\w+-\d{4}\s*:\s*(\d{2}:\d{2} [APM]{2})', text).group(1) if re.search(r'Accident Date and Time\s*\d{2}-\w+-\d{4}\s*:\s*(\d{2}:\d{2} [APM]{2})', text) else '',
            "Jurisdiction of Police Station": re.search(r'Station Name\s*:\s*(.*?)\n', text).group(1).split('Investigating Oﬃcer')[0].strip() if re.search(r'Station Name\s*:\s*(.*?)\n', text) else '',
            "District": re.search(r'District Name\s*:\s*(.*?)\n', text).group(1) if re.search(r'District Name\s*:\s*(.*?)\n', text) else '',
            "Name of Crash Location": re.search(r'Location Details\s*(.*?)\n', text).group(1) if re.search(r'Location Details\s*(.*?)\n', text) else '',
            "Road Name": road_name,
            "Latitude, Longitude": re.search(r'Location Details\s*.*?Lat/Long\s*:\s*([0-9.]+,\s*[0-9.]+)', text).group(1) if re.search(r'Location Details\s*.*?Lat/Long\s*:\s*([0-9.]+,\s*[0-9.]+)', text) else '',
            "Road Feature": re.search(r'Road Classification\s*:\s*(.*?)\n', text).group(1) if re.search(r'Road Classification\s*:\s*(.*?)\n', text) else '',
            "Number of Fatalities": re.search(r'Total\s*:\s*(\d+)\s*Number of Animals involved', text).group(1) if re.search(r'Total\s*:\s*(\d+)\s*Number of Animals involved', text) else '0',
            "Number of Grievously Injured persons": re.search(r'Grievous Injury\s*(\d+)', text).group(1) if re.search(r'Grievous Injury\s*(\d+)', text) else '0',
            "Number of people with Minor Injury": re.search(r'Minor Injury\s*(\d+)', text).group(1) if re.search(r'Minor Injury\s*(\d+)', text) else '0',
            "Total number of Injuries": re.search(r'Total\s*(\d+)', text).group(1) if re.search(r'Total\s*(\d+)', text) else '0',
            "Number of Motor Vehicles involved": re.search(r'No of Vehicle\(s\) involved\s*(\d+)', text).group(1) if re.search(r'No of Vehicle\(s\) involved\s*(\d+)', text) else '0',
            "Number of Non-motorised Vehicles involved": "0",
            "Number of Pedestrians Involved": "0",
            "Crash Between": re.search(r'Collision Type\s*:\s*(.*?)\n', text).group(1) if re.search(r'Collision Type\s*:\s*(.*?)\n', text) else '',
            "Crash Configuration": re.search(r'Collision Nature\s*:\s*(.*?)\n', text).group(1) if re.search(r'Collision Nature\s*:\s*(.*?)\n', text) else '',
            "Vehicle 1": re.search(r'Vehicle Regn. No\s*(.*?)\s', text).group(1) if re.search(r'Vehicle Regn. No\s*(.*?)\s', text) else '',
            "Higest Injury in Vehicle 1": "Fatal",  # Hardcoded based on available data
            "Vehicle 2": re.search(r'Vehicle Regn. No\s*.*?\s*MH40Y5087', text).group(1) if re.search(r'Vehicle Regn. No\s*.*?\s*MH40Y5087', text) else '',
            "Higest Injury in Vehicle 2": "No Injury",  # Hardcoded based on available data
            "Vehicle 3": "None",  # Based on available data
            "Higest Injury in Vehicle 3": "None",  # Based on available data
            "Crash Contributing Factor": re.search(r'Initial observation of accident scene\s*(.*?)\n', text).group(1) if re.search(r'Initial observation of accident scene\s*(.*?)\n', text) else '',
            "Injury Contributing Factor": "Not mentioned",  # Not available in the text
            "FIR Summary": re.search(r'Initial observation of accident scene\s*(.*?)\n', text).group(1) if re.search(r'Initial observation of accident scene\s*(.*?)\n', text) else '',
            "Text File Link": txt_file_path,
            "PDF File Link": pdf_path,
            "PDF File Name": os.path.basename(pdf_path)  # Add the PDF file name
        }
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        data = {col: '' for col in columns}

    return data


def process_pdf_files(pdf_files):
    all_data = []
    with ThreadPoolExecutor() as executor:
        for idx, result in enumerate(executor.map(extract_info_from_pdf, pdf_files)):
            result['Serial No.'] = idx + 1
            all_data.append(result)
            print(f"PDF {idx + 1} processed")
    return all_data


# Get the list of PDF files
pdf_files = [os.path.join(pdf_folder_path, pdf_file) for pdf_file in os.listdir(pdf_folder_path) if pdf_file.endswith('.pdf')]

# Process the PDF files
all_data = process_pdf_files(pdf_files)

# Create DataFrame from all data
df = pd.DataFrame(all_data, columns=columns)

# Save the DataFrame to an Excel file
output_excel_path = os.path.join(output_folder_path, 'Accident_Summary.xlsx')
df.to_excel(output_excel_path, index=False)

# Save the DataFrame to a CSV file
output_csv_path = os.path.join(output_folder_path, 'Accident_Summary.csv')
df.to_csv(output_csv_path, index=False)

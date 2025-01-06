import os
import PyPDF2
import pandas as pd

def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_data(text):
    sections = []
    fir_nos = []
    road_names = []
    injury_data = []
    
    section_start_marker = "1988Section"
    section_end_marker = "State Rule"
    
    fir_start_marker = "FIR/CSR Number   : "
    fir_end_marker = "FIR Date & Time"
    
    road_start_marker = "Street Name"
    road_end_marker = "Local Body"
    
    injury_start_marker = "Total"
    injury_end_marker = "Number of Animals involved in the"
    
    start_idx = 0
    while True:
        # Extract Section
        section_start_idx = text.find(section_start_marker, start_idx)
        if section_start_idx == -1:
            break
        section_end_idx = text.find(section_end_marker, section_start_idx)
        if section_end_idx == -1:
            break
        section = text[section_start_idx + len(section_start_marker):section_end_idx].strip()
        sections.append(section)
        
        # Extract FIR No.
        fir_start_idx = text.find(fir_start_marker, section_end_idx)
        if fir_start_idx == -1:
            break
        fir_end_idx = text.find(fir_end_marker, fir_start_idx)
        if fir_end_idx == -1:
            break
        fir_no = text[fir_start_idx + len(fir_start_marker):fir_end_idx].strip()
        fir_nos.append(fir_no)
        
        # Extract Road Name
        road_start_idx = text.find(road_start_marker, fir_end_idx)
        if road_start_idx == -1:
            break
        road_end_idx = text.find(road_end_marker, road_start_idx)
        if road_end_idx == -1:
            break
        road_name = text[road_start_idx + len(road_start_marker):road_end_idx].strip()
        road_names.append(road_name)
        
        # Extract Injury Data
        injury_start_idx = text.find(injury_start_marker, road_end_idx)
        if injury_start_idx == -1:
            break
        injury_end_idx = text.find(injury_end_marker, injury_start_idx)
        if injury_end_idx == -1:
            break
        injuries = text[injury_start_idx + len(injury_start_marker):injury_end_idx].strip().split()
        if len(injuries) == 5:
            injury_data.append(injuries)
        
        start_idx = injury_end_idx + len(injury_end_marker)
    
    return sections, fir_nos, road_names, injury_data

def save_to_files(data, output_text_file, output_csv_file):
    with open(output_text_file, 'w') as txt_file:
        for row in data:
            txt_file.write(f"Section: {row['Section']}\n")
            txt_file.write(f"FIR No.: {row['FIR No.']}\n")
            txt_file.write(f"Road Name: {row['Road Name']}\n")
            txt_file.write(f"Number of Fatalities: {row['Number of Fatalities']}\n")
            txt_file.write(f"Number of Grievously Injured persons: {row['Number of Grievously Injured persons']}\n")
            txt_file.write(f"Number of people with Minor Injury: {row['Number of people with Minor Injury']}\n")
            txt_file.write(f"No Injury: {row['No Injury']}\n")
            txt_file.write(f"Total: {row['Total']}\n\n")

    df = pd.DataFrame(data)
    df.to_csv(output_csv_file, index=False)

def main(folder_path):
    output_folder = os.path.join(folder_path, 'Nagpur Fatalities')
    os.makedirs(output_folder, exist_ok=True)
    
    data = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            text = extract_text_from_pdf(pdf_path)
            
            sections, fir_nos, road_names, injury_data = extract_data(text)

            for section, fir_no, road_name, injuries in zip(sections, fir_nos, road_names, injury_data):
                data.append({
                    "Section": section,
                    "FIR No.": fir_no,
                    "Road Name": road_name,
                    "Number of Fatalities": injuries[0],
                    "Number of Grievously Injured persons": injuries[1],
                    "Number of people with Minor Injury": injuries[2],
                    "No Injury": injuries[3],
                    "Total": injuries[4]
                })
    
    output_text_file = os.path.join(output_folder, 'extracted_sections.txt')
    output_csv_file = os.path.join(output_folder, 'extracted_sections.csv')
    
    save_to_files(data, output_text_file, output_csv_file)
    print(f"Extraction completed. Files saved in {output_folder}")

folder_path = "/Users/sukanyasaha/Desktop/PDFS2"
main(folder_path)

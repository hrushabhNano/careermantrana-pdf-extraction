import pandas as pd
import re
from pdf2image import convert_from_path
import pytesseract
import os
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('extraction.log'),
        logging.StreamHandler()
    ]
)

def pdf_to_ocr(pdf_path, output_text_file):
    logging.info(f"Starting OCR conversion for PDF: {pdf_path}")
    try:
        images = convert_from_path(pdf_path)
        logging.info(f"Converted PDF to {len(images)} images")
        full_text = ""
        for i, image in enumerate(images):
            logging.info(f"Performing OCR on page {i+1}")
            text = pytesseract.image_to_string(image)
            full_text += f"<PAGE{i+1}>\n<CONTENT_FROM_OCR>\n{text}\n</CONTENT_FROM_OCR>\n</PAGE{i+1}>\n"
        with open(output_text_file, 'w', encoding='utf-8') as f:
            f.write(full_text)
        logging.info(f"Raw OCR text saved to {output_text_file}")
        return full_text
    except Exception as e:
        logging.error(f"Error during OCR: {str(e)}")
        raise

def clean_ocr_text(text):
    logging.info("Starting OCR text cleanup")
    header_pattern = r'Government of Maharashtra\s+State Common Entrance Test Cell\s+Cut Off List for Maharashtra & Minority Seats of CAP Round \| for Admission to First Year of Four Year\s+Degree Courses In Engineering and Technology & Master of Engineering and Technology \(Integrated 5 Years\) for the Year 2023-24\s*'
    footer_pattern = r'Legends: Starting character G-General, L-Ladies, End character H-Home University, O-Other than Home University,S-State Level, Al- All India Seat\.\s+Maharashtra State Seats - Cut Off Indicates Maharashtra State General Merit No\.; Figures in bracket Indicates Merit Percentile\.\s*'
    pages = text.split('<PAGE')[1:]
    cleaned_pages = []
    for page in pages:
        page_content = page.split('<CONTENT_FROM_OCR>')[1].split('</CONTENT_FROM_OCR>')[0]
        cleaned_content = re.sub(header_pattern, '', page_content, flags=re.DOTALL)
        cleaned_content = re.sub(footer_pattern, '', cleaned_content, flags=re.DOTALL)
        cleaned_content = '\n'.join(line.strip() for line in cleaned_content.splitlines() if line.strip())
        cleaned_pages.append(f"<PAGE{page.split('>')[0]}>\n<CONTENT_FROM_OCR>\n{cleaned_content}\n</CONTENT_FROM_OCR>")
    cleaned_text = '\n'.join(cleaned_pages)
    with open('cleaned_ocr_output.txt', 'w', encoding='utf-8') as f:
        f.write(cleaned_text)
    logging.info(f"Cleaned OCR text saved to 'cleaned_ocr_output.txt'")
    return cleaned_text

def normalize_seat_type(seat_type):
    """Normalize seat types to correct OCR errors and standardize format."""
    seat_type = seat_type.replace(':', '').upper()  # Remove colons and convert to uppercase
    # Known OCR corrections
    corrections = {
        'EWWS': 'EWS',
        'GSCS': 'GSCS',  # Already correct, but ensures consistency
        # Add more corrections as needed based on observed OCR errors
    }
    return corrections.get(seat_type, seat_type)

def extract_data_to_excel(text, output_file):
    logging.info("Starting data extraction from cleaned OCR text")
    columns = ['Sr', 'District', 'Home University', 'College Code', 'Institute Name', 
               'Branch Code', 'Branch Name', 'Seat Type', 'Rank', 'Percentile']
    data = []
    sr_no = 1
    pages = text.split('<PAGE')[1:]
    logging.info(f"Found {len(pages)} pages in cleaned OCR text")

    college_pattern = r'(\d{4}) - (.+?),\s*([^,\n]+?)$'
    branch_pattern = r'(\d{7}) - (.+?)$'  # Capture full 7-digit branch code
    status_pattern = r'Status: (.+?)$'
    section_pattern = r'(Home University Seats Allotted to Home University Candidates|Other Than Home University Seats Allotted to Other Than Home University Candidates|Home University Seats Allotted to Other Than Home University Candidates|Other Than Home University Seats Allotted to Home University Candidates|State Level)'
    seat_type_pattern = r'Stage\s+(.+?)$'
    rank_pattern = r'^\s*[iI|W]\s+([\d\s,]+)$'
    percentile_pattern = r'^\s*\(([\d.\s\(\)]+)\)$'

    for page in pages:
        page_content = page.split('<CONTENT_FROM_OCR>')[1].split('</CONTENT_FROM_OCR>')[0]
        logging.info(f"Processing page: {page.split('>')[0]}")
        
        college_match = re.search(college_pattern, page_content, re.MULTILINE)
        if college_match:
            college_code = college_match.group(1)
            institute_name = college_match.group(2).strip()
            district = college_match.group(3).strip()
            logging.info(f"Extracted college: {college_code} - {institute_name}, {district}")
        else:
            logging.warning("No college details found in page")
            continue

        status_match = re.search(status_pattern, page_content, re.MULTILINE)
        home_university = status_match.group(1) if status_match else ''
        logging.info(f"Home University: {home_university}")

        lines = page_content.split('\n')
        current_branch_code = None
        current_branch_name = None
        current_section = None

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            branch_match = re.search(branch_pattern, line)
            if branch_match:
                current_branch_code = branch_match.group(1)  # Full 7-digit code
                current_branch_name = branch_match.group(2)
                logging.info(f"Extracted branch: {current_branch_code} - {current_branch_name}")
                if len(current_branch_code) != 7:
                    logging.warning(f"Branch code {current_branch_code} is not 7 digits, expected length 7")
                i += 1
                continue

            section_match = re.search(section_pattern, line)
            if section_match:
                current_section = section_match.group(1)
                logging.info(f"Section: {current_section}")
                i += 1
                continue

            if line.startswith('Stage'):
                seat_types_match = re.search(seat_type_pattern, line)
                if seat_types_match:
                    seat_types = [normalize_seat_type(st) for st in seat_types_match.group(1).split()]
                    logging.info(f"Normalized seat types: {seat_types}")
                    
                    i += 1
                    if i < len(lines):
                        rank_line = lines[i].strip()
                        rank_match = re.search(rank_pattern, rank_line)
                        if rank_match:
                            ranks = rank_match.group(1).replace(',', '').split()
                            logging.info(f"Ranks: {ranks}")
                            
                            i += 1
                            if i < len(lines):
                                percentile_line = lines[i].strip()
                                percentile_match = re.search(percentile_pattern, percentile_line)
                                if percentile_match:
                                    percentiles = percentile_match.group(1).split(') (')
                                    percentiles = [p.strip('()') for p in percentiles]
                                    logging.info(f"Percentiles: {percentiles}")
                                    
                                    for j, seat_type in enumerate(seat_types):
                                        if j < len(ranks) and j < len(percentiles):
                                            rank = ranks[j]
                                            percentile = percentiles[j]
                                            data.append([sr_no, district, home_university, college_code, institute_name, 
                                                         current_branch_code, current_branch_name, seat_type, rank, percentile])
                                            logging.info(f"Added row: Sr {sr_no}, Seat Type {seat_type}, Rank {rank}, Percentile {percentile}, Branch Code {current_branch_code}")
                                            sr_no += 1
            i += 1

    df = pd.DataFrame(data, columns=columns)
    logging.info(f"Created DataFrame with {len(df)} rows")
    df.to_excel(output_file, index=False)
    logging.info(f"Data extracted and saved to {output_file}")

def main():
    pdf_path = 'Engg_cap_1_trimmed.pdf'  # Update with your PDF path
    raw_ocr_text_file = 'raw_ocr_output.txt'
    output_excel_file = 'cut_off_list_2023_24.xlsx'
    
    logging.info("Starting script execution")
    if not os.path.exists(pdf_path):
        logging.error(f"PDF file '{pdf_path}' not found")
        return
    
    ocr_text = pdf_to_ocr(pdf_path, raw_ocr_text_file)
    cleaned_text = clean_ocr_text(ocr_text)
    extract_data_to_excel(cleaned_text, output_excel_file)
    
    logging.info("Script execution completed")

if __name__ == "__main__":
    main()
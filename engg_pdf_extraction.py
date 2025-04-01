import pandas as pd
import re
from pdf2image import convert_from_path
import pytesseract
import os
import logging
import gc  # Garbage collection for memory management

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('extraction.log'),
        logging.StreamHandler()
    ]
)

def pdf_to_ocr(pdf_path, output_text_file, batch_size=10):
    """Convert PDF to OCR text in small batches, saving incrementally."""
    logging.info(f"Starting OCR conversion for PDF: {pdf_path}")
    try:
        # Get total pages without loading all into memory
        from pdf2image import pdfinfo_from_path
        pdf_info = pdfinfo_from_path(pdf_path)
        total_pages = pdf_info["Pages"]
        logging.info(f"PDF has {total_pages} pages")
        
        # Clear file if it exists
        if os.path.exists(output_text_file):
            os.remove(output_text_file)
        
        for start in range(0, total_pages, batch_size):
            end = min(start + batch_size, total_pages)
            logging.info(f"Processing OCR batch: pages {start+1} to {end}")
            
            # Convert only the current batch of pages
            images = convert_from_path(pdf_path, first_page=start+1, last_page=end)
            batch_text = ""
            
            for i, image in enumerate(images):
                page_num = start + i + 1
                logging.info(f"Performing OCR on page {page_num}")
                text = pytesseract.image_to_string(image)
                batch_text += f"<PAGE{page_num}>\n<CONTENT_FROM_OCR>\n{text}\n</CONTENT_FROM_OCR>\n</PAGE{page_num}>\n"
                # Free image memory immediately
                del image
            
            # Append batch to file
            with open(output_text_file, 'a', encoding='utf-8') as f:
                f.write(batch_text)
            logging.info(f"Batch saved to {output_text_file} (pages {start+1}-{end})")
            
            # Clear memory
            del images
            del batch_text
            gc.collect()
        
        logging.info(f"Raw OCR text fully saved to {output_text_file}")
        
        # Read the full text from file instead of keeping in memory
        with open(output_text_file, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        logging.error(f"Error during OCR: {str(e)}")
        raise

def clean_ocr_text(text, batch_size=10):
    """Clean OCR text in batches."""
    logging.info("Starting OCR text cleanup")
    header_pattern = r'Government of Maharashtra\s+State Common Entrance Test Cell\s+Cut Off List for Maharashtra & Minority Seats of CAP Round \| for Admission to First Year of Four Year\s+Degree Courses In Engineering and Technology & Master of Engineering and Technology \(Integrated 5 Years\) for the Year 2023-24\s*'
    footer_pattern = r'Legends: Starting character G-General, L-Ladies, End character H-Home University, O-Other than Home University,S-State Level, Al- All India Seat\.\s+Maharashtra State Seats - Cut Off Indicates Maharashtra State General Merit No\.; Figures in bracket Indicates Merit Percentile\.\s*'
    pages = text.split('<PAGE')[1:]
    total_pages = len(pages)
    cleaned_text = ""
    
    # Clear previous cleaned file
    cleaned_file = 'cleaned_ocr_output.txt'
    if os.path.exists(cleaned_file):
        os.remove(cleaned_file)
    
    for start in range(0, total_pages, batch_size):
        end = min(start + batch_size, total_pages)
        logging.info(f"Cleaning batch: pages {start+1} to {end}")
        batch_cleaned = ""
        
        for page_idx in range(start, end):
            page = pages[page_idx]
            page_content = page.split('<CONTENT_FROM_OCR>')[1].split('</CONTENT_FROM_OCR>')[0]
            cleaned_content = re.sub(header_pattern, '', page_content, flags=re.DOTALL)
            cleaned_content = re.sub(footer_pattern, '', cleaned_content, flags=re.DOTALL)
            cleaned_content = '\n'.join(line.strip() for line in cleaned_content.splitlines() if line.strip())
            batch_cleaned += f"<PAGE{page.split('>')[0]}>\n<CONTENT_FROM_OCR>\n{cleaned_content}\n</CONTENT_FROM_OCR>\n"
        
        # Append to file
        with open(cleaned_file, 'a', encoding='utf-8') as f:
            f.write(batch_cleaned)
        logging.info(f"Batch cleaned and appended to {cleaned_file} (pages {start+1}-{end})")
        
        # Clear memory
        del batch_cleaned
        gc.collect()
    
    # Read cleaned text from file
    with open(cleaned_file, 'r', encoding='utf-8') as f:
        cleaned_text = f.read()
    logging.info(f"Cleaned OCR text fully saved to {cleaned_file}")
    return cleaned_text

def normalize_seat_type(seat_type):
    """Normalize seat types to correct OCR errors and standardize format."""
    seat_type = seat_type.replace(':', '').upper()
    corrections = {
        'EWWS': 'EWS',
    }
    return corrections.get(seat_type, seat_type)

def extract_data_to_excel(text, output_file, batch_size=10):
    """Extract data in batches and save incrementally."""
    logging.info("Starting data extraction from cleaned OCR text")
    columns = ['Sr', 'District', 'Institute Status', 'College Code', 'Institute Name', 
               'Branch Code', 'Branch Name', 'Seat Type', 'Rank', 'Percentile']
    data = []
    sr_no = 1
    pages = text.split('<PAGE')[1:]
    total_pages = len(pages)
    logging.info(f"Found {total_pages} pages in cleaned OCR text")

    college_pattern = r'(\d{4}) - (.+?),\s*([^,\n]+?)$'
    branch_pattern = r'(\d{9}) - (.+?)$'
    status_pattern = r'Status: (.+?)$'
    section_pattern = r'(Home University Seats Allotted to Home University Candidates|Other Than Home University Seats Allotted to Other Than Home University Candidates|Home University Seats Allotted to Other Than Home University Candidates|Other Than Home University Seats Allotted to Home University Candidates|State Level)'
    seat_type_pattern = r'Stage\s+(.+?)$'
    rank_pattern = r'^\s*[iI|W]\s+([\d\s,]+)$'
    percentile_pattern = r'^\s*\(([\d.\s\(\)]+)\)$'

    # Clear previous batch files
    base_name = output_file.rsplit('.', 1)[0]
    for f in os.listdir('.'):
        if f.startswith(base_name + '_batch_'):
            os.remove(f)

    for start in range(0, total_pages, batch_size):
        end = min(start + batch_size, total_pages)
        logging.info(f"Processing extraction batch: pages {start+1} to {end}")
        batch_data = []
        
        for page_idx in range(start, end):
            page = pages[page_idx]
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
            institute_status = status_match.group(1) if status_match else ''
            logging.info(f"Institute Status: {institute_status}")

            lines = page_content.split('\n')
            current_branch_code = None
            current_branch_name = None
            current_section = None

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                branch_match = re.search(branch_pattern, line)
                if branch_match:
                    current_branch_code = branch_match.group(1)
                    current_branch_name = branch_match.group(2)
                    logging.info(f"Extracted branch: {current_branch_code} - {current_branch_name}")
                    if len(current_branch_code) != 9:
                        logging.warning(f"Branch code {current_branch_code} is not 9 digits, expected length 9")
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
                                                batch_data.append([sr_no, district, institute_status, college_code, institute_name, 
                                                                  current_branch_code, current_branch_name, seat_type, rank, percentile])
                                                logging.info(f"Added row: Sr {sr_no}, Seat Type {seat_type}, Rank {rank}, Percentile {percentile}, Branch Code {current_branch_code}")
                                                sr_no += 1
                i += 1

        # Append batch data and save
        data.extend(batch_data)
        df = pd.DataFrame(data, columns=columns)
        batch_output_file = f"{output_file.rsplit('.', 1)[0]}_batch_{start+1}-{end}.xlsx"
        df.to_excel(batch_output_file, index=False)
        logging.info(f"Batch data saved to {batch_output_file} (pages {start+1}-{end}, {len(batch_data)} rows)")
        
        # Clear memory
        del batch_data
        gc.collect()

    # Final save
    df = pd.DataFrame(data, columns=columns)
    logging.info(f"Created final DataFrame with {len(df)} rows")
    df.to_excel(output_file, index=False)
    logging.info(f"Data extracted and saved to {output_file}")

def main():
    pdf_path = '2023ENGG_CAP1_CutOff.pdf'  # Update with your large PDF path
    raw_ocr_text_file = 'raw_ocr_output.txt'
    output_excel_file = 'cut_off_list_2023_24.xlsx'
    batch_size = 10  # Reduced to 10 pages per batch
    
    logging.info("Starting script execution")
    if not os.path.exists(pdf_path):
        logging.error(f"PDF file '{pdf_path}' not found")
        return
    
    ocr_text = pdf_to_ocr(pdf_path, raw_ocr_text_file, batch_size)
    cleaned_text = clean_ocr_text(ocr_text, batch_size)
    extract_data_to_excel(cleaned_text, output_excel_file, batch_size)
    
    logging.info("Script execution completed")

if __name__ == "__main__":
    main()
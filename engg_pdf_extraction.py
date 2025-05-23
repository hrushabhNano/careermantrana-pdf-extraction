import streamlit as st
import pandas as pd
import re
from pdf2image import convert_from_path, pdfinfo_from_path
import pytesseract
import os
import logging
import gc
import io
import requests

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)

# Set page config
st.set_page_config(
    page_title="CareerMantrana: MHT-CET PDF Extraction tool",
    layout="wide",
    initial_sidebar_state="collapsed",
    page_icon="📄"
)

# Custom CSS to style all buttons, including download button
st.markdown(
    """
    <style>
    /* Style for Process PDF button */
    div.stButton > button {
        background-color: #8e24aa;
        color: white !important;
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
    }
    div.stButton > button:hover {
        background-color: #6b1a82;
        color: white !important;
    }
    div.stButton > button:active {
        color: white !important;
    }
    /* Style for Browse Files button (file uploader) */
    div.stFileUploader button {
        background-color: #8e24aa;
        color: white !important;
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
    }
    div.stFileUploader button:hover {
        background-color: #6b1a82;
        color: white !important;
    }
    div.stFileUploader button:active {
        color: white !important;
    }
    /* Style for Download Cut-off Excel button (download button) */
    div.stDownloadButton > button {
        background-color: #8e24aa;
        color: white !important;
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
    }
    div.stDownloadButton > button:hover {
        background-color: #6b1a82;
        color: white !important;
    }
    div.stDownloadButton > button:active {
        color: white !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

def pdf_to_ocr(pdf_path, output_text_file, batch_size=10, dpi=200):
    logging.info(f"Starting OCR conversion for PDF: {pdf_path}")
    try:
        pdf_info = pdfinfo_from_path(pdf_path)
        total_pages = pdf_info["Pages"]
        logging.info(f"PDF has {total_pages} pages")
        
        if os.path.exists(output_text_file):
            os.remove(output_text_file)
        
        for start in range(0, total_pages, batch_size):
            end = min(start + batch_size, total_pages)
            logging.info(f"Processing OCR batch: pages {start+1} to {end}")
            images = convert_from_path(pdf_path, dpi=dpi, first_page=start+1, last_page=end)
            batch_text = ""
            
            for i, image in enumerate(images):
                page_num = start + i + 1
                logging.info(f"Performing OCR on page {page_num}")
                text = pytesseract.image_to_string(image)
                batch_text += f"<PAGE{page_num}>\n<CONTENT_FROM_OCR>\n{text}\n</CONTENT_FROM_OCR>\n</PAGE{page_num}>\n"
                del image
            
            with open(output_text_file, 'a', encoding='utf-8') as f:
                f.write(batch_text)
            logging.info(f"Batch saved to {output_text_file} (pages {start+1}-{end})")
            
            del images
            del batch_text
            gc.collect()
        
        logging.info(f"Raw OCR text fully saved to {output_text_file}")
        with open(output_text_file, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        logging.error(f"Error during OCR: {str(e)}")
        raise

def clean_ocr_text(text, batch_size=10):
    logging.info("Starting OCR text cleanup")
    header_pattern = r'Government of Maharashtra\s+State Common Entrance Test Cell\s+Cut Off List for Maharashtra & Minority Seats of CAP Round \| for Admission to First Year of Four Year\s+Degree Courses In Engineering and Technology & Master of Engineering and Technology \(Integrated 5 Years\) for the Year 2023-24\s*'
    footer_pattern = r'Legends: Starting character G-General, L-Ladies, End character H-Home University, O-Other than Home University,S-State Level, Al- All India Seat\.\s+Maharashtra State Seats - Cut Off Indicates Maharashtra State General Merit No\.; Figures in bracket Indicates Merit Percentile\.\s*'
    pages = text.split('<PAGE')[1:]
    total_pages = len(pages)
    cleaned_file = 'cleaned_ocr_output.txt'
    
    logging.info(f"Total pages to clean: {total_pages}")
    if os.path.exists(cleaned_file):
        logging.info(f"Removing existing {cleaned_file}")
        os.remove(cleaned_file)
    
    for start in range(0, total_pages, batch_size):
        end = min(start + batch_size, total_pages)
        logging.info(f"Cleaning batch: pages {start+1} to {end}")
        batch_cleaned = ""
        
        for page_idx in range(start, end):
            page = pages[page_idx]
            try:
                page_content = page.split('<CONTENT_FROM_OCR>')[1].split('</CONTENT_FROM_OCR>')[0]
            except IndexError:
                logging.warning(f"Page {page_idx + 1} has malformed OCR content: {page[:100]}...")
                continue
            cleaned_content = re.sub(header_pattern, '', page_content, flags=re.DOTALL)
            cleaned_content = re.sub(footer_pattern, '', cleaned_content, flags=re.DOTALL)
            cleaned_content = '\n'.join(line.strip() for line in cleaned_content.splitlines() if line.strip())
            batch_cleaned += f"<PAGE{page.split('>')[0]}>\n<CONTENT_FROM_OCR>\n{cleaned_content}\n</CONTENT_FROM_OCR>\n"
        
        logging.info(f"Writing batch to {cleaned_file} (pages {start+1}-{end})")
        try:
            full_path = os.path.abspath(cleaned_file)
            logging.info(f"Absolute path for {cleaned_file}: {full_path}")
            with open(cleaned_file, 'a', encoding='utf-8') as f:
                f.write(batch_cleaned)
            logging.info(f"Successfully wrote batch to {cleaned_file}")
        except Exception as e:
            logging.error(f"Failed to write to {cleaned_file}: {str(e)}")
        
        del batch_cleaned
        gc.collect()
    
    if not os.path.exists(cleaned_file):
        logging.error(f"{cleaned_file} was not created after processing")
        raise FileNotFoundError(f"{cleaned_file} was not created")
    
    logging.info(f"Cleaned OCR text fully saved to {cleaned_file}")
    with open(cleaned_file, 'r', encoding='utf-8') as f:
        return f.read()

def normalize_seat_type(seat_type):
    seat_type = seat_type.replace(':', '').upper()
    corrections = {
        'EWWS': 'EWS',
        'NT10': 'NT1O',
        'GNT10': 'GNT1O',
        'NT20': 'NT2O',
        'GNT20': 'GNT2O',
        'NT30': 'NT3O',
        'GNT30': 'GNT3O',
        'LVJSS': 'LVJS',
        'GNT30,': 'GNT3O',
        'LNT10': 'LNT1O'
    }
    corrected_seat_type = corrections.get(seat_type, seat_type)
    if re.match(r'^[GL]?[A-Z]{1,4}0$', corrected_seat_type):
        corrected_seat_type = corrected_seat_type[:-1] + 'O'
    return corrected_seat_type

def extract_data_to_excel(text, log_container, batch_size=10):
    logging.info("Starting data extraction from cleaned OCR text")
    columns = ['Sr', 'Stage', 'District', 'Institute Status', 'College Code', 'Institute Name', 
               'Branch Code', 'Branch Name', 'Seat Type', 'Rank', 'Percentile']
    data = []
    sr_no = 1
    pages = text.split('<PAGE')[1:]
    total_pages = len(pages)
    logging.info(f"Found {total_pages} pages in cleaned OCR text")

    college_pattern = r'(\d{4}) - (.+?)(?:,\s*([^,\n]+?))?$'
    branch_pattern = r'(\d{9}) - (.+?)$'
    status_pattern = r'Status: (.+?)$'
    section_pattern = r'(Home University Seats Allotted to Home University Candidates|Other Than Home University Seats Allotted to Other Than Home University Candidates|Home University Seats Allotted to Other Than Home University Candidates|Other Than Home University Seats Allotted to Home University Candidates|State Level)'
    seat_type_pattern = r'Stage\s+(.+?)$'
    rank_pattern = r'^\s*[iI1][l}l\s]*(.+)$'  # Handles "i", "I", "1", "il}", "l}"
    percentile_pattern = r'^\s*\(([\d.\s\(\)]+)\)$'

    # OCR correction dictionary for ranks
    ocr_corrections = {
        '2m': '201',
        'S77': '577',
        'M6': '346',
        'l}': '',  # Handle "l}" as noise
        'il}': ''  # Handle "il}" as noise
    }

    progress_bar = st.progress(0)
    status_text = st.empty()
    
    def normalize_seat_type(seat_type):
        seat_type = re.sub(r'[:;,\.]+$', '', seat_type.strip()).upper()
        corrections = {
            'EWWS': 'EWS',
            'NT10': 'NT1O',
            'GNT10': 'GNT1O',
            'NT20': 'NT2O',
            'GNT20': 'GNT2O',
            'NT30': 'NT3O',
            'GNT30': 'GNT3O',
            'LVJSS': 'LVJS'
        }
        if seat_type in corrections:
            return corrections[seat_type]
        if re.match(r'^[GL]?NT\d0$', seat_type):
            return seat_type[:-1] + 'O'
        return seat_type.strip()

    for start in range(0, total_pages, batch_size):
        end = min(start + batch_size, total_pages)
        logging.info(f"Processing extraction batch: pages {start+1} to {end}")
        batch_data = []
        
        for page_idx in range(start, end):
            page = pages[page_idx]
            page_content = page.split('<CONTENT_FROM_OCR>')[1].split('</CONTENT_FROM_OCR>')[0]
            logging.info(f"Processing page {page_idx + 1}: {page.split('>')[0]}")
            
            college_match = re.search(college_pattern, page_content, re.MULTILINE)
            if college_match:
                college_code = college_match.group(1)
                institute_name = college_match.group(2).strip()
                district = college_match.group(3).strip() if college_match.group(3) else "Unknown"
                logging.info(f"Extracted college: {college_code} - {institute_name}, {district}")
            else:
                logging.warning(f"No college details found in page {page_idx + 1}")
                continue

            status_match = re.search(status_pattern, page_content, re.MULTILINE)
            institute_status = status_match.group(1) if status_match else ''
            logging.info(f"Institute Status: {institute_status}")

            lines = page_content.split('\n')
            current_branch_code = None
            current_branch_name = None
            current_section = None
            base_seat_types = None
            seat_types = None
            ranks = None
            percentiles = None
            current_stage = 1

            def add_rows():
                nonlocal sr_no, batch_data
                if seat_types and ranks and current_branch_code:
                    for j, seat_type in enumerate(seat_types):
                        if j < len(ranks):
                            rank = ranks[j]
                            percentile = percentiles[j] if percentiles and j < len(percentiles) else None
                            batch_data.append([sr_no, current_stage, district, institute_status, college_code, institute_name, 
                                              current_branch_code, current_branch_name, seat_type, rank, percentile])
                            logging.info(f"Added row: Sr {sr_no}, Stage {current_stage}, Seat Type {seat_type}, Rank {rank}, Percentile {percentile}")
                            sr_no += 1
                else:
                    logging.warning(f"Skipping row addition: Missing data - Seat Types: {seat_types}, Ranks: {ranks}, Branch Code: {current_branch_code}")

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                branch_match = re.search(branch_pattern, line)
                if branch_match:
                    add_rows()
                    current_branch_code = branch_match.group(1)
                    current_branch_name = branch_match.group(2)
                    logging.info(f"Extracted branch: {current_branch_code} - {current_branch_name}")
                    base_seat_types = None
                    seat_types = None
                    ranks = None
                    percentiles = None
                    current_stage = 1
                    i += 1
                    continue

                section_match = re.search(section_pattern, line)
                if section_match:
                    add_rows()
                    current_section = section_match.group(1)
                    logging.info(f"Section: {current_section}")
                    base_seat_types = None
                    seat_types = None
                    ranks = None
                    percentiles = None
                    current_stage = 1
                    i += 1
                    continue

                seat_type_match = re.search(seat_type_pattern, line)
                if seat_type_match:
                    add_rows()
                    base_seat_types = [normalize_seat_type(st) for st in seat_type_match.group(1).split()]
                    seat_types = base_seat_types.copy()
                    logging.info(f"Normalized base seat types: {base_seat_types}")
                    current_stage = 1
                    i += 1
                    continue

                rank_match = re.search(rank_pattern, line)
                if rank_match:
                    if ranks and seat_types and current_branch_code:
                        add_rows()
                        current_stage += 1
                    # Use base_seat_types from Stage 1 if available, otherwise backtrack
                    if not base_seat_types and current_section:
                        prev_line_idx = i - 1
                        while prev_line_idx >= 0:
                            prev_line = lines[prev_line_idx].strip()
                            if prev_line and not re.search(section_pattern, prev_line) and not re.search(percentile_pattern, prev_line):
                                if re.search(seat_type_pattern, prev_line):
                                    base_seat_types = [normalize_seat_type(st) for st in re.search(seat_type_pattern, prev_line).group(1).split()]
                                    logging.info(f"Backtracked to seat types from Stage line: {base_seat_types}")
                                    break
                                prev_line_idx -= 1
                    # Extract ranks
                    rank_str = rank_match.group(1).strip()
                    rank_tokens = rank_str.split()
                    ranks = []
                    for token in rank_tokens:
                        if token == '}':  # Skip standalone "}" from "i}"
                            continue
                        corrected_token = ocr_corrections.get(token, token)
                        if corrected_token == '':  # Handle "l}" or "il}" by skipping to next token
                            continue
                        numbers = re.findall(r'\d+', corrected_token)
                        if numbers:
                            ranks.append(numbers[0])
                        else:
                            logging.warning(f"Could not extract number from rank token: {token}")
                            ranks.append(corrected_token)
                    if not ranks:
                        logging.warning(f"Failed to parse ranks from line: {line}")
                        ranks = None
                    else:
                        logging.info(f"Ranks after correction: {ranks}")
                        # Set seat_types for this stage, default to base_seat_types if not already set
                        if base_seat_types:
                            seat_types = base_seat_types[:len(ranks)]
                            logging.info(f"Adjusted seat types for Stage {current_stage}: {seat_types}")
                        else:
                            logging.warning(f"No base_seat_types available for Stage {current_stage}, skipping row addition")
                    i += 1
                    continue

                percentile_match = re.search(percentile_pattern, line)
                if percentile_match:
                    percentiles = percentile_match.group(1).split(') (')
                    percentiles = [p.strip('()') for p in percentiles]
                    logging.info(f"Percentiles: {percentiles}")
                    i += 1
                    continue

                i += 1

            add_rows()

        if batch_data:
            data.extend(batch_data)
            logging.info(f"Batch data added: {len(batch_data)} rows")
        
        progress = min((start + batch_size) / total_pages, 1.0)
        progress_bar.progress(progress)
        status_text.text(f"Processing batch: pages {start+1} to {end} ({len(batch_data)} rows extracted)")
        
        log_container.text_area("Processing Logs", value=st.session_state.logs, height=300, key=f"log_area_{start}")
        
        del batch_data
        gc.collect()

    logging.info(f"Total rows in data: {len(data)}")
    if not data:
        logging.error("No data extracted from the PDF")

    df = pd.DataFrame(data, columns=columns)
    logging.info(f"Created final DataFrame with {len(df)} rows")
    
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    progress_bar.progress(1.0)
    status_text.text("Processing complete!")
    return output

def main():
    # Create two columns: 70% left, 30% right
    col1, col2 = st.columns([7, 3])

    # Left column (70%) for functional components
    with col1:
        # Fetch and display the logo at the top
        logo_url = "https://www.careermantrana.com/images/mainLogo.svg"
        try:
            response = requests.get(logo_url)
            if response.status_code == 200:
                st.markdown(
                    f'<img src="{logo_url}" alt="Career Mantra Logo" style="max-width: 300px; display: block; margin: 0 auto;">',
                    unsafe_allow_html=True
                )
            else:
                st.warning("Could not load logo from URL.")
        except Exception as e:
            st.warning(f"Error loading logo: {str(e)}")

        st.title("CareerMantrana: MHT-CET PDF Extraction tool")
        st.write("Upload a PDF file to extract cut-off data into an Excel file.")

        # Initialize session state
        if 'logs' not in st.session_state:
            st.session_state.logs = ""
        if 'processing_complete' not in st.session_state:
            st.session_state.processing_complete = False
        
        uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
        
        if uploaded_file is not None:
            pdf_path = "temp_uploaded.pdf"
            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            log_container = st.empty()
            
            class StreamlitLogHandler(logging.Handler):
                def emit(self, record):
                    log_entry = self.format(record)
                    st.session_state.logs += log_entry + "\n"
            
            log_handler = StreamlitLogHandler()
            logging.getLogger().addHandler(log_handler)
            
            raw_ocr_text_file = 'raw_ocr_output.txt'
            output_excel_file = 'cut_off_list_2023_24.xlsx'
            batch_size = 10
            
            if not st.session_state.processing_complete:
                if st.button("Process PDF"):
                    with st.spinner("Processing..."):
                        ocr_text = pdf_to_ocr(pdf_path, raw_ocr_text_file, batch_size)
                        cleaned_text = clean_ocr_text(ocr_text, batch_size)
                        excel_bytes = extract_data_to_excel(cleaned_text, log_container, batch_size)
                        
                        st.session_state.processing_complete = True
                        st.session_state.excel_bytes = excel_bytes
            
            if st.session_state.processing_complete and 'excel_bytes' in st.session_state:
                log_container.text_area("Processing Logs", value=st.session_state.logs, height=300, key="log_area_final")
                if st.session_state.excel_bytes.getvalue():
                    st.download_button(
                        label="Download Cut-off Excel",
                        data=st.session_state.excel_bytes,
                        file_name=output_excel_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Generated Excel file is empty. Check logs for details.")
            
            for file in [pdf_path, raw_ocr_text_file]:
                if os.path.exists(file):
                    os.remove(file)
            
            logging.getLogger().removeHandler(log_handler)

    # Right column (30%) for the image
    with col2:
        image_url = "https://www.careermantrana.com/assets/heroSection1-irHkr2pB.svg"
        try:
            response = requests.get(image_url)
            if response.status_code == 200:
                st.image(image_url, use_column_width=True)
            else:
                st.warning("Could not load image from URL.")
        except Exception as e:
            st.warning(f"Error loading image: {str(e)}")

if __name__ == "__main__":
    main()
import streamlit as st
import pandas as pd
import re
from io import BytesIO
import base64
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
import logging

logging.basicConfig(
    filename='ocr_output.log',
    level=logging.DEBUG,  # Changed to DEBUG to capture debug messages
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def ocr_pdf_to_text(pdf_file, batch_size=5, dpi=200):
    """
    Convert PDF to text using OCR in batches to manage memory and log the output
    """
    pdf_content = pdf_file.read()
    full_text = ""
    
    # Get total page count
    try:
        images = convert_from_bytes(pdf_content, dpi=dpi, size=(None, None))  # Just to get count
        total_pages = len(images)
        del images  # Free memory immediately
    except Exception as e:
        st.error(f"Failed to get page count: {str(e)}")
        logging.error(f"Failed to get page count: {str(e)}")
        return ""
    
    # Process in batches
    for start_page in range(0, total_pages, batch_size):
        end_page = min(start_page + batch_size, total_pages)
        st.write(f"Processing pages {start_page + 1} to {end_page} of {total_pages}...")
        
        try:
            images = convert_from_bytes(pdf_content, dpi=dpi, first_page=start_page + 1, last_page=end_page)
            for i, image in enumerate(images):
                page_num = start_page + i + 1
                text = pytesseract.image_to_string(image)
                full_text += f"\n--- Page {page_num} ---\n{text}"
                # Log the OCR output for this page
                logging.info(f"Page {page_num} OCR Output:\n{text}")
                st.write(f"OCR extracted text from page {page_num}:\n{text[:500]}...")
                del image  # Free memory
            del images  # Free batch memory
        except Exception as e:
            st.error(f"Error processing pages {start_page + 1} to {end_page}: {str(e)}")
            logging.error(f"Error processing pages {start_page + 1} to {end_page}: {str(e)}")
            continue
    
    return full_text

def extract_data_from_text(text):
    """
    Extract cutoff data from OCR'd text with corrected district, rank, and percentile extraction
    """
    extracted_data = []
    
    # Extract district from institute name (more reliable than header)
    college_pattern = r'(\d{4,5})\s*[-–]\s*(.*?)(?=\n\d{4,5}\s*[-–]|\Z)'
    college_matches = re.finditer(college_pattern, text, re.DOTALL)
    
    for college_match in college_matches:
        college_code = college_match.group(1)
        college_block = college_match.group(2).strip()
        
        # Extract institute name and district
        institute_name_pattern = r'^(.*?)(?=\n\d{6,}\s*[-–]|\nStatus|\Z)'
        institute_match = re.search(institute_name_pattern, college_block, re.MULTILINE | re.DOTALL)
        institute_name = institute_match.group(1).strip() if institute_match else "Unknown"
        
        # Extract district from institute name (e.g., "Amravati" from "Government College of Engineering, Amravati")
        district_pattern = r',\s*(\w+)$|(\w+)\s*(?:University|College|Institute)'
        district_match = re.search(district_pattern, institute_name)
        district = district_match.group(1) or district_match.group(2) if district_match else "Unknown"
        
        # Extract branch and seat data
        branch_pattern = r'(\d{6,})\s*[-–]\s*(.*?)(?=\n\d{6,}\s*[-–]|\Z)'
        branch_matches = re.finditer(branch_pattern, college_block, re.DOTALL)
        
        for branch_match in branch_matches:
            branch_code = branch_match.group(1)
            branch_block = branch_match.group(2).strip()
            branch_name = branch_block.split('\n')[0].strip()
            
            # Extract status for Home University
            status_pattern = r'Status:\s*(.*?)(?=\n|$)'
            status_match = re.search(status_pattern, branch_block)
            status = status_match.group(1).strip() if status_match else "Unknown"
            home_university = "Autonomous Institute" if status_match and "Autonomous" in status_match.group(1) else status if status_match else "Unknown"
            
            # Extract seat data from all sections
            section_pattern = r'(State Level|Home University Seats.*?|Other Than Home University Seats.*?)\s*(?:Stage\s+([A-Z\s:]+?)\s*[\|\s]\s*([\d\s,]+?)(?:\s*(\(.*?)?)?)'
            section_matches = re.finditer(section_pattern, branch_block, re.DOTALL)
            
            for section_match in section_matches:
                seat_types_str = section_match.group(2).strip() if section_match.group(2) else ""
                ranks_str = section_match.group(3).strip() if section_match.group(3) else ""
                percentiles_str = section_match.group(4).strip() if section_match.group(4) else ""
                
                # Clean and split seat types
                seat_types = re.findall(r'[A-Z][A-Z0-9]+', seat_types_str.replace(':', ''))
                
                # Split ranks (ensure proper parsing)
                ranks = [int(rank.replace(',', '')) for rank in ranks_str.split() if rank.replace(',', '').isdigit()]
                
                # Split percentiles (capture full string after ranks)
                percentiles = re.findall(r'\((\d+\.\d+|\d+)\)', percentiles_str)
                percentiles = [float(p) for p in percentiles]
                
                # Debug logging
                logging.debug(f"Section: {section_match.group(1)}, Seat Types: {seat_types}, Ranks: {ranks}, Percentiles: {percentiles}")
                
                # Align lists
                min_length = min(len(seat_types), len(ranks), len(percentiles) if percentiles else len(ranks))
                if min_length == 0:
                    logging.warning(f"No valid data in section: {section_match.group(1)}")
                    continue
                seat_types = seat_types[:min_length]
                ranks = ranks[:min_length]
                percentiles = percentiles[:min_length] if percentiles else [None] * min_length
                
                # Pair seat types with ranks and percentiles
                for seat_type, rank, percentile in zip(seat_types, ranks, percentiles):
                    extracted_data.append({
                        "District": district,
                        "Home University": home_university,
                        "Institute Name": institute_name,
                        "College Code": college_code,
                        "Branch Code": branch_code,
                        "Branch Name": branch_name,
                        "Status": status,
                        "Seat Type": seat_type,
                        "Cutoff (Rank)": rank,
                        "Cutoff (Percentile)": percentile
                    })
    
    if not extracted_data:
        st.warning("No data matched the extraction patterns. Full OCR text:\n" + text[:1000])
        logging.warning("No data extracted. Sample OCR text:\n" + text[:1000])
    
    return extracted_data

def create_excel_file(data):
    """Convert data to Excel format"""
    if not data:
        return None
        
    df = pd.DataFrame(data)
    df.insert(0, "Sr", range(1, len(df) + 1))
    columns_order = ["Sr", "District", "Home University", "Institute Name", "College Code", 
                     "Branch Code", "Branch Name", "Status", "Seat Type", "Cutoff (Rank)", "Cutoff (Percentile)"]
    df = df[columns_order]
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Cutoff Data')
    return output.getvalue()

def get_download_link(file_bytes, filename):
    """Generate download link"""
    b64 = base64.b64encode(file_bytes).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download {filename}</a>'

def process_pdf_file(uploaded_file):
    """Process the uploaded PDF file using OCR"""
    ocr_text = ocr_pdf_to_text(uploaded_file, batch_size=5, dpi=200)
    extracted_data = extract_data_from_text(ocr_text)
    return extracted_data

def main():
    st.title("Students Mantrana: PDF to Excel Converter for B.Tech/B.E cutoffs)")
    st.write("Upload a PDF file containing college cutoff data to convert it to Excel format")
    
    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    
    if uploaded_file is not None:
        st.write("Processing file with OCR...")
        
        try:
            extracted_data = process_pdf_file(uploaded_file)
            
            if extracted_data:
                excel_bytes = create_excel_file(extracted_data)
                
                if excel_bytes:
                    st.markdown(
                        get_download_link(excel_bytes, f"cutoff_data_{uploaded_file.name.split('.')[0]}.xlsx"),
                        unsafe_allow_html=True
                    )
                    st.success("File processed successfully!")
                    
                    st.write("Preview of extracted data:")
                    df = pd.DataFrame(extracted_data)
                    df.insert(0, "Sr", range(1, len(df) + 1))
                    st.dataframe(df.head())
                else:
                    st.error("No data extracted from the PDF")
            else:
                st.error("No data could be extracted from the PDF after OCR")
                
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            logging.error(f"An error occurred in main: {str(e)}")

if __name__ == "__main__":
    main()
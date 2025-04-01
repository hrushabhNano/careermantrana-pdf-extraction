import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
import xlsxwriter
from tqdm import tqdm

# Function to extract data from a single page
def extract_data_from_page(page_text):
    # Define regex patterns for the data lines
    pattern = r'(\d+)\s+(\w+)\s+([\w\s]+)\s+(\d+)\s+(.+?)\s+(\d+)\s+(.+?)\s+([A-Z0-9]+)\s+(\d+)\s+([\d.]+)'
    
    # Find all matches in the page text
    matches = re.findall(pattern, page_text)
    
    # Store extracted data
    data = []
    for match in matches:
        sr, district, home_univ, college_code, institute_name, branch_code, branch_name, seat_type, rank, percentile = match
        data.append({
            "Sr": int(sr),
            "Institute Name": institute_name.strip(),
            "Institute Code": college_code,
            "University Status": "Government Autonomous" if college_code.startswith("1002") else "University Department" if college_code.startswith("1005") else "Government" if college_code.startswith("1012") else "Un-Aided",
            "District": district,
            "Seat Type": seat_type,
            "Cutoff Rank": int(rank),
            "Cutoff Percentile": float(percentile)
        })
    return data

# Function to process PDF in batches
def process_pdf(pdf_file, batch_size=50):
    all_data = []
    
    # Open the PDF file with pdfplumber
    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        st.write(f"Total pages in PDF: {total_pages}")
        
        # Process pages in batches
        for start_page in tqdm(range(0, total_pages, batch_size), desc="Processing batches"):
            end_page = min(start_page + batch_size, total_pages)
            batch_pages = pdf.pages[start_page:end_page]
            
            for page in batch_pages:
                page_text = page.extract_text()
                if page_text:
                    # Extract data from the page, ignoring headers/footers by focusing on data pattern
                    page_data = extract_data_from_page(page_text)
                    all_data.extend(page_data)
    
    return all_data

# Function to convert data to Excel
def convert_to_excel(data):
    # Create a DataFrame
    df = pd.DataFrame(data)
    
    # Remove the 'Sr' column from the DataFrame since we'll use autonumbering in Excel
    if 'Sr' in df.columns:
        df = df.drop(columns=['Sr'])
    
    # Create a BytesIO buffer for the Excel file
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Write DataFrame to Excel
    df.to_excel(writer, index=True, index_label="Sr No", sheet_name="Cutoff Data")
    
    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets["Cutoff Data"]
    
    # Set column widths for better readability
    worksheet.set_column('A:A', 10)  # Sr No
    worksheet.set_column('B:B', 50)  # Institute Name
    worksheet.set_column('C:C', 15)  # Institute Code
    worksheet.set_column('D:D', 25)  # University Status
    worksheet.set_column('E:E', 15)  # District
    worksheet.set_column('F:F', 20)  # Seat Type
    worksheet.set_column('G:G', 15)  # Cutoff Rank
    worksheet.set_column('H:H', 20)  # Cutoff Percentile
    
    # Save the Excel file
    writer.close()
    output.seek(0)
    
    return output

# Streamlit Interface
def main():
    st.title("PDF to Excel Converter for Engineering Cutoffs")
    st.write("Upload a PDF file containing college and branch-wise cutoff ranks & percentiles.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    
    if uploaded_file is not None:
        st.write("Processing the uploaded PDF...")
        
        # Process the PDF in batches
        extracted_data = process_pdf(uploaded_file, batch_size=50)
        
        if extracted_data:
            st.write(f"Extracted data from {len(extracted_data)} entries.")
            
            # Convert to Excel
            excel_file = convert_to_excel(extracted_data)
            
            # Provide download button
            st.download_button(
                label="Download Excel File",
                data=excel_file,
                file_name="Clone_FINAL_2024_FIRST_ROUND.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No valid data found in the PDF. Please check the file format.")

if __name__ == "__main__":
    main()
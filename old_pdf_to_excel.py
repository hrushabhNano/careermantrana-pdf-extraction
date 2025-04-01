import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import re
import xml.etree.ElementTree as ET
import openpyxl
from datetime import datetime

def parse_document_content(content):
    """Parse the XML-like content and extract cutoff data"""
    # Parse XML content
    try:
        root = ET.fromstring(f"<root>{content}</root>")
    except Exception as e:
        st.error(f"Error parsing document: {e}")
        return None

    data = []
    current_institute = ""
    current_code = ""
    current_district = ""

    # First check if there's a table format (like your second document)
    table_pattern = r"(\d+)\s+([^\t]+)\s+([^\t]+)\s+(\d+)\s+([^\t]+)\s+([^\t]+)\s+([^\t]+)\s+([^\t]+)\s+(\d+)\s+([\d.]+)"
    table_matches = re.findall(table_pattern, content)
    
    if table_matches:
        for match in table_matches:
            sr, district, home_uni, code, institute, branch_code, branch, seat_type, rank, percentile = match
            data.append({
                "Institute Name": institute.strip(),
                "Institute Code": code,
                "District": district.strip(),
                "Seat Type": seat_type.strip(),
                "Cutoff (Rank)": int(rank),
                "Cutoff (Percentile)": float(percentile)
            })
    else:
        # Parse the detailed format (like your first document)
        for page in root.findall(".//PAGE*"):
            text = page.find("CONTENT_FROM_OCR").text if page.find("CONTENT_FROM_OCR") is not None else ""
            
            # Extract institute info
            institute_match = re.search(r"(\d{4}) - ([^\n]+)", text)
            if institute_match:
                current_code = institute_match.group(1)
                current_institute = institute_match.group(2).strip()
                # Infer district from institute name (you might need to adjust this logic)
                current_district = "Amravati" if "Amravati" in current_institute else "Yavatmal" if "Yavatmal" in current_institute else "Unknown"

            # Extract cutoff data
            sections = ["Home University Seats", "Other Than Home University Seats", "State Level"]
            for section in sections:
                section_pattern = rf"{section}.*?Stage\s+([^\n]+)(.*?)(?=(?:{'|'.join(sections)}|$))"
                section_matches = re.findall(section_pattern, text, re.DOTALL)
                
                for stage, values in section_matches:
                    lines = values.strip().split('\n')
                    for line in lines:
                        if line.strip() and not line.strip().startswith("Stage"):
                            items = line.split()
                            if len(items) >= 2:
                                seat_type = items[0]
                                rank = int(items[1]) if items[1].isdigit() else None
                                percentile = float(items[2].strip('()')) if len(items) > 2 and items[2].startswith('(') else None
                                
                                if rank:
                                    data.append({
                                        "Institute Name": current_institute,
                                        "Institute Code": current_code,
                                        "District": current_district,
                                        "Seat Type": seat_type,
                                        "Cutoff (Rank)": rank,
                                        "Cutoff (Percentile)": percentile if percentile else 0.0
                                    })

    return data

def convert_to_excel(data, output_file):
    """Convert parsed data to Excel format"""
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Add autonumber column
    df.insert(0, "Sr No", range(1, len(df) + 1))
    
    # Reorder columns
    columns = ["Sr No", "Institute Name", "Institute Code", "District", "Seat Type", "Cutoff (Rank)", "Cutoff (Percentile)"]
    df = df[columns]
    
    # Save to Excel
    df.to_excel(output_file, index=False)
    return df

def main():
    st.title("Engineering Cutoff Converter")
    st.write("Upload a document containing cutoff ranks and percentiles to convert to Excel format")

    # File uploader
    uploaded_file = st.file_uploader("Choose a file", type=['txt', 'xml'])

    if uploaded_file is not None:
        # Read file content
        content = uploaded_file.read().decode('utf-8')
        
        # Process the content
        with st.spinner('Processing file...'):
            parsed_data = parse_document_content(content)
            
            if parsed_data:
                # Generate output filename
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = f"cutoff_output_{timestamp}.xlsx"
                
                # Convert to Excel
                df = convert_to_excel(parsed_data, output_file)
                
                # Display preview
                st.write("Preview of processed data:")
                st.dataframe(df.head())
                
                # Provide download link
                with open(output_file, 'rb') as f:
                    st.download_button(
                        label="Download Excel file",
                        data=f,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("No data could be extracted from the file")

if __name__ == "__main__":
    main()
import PyPDF2
import pandas as pd
import streamlit as st
import re
from io import BytesIO

def extract_cutoff_data(pdf_file):
    """
    Extracts cutoff rank and percentile data from the PDF.

    Args:
        pdf_file (streamlit.UploadedFile): Uploaded PDF file object.

    Returns:
        pandas.DataFrame: DataFrame containing the extracted data, or None if no data is found.
    """

    data = []
    institute_name = None
    institute_code = None
    university_status = None
    district = None  # Not directly available in this PDF structure

    try:
        reader = PyPDF2.PdfReader(pdf_file)
        total_pages = len(reader.pages)

        # Process PDF in chunks (batch processing)
        for page_num in range(total_pages):
            page = reader.pages[page_num]
            text = page.extract_text()
            if not text:
                continue

            lines = text.split('\n')

            for i, line in enumerate(lines):
                # Extract Institute Name and Code
                if i > 0 and " - " in lines[i-1] and "College" in lines[i-1] and "Engineering" in lines[i-1]:
                    try:
                      institute_code, institute_name = lines[i-1].split(" - ", 1)
                    except ValueError:
                      continue
                if "Status:" in line:
                    university_status = line.split("Status:")[1].strip()

                # Extract data rows
                if "Stage" in line and any(x in lines[i+1] for x in ["GOPEN", "GS", "GOBC", "TFWS", "EWS", "DEF", "LOPEN"]):

                  header_line = lines[i+1]
                  header_cols = header_line.split(",")
                  
                  for j in range(i+2, len(lines)):
                    if not any(x in lines[j] for x in ["(", ")"]):
                      continue
                    
                    row_line = lines[j]
                    row_values = row_line.split(",")
                    
                    record = {
                        "Institute Name": institute_name,
                        "Institute Code": institute_code,
                        "University Status": university_status,
                        "District": district,
                        "Seat Type": None,
                        "Cutoff Rank": None,
                        "Cutoff Percentile": None
                    }
                    
                    for k, header in enumerate(header_cols):
                      header = header.strip()
                      if k < len(row_values):
                        value = row_values[k].strip()
                        
                        # Extract rank and percentile
                        match = re.search(r'(\d+)\s*\(([\d.]+)\)', value)
                        if match:
                          record["Cutoff Rank"] = match.group(1)
                          record["Cutoff Percentile"] = match.group(2)
                          record["Seat Type"] = header
                          data.append(record)
                        else:
                          # If no rank/percentile, check for just rank
                          match = re.search(r'(\d+)', value)
                          if match:
                            record["Cutoff Rank"] = match.group(1)
                            record["Seat Type"] = header
                            data.append(record)

        if data:
            df = pd.DataFrame(data)
            df.insert(0, 'Auto Number', range(1, len(df) + 1))  # Add autonumber column
            return df
        else:
            return None

    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None

def main():
    st.title("PDF Data Extractor")

    uploaded_file = st.file_uploader("Upload PDF file", type="pdf")

    if uploaded_file is not None:
        df = extract_cutoff_data(uploaded_file)

        if df is not None:
            st.success("Data extracted successfully!")
            st.dataframe(df)  # Display the DataFrame

            # Create a download link for the Excel file
            excel_file = BytesIO()
            df.to_excel(excel_file, index=False)
            excel_file.seek(0)
            st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data could be extracted from the PDF.")

if __name__ == "__main__":
    main()
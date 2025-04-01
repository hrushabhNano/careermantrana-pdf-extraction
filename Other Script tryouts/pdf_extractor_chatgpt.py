import streamlit as st
import pandas as pd
import os

try:
    import pdfplumber
except ModuleNotFoundError:
    st.error("Module 'pdfplumber' is not installed. Please install it using 'pip install pdfplumber'")
    st.stop()

def extract_data_from_pdf(pdf_path, batch_size=10):
    extracted_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            
            for i in range(0, total_pages, batch_size):
                batch = pdf.pages[i:i+batch_size]
                for page in batch:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for row in table:
                                row = [cell.strip() if cell else "" for cell in row]  # Clean up empty values
                                if len(row) >= 7:  # Ensure minimum required columns
                                    extracted_data.append(row[:7])  # Trim extra columns if necessary
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
        st.stop()
    return extracted_data

def process_data_to_dataframe(data):
    if not data:
        st.error("No valid data extracted from PDF.")
        st.stop()
    
    columns = ["Institute Name", "Institute Code", "Status", "District", "Seat Type", "Cutoff Rank", "Cutoff Percentile"]
    df = pd.DataFrame(data, columns=columns)
    df.insert(0, "ID", range(1, len(df) + 1))  # Auto-number column
    return df

def main():
    st.title("PDF to Excel Converter - Engineering Cutoffs")
    
    uploaded_file = st.file_uploader("Upload the PDF file", type=["pdf"])
    if uploaded_file:
        temp_pdf_path = "temp.pdf"
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.read())
        
        st.write("Processing file...")
        raw_data = extract_data_from_pdf(temp_pdf_path)
        df = process_data_to_dataframe(raw_data)
        
        excel_path = "output.xlsx"
        df.to_excel(excel_path, index=False)
        
        st.success("File processed successfully!")
        st.download_button(label="Download Excel File", data=open(excel_path, "rb"), file_name="CutoffData.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        os.remove(temp_pdf_path)
        os.remove(excel_path)

if __name__ == "__main__":
    main()

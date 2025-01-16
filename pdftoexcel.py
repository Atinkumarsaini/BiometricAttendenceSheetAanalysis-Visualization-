import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime

def convert_date_format(date_str):
    """
    Convert date from 'DD-Mon-YYYY' to 'YYYY-MM-DD' format
    """
    try:
        # Parse the date string
        date_obj = datetime.strptime(date_str, '%d-%b-%Y')
        # Convert to desired format
        return date_obj.strftime('%Y-%m-%d')
    except Exception as e:
        print(f"Date conversion me error: {str(e)}")
        return date_str

def extract_page_attendance_date(page_text):
    """
    Extract attendance date from page text and convert format
    """
    try:
        date_pattern = r'Attendance Date\s+(\d{2}-[A-Za-z]+-\d{4})'
        match = re.search(date_pattern, page_text)
        
        if match:
            date_str = match.group(1)
            return convert_date_format(date_str)
        return None
            
    except Exception as e:
        print(f"Page se date extract karne me error aaya: {str(e)}")
        return None

def remove_header_footer(input_pdf_path, output_pdf_path, header_height=50, footer_height=50):
    """
    Remove header and footer from all pages of a PDF file.
    """
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()
        
        for page in reader.pages:
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)
            
            page.mediabox.upper_right = (
                page_width,
                page_height - header_height
            )
            page.mediabox.lower_left = (
                0,
                footer_height
            )
            
            writer.add_page(page)
        
        with open(output_pdf_path, 'wb') as output_file:
            writer.write(output_file)
            
        print(f"PDF process ho gaya. Save ho gaya yahan: {output_pdf_path}")
        return True
        
    except Exception as e:
        print(f"Error aaya header/footer remove karte time: {str(e)}")
        return False

def convert_pdf_to_excel(pdf_path, excel_path):
    """
    Convert PDF to Excel with attendance status
    """
    try:
        print("PDF ko Excel mein convert kar rahe hain...")
        all_tables = []
        
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                tables = page.extract_tables()
                page_date = extract_page_attendance_date(page_text)
                
                for table in tables:
                    if not table:
                        continue
                        
                    df = pd.DataFrame(table[1:], columns=table[0])
                    
                    # Remove empty rows
                    df = df.dropna(how='all')
                    
                    if page_date:
                        df['Attendance Date'] = page_date
                    
                    
                    all_tables.append(df)
        
        if len(all_tables) > 0:
            final_df = pd.concat(all_tables, ignore_index=True)
            
            # Reorder columns to put date and status at end
            cols = [col for col in final_df.columns if col not in ['Attendance Date']]
            cols = cols + ['Attendance Date']
            final_df = final_df[cols]
            
            # Clean the DataFrame
            final_df = final_df.dropna(how='all')
            final_df = final_df[final_df['Attendance Date'].notna()]
            
            # Save to Excel with auto-column width
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Attendance')
                worksheet = writer.sheets['Attendance']
                for idx, col in enumerate(final_df.columns):
                    max_length = max(
                        final_df[col].astype(str).apply(len).max(),
                        len(str(col))
                    ) + 2
                    worksheet.column_dimensions[chr(65 + idx)].width = max_length
            
            print(f"Excel file ban gayi! Location: {excel_path}")
            return True
        else:
            print("Koi table nahi mili PDF mein")
            return False
            
    except Exception as e:
        print(f"Excel conversion mein error aaya: {str(e)}")
        return False

def process_pdf_to_excel(input_pdf, output_excel, header_size=105, footer_size=50):
    """
    Complete process: Remove header/footer and convert to Excel
    """
    temp_pdf = "temp_without_header_footer.pdf"
    
    if remove_header_footer(input_pdf, temp_pdf, header_size, footer_size):
        if convert_pdf_to_excel(temp_pdf, output_excel):
            print("Process complete ho gaya!")
        else:
            print("Excel conversion fail ho gaya")
    else:
        print("Header/footer remove karne mein fail ho gaya")
    
    if os.path.exists(temp_pdf):
        try:
            os.remove(temp_pdf)
        except Exception as e:
            print(f"Warning: Temporary file delete nahi ho payi: {str(e)}")

if __name__ == "__main__":
    input_file = "attendance.pdf"
    output_excel = "attendance.xlsx"
    
    process_pdf_to_excel(
        input_pdf=input_file,
        output_excel=output_excel,
        header_size=105,
        footer_size=50
    )
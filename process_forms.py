
import pandas as pd
import datetime as dt
import os
import shutil
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject

def main():
    """
    This script processes an Excel file to fill a PDF form template with data from each row.
    """
    # Configuration
    excel_file = 'data2026.xlsx'
    pdf_template_path = 'form_template.pdf'
    output_directory = 'data'

    # Ensure the output directory exists
    os.makedirs(output_directory, exist_ok=True)

    # Step 1: Load the Excel file
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        print(f"Error: The file '{excel_file}' was not found.")
        return
        
    df['Date'] = pd.to_datetime(df['Date']).dt.date

    # Step 2: Loop through each row in the Excel file
    for index, row in df.iterrows():
        # Load the PDF form template for each iteration
        pdf_template = PdfReader(pdf_template_path)

        # Step 3: Create a dictionary mapping Excel columns to PDF fields
        data_map = {
            'Employee Name': row.get('Employee Name', ''),
            'Employee P': row.get('Employee P#', ''),
            'Direct Supervisor': row.get('Supervisor', ''),
            'Office Address': row.get('Address', ''),
            'Office Name': row.get('Location (Office)', ''),
            'Location Code': row.get('Location Code', ''),
            'Cost Center': row.get('Cost Center', ''),
            'Effective Date of Assignment': row.get('Date', ''),
            'Property Tag Number': row.get('Property Tag#', ''),
            'Service Tag Number': row.get('Service Tag #', ''),
            "Employee Telephone": row.get('Phone#', ''),
            'Asset Description Make  Model': row.get('Make & Model', ''),
        }
        
        # Add checkbox logic if needed in the future
        # 'Laptop': '/Yes' if row.get('Device') == 'Laptop' else '',
        # 'Desktop': '/Yes' if row.get('Device') == 'Desktop' else '',
        # 'iPad': '/Yes' if row.get('Device') == 'iPad' else '',

        # Step 4: Populate fields with data
        for page in pdf_template.pages:
            annotations = page.get('/Annots')
            if annotations:
                for annotation in annotations:
                    field_name_obj = annotation.get('/T')
                    if field_name_obj:
                        field_name = field_name_obj[1:-1]  # Remove parentheses
                        if field_name in data_map:
                            annotation.update(
                                PdfDict(V=f'{data_map[field_name]}')
                            )

        # Step 5: Update NeedAppearances to ensure fields are rendered properly
        if pdf_template.Root.AcroForm:
            pdf_template.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))

        # Step 6: Save the output PDF
        assigned_to_person = row.get('Assigned_to:', 'Unknown')
        property_tag = row.get('Property Tag#', 'NoTag')
        output_pdf_path = os.path.join(output_directory, f'2024_Assignment_form_{assigned_to_person}_{property_tag}.pdf')

        # Write the filled and updated PDF
        PdfWriter().write(output_pdf_path, pdf_template)

    print(f"Processed {len(df)} rows. Filled PDFs are in the '{output_directory}' directory.")

    # Step 7: Create a zip archive of the output directory
    shutil.make_archive("data", 'zip', output_directory)
    print("Created 'data.zip' with the filled forms.")

if __name__ == "__main__":
    main()

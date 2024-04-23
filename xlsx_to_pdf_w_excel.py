import os
import win32com.client

def excel_to_pdf_with_excel(excel_file_path, output_pdf_name):
    try:
        # Ensure that the output PDF directory exists (current working directory)
        output_dir = os.getcwd()
        os.makedirs(output_dir, exist_ok=True)
        print(f"Output directory: '{output_dir}'")

        # Connect to Excel application
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False

        # Open the Excel file
        workbook = excel_app.Workbooks.Open(os.path.abspath(excel_file_path))

        # Create the full path for the output PDF
        file_name = os.path.basename(excel_file_path)
        file = os.path.splitext(file_name)
        output_pdf_path = os.path.join(output_dir, (file[0] + ".pdf"))

        # Export the workbook to PDF
        workbook.ExportAsFixedFormat(0, output_pdf_path)

        # Close the workbook and quit Excel
        workbook.Close(False)
        excel_app.Quit()

        print(f"PDF created successfully: {output_pdf_path}")
        return output_pdf_path

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

# Example usage
excel_file_path = "North_temp.xlsx"
output_pdf_name = "output.pdf"
excel_to_pdf_with_excel(excel_file_path, output_pdf_name)

input("Press Enter to close...")
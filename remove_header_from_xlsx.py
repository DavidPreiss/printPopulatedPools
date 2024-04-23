import openpyxl

def remove_header_from_xlsx(excel_file_path):
    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(excel_file_path)

        # Iterate through each worksheet
        for sheet in workbook:
            # Clear the header text in print settings
            sheet.sheet_properties.headerFooter.center_header.text = None
            sheet.sheet_properties.headerFooter.left_header.text = None
            sheet.sheet_properties.headerFooter.right_header.text = None

        # Save the modified workbook (overwrite the original file)
        workbook.save(excel_file_path)

        print("Header removed successfully")
        return True

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return False

# Example usage
excel_file_path = "North_temp.xlsx"
# remove_header_from_xlsx(excel_file_path)

def remove_header_from_excel(excel_file_path):
    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(excel_file_path)

        # Iterate through each worksheet
        for sheet in workbook:
            # Clear the header text in print settings
            sheet.oddHeader = None
            sheet.oddFooter = None
            sheet.evenHeader = None
            sheet.evenFooter = None

        # Save the modified workbook (overwrite the original file)
        workbook.save(excel_file_path)

        print("Headers and Footers removed successfully")
        return True

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return False

# Example usage
excel_file_path = "North_temp.xlsx"
remove_header_from_excel(excel_file_path)

import os
import subprocess

def check_installations():
    # Check for Microsoft Excel
    excel_installed = False
    try:
        subprocess.check_call(["excel", "/?"])
        excel_installed = True
    except (subprocess.CalledProcessError, FileNotFoundError):
        pass

    # Check for LibreOffice
    libreoffice_installed = False
    try:
        subprocess.check_call(["libreoffice", "--version"])
        libreoffice_installed = True
    except (subprocess.CalledProcessError, FileNotFoundError):
        pass

    return excel_installed, libreoffice_installed

# Example usage
excel_installed, libreoffice_installed = check_installations()
print(f"Excel installed: {excel_installed}")
print(f"LibreOffice installed: {libreoffice_installed}")

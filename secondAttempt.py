import shutil
import os
import calendar
# System call
os.system("")

# Class of different styles
class style():
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'

# print(style.GREEN + f"Hello, World! {WEEK_NUMBER}" + style.RESET) #debug
import subprocess

try:
    import pikepdf
except ImportError as e:
    print(style.RED + f"!--ERROR:{e}\n pikepdf is not installed. Installing..." + style.RESET)
    subprocess.check_call(["pip", "install", "pikepdf"])
    print("Installation complete. You can now run the script.")
    # exit()


def compress_pdf(input_file, output_file):
    # Open the original PDF
    print("compress_pdf()")
    print(f"Opening {input_file} ...")
    with pikepdf.open(input_file, allow_overwriting_input=True) as pdf:
        print("opened!")
        # Create a new PDF to store compressed version
        pdf.save(output_file, compress_streams=True)
        print(f"saved {output_file}")

input_pdf_path = 'input.pdf'  # Replace with your input PDF path
output_pdf_path = 'output.pdf'  # Replace with your desired output path



#compress_pdf(input_pdf_path, output_pdf_path)
print("good")
print(f"current location: {os.getcwd()}")
ShrinkTargetPath = "C:\\Users\\David\\Desktop\\code\\Web Archive shrink"
test0 = ShrinkTargetPath+"/P002/P002 2024 M12 Week 1.pdf"
output_pdf_path = ShrinkTargetPath+"/P002/P002 2024 M12 Week 1 SHRINK.pdf"
#compress_pdf("input.pdf", "input.pdf")

#goto webarchive
for mydir in os.listdir(os.getcwd()):
    if os.path.isdir(mydir):
        print(mydir)
        for thing in os.listdir(os.getcwd()+"/"+mydir):
            print(thing)
            jam = os.getcwd()+"/"+mydir+"/"+thing
            print(jam)
            print(jam[-4:])
            if jam[-4:] == ".pdf":
                compress_pdf(jam, jam)
#for each folder in webarchive
    #for each input.pdf that doesn't end in SHRINK.pdf
        #compress_pdf(input.pdf, inputSHRINK.pdf)

# Prompt the user to press Enter before closing
input("Press Enter to close...")
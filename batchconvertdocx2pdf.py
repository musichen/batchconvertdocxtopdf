import os
import argparse
import shutil
import subprocess

# make sure PANDOC is installed! visit https://pandoc.org/installing.html to install 
# make sure pdflatex is installed   just visit https://www.tug.org/begin.html   or if your're using Mac then install it with brew
# brew install mactex 


# Check if Pandoc is installed
try:
    subprocess.run(['pandoc', '-v'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
except FileNotFoundError:
    print('Pandoc is not installed. Please visit https://pandoc.org/installing.html to install.')
    exit()

# Check if pdflatex is installed
try:
    subprocess.run(['pdflatex', '-v'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
except FileNotFoundError:
    print('pdflatex is not installed. Please install a LaTeX distribution, such as TeX Live (https://www.tug.org/texlive/) or MiKTeX (https://miktex.org/).')
    exit()



# Parse command-line arguments
parser = argparse.ArgumentParser(description='Batch convert docx files to pdf')
parser.add_argument('input_folder', type=str, help='path to the input folder containing docx files')
parser.add_argument('output_folder', type=str, help='path to the output folder for pdf files')
args = parser.parse_args()

# Loop over the docx files in the input folder
for filename in os.listdir(args.input_folder):
    if filename.endswith('.docx'):
        # Generate the output filename by replacing the extension
        output_filename = os.path.splitext(filename)[0] + '.pdf'
        
        # Call Pandoc to convert the input file to PDF
        input_path = os.path.join(args.input_folder, filename)
        output_path = os.path.join(args.output_folder, output_filename)
        os.system(f'pandoc "{input_path}" -o "{output_path}"')

# conversion done
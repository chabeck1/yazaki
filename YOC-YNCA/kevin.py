#import pypdf
import pdfplumber
#import pandas

# Open the PDF file
with pdfplumber.open("66401-080A_002_1_0.pdf") as pdf:
    for page in pdf.pages:
        print(page.extract_tables()[0][6])
        #print(page.extract_text())

import fitz  # PyMuPDF
import pandas as pd

# Open the PDF file
pdf_path = 'JAYARAM REPORT 1.pdf'
pdf_document = fitz.open(pdf_path)

# Prepare a list to hold the extracted data
columns = [
    "Student Name", "Father Name", "Student Id", "Mother Name",
    "Habitation", "Aadhaar uid No", "DOA", "DOB", "Sex",
    "Caste", "Minority", "BPL No.", "Disadvantage", "Admin",
    "Class Studying", "Previous Class", "Status if", "Medium and Syllabus",
    "Disability", "Correction"
]
data = []

# Extract text from each page and process it
for page_num in range(len(pdf_document)):
    page = pdf_document.load_page(page_num)
    text = page.get_text("text")
    
    # Split text into lines
    lines = text.split('\n')
    
    # Initialize an empty record
    record = {}
    
    # Iterate over the lines and assign values to the record
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if "1. Student Name" in line:
            record["Student Name"] = lines[i+3].strip()
            #print(line[i + 10].strip())
            record["Father Name"] = lines[i + 32].strip()  #father's name displayed accurately
            record["Student Id"] = lines[i + 46].strip()   #correct
            record["Mother Name"] = lines[i + 26].strip()    #mother name displaying accurately
            record["Habitation"] = lines[i + 50].strip()
            record["Aadhaar uid No"] = lines[i + 12].replace("6. Aadhaar uid No", "").strip()
            record["DOA"] = lines[i + 39].strip()    #DOA displaying accurately
            record["DOB"] = lines[i + 16].replace("8. DOB", "").strip()
            record["Sex"] = lines[i + 28][4:].strip()   #gender displaying accurately
            record["Caste"] = lines[i + 33].strip()    #caste displaying accurately
            record["Minority"] = lines[i + 35].strip() #Minority displaying acuurately
            record["BPL No."] = lines[i + 42].strip()  #bond number displayed properly
            record["Disadvantage"] = 0
            record["Admin"] = lines[i + 43].strip()  #correct
            record["Class Studying"] = lines[i + 44].strip() #correct
            record["Previous Class"] = lines[i + 31].replace("16. Previous Class", "").strip()
            record["Status if"] = lines[i + 45].strip()  #displayed accurately
            record["Medium and Syllabus"] = lines[i + 29].strip()   #displayed accuratetly
            record["Disability"] = lines[i + 30][5:].strip()  #disability displaying accurately
            record["Correction"] = 0
            
            data.append(record)
            record = {}
            
            # Move to the next record
            i += 40
        else:
            i += 1

# Convert list to DataFrame
df = pd.DataFrame(data, columns=columns)

# Write DataFrame to an Excel file
excel_path = 'JAYARAM_REPORT_1.xlsx'
df.to_excel(excel_path, index=False)

print(f"Data has been successfully written")

import fitz  # PyMuPDF
import pandas as pd
import google.generativeai as genai
import json
import re
import sys

# ==========================================
# 1. CONFIGURATION
# ==========================================
API_KEY = "AQ.Ab8RN6KEIYdFoEhtBA8GF13bVVQp4YeHldSeZ95S-luAF6qMcw"  # Paste your Gemini API key here

SYLLABUS_PDF_PATH = "BE-AEROSPACE-Syllabus-R2023.pdf"
BOOK_PDF_PATH = "mechanics of material.pdf"
OUTPUT_EXCEL_PATH = "Automated_Book_Mapping.xlsx"

# Specify the page numbers for the Book's Table of Contents (0-indexed). 
# Textbooks are huge; feeding the whole book into the API is unnecessary and slow.
# Find which pages the "Contents" or "Index" span across and input them here.
BOOK_TOC_START_PAGE = 4  
BOOK_TOC_END_PAGE = 20   

# ==========================================
# 2. SETUP AI MODEL
# ==========================================
genai.configure(api_key=API_KEY)
# Using gemini-1.5-flash as it is fast, cheap, and has a large context window
model = genai.GenerativeModel('gemini-1.5-flash')

# ==========================================
# 3. PDF EXTRACTION FUNCTIONS
# ==========================================
def extract_text_from_pdf(pdf_path, start_page=0, end_page=None):
    """Extracts text from a specific page range in a PDF."""
    try:
        doc = fitz.open(pdf_path)
        text = ""
        if end_page is None:
            end_page = len(doc) - 1
            
        for page_num in range(start_page, end_page + 1):
            page = doc.load_page(page_num)
            text += page.get_text("text") + "\n"
        return text
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        sys.exit(1)

# ==========================================
# 4. EXECUTE EXTRACTION & AI MAPPING
# ==========================================
print("Extracting text from Syllabus PDF...")
syllabus_text = extract_text_from_pdf(SYLLABUS_PDF_PATH)

print(f"Extracting Table of Contents from Book PDF (Pages {BOOK_TOC_START_PAGE}-{BOOK_TOC_END_PAGE})...")
book_toc_text = extract_text_from_pdf(BOOK_PDF_PATH, BOOK_TOC_START_PAGE, BOOK_TOC_END_PAGE)

print("Sending data to AI for semantic matching. This may take a minute...")

prompt = f"""
You are an expert academic curriculum designer. Your task is to map the topics from a university syllabus to the chapters/sections of a textbook.

Here is the Syllabus Text:
---
{syllabus_text[:15000]} # Limiting chars just in case the syllabus is massive
---

Here is the Textbook Table of Contents:
---
{book_toc_text}
---

Create a mapping of the syllabus units and topics to the most relevant textbook chapters and sections. 
Output the result strictly as a valid JSON array of objects. Do not include markdown formatting like ```json.
Use this exact JSON structure:
[
  {{
    "Semester": "Semester III",
    "Course": "Course Name",
    "Unit": "Unit Name",
    "Syllabus Topic": "Specific topic from syllabus",
    "Book Mapping": "Chapter X: Section Y - Title"
  }}
]
"""

try:
    response = model.generate_content(prompt)
    response_text = response.text.strip()
    
    # Clean up markdown formatting if the model accidentally includes it
    if response_text.startswith("```json"):
        response_text = response_text[7:]
    if response_text.endswith("```"):
        response_text = response_text[:-3]
        
    mapping_data = json.loads(response_text.strip())
    print("AI Mapping successful!")
    
except Exception as e:
    print(f"Failed to generate or parse AI response. Error: {e}")
    print("Raw output:")
    print(response.text if 'response' in locals() else "No response generated.")
    sys.exit(1)

# ==========================================
# 5. EXPORT TO EXCEL
# ==========================================
print(f"Generating Excel file: {OUTPUT_EXCEL_PATH}...")
df = pd.DataFrame(mapping_data)

with pd.ExcelWriter(OUTPUT_EXCEL_PATH, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Syllabus to Book Mapping')
    
    # Auto-adjust column widths for readability
    worksheet = writer.sheets['Syllabus to Book Mapping']
    for idx, col in enumerate(df.columns, 1):
        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.column_dimensions[worksheet.cell(row=1, column=idx).column_letter].width = max_len

print("Done! Check your folder for the Excel file.")
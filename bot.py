import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os
import pandas as pd
import pdfplumber
from docx import Document
import spacy

# --- Global Variables and Setup ---
try:
    nlp = spacy.load("en_core_web_lg")
except OSError:
    messagebox.showerror("Error", "spaCy model 'en_core_web_lg' not found. Please run 'python -m spacy download en_core_web_lg' in your environment.")
    nlp = None

selected_files = []

# --- Targeted Extraction Logic ---

def extract_from_text(text):
    """
    Targeted extraction focusing on contact information sections.
    """
    data = {
        "name": "N/A",
        "phone": "N/A",
        "email": "N/A",
        "postcode": "N/A"
    }

    # Split text into lines for better analysis
    lines = text.split('\n')
    
    # Look for contact information sections
    contact_sections = find_contact_sections(lines)
    
    # Extract from contact sections first, then fall back to full text
    if contact_sections:
        contact_text = ' '.join(contact_sections)
        data = extract_from_contact_section(contact_text)
    
    # If we're still missing data, search the entire text
    if data["name"] == "N/A" or data["phone"] == "N/A" or data["email"] == "N/A" or data["postcode"] == "N/A":
        fallback_data = extract_from_full_text(text)
        for key in data:
            if data[key] == "N/A" and fallback_data[key] != "N/A":
                data[key] = fallback_data[key]
    
    return data

def find_contact_sections(lines):
    """
    Identify sections of the document that likely contain contact information.
    """
    contact_keywords = [
        'contact', 'address', 'phone', 'mobile', 'email', 'e-mail',
        'telephone', 'location', 'details', 'information',
        'reach me at', 'get in touch', 'connect with',
        'call me', 'phone me', 'contact me'
    ]
    
    contact_sections = []
    in_contact_section = False
    current_section = []
    
    for i, line in enumerate(lines):
        line_clean = line.strip().lower()
        
        # Check if this line indicates a contact section
        is_contact_header = any(keyword in line_clean for keyword in contact_keywords)
        has_contact_data = has_contact_patterns(line)
        
        if is_contact_header or has_contact_data:
            in_contact_section = True
            current_section.append(line)
        elif in_contact_section:
            # Continue collecting if line has content but isn't too long (avoid body text)
            if line.strip() and len(line.strip()) < 100:
                current_section.append(line)
            else:
                # End of contact section
                if current_section:
                    contact_sections.extend(current_section)
                    current_section = []
                in_contact_section = False
    
    # Add any remaining section
    if current_section:
        contact_sections.extend(current_section)
    
    return contact_sections

def has_contact_patterns(text):
    """Check if text contains patterns of contact information."""
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'(\+44|0)[\s\-]?(\d[\s\-]?){9,10}'
    postcode_pattern = r'[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}'
    
    return (re.search(email_pattern, text) or 
            re.search(phone_pattern, text) or 
            re.search(postcode_pattern, text))

def extract_from_contact_section(text):
    """
    Extract data specifically from identified contact sections.
    """
    data = {
        "name": "N/A",
        "phone": "N/A",
        "email": "N/A",
        "postcode": "N/A"
    }
    
    # Extract email (most reliable)
    email = extract_email_targeted(text)
    if email != "N/A":
        data["email"] = email
    
    # Extract phone
    phone = extract_phone_targeted(text)
    if phone != "N/A":
        data["phone"] = phone
    
    # Extract postcode
    postcode = extract_postcode_targeted(text)
    if postcode != "N/A":
        data["postcode"] = postcode
    
    # Extract name (tricky - look for patterns)
    name = extract_name_targeted(text)
    if name != "N/A":
        data["name"] = name
    
    return data

def extract_from_full_text(text):
    """
    Fallback extraction from entire text.
    """
    data = {
        "name": "N/A",
        "phone": "N/A",
        "email": "N/A",
        "postcode": "N/A"
    }
    
    # Use more conservative patterns for full text search
    data["email"] = extract_email_conservative(text)
    data["phone"] = extract_phone_conservative(text)
    data["postcode"] = extract_postcode_conservative(text)
    data["name"] = extract_name_conservative(text)
    
    return data

def extract_name_targeted(text):
    """Extract name from contact section."""
    # Look for name patterns in contact context
    patterns = [
        r'name\s*[:\-]\s*([A-Z][a-zA-Z\s]{2,30})',
        r'full name\s*[:\-]\s*([A-Z][a-zA-Z\s]{2,30})',
        r'contact name\s*[:\-]\s*([A-Z][a-zA-Z\s]{2,30})',
        r'^([A-Z][a-zA-Z\s]{2,30})$',  # Standalone line with just name
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            name = match.group(1).strip()
            # Basic validation - should contain letters and spaces only
            if re.match(r'^[A-Za-z\s\.]+$', name) and len(name) > 2:
                return name.title()
    
    return "N/A"

def extract_name_conservative(text):
    """Conservative name extraction from full text."""
    if not nlp:
        return "N/A"
    
    # Only look at first few sentences for name
    sentences = text.split('.')[:5]
    first_part = '. '.join(sentences)
    
    doc = nlp(first_part)
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            name = ent.text.strip()
            # Validate it looks like a real name
            if (2 <= len(name.split()) <= 3 and 
                all(part[0].isupper() for part in name.split() if part) and
                len(name) <= 30):
                return name
    
    return "N/A"

def extract_phone_targeted(text):
    """Targeted phone extraction from contact sections."""
    # UK phone patterns
    patterns = [
        r'phone\s*[:\-]\s*([+\d\s\-\(\)]{10,})',
        r'mobile\s*[:\-]\s*([+\d\s\-\(\)]{10,})',
        r'telephone\s*[:\-]\s*([+\d\s\-\(\)]{10,})',
        r'contact\s*[:\-]\s*([+\d\s\-\(\)]{10,})',
        r'call me at\s*([+\d\s\-\(\)]{10,})',
        r'(\+44\s?7\d{3}|07\d{3})\s?\d{3}\s?\d{3}',
        r'(\+44\s?1\d{3}|01\d{3})\s?\d{3}\s?\d{3}',
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            if isinstance(match, tuple):
                match = match[0]
            phone = clean_phone(match)
            if is_valid_phone(phone):
                return format_phone(phone)
    
    return "N/A"

def extract_phone_conservative(text):
    """Conservative phone extraction."""
    # Simple UK mobile pattern
    mobile_pattern = r'07\d{3}\s?\d{3}\s?\d{3}'
    match = re.search(mobile_pattern, text)
    if match:
        phone = clean_phone(match.group(0))
        if is_valid_phone(phone):
            return format_phone(phone)
    
    return "N/A"

def extract_email_targeted(text):
    """Targeted email extraction."""
    patterns = [
        r'email\s*[:\-]\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
        r'e-mail\s*[:\-]\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
        r'contact\s*[:\-]\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
        r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            email = match.group(1) if not isinstance(match.group(1), tuple) else match.group(1)[0]
            if is_valid_email(email):
                return email.lower()
    
    return "N/A"

def extract_email_conservative(text):
    """Conservative email extraction."""
    email_pattern = r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b'
    match = re.search(email_pattern, text)
    if match:
        email = match.group(0)
        if is_valid_email(email):
            return email.lower()
    
    return "N/A"

def extract_postcode_targeted(text):
    """Targeted postcode extraction."""
    patterns = [
        r'postcode\s*[:\-]\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})',
        r'post code\s*[:\-]\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})',
        r'zipcode\s*[:\-]\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})',
        r'([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            postcode = match.group(1) if not isinstance(match.group(1), tuple) else match.group(1)[0]
            formatted = format_postcode(postcode.upper())
            if is_valid_postcode(formatted):
                return formatted
    
    return "N/A"

def extract_postcode_conservative(text):
    """Conservative postcode extraction."""
    postcode_pattern = r'[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}'
    match = re.search(postcode_pattern, text)
    if match:
        formatted = format_postcode(match.group(0).upper())
        if is_valid_postcode(formatted):
            return formatted
    
    return "N/A"

def clean_phone(phone):
    """Clean phone number."""
    return re.sub(r'[^\d+]', '', phone)

def is_valid_phone(phone):
    """Validate phone number."""
    if not phone:
        return False
    # Check if it's a reasonable UK phone number length
    clean = re.sub(r'[^\d]', '', phone)
    return 10 <= len(clean) <= 12

def format_phone(phone):
    """Format phone number."""
    clean = re.sub(r'[^\d]', '', phone)
    if clean.startswith('44'):
        clean = clean[2:]
    if clean.startswith('0') and len(clean) == 11:
        return f"{clean[:5]} {clean[5:8]} {clean[8:]}"
    return clean

def is_valid_email(email):
    """Basic email validation."""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def is_valid_postcode(postcode):
    """Basic postcode validation."""
    pattern = r'^[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}$'
    return re.match(pattern, postcode) is not None

def format_postcode(postcode):
    """Format postcode."""
    clean = re.sub(r'\s+', '', postcode.upper())
    if len(clean) >= 5:
        return f"{clean[:-3]} {clean[-3:]}"
    return clean

def extract_from_pdf(filepath):
    """PDF extraction with focus on structure."""
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                # Try to extract text maintaining some structure
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        
        return extract_from_text(text) if text.strip() else None
    except Exception as e:
        print(f"Error reading PDF file {filepath}: {e}")
        return None

def extract_from_docx(filepath):
    """DOCX extraction."""
    text = ""
    try:
        document = Document(filepath)
        for paragraph in document.paragraphs:
            if paragraph.text.strip():
                text += paragraph.text + "\n"
        
        return extract_from_text(text) if text.strip() else None
    except Exception as e:
        print(f"Error reading DOCX file {filepath}: {e}")
        return None

# Rest of the GUI code remains the same...
def process_files():
    """
    Main function to process all selected files and generate the CSV.
    """
    if not selected_files:
        messagebox.showwarning("Warning", "Please select one or more files first.")
        return

    if not nlp:
        messagebox.showerror("Error", "spaCy model is not loaded. Cannot proceed with name extraction.")
        return

    results = []
    total_files = len(selected_files)
    
    status_label.config(text=f"Processing 0 of {total_files} files...")
    app.update()

    for i, filepath in enumerate(selected_files):
        filename = os.path.basename(filepath)
        status_label.config(text=f"Processing {i+1} of {total_files} files: {filename}...")
        app.update()

        file_extension = os.path.splitext(filename)[1].lower()
        
        extracted_data = None
        if file_extension == ".pdf":
            extracted_data = extract_from_pdf(filepath)
        elif file_extension == ".docx":
            extracted_data = extract_from_docx(filepath)
        else:
            print(f"Skipping unsupported file type: {filename}")
            continue

        if extracted_data:
            extracted_data["filename"] = filename
            ordered_data = {
                "filename": extracted_data["filename"],
                "name": extracted_data["name"],
                "phone": extracted_data["phone"],
                "email": extracted_data["email"],
                "postcode": extracted_data["postcode"]
            }
            results.append(ordered_data)
            
            # Debug print
            print(f"Extracted from {filename}: {ordered_data}")
        else:
            results.append({
                "filename": filename,
                "name": "ERROR",
                "phone": "ERROR",
                "email": "ERROR",
                "postcode": "ERROR"
            })

    if results:
        df = pd.DataFrame(results)
        output_filename = "output.csv"
        try:
            df.to_csv(output_filename, index=False)
            # Show summary
            successful = len([r for r in results if all(r[field] not in ["N/A", "ERROR"] for field in ["name", "phone", "email", "postcode"])])
            messagebox.showinfo("Success", 
                              f"Extraction complete!\n\n"
                              f"Files processed: {len(results)}\n"
                              f"Fully successful: {successful}\n"
                              f"CSV saved as '{output_filename}'")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV file: {e}")
    else:
        messagebox.showwarning("Warning", "No data was successfully extracted.")

    selected_files.clear()
    file_count_label.config(text="0 files selected.")
    status_label.config(text="Ready.")

def select_files():
    global selected_files
    filetypes = [
        ("Supported Documents", "*.pdf *.docx"),
        ("PDF files", "*.pdf"),
        ("DOCX files", "*.docx"),
        ("All files", "*.*")
    ]
    
    filepaths = filedialog.askopenfilenames(
        title="Select PDF and DOCX Files",
        filetypes=filetypes
    )
    
    if filepaths:
        selected_files = list(filepaths)
        file_count_label.config(text=f"{len(selected_files)} files selected.")
        status_label.config(text="Files ready for extraction.")
    else:
        file_count_label.config(text="0 files selected.")
        status_label.config(text="Ready.")

# GUI Setup
app = tk.Tk()
app.title("Smart Document Data Extractor")
app.geometry("450x280")
app.resizable(False, False)

padding_options = {'padx': 10, 'pady': 10}

main_frame = tk.Frame(app, **padding_options)
main_frame.pack(fill='both', expand=True)

upload_button = tk.Button(
    main_frame,
    text="1. Select PDF/DOCX Files",
    command=select_files,
    height=2,
    width=35
)
upload_button.pack(**padding_options)

file_count_label = tk.Label(
    main_frame,
    text="0 files selected.",
    font=('Arial', 10)
)
file_count_label.pack(**padding_options)

extract_button = tk.Button(
    main_frame,
    text="2. Run Smart Extraction and Generate CSV",
    command=process_files,
    height=2,
    width=35,
    bg='lightgreen'
)
extract_button.pack(**padding_options)

status_label = tk.Label(
    app,
    text="Ready.",
    bd=1,
    relief=tk.SUNKEN,
    anchor=tk.W
)
status_label.pack(side=tk.BOTTOM, fill=tk.X)

if __name__ == "__main__":
    app.mainloop()
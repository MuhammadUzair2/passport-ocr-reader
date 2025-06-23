import os
import re
import io
import pandas as pd
from PIL import Image
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path

# ✅ SET YOUR TESSERACT PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\HP\Desktop\tesseract-ocr\tesseract.exe'

uploaded_files = set()

def extract_text(file_path):
    try:
        if file_path.lower().endswith('.pdf'):
            pages = convert_from_path(file_path, dpi=300)
            image = pages[0].convert('RGB')
        else:
            with open(file_path, 'rb') as f:
                image = Image.open(io.BytesIO(f.read())).convert('RGB')

        text = pytesseract.image_to_string(image)
        print("🧾 OCR TEXT:\n", text)
        return text
    except Exception as e:
        messagebox.showerror("Error Reading File", str(e))
        return ""

def extract_data(text):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    mrz_lines = [line for line in lines if len(line) >= 40 and '<' in line]

    passport_no, name, surname, dob, issue_date, expiry_date = '', '', '', '', '', ''

    # MRZ Parsing
    if len(mrz_lines) >= 2:
        mrz1, mrz2 = mrz_lines[-2], mrz_lines[-1]
        match = re.match(r'P<\w{3}([A-Z<]+)<<([A-Z<]+)', mrz1)
        if match:
            surname_raw = match.group(1).replace('<', ' ').strip()
            given_raw = match.group(2).replace('<', ' ').strip()
            name = f"{given_raw} {surname_raw}"
        passport_no = mrz2[0:9].replace('<', '').strip()
        dob_raw = mrz2[13:19]
        if len(dob_raw) == 6:
            dob = f"{dob_raw[4:6]}/{dob_raw[2:4]}/19{dob_raw[0:2]}"
        expiry_raw = mrz2[21:27]
        if len(expiry_raw) == 6:
            expiry_date = f"{expiry_raw[4:6]}/{expiry_raw[2:4]}/20{expiry_raw[0:2]}"

    # Surname (Father Name)
    for i, line in enumerate(lines):
        if "father" in line.lower():
            if i + 1 < len(lines):
                surname = lines[i + 1].strip()
            break
    if not surname:
        for line in lines:
            if line.isupper() and "PAKISTAN" not in line and len(line.split()) > 1:
                surname = line.strip()
                break

    # Date of Birth (human readable)
    dob_match = re.search(r'Date\s*of\s*Birth\s*[:\-]?\s*(\d{1,2}\s+[A-Z]+\s+\d{4})', text, re.IGNORECASE)
    if dob_match:
        dob = dob_match.group(1).strip()

    # Issue Date (Issuing Authority fallback)
    for i, line in enumerate(lines):
        if "issuing authority" in line.lower() or "date of issue" in line.lower():
            for j in range(i, min(i + 3, len(lines))):
                match = re.search(r'\d{1,2}\s+[A-Z]+\s+\d{4}', lines[j])
                if match:
                    issue_date = match.group(0)
                    break
            if issue_date:
                break

    # Expiry Date (human readable)
    expiry_match = re.search(r'(Date of Expiry|Valid Until)[^\n]*?(\d{1,2}\s+[A-Z]+\s+\d{4})', text, re.IGNORECASE)
    if expiry_match:
        expiry_date = expiry_match.group(2).strip()

    return {
        'Passport Number': passport_no,
        'Name': name,
        'Surname': surname,
        'Date of Birth': dob,
        'Date of Passport Issue': issue_date,
        'Date of Expiry': expiry_date
    }

def save_to_excel(data, filename='passport_data.xlsx', sheet_name='Sheet1'):
    df_new = pd.DataFrame([data])
    column_order = [
        'Passport Number',
        'Name',
        'Surname',
        'Date of Birth',
        'Date of Passport Issue',
        'Date of Expiry'
    ]
    df_new = df_new[column_order]

    try:
        if os.path.exists(filename):
            df_existing = pd.read_excel(filename, sheet_name=sheet_name)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
            df_combined.to_excel(writer, sheet_name=sheet_name, index=False)

    except PermissionError:
        backup = 'passport_data_backup.xlsx'
        with pd.ExcelWriter(backup, engine='openpyxl', mode='w') as writer:
            df_new.to_excel(writer, sheet_name=sheet_name, index=False)
        messagebox.showwarning("File In Use", f"Could not write to '{filename}'. Saved to '{backup}' instead.")

def upload_and_process_passport():
    file_path = filedialog.askopenfilename(
        title="Select Passport File",
        filetypes=[
            ("Supported Files", "*.jpg *.jpeg *.png *.bmp *.tiff *.webp *.jfif *.gif *.pdf"),
            ("All Files", ".")
        ]
    )
    if not file_path:
        return

    if file_path in uploaded_files:
        retry = messagebox.askyesno(
            "Duplicate File",
            "This passport was already added. Do you want to add it again?"
        )
        if not retry:
            return

    text = extract_text(file_path)
    if not text.strip():
        messagebox.showwarning("No Text Found", "Could not extract any text from this file.")
        return

    data = extract_data(text)
    if any(data.values()):
        save_to_excel(data)
        uploaded_files.add(file_path)
        messagebox.showinfo("✅ Success", f"Data extracted and saved:\n\n{data}")
    else:
        messagebox.showwarning("No Data Found", "Could not find valid passport data.")

# GUI Setup
root = tk.Tk()
root.title("Passport OCR Reader")
root.geometry("420x220")
root.resizable(False, False)

label = tk.Label(root, text="Upload a passport image or PDF file", font=("Arial", 12))
label.pack(pady=20)

upload_button = tk.Button(root, text="Upload Passport File",
                          command=upload_and_process_passport,
                          height=2, width=30, bg="blue", fg="white")
upload_button.pack(pady=10)

root.mainloop()


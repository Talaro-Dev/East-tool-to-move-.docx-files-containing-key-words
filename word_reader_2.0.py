import os
import shutil
import re
from docx import Document


def check_directory_for_files(directory_path):
    found_files = []
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                found_files.append(file_path)
    return found_files


def extract_risk_factor(file_path):
    doc = Document(file_path)
    pattern = r"risk factor (\d+)%" #pattern to read
    
    for para in doc.paragraphs:
        match = re.search(pattern, para.text, re.IGNORECASE)  #not caracter-sensitive search
        if match:
            return int(match.group(1))  # take the percentage as int
    return None  


def scan_directory_for_risk_factors(directory_path, low_risk_folder, high_risk_folder):
    found_files = check_directory_for_files(directory_path)

    low_risk_files = []
    high_risk_files = []
    
    for file_path in found_files:
        risk_factor = extract_risk_factor(file_path)
        if risk_factor is not None:
            if risk_factor < 4:
                low_risk_files.append(file_path)
            elif risk_factor >= 6:
                high_risk_files.append(file_path)

    
    if low_risk_files:
        move_files_to_folder(low_risk_files, low_risk_folder) #low risk folder
    if high_risk_files:
        move_files_to_folder(high_risk_files, high_risk_folder) #high risk folder

    return low_risk_files, high_risk_files


def move_files_to_folder(files_to_move, destination_folder):
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    for file_path in files_to_move:
        file_name = os.path.basename(file_path)
        destination_path = os.path.join(destination_folder, file_name)
        shutil.move(file_path, destination_path)
        print(f"Moved file: {file_name} to {destination_folder}")

# Input values (r contains raw string)
directory = r"directory path" #path to read files
low_risk_folder = r"low_risk folder path" #low risk path
high_risk_folder = r"high risk folder path" #high risk path


low_risk_files, high_risk_files = scan_directory_for_risk_factors(directory, low_risk_folder, high_risk_folder)

#print for debugging
if low_risk_files:
    print("Files with risk factors < 4% moved to low_risk_folder:")
    for file in low_risk_files:
        print(file)
else:
    print("No files with risk factors < 4% found.")

if high_risk_files:
    print("Files with risk factors >= 6% moved to high_risk_folder:")
    for file in high_risk_files:
        print(file)
else:
    print("No files with risk factors >= 6% found.")

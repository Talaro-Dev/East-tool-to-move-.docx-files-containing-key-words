import os
import shutil
from docx import Document

def check_directory_for_files(directory_path):
    # List all .docx files in the directory and its subdirectories
    found_files = []
    
    # Walk through all files in the directory
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                found_files.append(file_path)
    
    return found_files

def check_word_in_docx(file_path, word_to_check):
    doc = Document(file_path)
    
    for para in doc.paragraphs:
        if word_to_check in para.text:
            return True
    return False
    
    
def scan_directory_for_word(directory_path, word_to_check, destination_folder):
    found_files = []
    
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".docx"):
              file_path = os.path.join(root,file)
              if check_word_in_docx(file_path,word_to_check):
                  found_files.append(file_path)
    if found_files:
        move_files_to_folder(found_files, destination_folder)
    return found_files

def move_files_to_folder(files_to_move, destination_folder):
    
    #Create the destination folder if it doesn't exist
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    for file_path in files_to_move:
        file_name = os.path.basename(file_path) #get the file name
        destination_path = os.path.join(destination_folder, file_name)
        shutil.move(file_path, destination_path) #move the file
        #print("Moved files") #check if moved print


#Input values (r contains raw string)
directory = r"C:\Users\georg\Desktop\python\pap"
destination_directory = r"C:\Users\georg\Desktop\python\pap\found_docs"
word="γιώργος"

# Find and move files containing the word
found_docs = check_directory_for_files(directory)

if found_docs:
    print("Found the following .docx files:")
    for doc in found_docs:
        print(doc)
else:
    print(f"No .docx files found in the directory: {directory}")


found_docs = scan_directory_for_word(directory, word, destination_directory)
if found_docs:
    print("The word " + word + " was found in the following documents:")
    for doc in found_docs:
        print(doc)
else:
     print("The word " + word + " was not found in any documents")

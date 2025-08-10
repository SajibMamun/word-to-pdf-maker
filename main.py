import os
import shutil
import subprocess

# === Configuration ===
root_folder = r"C:\Users\Sajib Bin Mamun\Desktop\Word to PDF\Skripte & Themenliste-20250810T031855Z-1-003"          # Your main folder
word_folder = r"C:\Users\Sajib Bin Mamun\Desktop\Word to PDF\word_files"      # Folder to store collected Word files
pdf_folder = r"C:\Users\Sajib Bin Mamun\Desktop\Word to PDF\pdf_files"        # Folder to store converted PDFs

os.makedirs(word_folder, exist_ok=True)
os.makedirs(pdf_folder, exist_ok=True)

# Step 1: Collect all Word files
for dirpath, _, filenames in os.walk(root_folder):
    for filename in filenames:
        if filename.lower().endswith((".doc", ".docx")):
            source_path = os.path.join(dirpath, filename)
            target_path = os.path.join(word_folder, filename)
            shutil.copy2(source_path, target_path)
            print(f"Copied: {source_path} -> {target_path}")

# Step 2: Convert Word to PDF using LibreOffice (preserves original quality)
for filename in os.listdir(word_folder):
    if filename.lower().endswith((".doc", ".docx")):
        word_path = os.path.abspath(os.path.join(word_folder, filename))
        try:
            subprocess.run([
                "soffice", "--headless", "--convert-to", "pdf",
                "--outdir", os.path.abspath(pdf_folder), word_path
            ], check=True)
            print(f"Converted to PDF: {filename}")
        except subprocess.CalledProcessError as e:
            print(f"Failed to convert {filename}: {e}")

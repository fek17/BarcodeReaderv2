"""
*********************************************************
*                                                       *
*  Fatima Khan                                           *
*  Date: 2024-08-27                                      *
*                                                       *
*  Purpose:                                              *
*  This script is designed to extract barcodes          *
*                                                       *
*  Version: 2.0                                          *
*                                                       *
*********************************************************
"""

import os
import openpyxl
from pyzbar.pyzbar import decode
from PIL import Image

# Local folder path
folder_path = "C:\\Users\\Fatima.Khan\\Downloads\\iCloud Photos\\Spar"

# Excel file path (new file will be created)
excel_file = "C:\\Users\\Fatima.Khan\\Downloads\\barcode_extraction_results_spar.xlsx"

# Initialize a new workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Extraction Results"

# Add headers to the Excel sheet
ws.append(["File Name", "Extracted Barcode Number"])

# Function to extract barcode locally using pyzbar
def extract_barcode(image_path):
    image = Image.open(image_path)
    decoded_objects = decode(image)
    
    # Define a list of barcode types you're interested in
    valid_barcode_types = ['EAN13', 'EAN8', 'UPC-A', 'UPC-E', 'CODE128', 'CODE39']

    if decoded_objects:
        for obj in decoded_objects:
            if obj.type in valid_barcode_types:
                barcode_data = obj.data.decode('utf-8')
                return barcode_data
    return "No valid barcode found"

# Loop through each file in the local folder
for file_name in os.listdir(folder_path):
    if file_name.lower().endswith((".jpeg", ".jpg", ".png")):  # Process only image files
        print("Processing " + file_name)
        image_path = os.path.join(folder_path, file_name)
        
        # Extract barcode locally
        extracted_barcode_number = extract_barcode(image_path)
        
        # Append the result to the Excel sheet
        ws.append([file_name, extracted_barcode_number])

# Save the new Excel file
wb.save(excel_file)

print(f"Results have been saved to {excel_file}")

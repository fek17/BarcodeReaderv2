
"""
*********************************************************
*                                                       *
*  Fatima Khan                                           *
*  Date: 2024-08-27                                      *
*                                                       *
*  Purpose:                                              *
*  This script is designed to extract barcodes from       *
*  meat packaging                                          *
*                                                         *
*  Version: 2.0                                          *
*                                                       *
*********************************************************
"""

import os
import base64
import requests
import openpyxl
import json
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
    
    if decoded_objects:
        for obj in decoded_objects:
            barcode_data = obj.data.decode('utf-8')
            # You can set the confidence to 'high' here since pyzbar generally returns accurate results
            return barcode_data, "high"
    return "No Code Found", "N/A"

# Loop through each file in the local folder
for file_name in os.listdir(folder_path):
    if file_name.lower().endswith((".jpeg", ".jpg", ".png")):  # Process only image files
        print("Processing "+file_name)
        image_path = os.path.join(folder_path, file_name)
        
        # Extract barcode locally
        extracted_barcode_number, barcode_confidence = extract_barcode(image_path)
        
        ws.append([file_name, extracted_barcode_number])
        
    else:
        # Handle case where no valid response is received
        ws.append([file_name, extracted_barcode_number, "No Response", "N/A", "No Response", "N/A"])

# Save the new Excel file
wb.save(excel_file)

print(f"Results have been saved to {excel_file}")

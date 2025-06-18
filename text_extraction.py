import pytesseract
import os
import csv

### Initial Setup ###
tess_path = r"C:/Users/HP/AppData/Local/Programs/Tesseract-OCR/tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = tess_path
os.environ['TESSDATA_PREFIX'] = 'C:/Users/HP/AppData/Local/Programs/Tesseract-OCR/tessdata'


def auto_annotate_ocr(image_path, csv_file, output_csv):
    # Open the input CSV file for reading
    with open(csv_file, 'r', encoding='latin1') as infile:
        csv_reader = csv.reader(infile)
        
        # Open the output CSV file for writing
        with open(output_csv, 'w', newline='', encoding='utf-8') as outfile:
            csv_writer = csv.writer(outfile)

            # Write the header row to the output CSV file
            header = next(csv_reader)
            csv_writer.writerow(header)

            # Iterate through each row in the CSV file
            for row in csv_reader:
                if not row:
                    continue  # Skip empty rows
                
                image_name = row[0]  # Assuming the image name is in the first column
                image_name_full = os.path.join(image_path, image_name)
                
                try:
                    extracted_text = pytesseract.image_to_string(image_name_full, lang='ben+eng', config='--psm 4')
                except FileNotFoundError:
                    print(f"Error: File '{image_name_full}' not found.")
                    extracted_text = ""  # Set extracted text to empty string
                except pytesseract.TesseractError as e:
                    print(f"Error processing '{image_name_full}': {e}")
                    extracted_text = ""  # Set extracted text to empty string
                
                # Update the label column with extracted text
                row[-1] = extracted_text  # Assuming the label column is the last column
                
                # Write the updated row to the output CSV file
                csv_writer.writerow(row)

"""
# Example usage
image_folder_path = "C:/Users/HP/Desktop/med_lines/3.Bikas/images"
csv_file_path = "C:/Users/HP/Desktop/med_lines/3.Bikas/labels.csv"
output_csv_path = "C:/Users/HP/Desktop/med_lines/3.Bikas/updated_labels.csv"

auto_annotate(image_folder_path, csv_file_path, output_csv_path)
"""
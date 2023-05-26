import cv2
import pytesseract
import pandas as pd
import openpyxl

def extract_data_from_image(image_path):
    # Load the image using OpenCV
    image = cv2.imread(image_path)

    # Preprocess the image (grayscale)
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Apply OCR to extract the text from the image
    extracted_text = pytesseract.image_to_string(gray_image)

    # Split the extracted text into rows and columns
    rows = extracted_text.split('\n')
    data = [row.split('\t') for row in rows]

    # Create a DataFrame from the extracted data
    df = pd.DataFrame(data)

    return df

def input_data_into_excel(dataframe, output_file):
    # Create a new workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Input the data into the Excel sheet
    for row in dataframe.iterrows():
        row_values = row[1]
        sheet.append(list(row_values))

    # Save the workbook to the output file
    workbook.save(output_file)

# Example usage:
image_path = 'tableimage2.png'
output_file = 'test.xlsx'

extracted_data = extract_data_from_image(image_path)
input_data_into_excel(extracted_data, output_file)
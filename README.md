# Excel Extraction and Excel Input

This program extracts tabular data from an image, performs OCR (Optical Character Recognition) to convert the text, and inputs the data into an Excel spreadsheet. The script utilizes Python, OpenCV, pytesseract, pandas, and openpyxl libraries to accomplish these tasks.

## Installation

1. Clone the repository or download the program files.
2. Install the required libraries using pip:
 - pip install opencv-python
 - pip install pytesseract
 - pip install pandas
 - pip install openpyxl

## Usage

1. Place the image file (containing the table) in the same directory as the program files.
2. Update the `image_path` variable in the script with the path to your input image file.
3. Specify the output file name and path by modifying the `output_file` variable.
4. Run the script
5. The script will extract the tabular data from the image, perform OCR to convert it into text, and input the data into an Excel spreadsheet.
6. The output Excel file will be created in the specified location.

## Limitations

- The accuracy of the program depends on the quality of the image and the OCR performance. It may not be able to accurately extract data from images with low resolution, poor lighting, or complex layouts.
- The program assumes that the table in the image has regular row and column structures. Irregular tables may not be processed correctly.
- The layout of the extracted data in the Excel file is an approximation. The program replicates the row and column structure but does not retain cell formatting, such as merged cells, cell borders, or cell styles.

Feel free to modify the content according to your specific program and requirements. Include any additional details, instructions, or acknowledgments as needed.

Made By Conor Phan

# Metadata Extractor

This Python script is designed to extract metadata from various types of files including PDF, DOCX, PPTX, XLSX, TXT, JPG, and JPEG.

## Functionality

The script contains functions to extract metadata from each supported file type:

- `pdf_metadata_extractor(file_path)`: Extracts and prints metadata from a PDF file.
- `docx_metadata_extractor(file_path)`: Extracts and prints metadata from a DOCX file.
- `ppt_metadata_extractor(file_path)`: Extracts and prints metadata from a PPTX file.
- `image_metadata_extractor(file_path)`: Extracts and prints metadata from an image file (JPG or JPEG).
- `text_file_metadata_extractor(file_path)`: Extracts and prints metadata from a TXT file.
- `excel_file_metadata_extractor(file_path)`: Extracts and prints metadata from an XLSX file.

The `main()` function handles user input and calls the appropriate metadata extraction function based on the file extension.

## How to Run

1. Ensure you have Python installed on your system.
2. Install the required libraries by running `pip install PyPDF2 python-docx python-pptx pyexiv2 openpyxl`.
3. Save the script to your local machine.
4. Run the script in a terminal with the command `python <script_name>.py`.
5. When prompted, enter the path to the file you want to extract metadata from.

Please note that this script only supports PDF, DOCX, PPTX, XLSX, TXT, JPG, and JPEG files. If you provide a file with a different extension or an invalid path, the script will print an error message.

## Error Handling

The script includes error handling for file not found errors and other exceptions that may occur during metadata extraction. If an error occurs, the script will print an error message detailing what went wrong.

## Future Improvements

Future versions of this script could include support for additional file types and more robust error handling. Contributions are welcome!

Please let me know if you need any further assistance! 😊

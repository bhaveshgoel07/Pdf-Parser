# OCR Question Paper Transcription Tool

This project is an OCR-based automation tool designed to transcribe text from a collection of question paper PDFs and save the extracted text as editable DOC files. This tool is particularly useful for converting scanned documents into a structured text format, making the content accessible for further analysis and editing.

## Features

- **Batch Processing**: Processes multiple PDFs in a batch, suitable for large datasets like the 700 question paper PDFs used in this project.
- **OCR Text Extraction**: Utilizes PyTesseract to perform Optical Character Recognition (OCR) on each page image within the PDFs.
- **Document Conversion**: Saves the transcribed text as DOC files for ease of access and editing.
- **Customizable Pipeline**: Modify the OCR and PDF handling parameters to suit your specific document processing needs.

## Technologies Used

- **PyPDF2**: For handling and processing PDF files, enabling efficient page extraction.
- **PyTesseract**: For OCR functionality, enabling text extraction from image-based content in PDF files.

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/ocr-question-paper-tool.git
    cd ocr-question-paper-tool
    ```

2. Install the required libraries:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Place all PDF files to be processed in a designated folder (e.g., `input_pdfs`).
2. Run the main script to start processing:
    ```bash
    python main.py
    ```
3. Extracted text will be saved as DOC files in the specified output directory (e.g., `output_docs`).

## Project Structure

- `main.py`: Main script to process PDFs and save text as DOC files.
- `requirements.txt`: File containing all required libraries for the project.

## Requirements

All dependencies for this project are listed in `requirements.txt`.

## Example

After running the script, DOC files containing the transcribed text will be saved to the `output_docs` directory.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [PyTesseract](https://github.com/madmaze/pytesseract)
- [PyPDF2](https://github.com/mstamy2/PyPDF2)

---

Feel free to reach out with any questions or feedback!

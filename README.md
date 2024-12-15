### Resume Email Extractor

### Description
This project extracts email addresses from resumes in various formats (PDF and DOCX). It processes Google Drive links containing the resumes, converts them into downloadable links, and extracts emails using regular expressions. The results are stored in an Excel file with the email addresses added in a new column.

### Features
- Converts Google Drive view links to downloadable links
- Handles both PDF and DOCX file types as of now
- Extracts email addresses from PDF and DOCX files
- Outputs the extracted emails into an Excel file

### Prerequisites
To run this project, you'll need to have Python 3.x installed on your machine, as well as the following Python libraries:

- `pandas` – for working with Excel files and data manipulation
- `requests` – for downloading the files from the URLs
- `pyMuPDF` – for extracting text from PDF files
- `python-docx` – for extracting text from DOCX files
- `openpyxl` – for working with Excel files

You can install all the necessary dependencies by running:

```bash
pip install -r requirements.txt
```


### Setup
- Clone or download the project repository to your local machine.
- Install the required libraries by running pip install -r requirements.txt from the project folder.
- Place your Excel file (containing links to resumes) in the project folder.

### Usage
- Open the Python script and edit the file_path variable to point to your Excel file.
- The program will process the resume links, extract the emails, and save the updated data into a new Excel file.

### How It Works
The script loops through the Google Drive links provided in the Excel file and downloads the corresponding resumes (PDF or DOCX format).
It extracts text from the PDF using the pyMuPDF library and from DOCX using the python-docx library.
Regular expressions are used to extract email addresses from the text content.
The extracted email addresses are saved back into a new Excel file for review.

### Limitations
- The script currently supports extracting emails only from PDF and DOCX formats.
- Some resume formats may not allow for successful extraction due to text formatting issues or encryption.
- Supports RFC5322 format mail ids only


# PDF to Excel

This is a web application built with Flask that allows users to convert tables from PDF files into Excel format. Users can upload their PDF files, select specific tables they want to convert, and choose whether to combine multiple tables into a single Excel file. The converted Excel files can then be downloaded for further use.

## Features

- **Upload PDF**: Users can upload their PDF files containing tables.
- **Table Selection**: Users can select specific tables they want to convert into Excel format.
- **Combine Tables**: Users can choose to combine multiple selected tables into a single Excel file. NB: tables must have the same column structure
- **Password protected**: Files are converted and password protected. A unique .txt file is generated which contains the password needed to unlock the excel file.
- **Download**: Converted Excel files can be downloaded for further use.
- **Error Handling**: The application provides error handling for various scenarios, such as empty PDF files or unsupported table formats.

## Technologies Used

- **Flask**: Python-based web framework used for backend development.
- **Bootstrap**: Frontend framework for building responsive and visually appealing web pages.
- **Pandas**: Library for data manipulation and analysis, used for handling Excel files and data processing.
- **Tabula**: Library for extracting tables from PDF files.
- **BeautifulSoup**: Library for parsing HTML and XML documents, used for HTML table generation.
- **Win32com**: Library for interacting with Windows COM objects.
- **Pathlib**: Library for working with file paths.

## Usage

1. **Upload PDF**: Select a PDF file containing tables and upload it using the provided form.
2. **Select Tables**: Check the boxes next to the tables you want to convert into Excel format.
3. **Combine Tables** (Optional): Check the "Combine Tables" box if you want to merge multiple selected tables into a single Excel file.
4. **Generate Excel**: Click the "Generate Excel" button to initiate the conversion process.
5. **Download**: Once the conversion is complete, the download button will appear. Click it to download the converted Excel file.
6. **Unlock file**: A zip file will be downloaded containing a .txt password file and an excel file. Copy and paste the password and the file will be unlocked for use. 

#### Additional
- The code contains a variable named 'substrings_to_remove'. This variable has a list that allows you to exclude certain tables by name. Simply add the name of the sheet to the list and it will be excluded from the excel file. 

## Installation

1. Clone the repository:

```
git clone <repository_url>
```

2. Run the Flask application:

```
python app.py
```

3. Access the application in your web browser at `http://localhost:5000`.

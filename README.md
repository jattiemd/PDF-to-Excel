# PDF to Excel

This is a web application built with Flask that allows users to convert tables from PDF files into Excel format. Users can upload their PDF files, select specific tables they want to convert, and choose whether to combine multiple tables into a single Excel file. The converted Excel files can then be downloaded for further use.

## Features

- **Upload PDF**: Users can upload their PDF files containing tables.
- **Table Selection**: Users can select specific tables they want to convert into Excel format.
- **Combine Tables**: Users can choose to combine multiple selected tables into a single Excel file. NB: tables must have the same column structure
- **Conversion**: PDF tables are converted into Excel format with proper formatting.
- **Download**: Converted Excel files can be downloaded for further use.
- **Error Handling**: The application provides error handling for various scenarios, such as empty PDF files or unsupported table formats.

## Technologies Used

- **Flask**: Python-based web framework used for backend development.
- **Bootstrap**: Frontend framework for building responsive and visually appealing web pages.
- **Pandas**: Python library for data manipulation and analysis, used for handling Excel files and data processing.
- **Tabula**: Python library for extracting tables from PDF files.
- **BeautifulSoup**: Python library for parsing HTML and XML documents, used for HTML table generation.

## Usage

1. **Upload PDF**: Select a PDF file containing tables and upload it using the provided form.
2. **Select Tables**: Check the boxes next to the tables you want to convert into Excel format.
3. **Combine Tables** (Optional): Check the "Combine Tables" box if you want to merge multiple selected tables into a single Excel file.
4. **Generate Excel**: Click the "Generate Excel" button to initiate the conversion process.
5. **Download**: Once the conversion is complete, the download button will appear. Click it to download the converted Excel file.

## Installation

1. Clone the repository:

```
git clone <repository_url>
```

2. Install dependencies:

```
pip install -r requirements.txt
```

3. Run the Flask application:

```
python app.py
```

4. Access the application in your web browser at `http://localhost:5000`.

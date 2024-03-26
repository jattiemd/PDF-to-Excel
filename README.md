# PDF to Excel Converter

This Flask web application converts PDF files containing tables into Excel format. It allows users to upload PDF files, extract tables from them, and then generate Excel files based on selected tables. Users can download the resulting Excel files once the conversion is complete.

## Features

- **PDF to Excel Conversion**: Converts PDF files containing tables into Excel format.
- **Selective Table Generation**: Allows users to select specific tables from the PDF for conversion to Excel.
- **Downloadable Excel Files**: Users can download the converted Excel files.
- **Error Handling**: Provides error messages for cases where the PDF file has no tables or if no file is uploaded.

## Requirements

- Python 3.x
- Flask
- pandas
- tabula-py
- BeautifulSoup

## Usage

1. Clone this repository to your local machine.
   
   ```
   git clone https://github.com/your_username/your_repository.git
   ```

2. Navigate to the project directory.

   ```
   cd your_repository
   ```

3. Install the required dependencies.

   ```
   pip install -r requirements.txt
   ```

4. Run the Flask application.

   ```
   python app.py
   ```

5. Open your web browser and go to `http://localhost:5000`.

6. Upload a PDF file containing tables.

7. Select the desired tables for conversion or choose all tables.

8. Click on the "Generate Excel" button to initiate the conversion process.

9. Once the conversion is complete, download the Excel file.

## Folder Structure

- `file_handler/`: Directory for storing uploaded files.
- `templates/`: HTML templates for the web interface.

## Acknowledgments

- This project was inspired by the need for a simple PDF to Excel converter.
- Special thanks to the Flask, pandas, tabula-py, and BeautifulSoup developers for their excellent libraries.
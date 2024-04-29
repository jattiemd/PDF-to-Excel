from flask import Flask, flash, render_template, request, redirect, send_file, after_this_request, session
from bs4 import BeautifulSoup
from win32com.client.gencache import EnsureDispatch
from pathlib import PurePath
import pythoncom
import secrets, os, threading, tabula, zipfile, pandas as pd


app = Flask(__name__)
app.config['SECRET_KEY'] = secrets.token_hex()
UPLOAD_FOLDER = 'file_handler\\'
lock = threading.Lock()
password = secrets.token_hex()


# Main page
@app.route('/', methods=['GET', 'POST'])
def index():
    html_tables = {}
    checkbox_value = False
    global index_error

    if 'USER_IP' not in session:
        session['USER_IP'] = request.remote_addr
    if 'UPLOADS' not in session:
        session['UPLOADS'] = []

    # Catch request
    if request.method == 'POST':
        file = request.files.get('PdfFile', None)

        # Since POST is taking place more than once to this route
        # file will = None after the first POST. As such, this control flow remedies the error
        if file:
            do_conversion(file)
        if not session.get('EXCEL_FILE_NAME'):
            flash('Please upload a pdf file')

        # If PDF file was saved successfully
        if os.path.exists(f'{UPLOAD_FOLDER}{session.get("PDF_FILE_NAME")}'):
            excel_file_path = os.path.join(f'{UPLOAD_FOLDER}{session.get("EXCEL_FILE_NAME")}')
            
            # error handling: empty pdf file, no tables in pdf file
            if excel_file_path.endswith('None') or not os.path.exists(excel_file_path):
                remove_files(session.get('PDF_FILE_NAME'), session.get('EXCEL_FILE_NAME'), session.get('NEW_EXCEL_FILE_NAME'), None, None)
                session.clear()
                flash('Error while converting! Either your pdf file has no tables or no file has been uploaded.')
                flash('Please re-upload your file')
                return redirect(request.url)    
            
            excel_data = pd.ExcelFile(excel_file_path)
            substrings_to_remove = [] # Remove and replace with table name(s) to exclude specific tables from table generation.
            
            # Looping through all sheets to create html tables
            for sheet_name in excel_data.sheet_names:
                try:
                    sheet_data = excel_data.parse(sheet_name, header=1)
                except ValueError:
                    sheet_data = excel_data.parse(sheet_name)

                # Check if sheet_data is a DataFrame and not empty
                if isinstance(sheet_data, pd.DataFrame) and not sheet_data.empty:
                    # Handle the case where sheet_data.columns returns an integer
                    if isinstance(sheet_data.columns, int):
                        # print(f"* Warning: Sheet {sheet_name} has returned an integer for columns. Skipping...")
                        continue
                    
                    # Identifying unnamed columns
                    unnamed_columns = [col for col in sheet_data.columns if 'Unnamed' in str(col)]

                    # Replacing values in unnamed columns with an empty string
                    sheet_data[unnamed_columns] = sheet_data[unnamed_columns].fillna('')

                    # Resetting column names
                    sheet_data.columns = [col if 'Unnamed' not in str(col) else '' for col in sheet_data.columns]

                    # Converting DataFrame to HTML table with CSS styling for column width and removing NaN values
                    html_table = sheet_data.to_html(classes='table table-striped', index=False, na_rep='')
                    html_table = html_table.replace('<table>', '<table style="table-layout: auto; width: 100%;">')
                    html_table = html_table.replace('<th>', '<th style="text-align: left;">')

                    soup = BeautifulSoup(html_table, 'html.parser')

                    # Flag to indicate if the substring is found in the table
                    substring_found = False

                    # Remove tables based on whether the <th> tag contains any of the substrings
                    for th_tag in soup.find_all('th'):
                        if th_tag.string:
                            for substring_to_remove in substrings_to_remove:
                                if substring_to_remove in th_tag.string:
                                    # Marking substring as found
                                    substring_found = True
                                    break

                    # Stop processing current sheet if substring is found
                    if substring_found:
                        continue

                    # Converting the modified html content back to a string
                    modified_html_table = str(soup)
                    html_tables[sheet_name] = modified_html_table

            # Catching selected sheets and checking combine tables variable
            selected_sheets = request.form.getlist('selected_sheets[]')
            index_error = False

            if selected_sheets:
                if 'combineTables' in request.form:
                    checkbox_value = True 

                run_file_check_on(session.get('NEW_EXCEL_FILE_NAME', None))
                session['NEW_EXCEL_FILE_NAME'] = 'custom_' + session.get('EXCEL_FILE_NAME')
                session['NEW_EXCEL_FILE_NAME'] = session['NEW_EXCEL_FILE_NAME'].replace(" ", "_") 
                session['UPLOADS'].append(session['NEW_EXCEL_FILE_NAME'])

                # Only display download notification if there is no index error
                if not index_error:
                    flash(f'{len(selected_sheets)} tables selected')
                    flash('Click Download!')

                print('* Sheets selected')
                print('* Generating...')

                # Writing selected sheets to excel file
                with pd.ExcelWriter(os.path.join(f'{UPLOAD_FOLDER}{session.get("NEW_EXCEL_FILE_NAME")}'), engine='xlsxwriter') as new_excel_data:
                    if checkbox_value:
                        print(f"* Combine tables: {checkbox_value}")
                        combined_sheets = []
                        for sheet_name in selected_sheets:
                            sheet_data = excel_data.parse(sheet_name, header=1)

                            # Check if sheet_data is a DataFrame and not empty
                            if isinstance(sheet_data, pd.DataFrame) and not sheet_data.empty:
                                # Handle the case where sheet_data.columns returns an integer
                                if isinstance(sheet_data.columns, int):
                                    continue                 

                                unnamed_columns = [col for col in sheet_data.columns if 'Unnamed' in str(col)]                       
                                sheet_data[unnamed_columns] = sheet_data[unnamed_columns].fillna('')                       
                                sheet_data.columns = [col if 'Unnamed' not in str(col) else '' for col in sheet_data.columns]
                                combined_sheets.append(sheet_data)

                        print(f"* Index Error: {index_error}")
                        try:
                            combined_data = pd.concat(combined_sheets, ignore_index=True)
                            combined_data.to_excel(new_excel_data, sheet_name='Combined_sheets', index=False)
                        except pd.errors.InvalidIndexError as e:
                            index_error = True
                            print(f"* Index Error: {index_error}")
                            flash("Error while combining tables! Please reselect tables to combine. Ensure tables have the same column headers.")
                            redirect(request.url)
                    else:
                        # Writing individual sheets to Excel file without concatenation
                        print(f"* Combine tables: {checkbox_value}")
                        for sheet_name in selected_sheets:
                            sheet_data = excel_data.parse(sheet_name, header=1)

                            # Check if sheet_data is a DataFrame and not empty
                            if isinstance(sheet_data, pd.DataFrame) and not sheet_data.empty:
                                # Handle the case where sheet_data.columns returns an integer
                                if isinstance(sheet_data.columns, int):
                                    continue                 

                                unnamed_columns = [col for col in sheet_data.columns if 'Unnamed' in str(col)]                       
                                sheet_data[unnamed_columns] = sheet_data[unnamed_columns].fillna('')                       
                                sheet_data.columns = [col if 'Unnamed' not in str(col) else '' for col in sheet_data.columns]
                                sheet_data.to_excel(new_excel_data, sheet_name=sheet_name, index=False)
                
                dir_path = os.getcwd()
                with lock:
                    protect_this_file =  str(PurePath(dir_path, f"{UPLOAD_FOLDER}{session.get('NEW_EXCEL_FILE_NAME')}"))
                    password_protect_excel(protect_this_file, password)

                excel_data.close()
                print(f'* {len(selected_sheets)} sheets successfully generated!')
                print('* Awaiting download request...')
                generated_excel = os.path.join(UPLOAD_FOLDER, session.get("NEW_EXCEL_FILE_NAME"))

                return render_template('index.html', html_tables=html_tables, generated_excel=generated_excel, checkbox_value=checkbox_value, index_error=index_error)

    return render_template('index.html', html_tables=html_tables)


@app.route('/download_excel_file/<filename>', methods=['POST', 'GET'])
def download_excel_file(filename):
    excel_file_path = filename

    # Creating password file 
    session["PASSWORD_FILE"] = f"{session.get('USER_IP')}_password.txt"
    password_file_path = os.path.join(UPLOAD_FOLDER, session.get("PASSWORD_FILE"))
    with open(password_file_path, "w") as password_file:
        password_file.write(password)

    @after_this_request
    def remove_data(response):
        """Function that Flushes session data and files after download completes"""
        
        with lock:
            print('* Locking file resources')
            print('* Download Successful!')
            # Scheduling the removal of files in seperate thread after a short delay to acquire the file resource
            threading.Timer(1.0, remove_files, args=(session.get('PDF_FILE_NAME'), session.get('EXCEL_FILE_NAME'), session.get('NEW_EXCEL_FILE_NAME'), session.get("ZIP_FOLDER"), session.get("PASSWORD_FILE"))).start()
            session.clear()

        return response

    # Creating and preparing zip
    session["ZIP_FOLDER"] = f"{session.get('USER_IP')}_files.zip"
    zip_file_path = os.path.join(UPLOAD_FOLDER, session.get("ZIP_FOLDER"))
    with zipfile.ZipFile(zip_file_path, "w") as zipf:
        zipf.write(excel_file_path)
        zipf.write(password_file_path)
        print("* Files Zipped Successfully")

    zip_response = send_file(zip_file_path, as_attachment=True)
    return zip_response


def do_conversion(file):
    """
    Takes in pdf file, runs check if it has tables on or not. 
    If tables exist excel file will be created and saved for further processing.
    """

    if not file.filename.endswith('.pdf'):
        flash("File extension must be '.pdf' only.")
        threading.Timer(1.0, run_file_check_on, args=(session.get('PDF_FILE_NAME'),))
        threading.Timer(1.0, run_file_check_on, args=(session.get('EXCEL_FILE_NAME'),))
        threading.Timer(1.0, run_file_check_on, args=(session.get('NEW_EXCEL_FILE_NAME'),))

        return redirect(request.url)
    
    run_file_check_on(session.get('PDF_FILE_NAME', None))

    session['PDF_FILE_NAME'] = file.filename
    session['UPLOADS'].append(session['PDF_FILE_NAME'])
    file.save(os.path.join(UPLOAD_FOLDER, session['PDF_FILE_NAME']))
    
    print(f"* File '{session['PDF_FILE_NAME']}' has been saved!")
    print('* Converting...')

    # Excel conversion
    # Getting Files
    extracted_file_path = os.path.join(UPLOAD_FOLDER, session['PDF_FILE_NAME'])
    pdf_file = extracted_file_path
    excel_file = extracted_file_path.replace('.pdf', '.xlsx')

    # Reading tables
    tables = tabula.read_pdf(pdf_file, pages='all')

    # Writing tables to excel file
    if tables:        
        print('* Tables exist: True')
        excel_writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        for i, table in enumerate(tables):
            table.to_excel(excel_writer, sheet_name=f'Sheet_{i+1}', index=False)

        excel_writer.close()

        run_file_check_on(session.get('EXCEL_FILE_NAME', None))
        
        session['EXCEL_FILE_NAME'] = excel_file.replace(UPLOAD_FOLDER, '')
        session['UPLOADS'].append(session['EXCEL_FILE_NAME'])
        print('* Conversion Successful!')
        print(f"* File '{session["EXCEL_FILE_NAME"]}' contents displayed successfully!")
        flash('Conversion completed!')
        flash("Use the check boxes to select the tables that you want to generate for your Excel file. Then click 'Generate Excel' at the bottom of the page ")
    else:
        print('* Tables exist: False')
        run_file_check_on(session.get('EXCEL_FILE_NAME', None))

    redirect(request.url)


def remove_files(pdf_file_name, excel_file_name, new_excel_file_name, zip_folder, password_file):
    """Deleting files from file handler dir. Only if files exist within the dir"""

    with lock: 
        print('* Locking file resources')
        print('* Rinsing file handler')
        # PDF
        if pdf_file_name == None:
            print('* No Pdf file found')
        elif os.path.exists(os.path.join(UPLOAD_FOLDER, pdf_file_name)):
            os.remove(os.path.join(UPLOAD_FOLDER, pdf_file_name))
            print('* Pdf file found')
            print('* Removing pdf file...')

        # EXCEL
        if excel_file_name == None:
            print('* No Excel file found')
        elif os.path.exists(os.path.join(UPLOAD_FOLDER, excel_file_name)):
            os.remove(os.path.join(UPLOAD_FOLDER, excel_file_name))
            print('* Excel file found')
            print('* Removing Excel file...')

        # NEW EXCEL
        if new_excel_file_name == None:
            print('* No New Excel file found')
        elif os.path.exists(os.path.join(UPLOAD_FOLDER, new_excel_file_name)):
            os.remove(os.path.join(UPLOAD_FOLDER, new_excel_file_name))
            print('* New Excel file found')
            print('* Removing New Excel file...')

        # ZIP FOLDER
        if zip_folder == None:
            print('* No Zip folder found')
        elif os.path.exists(os.path.join(UPLOAD_FOLDER, zip_folder)):
            os.remove(os.path.join(UPLOAD_FOLDER, zip_folder))
            print('* Zip folder found')
            print('* Removing Zip folder...')

        # PASSWORD FILE
        if password_file == None:
            print('* No Password file found')
        elif os.path.exists(os.path.join(UPLOAD_FOLDER, password_file)):
            os.remove(os.path.join(UPLOAD_FOLDER, password_file))
            print('* Password file found')
            print('* Removing Password file...')


def run_file_check_on(file_name):
    """Check if a file already exists"""
    if file_name is not None:
        os.remove(os.path.join(UPLOAD_FOLDER, file_name))


def password_protect_excel(file_dir_path, password):
    """Password protect file"""
    pythoncom.CoInitialize()
    xl_file = EnsureDispatch("Excel.Application")
    wb = xl_file.Workbooks.Open(file_dir_path)
    xl_file.DisplayAlerts = False
    wb.Visible = False
    wb.SaveAs(file_dir_path, Password=password)
    wb.Close()
    xl_file.Quit()
    print("* File protected successfully")
    pythoncom.CoUninitialize()

if __name__ == '__main__':
    app.run(debug=True)
from flask import Flask, flash, send_file, after_this_request, render_template
from bs4 import BeautifulSoup
from pathlib import PurePath
from functions import *
import secrets, zipfile


app = Flask(__name__)
app.config['SECRET_KEY'] = secrets.token_hex()
password = secrets.token_hex()


# Main page
@app.route('/', methods=['GET', 'POST'])
def index():
    html_tables = {}
    combine_tables = False
    encrypt_workbook = False
    encrypt_worksheets = False
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
                # Check for Combine Tables checkbox
                if 'combineTables' in request.form:
                    combine_tables = True 

                # Check for Encrypt Workbook checkbox
                if 'encryptWorkbook' in request.form:
                    encrypt_workbook = True
                print(f"* Encrypt Workbook: {encrypt_workbook}")

                # Check for Encrypt Worksheets checkbox
                if 'encryptSheets' in request.form:
                    encrypt_worksheets = True
                print(f"* Encrypt Worksheets: {encrypt_worksheets}")

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
                    if combine_tables:
                        print(f"* Combine tables: {combine_tables}")
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
                        print(f"* Combine tables: {combine_tables}")
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
                protect_this_file =  str(PurePath(dir_path, f"{UPLOAD_FOLDER}{session.get('NEW_EXCEL_FILE_NAME')}"))
                # If user wants to encrypt the Excel Workbook
                if encrypt_workbook:
                    with lock:
                        password_protect_excel(protect_this_file, password)
                # If user wants to encrypt the sheets within the excel workbook
                if encrypt_worksheets:
                    with lock:
                        password_protect_sheets(protect_this_file, password)

                excel_data.close()
                print(f'* {len(selected_sheets)} sheets successfully generated!')
                print('* Awaiting download request...')
                generated_excel = os.path.join(UPLOAD_FOLDER, session.get("NEW_EXCEL_FILE_NAME"))

                return render_template('index.html', html_tables=html_tables, generated_excel=generated_excel, combine_tables=combine_tables, encrypt_workbook=encrypt_workbook, encrypt_worksheets=encrypt_worksheets, index_error=index_error)

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


if __name__ == '__main__':
    app.run(debug=True)
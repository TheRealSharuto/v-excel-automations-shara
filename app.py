from flask import Flask, render_template, request, send_file
import pandas as pd
from werkzeug.utils import secure_filename
import zipfile
from io import BytesIO

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        # Get the file from the form
        excel_file = request.files["file"]
        rows_per_sheet = int(request.form["rows"])
        sheet_name = request.form["sheet_name"]

        # Read the Excel file in memory
        df = pd.read_excel(excel_file)
        # Remove leading and trailing whitespace from column names
        df.columns = df.columns.str.strip()
        sheets = [df[i:i + rows_per_sheet] for i in range(0, df.shape[0], rows_per_sheet)]

        # Create a zip file in memory
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for i, sheet in enumerate(sheets):
                output_buffer = BytesIO()
                sheet.to_excel(output_buffer, index=False, header=True)
                output_buffer.seek(0)
                zipf.writestr(f'{sheet_name}{i+1}.xlsx', output_buffer.read())

        zip_buffer.seek(0)
        return send_file(zip_buffer, as_attachment=True, download_name=f"{sheet_name}.zip", mimetype='application/zip')
    
    return render_template("index.html")


@app.route('/excel-data-extractor', methods=["GET", "POST"])
def extractor():
    if request.method == "POST":
        ex_excel_file = request.files["file"]
        
        # Validate file type
        if not ex_excel_file.filename.endswith('.xlsx'):
            return "Invalid file type", 400
            
        col_name = request.form["col_name"]
        col_value = request.form["col_value"]
        workbook_name = request.form["ext_workbook_name"]

        # Read the uploaded Excel file into a BytesIO object
        file_bytes = BytesIO(ex_excel_file.read())
        
        # Extract from the Excel sheet
        df = pd.read_excel(file_bytes)
        # Remove leading and trailing whitespace from column names
        df.columns = df.columns.str.strip()

        if col_value != "0":
            # Filter rows where the column matches the given value
            extracted_df = df[df[col_name] == col_value]
            output = BytesIO()
            extracted_df.to_excel(output, index=False, header=True, engine='openpyxl')
            output.seek(0)  # Reset pointer for sending file
            return send_file(output, as_attachment=True, download_name=f"{workbook_name}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        elif col_value == "0":
            # Create a BytesIO object for the zip file
            zip_buffer = BytesIO()

            with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                unique_values = df[col_name].dropna().unique()
                for value in unique_values:
                    filtered_df = df[df[col_name] == value]
                    output = BytesIO()
                    filtered_df.to_excel(output, index=False, header=True, engine='openpyxl')
                    output.seek(0)  # Reset pointer for adding to zip
                    # Write the in-memory Excel file to the zip
                    zipf.writestr(f'{col_name}_{value}.xlsx', output.getvalue())

            zip_buffer.seek(0)  # Reset pointer for sending zip file
            return send_file(zip_buffer, as_attachment=True, download_name=f"{col_name}_extracted.zip", mimetype='application/zip')

        else:
            # Filter rows where the column is blank
            extracted_df = df[df[col_name].isna()]
            output = BytesIO()
            extracted_df.to_excel(output, index=False, header=True, engine='openpyxl')
            output.seek(0)  # Reset pointer for sending file
            return send_file(output, as_attachment=True, download_name=f"{workbook_name}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    return render_template('excel-data-extractor.html')

@app.route('/excel-column-puller', methods=["GET", "POST"])
def excel_column_puller():
    if request.method == "POST":
        files = request.files.getlist("files")
        col_name = request.form["col_name"]
        new_col_name = request.form["new_col_name"]

        # Create a BytesIO object to store all the new workbooks in a zip file
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            # Process each uploaded file
            for file in files:
                file_bytes = BytesIO(file.read())

                # Read the Excel file and extract the specified column
                df = pd.read_excel(file_bytes)

                # Remove leading and trailing whitespace from column names
                df.columns = df.columns.str.strip()

                if col_name in df.columns:
                    if new_col_name:
                        extracted_df = df[[col_name]].rename(columns={col_name: new_col_name})
                    else:
                        extracted_df = df[[col_name]]

                    # Create a new workbook name based on the column name and original workbook name
                    original_filename = secure_filename(file.filename)
                    new_workbook_name = f"{new_col_name}-{original_filename}" if new_col_name else f"{col_name}-{original_filename}"

                    # Create a BytesIO object for the new workbook
                    output = BytesIO()
                    extracted_df.to_excel(output, index=False, engine='openpyxl')
                    output.seek(0)  # Reset pointer for adding to zip

                    # Add the new workbook to the zip file
                    zipf.writestr(new_workbook_name, output.getvalue())
                else:
                    return f"Column '{col_name}' not found in file '{file.filename}'", 400

        zip_buffer.seek(0)  # Reset pointer for sending zip file
        return send_file(zip_buffer, as_attachment=True, download_name=f"{col_name}_extracted.zip", mimetype='application/zip')
    
    return render_template('excel-column-puller.html')

if __name__ == '__main__':
    app.run(debug=True)
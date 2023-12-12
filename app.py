import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

# Set the upload folder and allowed file extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create the 'uploads' directory if it doesn't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Function to check if the file extension is allowed
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file1' not in request.files or 'file2' not in request.files:
            return redirect(request.url)

        file1 = request.files['file1']
        file2 = request.files['file2']

        if file1.filename == '' or file2.filename == '':
            return redirect(request.url)

        if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename):
            # Save uploaded files
            file1_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file1.filename))
            file2_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file2.filename))
            file1.save(file1_path)
            file2.save(file2_path)

            # Read the uploaded Excel files into DataFrames
            sheet1 = pd.read_excel(file1_path)
            sheet2 = pd.read_excel(file2_path)

            # Merge the two sheets based on the reference number
            merged_data = pd.merge(sheet1, sheet2, on='Reference', how='outer', suffixes=('_Autocab', '_CMAC'))

            # Create a column to indicate if it's a match or not
            merged_data['Match'] = (merged_data['Price_Autocab'] == merged_data['Price_CMAC'])

            # Sort the results by 'Match' column in descending order (mismatch first)
            merged_data.sort_values(by='Match', ascending=False, inplace=True)

            # Calculate the price difference
            merged_data['Price_Difference'] = merged_data['Price_Autocab'] - merged_data['Price_CMAC']

            # Clean the 'Reference' column by replacing non-finite values with NaN
            merged_data['Reference'] = pd.to_numeric(merged_data['Reference'], errors='coerce')

            # Convert the 'Reference' column to integers, dropping NaN values
            merged_data['Reference'] = merged_data['Reference'].astype('Int64')

            # Create a column to indicate which file the price has come from
            merged_data['Description'] = 'Autocab' + merged_data['Match'].apply(lambda x: ' (Match)' if x else ' (Mismatch)') + \
                                        ', CMAC' + merged_data['Match'].apply(lambda x: ' (Match)' if x else ' (Mismatch)')

            # Select only the relevant columns
            result_columns = ['Reference', 'Price_Autocab', 'Price_CMAC', 'Price_Difference', 'Description']

            # Render the results as an HTML report
            return render_template('report.html', results=merged_data[result_columns])

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

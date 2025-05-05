from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from prodNameExtraction import extract_product_name_description

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    if file:
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        output_path = os.path.join(PROCESSED_FOLDER, f"processed_{file.filename}")
        file.save(input_path)

        # Process the file
        df = pd.read_excel(input_path)
        desc_col = 'Product Description'
        if desc_col not in df.columns:
            return f"Column '{desc_col}' not found in the uploaded file", 400

        df['Product Name'] = df[desc_col].apply(lambda x: extract_product_name_description(x)[0])

        # Save the processed file
        df.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
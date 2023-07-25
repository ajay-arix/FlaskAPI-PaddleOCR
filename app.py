from flask import Flask, request, jsonify, render_template, send_file
from paddleocr import PaddleOCR
from PIL import Image
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
import os
import numpy as np
import openpyxl

app = Flask(__name__)

# Load the OCR model
ocr = PaddleOCR()

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/ocr', methods=['POST'])
def ocr_endpoint():
    if request.method == 'POST':
        # Check if the POST request has the 'file' field (name attribute in the form)
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']

        # Check if the file is empty
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        # Check if the file is a PDF
        if file.filename.lower().endswith(('.pdf')):
            # Create the 'tmp' directory if it doesn't exist
            if not os.path.exists('tmp'):
                os.makedirs('tmp')

            # Save the PDF file to a temporary location
            filename = secure_filename(file.filename)
            pdf_path = f'tmp/{filename}'
            print(f"Saving PDF to: {pdf_path}")
            file.save(pdf_path)

            try:
                # Read the PDF and convert to images
                pdf_images = convert_from_path(pdf_path)
            except Exception as e:
                return jsonify({'error': f'Error converting PDF to images: {str(e)}'}), 400

            # Perform OCR on each image and store results in a list
            extracted_text = []
            for page_num, image in enumerate(pdf_images, start=1):
                # Convert PIL image to RGB mode
                image = image.convert("RGB")

                # Convert PIL image to NumPy array
                image_np = np.array(image)

                # Handle channel ordering for PaddleOCR (assuming image is in RGB mode)
                if image_np.shape[2] == 4:
                    # If the image has an alpha channel (RGBA), remove it
                    image_np = image_np[:, :, :3]

                result = ocr.ocr(image_np)

                # Handle the case when result is not None
                if result is not None:
                    # Extract the recognized text
                    page_text = [line[1] for line in result[0]]

                    # Store OCR results for the page
                    extracted_text.append({'Page': page_num, 'Text': page_text})

            # Do not remove the temporary PDF file at this stage

        else:
            # The file is an image
            try:
                image = Image.open(file)

                # Convert PIL image to RGB mode
                image = image.convert("RGB")

                # Convert PIL image to NumPy array
                image_np = np.array(image)

                # Handle channel ordering for PaddleOCR (assuming image is in RGB mode)
                if image_np.shape[2] == 4:
                    # If the image has an alpha channel (RGBA), remove it
                    image_np = image_np[:, :, :3]

                # Perform OCR using PaddleOCR
                result = ocr.ocr(image_np)

                # Handle the case when result is None
                if result is None:
                    extracted_text = []
                else:
                    # Extract the recognized text
                    extracted_text = [line[1] for line in result[0]]

            except Exception as e:
                return jsonify({'error': f'Error processing the image: {str(e)}'}), 400

        # Export OCR results to Excel
        excel_filename = f'tmp/{filename.replace(".pdf", ".xlsx")}'  # Use xlsx extension for Excel file
        export_to_excel(extracted_text, excel_filename)

        # Render the 'result.html' template with the extracted_text data
        download_link = f'/download/{filename.replace(".pdf", ".xlsx")}'  # Use xlsx extension for download link
        return render_template('result.html', extracted_text=extracted_text, download_link=download_link)

    else:
        # Return a response with instructions for using the API
        response = {
            'message': 'This endpoint only accepts POST requests. '
                       'Please use POST method to upload the file and perform OCR.'
        }
        return jsonify(response), 405

def export_to_excel(extracted_text, filename):
    # Create the 'tmp' directory if it doesn't exist
    if not os.path.exists('tmp'):
        os.makedirs('tmp')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Page', 'Text'])

    for page_data in extracted_text:
        page_num = page_data['Page']
        text = page_data.get('Text', '')
        
        # Convert 'text' to a string if it's a tuple
        if isinstance(text, tuple):
            text = '\n'.join(text)
        else:
            text = str(text)

        ws.append([page_num, text])

    wb.save(filename)

@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    excel_file = f'tmp/{filename}'
    return send_file(excel_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

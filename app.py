from io import BytesIO  
from flask import Flask, flash, redirect, render_template, request, url_for, send_file
from flask_wtf import FlaskForm
from flask_wtf.file import FileField
from werkzeug.utils import secure_filename
from werkzeug.datastructures import FileStorage
from injection_surplus import process_surplus
from injection_totale import process_total
from pdf_converter import convert_to_image
from datetime import datetime
from traceback import print_last
from docx import Document
from docx.table import Table
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT,WD_COLOR_INDEX
from flask import Flask, flash
import chardet
import re
import subprocess
from subprocess import run, PIPE
import logging
import os
import imghdr


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    # Get the uploaded DOC file
    doc_file = request.files['doc_file']
    image = request.files['schema_unifulaire']
    tiers = request.form['tiers']
    tiers_name = request.form['tiers_name']
    tiers_email = request.form['tiers_email']
    surplus = False
    schema_unifulaire = 'schema_unifulaire_1.png'


    # Check if the file is a PDF
    if image and image.filename.lower().endswith('.pdf'):
        schema_unifulaire_path = '/tmp/uploaded_file.pdf'
        image.save(schema_unifulaire_path)
        schema_unifulaire = convert_to_image(schema_unifulaire_path)

    # Check if the file is an image (JPEG or PNG)
    image_type = imghdr.what(image)
    if image_type in ['jpeg', 'png']:
        image.save(image.filename)
        schema_unifulaire = image.filename
        # return 'JPEG' if image_type == 'jpeg' else 'PNG'

        
       



    # Save the DOC file temporarily
    temp_doc_path = 'temp_input.doc'
    doc_file.save(temp_doc_path)

    # Define the output DOCX file path
    output_docx_path = 'converted_output.docx'

    # Convert DOC to DOCX using unoconv
    try:
        subprocess.run(['unoconv', '-f', 'docx', '-o', output_docx_path, temp_doc_path], check=True)
        print(f"Conversion successful: {temp_doc_path} -> {output_docx_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error during conversion: {e}")
        return "Conversion failed. Please check the logs for details."

    # Clean up: Remove temporary DOC file
    os.remove(temp_doc_path)

    doc = Document(output_docx_path)
    for p in doc.paragraphs:
        
        if "soutirée sur le PDL d’Injection" in p.text:
            print('SURPLUS FOUND')
            surplus = True



    if surplus == True:
        processed_doc = process_surplus(doc,tiers,tiers_name,tiers_email,schema_unifulaire)

    if surplus == False:
        processed_doc = process_total(doc, tiers, tiers_name, tiers_email,schema_unifulaire)
    

    # Redirect to a page displaying the download link or perform other actions as needed
    return redirect(url_for('download', filename=processed_doc))

@app.route('/download/<filename>')
def download(filename):
    # Serve the converted DOCX file for download
    return send_file(filename, as_attachment=True)

    # # Serve the converted DOCX file for download
    # return redirect(url_for('static', filename=filename, _external=True))

if __name__ == '__main__':
    app.run(debug=True)

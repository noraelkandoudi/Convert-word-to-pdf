from flask import Flask, render_template, request, send_file, redirect, url_for
import os
from werkzeug.utils import secure_filename
import pdfkit
from docx import Document

app = Flask(__name__)

# Cartella dove salveremo i file caricati e convertiti
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Pagina iniziale
@app.route('/')
def index():
    return render_template('index.html')

# Caricamento del file
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(url_for('index'))
    
    file = request.files['file']
    if file.filename == '':
        return redirect(url_for('index'))
    
    if file and file.filename.endswith('.docx'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        return render_template('convert.html', filename=filename)

    return redirect(url_for('index'))

# Conversione del file
@app.route('/convert/<filename>', methods=['POST'])
def convert_file(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    output_filepath = filepath.replace('.docx', '.pdf')

    # Converti il file Word in HTML
    doc = Document(filepath)
    html_content = ''
    for para in doc.paragraphs:
        html_content += f'<p>{para.text}</p>'

    # Salva l'HTML in un file temporaneo
    html_filepath = filepath.replace('.docx', '.html')
    with open(html_filepath, 'w') as f:
        f.write(html_content)

    # Converti l'HTML in PDF usando pdfkit
    pdfkit.from_file(html_filepath, output_filepath)

    # Ritorna la pagina di download
    return render_template('download.html', filename=output_filepath)

# Scaricare il file
@app.route('/download/<path:filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True,host="0.0.0.0",port=0000)

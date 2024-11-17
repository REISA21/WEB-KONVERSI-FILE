from flask import Flask, render_template, request, jsonify, send_from_directory
import os
import logging
import traceback
from flask_cors import CORS
from pdf2docx import Converter
from docx2pdf import convert
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import A4, letter  # Tambahkan letter di sini
from reportlab.lib import colors
from pptx import Presentation
import comtypes.client  # Required for PowerPoint to PDF conversion on Windows
from io import BytesIO
import comtypes.client
from comtypes import CoInitialize, CoUninitialize
from reportlab.pdfgen import canvas  # Pastikan ini diimpor
from PIL import Image, ImageDraw  # Pastikan ini diimpor

app = Flask(__name__)
CORS(app)

# Enable detailed logging
logging.basicConfig(level=logging.DEBUG)

# Ensure the temp folder exists
if not os.path.exists("temp"):
    os.makedirs("temp")

# Route for the main page
@app.route('/')
def index():
    return render_template('index.html')

# Route for PDF to Word page
@app.route('/page_pdf_to_word.html')
def pdf_to_word_page():
    return render_template('page_pdf_to_word.html')

# Route for Word to PDF page
@app.route('/page_word_to_pdf.html')
def word_to_pdf_page():
    return render_template('page_word_to_pdf.html')

# Route for PPT to PDF page
@app.route('/page_ppt_to_pdf.html')
def ppt_to_pdf_page():
    return render_template('page_ppt_to_pdf.html')

# Route for Explore Files page
@app.route('/explore.html')
def explore():
    return render_template('explore.html')

# Route for other conversion pages
@app.route('/page_excel_to_pdf.html')
def excel_to_pdf_page():
    return render_template('page_excel_to_pdf.html')

@app.route('/page_pdf_to_excel.html')
def pdf_to_excel_page():
    return render_template('page_pdf_to_excel.html')

@app.route('/page_pdf_to_ppt.html')
def pdf_to_ppt_page():
    return render_template('page_pdf_to_ppt.html')

@app.route('/page_word_to_txt.html')
def word_to_txt_page():
    return render_template('page_word_to_txt.html')

@app.route('/page_txt_to_word.html')
def txt_to_word_page():
    return render_template('page_txt_to_word.html')

@app.route('/page_txt_to_pdf.html')
def txt_to_pdf_page():
    return render_template('page_txt_to_pdf.html')

@app.route('/page_pdf_to_txt.html')
def pdf_to_txt_page():
    return render_template('page_pdf_to_txt.html')

@app.route('/page_jpg_to_png.html')
def jpg_to_png_page():
    return render_template('page_jpg_to_png.html')

@app.route('/page_png_to_jpg.html')
def png_to_jpg_page():
    return render_template('page_png_to_jpg.html')

@app.route('/page_heic_to_jpg.html')
def heic_to_jpg_page():
    return render_template('page_heic_to_jpg.html')

@app.route('/page_epub_to_pdf.html')
def epub_to_pdf_page():
    return render_template('page_epub_to_pdf.html')

@app.route('/page_pdf_to_epub.html')
def pdf_to_epub_page():
    return render_template('page_pdf_to_epub.html')

@app.route('/page_rar_to_zip.html')
def rar_to_zip_page():
    return render_template('rar_to_zip.html')

@app.route('/page_excel_to_csv.html')
def excel_to_csv_page():
    return render_template('page_excel_to_csv.html')

@app.route('/page_csv_to_excel.html')
def csv_to_excel_page():
    return render_template('page_csv_to_excel.html')

@app.route('/page_json_to_csv.html')
def json_to_csv_page():
    return render_template('json_to_csv.html')

@app.route('/page_json_to_xml.html')
def json_to_xml_page():
    return render_template('json_to_xml.html')

# Endpoint to upload PDF and convert to Word
@app.route('/upload', methods=['POST'])
def upload_file():
    app.logger.info("Upload endpoint hit")
    
    if 'pdf_file' not in request.files:
        app.logger.error("No file part in the request")
        return jsonify(success=False, message="No file part"), 400
    
    file = request.files['pdf_file']
    if file.filename == '':
        app.logger.error("No selected file")
        return jsonify(success=False, message="No selected file"), 400

    app.logger.info(f"File received: {file.filename}")
    if file and file.filename.endswith('.pdf'):
        try:
            pdf_path = os.path.join("temp", file.filename)
            file.save(pdf_path)
            app.logger.info(f"File saved to {pdf_path}")
            
            docx_path = pdf_path.replace('.pdf', '.docx')
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            app.logger.info(f"Conversion successful: {docx_path}")
            
            return jsonify(success=True, message="File converted successfully", download_url=f"/download/{os.path.basename(docx_path)}")
        
        except Exception as e:
            app.logger.error("Error during file processing")
            app.logger.error(traceback.format_exc())
            return jsonify(success=False, message=f"Error during conversion: {str(e)}"), 500
    else:
        app.logger.error("Invalid file format")
        return jsonify(success=False, message="Invalid file format"), 400

# Endpoint for Word to PDF conversion
@app.route('/upload_word_to_pdf', methods=['POST'])
def upload_word_to_pdf():
    app.logger.info("Word to PDF upload endpoint hit")
    
    if 'word_file' not in request.files:
        app.logger.error("No file part in the request")
        return jsonify(success=False, message="No file part"), 400
    
    file = request.files['word_file']
    if file.filename == '':
        app.logger.error("No selected file")
        return jsonify(success=False, message="No selected file"), 400

    if file and file.filename.endswith('.docx'):
        try:
            word_path = os.path.join("temp", file.filename)
            file.save(word_path)
            app.logger.info(f"File saved to {word_path}")

            pdf_path = word_path.replace('.docx', '.pdf')
            convert(word_path, pdf_path)
            app.logger.info(f"Conversion successful: {pdf_path}")
            
            return jsonify(success=True, message="File converted successfully", download_url=f"/download/{os.path.basename(pdf_path)}")
        
        except Exception as e:
            app.logger.error("Error during file processing")
            app.logger.error(traceback.format_exc())
            return jsonify(success=False, message=f"Error during conversion: {str(e)}"), 500
    else:
        app.logger.error("Invalid file format")
        return jsonify(success=False, message="Invalid file format"), 400


# Endpoint for Excel to PDF conversion
@app.route('/upload_excel_to_pdf', methods=['POST'])
def upload_excel_to_pdf():
    app.logger.info("Excel to PDF upload endpoint hit")
    
    if 'excel_file' not in request.files:
        app.logger.error("No excel_file part in the request")
        return jsonify(success=False, message="No file part"), 400

    file = request.files['excel_file']
    if file.filename == '':
        app.logger.error("No selected file")
        return jsonify(success=False, message="No selected file"), 400

    if file and file.filename.endswith('.xlsx'):
        try:
            excel_path = os.path.join("temp", file.filename)
            file.save(excel_path)
            app.logger.info(f"File saved to {excel_path}")
            
            pdf_path = excel_path.replace('.xlsx', '.pdf')
            df = pd.read_excel(excel_path, header=None).fillna('')
            
            doc = SimpleDocTemplate(pdf_path, pagesize=A4)
            elements = []

            data = [df.columns.values.tolist()] + df.values.tolist()
            table = Table(data, colWidths=[70] * len(df.columns))
            
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ])
            table.setStyle(style)
            elements.append(table)
            doc.build(elements)
            
            return jsonify(success=True, download_url=f"/download/{os.path.basename(pdf_path)}")
        
        except Exception as e:
            return jsonify(success=False, message=f"Error during conversion: {str(e)}"), 500
    else:
        return jsonify(success=False, message="Invalid file format, please upload an .xlsx file"), 400


# Fungsi konversi PPT ke PDF
def convert_ppt_to_pdf(ppt_path, pdf_path):
    prs = Presentation(ppt_path)
    pdf_canvas = canvas.Canvas(pdf_path)
    pdf_width, pdf_height = 595, 842  # Ukuran standar halaman A4 dalam poin

    for slide_num, slide in enumerate(prs.slides):
        try:
            # Mendapatkan ukuran slide
            width = prs.slide_width
            height = prs.slide_height
            print(f"Memproses slide {slide_num + 1} dengan lebar {width} dan tinggi {height}")

            # Skala untuk menyesuaikan slide dengan halaman PDF
            scale = min(pdf_width / width, pdf_height / height)
            img_width = int(width * scale)
            img_height = int(height * scale)

            # Membuat gambar kosong untuk slide
            img = Image.new("RGB", (img_width, img_height), (255, 255, 255))
            draw = ImageDraw.Draw(img)
            draw.rectangle([(0, 0), (img_width, img_height)], fill="white")
            draw.text((10, 10), f"Slide {slide_num + 1}", fill="black")  # Placeholder teks

            # Simpan gambar ke buffer byte
            img_buffer = BytesIO()
            img.save(img_buffer, format="PNG")
            img_buffer.seek(0)

            # Menempatkan gambar di halaman PDF
            pdf_canvas.drawImage(img_buffer, 0, 0, width=img_width, height=img_height)
            pdf_canvas.showPage()
            img_buffer.close()

        except Exception as e:
            print(f"Kesalahan selama pemrosesan slide: {str(e)}")
            continue

    pdf_canvas.save()
    print("Pembuatan PDF selesai dengan sukses.")

# Endpoint untuk mengunggah dan mengonversi PPT ke PDF
@app.route('/upload_ppt_to_pdf', methods=['POST'])
def upload_ppt_to_pdf():
    if 'ppt_file' not in request.files:
        app.logger.error("No file part in the request")
        return jsonify(success=False, message="No file part"), 400

    file = request.files['ppt_file']
    if file.filename == '':
        app.logger.error("No selected file")
        return jsonify(success=False, message="No selected file"), 400

    if file and file.filename.endswith(('.ppt', '.pptx')):
        try:
            # Tentukan jalur file PPT dan PDF menggunakan jalur absolut
            ppt_path = os.path.abspath(os.path.join("temp", file.filename))
            file.save(ppt_path)
            pdf_path = os.path.abspath(ppt_path.replace('.pptx', '.pdf').replace('.ppt', '.pdf'))

            # Log jalur untuk memastikan jalur file sudah benar
            app.logger.info(f"File PPT disimpan di: {ppt_path}")
            app.logger.info(f"PDF akan disimpan di: {pdf_path}")

            # Inisialisasi COM
            CoInitialize()
            try:
                # Buka PowerPoint dan konversi
                powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
                powerpoint.Visible = 1

                presentation = powerpoint.Presentations.Open(ppt_path)
                presentation.SaveAs(pdf_path, 32)  # 32 adalah format PDF
                presentation.Close()
                powerpoint.Quit()

                app.logger.info(f"Konversi berhasil, PDF disimpan di {pdf_path}")
                return jsonify(success=True, download_url=f"/download/{os.path.basename(pdf_path)}")
            finally:
                CoUninitialize()

        except Exception as e:
            app.logger.error("Kesalahan saat mengonversi file")
            app.logger.error(traceback.format_exc())
            return jsonify(success=False, message=f"File conversion failed: {str(e)}"), 500

    else:
        app.logger.error("Format file tidak valid")
        return jsonify(success=False, message="Format file tidak valid, harap unggah file .ppt atau .pptx"), 400
# Endpoint untuk mengunduh file yang dikonversi
@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        return send_from_directory('temp', filename, as_attachment=True)
    except FileNotFoundError:
        logging.error("File not found for download")
        return jsonify(success=False, message="File not found"), 404

if __name__ == '__main__':
    app.run(debug=True)
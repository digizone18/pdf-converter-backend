from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile
import subprocess
import json

# Library konversi
import PyPDF2
from docx import Document
from openpyxl import Workbook, load_workbook
import pdf2docx
import tabula
import pandas as pd
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib import colors
import io

app = Flask(__name__)
CORS(app)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.config['UPLOAD_FOLDER'] = '/tmp'

# ============ PDF TO WORD ============
@app.route('/convert/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Empty filename'}), 400
    
    # Save temp file
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    file.save(temp_input.name)
    temp_input.close()
    
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    temp_output.close()
    
    try:
        # Method 1: pake pdf2docx (lebih akurat)
        try:
            converter = pdf2docx.Converter(temp_input.name)
            converter.convert(temp_output.name, start=0, end=None)
            converter.close()
            result_file = temp_output.name
        except:
            # Method 2: fallback ke PyPDF2 (cuma teks)
            pdf_reader = PyPDF2.PdfReader(temp_input.name)
            doc = Document()
            
            for page in pdf_reader.pages:
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)
                doc.add_page_break()
            
            doc.save(temp_output.name)
            result_file = temp_output.name
        
        return send_file(
            result_file,
            as_attachment=True,
            download_name='output.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(temp_input.name)
        try:
            os.unlink(temp_output.name)
        except:
            pass

# ============ PDF TO EXCEL ============
@app.route('/convert/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    file.save(temp_input.name)
    temp_input.close()
    
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    temp_output.close()
    
    try:
        # Ekstrak tabel dengan tabula
        try:
            dfs = tabula.read_pdf(temp_input.name, pages='all', multiple_tables=True)
            
            if not dfs:
                # Fallback: ekstrak teks biasa
                pdf_reader = PyPDF2.PdfReader(temp_input.name)
                text = ''
                for page in pdf_reader.pages:
                    text += page.extract_text()
                
                # Ubah teks ke excel
                wb = Workbook()
                ws = wb.active
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if line.strip():
                        ws.cell(row=i+1, column=1, value=line.strip())
            else:
                # Buat Excel dari tabel
                wb = Workbook()
                for i, df in enumerate(dfs):
                    if i == 0:
                        ws = wb.active
                        ws.title = f'Table_{i+1}'
                    else:
                        ws = wb.create_sheet(f'Table_{i+1}')
                    
                    # Tulis header & data
                    for col, header in enumerate(df.columns, 1):
                        ws.cell(row=1, column=col, value=str(header))
                    
                    for row_idx, row in enumerate(df.values, 2):
                        for col_idx, val in enumerate(row, 1):
                            ws.cell(row=row_idx, column=col_idx, value=str(val) if pd.notna(val) else '')
        except:
            # Ultra fallback
            wb = Workbook()
            ws = wb.active
            ws.cell(row=1, column=1, value='Terjadi error ekstraksi, silakan coba file lain')
        
        wb.save(temp_output.name)
        
        return send_file(
            temp_output.name,
            as_attachment=True,
            download_name='output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(temp_input.name)
        try:
            os.unlink(temp_output.name)
        except:
            pass

# ============ WORD TO PDF ============
@app.route('/convert/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    file.save(temp_input.name)
    temp_input.close()
    
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_output.close()
    
    try:
        # Baca DOCX
        doc = Document(temp_input.name)
        
        # Buat PDF dengan ReportLab
        doc_pdf = SimpleDocTemplate(temp_output.name, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        
        # Tambah semua paragraf
        for para in doc.paragraphs:
            if para.text.strip():
                story.append(Paragraph(para.text, styles['Normal']))
                story.append(Spacer(1, 12))
        
        # Tambah tabel kalo ada
        for table in doc.tables:
            data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                data.append(row_data)
            
            if data:
                from reportlab.platypus import Table, TableStyle
                t = Table(data)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 14),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(t)
                story.append(Spacer(1, 20))
        
        doc_pdf.build(story)
        
        return send_file(
            temp_output.name,
            as_attachment=True,
            download_name='output.pdf',
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(temp_input.name)
        try:
            os.unlink(temp_output.name)
        except:
            pass

# ============ EXCEL TO PDF ============
@app.route('/convert/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    file.save(temp_input.name)
    temp_input.close()
    
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_output.close()
    
    try:
        # Baca Excel
        wb = load_workbook(temp_input.name)
        ws = wb.active
        
        # Buat PDF
        doc_pdf = SimpleDocTemplate(temp_output.name, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        
        # Ambil semua data
        data = []
        for row in ws.iter_rows(values_only=True):
            row_data = [str(cell) if cell is not None else '' for cell in row]
            if any(row_data):
                data.append(row_data)
        
        # Buat tabel
        if data:
            from reportlab.platypus import Table, TableStyle
            t = Table(data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(t)
        
        doc_pdf.build(story)
        
        return send_file(
            temp_output.name,
            as_attachment=True,
            download_name='output.pdf',
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(temp_input.name)
        try:
            os.unlink(temp_output.name)
        except:
            pass

# ============ HEALTH CHECK ============
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'message': 'Server is running!'})

@app.route('/', methods=['GET'])
def index():
    return jsonify({
        'name': 'PDF Converter API',
        'version': '1.0',
        'endpoints': [
            '/convert/pdf-to-word',
            '/convert/pdf-to-excel',
            '/convert/word-to-pdf',
            '/convert/excel-to-pdf'
        ]
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
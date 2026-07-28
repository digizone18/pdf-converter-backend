from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile
import PyPDF2
from docx import Document
from openpyxl import Workbook, load_workbook
import pdf2docx
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

app = Flask(__name__)
CORS(app)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# ============ PDF TO WORD ============
@app.route('/convert/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    file.save(temp_input.name)
    temp_input.close()
    
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    temp_output.close()
    
    try:
        pdf_reader = PyPDF2.PdfReader(temp_input.name)
        doc = Document()
        
        for page in pdf_reader.pages:
            text = page.extract_text()
            if text:
                doc.add_paragraph(text)
            doc.add_page_break()
        
        doc.save(temp_output.name)
        
        return send_file(
            temp_output.name,
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
        pdf_reader = PyPDF2.PdfReader(temp_input.name)
        text = ''
        for page in pdf_reader.pages:
            text += page.extract_text()
        
        wb = Workbook()
        ws = wb.active
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if line.strip():
                ws.cell(row=i+1, column=1, value=line.strip())
        
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
        doc = Document(temp_input.name)
        doc_pdf = SimpleDocTemplate(temp_output.name, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        
        for para in doc.paragraphs:
            if para.text.strip():
                story.append(Paragraph(para.text, styles['Normal']))
                story.append(Spacer(1, 12))
        
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
        wb = load_workbook(temp_input.name)
        ws = wb.active
        
        doc_pdf = SimpleDocTemplate(temp_output.name, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        
        data = []
        for row in ws.iter_rows(values_only=True):
            row_data = [str(cell) if cell is not None else '' for cell in row]
            if any(row_data):
                data.append(row_data)
        
        if data:
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
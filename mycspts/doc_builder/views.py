import os
import re
import json
import uuid
import subprocess
from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from docx import Document

LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

def extract_fields_from_docx(filepath):
    doc = Document(filepath)
    full_text = "\n".join([p.text for p in doc.paragraphs])
    pattern = r"([ก-๙a-zA-Z0-9\s(){}\[\]\-:]+?)\s*\.{3,}"
    matches = re.finditer(pattern, full_text)
    
    fields = []
    seen = set()
    for match in matches:
        label = match.group(1).strip()
        label = re.sub(r'[\.\s:]+$', '', label)
        if label and label not in seen:
            fields.append(label)
            seen.add(label)
    return fields

def fill_docx(template_path, data, output_path):
    doc = Document(template_path)
    for p in doc.paragraphs:
        for i in range(len(p.runs)):
            run_text = p.runs[i].text
            for label, value in data.items():
                if label in p.text:
                    if '...' in run_text:
                        p.runs[i].text = re.sub(r'\.{3,}', value, run_text, count=1)
    doc.save(output_path)

def convert_to_pdf(docx_path, output_dir):
    try:
        command = [
            LIBREOFFICE_PATH,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            docx_path
        ]
        subprocess.run(command, check=True)
        return docx_path.replace('.docx', '.pdf')
    except Exception as e:
        print(f"PDF Conversion Error: {e}")
        return None

def index(request):
    return render(request, 'doc_builder/index.html')

def upload_file(request):
    if request.method == 'POST' and request.FILES.get('file'):
        file = request.FILES['file']
        fs = FileSystemStorage(location=os.path.join(settings.MEDIA_ROOT, 'uploads'))
        filename = fs.save(f"{uuid.uuid4()}_{file.name}", file)
        filepath = fs.path(filename)
        
        fields = extract_fields_from_docx(filepath)
        return JsonResponse({
            'filename': filename,
            'fields': fields
        })
    return JsonResponse({'error': 'No file uploaded'}, status=400)

def generate(request):
    if request.method == 'POST':
        data = json.loads(request.body)
        filename = data.get('filename')
        form_data = data.get('formData')
        
        input_path = os.path.join(settings.MEDIA_ROOT, 'uploads', filename)
        output_dir = os.path.join(settings.MEDIA_ROOT, 'output')
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        output_docx_name = f"filled_{filename}"
        output_docx_path = os.path.join(output_dir, output_docx_name)
        output_json_path = output_docx_path.replace('.docx', '.json')
        
        # Fill Docx
        fill_docx(input_path, form_data, output_docx_path)
        
        # Save JSON
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(form_data, f, ensure_ascii=False, indent=4)
            
        # Convert PDF
        pdf_path = convert_to_pdf(output_docx_path, output_dir)
        
        if pdf_path and os.path.exists(pdf_path):
            return JsonResponse({
                'success': True,
                'pdf_url': f"/media/output/{os.path.basename(pdf_path)}",
                'json_url': f"/media/output/{os.path.basename(output_json_path)}"
            })
    return JsonResponse({'error': 'Failed'}, status=500)

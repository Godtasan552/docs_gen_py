import os
import re
import json
import uuid
import subprocess
from django.shortcuts import render
from django.http import JsonResponse
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from docx import Document

LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

def extract_fields_from_docx(filepath):
    """
    Highly robust field extraction. 
    Handles:
    - Normal dots (.), Ellipsis (…), Underlines (_)
    - Labels on same line or line above
    - Multiple dot groups on one line
    """
    doc = Document(filepath)
    fields = []
    seen = set()
    
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    
    # Regex including Unicode ellipsis \u2026 and dots
    # Looks for any sequence of 2 or more dot-like characters
    dot_pattern = r"[…\._]{2,}"
    
    for i, text in enumerate(paragraphs):
        if re.search(dot_pattern, text):
            # Case 1: Text before dots on the same line
            # Capture labels (supporting Thai, English, Numbers)
            match = re.search(r"([ก-๙a-zA-Z0-9\s(){}\[\]\-:]+?)\s*" + dot_pattern, text)
            if match:
                label = match.group(1).strip().strip(':').strip()
                if label and label not in seen and len(label) < 100:
                    fields.append(label)
                    seen.add(label)
            # Case 2: Dots on a new line, check line above
            elif i > 0:
                prev_text = paragraphs[i-1].strip().strip(':').strip()
                if prev_text and not re.search(dot_pattern, prev_text) and len(prev_text) < 100:
                    if prev_text not in seen:
                        fields.append(prev_text)
                        seen.add(prev_text)
                        
    return fields

def fill_docx(template_path, data, output_path):
    """
    Fills document by replacing dot-like characters with values.
    """
    doc = Document(template_path)
    dot_pattern = r"[…\._]{2,}"
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        for label, value in data.items():
            # Match if label is in this paragraph or previous
            if label in text or (i > 0 and label in doc.paragraphs[i-1].text):
                for run in p.runs:
                    if re.search(dot_pattern, run.text):
                        # Replace dots/ellipsis/underlines with the user's value
                        run.text = re.sub(dot_pattern, value, run.text)
    
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
        # Create media root subdirs if not exist
        upload_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
        if not os.path.exists(upload_dir):
            os.makedirs(upload_dir)
            
        fs = FileSystemStorage(location=upload_dir)
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
        try:
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
            
            fill_docx(input_path, form_data, output_docx_path)
            
            with open(output_json_path, 'w', encoding='utf-8') as f:
                json.dump(form_data, f, ensure_ascii=False, indent=4)
                
            pdf_path = convert_to_pdf(output_docx_path, output_dir)
            
            if pdf_path and os.path.exists(pdf_path):
                return JsonResponse({
                    'success': True,
                    'pdf_url': f"{settings.MEDIA_URL}output/{os.path.basename(pdf_path)}",
                    'json_url': f"{settings.MEDIA_URL}output/{os.path.basename(output_json_path)}"
                })
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)
    return JsonResponse({'error': 'Failed'}, status=500)

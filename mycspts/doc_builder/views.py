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
    """
    Smarter field extraction. Handles cases where dots are:
    1. On the same line as the label (e.g., 'Name.......')
    2. On the line below the label (common in Thai academic forms)
    """
    doc = Document(filepath)
    fields = []
    seen = set()
    
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    
    for i, text in enumerate(paragraphs):
        # Case 1: Label and dots on the same line
        if '...' in text:
            # Look for non-dot text before the dots
            match = re.search(r"([ก-๙a-zA-Z0-9\s(){}\[\]\-:]+?)\s*\.{3,}", text)
            if match:
                label = match.group(1).strip().strip(':')
                if label and label not in seen:
                    fields.append(label)
                    seen.add(label)
            # Case 2: Only dots on this line, label was likely on the line above
            elif i > 0:
                prev_text = paragraphs[i-1].strip().strip(':')
                # If previous line had no dots and was reasonably short (likely a label)
                if prev_text and '...' not in prev_text and len(prev_text) < 100:
                    if prev_text not in seen:
                        fields.append(prev_text)
                        seen.add(prev_text)
                        
    return fields

def fill_docx(template_path, data, output_path):
    """
    Fills the document by replacing dots.
    Optimized for both same-line and next-line patterns.
    """
    doc = Document(template_path)
    
    # Track which fields we have filled to avoid duplicates if desired, 
    # but usually one field replaces multiple lines of dots.
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        for label, value in data.items():
            # If label is in this paragraph or the one above
            found_in_this = label in text
            found_in_prev = False
            if i > 0:
                found_in_prev = label in doc.paragraphs[i-1].text
                
            if found_in_this or found_in_prev:
                # Replace dots in THIS paragraph
                for run in p.runs:
                    if '...' in run.text:
                        # Replace 3 or more dots with the value
                        # We use count=1 to replace only the first occurrence if multiple, 
                        # but usually a paragraph has one semantic dot sequence.
                        run.text = re.sub(r'\.{3,}', value, run.text)
    
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
                    'pdf_url': f"/media/output/{os.path.basename(pdf_path)}",
                    'json_url': f"/media/output/{os.path.basename(output_json_path)}"
                })
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)
    return JsonResponse({'error': 'Failed'}, status=500)

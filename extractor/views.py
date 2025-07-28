# views.py
import json
import traceback
from django.http import JsonResponse
from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt
from django import forms
from io import BytesIO
from .extractor import PDFOCRExtractor

class PDFUploadForm(forms.Form):
    pdf_file = forms.FileField(
        label="Select PDF File",
        widget=forms.ClearableFileInput(attrs={
            'accept': '.pdf',
            'class': 'form-control'
        }),
        help_text="Max file size: 10MB. Only PDF files are supported."
    )

def upload_pdf(request):
    """Handle both GET (show form) and POST (process upload)"""
    
    if request.method == 'GET':
        # Show the upload form
        form = PDFUploadForm()
        return render(request, 'upload.html', {'form': form})
    
    elif request.method == 'POST':
        print("=" * 50)
        print("🔍 DJANGO PDF UPLOAD DEBUG")
        print("=" * 50)
        
        form = PDFUploadForm(request.POST, request.FILES)
        
        if not form.is_valid():
            print("❌ Form validation failed")
            print(f"Form errors: {form.errors}")
            return render(request, 'upload.html', {
                'form': form,
                'error': {
                    'message': 'Form validation failed',
                    'details': str(form.errors)
                }
            })
        
        pdf_file = form.cleaned_data['pdf_file']
        
        # Debug file info
        print(f"📄 File name: {pdf_file.name}")
        print(f"📄 File size: {pdf_file.size}")
        print(f"📄 Content type: {pdf_file.content_type}")
        
        # Validate PDF
        pdf_file.seek(0)
        first_bytes = pdf_file.read(10)
        pdf_file.seek(0)
        
        print(f"📄 First 10 bytes: {first_bytes}")
        print(f"📄 Is valid PDF: {first_bytes.startswith(b'%PDF')}")
        
        if not first_bytes.startswith(b'%PDF'):
            return render(request, 'upload.html', {
                'form': form,
                'error': {
                    'message': 'Invalid PDF file',
                    'details': f'File does not appear to be a valid PDF. File starts with: {first_bytes}'
                }
            })
        
        try:
            # Read file content and create BytesIO
            pdf_file.seek(0)
            pdf_content = pdf_file.read()
            print(f"📄 Read {len(pdf_content)} bytes from uploaded file")
            
            # Create BytesIO object for your extractor
            pdf_file_io = BytesIO(pdf_content)
            
            print("🚀 Starting extraction...")
            
            # Use your extractor
            extractor = PDFOCRExtractor()
            result = extractor.extract(pdf_file_io)
            
            if result.get('error'):
                print(f"❌ Extraction failed: {result.get('error')}")
                return render(request, 'upload.html', {
                    'form': form,
                    'error': {
                        'message': result.get('error'),
                        'details': result.get('details')
                    },
                    'debug': result.get('debug')
                })
            
            print("✅ Extraction successful!")
            print(f"Global fields: {len(result.get('global', {}))}")
            print(f"Items found: {len(result.get('items', []))}")
            
            # Prepare data for template
            template_data = {
                'form': PDFUploadForm(),  # Fresh form for next upload
                'data': {
                    'raw': result,
                    'pretty': json.dumps(result, indent=2)
                },
                'debug': {
                    'stats': {
                        'global_fields': len(result.get('global', {})),
                        'items': len(result.get('items', [])),
                        'components': sum(len(item.get('Components', [])) for item in result.get('items', [])),
                        'ocr_text_length': result.get('debug', {}).get('ocr_text_length', 0)
                    },
                    'processing_steps': result.get('debug', {}).get('processing_steps', [])
                }
            }
            
            return render(request, 'upload.html', template_data)
            
        except Exception as e:
            print(f"❌ Processing exception: {e}")
            traceback.print_exc()
            
            return render(request, 'upload.html', {
                'form': form,
                'error': {
                    'message': 'Processing failed',
                    'details': f'An unexpected error occurred: {str(e)}'
                }
            })

# Keep the API endpoint for AJAX calls if needed
@csrf_exempt
def extract_pdf_data(request):
    """API endpoint for AJAX uploads"""
    if request.method == 'POST':
        if 'pdf_file' not in request.FILES:
            return JsonResponse({
                'error': 'No PDF file uploaded',
                'available_files': list(request.FILES.keys())
            })
        
        pdf_file = request.FILES['pdf_file']
        
        # Validate PDF
        pdf_file.seek(0)
        first_bytes = pdf_file.read(10)
        pdf_file.seek(0)
        
        if not first_bytes.startswith(b'%PDF'):
            return JsonResponse({
                'error': 'Invalid PDF file',
                'details': f'File starts with: {first_bytes}'
            })
        
        try:
            # Process the file
            pdf_file.seek(0)
            pdf_content = pdf_file.read()
            pdf_file_io = BytesIO(pdf_content)
            
            extractor = PDFOCRExtractor()
            result = extractor.extract(pdf_file_io)
            
            if result.get('error'):
                return JsonResponse({
                    'error': result.get('error'),
                    'details': result.get('details')
                })
            
            return JsonResponse({
                'success': True,
                'data': result
            })
            
        except Exception as e:
            return JsonResponse({
                'error': 'Processing failed',
                'details': str(e)
            })
    
    return JsonResponse({'error': 'Only POST method allowed'})
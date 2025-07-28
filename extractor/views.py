from django.shortcuts import render
from django.views.decorators.http import require_http_methods
from django.core.files.uploadedfile import InMemoryUploadedFile
from .forms import PDFUploadForm
from .extractor import PDFOCRExtractor
import json
import logging

logger = logging.getLogger(__name__)

@require_http_methods(["GET", "POST"])
def upload_pdf(request):
    form = PDFUploadForm()
    context = {
        'form': form,
        'data': None,
        'error': None,
        'debug': None
    }

    if request.method == "POST":
        form = PDFUploadForm(request.POST, request.FILES)
        
        if form.is_valid():
            try:
                pdf_file = form.cleaned_data["pdf_file"]
                
                # Validate file type more thoroughly
                if not isinstance(pdf_file, InMemoryUploadedFile):
                    raise ValueError("Invalid file upload")
                
                if not pdf_file.name.lower().endswith('.pdf'):
                    raise ValueError("Only PDF files are allowed")
                
                # Process the PDF - SINGLE extractor creation
                extractor = PDFOCRExtractor()
                result = extractor.extract(pdf_file)
                
                # Check if we got meaningful data
                global_count = len(result.get('global', {}))
                item_count = len([item for item in result.get('items', []) if any(v for k, v in item.items() if k != 'Components' and v)])
                component_count = sum(len(item.get('Components', [])) for item in result.get('items', []))
                
                if global_count == 0 and item_count == 0 and component_count == 0:
                    logger.warning("No extractable data found in PDF")
                    context['error'] = {
                        'message': "No extractable data found in PDF",
                        'details': "The PDF may be corrupted, password-protected, or contain non-text content"
                    }
                else:
                    context['data'] = {
                        "pretty": json.dumps(result, indent=2),
                        "raw": result
                    }
                    
                    # Add debug information
                    context['debug'] = {
                        'stats': {
                            'global_fields': global_count,
                            'items': item_count,
                            'components': component_count,
                            'ocr_text_length': result.get('debug', {}).get('ocr_text_length', 0)
                        },
                        'processing_steps': result.get('debug', {}).get('processing_steps', [])
                    }
                
                # Log successful extraction (without sensitive data)
                logger.info(f"Successfully processed PDF: {pdf_file.name} - "
                          f"Global: {global_count}, Items: {item_count}, Components: {component_count}")
                
            except Exception as e:
                logger.error(f"PDF processing error: {str(e)}", exc_info=True)
                context['error'] = {
                    'message': "An error occurred during PDF processing",
                    'details': str(e),
                    'type': type(e).__name__
                }
        else:
            context['error'] = {
                'message': "Form validation failed",
                'details': dict(form.errors)
            }

    return render(request, "upload.html", context)
# views.py
import json
import traceback
import pandas as pd
import os
from django.http import JsonResponse, HttpResponse
from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt
from django import forms
from datetime import datetime
from io import BytesIO
from .extractor import PDFOCRExtractor
from .extractor import export_to_excel

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
        
        # Check if this is a download request (no file upload)
        # Check if this is a download request (no file upload)
                # Check if this is a download request (no file upload)
                # Check if this is a download request (no file upload)
        if 'download_excel' in request.POST:
            print("📊 Download Excel request detected")
            if 'extracted_data' in request.session:
                print("📊 Processing Excel data...")
                try:
                    result = request.session['extracted_data']
                    
                    # Create daily filename on Desktop
                    today = datetime.now().strftime("%Y-%m-%d")
                    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                    excel_filename = os.path.join(desktop_path, f"Daily_PO_Extracts_{today}.xlsx")
                    
                    # Check if daily file exists on Desktop
                    file_exists = os.path.exists(excel_filename)
                    
                    if file_exists:
                        print(f"📊 Updating existing file on Desktop: {excel_filename}")
                        try:
                            existing_data = pd.read_excel(excel_filename)
                        except:
                            existing_data = pd.DataFrame()
                    else:
                        print(f"📊 Creating new file on Desktop: {excel_filename}")
                        existing_data = pd.DataFrame()
                    
                    # Prepare new data in flat structure
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    new_rows = []
                    
                    # Get global data
                    global_data = result.get('global', {})
                    
                    # Process each item and its components
                    for item in result.get('items', []):
                        components = item.get('Components', [])
                        
                        if components:
                            # Create row for each component
                            for component in components:
                                row = {
                                    # PO Level Data
                                    'PO #': global_data.get('PO #', ''),
                                    'Job #': item.get('Job #', ''),
                                    'Location': global_data.get('Location', ''),
                                    'PO Date': global_data.get('PO Date', ''),
                                    'Due Date': global_data.get('Due Date', ''),
                                    'Vendor ID #': global_data.get('Vendor ID #', ''),
                                    'Vendor Name': global_data.get('Vendor Name', ''),
                                    'Order Type': global_data.get('Order Type', ''),
                                    'Gold': global_data.get('Gold Rate', ''),
                                    'Platinum': global_data.get('Platinum Rate', ''),
                                    'Silver': global_data.get('Silver Rate', ''),
                                    
                                    # Job/Item Level Data
                                    'Richline Item #': item.get('Richline Item #', ''),
                                    'Vendor Item #': item.get('Vendor Item #', ''),
                                    'Pcs': item.get('Pieces/Carats', ''),
                                    'Cast Fin Wt Gold': item.get('Fin Weight (Gold)', ''),
                                    'CAST Fin WT Silver': item.get('Fin Weight (Silver)', ''),
                                    'Gold Loss %': item.get('LOSS %', {}).get('Gold', ''),
                                    'Silver Loss %': item.get('LOSS %', {}).get('Silver', ''),
                                    'Diamond Details': item.get('Diamond Details', ''),
                                    'Stone Pc': item.get('Stone Labor', ''),
                                    'Labor Pc': item.get('Labor PC', ''),
                                    'Unit Price': item.get('Unit Cost', ''),
                                    'Metal 1': item.get('Metal 1', ''),
                                    'Metal2': item.get('Metal 2', ''),
                                    
                                    # Component Level Data
                                    'Component': component.get('Component', ''),
                                    'Supply Policy': component.get('Supply Policy', ''),
                                    'Total Wt': component.get('Tot. Weight', ''),
                                    'Rate': component.get('Cost ($)', ''),
                                    
                                    # Tracking Data
                                    'Extraction_Time': timestamp,
                                    'Extraction_Date': today
                                }
                                new_rows.append(row)
                        else:
                            # If no components, create one row for the item
                            row = {
                                # PO Level Data
                                'PO #': global_data.get('PO #', ''),
                                'Location': global_data.get('Location', ''),
                                'PO Date': global_data.get('PO Date', ''),
                                'Due Date': global_data.get('Due Date', ''),
                                'Vendor ID #': global_data.get('Vendor ID #', ''),
                                'Vendor Name': global_data.get('Vendor Name', ''),
                                'Order Type': global_data.get('Order Type', ''),
                                'Gold': global_data.get('Gold Rate', ''),
                                'Platinum': global_data.get('Platinum Rate', ''),
                                'Silver': global_data.get('Silver Rate', ''),
                                
                                # Job/Item Level Data
                                'Job #': item.get('Job #', ''),
                                'Richline Item #': item.get('Richline Item #', ''),
                                'Vendor Item #': item.get('Vendor Item #', ''),
                                'Pcs': item.get('Pieces/Carats', ''),
                                'Cast Fin Wt Gold': item.get('Fin Weight (Gold)', ''),
                                'CAST Fin WT Silver': item.get('Fin Weight (Silver)', ''),
                                'Gold Loss %': item.get('LOSS %', {}).get('Gold', ''),
                                'Silver Loss %': item.get('LOSS %', {}).get('Silver', ''),
                                'Diamond Details': item.get('Diamond Details', ''),
                                'Stone Pc': item.get('Stone Labor', ''),
                                'Labor Pc': item.get('Labor PC', ''),
                                'Unit Price': item.get('Unit Cost', ''),
                                'Metal 1': item.get('Metal 1', ''),
                                'Metal2': item.get('Metal 2', ''),
                                
                                # Component Level Data (empty for items without components)
                                'Component': '',
                                'Supply Policy': '',
                                'Total Wt': '',
                                'Rate': '',
                                
                                # Tracking Data
                                'Extraction_Time': timestamp,
                                'Extraction_Date': today
                            }
                            new_rows.append(row)
                    
                    # Create new DataFrame
                    new_data = pd.DataFrame(new_rows)
                    
                    # Combine with existing data
                    if not existing_data.empty and not new_data.empty:
                        combined_data = pd.concat([existing_data, new_data], ignore_index=True)
                    elif not new_data.empty:
                        combined_data = new_data
                    else:
                        combined_data = existing_data
                    
                    # Save to Desktop (no download)
                    combined_data.to_excel(excel_filename, index=False)
                    
                    print(f"📊 File saved to Desktop: {excel_filename}")
                    print(f"📊 Total rows: {len(combined_data)}")
                    
                    # Return success message instead of download
                    template_data = {
                        'form': PDFUploadForm(),
                        'success_message': f'Excel file updated successfully! Check your Desktop: Daily_PO_Extracts_{today}.xlsx',
                        'file_info': {
                            'filename': f'Daily_PO_Extracts_{today}.xlsx',
                            'location': 'Desktop',
                            'total_rows': len(combined_data),
                            'new_rows_added': len(new_data),
                            'file_existed': file_exists
                        },
                        'show_download': True
                    }
                    
                    return render(request, 'upload.html', template_data)
                        
                except Exception as e:
                    print(f"❌ Excel save failed: {e}")
                    traceback.print_exc()
                    return render(request, 'upload.html', {
                        'form': PDFUploadForm(),
                        'error': {
                            'message': 'Excel save failed',
                            'details': str(e)
                        }
                    })
            else:
                print("❌ No extracted data found in session")
                return render(request, 'upload.html', {
                    'form': PDFUploadForm(),
                    'error': {
                        'message': 'No data to save',
                        'details': 'Please extract data first before saving to Excel.'
                    }
                })
        
        # Regular file upload and extraction
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
            
            # Store extracted data in session for later download
            request.session['extracted_data'] = result
            print("💾 Stored extracted data in session")
            
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
                },
                'show_download': True  # Flag to show download button
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
    
    # Fallback return (should never reach here)
    return render(request, 'upload.html', {'form': PDFUploadForm()})

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
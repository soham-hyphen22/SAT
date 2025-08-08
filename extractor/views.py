# views.py
import json
import traceback
import pandas as pd
import os
from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt
from django import forms
from datetime import datetime, timedelta
from io import BytesIO

# Make sure to import the correct class name
from .extractor import HybridPDFOCRExtractor as FastPDFOCRExtractor

### REFACTORED ###
# This function now iterates through multiple Purchase Orders
def export_to_excel(multi_po_data, base_filename="Daily_PO_Extracts"):
    """Export extracted data from multiple POs to a single daily Excel file on the Desktop."""

    today = datetime.now().strftime("%Y-%m-%d")
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = os.path.join(desktop_path, f"{base_filename}_{today}.xlsx")

    try:
        existing_df = pd.read_excel(filename) if os.path.exists(filename) else pd.DataFrame()
    except Exception as e:
        print(f"Warning: Could not read existing Excel file '{filename}'. Starting fresh. Error: {e}")
        existing_df = pd.DataFrame()

    all_new_rows = []

    # Iterate through each purchase order found in the PDF
    for po_result in multi_po_data.get('purchase_orders', []):
        timestamp = datetime.now().strftime("%H:%M:%S")
        global_data = po_result.get('global', {})
        po_number = global_data.get('PO #', 'UNKNOWN')

        # Check if this specific PO already exists in the file
        po_exists_in_file = False
        if not existing_df.empty and 'PO #' in existing_df.columns:
            po_exists_in_file = existing_df['PO #'].astype(str).eq(str(po_number)).any()

        if po_exists_in_file:
            print(f"Skipping PO# {po_number} as it already exists in the Excel file.")
            continue # Skip to the next PO

        # Prepare the PO-level data that will be repeated for this PO's first row
        po_header_data = {
            'PO #': po_number,
            'Location': global_data.get('Location', ''),
            'PO Date': global_data.get('PO Date', ''),
            'Due Date': global_data.get('Due Date', ''),
            'Vendor ID #': global_data.get('Vendor ID #', ''),
            'Order Type': global_data.get('Order Type', ''),
            'Gold Rate': global_data.get('Gold Rate', ''),
            'Platinum Rate': global_data.get('Platinum Rate', ''),
            'Silver Rate': global_data.get('Silver Rate', ''),
            'Extraction_Time': timestamp,
            'Extraction_Date': today
        }

        # Process each item within the current PO
        items_in_po = po_result.get('items', [])
        if not items_in_po: # Handle case where a PO has global data but no items
             # Add a single row with just the PO header data
            row = po_header_data.copy()
            # Fill in blank item/component data
            row.update({ 'Job #': 'N/A', 'Richline Item #': 'N/A', 'Component': 'N/A'})
            all_new_rows.append(row)
            continue

        for item_idx, item in enumerate(items_in_po):
            item_data = {
                'Job #': item.get('Job #', ''),
                'Richline Item #': item.get('Richline Item #', ''),
                'Vendor Item #': item.get('Vendor Item #', ''),
                'Fin Weight (Gold)': item.get('Fin Weight (Gold)', ''),
                'Stone Labor': item.get('Stone Labor', ''),
            }

            components = item.get('Components', [])
            if not components: # Handle item with no components
                row = item_data.copy()
                if item_idx == 0: # Add PO header to the first item of this PO
                    row.update(po_header_data)
                # Fill in blank component data
                row.update({ 'Component': 'N/A', 'Supply Policy': 'N/A'})
                all_new_rows.append(row)
                continue

            for comp_idx, component in enumerate(components):
                row = {}
                # Add PO header and item data only to the very first row of the item
                if comp_idx == 0:
                    row.update(item_data)
                    if item_idx == 0:
                        row.update(po_header_data)

                # Always add component data
                row.update({
                    'Component': component.get('Component', ''),
                    'Supply Policy': component.get('Supply Policy', ''),
                    'Tot. Weight': component.get('Tot. Weight', ''),  # Fixed field name
                    'Cost ($)': component.get('Cost ($)', ''),  # Fixed field name
                })
                all_new_rows.append(row)

    if not all_new_rows:
        print("No new data to add to Excel.")
        return filename, os.path.exists(filename)

    new_df = pd.DataFrame(all_new_rows)
    # Define the desired column order
    column_order = [
        'Extraction_Date', 'Extraction_Time', 'PO #', 'Location', 'PO Date', 'Due Date', 'Vendor ID #', 'Order Type',
        'Gold Rate', 'Platinum Rate', 'Silver Rate', 'Job #', 'Richline Item #', 'Vendor Item #',
        'Fin Weight (Gold)', 'Stone Labor', 'Component', 'Supply Policy', 'Tot. Weight', 'Cost ($)'
    ]
    # Reorder the DataFrame, adding missing columns as blank
    new_df = new_df.reindex(columns=column_order)

    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    combined_df.to_excel(filename, index=False)

    return filename, os.path.exists(filename)


class PDFUploadForm(forms.Form):
    pdf_file = forms.FileField(
        label="Select a PDF File",
        widget=forms.ClearableFileInput(attrs={'accept': '.pdf', 'class': 'form-control'})
    )

### REFACTORED ###
# This is the custom JSON encoder needed for the session data
class CustomJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, timedelta)):
            return str(obj)
        return super().default(obj)

def upload_pdf(request):
    if request.method == 'POST' and request.FILES.get('pdf_file'):
        try:
            pdf_file = request.FILES['pdf_file']
            
            extractor = FastPDFOCRExtractor()
            result = extractor.extract_with_adaptive_quality(pdf_file)
            
            if "error" in result:
                context = {
                    'error': {'message': result.get('error', 'Unknown error'), 'details': result.get('details', '')},
                    'success': False,
                    'form': PDFUploadForm()
                }
                return render(request, 'upload.html', context)
            
            # Enhanced debug analysis
            debug_info = []
            
            if result.get('purchase_orders'):
                # Multiple POs
                debug_info.append(f"=== MULTIPLE PO EXTRACTION ANALYSIS ===")
                debug_info.append(f"Total Purchase Orders: {len(result['purchase_orders'])}")
                
                for i, po in enumerate(result['purchase_orders'][:3]):  # First 3 POs
                    debug_info.append(f"\nPO {i+1}: {po['po_number']}")
                    debug_info.append(f"  Global fields: {len(po['global'])}")
                    for key, value in po['global'].items():
                        debug_info.append(f"    {key}: {value}")
                    
                    debug_info.append(f"  Items: {len(po['items'])}")
                    for j, item in enumerate(po['items'][:2]):  # First 2 items per PO
                        debug_info.append(f"    Item {j+1}: {item.get('Richline Item #', 'Unknown')}")
                        debug_info.append(f"      Job #: {item.get('Job #', 'Not found')}")
                        debug_info.append(f"      Vendor Item #: {item.get('Vendor Item #', 'Not found')}")
                        debug_info.append(f"      Components: {len(item.get('Components', []))}")
                        
                        for k, comp in enumerate(item.get('Components', [])[:2]):  # First 2 components
                            debug_info.append(f"        Component {k+1}: {comp.get('Component', 'No name')}")
                            debug_info.append(f"          Cost: {comp.get('Cost ($)', 'No cost')}")
                            debug_info.append(f"          Weight: {comp.get('Tot. Weight', 'No weight')}")
                            debug_info.append(f"          Policy: {comp.get('Supply Policy', 'No policy')}")
            
            else:
                # Single PO (existing logic)
                debug_info.append(f"=== SINGLE PO EXTRACTION ANALYSIS ===")
                # ... existing debug code ...
            
            try:
                result_json_pretty = json.dumps(result, indent=2, ensure_ascii=False, cls=CustomJSONEncoder)
            except Exception as json_error:
                result_json_pretty = f"Error serializing JSON: {str(json_error)}"
            
            context = {
                'result': result,
                'result_json': result_json_pretty,
                'debug_analysis': '\n'.join(debug_info),
                'success': True,
                'filename': pdf_file.name,
                'form': PDFUploadForm()
            }
            
            return render(request, 'upload.html', context)
            
        except Exception as e:
            traceback.print_exc()
            context = {
                'error': {'message': 'An unexpected error occurred during processing.', 'details': str(e)},
                'success': False,
                'form': PDFUploadForm()
            }
            return render(request, 'upload.html', context)
    
    return render(request, 'upload.html', {'form': PDFUploadForm()})
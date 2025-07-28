import re
import pytesseract
import pdf2image
import cv2
import numpy as np
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Samuel Aaron\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

class PDFOCRExtractor:
    def __init__(self):
        self.global_patterns = {
            "PO #": r"Order Number:\s*(RPO\d+)",
            "PO Date": r"Order Date:\s*(\d{2}/\d{2}/\d{2,4})",
            "Location": r"Location:\s*([A-Z]{2})",
        }
        self.product_patterns = {
            "main_product": r'\b([DA|BA]\d{3,4}[A-Z0-9]+)\b',
            "job_number": r'\b(RFP\d{6,}|RSET\d{6,})\b',
            "vendor_style": r'([A-Z0-9]+(?:-[A-Z0-9]+)*)\s*\|?\s*(?:EXCL|EXCLUSIVE)?'
        }
        self.job_pattern = r"(RFP\d{6,}|RSET\d{6,})"
        self.component_heading_keywords = [
            "Component",
            "Cost",
            "Tot. Weight",
            "Supplied By",
        ]

    def convert_pdf_to_image(self, pdf_file):
        poppler_path = r"C:\Users\Samuel Aaron\Documents\Release-24.08.0-0\poppler-24.08.0\Library\bin"
        try:
            pdf_file.seek(0)
            images = pdf2image.convert_from_bytes(pdf_file.read(), dpi=300, poppler_path=poppler_path)
            return images if images else None
        except Exception as e:
            print("---")
            print("!!! PDF to Image Conversion FAILED !!!")
            print(f"!!! The script tried to use this Poppler path: '{poppler_path}'")
            print(f"!!! Full Error: {e}")
            print("---")
            return None

    def preprocess_image(self, image):
        try:
            open_cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
            denoised = cv2.fastNlMeansDenoising(gray, h=10)
            _, thresh = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            return Image.fromarray(thresh)
        except Exception as e:
            print(f"Image preprocessing failed: {e}")
            return image

    def extract_text(self, image):
        try:
            return pytesseract.image_to_string(image)
        except Exception as e:
            print(f"OCR extraction failed: {e}")
            return ""

    def extract_global_fields(self, full_text):
        """Extract global fields that apply to all products"""
        result = {}
        lines = full_text.split('\n')
        
        # Extract PO Number
        po_match = re.search(r"Order Number:\s*(RPO\d+)", full_text, re.IGNORECASE)
        if po_match:
            result["PO #"] = po_match.group(1)
        
        # Extract PO Date
        date_match = re.search(r"Order Date:\s*(\d{2}/\d{2}/\d{2,4})", full_text, re.IGNORECASE)
        if date_match:
            result["PO Date"] = date_match.group(1)
        
        # Extract Location
        location_match = re.search(r"Location:\s*([A-Z]{2})", full_text, re.IGNORECASE)
        if location_match:
            result["Location"] = location_match.group(1)
        
        # Extract Vendor ID (handle both formats)
        vendor_patterns = [
            r"Vendor ID\s+([A-Z0-9]+)",
            r"Vendor ID\s*#?\s*([A-Z0-9]+)"
        ]
        for pattern in vendor_patterns:
            vendor_match = re.search(pattern, full_text, re.IGNORECASE)
            if vendor_match:
                result["Vendor ID #"] = vendor_match.group(1)
                break
        
        # Extract Vendor Name - Fixed logic
        if "Vendor ID #" in result:
            vendor_id = result["Vendor ID #"]
            # Look for vendor name in the lines after Vendor ID
            for i, line in enumerate(lines):
                if f"Vendor ID {vendor_id}" in line or re.search(rf"Vendor ID\s+{vendor_id}", line):
                    # Look in next few lines for vendor name
                    for j in range(i + 1, min(i + 5, len(lines))):
                        candidate_line = lines[j].strip()
                        # Skip empty lines and lines that look like addresses or other data
                        if (candidate_line and 
                            len(candidate_line) > 3 and 
                            not candidate_line.startswith(('Ship To', 'Due Date', 'Terms', 'Order Type', 'Gold:', 'Page:')) and
                            not re.match(r'^[A-Z]{2}-\d+', candidate_line) and  # Skip address patterns
                            not candidate_line.isdigit()):
                            result["Vendor Name"] = candidate_line
                            break
                    break
        
        # Extract Order Type
        order_type_match = re.search(r"Order Type:\s*([A-Z]+)", full_text, re.IGNORECASE)
        if order_type_match:
            result["Order Type"] = order_type_match.group(1)
        
        # Extract Due Date
        due_date_match = re.search(r"Due Date:\s*([^\n]+)", full_text, re.IGNORECASE)
        if due_date_match:
            result["Due Date"] = due_date_match.group(1).strip()
        
        # Extract Metal Rates
        metal_rates_section = re.search(
            r"Gold:\s*([\d,]+\.?\d*)\s*\nPlatinum:\s*([\d,]+\.?\d*)\s*\nSilver:\s*([\d,]+\.?\d*)", 
            full_text, re.MULTILINE
        )
        if metal_rates_section:
            result["Gold Rate"] = metal_rates_section.group(1).replace(",", "")
            result["Platinum Rate"] = metal_rates_section.group(2).replace(",", "")
            result["Silver Rate"] = metal_rates_section.group(3).replace(",", "")
        
        return result

    def segment_products_by_boundaries(self, lines):
        """Segment products using flexible boundary detection"""
        products = []
        current_product = None
        
        print(f"DEBUG: Processing {len(lines)} lines for product segmentation")
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # More flexible product header detection
            # Look for lines that contain product codes and pricing info
            product_patterns = [
                # Pattern 1: Full table format with pipes
                r'([DA|BA]\d{3,4}[A-Z0-9]+)\s*\|\s*([^|]+)\|\s*([\d.]+)\s*\|\s*(\d+)\s*EA\s*\|\s*([\d.]+)\s*\|\s*([\d,.]+)',
                # Pattern 2: Space-separated format
                r'([DA|BA]\d{3,4}[A-Z0-9]+)\s+([^|]+?)\s+([\d.]+)\s+(\d+)\s+EA\s+([\d.]+)\s+([\d,.]+)',
                # Pattern 3: Just product code with description and price
                r'([DA|BA]\d{3,4}[A-Z0-9]+)\s+.*?\s+([\d.]+)\s+(\d+)\s+EA'
            ]
            
            product_match = None
            for pattern in product_patterns:
                product_match = re.search(pattern, line)
                if product_match:
                    print(f"DEBUG: Found product match with pattern: {pattern}")
                    print(f"DEBUG: Line: {line}")
                    break
            
            if product_match:
                # Save previous product if exists
                if current_product:
                    products.append(current_product)
                    print(f"DEBUG: Saved previous product: {current_product['richline_item']}")
                
                # Start new product
                groups = product_match.groups()
                current_product = {
                    'richline_item': groups[0],
                    'description': groups[1].strip() if len(groups) > 1 else "",
                    'unit_cost': groups[2] if len(groups) > 2 else "",
                    'pieces': groups[3] if len(groups) > 3 else "",
                    'ext_gross_wt': groups[4] if len(groups) > 4 else "",
                    'ext_cost': groups[5].replace(',', '') if len(groups) > 5 else "",
                    'lines': [line],
                    'start_index': i
                }
                
                print(f"DEBUG: Started new product: {current_product['richline_item']}")
                
                # Look for vendor style in next line
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    vendor_match = re.search(r'^([A-Z0-9]+(?:-[A-Z0-9]+)*)', next_line)
                    if vendor_match and not re.search(r'[DA|BA]\d{3,4}[A-Z0-9]+', next_line):
                        current_product['vendor_item'] = vendor_match.group(1)
                        current_product['lines'].append(next_line)
                        i += 1
                
            elif current_product:
                # Add line to current product
                current_product['lines'].append(line)
                
                # Check if we've reached the end of this product
                if self.is_product_end(line, lines, i):
                    current_product['end_index'] = i
                    products.append(current_product)
                    print(f"DEBUG: Ended product: {current_product['richline_item']}")
                    current_product = None
            
            i += 1
        
        # Don't forget the last product
        if current_product:
            current_product['end_index'] = len(lines) - 1
            products.append(current_product)
            print(f"DEBUG: Added final product: {current_product['richline_item']}")
        
        print(f"DEBUG: Total products found: {len(products)}")
        return products

    def is_product_end(self, line, lines, current_index):
        """Determine if we've reached the end of a product section"""
        # Check if next few lines contain another product
        for i in range(current_index + 1, min(current_index + 5, len(lines))):
            if i < len(lines):
                next_line = lines[i]
                # Look for product codes
                if re.search(r'[DA|BA]\d{3,4}[A-Z0-9]+', next_line):
                    # Make sure it's a product header, not just a reference
                    if ('|' in next_line or 'EA' in next_line) and re.search(r'[\d.]+', next_line):
                        return True
        # Check for section end markers
        end_markers = ['Totals:', 'PLEASE COMMUNICATE', 'There is a +/-']
        if any(marker in line for marker in end_markers):
            return True
        return False

    def extract_product_data(self, product_info):
        """Extract detailed data for a single product"""
        lines = product_info['lines']
        product_text = '\n'.join(lines)
        product_data = {
            'Richline Item': product_info['richline_item'],
            'Description': product_info['description'],
            'Unit Price': product_info['unit_cost'],
            'Pieces/Carats': product_info['pieces'],
            'Ext. Gross Wt.': product_info['ext_gross_wt'],
            'Ext. Cost': product_info['ext_cost'],
            'Components': [],
            'CAST Fin WT': {},
            'LOSS %': {}
        }
        # Add vendor item if found
        if 'vendor_item' in product_info:
            product_data['Vendor Item'] = product_info['vendor_item']
        # Extract Job Number (RFP/RSET)
        job_match = re.search(r'\b(RFP\d{6,}|RSET\d{6,})\b', product_text)
        if job_match:
            product_data['Job #'] = job_match.group(1)
        # Extract product-specific fields
        self.extract_product_fields(product_data, product_text, lines)
        # Extract components
        product_data['Components'] = self.extract_product_components(lines)
        return product_data

    def extract_product_fields(self, product_data, product_text, lines):
        """Extract specific fields for this product"""
        # Extract Stone PC
        stone_patterns = [
            r'Stone PC\s*\|\s*([^|\n]+)',
            r'Stone PC\s+([0-9]+\.?[0-9]*)',
            r'Stone PC[:\s]+([0-9]+\.?[0-9]*)'
        ]
        for pattern in stone_patterns:
            stone_match = re.search(pattern, product_text)
            if stone_match:
                stone_value = stone_match.group(1).strip()
                if stone_value != '*********':
                    product_data['Stone PC'] = stone_value
                break
        # Extract Labor PC
        labor_patterns = [
            r'Labor PC\s*\|\s*([\d.]+)',
            r'Labor PC\s+([\d.]+)',
            r'Labor PC[:\s]+([\d.]+)'
        ]
        for pattern in labor_patterns:
            labor_match = re.search(pattern, product_text)
            if labor_match:
                product_data['Labor PC'] = labor_match.group(1)
                break
        # Extract Diamond TW
        diamond_patterns = [
            r'Diamond TW\s*\|\s*([\d.]+)',
            r'Diamond TW\s+([\d.]+)',
            r'Diamond TW[:\s]+([\d.]+)'
        ]
        for pattern in diamond_patterns:
            diamond_match = re.search(pattern, product_text)
            if diamond_match:
                product_data['Diamond Details'] = f"Diamond TW: {diamond_match.group(1)}"
                break
        # Extract CAST Fin WT
        cast_match = re.search(r'CAST Fin WT:\s*Gold:\s*([\d.]+)(?:\s*Silver:\s*([\d.]+))?', product_text)
        if cast_match:
            product_data['CAST Fin WT']['Gold'] = cast_match.group(1)
            product_data['Fin Weight (Gold)'] = cast_match.group(1)
            if cast_match.group(2):
                product_data['CAST Fin WT']['Silver'] = cast_match.group(2)
                product_data['Fin Weight (Silver)'] = cast_match.group(2)
        # Extract LOSS %
        loss_match = re.search(r'LOSS %:\s*Gold:\s*([\d.]+)%?(?:\s*Silver:\s*([\d.]+)%?)?', product_text)
        if loss_match:
            product_data['LOSS %']['Gold'] = loss_match.group(1)
            if loss_match.group(2):
                product_data['LOSS %']['Silver'] = loss_match.group(2)
        # Extract Metal Category from description
        metal_category_match = re.search(r'\b(\d+K[A-Z]?)\b', product_data['Description'])
        if metal_category_match:
            product_data['Metal Category'] = metal_category_match.group(1)

    def extract_product_components(self, lines):
        """Extract components table for this specific product"""
        components = []
        # Find component table start
        table_start = -1
        for i, line in enumerate(lines):
            if 'Supplied By' in line and 'Component' in line:
                table_start = i + 1
                break
        if table_start == -1:
            return components
        # Find table end
        table_end = len(lines)
        for i in range(table_start, len(lines)):
            line = lines[i].strip()
            if (not line or 
                re.search(r'[DA|BA]\d{3,4}[A-Z0-9]+', line) or
                any(marker in line for marker in ['Totals:', 'There is a +/-', 'PLEASE COMMUNICATE'])):
                table_end = i
                break
        # Parse component rows
        for i in range(table_start, table_end):
            if i >= len(lines):
                break
            line = lines[i].strip()
            if not line:
                continue
            component = self.parse_component_row(line)
            if component:
                components.append(component)
        return components

    def parse_component_row(self, line):
        """Parse a single component row"""
        # Split by | or multiple spaces
        if '|' in line:
            parts = [part.strip() for part in line.split('|') if part.strip()]
        else:
            parts = re.split(r'\s{2,}', line.strip())
            parts = [part.strip() for part in parts if part.strip()]
        if len(parts) < 3:
            return None
        component = {
            "Component": parts[0] if parts else "",
            "Cost ($)": "",
            "Tot. Weight": "",
            "Supply Policy": ""
        }
        # Parse the remaining parts to extract cost, weight, and policy
        for part in parts[1:]:
            # Cost pattern (CT values)
            if re.match(r'[\d.]+\s*CT$', part, re.IGNORECASE):
                if not component["Cost ($)"]:
                    component["Cost ($)"] = part
                else:
                    component["Tot. Weight"] = part
            # Weight pattern (GR values)
            elif re.match(r'[\d.]+\s*GR$', part, re.IGNORECASE):
                component["Tot. Weight"] = part
            # Supply policy patterns
            elif any(policy in part for policy in ["Send To", "In House", "By Vendor"]):
                component["Supply Policy"] = part
        return component if component["Component"] else None

    def validate_extracted_data(self, global_data, items_data, text):
        """Validate the quality of extracted data"""
        if not text or len(text.strip()) < 10:
            return False, "No readable text found in PDF"
        key_indicators = [
            bool(global_data.get("PO #")),
            bool(global_data.get("Vendor ID #")),
            len(items_data) > 0,
            len(global_data) > 0,
        ]
        if sum(key_indicators) >= 2:
            return True, "Valid data extracted"
        if "password" in text.lower():
            return False, "PDF appears to be password-protected"
        if len(text.strip()) < 50:
            return False, "PDF may be corrupted or contain non-text content"
        return False, "No extractable data found in PDF"

    def extract(self, pdf_file):
        """Main extraction method with improved product segmentation"""
        debug = {"processing_steps": []}
        try:
            # Convert PDF to images
            images = self.convert_pdf_to_image(pdf_file)
            debug["processing_steps"].append("PDF converted to images")
            if not images:
                return {
                    "error": "Failed to convert PDF to images",
                    "details": "The PDF may be corrupted, password-protected, or contain non-text content",
                    "debug": debug,
                }
            # Process all pages and combine text
            all_text = ""
            all_lines = []
            for i, image in enumerate(images):
                preprocessed_image = self.preprocess_image(image)
                page_text = self.extract_text(preprocessed_image)
                all_text += f"\n#page {i+1}\n" + page_text
                all_lines.extend(page_text.splitlines())
                debug["processing_steps"].append(f"Processed page {i+1}")
            debug["processing_steps"].append("Text extracted via OCR from all pages")
            debug["total_pages"] = len(images)
            if not all_text:
                return {
                    "error": "No extractable data found in PDF",
                    "details": "OCR failed to extract any text from the PDF",
                    "debug": debug,
                }
            debug["ocr_text_length"] = len(all_text)
            # Extract global data
            global_data = self.extract_global_fields(all_text)
            debug["processing_steps"].append(f"Extracted {len(global_data)} global fields")
            # Segment products by boundaries
            product_segments = self.segment_products_by_boundaries(all_lines)
            debug["processing_steps"].append(f"Segmented {len(product_segments)} products")
            # Extract data for each product
            items_data = []
            for product_info in product_segments:
                product_data = self.extract_product_data(product_info)
                items_data.append(product_data)
            debug["processing_steps"].append(f"Extracted data for {len(items_data)} items")
            # Validate extraction
            is_valid, validation_message = self.validate_extracted_data(
                global_data, items_data, all_text
            )
            debug["validation_message"] = validation_message
            if not is_valid:
                return {
                    "error": "No extractable data found in PDF",
                    "details": validation_message,
                    "debug": debug,
                    "raw_text_sample": all_text[:500] + "..." if len(all_text) > 500 else all_text,
                }
            # Count total components across all items
            total_components = sum(len(item.get("Components", [])) for item in items_data)
            debug["processing_steps"].append(f"Extracted {total_components} total components")
            debug["total_fields_extracted"] = len(global_data) + sum(len(item) for item in items_data)
            return {
                "global": global_data,
                "items": items_data,
                "debug": debug
            }
        except Exception as e:
            return {
                "error": "Processing failed",
                "details": f"An unexpected error occurred: {str(e)}",
                "debug": debug,
            }
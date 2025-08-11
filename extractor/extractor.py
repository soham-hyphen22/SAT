import re
import pytesseract
import pdf2image
import cv2
import numpy as np
from PIL import Image
from datetime import datetime
import concurrent.futures
import traceback
import json
import pandas as pd
import os
from pathlib import Path
import time # Imported for timing

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Samuel Aaron\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

class AccuracyIntelligence:
    """Enhanced accuracy validation"""
    def __init__(self):
        self.field_accuracy = {}
        self.extraction_history = []
        self.accuracy_patterns = {}
    
    def validate_extraction(self, result):
        validation = {
            "accuracy_score": 0,
            "field_scores": {},
            "issues": [],
            "suggestions": []
        }
        
        # Handle both single and multiple RPO structures
        if "purchase_orders" in result:
            # Multiple RPOs
            all_scores = []
            for po in result["purchase_orders"]:
                global_data = po.get("global", {})
                scores = {field: self.validate_field(field, value) for field, value in global_data.items()}
                all_scores.extend(scores.values())
                validation["field_scores"].update(scores)
        else:
            # Single RPO
            global_data = result.get("global", {})
            validation["field_scores"] = {field: self.validate_field(field, value) 
                                        for field, value in global_data.items()}
            all_scores = list(validation["field_scores"].values())
        
        if all_scores:
            validation["accuracy_score"] = sum(all_scores) / len(all_scores)
        
        return validation
    
    def validate_field(self, field_name, field_value):
        if not field_value:
            return 0.0
        
        validators = {
            "PO #": lambda x: 1.0 if re.match(r'^RPO\d+$', str(x)) else 0.3,
            "Vendor ID #": lambda x: 1.0 if 3 <= len(str(x)) <= 20 else 0.5,
            "Due Date": lambda x: 1.0 if re.match(r'\d{1,2}/\d{1,2}/\d{2,4}', str(x)) else 0.3,
            "Order Type": lambda x: 1.0 if str(x).upper() in ["STOCK", "MCH", "SPC", "ASSAY", "ASSET"] else 0.4,
            "Gold Rate": lambda x: self.validate_rate(x, 1500, 3000),
            "Silver Rate": lambda x: self.validate_rate(x, 15, 50),
            "Platinum Rate": lambda x: self.validate_rate(x, 800, 1500),
            "Vendor Name": lambda x: 1.0 if len(str(x)) >= 5 else 0.6,
            "Location": lambda x: 1.0 if len(str(x)) >= 2 else 0.4
        }
        
        if field_name in validators:
            return validators[field_name](field_value)
        return 0.8
    
    def validate_rate(self, rate_value, min_val, max_val):
        try:
            rate = float(str(rate_value).replace(',', ''))
            return 1.0 if min_val <= rate <= max_val else 0.4
        except:
            return 0.2

class HybridPDFOCRExtractor:
    def __init__(self, output_folder=None):
        # Original patterns for backward compatibility
        self.global_patterns = {
            "PO #": r"(RPO\d+)",
            "PO Date": r"\b(\d{2}/\d{2}/\d{2,4})\b",
            "Location": r"Location[:\s]*([A-Z]{2,})"
        }
        self.job_pattern = r"(RFP\d{6,}|RSET\d{6,})"
        self.accuracy_intelligence = AccuracyIntelligence()
        
        # Processing parameters
        self.fast_dpi = 200
        self.accurate_dpi = 300
        self.max_workers = 4
        
        # Expected fields for consistency
        self.GLOBAL_FIELDS = [
            "PO #", "PO Date", "Location", "Vendor ID #", "Vendor Name", 
            "Due Date", "Order Type", "Gold Rate", "Silver Rate", "Platinum Rate"
        ]
        
        self.ITEM_FIELDS = [
            "Richline Item #", "Vendor Item #", "Job #", "Metal 1", "Metal 2",
            "Stone PC", "Labor PC", "Diamond TW", "Fin Weight (Gold)", 
            "Fin Weight (Silver)", "Loss % (Gold)", "Loss % (Silver)", 
            "Pieces/Carats", "Ext. Gross Wt."
        ]

    # ===============================
    # MAIN ENTRY POINTS (All lead to the same extraction logic)
    # ===============================
    
    def extract(self, pdf_file):
        """Main extract method for backward compatibility"""
        return self.extract_with_adaptive_quality(pdf_file)
    
    def extract_from_pdf(self, pdf_file):
        """Alternative method name"""
        return self.extract_with_adaptive_quality(pdf_file)
    
    def process_pdf(self, pdf_file):
        """Another alternative method name"""
        return self.extract_with_adaptive_quality(pdf_file)
    
    def extract_data(self, pdf_file):
        """Another common method name"""
        return self.extract_with_adaptive_quality(pdf_file)

    def extract_with_adaptive_quality(self, pdf_file):
        """Enhanced main extraction method using state machine for better accuracy"""
        start_time = datetime.now()
        debug = {"processing_steps": []}
        
        try:
            # Try enhanced state machine approach first
            debug["processing_steps"].append("Starting enhanced state machine extraction...")
            result = self.extract_with_state_machine_internal(pdf_file, debug)
            
            if "error" in result:
                # Fallback to original method if state machine fails
                debug["processing_steps"].append("State machine failed, falling back to original method...")
                result = self._extract_fast(pdf_file, debug)
            
            if "error" not in result:
                # Validate accuracy
                accuracy_check = self.accuracy_intelligence.validate_extraction(result)
                result["accuracy"] = accuracy_check
                debug["processing_steps"].append(f"Extraction accuracy: {accuracy_check['accuracy_score']:.2f}")
            
            debug["processing_time"] = str(datetime.now() - start_time)
            result["debug"] = debug
            
            return result
            
        except Exception as e:
            debug["processing_time"] = str(datetime.now() - start_time)
            return {
                "error": "Processing failed",
                "details": str(e),
                "debug": debug,
                "traceback": traceback.format_exc()
            }

    # ===============================
    # PDF PROCESSING
    # ===============================

    def convert_pdf_to_image(self, pdf_file, dpi=200, use_jpeg=True):
        """Convert PDF to images"""
        poppler_path = r"C:\Users\Samuel Aaron\Documents\Release-24.08.0-0\poppler-24.08.0\Library\bin"
        try:
            pdf_file.seek(0)        
            images = pdf2image.convert_from_bytes(
                pdf_file.read(),
                dpi=dpi,
                poppler_path=poppler_path,
                thread_count=self.max_workers,
                fmt='jpeg' if use_jpeg else 'ppm'
            )
            return images if images else None
        except Exception as e:
            print(f"PDF to Image Conversion FAILED: {e}")
            return None

    def preprocess_image_adaptive(self, image, enhanced=False):
        """Image preprocessing"""
        try:
            open_cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
            
            if enhanced:
                gray = cv2.fastNlMeansDenoising(gray, h=10)
                
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            return Image.fromarray(thresh)
        except Exception as e:
            print(f"Image preprocessing failed: {e}")
            return image

    def extract_text_adaptive(self, image, enhanced=False):
        """OCR text extraction"""
        try:
            if enhanced:
                return pytesseract.image_to_string(image, config='--oem 3 --psm 6')
            else:
                return pytesseract.image_to_string(image, config='--oem 3 --psm 6')
        except Exception as e:
            print(f"OCR extraction failed: {e}")
            return ""

    def process_page_parallel(self, page_data, enhanced=False):
        """Parallel page processing"""
        page_num, image = page_data
        try:
            preprocessed_image = self.preprocess_image_adaptive(image, enhanced)
            page_text = self.extract_text_adaptive(preprocessed_image, enhanced)
            return page_num, page_text
        except Exception as e:
            print(f"Error processing page {page_num}: {e}")
            return page_num, ""

    def extract_text_with_coordinates(self, images):
        """Extract text with coordinate information"""
        all_text = ""
        all_lines = []
        text_with_coords = []
        
        for page_num, image in enumerate(images):
            try:
                # Get text with bounding boxes
                data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
                page_text = pytesseract.image_to_string(image)
                
                # Store coordinate information
                for i in range(len(data['text'])):
                    if data['text'][i].strip():
                        text_with_coords.append({
                            'text': data['text'][i],
                            'x': data['left'][i],
                            'y': data['top'][i],
                            'width': data['width'][i],
                            'height': data['height'][i],
                            'page': page_num
                        })
                
                all_text += f"\n#page {page_num + 1}\n" + page_text
                all_lines.extend(page_text.splitlines())
                
            except Exception as e:
                print(f"Error processing page {page_num}: {e}")
                # Fallback to simple text extraction
                page_text = pytesseract.image_to_string(image)
                all_text += f"\n#page {page_num + 1}\n" + page_text
                all_lines.extend(page_text.splitlines())
                
        return all_text, all_lines, text_with_coords

    # ===============================
    # STATE MACHINE EXTRACTION
    # ===============================

    def extract_with_state_machine_internal(self, pdf_file, debug):
        """Enhanced state machine extraction"""
        try:
            # Step 1: Extract text with coordinates
            images = self.convert_pdf_to_image(pdf_file, dpi=self.accurate_dpi)
            if not images:
                return {"error": "Failed to convert PDF to images"}
            
            all_text, all_lines, text_with_coords = self.extract_text_with_coordinates(images)
            debug["processing_steps"].append(f"Extracted text from {len(images)} pages")
            
            # Step 2: Split into RPO blocks using state machine
            rpo_blocks = self.split_into_rpo_blocks(all_lines, debug)
            debug["processing_steps"].append(f"Found {len(rpo_blocks)} RPO blocks")
            
            # Step 3: Process each RPO block
            processed_rpos = []
            for rpo_block in rpo_blocks:
                rpo_result = self.process_rpo_block(rpo_block, all_lines, text_with_coords, debug)
                if rpo_result:
                    processed_rpos.append(rpo_result)
            
            # Step 4: Format final result
            return self.format_final_result(processed_rpos, debug)
            
        except Exception as e:
            return {
                "error": "State machine processing failed",
                "details": str(e)
            }

    def split_into_rpo_blocks(self, all_lines, debug):
        """FIXED: Properly detect multiple RPOs"""
        rpo_blocks = []
        rpo_pattern = r'\b(RPO?\d+)\b'  # Added ? to handle RP0915176 vs RPO911481
        
        # Find all RPO occurrences with their line positions
        rpo_occurrences = []
        for i, line in enumerate(all_lines):
            rpo_matches = re.findall(rpo_pattern, line, re.IGNORECASE)
            for rpo in rpo_matches:
                # Normalize RPO format
                if rpo.upper().startswith('RP0'):
                    rpo = 'RPO' + rpo[3:]  # Convert RP0915176 to RPO915176
                rpo_occurrences.append((i, rpo.upper()))
        
        # Group by unique RPO numbers
        unique_rpos = {}
        for line_num, rpo in rpo_occurrences:
            if rpo not in unique_rpos:
                unique_rpos[rpo] = []
            unique_rpos[rpo].append(line_num)
        
        debug.setdefault("state_transitions", []).append(f"Found unique RPOs: {list(unique_rpos.keys())}")
        
        # Create blocks for each unique RPO
        if len(unique_rpos) == 1:
            # Single RPO
            single_rpo = list(unique_rpos.keys())[0]
            rpo_blocks.append({
                "rpo_number": single_rpo,
                "start_line": 0,
                "end_line": len(all_lines),
                "lines": all_lines
            })
            debug.setdefault("state_transitions", []).append(f"Single RPO detected: {single_rpo}")
            
        elif len(unique_rpos) > 1:
            # Multiple RPOs - create separate blocks
            rpo_list = list(unique_rpos.items())
            rpo_list.sort(key=lambda x: x[1][0])  # Sort by first occurrence
            
            for i, (rpo_number, line_positions) in enumerate(rpo_list):
                start_line = line_positions[0]  # First occurrence
                
                # End line is start of next RPO or end of document
                if i + 1 < len(rpo_list):
                    end_line = rpo_list[i + 1][1][0]  # Start of next RPO
                else:
                    end_line = len(all_lines)
                
                rpo_blocks.append({
                    "rpo_number": rpo_number,
                    "start_line": start_line,
                    "end_line": end_line,
                    "lines": all_lines[start_line:end_line]
                })
                debug.setdefault("state_transitions", []).append(f"Created block for {rpo_number}: lines {start_line}-{end_line}")
        
        else:
            # No RPO found
            rpo_blocks.append({
                "rpo_number": "RPO001",
                "start_line": 0,
                "end_line": len(all_lines),
                "lines": all_lines
            })
            debug.setdefault("state_transitions", []).append("No RPO found - created default")
        
        return rpo_blocks

    def process_rpo_block(self, rpo_block, all_lines, text_with_coords, debug):
        """Process a single RPO block"""
        rpo_lines = rpo_block["lines"]
        rpo_text = "\n".join(rpo_lines)
        
        # Pre-populate all global fields
        global_data = {field: "" for field in self.GLOBAL_FIELDS}
        global_data["PO #"] = rpo_block["rpo_number"]
        
        # Extract global data with enhanced patterns and fallbacks
        extracted_global = self.extract_global_data_enhanced(rpo_lines, rpo_text, text_with_coords, debug)
        global_data.update(extracted_global)
        
        # Split RPO block into item blocks
        item_blocks = self.split_rpo_into_item_blocks(rpo_lines, debug)
        debug.setdefault("state_transitions", []).append(f"RPO {rpo_block['rpo_number']}: Found {len(item_blocks)} item blocks")
        
        # Process each item block
        processed_items = []
        for item_block in item_blocks:
            item_result = self.process_item_block(item_block, rpo_block["start_line"], all_lines, text_with_coords, debug)
            if item_result:
                processed_items.append(item_result)
        
        return {
            "po_number": rpo_block["rpo_number"],
            "global": global_data,
            "items": processed_items,
            "item_count": len(processed_items),
            "component_count": sum(len(item.get("Components", [])) for item in processed_items)
        }

    def split_rpo_into_item_blocks(self, rpo_lines, debug):
        """Split RPO block into item blocks"""
        item_blocks = []
        current_block = {"item_number": None, "start_line": 0, "lines": []}
        
        # Enhanced item patterns
        item_patterns = [
            r'\*\*([A-Z]{2}\d{4}[A-Z0-9]+)\*\*',  # **ITEM**
            r'\b([A-Z]{2}\d{4}[A-Z0-9]+)\b(?=\s+[A-Z]{2}\d{3,6})',  # ITEM followed by vendor item
            r'^\s*([A-Z]{2}\d{4}[A-Z0-9]+)\s+',  # ITEM at line start
            r'^\s*([0-9]{5,}[A-Z]{2}[A-Z0-9]*)\s+',  # Numeric-prefix items
        ]
        
        for i, line in enumerate(rpo_lines):
            item_found = False
            
            for pattern in item_patterns:
                matches = re.finditer(pattern, line)
                for match in matches:
                    item_number = match.group(1).replace('O', '0').replace('B', '8')
                    
                    # Close previous item block
                    if current_block["item_number"] is not None:
                        current_block["end_line"] = i
                        item_blocks.append(current_block)
                        debug.setdefault("state_transitions", []).append(f"Closed item {current_block['item_number']} at line {i}")
                    
                    # Start new item block
                    current_block = {
                        "item_number": item_number,
                        "start_line": i,
                        "lines": [line],
                        "item_line": line
                    }
                    debug.setdefault("state_transitions", []).append(f"Started item {item_number} at line {i}")
                    item_found = True
                    break
                
                if item_found:
                    break
            
            if not item_found and current_block["item_number"] is not None:
                current_block["lines"].append(line)
        
        # Close final item block
        if current_block["item_number"] is not None:
            current_block["end_line"] = len(rpo_lines)
            item_blocks.append(current_block)
            debug.setdefault("state_transitions", []).append(f"Closed final item {current_block['item_number']}")
        
        return item_blocks

    def process_item_block(self, item_block, global_start_idx, all_lines, text_with_coords, debug):
        """Process a single item block"""
        # Pre-populate all item fields
        item = {field: "" for field in self.ITEM_FIELDS}
        item["Richline Item #"] = item_block["item_number"]
        item["Components"] = []
        
        item_lines = item_block["lines"]
        item_text = "\n".join(item_lines)
        item_line = item_block["item_line"]
        
        # Extract item data using enhanced methods
        self.extract_item_data_enhanced(item, item_line, item_text, debug)
        
        # Extract components with cross-page awareness
        item["Components"] = self.extract_components_state_machine(
            item_block, global_start_idx, all_lines, text_with_coords, debug
        )
        
        debug.setdefault("state_transitions", []).append(f"Item {item_block['item_number']}: {len(item['Components'])} components")
        
        return item

    # ===============================
    # GLOBAL DATA EXTRACTION
    # ===============================

    def extract_global_data_enhanced(self, rpo_lines, rpo_text, text_with_coords, debug):
        """Enhanced global data extraction with fallbacks"""
        global_data = {}
        
        # Location extraction
        location = self.extract_location_enhanced(rpo_lines, rpo_text, text_with_coords)
        if location:
            global_data["Location"] = location
        
        # Vendor extraction with enhanced multi-line support
        vendor_data = self.extract_vendor_data_enhanced(rpo_lines, rpo_text, text_with_coords)
        global_data.update(vendor_data)
        
        # Metal rates with positional fallback
        rates = self.extract_metal_rates_enhanced(rpo_lines, rpo_text, text_with_coords)
        global_data.update(rates)
        
        # Other fields with original patterns
        other_fields = self.extract_other_global_fields(rpo_text)
        global_data.update(other_fields)
        
        return global_data

    def extract_location_enhanced(self, rpo_lines, rpo_text, text_with_coords):
        """Enhanced location extraction"""
        location_patterns = [
            r"Location[:\s]*([A-Z]{2,4})(?:\s+(?:Vendor|Printed|Tel|\n|$))",
            r"Location[:\s]*([A-Z]{2,4})\s*\n",
            r"Location[:\s]*([A-Z]{2,4})\s*$",
            r"Ship\s+To[:\s]*.*?([A-Z]{3,4})",
            r"\b([A-Z]{3,4})\s+(?:WAREHOUSE|LOCATION|FACILITY)",
        ]
        
        for pattern in location_patterns:
            match = re.search(pattern, rpo_text, re.IGNORECASE | re.MULTILINE)
            if match:
                location = match.group(1).strip().upper()
                if len(location) >= 2 and location not in ['THE', 'AND', 'FOR', 'YOU', 'ARE', 'TEL', 'FAX', 'PO']:
                    return location
        
        # Positional fallback using coordinates if available
        if text_with_coords:
            for coord_data in text_with_coords:
                if (coord_data['y'] < 500 and  # Top area of page
                    re.match(r'^[A-Z]{3,4}$', coord_data['text']) and
                    coord_data['text'] not in ['THE', 'AND', 'FOR', 'YOU', 'ARE']):
                    return coord_data['text']
        
        return None

    def extract_vendor_data_enhanced(self, rpo_lines, rpo_text, text_with_coords):
        """Enhanced vendor extraction"""
        vendor_data = {}
        
        # Find Vendor ID line
        vendor_id_line_idx = None
        for i, line in enumerate(rpo_lines):
            if re.search(r"Vendor\s+ID", line, re.IGNORECASE):
                vendor_id_line_idx = i
                
                # Extract Vendor ID
                vendor_id_match = re.search(r"Vendor\s+ID\s*[:#]?\s*([A-Za-z0-9-]+)", line, re.IGNORECASE)
                if vendor_id_match:
                    vendor_data["Vendor ID #"] = vendor_id_match.group(1).strip()
                elif i + 1 < len(rpo_lines):
                    next_line = rpo_lines[i + 1].strip()
                    if next_line and len(next_line) < 30 and re.match(r'^[A-Za-z0-9-]+$', next_line):
                        vendor_data["Vendor ID #"] = next_line
                break
        
        # Enhanced vendor name extraction
        if vendor_id_line_idx is not None:
            vendor_name_parts = []
            skip_keywords = [
                "ID", "DATE", "PO", "ORDER", "LOCATION", "RATE", "TYPE", "#", "SHIP", "BILL",
                "SUPPLY", "CERT", "SEND", "POLICY", "DUE", "TO:", "UNIT", "VENDOR", "TEL", "FAX",
                "PHONE", "ADDRESS", "ZIP", "EMAIL"
            ]
            
            # Look through more lines for vendor name (up to 30 lines)
            for j in range(1, 30):
                if vendor_id_line_idx + j >= len(rpo_lines):
                    break
                    
                possible_name = rpo_lines[vendor_id_line_idx + j].strip()
                if not possible_name or len(possible_name) < 2:
                    continue
                
                # Enhanced skip logic
                if any(possible_name.upper().startswith(word) for word in skip_keywords):
                    break
                if re.match(r'^\d+$', possible_name):  # Skip pure numbers
                    continue
                if re.match(r'^\d{5}$', possible_name):  # Skip ZIP codes
                    continue
                if ":" in possible_name and len(possible_name.split(":")[0]) < 8:
                    continue
                if re.match(r'^[A-Z]{2}\d{4}', possible_name):  # Skip item numbers
                    break
                
                # Handle "Ship To:" cases
                if re.match(r'^Ship\s+To:', possible_name, re.IGNORECASE):
                    break
                    
                if "To:" in possible_name:
                    before_to = possible_name.split("To:")[0].strip()
                    if before_to and not before_to.upper().startswith("SHIP"):
                        vendor_name_parts.append(before_to)
                    break
                
                # Add to vendor name if it looks like a company name
                if not possible_name.upper().startswith("SHIP"):
                    # Additional validation for company names
                    if (len(possible_name) > 3 and 
                        not re.match(r'^\d+[\d\.\s]*$', possible_name) and  # Not just numbers
                        not possible_name.upper() in ["DUE", "DATE", "ORDER", "TYPE"] and
                        not re.match(r'^[A-Z]{1,3}$', possible_name)):  # Not short codes
                        vendor_name_parts.append(possible_name)
                
                # Stop at company indicators or obvious non-vendor content
                if (re.search(r'\b(LTD|LIMITED|INC|CORP|CORPORATION|PVT\.?\s*LTD|LLC)\b', possible_name, re.IGNORECASE) or
                    re.search(r'\b(TERMS|CONDITIONS|PAYMENT|DUE|DAYS)\b', possible_name, re.IGNORECASE)):
                    break
            
            if vendor_name_parts:
                vendor_full = " ".join(vendor_name_parts)
                # Clean up vendor name
                vendor_patterns = [
                    r"(.+?(?:LTD|LIMITED|INC|CORP|CORPORATION|PVT\.?\s*LTD|LLC)\.?)",
                    r"(.+)"
                ]
                
                for pattern in vendor_patterns:
                    match = re.search(pattern, vendor_full, re.IGNORECASE)
                    if match:
                        vendor_name = match.group(1).strip(" .:-")
                        # Remove common prefixes
                        vendor_name = re.sub(r'^(Shi\s+|Ship\s+|Vendor\s+)', '', vendor_name, flags=re.IGNORECASE)
                        if len(vendor_name) > 3:  # Minimum length check
                            vendor_data["Vendor Name"] = vendor_name
                        break
        
        return vendor_data

    def extract_metal_rates_enhanced(self, rpo_lines, rpo_text, text_with_coords):
        """Enhanced metal rates extraction"""
        rates = {}
        
        # Enhanced rate patterns
        rate_patterns = [
            r"Order Type\s+Gold\s+Platinum\s+Silver\s*\n.*?\b[A-Z]+\b\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)",
            r"\b[A-Z]+\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s*(?=\n|$)",
            r"([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s*\n.*?(?:ALL MDSE|TERMS)",
            r"Gold\s*[:\s]*([\d,]+\.?\d*)\s*Platinum\s*[:\s]*([\d,]+\.?\d*)\s*Silver\s*[:\s]*([\d,]+\.?\d*)",
            r"(?:STOCK|MCH|SPC)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)",
        ]
        
        for pattern in rate_patterns:
            rate_match = re.search(pattern, rpo_text, re.IGNORECASE | re.MULTILINE)
            if rate_match:
                rates["Gold Rate"] = rate_match.group(1).replace(',', '')
                rates["Platinum Rate"] = rate_match.group(2).replace(',', '')
                rates["Silver Rate"] = rate_match.group(3).replace(',', '')
                break
        
        return rates

    def extract_other_global_fields(self, rpo_text):
        """Extract other global fields"""
        fields = {}
        
        # Due Date patterns
        due_date_patterns = [
            r"Due Date[:\s]*([A-Za-z]+ \d{1,2},?\s+\d{4})",
            r"Due Date[:\s]*(\d{1,2}/\d{1,2}/\d{2,4})",
            r"\b(January \d{1,2},?\s+\d{4}|February \d{1,2},?\s+\d{4}|March \d{1,2},?\s+\d{4}|April \d{1,2},?\s+\d{4}|May \d{1,2},?\s+\d{4}|June \d{1,2},?\s+\d{4}|July \d{1,2},?\s+\d{4}|August \d{1,2},?\s+\d{4}|September \d{1,2},?\s+\d{4}|October \d{1,2},?\s+\d{4}|November \d{1,2},?\s+\d{4}|December \d{1,2},?\s+\d{4})\b"
        ]
        
        for pattern in due_date_patterns:
            match = re.search(pattern, rpo_text, re.IGNORECASE)
            if match:
                due_date = match.group(1)
                due_date = re.sub(r'\s+', ' ', due_date).strip()
                fields["Due Date"] = due_date
                break
        
        # Order Type
        order_types = [
            "STOCK", "MCH", "SPC", "ASSAY", "ASSET", "ASSETKM-AD", "CHARGEBACK", "CONFONLY", "CORRECT", "DNP",
            "DOTCOM", "DOTCOMB", "EXTEND", "FL-RECIEVE", "IGI", "MANUAL", "MC", "MCH-REV", "MST", "NEW-CLR",
            "PCM", "PKG", "PSAMPLE", "REP", "RMC", "RPR", "RTV", "SGI", "SHW", "SLD", "SLDSPC", "SMG", "SMP",
            "SMPGEM", "SPO-BUILD", "SUPPLY", "TST"
        ]
        
        order_type_patterns = [
            r"Order Type\s+Gold\s+Platinum\s+Silver\s*\n.*?\b(" + "|".join(order_types) + r")\b",
            r"\b(" + "|".join(order_types) + r")\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*",
            r"Terms\s+Order Type\s+Gold\s+Platinum\s+Silver\s*\n.*?\b(" + "|".join(order_types) + r")\b"
        ]
        
        for pattern in order_type_patterns:
            match = re.search(pattern, rpo_text, re.IGNORECASE | re.MULTILINE)
            if match:
                groups = match.groups()
                for group in reversed(groups):
                    if group and group.upper() in order_types:
                        fields["Order Type"] = group.upper()
                        break
                if "Order Type" in fields:
                    break
        
        # PO Date
        po_date_match = re.search(r"\b(\d{2}/\d{2}/\d{2,4})\b", rpo_text)
        if po_date_match:
            fields["PO Date"] = po_date_match.group(1)
        
        return fields

    # ===============================
    # ITEM EXTRACTION
    # ===============================

    def extract_item_data_enhanced(self, item, item_line, item_text, debug):
        """Enhanced item data extraction"""
        # Vendor Item extraction
        vendor_item = self.extract_vendor_item_enhanced(item_line, item_text)
        if vendor_item and vendor_item != item["Richline Item #"]:
            item["Vendor Item #"] = vendor_item
        
        # Job number extraction
        job_patterns = [r"(RFP\s*\d{6,})", r"(RSET\s*\d{6,})"]
        for pattern in job_patterns:
            match = re.search(pattern, item_text)
            if match:
                item["Job #"] = match.group(1).replace(" ", "")
                break
        
        # Metal extraction
        metal1, metal2 = self.extract_metal_from_description_fixed(item_line)
        if metal1:
            item["Metal 1"] = metal1
        if metal2:
            item["Metal 2"] = metal2
        
        # Financial data
        self.extract_item_financial_data_enhanced(item, item_text)
        
        # Technical data
        self.extract_item_technical_data(item, item_text)

    def extract_vendor_item_enhanced(self, item_line, item_text):
        """Vendor item extraction"""
        if not item_line or not item_text:
            return None

        try:
            # Pattern 1: Item/Vendor Style format
            item_vendor_pattern = r'Item \d+:\s*([A-Z0-9]+)\s+Vendor Style:\s*([A-Z0-9]+)'
            match = re.search(item_vendor_pattern, item_text)
            if match:
                return match.group(1)
            
            # Pattern 2: Extract from item line
            richline_match = re.search(r'^([A-Z0-9]+)', item_line)
            if richline_match:
                richline_item = richline_match.group(1)
                lines = item_text.split('\n')

                for line in lines:
                    if not line:
                        continue
                    line = line.strip()

                    vendor_match = re.match(r'^([A-Z0-9-]+)', line)
                    if vendor_match:
                        potential_vendor = vendor_match.group(1)

                        if (len(potential_vendor) < len(richline_item) and richline_item.startswith(potential_vendor)):
                            return potential_vendor
                        elif 5 <= len(potential_vendor) <= 15 and '-' in potential_vendor:
                            return potential_vendor
                        
                    if re.match(r'^[A-Z0-9]+$', line):
                        if len(line) < len(richline_item):
                            if richline_item.startswith(line):
                                return line
                            elif 5 <= len(line) <= 15:
                                return line
            
            return None
        
        except Exception as e:
            return None

    def extract_metal_from_description_fixed(self, description):
        """Metal extraction from description"""
        if not description:
            return None, None
        
        valid_metals = [
            '10K', '10KA', '10KB', '10KC', '10KD', '10KE', '10KF', '10KG', '10KH', '10KI', '10KJ', '10KK', 
            '10KL', '10KM', '10KN', '10KO', '10KP', '10KR', '10KS', '10KT', '10KW', '10KX', '10KY', 
            '14K', '14KA', '14KB', '14KC', '14KD', '14KE', '14KF', '14KG', '14KH', '14KI', '14KJ', '14KK', 
            '14KL', '14KM', '14KN', '14KO', '14KP', '14KR', '14KS', '14KT', '14KW', '14KX', '14KY', 
            '18K', '18KA', '18KB', '18KC', '18KD', '18KE', '18KF', '18KG', '18KH', '18KI', '18KJ', '18KK', 
            '18KL', '18KM', '18KN', '18KO', '18KP', '18KR', '18KS', '18KT', '18KW', '18KX', '18KY', 
            'SS', 'SILVER', 'GOLD', 'GOS', 'BRASS', 'BRONZE'
        ]
        
        description_upper = description.upper()

        # Check for bimetal patterns first
        bimetal_patterns = [
            r'\b(SS)\s*/\s*(10KY|10KW|10KR|14KY|14KW|14KR|18KY|18KW|18KR)\b',
            r'\b(10KY|10KW|10KR|14KY|14KW|14KR|18KY|18KW|18KR)\s*/\s*(SS)\b',
            r'\b(SILVER)\s*/\s*(GOLD|10K|14K|18K)\b',
            r'\b(GOLD|10K|14K|18K)\s*/\s*(SILVER)\b',
            r'\b(SS)\s+.*?\b(10KY|10KW|10KR|14KY|14KW|14KR|18KY|18KW|18KR)\b',
            r'\b(10KY|10KW|10KR|14KY|14KW|14KR|18KY|18KW|18KR)\s+.*?\b(SS)\b',
        ]
        
        for pattern in bimetal_patterns:
            match = re.search(pattern, description_upper)
            if match:
                metal1, metal2 = match.groups()
                if metal1 in valid_metals and metal2 in valid_metals:
                    return metal1, metal2
            
        # Find single metals
        found_metals = []
        for metal in valid_metals:
            patterns = [
                r'\b' + re.escape(metal) + r'\b',
                r'\b' + re.escape(metal) + r'(?=\s)',
                r'(?<=\s)' + re.escape(metal) + r'\b',
            ]
            for pattern in patterns:
                if re.search(pattern, description_upper):
                    if metal not in found_metals:
                        found_metals.append(metal)
                    break

        if len(found_metals) >= 2:
            return found_metals[0], found_metals[1]
        elif len(found_metals) == 1:
            return found_metals[0], None
        
        return None, None

    def extract_item_financial_data_enhanced(self, item, item_text):
        """Financial data extraction"""
        # Stone PC patterns including asterisk patterns
        stone_patterns = [
            r'Stone PC[:\s]+(\d+\.\d+)', 
            r'Stone\s+Labor[:\s]+(\d+\.\d+)', 
            r'Stone[:\s]+(\d+\.\d+)(?!\s*CT)',
            r'Stone Labor[:\s]+(\d+\.\d+)',
            r'\*+([0-9.]*)\*+',
        ]
        for pattern in stone_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                value = match.group(1)
                if not value or '*' in value:
                    item["Stone PC"] = "***********"
                else:
                    item["Stone PC"] = value
                break
        
        if "Stone PC" not in item:
            asterisk_match = re.search(r'\*{5,}', item_text)
            if asterisk_match:
                item["Stone PC"] = "***********"
        
        # Labor PC patterns
        labor_patterns = [
            r'Labor PC[:\s]+(\d+\.\d+)', 
            r'Labor[:\s]+PC[:\s]+(\d+\.\d+)',
            r'PC\s+Labor[:\s]+(\d+\.\d+)',
            r'Labor[:\s]+(\d+\.\d+)(?!\s*CT|GR|EA)',
        ]
        for pattern in labor_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                item["Labor PC"] = match.group(1)
                break

    def extract_item_technical_data(self, item, item_text):
        """Technical data extraction"""
        # Cast Fin Weight patterns
        cast_patterns = [
            r'CAST Fin WT[:\s]*Gold[:\s]*(\d+\.\d+)(?:\s*Silver[:\s]*(\d+\.\d+))?',
            r'CAST Fin Wt[:\s]*Gold[:\s]*(\d+\.\d+)(?:\s*Silver[:\s]*(\d+\.\d+))?',
            r'Fin WT[:\s]*Gold[:\s]*(\d+\.\d+)(?:\s*Silver[:\s]*(\d+\.\d+))?'
        ]
        for pattern in cast_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                if match.group(1):
                    item["Fin Weight (Gold)"] = match.group(1)
                if match.group(2):
                    item["Fin Weight (Silver)"] = match.group(2)
                break
        
        # Loss percentage patterns
        loss_patterns = [
            r'LOSS %[:\s]*Gold[:\s]*(\d+\.\d+)%?(?:\s*Silver[:\s]*(\d+\.\d+)%?)?',
            r'Loss[:\s]*Gold[:\s]*(\d+\.\d+)%?(?:\s*Silver[:\s]*(\d+\.\d+)%?)?'
        ]
        for pattern in loss_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                if match.group(1):
                    item["Loss % (Gold)"] = f"{match.group(1)}%"
                if match.group(2):
                    item["Loss % (Silver)"] = f"{match.group(2)}%"
                break

    # ===============================
    # COMPONENT EXTRACTION
    # ===============================

    def extract_components_state_machine(self, item_block, global_start_idx, all_lines, text_with_coords, debug):
        """FIXED: Enhanced component extraction with better cross-page logic"""
        components = []
        
        # First, try within item block
        components = self.extract_components_from_lines(item_block["lines"])
        if components:
            debug.setdefault("state_transitions", []).append(f"Found {len(components)} components in item block")
            return components
        
        # FIXED: More targeted cross-page search
        item_global_pos = global_start_idx + item_block["start_line"]
        
        # Look for component table within reasonable range
        search_start = max(0, item_global_pos - 10)  # Reduced range
        search_end = min(len(all_lines), item_global_pos + 30)  # Look ahead more
        
        # Find the exact component table for this item
        for i in range(search_start, search_end):
            line = all_lines[i]
            
            # Look for component table header
            if re.search(r'supplied\s+by\s+component', line, re.IGNORECASE):
                # Found component table, extract from here
                component_lines = all_lines[i:i+15]  # Next 15 lines
                components = self.extract_components_from_lines(component_lines)
                if components:
                    debug.setdefault("state_transitions", []).append(f"Found {len(components)} components via targeted search")
                    break
        
        return components

    def extract_components_from_lines(self, component_lines):
        """FIXED: Better component extraction for your specific format"""
        components = []
        in_component_section = False
        
        for line in component_lines:
            line = line.strip()
            if not line or len(line) < 5:
                continue

            line_lower = line.lower()
            
            # FIXED: Detect component section start
            if (re.search(r'supplied\s+by\s+component', line_lower) or
                re.search(r'component.*setting.*cost', line_lower) or
                line_lower.startswith('supplied by')):
                in_component_section = True
                continue
            
            # Stop conditions
            stop_conditions = [
                "there is a", "market price", "page:", "purchase order",
                "richline group", "total", "weight tolerance"
            ]
            if any(stop in line_lower for stop in stop_conditions):
                break
            
            # Skip if not in component section yet
            if not in_component_section:
                continue
            
            # Stop if we hit another item
            if re.search(r'\b[A-Z]{2}\d{4}[A-Z0-9]+\b.*\d+\.\d+.*(EA|PR)', line):
                break
                
            component = self.parse_component_line_fixed_for_your_format(line)
            if component and component.get("Component"):
                components.append(component)
        
        return components

    def parse_component_line_fixed_for_your_format(self, line):
        """FIXED: Parse component lines for your specific OCR format"""
        component = {
            "Component": "",
            "Cost ($)": "",
            "Tot. Weight": "",
            "Supply Policy": ""
        }

        line = line.strip()
        if not line or len(line) < 5:
            return None

        # Skip obvious non-component lines
        skip_patterns = [
            r'^\d+\s*$',  # Just numbers
            r'^setting\s+typ',  # Headers
            r'^qty\.\s+per',  # Headers
            r'^ext\.\s+weight',  # Headers
        ]
        
        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                return None

        # FIXED: Component patterns specific to your OCR
        component_patterns = [
            # For your format: DIA125, LDSH/.33, PKG05679, etc.
            r'\b(DIA\d+)\b',
            r'\b(LDSH/\.\d+)\b',
            r'\b(PKG\d+)\b', 
            r'\b(CUN\.UN\.\d+\.\d+)\b',
            r'\b([A-Z]+\d+[A-Z]*(?:/\.\d+)?)\b',  # General pattern
            r'\b([A-Z]{2,4}\d{2,6}[A-Z0-9]*)\b',
            # CS patterns - most common
            r'\b(CS[0-9/\.-]+(?:-[A-Z0-9]+)*)', r'\b(CS[A-Z0-9\./\-]+)',
            r'\b(CS\d+/\d+(?:\.\d+)?(?:NV|OV|PS|HS|RDP)-[A-Z0-9]+)',
            r'\b(CS\d+(?:/\d+(?:\.\d+)?)?-[A-Z0-9]+-[A-Z0-9]+)',
            r'\b(CS\d+(?:/\d+(?:\.\d+)?)?[A-Z]+-[A-Z0-9]+)', 
            # Other patterns
            r'\b(THP-WH\d+-[A-Z]+)', r'\b(THP-[A-Z0-9\-]+)', 
            r'\b([0-9]{2}XX[0-9]{4}-[A-Z0-9]+)', r'\b(SSC[0-9]+[A-Z0-9]*)', 
            r'\b(PKG[0-9]+)', r'\b(TRC[0-9]+[A-Z0-9]*)', 
            r'\b(CHR[A-Z0-9]+W?-\d+[A-Z]?)', r'\b(OT-[A-Z0-9]+)',
            r'\b(OT-CHR\d+-\d+)', r'\b(OT-[A-Z]+\d*)', 
            r'\b(R[0-9]{3}-[0-9]+[A-Z\-]*)', r'\b(CN\d{4}-[A-Z0-9]+-\d+)',
            r'\b(CN[0-9]{4}-[A-Z0-9\-]+)', r'\b(MS\d{4}-[A-Z0-9]+)', 
            r'\b(MS[0-9]{4}-[A-Z0-9]+)', r'\b(A\d+/\.\d+)',
            r'\b(A[0-9]+/\.[0-9]+)', r'\b(CHSBOXIML-\d+)',
            r'\b(CHS[A-Z]+ML-\d+)', r'\b(LD[A-Z]*[0-9]+[A-Z]*[/\.][0-9\.]+)', 
            r'\b(LD[A-Z]*[0-9]+/[0-9\.]+)', r'\b(LD[A-Z0-9\.]+)',   
            r'\b(PKG\d+)', r'\b([A-Z]+[\d/.]+)', r'\b(H[0-9]+/[0-9\.]+)',
            r'\b([A-Z]{1,4}\d{1,6}[A-Z0-9\./\-]*)',
            r'\b([A-Z]+\d+[A-Z]*(?:[/\.-][A-Z0-9]+)*)',
            r'\b([0-9]{4}-[A-Z]+-[A-Z]+)', r'\b([0-9]/[0-9\.]+-[A-Z]+-[A-Z]+)', 
            r'\b([0-9]/[0-9\.]+)', r'\b([0-9]+/[0-9\.]+[A-Z]*-[A-Z0-9]+)',
            # Fallback
        ]
        
        for pattern in component_patterns:
            match = re.search(pattern, line)
            if match:
                component["Component"] = match.group(1)
                break
        
        # FIXED: Extract cost and weight - specific to your format
        # Your format: "90.25 CT", "0.33333 CT", "0.02329 EA"
        cost_weight_pattern = r'(\d+\.?\d*)\s+(CT|EA|GR)'
        values = re.findall(cost_weight_pattern, line)
        
        if values:
            # First value is usually cost, second is weight
            if len(values) >= 1:
                component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"
            if len(values) >= 2:
                component["Tot. Weight"] = f"{values[1][0]} {values[1][1]}"
            elif len(values) == 1:
                # Single value - determine by magnitude
                val = float(values[0][0])
                if val > 5:  # Likely cost
                    component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"
                else:  # Likely weight
                    component["Tot. Weight"] = f"{values[0][0]} {values[0][1]}"
        
        # FIXED: Extract supply policy
        if re.search(r'by\s+vendor', line, re.IGNORECASE):
            component["Supply Policy"] = "By Vendor"
        elif re.search(r'send\s+to', line, re.IGNORECASE):
            component["Supply Policy"] = "Send To"
        
        return component if component["Component"] else None

    # ===============================
    # RESULT FORMATTING
    # ===============================

    def format_final_result(self, processed_rpos, debug):
        """Format the final result structure"""
        if len(processed_rpos) == 1:
            # Single RPO
            return processed_rpos[0]
        else:
            # Multiple RPOs
            return {
                "purchase_orders": processed_rpos,
                "summary": {
                    "total_pos": len(processed_rpos),
                    "total_items": sum(po["item_count"] for po in processed_rpos),
                    "total_components": sum(po["component_count"] for po in processed_rpos)
                }
            }

    # ===============================
    # FALLBACK & FAST MODE METHODS
    # ===============================

    def _extract_fast(self, pdf_file, debug):
        """Fast extraction fallback using original logic"""
        # Convert PDF with fast settings
        images = self.convert_pdf_to_image(pdf_file, dpi=self.fast_dpi, use_jpeg=True)
        if not images:
            return {"error": "Failed to convert PDF to images", "debug": debug}
        
        debug["processing_steps"].append(f"PDF converted to {len(images)} images (Fast mode)")

        # Parallel OCR processing
        all_text, all_lines = self.extract_text_simple(images)

        return self.process_text_fast(all_text, all_lines, debug)

    def process_text_fast(self, all_text, all_lines, debug):
        """Fast processing of text without coordinates, re-using original logic"""
        return self._process_extracted_text_original(all_text, all_lines, debug)

    def _process_extracted_text_original(self, all_text, all_lines, debug):
        """Original text processing method as fallback"""
        if not all_text:
            return {"error": "No extractable data found in PDF", "debug": debug}

        debug["ocr_text_length"] = len(all_text)

        # Extract global data
        global_data = self.extract_global_fields_enhanced_original(all_lines, all_text)
        debug["processing_steps"].append(f"Extracted {len(global_data)} global fields")

        # Find all RPOs and associate items
        items_by_rpo = self.find_items_with_rpo_association(all_lines)
        debug["processing_steps"].append(f"Found items for RPOs: {list(items_by_rpo.keys())}")

        # Process each RPO with its items
        purchase_orders = []
        
        for rpo_number, rpo_items in items_by_rpo.items():
            # Create RPO-specific global data
            rpo_global = global_data.copy()
            rpo_global["PO #"] = rpo_number
            
            # Process items for this RPO
            processed_items = []
            
            for idx, (item_line_idx, item_number, item_line) in enumerate(rpo_items):
                # Determine item text boundaries
                next_item_idx = rpo_items[idx + 1][0] if idx + 1 < len(rpo_items) else len(all_lines)
                item_lines = all_lines[item_line_idx:next_item_idx]
                
                # Extract complete item data
                item = self.extract_single_item_enhanced(
                    item_number, 
                    item_line, 
                    item_lines,
                    item_line_idx,
                    all_lines
                )
                
                if item:
                    processed_items.append(item)
            
            # Create RPO entry
            rpo_entry = {
                "po_number": rpo_number,
                "global": rpo_global,
                "items": processed_items,
                "item_count": len(processed_items),
                "component_count": sum(len(item.get("Components", [])) for item in processed_items)
            }
            
            purchase_orders.append(rpo_entry)
            debug["processing_steps"].append(f"Processed RPO {rpo_number}: {len(processed_items)} items")

        debug["total_rpos"] = len(purchase_orders)

        # Return appropriate structure
        return self.format_final_result(purchase_orders, debug)

    def extract_global_fields_enhanced_original(self, lines, full_text):
        """Original global field extraction method"""
        result = {}

        # Location extraction
        location_patterns = [
            r"Location[:\s]*([A-Z]{2,4})(?:\s+(?:Vendor|Printed|Tel|\n|$))",
            r"Location[:\s]*([A-Z]{2,4})\s*\n",
            r"Location[:\s]*([A-Z]{2,4})\s*$",
            r"Location[:\s]*([A-Z]{2,4})",  # Fallback
        ]
        
        for pattern in location_patterns:
            match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE)
            if match:
                location = match.group(1).strip().upper()
                if len(location) >= 2 and location not in ['THE', 'AND', 'FOR', 'YOU', 'ARE', 'TEL', 'FAX']:
                    result["Location"] = location
                    break

        # Original field extraction logic
        first_page_fields = ["PO #", "PO Date"]
        pages = full_text.split('#page')
    
        for field, pattern in self.global_patterns.items():
            if field == "Location":
                continue  # Already handled above
                
            match_found = False
            if field in first_page_fields and len(pages) > 1:
                first_page = pages[1] if len(pages) > 1 else full_text
                match = re.search(pattern, first_page, re.IGNORECASE)
                if match:
                    result[field] = match.group(1).replace(",", "")
                    match_found = True
        
            if not match_found:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    result[field] = match.group(1).replace(",", "")

        return result

    def find_items_with_rpo_association(self, lines):
        """Original multi-RPO handling"""
        # Step 1: Find all RPO numbers with their line positions
        rpo_positions = []
        for i, line in enumerate(lines):
            rpo_matches = re.findall(r'RPO\d+', line, re.IGNORECASE)
            for rpo in rpo_matches:
                rpo_positions.append((i, rpo.upper()))
        
        # Step 2: Find all items with their line positions
        item_patterns = [
            r'\*\*([A-Z]{2}\d{4}[A-Z0-9]+)\*\*',
            r'\b([A-Z]{2}\d{4}[A-Z0-9]+)\b(?=\s+[A-Z]{2}\d{3,6})',
            r'^\s*([A-Z]{2}\d{4}[A-Z0-9]+)\s+',
            r'^\s*([0-9]{5,}[A-Z]{2}[A-Z0-9]*)\s+',
        ]
        
        item_positions = []
        for i, line in enumerate(lines):
            for pattern in item_patterns:
                matches = re.finditer(pattern, line)
                for match in matches:
                    item_number = match.group(1).replace('O', '0').replace('B', '8')
                    item_positions.append((i, item_number, line))
        
        # Remove duplicate items
        seen_items = set()
        unique_items = []
        for pos in item_positions:
            if pos[1] not in seen_items:
                seen_items.add(pos[1])
                unique_items.append(pos)
        
        unique_items.sort(key=lambda x: x[0])
        
        # Step 3: Associate each item with the nearest preceding RPO
        items_by_rpo = {}
        
        for item_line, item_number, item_text in unique_items:
            closest_rpo = None
            closest_distance = float('inf')
            
            for rpo_line, rpo_number in rpo_positions:
                if rpo_line <= item_line:
                    distance = item_line - rpo_line
                    if distance < closest_distance:
                        closest_distance = distance
                        closest_rpo = rpo_number
            
            if closest_rpo is None and rpo_positions:
                closest_rpo = rpo_positions[0][1]
            
            if closest_rpo:
                if closest_rpo not in items_by_rpo:
                    items_by_rpo[closest_rpo] = []
                items_by_rpo[closest_rpo].append((item_line, item_number, item_text))
        
        return items_by_rpo

    def extract_single_item_enhanced(self, item_number, item_line, item_lines, global_start_idx, all_lines):
        """Original item extraction method"""
        item = {"Components": [], "CAST Fin WT": {}, "LOSS %": {},  "Richline Item #": item_number}
        
        item_text = "\n".join(item_lines)
        
        # Extract vendor item
        vendor_item = self.extract_vendor_item_enhanced(item_line, item_text)
        if vendor_item and vendor_item != item_number:
            if len(item_number) > len(vendor_item):
                item["Richline Item #"] = item_number
                item["Vendor Item #"] = vendor_item
            else:
                item["Richline Item #"] = vendor_item
                item["Vendor Item #"] = item_number
        else:
            item["Richline Item #"] = item_number

        # Extract job number
        job_patterns = [r"(RFP\s*\d{6,})", r"(RSET\s*\d{6,})"]
        for pattern in job_patterns:
            match = re.search(pattern, item_text)
            if match:
                job_number = match.group(1).replace(" ", "")
                item["Job #"] = job_number
                break

        # Extract metals
        metal1, metal2 = self.extract_metal_from_description_fixed(item_line)
        if not metal1 and not metal2:
            metal1, metal2 = self.extract_metal_from_description_fixed(item_text)

        if metal1:
            item["Metal 1"] = metal1
        if metal2:
            item["Metal 2"] = metal2
        
        # Extract financial, technical, and physical data
        self.extract_item_financial_data_enhanced(item, item_text)
        self.extract_item_technical_data(item, item_text)
        
        # Components extraction
        item["Components"] = self.extract_components_enhanced_fixed(item_lines, global_start_idx, all_lines)
        
        return item

    def extract_components_enhanced_fixed(self, item_lines, global_start_idx, all_lines):
        """Original component extraction method"""
        components = []

        components = self.extract_components_from_lines(item_lines)

        if components:
            return components

        search_ranges = [
            (max(0, global_start_idx - 100), global_start_idx + 5),
            (max(0, global_start_idx - 50), global_start_idx + 10),
            (max(0, global_start_idx - 25), global_start_idx + 15),
        ]
        for start_idx, end_idx in search_ranges:
            search_lines = all_lines[start_idx:min(end_idx, len(all_lines))]

            component_start = -1
        
            header_patterns = [
                r'supplied by.*component.*cost',
                r'^\s*\|\s*Supplied by\s*\|',
                r'Component\s+Setting Typ\s+Qty',
                r'Component\s+Cost\s+\$\s+Tot',
                r'Supplied by\s+Component\s+Setting',
                r'\|\s*Supplied by\s*\|\s*Component',
                r'Component.*Cost.*Weight',
                r'Component.*Setting.*Cost.*Weight'
            ]

            for i, line in enumerate(search_lines):
                for pattern in header_patterns:
                    if re.search(pattern, line, re.IGNORECASE):
                        component_start = i + 1
                        break
                if component_start != -1:
                    break
        
            if component_start != -1:
                component_lines = search_lines[component_start:component_start + 20]
                components = self.extract_components_from_lines(component_lines)
                if components:
                    return components
            
        return components


# Example usage function
def main():
    """Example usage with enhanced state machine extractor"""
    extractor = HybridPDFOCRExtractor()
    
    pdf_path = "path_to_your_pdf.pdf"
    
    try:
        with open(pdf_path, 'rb') as pdf_file:
            print("--- Running Standard Extraction ---")
            result = extractor.extract(pdf_file)
            
            if "error" not in result:
                print("✅ Standard Extraction successful!")
                print(f"🎯 Accuracy Score: {result.get('accuracy', {}).get('accuracy_score', 'N/A'):.2f}")
            else:
                print("❌ Extraction failed:", result['error'])

        print("\n" + "="*40 + "\n")

        with open(pdf_path, 'rb') as pdf_file:
            print("--- Running Fast-Only Extraction ---")
            fast_result = extractor.extract_fast_only(pdf_file)
            
            if "error" not in fast_result:
                print("✅ Fast-Only Extraction successful!")
            else:
                print("❌ Fast extraction failed:", fast_result['error'])
        
        print("\n" + "="*40 + "\n")

        with open(pdf_path, 'rb') as pdf_file:
            print("--- Running Extraction with Timing ---")
            timed_result = extractor.extract_with_timing(pdf_file)
            if "error" in timed_result:
                 print("❌ Timed extraction failed:", timed_result['error'])
            else:
                 print("✅ Timed Extraction successful!")


    except FileNotFoundError:
        print(f"❌ File not found: {pdf_path}")
    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")

if __name__ == "__main__":
    main()
import re
import pytesseract
import pdf2image
import cv2
import numpy as np
from PIL import Image
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Samuel Aaron\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

class PDFOCRExtractor:
    def __init__(self):
        self.global_patterns = {
            "PO #": r"(RPO\d+)",
            "PO Date": r"\b(\d{2}/\d{2}/\d{2,4})\b",
            "Location": r"Location[:\s]*([A-Z]{2})",
        }
        self.job_pattern = r"(RFP\d{6,}|RSET\d{6,})"

    def convert_pdf_to_image(self, pdf_file):
        poppler_path = r"C:\Users\Samuel Aaron\Documents\Release-24.08.0-0\poppler-24.08.0\Library\bin"
        try:
            pdf_file.seek(0)
            images = pdf2image.convert_from_bytes(
                pdf_file.read(),
                dpi=300,
                poppler_path=poppler_path,
            )
            return images if images else None
        except Exception as e:
            print(f"PDF to Image Conversion FAILED: {e}")
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

    def extract_global_fields(self, lines, full_text):
        result = {}

        # Extract basic patterns
        for field, pattern in self.global_patterns.items():
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                result[field] = match.group(1).replace(",", "")

        # Extract Due Date - FIXED to capture full date
        due_date_patterns = [
            r"Due Date[:\s]*([A-Za-z]+ \d{1,2},?\s+\d{4})",  # "Due Date June 20, 2025"
            r"Due Date[:\s]*(\d{1,2}/\d{1,2}/\d{2,4})",
            r"\b(January \d{1,2},?\s+\d{4}|February \d{1,2},?\s+\d{4}|March \d{1,2},?\s+\d{4}|April \d{1,2},?\s+\d{4}|May \d{1,2},?\s+\d{4}|June \d{1,2},?\s+\d{4}|July \d{1,2},?\s+\d{4}|August \d{1,2},?\s+\d{4}|September \d{1,2},?\s+\d{4}|October \d{1,2},?\s+\d{4}|November \d{1,2},?\s+\d{4}|December \d{1,2},?\s+\d{4})\b"
        ]
        
        for pattern in due_date_patterns:
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                due_date = match.group(1)
                due_date = re.sub(r'\s+', ' ', due_date).strip()
                result["Due Date"] = due_date
                break

        # Extract Order Type - Enhanced with table context
        order_types = [
            "STOCK", "MCH", "SPC", "ASSAY", "ASSET", "ASSETKM-AD", "CHARGEBACK", 
            "CONFONLY", "CORRECT", "DNP", "DOTCOM", "DOTCOMB", "EXTEND", 
            "FL-RECIEVE", "IGI", "MANUAL", "MC", "MCH-REV", "MST", "NEW-CLR", 
            "PCM", "PKG", "PSAMPLE", "REP", "RMC", "RPR", "RTV", "SGI", "SHW", 
            "SLD", "SLDSPC", "SMG", "SMP", "SMPGEM", "SPC", "SPO-BUILD", "SUPPLY", "TST"
        ]
        
        # Look for order type in the table with Gold/Platinum/Silver rates
        order_type_patterns = [
            r"Order Type\s+Gold\s+Platinum\s+Silver\s*\n.*?\b(" + "|".join(order_types) + r")\b",
            r"\b(" + "|".join(order_types) + r")\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*",
            r"Terms\s+Order Type\s+Gold\s+Platinum\s+Silver\s*\n.*?\b(" + "|".join(order_types) + r")\b"
        ]
        
        for pattern in order_type_patterns:
            match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE)
            if match:
                # Get the last group that contains the order type
                groups = match.groups()
                for group in reversed(groups):
                    if group and group.upper() in order_types:
                        result["Order Type"] = group.upper()
                        break
                if "Order Type" in result:
                    break

        # Extract Metal Rates (Gold, Platinum, Silver) - Enhanced
        rate_patterns = [
            r"Order Type\s+Gold\s+Platinum\s+Silver\s*\n.*?\b[A-Z]+\b\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)",
            r"\b[A-Z]+\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s*(?=\n|$)",
            r"([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s*\n.*?ALL MDSE MUST BE STAMPED"
        ]
        
        for pattern in rate_patterns:
            rate_match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE)
            if rate_match:
                result["Gold Rate"] = rate_match.group(1).replace(',', '')
                result["Platinum Rate"] = rate_match.group(2).replace(',', '')
                result["Silver Rate"] = rate_match.group(3).replace(',', '')
                break

        # Extract Vendor ID and Name
        vendor_extracted = False
        for i, line in enumerate(lines):
            if "Vendor ID" in line and not vendor_extracted:
                # Extract Vendor ID
                vendor_id_match = re.search(r"Vendor ID\s*[:#]?\s*([A-Za-z0-9-]+)", line, re.IGNORECASE)
                if vendor_id_match:
                    result["Vendor ID #"] = vendor_id_match.group(1).strip()
                elif i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if next_line and len(next_line) < 30 and re.match(r'^[A-Za-z0-9-]+$', next_line):
                        result["Vendor ID #"] = next_line

                # Extract vendor name
                vendor_name_parts = []
                skip_keywords = [
                    "ID", "DATE", "PO", "ORDER", "LOCATION", "RATE", "TYPE", "#",
                    "SHIP", "BILL", "SUPPLY", "CERT", "SEND", "POLICY", "DUE", "TO:", "UNIT"
                ]

                for j in range(1, 10):
                    if i + j < len(lines):
                        possible_name = lines[i + j].strip()
                        if not possible_name or len(possible_name) < 2:
                            continue
                        
                        if any(possible_name.upper().startswith(word) for word in skip_keywords):
                            break
                        if re.match(r'^\d+$', possible_name):
                            continue
                        if ":" in possible_name and len(possible_name.split(":")[0]) < 8:
                            continue
                        
                        if re.match(r'^Ship\s+To:', possible_name, re.IGNORECASE):
                            break
                            
                        if "To:" in possible_name:
                            before_to = possible_name.split("To:")[0].strip()
                            if before_to and not before_to.upper().startswith("SHIP"):
                                vendor_name_parts.append(before_to)
                            break
                        
                        if not possible_name.upper().startswith("SHIP"):
                            vendor_name_parts.append(possible_name)
                        
                        if re.search(r'\b(LTD|LIMITED|INC|CORP|CORPORATION|PVT\.?\s*LTD)\.?\b', possible_name, re.IGNORECASE):
                            break

                if vendor_name_parts:
                    vendor_full = " ".join(vendor_name_parts)
                    vendor_patterns = [
                        r"(.+?(?:LTD|LIMITED|INC|CORP|CORPORATION|PVT\.?\s*LTD)\.?)",
                        r"(.+)"
                    ]
                    
                    for pattern in vendor_patterns:
                        match = re.search(pattern, vendor_full, re.IGNORECASE)
                        if match:
                            vendor_name = match.group(1).strip(" .:-")
                            vendor_name = re.sub(r'^Shi\s+', '', vendor_name, flags=re.IGNORECASE)
                            vendor_name = re.sub(r'^Ship\s+', '', vendor_name, flags=re.IGNORECASE)
                            result["Vendor Name"] = vendor_name
                            break
                
                vendor_extracted = True
                break

        return result

    def extract_items_from_text(self, lines):
        """Enhanced item extraction with better table parsing"""
        items = []
        
        # Find all item numbers first - ENHANCED PATTERNS
        item_patterns = [
            r'\*\*([A-Z]{2}\d{4}[A-Z0-9]+)\*\*',  # Bold item numbers
            r'\b([A-Z]{2}\d{4}[A-Z0-9]+)\b(?=\s+[A-Z]{2}\d{3,6})',  # Item with vendor style
            r'^\s*([A-Z]{2}\d{4}[A-Z0-9]+)\s+',
            r'^\s*([0-9]{5,}[A-Z]{2}[A-Z0-9]*)\s+',  # Line starting with item number
        ]
        
        item_locations = []
        for i, line in enumerate(lines):
            for pattern in item_patterns:
                matches = re.finditer(pattern, line)
                for match in matches:
                    item_number = match.group(1)
                    # Fix O to 0 in item numbers
                    item_number = item_number.replace('O', '0').replace('B', '8')
                    item_locations.append((i, item_number, line))
        
        # Remove duplicates and sort by line number
        seen_items = set()
        unique_items = []
        for loc in item_locations:
            if loc[1] not in seen_items:
                seen_items.add(loc[1])
                unique_items.append(loc)
        
        unique_items.sort(key=lambda x: x[0])
        
        # Process each item
        for idx, (line_idx, item_number, item_line) in enumerate(unique_items):
            # Determine item boundary
            next_item_idx = unique_items[idx + 1][0] if idx + 1 < len(unique_items) else len(lines)
            
            item = self.extract_single_item_enhanced(
                item_number, 
                item_line, 
                lines[line_idx:next_item_idx],
                line_idx,
                lines
            )
            
            if item:
                items.append(item)
        
        return items

    def extract_vendor_item_from_ocr(self, item_line, item_text):
        """Extract vendor item number based on OCR structure"""
        
        if not item_line or not item_text:
            print("DEBUG - item_line or item_text is None/empty")
            return None

        print(f"DEBUG - item_line: {item_line}")
        print(f"DEBUG - item_text first 200 chars: {item_text[:200]}")

        try:
            item_vendor_pattern = r'Item \d+:\s*([A-Z0-9]+)\s+Vendor Style:\s*([A-Z0-9]+)'
            match = re.search(item_vendor_pattern, item_text)
            if match:
                print(f"DEBUG - Pattern 0 matched: {match.group(1)}")
                return match.group(1)
            
            richline_match = re.search(r'^([A-Z0-9]+)', item_line)
            if richline_match:
                richline_item = richline_match.group(1)
                print(f"DEBUG - Richline item extracted: {richline_item}")
                lines = item_text.split('\n')
                print(f"DEBUG - Lines in item_text: {lines}")

                for line in lines:
                    if not line:
                        continue
                    line = line.strip()
                    print(f"DEBUG - Checking line: '{line}'")

                    vendor_match = re.match(r'^([A-Z0-9-]+)', line)
                    if vendor_match:
                        potential_vendor = vendor_match.group(1)
                        print(f"DEBUG - Potential vendor extracted: {potential_vendor}")

                        if (len(potential_vendor) < len(richline_item) and richline_item.startswith(potential_vendor)):
                            print(f"DEBUG - Pattern 1 matched: {potential_vendor}")
                            return potential_vendor
                        elif 5 <= len(potential_vendor) <= 15 and '-' in potential_vendor:
                            print(f"DEBUG - Pattern 1 matched (hyphenated): {potential_vendor}")
                            return potential_vendor
                        
                    if re.match(r'^[A-Z0-9]+$', line):
                        print(f"DEBUG - Line matches alphanumeric pattern")
                        if len(line) < len(richline_item):
                            print(f"DEBUG - Line is shorter than richline item")
                            if richline_item.startswith(line):
                                print(f"DEBUG - Pattern 1 matched: {line}")
                                return line
                            elif 5 <= len(line) <= 15:
                                print(f"DEBUG - Pattern 1 matched (non-prefix): {line}")
                                return line
                            else:
                                print(f"DEBUG - Richline item doesn't start with this line")
                        else:
                            print(f"DEBUG - Line is not shorter than richline item")
                    else:
                        print(f"DEBUG - Line doesn't match alphanumeric pattern")
            
            vendor_pattern1 = r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([A-Z]{2}\d{3,6}[A-Z0-9]*)'
            match = re.search(vendor_pattern1, item_line)
            if match:
                print(f"DEBUG - Pattern 2 matched: {match.group(1)}")
                return match.group(1)   
            
            vendor_patterns = [
                r'Vendor Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)',
                r'\b([A-Z]{2}\d{3,6}[A-Z0-9]*)\b(?=\s+\d+\.\d+\s+[A-Z]+)',
                r'Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
            ]

            for i, pattern in enumerate(vendor_patterns):
                match = re.search(pattern, item_text)
                if match:
                    print(f"DEBUG - Pattern 3.{i} matched: {match.group(1)}")
                    return match.group(1)
            print("DEBUG - Entering Universal Pattern")
            if richline_match:
                richline_item = richline_match.group(1)
                print(f"DEBUG - Universal Pattern: richline_item = {richline_item}")
                lines = item_text.split('\n')

                candidates = []
                for line in lines:
                    if not line:
                        continue
                    line = line.strip()
                    print(f"DEBUG - Universal checking: '{line}', length: {len(line)}")

                    if (re.match(r'^[A-Z0-9-]+$', line) and 5 <= len(line) <= 15 and line != richline_item): 
                        candidates.append(line)
                        print(f"DEBUG - Universal candidate found: {line}")

                print(f"DEBUG - Universal candidates: {candidates}")
                if candidates:
                    print(f"DEBUG - Universal pattern matched: {candidates[0]}")
                    return candidates[0]
            else:
                print("DEBUG - No richline_match for Universal Pattern")    
            print("DEBUG - No patterns matched")
            return None
        
        except Exception as e:
            print(f"DEBUG - Error in extract_vendor_item_from_ocr: {e}")
            return None                       


    def extract_size_from_description(self, description):
        """Extract size from metal description"""
        if not description:
            return None
            
        # Size patterns from actual OCR data
        size_patterns = [
            r'^(\d+\.\d+)\s+',  # "7.00 10KY BAND..."
            r'^(\d+)\s+',       # "18 10KW PEN..."
            r'(\d+\.\d+)(?:\s+[A-Z]+)',  # Size followed by metal
        ]
        
        for pattern in size_patterns:
            match = re.search(pattern, description)
            if match:
                return f"SIZE {match.group(1)}"
        
        return None

    def extract_metal_from_description(self, description):
        """Extract metal information from description"""
        if not description:
            return None, None
        
        valid_metals = [
        '02.4K', '02.4KR', '02.4KY', '02KT', '03.5KT', '03KT', '04KT', '06KR', '06KT', '06KW', '06KY', 
        '08KP', '08KT', '08KW', '08KY', '09KB', '09KP', '09KT', '09KW', '09KY', '100P', '10K', '10KA', 
        '10KB', '10KC', '10KD', '10KE', '10KF', '10KG', '10KH', '10KI', '10KJ', '10KK', '10KL', '10KM', 
        '10KN', '10KO', '10KP', '10KR', '10KS', '10KT', '10KW', '10KX', '10KY', '14', '14K', '14KA', 
        '14KB', '14KC', '14KD', '14KE', '14KF', '14KG', '14KH', '14KI', '14KJ', '14KK', '14KL', '14KM', 
        '14KN', '14KO', '14KP', '14KR', '14KS', '14KT', '14KW', '14KX', '14KY', '14S', '14TT', '18GG', 
        '18K', '18KA', '18KB', '18KC', '18KD', '18KE', '18KF', '18KG', '18KH', '18KI', '18KJ', '18KK', 
        '18KL', '18KM', '18KN', '18KO', '18KP', '18KR', '18KS', '18KT', '18KW', '18KX', '18KY', '1KR', 
        '1KW', '1KY', '21K', '22K', '22KY', '24K', '24KT', '24KTY', '24KW', '24KY', '3KR', '3KW', '3KY', 
        '585P', '8K', '8KW', '8KY', '9K', '9KR', '9KT', '9KW', '9KX', '9KY', 'BRASS', 'BRONZE', 'CB', 
        'GF', 'GOS', 'GF0', 'GF00', 'GF04', 'GF05', 'GF4', 'GF44', 'GF88', 'GFT4', 'GP', 'NIP', 'NIS', 
        'NO', 'NON', 'NONY', 'P', 'P10I', 'PD', 'PN', 'RH', 'S0', 'SS', 'SSF', 'STS', 'T', 'V', 'Y'
    ]
        description_upper = description.upper()  

        # Enhanced metal patterns based on OCR data
        bimetal_pattern = r'\b([A-Z0-9]+)\s*/\s*([A-Z0-9]+)\b'
        bimetal_matches = re.findall(bimetal_pattern, description_upper)
        
        for metal1_candidate, metal2_candidate in bimetal_matches:
            if metal1_candidate in valid_metals and metal2_candidate in valid_metals:
                return metal1_candidate, metal2_candidate
            
        found_metals = []

        for metal in valid_metals:
            if re.search(r'\b' + re.escape(metal) + r'\b', description_upper):
                if metal not in found_metals:
                    found_metals.append(metal)

        if found_metals:
            return found_metals[0], None
        
        return None, None

    def extract_single_item_enhanced(self, item_number, item_line, item_lines, global_start_idx, all_lines):
        """Extract single item with all required fields"""
        item = {"Components": [], "CAST Fin WT": {}, "LOSS %": {},  "Richline Item #": item_number}

        
        # Get full item text for pattern matching
        item_text = "\n".join(item_lines)
        
        # Parse the main item line for basic data - ENHANCED
        self.parse_item_table_row(item, item_line)
        
        # Extract vendor item number - ENHANCED
        vendor_item = self.extract_vendor_item_from_ocr(item_line, item_text)
        if vendor_item and vendor_item != item_number:
            if len(item_number) > len(vendor_item):
                item["Richline Item #"] = item_number
                item["Vendor Item #"] = vendor_item
            else:
                item["Richline Item #"] = vendor_item
                item["Vendor Item #"] = item_number
        else:
            item["Richline Item #"] = item_number

        size_match = re.search(r'\b(\d{1,2}\.\d{2})\b', item_text)
        if size_match:
            item["Size"] = f"SIZE {size_match.group(1)}"

        # metal_match = re.findall(r'\b(10K[YWR]|14K[YWR]|18K[YWR]|SS|GOS|SILVER|GOLD)\b', item_text.upper())
        # if metal_match:
        #     item["Metal 1"] = metal_match[0]
        #     if len(metal_match) > 1:
        #         item["Metal 2"] = metal_match[1]

        job_match = re.search(r'\b(RSET\d{6}|RFP\d{6})\b', item_text)
        if job_match:
            item["Job #"] = job_match.group(1)

        # Extract size from description - NEW
        size = self.extract_size_from_description(item.get("Metal Description", ""))
        if size:
            item["Size"] = size
        
        # Extract metal information from description - ENHANCED
        metal1, metal2 = self.extract_metal_from_description(item.get("Metal Description", ""))
        if metal1:
            item["Metal 1"] = metal1
        if metal2:
            item["Metal 2"] = metal2
        
        # Extract Job # - ENHANCED to find job numbers for each item more accurately
        job_patterns = [r"(RFP\s*\d{6,})", r"(RSET\s*\d{6,})"]
        
        print(f"DEBUG - Looking for job # in item: {item.get('Richline Item #', 'Unknown')}")
        print(f"DEBUG - Item text first 300 chars: {item_text[:300]}")

        # Look for job number in the item's context
        for pattern in job_patterns:
            # First, look in the immediate item context
            match = re.search(pattern, item_text)
            if match:
                job_number = match.group(1).replace("", "")
                item["Job #"] = job_number
                print(f"DEBUG - Job # found: {job_number}")
                break
        else:
            print(f"DEBUG - No match for pattern: {pattern}")
        # If not found in item context, look in a broader window around this item
        if "Job #" not in item:
            print("DEBUG - No Job # found for this item")
            for pattern in job_patterns:
                # Look in a window around this item
                start_idx = max(0, global_start_idx - 15)
                end_idx = min(len(all_lines), global_start_idx + 40)
                extended_text = "\n".join(all_lines[start_idx:end_idx])
                
                # Find all matches and get the one closest to current item
                matches = []
                for match in re.finditer(pattern, extended_text):
                    match_line = extended_text[:match.start()].count('\n')
                    distance = abs(match_line - (global_start_idx - start_idx))
                    matches.append((match.group(1), distance))
                
                if matches:
                    # Sort by distance and take the closest
                    matches.sort(key=lambda x: x[1])
                    item["Job #"] = matches[0][0]
                    break

        # If still no vendor item found, use fallback method
        if "Vendor Item #" not in item:
            # Look in the item line for vendor item patterns
            vendor_item_line_pattern = r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
            vendor_match = re.search(vendor_item_line_pattern, item_line)
            if vendor_match:
                vendor_item = vendor_match.group(1)
                if vendor_item != item_number:
                    item["Vendor Item #"] = vendor_item
            
            # If not found in item line, look in item text
            if "Vendor Item #" not in item:
                vendor_patterns = [
                    r'Vendor Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)',
                    r'\b([A-Z]{2}\d{3,6}[A-Z0-9-]*)\b(?!\s*\d+\.\d+)',
                    r'Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
                ]
                for pattern in enumerate(vendor_patterns):
                    match = re.search(pattern, item_text)
                    if match:
                        print(f"DEBUG - Pattern 3.{i} matched: {match.group(1)}")
                        return match.group(1)
                
                print("DEBUG - No patterns matched")
                return None

        # Extract detailed fields
        self.extract_item_financial_data(item, item_text)
        self.extract_item_technical_data(item, item_text)
        self.extract_item_physical_data(item, item_text)
        
        # Extract components
        item["Components"] = self.extract_components_enhanced(item_lines, global_start_idx, all_lines)
        
        return item

    def parse_item_table_row(self, item, item_line):
        """Enhanced item table row parsing based on actual OCR structure"""
        
        # Extract description - ENHANCED patterns based on actual OCR format
        desc_patterns = [
            # Pattern: SIZE METAL DESCRIPTION
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([0-9.]+\s+[A-Z/]+\s+[^|]+?)(?=\s+[\d.]+\s+\d+\s+EA)',
            # Fallback pattern
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+(.+?)(?=\s+[\d.]+\s+\d+\s+(?:EA|PR))',
            # Another fallback
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+(.+?)(?=\s+[\d.]+)',
            # Original pattern as fallback
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([^0-9]+?)(?=\s+\d+\.\d+)'
        ]
        
        for pattern in desc_patterns:
            desc_match = re.search(pattern, item_line)
            if desc_match:
                description = desc_match.group(1).strip()
                # Clean up the description
                description = re.sub(r'\s+', ' ', description)
                description = re.sub(r'\s*\|\s*', ' ', description)  # Remove table separators
                if len(description) > 5:  # Ensure meaningful description
                    item["Metal Description"] = description
                    break

        # Extract Pieces/Carats and Ext. Gross Wt. - ENHANCED for actual format
        # Pattern: ITEM_NUMBER DESCRIPTION UNIT_COST PIECES UOM EXT_GROSS_WT EXT_COST
        table_patterns = [
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+.+?\s+([\d.]+)\s+(\d+)\s+(EA|PR)\s+([\d.]+)',
            r'(\d+)\s+(EA|PR)\s+(\d+\.\d+)',  # Original pattern as fallback
        ]
        
        for pattern in table_patterns:
            table_match = re.search(pattern, item_line)
            if table_match:
                if len(table_match.groups()) == 4:  # First pattern
                    item["Pieces/Carats"] = table_match.group(2)
                    item["Ext. Gross Wt."] = f"{table_match.group(4)} GR"
                elif len(table_match.groups()) == 3:  # Second pattern
                    item["Pieces/Carats"] = table_match.group(1)
                    item["Ext. Gross Wt."] = f"{table_match.group(3)} GR"
                break

    def extract_item_financial_data(self, item, item_text):
        """Extract financial data from item text"""
        
        # Stone Labor (Stone PC)
        stone_patterns = [
            r'Stone PC[:\s]+(\d+\.\d+)',
            r'Stone[:\s]+(\d+\.\d+)',
            r'Stone Labor[:\s]+(\d+\.\d+)'
        ]
        
        for pattern in stone_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                item["Stone Labor"] = match.group(1)
                break
        
        # Labor PC
        labor_patterns = [
            r'Labor PC[:\s]+(\d+\.\d+)',
            r'Labor[:\s]+(\d+\.\d+)'
        ]
        
        for pattern in labor_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                item["Labor PC"] = match.group(1)
                break

    def extract_item_technical_data(self, item, item_text):
        """Extract technical data from item text"""
        
        # CAST Fin WT
        cast_patterns = [
            r'CAST Fin WT[:\s]*Gold[:\s]*(\d+\.\d+)(?:\s*Silver[:\s]*(\d+\.\d+))?',
            r'CAST Fin Wt[:\s]*Gold[:\s]*(\d+\.\d+)(?:\s*Silver[:\s]*(\d+\.\d+))?',
            r'Fin WT[:\s]*Gold[:\s]*(\d+\.\d+)(?:\s*Silver[:\s]*(\d+\.\d+))?'
        ]
        
        for pattern in cast_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                if match.group(1):
                    item["CAST Fin WT"]["Gold"] = match.group(1)
                    item["Fin Weight (Gold)"] = match.group(1)
                if match.group(2):
                    item["CAST Fin WT"]["Silver"] = match.group(2)
                    item["Fin Weight (Silver)"] = match.group(2)
                break
        
        # Enhanced silver weight extraction - NEW
        if "Silver" not in item["CAST Fin WT"]:
            silver_patterns = [
                r'CAST Fin WT[:\s]*.*?Silver[:\s]*(\d+\.\d+)',
                r'Fin WT[:\s]*.*?Silver[:\s]*(\d+\.\d+)',
                r'Silver[:\s]*(\d+\.\d+)(?:\s*GR)?',
                r'CAST.*?Silver[:\s]*(\d+\.\d+)'
            ]
            
            for pattern in silver_patterns:
                match = re.search(pattern, item_text, re.IGNORECASE)
                if match:
                    item["CAST Fin WT"]["Silver"] = match.group(1)
                    item["Fin Weight (Silver)"] = match.group(1)
                    break
        
        # LOSS % - FIXED to extract percentages correctly
        loss_patterns = [
            r'LOSS %[:\s]*Gold[:\s]*(\d+\.\d+)%?(?:\s*Silver[:\s]*(\d+\.\d+)%?)?',
            r'Loss[:\s]*Gold[:\s]*(\d+\.\d+)%?(?:\s*Silver[:\s]*(\d+\.\d+)%?)?'
        ]
        
        for pattern in loss_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                if match.group(1):
                    item["LOSS %"]["Gold"] = f"{match.group(1)}%"
                if match.group(2):
                    item["LOSS %"]["Silver"] = f"{match.group(2)}%"
                break
        
        # If Silver LOSS % not found with Gold, look for it separately
        if "Silver" not in item["LOSS %"]:
            silver_loss_patterns = [
                r'Silver[:\s]*(\d+\.\d+)%',
                r'LOSS %.*?Silver[:\s]*(\d+\.\d+)%',
                r'Silver:\s*(\d+)%',
                r'Silver\s+(\d+)%'
            ]
            for pattern in silver_loss_patterns:
                match = re.search(pattern, item_text, re.IGNORECASE)
                if match:
                    item["LOSS %"]["Silver"] = f"{match.group(1)}%"
                    break
            
            if "Silver" not in item["LOSS %"]:
                combined_loss_pattern = r'Gold:\s*(\d+\.\d+)%?\s*Silver:\s*(\d+)%?'
                match = re.search(combined_loss_pattern, item_text, re.IGNORECASE)
                if match:
                    item["LOSS %"]["Gold"] = f"{match.group(1)}%"
                    item["LOSS %"]["Silver"] = f"{match.group(2)}%"
        
        # Diamond Details
        diamond_patterns = [
            r'Diamond TW[:\s]*(\d+\.\d+)',
            r'TW[:\s]*(\d+\.\d+)',
            r'LAB DIA.*?TW[:\s]*(\d+/\d+|\d+\.\d+)',
            r'Diamond.*?(\d+\.\d+)\s*CT',
            r'(\d+\.\d+)\s*CT.*?Diamond'
        ]
        
        for pattern in diamond_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                item["Diamond Details"] = f"Diamond TW: {match.group(1)}"
                break

    def extract_item_physical_data(self, item, item_text):
        """Extract physical characteristics from item text"""

        # Size - ENHANCED with description extraction
        if "Size" not in item:  # Only if not already extracted from description
            size_patterns = [
                r'SIZE[:\s]+(\d+(?:\.\d+)?)',
                r'Size[:\s]+(\d+(?:\.\d+)?)',
                                r'(\d+\.\d+)\s+(?:10K|14K|18K|SS|GOS)',
                r'Size[:\s]*(\d+)'
            ]

            for pattern in size_patterns:
                match = re.search(pattern, item_text, re.IGNORECASE)
                if match:
                    item["Size"] = f"SIZE {match.group(1)}"
                    break

        # Metal Category - ENHANCED to properly split metals from description
        if "Metal 1" not in item:  # Only if not already extracted from description
            metal_description = item.get("Metal Description", "")

            # ENHANCED: Look for SS/ patterns specifically and other combinations
            # Check for SS/ first (SS will always be present in SS/... patterns)
            combined_metal_pattern = r'\b(SS|10K[YMWR]|14K[YMWR]|18K[YMWR]|GOLD|SILVER)/(SS|10K[YMWR]|14K[YMWR]|18K[YMWR]|GOLD|SILVER)\b'
            combined_match = re.search(combined_metal_pattern, metal_description.upper())

            if combined_match:
                item["Metal 1"] = combined_match.group(1)
                item["Metal 2"] = combined_match.group(2)
            else:
                single_metal_patterns = [
                    r'\b(10KM|14KM|18KM|10KY|10KW|14KY|14KW|18KY|18KW|SS|SILVER|GOLD)\b'
                ]

                for pattern in single_metal_patterns:
                    match = re.search(pattern, metal_description.upper())
                    if match:
                        item["Metal 1"] = match.group(1)
                        break

                if "Metal 1" not in item:
                    for pattern in single_metal_patterns:
                        match = re.search(pattern, item_text.upper())
                        if match:
                            item["Metal 1"] = match.group(1)
                            break
        
    def extract_components_enhanced(self, item_lines, global_start_idx, all_lines):
        """Enhanced component extraction"""
        components = []
        
        # Find component table start
        component_start = -1
        component_end = len(item_lines)
        
        for i, line in enumerate(item_lines):
            line_lower = line.lower()
            
            if ("supplied by" in line_lower and 
                ("component" in line_lower or "cost" in line_lower or "weight" in line_lower)):
                component_start = i
                break
            
            if re.search(r'^\s*\|\s*Supplied by\s*\|', line, re.IGNORECASE):
                component_start = i
                break
        
        # If still not found, look for component data patterns directly
        if component_start == -1:
            for i, line in enumerate(item_lines):
                if (re.search(r'\b(HS/[\d\.]+|H\d+/[\d\.]+|LD\d+/[\d\.]+|CHSBOXML|PKG\d+|CS\d+|LDZD|LDSH|[A-Z]+\d+|[A-Z0-9/\.-]{3,20})\b', line) and
                    re.search(r'\d+\.\d+\s*(CT|EA|GR)', line)):
                    component_start = i
                    break
        
        if component_start == -1:
            return components
        
        # Process component rows
        header_skipped = False
        for i in range(component_start, component_end):
            if i >= len(item_lines):
                break
                
            line = item_lines[i].strip()
            
            if not line:
                continue
            
            # Skip header line
            line_lower = line.lower()
            if (("supplied by" in line_lower and "component" in line_lower and 
                "cost" in line_lower and "weight" in line_lower) and not header_skipped):
                header_skipped = True
                continue
            
            # Stop conditions
            stop_conditions = [
                "total", "subtotal", "grand", "summary", "there is a",
                "item no.", "description", "unit cost", "please communicate",
                "page:", "richline group", "purchase order"
            ]
            
            if any(stop in line_lower for stop in stop_conditions):
                break
            
            # Stop if we hit another item number
            if re.search(r'\b[A-Z]{2}\d{4}[A-Z0-9]+\b', line):
                break
            
            # Parse component data
            if (re.search(r'\b(HS/[\d\.]+|H\d+/[\d\.]+|LD\d+/[\d\.]+|CHSBOXML|PKG\d+|CS\d+|LDZD|LDSH|[A-Z0-9/\.-]{3,20})\b', line) or
                re.search(r'\d+\.\d+\s*(CT|EA|GR)', line) or
                re.search(r'\b\d{4,5}\b', line)):
                
                component = self.parse_component_line_enhanced(line)
                if component and component.get("Component"):
                    components.append(component)
        
        return components

    def parse_component_line_enhanced(self, line):
        """Parse a single component line - ENHANCED to clean component names properly"""
        print(f"\n=== COMPONENT DEBUG ===")
        print(f"Raw line: '{line}'")

        test_patterns = [
            (r'\b(LD\d+/[\d\.]+)', "LD pattern 1"),
            (r'\b(LD[A-Z]*\d+/[\d\.]+)', "LD pattern 2"),
            (r'(LD\d+/\.?\d*)', "LD pattern 3"),
            (r'(LD\d+/\.\d+)', "LD pattern 4"),
        ]

        for pattern, name in test_patterns:
            match = re.search(pattern, line)
            if match:
                print(f"✅ {name} MATCHED: '{match.group(1)}'")
            else:
                print(f"❌ {name} NOT MATCHED")
        
        print(f"=== END COMPONENT DEBUG ===\n")

        component = {
            "Component": "",
            "Cost ($)": "",
            "Tot. Weight": "",
            "Supply Policy": ""
        }
        
        line = line.strip()
        if not line or len(line) < 3:
            return None
        
        # Add validation to prevent malformed component names - NEW
        skip_patterns = [
            r'^\s*\|\s*Supplied by\s*\|',
            r'^\s*Component\s+Cost\s+Weight',
            r'^\s*Total\s*:',
            r'^\d+\s+\d+\.\d+\s+CT\s+\d+\.\d+\s+CT.*Send To$'  # Skip malformed lines
        ]
        
        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                return None
            
        print(f"\n=== TABLE PARSING DEBUG ===")
        print(f"Raw line: '{line}'")
        
        # Handle table format with | separators
        if "|" in line:
            columns = [col.strip() for col in line.split("|") if col.strip()]
            print(f"Pipe-separated columns: {columns}")
        else:
            parts = line.split()
            if parts:
                component_name = parts[0]
                if re.match(r'^[A-Z0-9/\.-]+$', component_name) and len(component_name) >= 3:
                    columns = [component_name] + parts[1:]
                else:
                    columns = re.split(r'\s{2,}|\t+', line)
                    columns = [col.strip() for col in columns if col.strip()]
            else:
                columns = []

        
        if columns:
            print(f"First column (should be component): '{columns[0]}'")
            if len(columns) > 1:
                print(f"Second column: '{columns[1]}'")
            if len(columns) > 2:
                print(f"Third column: '{columns[2]}'")
            
        print(f"=== END TABLE DEBUG ===\n")
        
        # Extract component name - ENHANCED to get clean component names
        component_name = ""
        
        # Look for specific component patterns and extract ONLY the component name
        component_patterns = [
            r'\b(HS/\d+[A-Z]*)',
            r'\b(CHSBOXML-\d+)',
            r'\b(PKG\d+)',
            r'\b(CS[A-Z0-9\.\-/]+)',
            r'\b(CS[A-Z0-9\.-]+(?:-[A-Z0-9]+)*)',  # Handles CS32.5NV-CR, CS97-OV-CS, CS86-ONY-1PC
            r'\b(LD[A-Z]*\d+/[\d\.]+)',        # LD24/.14, LDS2/.006, LD25/.0066, etc.
            r'\b(LDZD[/\.\d]*)',
            r'\b(LDSH[/\.\d]*)',
            r'\b(H\d+/[\d\.]+)',               # Added: H6/.003, H6/.0033, etc.
            r'\b([A-Z]+\d*/[\d\.]+[A-Z]*)',
            r'\b([A-Z0-9]+/[A-Z0-9\.]+[A-Z0-9]*)',
            r'\b([A-Z0-9]+-[A-Z0-9]+)',
            r'\b([A-Z]{2,}[0-9]+[A-Z0-9/\.-]*)',
        ]
        
        for pattern in component_patterns:
            match = re.search(pattern, line)
            if match:
                component_name = match.group(1)
                break
        
        # If no specific pattern found, try to extract from first column but clean it
        if not component_name and columns:
            first_col = columns[0]
            
            # Try to find component patterns in first column
            for pattern in component_patterns:
                match = re.search(pattern, first_col)
                if match:
                    component_name = match.group(1)
                    break
            
            # If still no match, clean the first column by removing extra data
            if not component_name:
                # Remove vendor IDs, quantities, costs, and weights
                clean_name = re.sub(r'^\d{4,5}\s+', '', first_col)  # Remove vendor ID at start
                clean_name = re.sub(r'\s+\d+\s+\d+\.\d+\s+(EA|CT|GR).*$', '', clean_name)  # Remove quantity + cost + weight
                clean_name = re.sub(r'\s+\d+\s*$', '', clean_name)  # Remove trailing numbers
                clean_name = re.sub(r'\s+(PHW|PS|PH|MS|PM|PT|SS|PF|SAI-HS)\s*\d*\s*$', '', clean_name)  # Remove setting types
                
                if len(clean_name) > 2 and not re.match(r'^\d+$', clean_name):
                    component_name = clean_name.strip()
        
        # If still no component name, check if this line has vendor ID
        if not component_name:
            vendor_id_match = re.search(r'\b(\d{4,5})\b', line)
            if vendor_id_match:
                remaining_text = line.split(vendor_id_match.group(1), 1)
                if len(remaining_text) > 1:
                    for pattern in component_patterns:
                        match = re.search(pattern, remaining_text[1])
                        if match:
                            component_name = match.group(1)
                            break
        
        # Validate component name before proceeding - NEW
        if not component_name:
            return None
        
        # Check if component name looks valid - NEW
        if re.match(r'^[\d\s\.]+$', component_name):  # Only numbers/spaces/dots
            return None
        
        component["Component"] = component_name
        
        # Extract cost
        cost_patterns = [
            r'(\d+\.\d+)\s*CT\b',
            r'(\d+\.\d+)CT\b',
            r'\$(\d+\.\d+)',
            r'(\d+\.\d+)\s*EA\b'
        ]
        
        for pattern in cost_patterns:
            match = re.search(pattern, line)
            if match:
                cost_value = match.group(1)
                if "CT" in match.group(0):
                    component["Cost ($)"] = f"{cost_value} CT"
                elif "EA" in match.group(0):
                    component["Cost ($)"] = f"{cost_value} EA"
                else:
                    component["Cost ($)"] = cost_value
                break
        
        # Extract weight
        weight_patterns = [
            r'(\d+\.\d+)\s*GR\b',
            r'(\d+\.\d+)GR\b',
        ]
        
        # For CT weights, make sure they're not the same as costs
        ct_weight_matches = re.findall(r'(\d+\.\d+)\s*CT\b', line)
        cost_value = component["Cost ($)"].replace(" CT", "") if "CT" in component["Cost ($)"] else ""
        
        for pattern in weight_patterns:
            match = re.search(pattern, line)
            if match:
                component["Tot. Weight"] = match.group(0)
                break
        
        # If no GR weight found, look for CT weights that are different from cost
        if not component["Tot. Weight"] and ct_weight_matches:
            for ct_value in ct_weight_matches:
                if ct_value != cost_value:  # Different from cost value
                    component["Tot. Weight"] = f"{ct_value} CT"
                    break
        
        # Extract supply policy
        line_lower = line.lower()
        if "send to" in line_lower:
            component["Supply Policy"] = "Send To"
        elif "by vendor" in line_lower:
            component["Supply Policy"] = "By Vendor"
        elif "in house" in line_lower:
            component["Supply Policy"] = "In House"
        elif "ship sep" in line_lower:
            component["Supply Policy"] = "Ship Sep"
        
        return component

    def validate_extracted_data(self, global_data, items_data, text):
        """Enhanced validation with better error detection"""
        if not text or len(text.strip()) < 10:
            return False, "No readable text found in PDF"

        # Check for password protection
        if any(keyword in text.lower() for keyword in ["password", "encrypted", "protected"]):
            return False, "PDF appears to be password-protected"

        # Check for minimum required fields
        required_indicators = [
            bool(global_data.get("PO #")),
            bool(global_data.get("Vendor ID #")),
            len(items_data) > 0,
            len(global_data) >= 3,
            "richline" in text.lower() or "rpo" in text.lower()
        ]

        if sum(required_indicators) >= 3:
            return True, "Valid data extracted"

        if len(text.strip()) < 100:
            return False, "PDF may be corrupted or contain insufficient text content"

        return False, "No extractable data found in PDF - may not be a valid purchase order"

    def analyze_field_coverage(self, global_data, items_data):
        """Comprehensive analysis of field extraction coverage"""
        required_global_fields = [
            "PO #", "Location", "PO Date", "Due Date", "Vendor ID #", 
            "Vendor Name", "Order Type", "Gold Rate", "Platinum Rate", "Silver Rate"
        ]
        
        required_item_fields = [
            "Richline Item #", "Job #", "Vendor Item #", "Pieces/Carats", "Fin Weight (Gold)", 
            "Fin Weight (Silver)", "Size", "Metal Description", "Diamond Details", 
            "Stone Labor", "Labor PC", "Ext. Gross Wt.", "Metal 1", "Metal 2",
            "LOSS %", "Rate", "Total Weight"  # Added new fields
        ]
        
        coverage = {
            "global_fields_found": [],
            "global_fields_missing": [],
            "item_fields_coverage": {},
            "components_found": 0,
            "total_fields_extracted": 0,
            "extraction_percentage": 0
        }
        
        # Analyze global fields
        for field in required_global_fields:
            if field in global_data and global_data[field]:
                coverage["global_fields_found"].append(field)
            else:
                coverage["global_fields_missing"].append(field)
        
        # Analyze item fields
        for field in required_item_fields:
            found_count = sum(1 for item in items_data if field in item and item[field])
            coverage["item_fields_coverage"][field] = {
                "found_in_items": found_count,
                "total_items": len(items_data),
                "percentage": (found_count / len(items_data) * 100) if items_data else 0
            }
        
        # Count components
        coverage["components_found"] = sum(len(item.get("Components", [])) for item in items_data)
        
        # Calculate overall extraction percentage
        total_possible_fields = len(required_global_fields) + (len(required_item_fields) * len(items_data))
        total_extracted_fields = len(coverage["global_fields_found"]) + sum(
            coverage["item_fields_coverage"][field]["found_in_items"] 
            for field in required_item_fields
        )
        
        coverage["total_fields_extracted"] = total_extracted_fields
        coverage["extraction_percentage"] = (total_extracted_fields / total_possible_fields * 100) if total_possible_fields > 0 else 0
        
        return coverage

    def extract(self, pdf_file):
        """Main extraction method with comprehensive processing"""
        debug = {"processing_steps": []}

        try:
            # Convert ALL pages to images
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

            # Extract global data with enhanced patterns
            global_data = self.extract_global_fields(all_lines, all_text)
            debug["processing_steps"].append(f"Extracted {len(global_data)} global fields")

            # Extract all items with enhanced method
            items_data = self.extract_items_from_text(all_lines)
            debug["processing_steps"].append(f"Extracted {len(items_data)} items")

            # Add debugging info for components
            total_components = sum(len(item.get("Components", [])) for item in items_data)
            debug["processing_steps"].append(f"Extracted {total_components} total components")
            
            # Add field coverage analysis
            field_coverage = self.analyze_field_coverage(global_data, items_data)
            debug["field_coverage"] = field_coverage

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
                    "raw_text_sample": all_text[:1000] + "..." if len(all_text) > 1000 else all_text,
                }

            debug["total_fields_extracted"] = len(global_data) + sum(len(item) for item in items_data)

            return {
                "global": global_data,
                "items": items_data,
                "debug": debug
            }

        except Exception as e:
            import traceback
            return {
                "error": "Processing failed",
                "details": f"An unexpected error occurred: {str(e)}",
                "debug": debug,
                "traceback": traceback.format_exc()
            }
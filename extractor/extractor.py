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
        
        # Find table boundaries
        table_started = False
        item_rows = []
        
        for i, line in enumerate(lines):
            # Detect table header
            if "Item No." in line and "Description" in line:
                table_started = True
                continue
            
            if table_started:
                # Look for item number patterns (Richline Item #) - FIXED to handle O vs 0
                item_match = re.search(r'\b([A-Z]{2}\d{4}[A-Z0-9]+)\b', line)
                if item_match:
                    item_number = item_match.group(1)
                    # Fix O to 0 in item numbers
                    item_number = item_number.replace('O', '0')
                    item_rows.append((i, item_number, line))
        
        # Process each item
        for idx, (line_idx, item_number, item_line) in enumerate(item_rows):
            # Determine item boundary
            next_item_idx = item_rows[idx + 1][0] if idx + 1 < len(item_rows) else len(lines)
            
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

    def extract_single_item_enhanced(self, item_number, item_line, item_lines, global_start_idx, all_lines):
        """Extract single item with all required fields"""
        item = {"Components": [], "CAST Fin WT": {}, "LOSS %": {}}
        
        # Set Richline Item #
        item["Richline Item #"] = item_number
        
        # Get full item text for pattern matching
        item_text = "\n".join(item_lines)
        
        # Parse the main item line for basic data
        self.parse_item_table_row(item, item_line)
        
        # Extract Job # - ENHANCED to find job numbers for each item more accurately
        job_patterns = [r"(RFP\d{6,})", r"(RSET\d{6,})"]
        
        # Look for job number in the item's context
        for pattern in job_patterns:
            # First, look in the immediate item context
            match = re.search(pattern, item_text)
            if match:
                item["Job #"] = match.group(1)
                break
        
        # If not found in item context, look in a broader window around this item
        if "Job #" not in item:
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

        # Extract Vendor Item # - Look for vendor item codes
        vendor_patterns = [
            r'Vendor Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)',
            r'\b([A-Z]{2}\d{3,6}[A-Z0-9-]*)\b(?!\s*\d+\.\d+)',
            r'Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
        ]
        
        # Look in the item line for vendor item patterns
        vendor_item_line_pattern = r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
        vendor_match = re.search(vendor_item_line_pattern, item_line)
        if vendor_match:
            vendor_item = vendor_match.group(1)
            if vendor_item != item_number:
                item["Vendor Item #"] = vendor_item
        
        # If not found in item line, look in item text
        if "Vendor Item #" not in item:
            for pattern in vendor_patterns:
                matches = re.findall(pattern, item_text)
                for match in matches:
                    if match != item_number:
                        item["Vendor Item #"] = match
                        break
                if "Vendor Item #" in item:
                    break

                # Extract detailed fields
        self.extract_item_financial_data(item, item_text)
        self.extract_item_technical_data(item, item_text)
        self.extract_item_physical_data(item, item_text)
        
        # Extract components
        item["Components"] = self.extract_components_enhanced(item_lines, global_start_idx, all_lines)
        
        return item

    def parse_item_table_row(self, item, item_line):
        """Parse the main table row for basic item data"""
        
        # Extract description (between item number and first price)
        desc_pattern = r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([^0-9]+?)(?=\s+\d+\.\d+)'
        desc_match = re.search(desc_pattern, item_line)
        if desc_match:
            description = desc_match.group(1).strip()
            description = re.sub(r'\s+', ' ', description)
            item["Metal Description"] = description

        # Extract Pieces/Carats and Ext. Gross Wt. based on EA position
        ea_pattern = r'(\d+)\s+(EA|PR)\s+(\d+\.\d+)'
        ea_match = re.search(ea_pattern, item_line)
        if ea_match:
            item["Pieces/Carats"] = ea_match.group(1)
            item["Ext. Gross Wt."] = f"{ea_match.group(3)} GR"

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
        
        # Size
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
        
        # REMOVED Total Weight extraction as requested
        
        # Metal Category - FIXED to properly split metals from description
        metal_description = item.get("Metal Description", "")
        
        # ENHANCED: Look for SS/ patterns specifically and other combinations
        # Check for SS/ first (SS will always be present in SS/... patterns)
        combined_metal_pattern =  r'\b(SS|10K[YMWR]|14K[YMWR]|18K[YMWR]|GOLD|SILVER)/(SS|10K[YMWR]|14K[YMWR]|18K[YMWR]|GOLD|SILVER)\b'
        combined_match  = re.search(combined_metal_pattern, metal_description.upper())
        
        if combined_match:
            item["Metal 1"] = combined_match.group(1)
            item["Metal 2"] = combined_match.group(2)
        else:
            single_metal_patterns = [
                r'\b(10KM|14KM|18KM|10KY|10KW|14KY|14KW|18KY|18KW|SS|SILVER|GOLD)\b']
            
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
                if (re.search(r'\b(HS/\d+|CHSBOXML|PKG\d+|CS\d+|LD\d+|LDZD|LDSH)\b', line) and
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
            if (re.search(r'\b(HS/\d+|CHSBOXML|PKG\d+|CS\d+|LD\d+|LDZD|LDSH|[A-Z]+\d+)\b', line) or
                re.search(r'\d+\.\d+\s*(CT|EA|GR)', line) or
                re.search(r'\b\d{4,5}\b', line)):
                
                component = self.parse_component_line_enhanced(line)
                if component and component.get("Component"):
                    components.append(component)
        
        return components

    def parse_component_line_enhanced(self, line):
        """Parse a single component line - FIXED to clean component names properly"""
        component = {
            "Component": "",
            "Cost ($)": "",
            "Tot. Weight": "",
            "Supply Policy": ""
        }
        
        line = line.strip()
        if not line or len(line) < 3:
            return None
        
        # Handle table format with | separators
        if "|" in line:
            columns = [col.strip() for col in line.split("|") if col.strip()]
        else:
            columns = re.split(r'\s{2,}|\t+', line)
            columns = [col.strip() for col in columns if col.strip()]
        
        if not columns:
            return None
        
        # Extract component name - FIXED to get clean component names
        component_name = ""
        
        # Look for specific component patterns and extract ONLY the component name
        component_patterns = [
            r'\b(HS/\d+[A-Z]*)',
            r'\b(CHSBOXML-\d+)',
            r'\b(PKG\d+)',
            r'\b(CS\d+[A-Z-]*)',
            r'\b(LD\d+[A-Z/]*)',
            r'\b(LDZD[/\.\d]*)',
            r'\b(LDSH[/\.\d]*)',
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
        
        if not component_name:
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
            "LOSS %"
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
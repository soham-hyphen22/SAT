import re
import pytesseract
import pdf2image
import cv2
import numpy as np
from PIL import Image
import pandas as pd
from datetime import datetime
import os

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
            # PDF to Image Conversion FAILED
            return None

    def preprocess_image(self, image):
        try:
            open_cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
            denoised = cv2.fastNlMeansDenoising(gray, h=10)
            _, thresh = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            return Image.fromarray(thresh)
        except Exception as e:
            # Image preprocessing failed
            return image

    def extract_text(self, image):
        try:
            return pytesseract.image_to_string(image)
        except Exception as e:
            # OCR extraction failed
            return ""

    def extract_global_fields(self, lines, full_text):
        result = {}

        # Extract basic patterns
        for field, pattern in self.global_patterns.items():
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                result[field] = match.group(1).replace(",", "")

        # Extract Due Date
        due_date_patterns = [
            r"Due Date[:\s]*([A-Za-z]+ \d{1,2},?\s+\d{4})",
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

        # Extract Order Type
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
            match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE)
            if match:
                groups = match.groups()
                for group in reversed(groups):
                    if group and group.upper() in order_types:
                        result["Order Type"] = group.upper()
                        break
                if "Order Type" in result:
                    break

        # Extract Metal Rates
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
                vendor_id_match = re.search(r"Vendor ID\s*[:#]?\s*([A-Za-z0-9-]+)", line, re.IGNORECASE)
                if vendor_id_match:
                    result["Vendor ID #"] = vendor_id_match.group(1).strip()
                elif i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if next_line and len(next_line) < 30 and re.match(r'^[A-Za-z0-9-]+$', next_line):
                        result["Vendor ID #"] = next_line

                vendor_name_parts = []
                skip_keywords = [
                    "ID", "DATE", "PO", "ORDER", "LOCATION", "RATE", "TYPE", "#", "SHIP", "BILL",
                    "SUPPLY", "CERT", "SEND", "POLICY", "DUE", "TO:", "UNIT"
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
        items = []
        item_patterns = [
            r'\*\*([A-Z]{2}\d{4}[A-Z0-9]+)\*\*',
            r'\b([A-Z]{2}\d{4}[A-Z0-9]+)\b(?=\s+[A-Z]{2}\d{3,6})',
            r'^\s*([A-Z]{2}\d{4}[A-Z0-9]+)\s+',
            r'^\s*([0-9]{5,}[A-Z]{2}[A-Z0-9]*)\s+',
        ]
        
        item_locations = []
        for i, line in enumerate(lines):
            for pattern in item_patterns:
                matches = re.finditer(pattern, line)
                for match in matches:
                    item_number = match.group(1).replace('O', '0').replace('B', '8')
                    item_locations.append((i, item_number, line))
        
        seen_items = set()
        unique_items = []
        for loc in item_locations:
            if loc[1] not in seen_items:
                seen_items.add(loc[1])
                unique_items.append(loc)
        
        unique_items.sort(key=lambda x: x[0])
        
        for idx, (line_idx, item_number, item_line) in enumerate(unique_items):
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
        if not item_line or not item_text:
            return None

        try:
            item_vendor_pattern = r'Item \d+:\s*([A-Z0-9]+)\s+Vendor Style:\s*([A-Z0-9]+)'
            match = re.search(item_vendor_pattern, item_text)
            if match:
                return match.group(1)
            
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
            
            vendor_pattern1 = r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([A-Z]{2}\d{3,6}[A-Z0-9]*)'
            match = re.search(vendor_pattern1, item_line)
            if match:
                return match.group(1)   
            
            vendor_patterns = [
                r'Vendor Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)',
                r'\b([A-Z]{2}\d{3,6}[A-Z0-9]*)\b(?=\s+\d+\.\d+\s+[A-Z]+)',
                r'Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
            ]

            for i, pattern in enumerate(vendor_patterns):
                match = re.search(pattern, item_text)
                if match:
                    return match.group(1)

            if richline_match:
                richline_item = richline_match.group(1)
                lines = item_text.split('\n')

                candidates = []
                for line in lines:
                    if not line:
                        continue
                    line = line.strip()

                    if (re.match(r'^[A-Z0-9-]+$', line) and 5 <= len(line) <= 15 and line != richline_item): 
                        candidates.append(line)

                if candidates:
                    return candidates[0]
            
            return None
        
        except Exception as e:
            return None                       

    def extract_size_from_description(self, description):
        if not description:
            return None
            
        size_patterns = [
            r'^(\d+\.\d+)\s+',
            r'^(\d+)\s+',
            r'(\d+\.\d+)(?:\s+[A-Z]+)',
        ]
        
        for pattern in size_patterns:
            match = re.search(pattern, description)
            if match:
                return f"SIZE {match.group(1)}"
        
        return None

    def extract_metal_from_description(self, description):
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
        item = {"Components": [], "CAST Fin WT": {}, "LOSS %": {},  "Richline Item #": item_number}
        
        item_text = "\n".join(item_lines)
        
        self.parse_item_table_row(item, item_line)
        
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

        job_match = re.search(r'\b(RSET\d{6}|RFP\d{6})\b', item_text)
        if job_match:
            item["Job #"] = job_match.group(1)

        size = self.extract_size_from_description(item.get("Metal Description", ""))
        if size:
            item["Size"] = size
        
        metal1, metal2 = self.extract_metal_from_description(item.get("Metal Description", ""))
        if metal1:
            item["Metal 1"] = metal1
        if metal2:
            item["Metal 2"] = metal2
        
        job_patterns = [r"(RFP\s*\d{6,})", r"(RSET\s*\d{6,})"]
        
        for pattern in job_patterns:
            match = re.search(pattern, item_text)
            if match:
                job_number = match.group(1).replace(" ", "")
                item["Job #"] = job_number
                break

        if "Job #" not in item:
            for pattern in job_patterns:
                start_idx = max(0, global_start_idx - 15)
                end_idx = min(len(all_lines), global_start_idx + 40)
                extended_text = "\n".join(all_lines[start_idx:end_idx])
                
                matches = []
                for match in re.finditer(pattern, extended_text):
                    match_line = extended_text[:match.start()].count('\n')
                    distance = abs(match_line - (global_start_idx - start_idx))
                    matches.append((match.group(1), distance))
                
                if matches:
                    matches.sort(key=lambda x: x[1])
                    item["Job #"] = matches[0][0]
                    break

        if "Vendor Item #" not in item:
            vendor_item_line_pattern = r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
            vendor_match = re.search(vendor_item_line_pattern, item_line)
            if vendor_match:
                vendor_item = vendor_match.group(1)
                if vendor_item != item_number:
                    item["Vendor Item #"] = vendor_item
            
            if "Vendor Item #" not in item:
                vendor_patterns = [
                    r'Vendor Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)',
                    r'\b([A-Z]{2}\d{3,6}[A-Z0-9-]*)\b(?!\s*\d+\.\d+)',
                    r'Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
                ]
                for i, pattern in enumerate(vendor_patterns):
                    match = re.search(pattern, item_text)
                    if match:
                        item["Vendor Item #"] = match.group(1)
                        break
        
        self.extract_item_financial_data(item, item_text)
        self.extract_item_technical_data(item, item_text)
        self.extract_item_physical_data(item, item_text)
        item["Components"] = self.extract_components_enhanced(item_lines, global_start_idx, all_lines)
        
        return item

    def parse_item_table_row(self, item, item_line):
        desc_patterns = [
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([0-9.]+\s+[A-Z/]+\s+[^|]+?)(?=\s+[\d.]+\s+\d+\s+EA)',
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+(.+?)(?=\s+[\d.]+\s+\d+\s+(?:EA|PR))',
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+(.+?)(?=\s+[\d.]+)',
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([^0-9]+?)(?=\s+\d+\.\d+)'
        ]
        
        for pattern in desc_patterns:
            desc_match = re.search(pattern, item_line)
            if desc_match:
                description = desc_match.group(1).strip()
                description = re.sub(r'\s+', ' ', description)
                description = re.sub(r'\s*\|\s*', ' ', description)
                if len(description) > 5:
                    item["Metal Description"] = description
                    break

        table_patterns = [
            r'[A-Z]{2}\d{4}[A-Z0-9]+\s+.+?\s+([\d.]+)\s+(\d+)\s+(EA|PR)\s+([\d.]+)',
            r'(\d+)\s+(EA|PR)\s+(\d+\.\d+)',
        ]
        
        for pattern in table_patterns:
            table_match = re.search(pattern, item_line)
            if table_match:
                if len(table_match.groups()) == 4:
                    item["Pieces/Carats"] = table_match.group(2)
                    item["Ext. Gross Wt."] = f"{table_match.group(4)} GR"
                elif len(table_match.groups()) == 3:
                    item["Pieces/Carats"] = table_match.group(1)
                    item["Ext. Gross Wt."] = f"{table_match.group(3)} GR"
                break

    def extract_item_financial_data(self, item, item_text):
        stone_patterns = [r'Stone PC[:\s]+(\d+\.\d+)', r'Stone[:\s]+(\d+\.\d+)', r'Stone Labor[:\s]+(\d+\.\d+)']
        for pattern in stone_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                item["Stone Labor"] = match.group(1)
                break
        
        labor_patterns = [r'Labor PC[:\s]+(\d+\.\d+)', r'Labor[:\s]+(\d+\.\d+)']
        for pattern in labor_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                item["Labor PC"] = match.group(1)
                break

    def extract_item_technical_data(self, item, item_text):
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
        
        if "Silver" not in item["CAST Fin WT"]:
            silver_patterns = [
                r'CAST Fin WT[:\s]*.*?Silver[:\s]*(\d+\.\d+)', r'Fin WT[:\s]*.*?Silver[:\s]*(\d+\.\d+)',
                r'Silver[:\s]*(\d+\.\d+)(?:\s*GR)?', r'CAST.*?Silver[:\s]*(\d+\.\d+)'
            ]
            for pattern in silver_patterns:
                match = re.search(pattern, item_text, re.IGNORECASE)
                if match:
                    item["CAST Fin WT"]["Silver"] = match.group(1)
                    item["Fin Weight (Silver)"] = match.group(1)
                    break
        
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
        
        if "Silver" not in item["LOSS %"]:
            silver_loss_patterns = [r'Silver[:\s]*(\d+\.\d+)%', r'LOSS %.*?Silver[:\s]*(\d+\.\d+)%', r'Silver:\s*(\d+)%', r'Silver\s+(\d+)%']
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
        
        diamond_patterns = [
            r'Diamond TW[:\s]*(\d+\.\d+)', r'TW[:\s]*(\d+\.\d+)', r'LAB DIA.*?TW[:\s]*(\d+/\d+|\d+\.\d+)',
            r'Diamond.*?(\d+\.\d+)\s*CT', r'(\d+\.\d+)\s*CT.*?Diamond'
        ]
        for pattern in diamond_patterns:
            match = re.search(pattern, item_text, re.IGNORECASE)
            if match:
                item["Diamond Details"] = f"Diamond TW: {match.group(1)}"
                break

    def extract_item_physical_data(self, item, item_text):
        if "Size" not in item:
            size_patterns = [
                r'SIZE[:\s]+(\d+(?:\.\d+)?)', r'Size[:\s]+(\d+(?:\.\d+)?)',
                r'(\d+\.\d+)\s+(?:10K|14K|18K|SS|GOS)', r'Size[:\s]*(\d+)'
            ]
            for pattern in size_patterns:
                match = re.search(pattern, item_text, re.IGNORECASE)
                if match:
                    item["Size"] = f"SIZE {match.group(1)}"
                    break

        if "Metal 1" not in item:
            metal_description = item.get("Metal Description", "")
            combined_metal_pattern = r'\b(SS|10K[YMWR]|14K[YMWR]|18K[YMWR]|GOLD|SILVER)/(SS|10K[YMWR]|14K[YMWR]|18K[YMWR]|GOLD|SILVER)\b'
            combined_match = re.search(combined_metal_pattern, metal_description.upper())

            if combined_match:
                item["Metal 1"] = combined_match.group(1)
                item["Metal 2"] = combined_match.group(2)
            else:
                single_metal_patterns = [r'\b(10KM|14KM|18KM|10KY|10KW|14KY|14KW|18KY|18KW|SS|SILVER|GOLD)\b']
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
        components = []
        component_start = -1
        
        # More robust detection of the component table header
        header_patterns = [
            r'supplied by.*component.*cost',
            r'^\s*\|\s*Supplied by\s*\|',
            r'Component\s+Setting Typ\s+Qty',
            r'Component\s+Cost\s+KATEX_INLINE_OPEN\$KATEX_INLINE_CLOSE\s+Tot'
        ]

        for i, line in enumerate(item_lines):
            for pattern in header_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    component_start = i + 1
                    break
            if component_start != -1:
                break
        
        if component_start == -1:
            return components
        
        for i in range(component_start, len(item_lines)):
            line = item_lines[i].strip()
            if not line:
                continue

            line_lower = line.lower()
            stop_conditions = [
                "total", "subtotal", "grand", "summary", "there is a",
                "please communicate", "page:", "richline group", "purchase order"
            ]
            if any(stop in line_lower for stop in stop_conditions):
                break
            
            if re.search(r'\b[A-Z]{2}\d{4}[A-Z0-9]+\b', line) and i > component_start:
                break
            
            component = self.parse_component_line_enhanced(line)
            if component and component.get("Component"):
                components.append(component)

        return components

    def parse_component_line_enhanced(self, line):
        """Parse component line with correct Tot. Weight extraction"""
        component = {
            "Component": "",
            "Cost ($)": "",
            "Tot. Weight": "",
            "Supply Policy": ""
        }

        line = line.strip()
        if not line or len(line) < 3:
            return None

        # Skip header and invalid lines
        skip_patterns = [
            r'^\s*\|\s*Supplied by\s*\|',
            r'^\s*Component\s+Cost\s+Weight',
            r'^\s*Total\s*:',
            r'^\s*Supplied by\s+Component\s+Setting Typ',
            r'^\s*\|\s*\-+\s*\|'
        ]

        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                return None

        # Parse pipe-separated table format (MAIN FIX HERE)
        if "|" in line:
            columns = [col.strip() for col in line.split("|")]
        
            # Remove empty columns at start/end
            while columns and not columns[0]:
                columns.pop(0)
            while columns and not columns[-1]:
                columns.pop()
            
            if len(columns) >= 7:
                try:
                    component["Component"] = columns[1]
                    component["Cost ($)"] = columns[4]
                    component["Tot. Weight"] = columns[6]
                
                    for col in columns[7:]:
                        if re.search(r'(Send To|Drop Ship|By Vendor)', col, re.IGNORECASE):
                            component["Supply Policy"] = col
                            break
                        
                except IndexError:
                    # Fallback to non-pipe parsing
                    pass
            else:
                 # Fallback to non-pipe parsing if columns are not as expected
                 pass
        
        # If pipe parsing fails or isn't applicable, use regex fallback
        if not component["Component"]:
            component_patterns = [
                # CS patterns - most common
                r'\b(CS\d{1,4}(?:/\d{1,4}(?:\.\d+)?)?[A-Z]{0,3}(?:-[A-Z0-9]+)+)\b',
                r'\b(CS[0-9/\.-]+(?:-[A-Z0-9]+)*)', r'\b(CS[A-Z0-9\./\-]+)',
                r'\b(CS\d+/\d+(?:\.\d+)?(?:NV|OV|PS|HS|RDP)-[A-Z0-9]+)',  # CS5/2.5NV-CR, CS9/7-OV-CS
                r'\b(CS\d+(?:/\d+(?:\.\d+)?)?-[A-Z0-9]+-[A-Z0-9]+)',      # CS0025-R-CR, CS0020-R-E2A
                r'\b(CS\d+(?:/\d+(?:\.\d+)?)?[A-Z]+-[A-Z0-9]+)',
                r'\b(CS\d+[/-]|LD\d+[/.]|OT-CHR\d+-\d+)\b', 
                # Specific component types
                r'\b(CUN\.[A-Z]+\.\d+\.\d+)\b', 
                r'\b(\d{1,4}/\d{1,4}(?:\.\d+)?[A-Z]{0,3}(?:-[A-Z0-9]+)+)\b',
                r'\b(THP-[A-Z0-9\-]+)\b',
                r'\b(THP-WH\d+-[A-Z]+)',
                r'\b(THP-[A-Z0-9\-]+)', 
                r'\b([0-9]{2}XX[0-9]{4}-[A-Z0-9]+)',
                r'\b(SSC\d{6,8}[A-Z0-9]*)\b',
                r'\b(SSC[0-9]+[A-Z0-9]*)', 
                r'\b(PKG[0-9]+)',
                r'\b(PKG\d{4,6})\b',
                r'\b(TRC\d{6,8}[A-Z0-9]*)\b',
                r'\b(TRC[0-9]+[A-Z0-9]*)',
                r'\b(CHR[A-Z0-9]+W?-\d+[A-Z]?)\b', 
                r'\b(CHR[A-Z0-9]+W?-\d+[A-Z]?)',
                r'\b(OT-[A-Z0-9\-]+)\b',
                r'\b(OT-[A-Z0-9]+)',
                r'\b(OT-CHR\d+-\d+)', 
                r'\b(OT-[A-Z]+\d*)', 
                r'\b(R\d{3,5}-\d{2,4}[A-Z0-9\-]*)\b',
                r'\b(R[0-9]{3}-[0-9]+[A-Z\-]*)',
                r'\b(CN\d{4,6}-[A-Z0-9\-]+)\b',
                r'\b(CN\d{4}-[A-Z0-9]+-\d+)',
                r'\b(CN[0-9]{4}-[A-Z0-9\-]+)',
                r'\b(MS\d{4,6}-[A-Z0-9]+)\b',
                r'\b(MS\d{4}-[A-Z0-9]+)', 
                r'\b(MS[0-9]{4}-[A-Z0-9]+)',
                r'\b(A\d{1,2}/\.\d{3,4})\b',
                r'\b(A\d+/\.\d+)',
                r'\b(A[0-9]+/\.[0-9]+)',
                # CHSBOXIML pattern - ADD THIS
                r'\b(CHSBOXIML-\d+)',
                r'\b(CHS[A-Z]+ML-\d+)',
                # LD patterns
                r'\b(LD[A-Z]*[0-9]+[A-Z]*[/\.][0-9\.]+)', r'\b(LD[A-Z]*[0-9]+/[0-9\.]+)',
                r'\b(LD[A-Z0-9\.]+)',
                # Other patterns
                r'\b(H[0-9]+/[0-9\.]+)',
                r'\b([A-Z]{1,4}\d{1,6}[A-Z0-9\./\-]*)',
                r'\b([A-Z]+\d+[A-Z]*(?:[/\.-][A-Z0-9]+)*)', 
                r'\b([A-Z0-9]{2,}[/-][A-Z0-9./-]{2,})\b',
                r'\b([0-9]{4}-[A-Z]+-[A-Z]+)',
                r'\b([0-9]/[0-9\.]+-[A-Z]+-[A-Z]+)', 
                r'\b([0-9]/[0-9\.]+)',
                r'\b([A-Z0-9\-]{5,})\b',
                r'\b([0-9]+/[0-9\.]+[A-Z]*-[A-Z0-9]+)',
            ]
        
            for pattern in component_patterns:
                match = re.search(pattern, line)
                if match:
                    component["Component"] = match.group(1)
                    break
        
            values = re.findall(r'(\d+\.\d+)\s*(CT|EA|GR)', line)
        
            if len(values) >= 3:
                component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"
                component["Tot. Weight"] = f"{values[2][0]} {values[2][1]}"
            elif len(values) >= 2:
                component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"
                component["Tot. Weight"] = f"{values[1][0]} {values[1][1]}"
            elif len(values) == 1:
                component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"

        if not component["Supply Policy"]:
            line_lower = line.lower()
            if "send to" in line_lower:
                component["Supply Policy"] = "Send To"
            elif "drop ship" in line_lower:
                component["Supply Policy"] = "Drop Ship"
            elif "by vendor" in line_lower:
                component["Supply Policy"] = "By Vendor"

        return component if component["Component"] else None

    def validate_extracted_data(self, global_data, items_data, text):
        if not text or len(text.strip()) < 10:
            return False, "No readable text found in PDF"

        if any(keyword in text.lower() for keyword in ["password", "encrypted", "protected"]):
            return False, "PDF appears to be password-protected"

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
        required_global_fields = [
            "PO #", "Location", "PO Date", "Due Date", "Vendor ID #", "Vendor Name",
            "Order Type", "Gold Rate", "Platinum Rate", "Silver Rate"
        ]
        
        required_item_fields = [
            "Richline Item #", "Job #", "Vendor Item #", "Pieces/Carats", "Fin Weight (Gold)", 
            "Fin Weight (Silver)", "Size", "Metal Description", "Diamond Details", "Stone Labor",
            "Labor PC", "Ext. Gross Wt.", "Metal 1", "Metal 2", "LOSS %", "Rate", "Total Weight"
        ]
        
        coverage = {
            "global_fields_found": [], "global_fields_missing": [], "item_fields_coverage": {},
            "components_found": 0, "total_fields_extracted": 0, "extraction_percentage": 0
        }
        
        for field in required_global_fields:
            if field in global_data and global_data[field]:
                coverage["global_fields_found"].append(field)
            else:
                coverage["global_fields_missing"].append(field)
        
        for field in required_item_fields:
            found_count = sum(1 for item in items_data if field in item and item[field])
            coverage["item_fields_coverage"][field] = {
                "found_in_items": found_count, "total_items": len(items_data),
                "percentage": (found_count / len(items_data) * 100) if items_data else 0
            }
        
        coverage["components_found"] = sum(len(item.get("Components", [])) for item in items_data)
        
        total_possible_fields = len(required_global_fields) + (len(required_item_fields) * len(items_data))
        total_extracted_fields = len(coverage["global_fields_found"]) + sum(
            coverage["item_fields_coverage"][field]["found_in_items"] for field in required_item_fields
        )
        
        coverage["total_fields_extracted"] = total_extracted_fields
        coverage["extraction_percentage"] = (total_extracted_fields / total_possible_fields * 100) if total_possible_fields > 0 else 0
        
        return coverage

    def extract(self, pdf_file):
        debug = {"processing_steps": []}
        try:
            images = self.convert_pdf_to_image(pdf_file)
            debug["processing_steps"].append("PDF converted to images")

            if not images:
                return {
                    "error": "Failed to convert PDF to images",
                    "details": "The PDF may be corrupted, password-protected, or contain non-text content",
                    "debug": debug,
                }

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

            global_data = self.extract_global_fields(all_lines, all_text)
            debug["processing_steps"].append(f"Extracted {len(global_data)} global fields")

            items_data = self.extract_items_from_text(all_lines)
            debug["processing_steps"].append(f"Extracted {len(items_data)} items")

            total_components = sum(len(item.get("Components", [])) for item in items_data)
            debug["processing_steps"].append(f"Extracted {total_components} total components")
            
            field_coverage = self.analyze_field_coverage(global_data, items_data)
            debug["field_coverage"] = field_coverage

            is_valid, validation_message = self.validate_extracted_data(global_data, items_data, all_text)
            debug["validation_message"] = validation_message

            if not is_valid:
                return {
                    "error": "No extractable data found in PDF",
                    "details": validation_message, "debug": debug,
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
        
def export_to_excel(extracted_data, base_filename="Daily_PO_Extracts"):
    """Export extracted PDF data to daily Excel file"""
    
    # Create daily filename
    today = datetime.now().strftime("%Y-%m-%d")
    filename = f"{base_filename}_{today}.xlsx"
    
    # Check if file exists
    file_exists = os.path.exists(filename)
    
    if file_exists:
        # Read existing data
        try:
            existing_global = pd.read_excel(filename, sheet_name='Global_Data')
            existing_items = pd.read_excel(filename, sheet_name='Items_Summary')
            existing_components = pd.read_excel(filename, sheet_name='Components_Detail')
        except:
            # If sheets don't exist, create empty dataframes
            existing_global = pd.DataFrame()
            existing_items = pd.DataFrame()
            existing_components = pd.DataFrame()
    else:
        # Create empty dataframes for new file
        existing_global = pd.DataFrame()
        existing_items = pd.DataFrame()
        existing_components = pd.DataFrame()
    
    # Prepare new data
    new_global = pd.DataFrame()
    new_items = pd.DataFrame()
    new_components = pd.DataFrame()
    
    # Add timestamp and PO number for tracking
    timestamp = datetime.now().strftime("%H:%M:%S")
    po_number = extracted_data.get('global', {}).get('PO #', 'Unknown')
    
    # 1. Process Global Data
    if 'global' in extracted_data:
        global_data = extracted_data['global'].copy()
        global_data['Extraction_Time'] = timestamp
        global_data['Extraction_Date'] = today
        new_global = pd.DataFrame([global_data])
    
    # 2. Process Items Data
    if 'items' in extracted_data:
        items_list = []
        for item_idx, item in enumerate(extracted_data['items']):
            item_row = {
                'PO_Number': po_number,
                'Extraction_Time': timestamp,
                'Extraction_Date': today,
                'Item_Index': item_idx + 1
            }
            
            # Basic item info
            for key, value in item.items():
                if key != 'Components' and not isinstance(value, dict):
                    item_row[key] = value
                elif isinstance(value, dict):
                    # Flatten nested dictionaries
                    for sub_key, sub_value in value.items():
                        item_row[f"{key}_{sub_key}"] = sub_value
            
            # Component count
            item_row['Component_Count'] = len(item.get('Components', []))
            items_list.append(item_row)
        
        if items_list:
            new_items = pd.DataFrame(items_list)
    
    # 3. Process Components Data
    if 'items' in extracted_data:
        components_list = []
        for item_idx, item in enumerate(extracted_data['items']):
            richline_item = item.get('Richline Item #', f'Item_{item_idx+1}')
            vendor_item = item.get('Vendor Item #', '')
            job_number = item.get('Job #', '')
            
            for comp in item.get('Components', []):
                comp_row = {
                    'PO_Number': po_number,
                    'Extraction_Time': timestamp,
                    'Extraction_Date': today,
                    'Item_Index': item_idx + 1,
                    'Richline_Item': richline_item,
                    'Vendor_Item': vendor_item,
                    'Job_Number': job_number,
                    'Component': comp.get('Component', ''),
                    'Cost': comp.get('Cost ($)', ''),
                    'Tot_Weight': comp.get('Tot. Weight', ''),
                    'Supply_Policy': comp.get('Supply Policy', '')
                }
                components_list.append(comp_row)
        
        if components_list:
            new_components = pd.DataFrame(components_list)
    
    # Combine existing and new data
    if not existing_global.empty and not new_global.empty:
        combined_global = pd.concat([existing_global, new_global], ignore_index=True)
    elif not new_global.empty:
        combined_global = new_global
    else:
        combined_global = existing_global
    
    if not existing_items.empty and not new_items.empty:
        combined_items = pd.concat([existing_items, new_items], ignore_index=True)
    elif not new_items.empty:
        combined_items = new_items
    else:
        combined_items = existing_items
    
    if not existing_components.empty and not new_components.empty:
        combined_components = pd.concat([existing_components, new_components], ignore_index=True)
    elif not new_components.empty:
        combined_components = new_components
    else:
        combined_components = existing_components
    
    # Write to Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        if not combined_global.empty:
            combined_global.to_excel(writer, sheet_name='Global_Data', index=False)
        if not combined_items.empty:
            combined_items.to_excel(writer, sheet_name='Items_Summary', index=False)
        if not combined_components.empty:
            combined_components.to_excel(writer, sheet_name='Components_Detail', index=False)
    
    return filename, len(existing_global) > 0  
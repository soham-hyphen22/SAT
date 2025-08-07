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

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Samuel Aaron\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

class FastPDFOCRExtractor:
    def __init__(self, output_folder=None):
        self.global_patterns = {
            "PO #": r"(RPO\d+)",
            "PO Date": r"\b(\d{2}/\d{2}/\d{2,4})\b",
            "Location": r"Location[:\s]*([A-Z]{2,})"
        }
        self.job_pattern = r"(RFP\d{6,}|RSET\d{6,})"

    def convert_pdf_to_image(self, pdf_file):
        """OPTIMIZED: Faster PDF conversion"""
        poppler_path = r"C:\Users\Samuel Aaron\Documents\Release-24.08.0-0\poppler-24.08.0\Library\bin"
        try:
            pdf_file.seek(0)        
            images = pdf2image.convert_from_bytes(
                pdf_file.read(),
                dpi=200,
                poppler_path=poppler_path,
                thread_count=4,
                fmt='jpeg'
            )
            return images if images else None
        except Exception as e:
            print(f"PDF to Image Conversion FAILED: {e}")
            return None

    def preprocess_image(self, image):
        """OPTIMIZED: Faster preprocessing"""
        try:
            open_cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            return Image.fromarray(thresh)
        except Exception as e:
            print(f"Image preprocessing failed: {e}")
            return image

    def extract_text(self, image):
        """OPTIMIZED: Faster OCR"""
        try:
            return pytesseract.image_to_string(image, config='--oem 3 --psm 6')
        except Exception as e:
            print(f"OCR extraction failed: {e}")
            return ""

    def process_page_parallel(self, page_data):
        """Process single page for parallel execution"""
        page_num, image = page_data
        try:
            preprocessed_image = self.preprocess_image(image)
            page_text = self.extract_text(preprocessed_image)
            return page_num, page_text
        except Exception as e:
            print(f"Error processing page {page_num}: {e}")
            return page_num, ""

    def extract_global_fields(self, lines, full_text):
        """Extract global fields - SAME AS YOUR WORKING CODE"""
        result = {}

        first_page_fields = ["Location", "PO #", "PO Date"]
        pages = full_text.split('#page')
    
        for field, pattern in self.global_patterns.items():
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

    def find_items_with_rpo_association(self, lines):
        """FIXED: Find items and associate them with their RPO based on proximity"""
        
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
            # Find the closest RPO that appears before this item
            closest_rpo = None
            closest_distance = float('inf')
            
            for rpo_line, rpo_number in rpo_positions:
                if rpo_line <= item_line:  # RPO must come before or at the item
                    distance = item_line - rpo_line
                    if distance < closest_distance:
                        closest_distance = distance
                        closest_rpo = rpo_number
            
            # If no RPO found before item, use the first RPO in document
            if closest_rpo is None and rpo_positions:
                closest_rpo = rpo_positions[0][1]
            
            # Group items by RPO
            if closest_rpo:
                if closest_rpo not in items_by_rpo:
                    items_by_rpo[closest_rpo] = []
                items_by_rpo[closest_rpo].append((item_line, item_number, item_text))
        
        return items_by_rpo

    def extract_single_item_enhanced(self, item_number, item_line, item_lines, global_start_idx, all_lines):
        """Extract single item - SAME AS YOUR WORKING CODE"""
        item = {"Components": [], "CAST Fin WT": {}, "LOSS %": {},  "Richline Item #": item_number}
        
        item_text = "\n".join(item_lines)
        
        # Parse item table row
        self.parse_item_table_row(item, item_line)
        
                # Extract vendor item
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

        # Extract size
        size_match = re.search(r'\b(\d{1,2}\.\d{2})\b', item_text)
        if size_match:
            item["Size"] = f"SIZE {size_match.group(1)}"

        # Extract job number
        job_patterns = [r"(RFP\s*\d{6,})", r"(RSET\s*\d{6,})"]
        
        for pattern in job_patterns:
            match = re.search(pattern, item_text)
            if match:
                job_number = match.group(1).replace(" ", "")
                item["Job #"] = job_number
                break

        # Extract metals
        metal1, metal2 = self.extract_metal_from_description(item.get("Metal Description", ""))
        if metal1:
            item["Metal 1"] = metal1
        if metal2:
            item["Metal 2"] = metal2
        
        # Extract financial, technical, and physical data
        self.extract_item_financial_data(item, item_text)
        self.extract_item_technical_data(item, item_text)
        self.extract_item_physical_data(item, item_text)
        
        # Extract components
        item["Components"] = self.extract_components_enhanced(item_lines, global_start_idx, all_lines)
        
        return item

    def extract_vendor_item_from_ocr(self, item_line, item_text):
        """Extract vendor item - SAME AS YOUR WORKING CODE"""
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
            
            # Pattern 3: Direct vendor patterns
            vendor_pattern1 = r'[A-Z]{2}\d{4}[A-Z0-9]+\s+([A-Z]{2}\d{3,6}[A-Z0-9]*)'
            match = re.search(vendor_pattern1, item_line)
            if match:
                return match.group(1)   
            
            # Pattern 4: Style patterns
            vendor_patterns = [
                r'Vendor Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)',
                r'\b([A-Z]{2}\d{3,6}[A-Z0-9]*)\b(?=\s+\d+\.\d+\s+[A-Z]+)',
                r'Style[:\s]*([A-Z]{2}\d{3,6}[A-Z0-9-]*)'
            ]

            for pattern in vendor_patterns:
                match = re.search(pattern, item_text)
                if match:
                    return match.group(1)
            
            # Pattern 5: Universal pattern
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

    def parse_item_table_row(self, item, item_line):
        """Parse item table row - SAME AS YOUR WORKING CODE"""
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

    def extract_metal_from_description(self, description):
        """Extract metals - SAME AS YOUR WORKING CODE"""
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

        # Check for bimetal patterns
        bimetal_pattern = r'\b([A-Z0-9]+)\s*/\s*([A-Z0-9]+)\b'
        bimetal_matches = re.findall(bimetal_pattern, description_upper)
        
        for metal1_candidate, metal2_candidate in bimetal_matches:
            if metal1_candidate in valid_metals and metal2_candidate in valid_metals:
                return metal1_candidate, metal2_candidate
            
        # Find single metals
        found_metals = []
        for metal in valid_metals:
            if re.search(r'\b' + re.escape(metal) + r'\b', description_upper):
                if metal not in found_metals:
                    found_metals.append(metal)

        if found_metals:
            return found_metals[0], None
        
        return None, None

    def extract_item_financial_data(self, item, item_text):
        """Extract financial data - SAME AS YOUR WORKING CODE"""
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
        """Extract technical data - SAME AS YOUR WORKING CODE"""
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
                    item["CAST Fin WT"]["Gold"] = match.group(1)
                    item["Fin Weight (Gold)"] = match.group(1)
                if match.group(2):
                    item["CAST Fin WT"]["Silver"] = match.group(2)
                    item["Fin Weight (Silver)"] = match.group(2)
                break
        
        # Silver patterns if not found above
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
        
        # Loss percentage patterns
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
        
        # Silver loss patterns if not found above
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
        
        # Diamond details
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
        """Extract physical data - SAME AS YOUR WORKING CODE"""
        if "Size" not in item:
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

    def extract_components_enhanced(self, item_lines, global_start_idx, all_lines):
        """Extract components - SAME AS YOUR WORKING CODE"""
        components = []
        component_start = -1
        
        # Find component table header
        header_patterns = [
            r'supplied by.*component.*cost',
            r'^\s*\|\s*Supplied by\s*\|',
            r'Component\s+Setting Typ\s+Qty',
            r'Component\s+Cost\s+\$\s+Tot'
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
                "please communicate", "page:", "richline group", "purchase order",
                "cast fin wt", "loss %"
            ]
            if any(stop in line_lower for stop in stop_conditions):
                break
            
            # Stop if we hit another item
            if re.search(r'\b[A-Z]{2}\d{4}[A-Z0-9]+\b', line) and i > component_start:
                break
            if re.search(r'\b(RSET\d{6}|RFP\d{6})\b', line): # Skip job number lines
                continue
            if re.match(r'^[A-Z]{2}\d{4,}', line.strip()): # Skip Richline Item # lines
                continue
            
            component = self.parse_component_line_enhanced(line)
            if component and component.get("Component"):
                components.append(component)

        return components

    def parse_component_line_enhanced(self, line):
        """Parse component line - SAME AS YOUR WORKING CODE"""
        component = {
            "Component": "",
            "Cost ($)": "",
            "Tot. Weight": "",
            "Supply Policy": ""
        }   

        line = line.strip()
        if not line or len(line) < 5:
            return None

        # Skip header lines
        skip_patterns = [
            r'Supplied By', r'^\s*Total\s*:', r'^\s*\|\s*\-+\s*\|',
            r'Supplied By\s+Component\s+Setting',
            r'^\s*\|\s*Supplied by\s*\|',
            r'^\s*Component\s+Setting\s+Typ',
            r'\b(RSET\d{6}|RFP\d{6})\b',
            r'^\s*Component\s+Cost\s+Weight',
            r'\b(CAS\d+|DA\d+|CA\d+)\b',
            r'^\s*Total\s*:',
            r'^\s*\|\s*\-+\s*\|'
        ]

        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                return None
            
        parts = re.split(r'\s{2,}', line)
        if len(parts) == 1: parts = line.split()

        component_candidate = ""

        for part in parts:
            if len(part) > 3 and re.search(r'[A-Za-z]', part) and re.search(r'[\d/\.-]', part):
                if not re.match(r'^\d+\.\d+$', part):
                    component_candidate = part
                    break

        if not component_candidate:
            return None
        component["Component"] = component_candidate.strip()
        
        # Handle pipe-delimited format
        if "|" in line:
            columns = [col.strip() for col in line.split("|")]
            
            # Remove empty columns
            while columns and not columns[0]:
                columns.pop(0)
            while columns and not columns[-1]:
                columns.pop()

            if len(columns) >= 7:
                try:
                    component["Component"] = columns[1]
                    component["Cost ($)"] = columns[4]
                    component["Tot. Weight"] = columns[6]

                    # Look for supply policy in remaining columns
                    for col in columns[7:]:
                        if re.search(r'(Send To|Drop Ship|By Vendor|In House)', col, re.IGNORECASE):
                            component["Supply Policy"] = col
                            break
                except IndexError:
                    pass

        # Handle space-separated format
        if not component["Component"]:
            # Extract component name using enhanced patterns
            component_patterns = [
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
                ]
            
            for pattern in component_patterns:
                match = re.search(pattern, line)
                if match:
                    component["Component"] = match.group(1)
                    break
            
            # Extract values with units - CORRECTED LOGIC
            values = re.findall(r'(\d+\.\d+)\s*(CT|EA|GR)', line)
            
            if len(values) >= 3:
                # First value = Cost, Third value = Weight (based on your working code)
                component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"
                component["Tot. Weight"] = f"{values[2][0]} {values[2][1]}"
            elif len(values) >= 2:
                # Apply business logic: higher value typically = cost
                val1 = float(values[0][0])
                val2 = float(values[1][0])
                
                if val1 > val2:
                    component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"
                    component["Tot. Weight"] = f"{values[1][0]} {values[1][1]}"
                else:
                    component["Cost ($)"] = f"{values[1][0]} {values[1][1]}"
                    component["Tot. Weight"] = f"{values[0][0]} {values[0][1]}"
            elif len(values) == 1:
                # Single value - determine if cost or weight based on magnitude
                value = float(values[0][0])
                if value > 25:  # Likely cost
                    component["Cost ($)"] = f"{values[0][0]} {values[0][1]}"
                else:  # Likely weight
                    component["Tot. Weight"] = f"{values[0][0]} {values[0][1]}"

            # Extract supply policy
            policy_keywords = ["Send To", "Drop Ship", "By Vendor", "In House"]
            for keyword in policy_keywords:
                if keyword.lower() in line.lower():
                    component["Supply Policy"] = keyword
                    break

        return component if component["Component"] else None

    def extract(self, pdf_file):
        """MAIN METHOD: FIXED approach for multiple RPOs"""
        start_time = datetime.now()
        debug = {"processing_steps": []}
        
        try:
            # Convert PDF to images with parallel processing
            images = self.convert_pdf_to_image(pdf_file)
            if not images:
                return {"error": "Failed to convert PDF to images", "debug": debug}
            
            debug["processing_steps"].append(f"PDF converted to {len(images)} images")

            # Parallel OCR processing
            all_text = ""
            all_lines = []
            
            max_workers = min(4, len(images))
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                page_futures = {
                    executor.submit(self.process_page_parallel, (i, image)): i 
                    for i, image in enumerate(images)
                }
                
                page_results = {}
                for future in concurrent.futures.as_completed(page_futures):
                    page_num, page_text = future.result()
                    page_results[page_num] = page_text
                
                for i in range(len(images)):
                    if i in page_results:
                        all_text += f"\n#page {i+1}\n" + page_results[i]
                        all_lines.extend(page_results[i].splitlines())

            debug["processing_steps"].append("OCR completed with parallel processing")
            debug["total_pages"] = len(images)

            if not all_text:
                return {"error": "No extractable data found in PDF", "debug": debug}

            debug["ocr_text_length"] = len(all_text)

            # STEP 1: Extract global data ONCE (shared across all RPOs)
            global_data = self.extract_global_fields(all_lines, all_text)
            debug["processing_steps"].append(f"Extracted {len(global_data)} global fields")

            # STEP 2: Find all RPOs and associate items with them
            items_by_rpo = self.find_items_with_rpo_association(all_lines)
            debug["processing_steps"].append(f"Found items for RPOs: {list(items_by_rpo.keys())}")

            # STEP 3: Process each RPO with its items
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

            debug["processing_time"] = str(datetime.now() - start_time)
            debug["total_rpos"] = len(purchase_orders)

            # Return appropriate structure
            if len(purchase_orders) == 1:
                # Single RPO - return simple structure
                single_rpo = purchase_orders[0]
                return {
                    "po_number": single_rpo["po_number"],
                    "global": single_rpo["global"],
                    "items": single_rpo["items"],
                    "item_count": single_rpo["item_count"],
                    "component_count": single_rpo["component_count"],
                    "debug": debug
                }
            else:
                # Multiple RPOs - return purchase_orders structure
                return {
                    "purchase_orders": purchase_orders,
                    "summary": {
                        "total_pos": len(purchase_orders),
                        "total_items": sum(po["item_count"] for po in purchase_orders),
                        "total_components": sum(po["component_count"] for po in purchase_orders)
                    },
                    "debug": debug
                }

        except Exception as e:
            debug["processing_time"] = str(datetime.now() - start_time)
            return {
                "error": "Processing failed",
                "details": str(e),
                "debug": debug,
                "traceback": traceback.format_exc()
            }

# Usage Example
def main():
    """Example usage"""
    extractor = FastPDFOCRExtractor()
    
    # Test with your PDF file
    pdf_path = "path_to_your_pdf.pdf"
    
    try:
        with open(pdf_path, 'rb') as pdf_file:
            result = extractor.extract(pdf_file)
            
            if "error" not in result:
                print("✅ Extraction successful!")
                
                if "purchase_orders" in result:
                    print(f"📊 Found {len(result['purchase_orders'])} RPOs:")
                    
                    for i, po in enumerate(result["purchase_orders"]):
                        print(f"\n📋 RPO {i+1}: {po['po_number']}")
                        print(f"   ├── Global Fields: {len(po['global'])}")
                        print(f"   ├── Items: {po['item_count']}")
                        print(f"   └── Components: {po['component_count']}")
                        
                        # Show first few items
                        for j, item in enumerate(po['items'][:3]):
                            print(f"      Item {j+1}: {item.get('Richline Item #', 'N/A')}")
                            print(f"         ├── Job #: {item.get('Job #', 'N/A')}")
                            print(f"         ├── Metal: {item.get('Metal 1', 'N/A')}")
                            print(f"         └── Components: {len(item.get('Components', []))}")
                else:
                    print(f"📋 Single RPO: {result.get('po_number', 'N/A')}")
                    print(f"   ├── Items: {result['item_count']}")
                    print(f"   └── Components: {result['component_count']}")
                
                print(f"\n⏱️ Processing time: {result['debug'].get('processing_time', 'N/A')}")
                
            else:
                print("❌ Extraction failed:")
                print(f"   Error: {result['error']}")
                print(f"   Details: {result.get('details', 'N/A')}")
                
    except FileNotFoundError:
        print(f"❌ File not found: {pdf_path}")
    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")

if __name__ == "__main__":
    main()
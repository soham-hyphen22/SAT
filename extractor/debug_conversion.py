import re
import pytesseract
import pdf2image
import cv2
import numpy as np
from PIL import Image

class PDFOCRExtractor:
    def __init__(self):
        self.global_patterns = {
            "PO #": r"(RPO\d+)",
            "PO Date": r"\b(\d{2}/\d{2}/\d{2,4})\b",
            "Location": r"Location[:\s]*([A-Z]{2})",
        }
        self.job_pattern = r"(RFP\d{6,}|RSET\d{6,})"
        self.component_heading_keywords = ["Component", "Cost", "Tot. Weight", "Supply"]

    def convert_pdf_to_image(self, pdf_file):
        poppler_path = r"C:\Users\Admin\Documents\Release-24.08.0-0\poppler-24.08.0\Library\bin"
        try:
            # Reset file pointer to beginning
            pdf_file.seek(0)
            images = pdf2image.convert_from_bytes(
                pdf_file.read(),
                dpi=300,
                first_page=1,
                last_page=1,
                poppler_path=poppler_path
            )
            return images[0] if images else None
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
            return image  # Return original image if preprocessing fails

    def extract_text(self, image):
        try:
            return pytesseract.image_to_string(image)
        except Exception as e:
            print(f"OCR extraction failed: {e}")
            return ""

    def validate_extracted_data(self, global_data, item_data, text):
        """Validate if meaningful data was extracted"""
        # Check if we have any meaningful text
        if not text or len(text.strip()) < 10:
            return False, "No readable text found in PDF"
        
        # Check if we have at least some key fields
        key_indicators = [
            bool(global_data.get("PO #")),
            bool(global_data.get("Vendor ID #")),
            bool(item_data.get("Job  ")),
            bool(item_data.get("Richline Item  ")),
            len(global_data) > 0,
            len(item_data) > 2  # More than just Components and CAST Fin WT
        ]
        
        # If we have at least 2 key indicators, consider it valid
        if sum(key_indicators) >= 2:
            return True, "Valid data extracted"
        
        # Check for common PDF issues
        if "password" in text.lower():
            return False, "PDF appears to be password-protected"
        
        if len(text.strip()) < 50:
            return False, "PDF may be corrupted or contain non-text content"
        
        return False, "No extractable data found in PDF"

    def extract_global_fields(self, lines, full_text):
        result = {}

        # Extract Gold, Platinum, Silver from summary table
        for i, line in enumerate(lines):
            if all(word in line for word in ["Gold", "Platinum", "Silver"]):
                if i + 1 < len(lines):
                    number_line = lines[i + 1]
                    numbers = re.findall(r"[\d,]+\.\d+", number_line)
                    if len(numbers) == 3:
                        result["Gold Rate"] = numbers[0].replace(",", "") 
                        result["Platinum Rate"] = numbers[1].replace(",", "") 
                        result["Silver Rate"] = numbers[2].replace(",", "") 
                break

        # Extract all other global fields
        for field, pattern in self.global_patterns.items():
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                result[field] = match.group(1).replace(",", "")

        # Vendor ID and Vendor Name
        for i, line in enumerate(lines):
            if "Vendor ID" in line:
                # Vendor ID
                match = re.search(
                    r"Vendor ID\s*#?\s*[:\-]?\s*([A-Za-z0-9\-]+)", line, re.IGNORECASE
                )
                if match:
                    result["Vendor ID #"] = match.group(1).strip()
                else:
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if next_line and len(next_line) < 30:
                            result["Vendor ID #"] = next_line

                # Vendor Name
                vendor_name_lines = []
                skip_keywords = [
                    "ID", "DATE", "PO", "ORDER", "LOCATION", "RATE", "TYPE", "#",
                    "SHIP", "BILL", "TO", "SUPPLY", "CERT", "SEND", "POLICY"
                ]
                for j in range(1, 4):
                    if i + j < len(lines):
                        possible_name = lines[i + j].strip()
                        if not possible_name:
                            continue
                        if any(possible_name.upper().startswith(word) for word in skip_keywords):
                            continue
                        if len(possible_name) < 4:
                            continue
                        vendor_name_lines.append(possible_name)
                if vendor_name_lines:
                    vendor_full = " ".join(vendor_name_lines)
                    vendor_full = re.split(
                        r"\bTo\s*[:\-]?\s|Ship\s*To|Bill\s*To|Send\s*To|Deliver\s*To|For\s*[:\-]?\s",
                        vendor_full,
                        flags=re.IGNORECASE,
                    )[0].strip(" .:-")
                    result["Vendor Name"] = vendor_full
                break

        # Order Type
        for i, line in enumerate(lines):
            if "Order Type" in line:
                match = re.search(
                    r"\b(ASSAY|ASSET|ASSETKM-AD|CHARGEBACK|CONFONLY|CORRECT|DNP|DOTCOM|DOTCOMB|EXTEND|FL-RECIEVE|IGI|MANUAL|MC|MCH|MCH-REV|MST|NEW-CLR|PCM|PKG|PSAMPLE|REP|RMC|RPR|RTV|SGI|SHW|SLD|SLDSPC|SMG|SMP|SMPGEM|SPC|SPO-BUILD|STOCK|SUPPLY|TST)\b",
                    line,
                    re.IGNORECASE,
                )
                if match:
                    result["Order Type"] = match.group(1).upper()
                elif i + 1 < len(lines):
                    match = re.search(
                        r"\b(ASSAY|ASSET|ASSETKM-AD|CHARGEBACK|CONFONLY|CORRECT|DNP|DOTCOM|DOTCOMB|EXTEND|FL-RECIEVE|IGI|MANUAL|MC|MCH|MCH-REV|MST|NEW-CLR|PCM|PKG|PSAMPLE|REP|RMC|RPR|RTV|SGI|SHW|SLD|SLDSPC|SMG|SMP|SMPGEM|SPC|SPO-BUILD|STOCK|SUPPLY|TST)\b",
                        lines[i + 1],
                        re.IGNORECASE,
                    )
                    if match:
                        result["Order Type"] = match.group(1).upper()
                break

        # Due Date
        for line in lines:
            match = re.search(
                r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}",
                line,
            )
            if match:
                result["Due Date"] = match.group(0)
                break

        return result

    def extract_item_fields(self, lines):
        item = {"Components": [], "CAST Fin WT": {}}

        for i, line in enumerate(lines):
            if "Item No." in line or "Vendor Style" in line:
                for j in range(1, 4):
                    if i + j >= len(lines):
                        break
                    next_line = lines[i + j].strip()
                    match = re.search(r"\b[A-Z]{2}\d{4}[A-Z0-9]{4,8}\b", next_line)
                    if match:
                        item["Richline Item  "] = match.group(0)
                        if i + j + 1 < len(lines):
                            vendor_line = lines[i + j + 1].strip()
                            vendor_match = re.search(
                                r"\b(?:DGE|DGN|BB|LA|BA)\d{3,7}[A-Z]?\b", vendor_line
                            )
                            if vendor_match:
                                item["Vendor Item  "] = vendor_match.group(0)
                        break

        for i, line in enumerate(lines):
            match = re.search(self.job_pattern, line)
            if match:
                item["Job  "] = match.group(1)

            if "Stone PC" in line:
                match = re.search(r"Stone PC[:\s]+([0-9]+\.?[0-9]*)", line)
                if match:
                    item["Stone Labor"] = match.group(1)

            if "Labor PC" in line:
                match = re.search(r"Labor PC[:\s]+([0-9]+\.?[0-9]*)", line)
                if match:
                    item["Labor PC"] = match.group(1)

            if "Diamond TW" in line:
                match = re.search(r"Diamond TW[:\s]+([0-9]+\.?[0-9]*)", line)
                if match:
                    item["Diamond Details"] = f"Diamond TW: {match.group(1)}"

            if "LOSS %" in line:
                item["LOSS %"] = {}
                combined_match = re.search(
                    r"Gold:\s*(\d+\.?\d*)%?\s+Silver:\s*(\d+\.?\d*)%?", line
                )
                if combined_match:
                    item["LOSS %"]["Gold"] = combined_match.group(1)
                    item["LOSS %"]["Silver"] = combined_match.group(2)
                else:
                    gold_match = re.search(r"Gold:\s*(\d+\.?\d*)%?", line)
                    if gold_match:
                        item["LOSS %"]["Gold"] = gold_match.group(1)
                    silver_match = re.search(r"Silver:\s*(\d+\.?\d*)%?", line)
                    if silver_match:
                        item["LOSS %"]["Silver"] = silver_match.group(1)

            if "CAST Fin WT" in line:
                match = re.search(r"Gold:\s*([0-9]+\.?[0-9]+)", line)
                if match:
                    item["Fin Weight (Gold)"] = match.group(1)
                    item["CAST Fin WT"]["Gold"] = match.group(1)

            if re.search(r"\bPieces\b.*\bCarats\b", line) or "Pieces/Carats" in line:
                for j in range(1, min(4, len(lines) - i)):
                    next_line = lines[i + j].strip()
                    match = re.search(r"^\s*(\d+)\s*$", next_line)
                    if match:
                        item["Pieces/Carats"] = match.group(1)
                        break

            if re.search(r"Ext\.?\s*Gross.*Wt", line, re.IGNORECASE):
                for j in range(1, min(4, len(lines) - i)):
                    next_line = lines[i + j].strip()
                    match = re.search(r"^\s*([0-9]+\.?[0-9]+)\s*$", next_line)
                    if match:
                        item["Ext. Gross Wt."] = match.group(1)
                        break

            if re.search(r"\bRate\b", line) and not re.search(
                r"(Gold|Platinum|Silver)\s*Rate", line, re.IGNORECASE
            ):
                match = re.search(r"Rate[:\s]*([0-9]+\.?[0-9]*)", line)
                if match:
                    item["Rate"] = match.group(1)
                elif i + 1 < len(lines):
                    match = re.search(r"([0-9]+\.?[0-9]*)", lines[i + 1].strip())
                    if match:
                        item["Rate"] = match.group(1)

            match = re.search(r"\b([0-9]+\.?[0-9]+)\s*GR\b", line, re.IGNORECASE)
            if match:
                item["Total Weight"] = match.group(1)

            if re.search(r"\bLock\b", line, re.IGNORECASE):
                item["Lock"] = line.strip()

            # Metal Description, Pieces/Carats, Ext. Gross Wt.
            metal_types = [
                "10K", "14K", "18K", "22K", "24K", "14KW", "14KY", "18KW", "18KY",
                "10KW", "10KY", "3KW", "3KY"
            ]
            jewelry_types = [
                "RING", "BAND", "CHAIN", "PENDANT", "EARRING", "RNG", "BRIDAL", "SET"
            ]

            if any(mt in line.upper() for mt in metal_types) and any(
                jt in line.upper() for jt in jewelry_types
            ):
                item["Metal Description"] = line.strip()

                tokens = line.split()
                try:
                    ea_idx = tokens.index("EA")
                    # Pieces/Carats
                    if ea_idx - 1 >= 0 and re.match(
                        r"^\d+$", tokens[ea_idx - 1].replace(",", "")
                    ):
                        item["Pieces/Carats"] = tokens[ea_idx - 1].replace(",", "")
                    # Ext. Gross Wt.
                    if ea_idx + 1 < len(tokens) and re.match(
                        r"^\d+(\.\d+)?$", tokens[ea_idx + 1].replace(",", "")
                    ):
                                                item["Ext. Gross Wt."] = tokens[ea_idx + 1].replace(",", "")
                except ValueError:
                    pass

            match = re.search(r"\b(\d+K[MWY]?)\b", line, re.IGNORECASE)
            if match:
                item["Metal Category"] = match.group(1).upper()

            match = re.search(r"SIZE\s+(\d+(?:\.\d+)?)", line, re.IGNORECASE)
            if match:
                item["Size"] = match.group(1)

            if re.search(r"Send To|Supply Policy", line, re.IGNORECASE):
                policy_match = re.search(
                    r"(?:Send To|Supply Policy)[:\s]+([A-Za-z\s]+)", line, re.IGNORECASE
                )
                if policy_match:
                    item["Supply Policy"] = policy_match.group(1).strip()
                elif i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if (
                        next_line
                        and len(next_line) < 50
                        and re.match(r"^[A-Za-z\s]+$", next_line)
                    ):
                        item["Supply Policy"] = next_line

        item["Components"] = self.extract_components(lines)
        return item

    def extract_components(self, lines):
        components = []
        in_table = False

        for idx, line in enumerate(lines):
            if not in_table and all(
                keyword.lower() in line.lower()
                for keyword in self.component_heading_keywords
            ):
                in_table = True
                continue

            if in_table:
                columns = re.split(r"\s{2,}|\t+", line.strip())

                if len(columns) < 2:
                    break

                if len(columns) >= 4 and re.search(r"\d", columns[3]):
                    comp = {
                        "Component": columns[0] if len(columns) > 0 else "",
                        "Cost ($)": columns[3] if len(columns) > 3 else "",
                        "Tot. Weight": columns[6] if len(columns) > 6 else "",
                        "Supply Policy": columns[10] if len(columns) > 10 else "",
                    }
                    components.append(comp)

        return components

    def extract(self, pdf_file):
        debug = {"processing_steps": []}

        try:
            # Convert PDF to image
            image = self.convert_pdf_to_image(pdf_file)
            debug["processing_steps"].append("PDF converted to image")

            if not image:
                return {
                    "error": "Failed to convert PDF to image", 
                    "details": "The PDF may be corrupted, password-protected, or contain non-text content",
                    "debug": debug
                }

            # Preprocess image
            preprocessed_image = self.preprocess_image(image)
            debug["processing_steps"].append("Image preprocessed")

            # Extract text via OCR
            text = self.extract_text(preprocessed_image)
            debug["processing_steps"].append("Text extracted via OCR")
            
            if not text:
                return {
                    "error": "No extractable data found in PDF",
                    "details": "OCR failed to extract any text from the PDF",
                    "debug": debug
                }

            lines = text.splitlines()
            debug["ocr_text_length"] = len(text)

            # Extract data
            global_data = self.extract_global_fields(lines, text)
            debug["processing_steps"].append(f"Extracted {len(global_data)} global fields")

            item_data = self.extract_item_fields(lines)
            debug["processing_steps"].append("Extracted item fields")

            # Validate extracted data
            is_valid, validation_message = self.validate_extracted_data(global_data, item_data, text)
            debug["validation_message"] = validation_message

            if not is_valid:
                return {
                    "error": "No extractable data found in PDF",
                    "details": validation_message,
                    "debug": debug,
                    "raw_text_sample": text[:200] + "..." if len(text) > 200 else text
                }

            # Process successful extraction
            if "Job #" in item_data:
                global_data["Job #"] = item_data["Job #"]

            debug["processing_steps"].append(
                f"Extracted {len(item_data['Components'])} components"
            )

            all_fields = {
                **global_data,
                **{k: v for k, v in item_data.items() if k != "Components"},
            }
            debug["total_fields_extracted"] = len(all_fields)

            return {
                "global": global_data, 
                "items": [item_data], 
                "debug": debug
            }

        except Exception as e:
            return {
                "error": "Processing failed",
                "details": f"An unexpected error occurred: {str(e)}",
                "debug": debug
            }
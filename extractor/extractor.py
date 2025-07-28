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
        self.component_heading_keywords = [
            "Component",
            "Cost",
            "Tot. Weight",
            "Supplied By",
        ]

    def convert_pdf_to_image(self, pdf_file):
        poppler_path = (
            r"C:\Users\Admin\Documents\Release-24.08.0-0\poppler-24.08.0\Library\bin"
        )
        try:
            pdf_file.seek(0)
            # Convert ALL pages, not just the first one
            images = pdf2image.convert_from_bytes(
                pdf_file.read(),
                dpi=300,
                poppler_path=poppler_path,
            )
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
            _, thresh = cv2.threshold(
                denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
            )
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

    def validate_extracted_data(self, global_data, items_data, text):
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

    def extract_global_fields(self, lines, full_text):
        result = {}

        # Extract metal rates
        for i, line in enumerate(lines):
            if all(word in line for word in ["Gold", "Platinum", "Silver"]):
                if i + 1 < len(lines):
                    number_line = lines[i + 1]
                    numbers = re.findall(r"[\d,]+.\d+", number_line)
                    if len(numbers) == 3:
                        result["Gold Rate"] = numbers[0].replace(",", "")
                        result["Platinum Rate"] = numbers[1].replace(",", "")
                        result["Silver Rate"] = numbers[2].replace(",", "")
                        break

        # Extract PO patterns
        for field, pattern in self.global_patterns.items():
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                result[field] = match.group(1).replace(",", "")

        # Extract Vendor ID and Name
        for i, line in enumerate(lines):
            if "Vendor ID" in line:
                match = re.search(
                    r"Vendor ID\s*#?\s*[:-]?\s*([A-Za-z0-9-]+)", line, re.IGNORECASE
                )
                if match:
                    result["Vendor ID #"] = match.group(1).strip()
                else:
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if next_line and len(next_line) < 30:
                            result["Vendor ID #"] = next_line

                # Extract vendor name
                vendor_name_lines = []
                skip_keywords = [
                    "ID", "DATE", "PO", "ORDER", "LOCATION", "RATE", "TYPE", "#",
                    "SHIP", "BILL", "TO", "SUPPLY", "CERT", "SEND", "POLICY",
                ]

                for j in range(1, 4):
                    if i + j < len(lines):
                        possible_name = lines[i + j].strip()
                        if not possible_name:
                            continue
                        if any(
                            possible_name.upper().startswith(word)
                            for word in skip_keywords
                        ):
                            continue
                        if len(possible_name) < 4:
                            continue
                        vendor_name_lines.append(possible_name)

                if vendor_name_lines:
                    vendor_full = " ".join(vendor_name_lines)
                    vendor_full = re.split(
                        r"\bTo\s*[:-]?\s|Ship\sTo|Bill\sTo|Send\sTo|Deliver\sTo|For\s*[:-]?\s",
                        vendor_full,
                        flags=re.IGNORECASE,
                    )[0].strip(" .:-")
                    result["Vendor Name"] = vendor_full
                break

        # Extract Order Type
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

        # Extract Due Date
        for line in lines:
            match = re.search(
                r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}",
                line,
            )
            if match:
                result["Due Date"] = match.group(0)
                break

        return result

    def extract_items_from_text(self, lines):
        """Extract all items from the text"""
        items = []
        current_item = None

        for i, line in enumerate(lines):
            # Look for item number patterns (DA followed by numbers and letters)
            item_match = re.search(r'\b(DA\d{4}[A-Z0-9]+)\b', line)
            if item_match:
                # If we have a current item, save it
                if current_item:
                    items.append(current_item)

                # Start new item
                current_item = {"Components": [], "CAST Fin WT": {}}
                current_item["Richline Item "] = item_match.group(1)

                # Look for vendor item in the same line or next few lines
                vendor_match = re.search(r'\b(DA\d{4})\b', line)
                if vendor_match and vendor_match.group(1) != item_match.group(1):
                    current_item["Vendor Item "] = vendor_match.group(1)

                # Extract other data from this line and following lines
                self.extract_item_data_from_context(current_item, lines, i)

        # Don't forget the last item
        if current_item:
            items.append(current_item)

        return items

    def extract_item_data_from_context(self, item, lines, start_idx):
        """Extract item-specific data from the context around the item"""
        # Look in the next 20 lines for item-specific data
        end_idx = min(start_idx + 20, len(lines))

        for i in range(start_idx, end_idx):
            line = lines[i]

            # Stop if we hit another item
            if i > start_idx and re.search(r'\bDA\d{4}[A-Z0-9]+\b', line):
                break

            # Extract Job number
            job_match = re.search(self.job_pattern, line)
            if job_match:
                item["Job "] = job_match.group(1)

            # Extract Stone PC
            if "Stone PC" in line:
                match = re.search(r"Stone PC[:\s]+([0-9]+\.?[0-9]*)", line)
                if match:
                    item["Stone PC"] = match.group(1)

            # Extract Labor PC
            if "Labor PC" in line:
                match = re.search(r"Labor PC[:\s]+([0-9.]+)", line, re.IGNORECASE)
                if match:
                    item["Labor PC"] = match.group(1).strip()

            # Extract Diamond TW
            if "Diamond TW" in line:
                match = re.search(r"Diamond TW[:\s]+([0-9]+\.?[0-9]*)", line)
                if match:
                    item["Diamond Details"] = f"Diamond TW: {match.group(1)}"

            # Extract CAST Fin WT
            if "CAST Fin WT" in line:
                gold_match = re.search(r"Gold:\s*([0-9]+\.?[0-9]*)", line)
                silver_match = re.search(r"Silver:\s*([0-9]+\.?[0-9]*)", line)
                if gold_match:
                    item["CAST Fin WT"]["Gold"] = gold_match.group(1)
                    item["Fin Weight (Gold)"] = gold_match.group(1)
                if silver_match:
                    item["CAST Fin WT"]["Silver"] = silver_match.group(1)
                    item["Fin Weight (Silver)"] = silver_match.group(1)

            # Extract LOSS %
            if "LOSS %" in line:
                item["LOSS %"] = {}
                gold_match = re.search(r"Gold:\s*([0-9]+\.?[0-9]*)%?", line)
                silver_match = re.search(r"Silver:\s*([0-9]+\.?[0-9]*)%?", line)
                if gold_match:
                    item["LOSS %"]["Gold"] = gold_match.group(1)
                if silver_match:
                    item["LOSS %"]["Silver"] = silver_match.group(1)

            # Extract pieces/carats and other data from description lines
            if any(unit in line for unit in ["EA", "PR"]) and any(mt in line.upper() for mt in ["EA", "PR"]):
                tokens = line.split()
                try:
                    unit_idx = -1
                    for unit in ["EA", "PR"]:
                        if unit in tokens:
                            unit_idx = tokens.index(unit)
                            break
                    if unit_idx > 0:
                        item["Pieces/Carats"] = tokens[unit_idx - 1]
                    if unit_idx >= 0 and unit_idx + 1 < len(tokens):
                        item["Ext. Gross Wt."] = tokens[unit_idx + 1]
                except (ValueError, IndexError):
                    pass

            # Extract metal category
            if "Metal Category" not in item and any(unit in line for unit in ["EA", "PR"]):
                karat_match = re.search(r'\b(\d+K[A-Z]*)\b', line.upper())
                if karat_match:
                    item["Metal Category"] = karat_match.group(1)

            # Extract total weight
            weight_match = re.search(r"\b([0-9]+\.?[0-9]+)\s*GR\b", line, re.IGNORECASE)
            if weight_match:
                item["Total Weight"] = weight_match.group(1)

        # Extract components for this item
        item["Components"] = self.extract_components_for_item(lines, start_idx, end_idx)

    def extract_components_for_item(self, lines, start_idx, end_idx):
        """Extract components for a specific item"""
        components = []

        # Look for component table within the item's section
        table_start = -1

        for idx in range(start_idx, end_idx):
            if idx >= len(lines):
                break

            line = lines[idx]
            line_lower = line.lower()

            # Look for component table header
            table_keywords = ["supplied", "component", "cost", "weight", "policy"]
            keyword_count = sum(1 for keyword in table_keywords if keyword in line_lower)

            if keyword_count >= 3:
                table_start = idx
                break

        if table_start == -1:
            return components

        # Process component rows
        for idx in range(table_start + 1, end_idx):
            if idx >= len(lines):
                break

            line = lines[idx].strip()

            if not line:
                continue

            # Stop if we hit another item or obvious end markers
            if (re.search(r'\bDA\d{4}[A-Z0-9]+\b', line) or
                any(marker in line for marker in ["Total", "Subtotal", "Grand", "Summary"])):
                break

            # Skip header-like lines (This loop is logically ineffective but kept from original)
            for keyword in ["supplied", "component", "cost", "weight", "policy"]:
                continue

            # Parse component data
            row_data = line.split()
            row_data = [col.strip() for col in row_data if col.strip()]

            if not row_data:
                continue

            # Extract component information
            component = ""
            cost = ""
            weight = ""
            policy = ""

            # First item is usually the component name
            component = row_data[0] if row_data else ""

            # Look for cost pattern (number followed by CT)
            for i, item in enumerate(row_data):
                if (re.match(r"^\d+\.\d+$", item) and
                    i + 1 < len(row_data) and
                    row_data[i + 1].upper() == "CT"):
                    cost = f"{item} CT"
                    break

            # Look for weight pattern (number followed by CT without space)
            for item in row_data:
                if re.match(r"^\d+\.\d+CT$", item, re.IGNORECASE):
                    weight = item
                    break

            # Look for "Send To" pattern
            for i, item in enumerate(row_data):
                if (item.lower() == "send" and
                    i + 1 < len(row_data) and
                    row_data[i + 1].lower() == "to"):
                    policy = "Send To"
                    break
                elif "Vendor" in line:
                    policy = "By Vendor"

            # Only add if we have a valid component
            if component and len(component) > 1 and not component.isdigit():
                comp = {
                    "Component": component,
                    "Cost ($)": cost,
                    "Tot. Weight": weight,
                    "Supply Policy": policy,
                }
                components.append(comp)

        return components

    def extract(self, pdf_file):
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

            # Extract global data (same for all items)
            global_data = self.extract_global_fields(all_lines, all_text)
            debug["processing_steps"].append(f"Extracted {len(global_data)} global fields")

            # Extract all items from the combined text
            items_data = self.extract_items_from_text(all_lines)
            debug["processing_steps"].append(f"Extracted {len(items_data)} items")

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
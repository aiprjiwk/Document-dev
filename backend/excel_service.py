import io
import logging
import openpyxl
from typing import Dict, Any, Union

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("ExcelExportService")

class ExcelExportService:
    """
    Service responsible for loading an Excel workbook template, populating cells
    using mapped OCR keys, and exporting the completed file.
    """
    
    def __init__(self, target_sheet_name: str = "Adjustment Table"):
        self.target_sheet_name = target_sheet_name

    def export_excel(
        self,
        template_bytes: bytes,
        ocr_data: Dict[str, Any],
        mapping: Dict[str, str]
    ) -> bytes:
        """
        Populates Excel template cells based on the provided mapping dictionary and OCR results.
        
        Args:
            template_bytes (bytes): The raw bytes of the template Excel workbook.
            ocr_data (Dict[str, Any]): Dictionary containing field names and extracted values.
            mapping (Dict[str, str]): Dictionary mapping field names to Excel cell addresses.
            
        Returns:
            bytes: The completed Excel workbook bytes ready for export.
        """
        logger.info("Initializing Excel template writeback process...")
        
        try:
            # Load template workbook
            wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
            
            # Select target worksheet
            sheet_name = self.target_sheet_name
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                logger.info(f"Target sheet '{sheet_name}' selected successfully.")
            else:
                ws = wb.active
                logger.warning(
                    f"Target sheet '{sheet_name}' not found. "
                    f"Falling back to active sheet: '{ws.title}'"
                )
                
            # Iterate through cell mapping keys
            write_count = 0
            for field_name, cell_coord in mapping.items():
                if field_name in ocr_data:
                    raw_val = ocr_data[field_name]
                    
                    # Convert to string and clean spaces
                    val_str = str(raw_val).strip() if raw_val is not None else ""
                    
                    if val_str == "":
                        ws[cell_coord] = None
                        logger.info(f"Cleared cell {cell_coord} for empty Field: '{field_name}'")
                    # Preserve leading zeros for values like "002"
                    elif val_str.startswith('0') and len(val_str) > 1 and not val_str.startswith('0.'):
                        ws[cell_coord] = val_str
                        logger.info(f"Wrote string cell {cell_coord} for Field: '{field_name}' -> '{val_str}' (preserved leading zeros)")
                    else:
                        # Attempt float/int parsing, fallback to raw string on failure
                        try:
                            if '.' in val_str:
                                ws[cell_coord] = float(val_str)
                            else:
                                ws[cell_coord] = int(val_str)
                        except ValueError:
                            ws[cell_coord] = raw_val
                        logger.info(f"Wrote numeric cell {cell_coord} for Field: '{field_name}' -> '{ws[cell_coord].value}'")
                        
                    write_count += 1
                    
            logger.info(f"Completed excel writeback. Wrote {write_count} cells.")
            
            # Save completed workbook to memory bytes
            out_bytes = io.BytesIO()
            wb.save(out_bytes)
            return out_bytes.getvalue()
            
        except Exception as e:
            logger.error(f"Error during excel templates generation: {e}")
            raise ValueError(f"Failed to compile Excel templates: {str(e)}")

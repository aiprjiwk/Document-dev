import io
import logging
import openpyxl
from typing import Dict, Optional

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("MappingService")

class MappingService:
    """
    Service responsible for loading mapping sheet configurations from Excel workbooks
    and returning them as a dictionary of FieldName -> CellAddress.
    """
    
    def __init__(self, default_sheet_name: str = "Mapping"):
        self.default_sheet_name = default_sheet_name

    def load_mappings_from_excel(self, excel_bytes: bytes) -> Dict[str, str]:
        """
        Parses an uploaded Excel template's mapping sheet and extracts cell assignments.
        
        Args:
            excel_bytes (bytes): The raw bytes of the uploaded template workbook.
            
        Returns:
            Dict[str, str]: A dictionary mapping FieldName to CellAddress (e.g. {"OrderNo": "B3"}).
        """
        logger.info("Starting cell mapping extraction from Excel bytes...")
        mappings: Dict[str, str] = {}
        
        try:
            # Load workbook in read-only mode for performance
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
            
            # Find the mapping sheet
            sheet_name = self.default_sheet_name
            if sheet_name not in wb.sheetnames:
                # Fallback to case-insensitive check or active sheet
                found = False
                for name in wb.sheetnames:
                    if "mapping" in name.lower():
                        sheet_name = name
                        found = True
                        break
                if not found:
                    sheet_name = wb.sheetnames[0]
                    logger.warning(
                        f"Sheet '{self.default_sheet_name}' not found. "
                        f"Falling back to first available sheet: '{sheet_name}'"
                    )
            
            sheet = wb[sheet_name]
            logger.info(f"Using sheet '{sheet_name}' for cell mappings.")
            
            # Parse header row to locate "FieldName" and "Cell" columns
            field_name_col_idx: Optional[int] = None
            cell_col_idx: Optional[int] = None
            
            # Read first row for headers
            max_cols = min(sheet.max_column or 10, 100)
            for col in range(1, max_cols + 1):
                val = sheet.cell(row=1, column=col).value
                if val is not None:
                    header_str = str(val).strip().lower()
                    if "field" in header_str or "key" in header_str:
                        field_name_col_idx = col
                    elif "cell" in header_str or "addr" in header_str:
                        cell_col_idx = col
                        
            # Fallback to column 1 and 2 if headers aren't detected
            if field_name_col_idx is None:
                field_name_col_idx = 1
                logger.info("FieldName column not detected in header. Defaulting to Column 1.")
            if cell_col_idx is None:
                cell_col_idx = 2
                logger.info("Cell coordinate column not detected in header. Defaulting to Column 2.")
                
            # Iterate through rows starting from row 2 (skipping header)
            max_rows = min(sheet.max_row or 100, 1000)
            for r in range(2, max_rows + 1):
                f_val = sheet.cell(row=r, column=field_name_col_idx).value
                c_val = sheet.cell(row=r, column=cell_col_idx).value
                
                if f_val is not None and c_val is not None:
                    f_name = str(f_val).strip()
                    c_addr = str(c_val).strip().upper()
                    
                    if f_name and c_addr:
                        # Basic alphanumeric verification of Excel cell formats (e.g. B3, F12)
                        if re.match(r"^[A-Z]+\d+$", c_addr):
                            mappings[f_name] = c_addr
                        else:
                            logger.warning(f"Skipped invalid cell address coordinate: '{c_addr}' on row {r}")
                            
            logger.info(f"Extracted {len(mappings)} valid cell mappings dynamically.")
            return mappings
            
        except Exception as e:
            logger.error(f"Error loading mappings from workbook: {e}")
            # Do not crash the application; raise a clean exception for the caller
            raise ValueError(f"Failed to parse cell mapping configurations: {str(e)}")

# Ensure regex module is imported
import re

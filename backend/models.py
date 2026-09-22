import datetime
from sqlalchemy import Column, Integer, String, Float, DateTime
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()

class OCRData(Base):
    """
    SQLAlchemy Model representing the 'ocr_data' table.
    Stores OCR results, user corrections, confidence levels, and audit logs.
    """
    __tablename__ = "ocr_data"

    id = Column(Integer, primary_key=True, autoincrement=True, index=True)
    document_id = Column(String, nullable=False, index=True)
    field_name = Column(String, nullable=False, index=True)
    ocr_value = Column(String, nullable=True)
    corrected_value = Column(String, nullable=True)
    confidence = Column(Float, nullable=True)
    reviewed_by = Column(String, nullable=True)
    reviewed_date = Column(DateTime, default=datetime.datetime.utcnow, nullable=True)
    status = Column(String, default="Pending", nullable=True)

    def __repr__(self) -> str:
        return (
            f"<OCRData(id={self.id}, document_id='{self.document_id}', "
            f"field_name='{self.field_name}', ocr_value='{self.ocr_value}', "
            f"corrected_value='{self.corrected_value}', status='{self.status}')>"
        )


class OQHMITemplateMaintenance(Base):
    """
    SQLAlchemy Model representing the 'oq_hmi_template_maintenance' table.
    Stores machine-specific screen definitions, Visu types, Function codes, and expected header titles.
    """
    __tablename__ = "oq_hmi_template_maintenance"

    id = Column(Integer, primary_key=True, autoincrement=True, index=True)
    machine_type = Column(String, nullable=False, index=True)
    visu_type = Column(String, default="IPC", nullable=False)
    function_code = Column(String, nullable=True, index=True)
    screen_name = Column(String, nullable=False)
    tab_index = Column(Integer, default=0)
    expected_header_title = Column(String, nullable=False)
    created_at = Column(DateTime, default=datetime.datetime.utcnow, nullable=True)

    def __repr__(self) -> str:
        return (
            f"<OQHMITemplateMaintenance(id={self.id}, machine_type='{self.machine_type}', "
            f"visu_type='{self.visu_type}', function_code='{self.function_code}', "
            f"expected_header_title='{self.expected_header_title}')>"
        )


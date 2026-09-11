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

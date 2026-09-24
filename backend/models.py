import datetime
import os
from sqlalchemy import Column, Integer, String, Float, DateTime, Text, create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

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
    sheet_name = Column(String, nullable=True)
    excel_cell = Column(String, nullable=True)
    ocr_value = Column(String, nullable=True)
    corrected_value = Column(String, nullable=True)
    confidence = Column(Float, nullable=True)
    engine_used = Column(String, nullable=True)
    reviewed_by = Column(String, nullable=True)
    reviewed_date = Column(DateTime, default=datetime.datetime.utcnow, nullable=True)
    status = Column(String, default="Pending", nullable=True)

    def __repr__(self) -> str:
        return (
            f"<OCRData(id={self.id}, document_id='{self.document_id}', "
            f"field_name='{self.field_name}', ocr_value='{self.ocr_value}', "
            f"corrected_value='{self.corrected_value}', status='{self.status}')>"
        )


class CalibrationProfile(Base):
    """
    SQLAlchemy Model representing the 'calibration_profile' table.
    Stores user-specific handwriting character embeddings for few-shot classification.
    """
    __tablename__ = "calibration_profile"

    id = Column(Integer, primary_key=True, autoincrement=True, index=True)
    user_name = Column(String, nullable=False, index=True)
    char_code = Column(String, nullable=False, index=True)  # '0'-'9', 'A'-'Z'
    image_path = Column(String, nullable=True)
    feature_vector_json = Column(Text, nullable=False)  # JSON string of normalized float embedding
    created_at = Column(DateTime, default=datetime.datetime.utcnow, nullable=True)

    def __repr__(self) -> str:
        return (
            f"<CalibrationProfile(id={self.id}, user_name='{self.user_name}', "
            f"char_code='{self.char_code}')>"
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

# Database helper functions
DB_PATH = os.path.join(os.path.dirname(__file__), "..", "database", "app_database.db")

def get_db_engine():
    db_dir = os.path.dirname(DB_PATH)
    os.makedirs(db_dir, exist_ok=True)
    engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
    return engine

def init_db():
    engine = get_db_engine()
    Base.metadata.create_all(bind=engine)
    return engine

def get_db_session():
    engine = init_db()
    SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
    return SessionLocal()

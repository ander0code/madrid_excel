from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    
    APP_TITLE: str = "API de Control de Marcaciones"
    APP_VERSION: str = "1.0.0"
    APP_DESCRIPTION: str = "API para generación de reportes de marcaciones y asistencia"
    
    EXTERNAL_API_URL: str 
    
    EXCEL_OUTPUT_FILE: str = "marcaciones_personal.xlsx"
    EXCEL_SHEET_TITLE: str = "FEBRERO 2025"
    
    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

settings = Settings()
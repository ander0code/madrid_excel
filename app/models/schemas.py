from pydantic import BaseModel, Field
from typing import List, Optional, Union
from datetime import datetime

class Marcacion(BaseModel):
    """Modelo para una marcación de asistencia."""
    fecha: Union[str, datetime] = Field(...)
    hora_ingreso: Optional[str] = None
    hora_salida: Optional[str] = None
    diferencia_ingreso: Optional[int] = None
    diferencia_salida: Optional[int] = None
    marco_ingreso: Optional[bool] = None
    marco_salida: Optional[bool] = None
    ingreso_tarde: Optional[bool] = None
    salida_temprano: Optional[bool] = None

class EmpleadoMarcaciones(BaseModel):
    """Modelo para un empleado con sus marcaciones."""
    emp_code: str
    first_name: Optional[str] = None
    last_name: Optional[str] = None
    hire_date: Optional[Union[str, datetime]] = None
    fecha_cese: Optional[Union[str, datetime]] = None
    is_unactive: Optional[bool] = False
    marcaciones: List[Marcacion] = []
    position_name: Optional[str] = None
    dept_name: Optional[str] = None
    hora_ingreso: Optional[str] = None
    hora_salida: Optional[str] = None
    dias_labores: Optional[str] = None
    dias_descanso: Optional[str] = None
    dias_remoto: List[str] = []

class ReporteRequest(BaseModel):
    """Modelo para la solicitud de generación de reporte."""
    empleados_data: List[EmpleadoMarcaciones]
    fecha_inicio: Optional[str] = None
    fecha_fin: Optional[str] = None

class ResponseEmpleados(BaseModel):
    """Modelo para la respuesta con lista de empleados."""
    empleados: List[EmpleadoMarcaciones]

class ErrorResponse(BaseModel):
    """Modelo para respuestas de error."""
    detail: str
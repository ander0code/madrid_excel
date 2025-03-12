from pydantic import BaseModel
from typing import List, Optional
import datetime

class Marcacion(BaseModel):
    """Modelo para una marcación de asistencia."""
    fecha: Optional[datetime.datetime]
    hora_ingreso: Optional[str]
    hora_salida: Optional[str]
    diferencia_ingreso: Optional[int]
    diferencia_salida: Optional[int]
    marco_ingreso: Optional[bool]
    marco_salida: Optional[bool]
    ingreso_tarde: Optional[bool]
    salida_temprano: Optional[bool]

class EmpleadoMarcaciones(BaseModel):
    """Modelo para un empleado con sus marcaciones."""
    emp_code: str
    first_name: Optional[str]
    last_name: Optional[str]
    hire_date: Optional[datetime.datetime] = None
    fecha_cese: Optional[datetime.datetime] = None
    is_unactive: Optional[bool] = False
    marcaciones: List[Marcacion] = []
    position_name: Optional[str] = None
    dept_name: Optional[str] = None
    hora_ingreso: Optional[str] = None
    hora_salida: Optional[str] = None
    dias_labores: Optional[str] = None
    dias_descanso: Optional[str] = None
    dias_remoto: List[str] = []

class ResponseEmpleados(BaseModel):
    """Modelo para la respuesta con lista de empleados."""
    empleados: List[EmpleadoMarcaciones]

class ErrorResponse(BaseModel):
    """Modelo para respuestas de error."""
    detail: str
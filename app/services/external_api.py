import json
from typing import Optional, List, Dict, Any
from fastapi import HTTPException


async def process_empleados_data(data: List[Dict[str, Any]], fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Procesa los datos de empleados recibidos directamente en el endpoint.

    Args:
        data: Datos de empleados recibidos en el body del request
        fecha_inicio: Fecha inicial en formato YYYY-MM-DD (para referencia)
        fecha_fin: Fecha final en formato YYYY-MM-DD (para referencia)

    Returns:
        Lista de diccionarios con datos de empleados procesados
    """
    try:
        # Validar que se recibieron datos
        if not data:
            raise HTTPException(
                status_code=400,
                detail="No se proporcionaron datos de empleados"
            )

        # Asegurar que los datos están en formato lista
        if not isinstance(data, list):
            data = [data]

        print(f"Procesando datos de {len(data)} empleados")

        # Añade logs detallados para depuración
        if len(data) > 0:
            print(
                f"Ejemplo del primer empleado: {json.dumps(data, default=str)[:500]}...")

        return data

    except Exception as e:
        # Registrar el error con detalles
        print(f"Error al procesar los datos: {str(e)}")

        try:
            with open("test.json", "r", encoding="utf-8") as file:
                fallback_data = json.load(file)
                if not isinstance(fallback_data, list):
                    fallback_data = [fallback_data]

                print("Usando datos de prueba del archivo local")
                return fallback_data
        except Exception as fallback_error:
            print(f"Error al cargar datos de respaldo: {str(fallback_error)}")
            raise HTTPException(
                status_code=503,
                detail=f"Error al procesar los datos: {str(e)}"
            )

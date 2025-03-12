from fastapi import APIRouter, Query
from fastapi.responses import FileResponse, JSONResponse
from typing import Optional

from services.external_api import get_empleados_data
from services.excel_service import generar_excel_marcaciones
from config import settings

router = APIRouter()

@router.get("/marcaciones-excel", response_class=FileResponse)
async def obtener_marcaciones_excel(
    fecha_inicio: Optional[str] = Query(None, description="Fecha inicial en formato YYYY-MM-DD"),
    fecha_fin: Optional[str] = Query(None, description="Fecha final en formato YYYY-MM-DD")
):
    """
    Genera un archivo Excel con marcaciones de personal.
    Opcionalmente filtrado por rango de fechas.
    """
    try:
        empleados_data = await get_empleados_data(fecha_inicio, fecha_fin)
        
        archivo_excel, exito = generar_excel_marcaciones(empleados_data)
        
        if not exito:
            return JSONResponse(
                status_code=500,
                content={"success": False, "message": "Error al generar el archivo Excel"}
            )
        
        nombre_archivo = settings.EXCEL_OUTPUT_FILE
        if fecha_inicio or fecha_fin:
            periodo = f"{fecha_inicio or 'inicio'}_a_{fecha_fin or 'fin'}"
            nombre_archivo = f"marcaciones_{periodo}.xlsx"
        
        headers = {"X-Generation-Success": "true"}
        
        return FileResponse(
            path=archivo_excel,
            filename=nombre_archivo,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )
        
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "message": f"Error al generar el excel: {str(e)}"
            }
        )
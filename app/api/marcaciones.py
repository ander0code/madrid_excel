from fastapi import APIRouter, HTTPException, Response, Request

import traceback

from models.schemas import ReporteRequest
from services.external_api import process_empleados_data
from services.excel_service import generate_excel_report

router = APIRouter()


@router.get("/ping")
async def ping():
    """Endpoint simple para verificar si el servicio está disponible"""
    return {"status": "ok", "message": "Excel service is running"}


@router.post("/marcaciones-excel")
async def generar_reporte_excel(request: ReporteRequest, req: Request):
    """
    Genera un reporte Excel a partir de los datos de empleados recibidos directamente.
    """
    try:
        # Log para debugging
        content_length = req.headers.get("content-length", "desconocido")
        print(
            f"Recibiendo solicitud con Content-Length: {content_length} bytes")

        # Mostrar muestra de los datos recibidos
        if request.empleados_data:
            muestra_empleado = request.empleados_data[0]
            print(f"Empleados: {muestra_empleado}")

        # Procesar los datos recibidos
        print("Procesando datos recibidos...")
        try:
            empleados_data = await process_empleados_data(
                [empleado.model_dump() for empleado in request.empleados_data],
                request.fecha_inicio,
                request.fecha_fin
            )
            print(
                f"Datos procesados correctamente. {len(empleados_data)} empleados listos.")
        except Exception as proc_error:
            print(f"Error procesando datos: {str(proc_error)}")
            traceback.print_exc()
            raise HTTPException(
                status_code=422,
                detail=f"Error al procesar los datos de empleados: {str(proc_error)}"
            )

        # Generar el Excel
        print("Generando Excel...")
        try:
            excel_bytes = await generate_excel_report(
                empleados_data,
                request.fecha_inicio,
                request.fecha_fin
            )
            print(
                f"Excel generado correctamente. Tamaño: {len(excel_bytes) / 1024:.2f} KB")
        except Exception as excel_error:
            print(f"Error generando Excel: {str(excel_error)}")
            traceback.print_exc()
            raise HTTPException(
                status_code=500,
                detail=f"Error al generar el Excel: {str(excel_error)}"
            )

        # Nombre de archivo con fechas si están disponibles
        filename = "marcaciones"
        if request.fecha_inicio:
            filename += f"_desde_{request.fecha_inicio}"
        if request.fecha_fin:
            filename += f"_hasta_{request.fecha_fin}"
        filename += ".xlsx"

        # Retornar el archivo Excel
        response = Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )

        print("Respondiendo con el Excel generado")
        return response

    except HTTPException:
        # Reenviar excepciones HTTP ya creadas
        raise
    except Exception as e:
        print(f"Error no manejado: {str(e)}")
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=f"Error al generar el reporte Excel: {str(e)}"
        )

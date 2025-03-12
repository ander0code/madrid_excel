import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
from typing import List, Dict, Any, Tuple

from utils.formatters import formatear_dias_teletrabajo, formato_tiempo
from config import settings

def generar_excel_marcaciones(empleados_data: List[Dict[str, Any]]) -> Tuple[str, bool]:
    """
    Genera un archivo Excel con las marcaciones de los empleados.
    
    Args:
        empleados_data: Lista de datos de empleados con sus marcaciones
        
    Returns:
        Tuple con (ruta del archivo Excel generado, bool indicando si fue exitoso)
    """
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = settings.EXCEL_SHEET_TITLE


        ws.merge_cells('A1:Z1')
        ws['A1'] = "REPORTE DE CONTROL DE MARCACIONES Y ASISTENCIA DEL PERSONAL"
        ws['A1'].font = Font(size=14, bold=True)
        ws['A1'].alignment = Alignment(horizontal='left')

        subtitulos = [
            "Gestión del Talento Humano",
            "MADRID INGENIEROS SAC",
            "Horario General: 08:30 AM - 6:30 PM",
            "Horario Área de Ventas: 09:00AM - 07:00PM",
            "Horario Practicantes Pre (08:30AM - 03:30PM)"
        ]
        for idx, texto in enumerate(subtitulos, start=2):
            ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=50)
            ws.cell(row=idx, column=1, value=texto).alignment = Alignment(horizontal='left')

        encabezados = [
            "N.", "DNI", "TRABAJADOR", "FECHA INGRESO", "FECHA DE CESE", "CARGO",
            "AREA", "GERENCIA", "ESTADO", "REGISTRO", "DIAS DE LABORES", "DSO",
            "HORARIO OFICIAL", "DÍAS DE TELETRABAJO JD 2025"
        ]
        col = 1
        for encabezado in encabezados:
            ws.merge_cells(start_row=8, start_column=col, end_row=10, end_column=col)
            ws.cell(row=8, column=col, value=encabezado).alignment = Alignment(horizontal='center', vertical='center')
            col += 1

        fechas_dias = [
            ("03 FEBRERO 2025", "LUNES"), ("04 FEBRERO 2025", "MARTES"), ("05 FEBRERO 2025", "MIÉRCOLES"),
            ("06 FEBRERO 2025", "JUEVES"), ("07 FEBRERO 2025", "VIERNES"), ("08 FEBRERO 2025", "SÁBADO"),
            ("09 FEBRERO 2025", "DOMINGO"), ("10 FEBRERO 2025", "LUNES"), ("11 FEBRERO 2025", "MARTES"),
            ("12 FEBRERO 2025", "MIÉRCOLES"), ("13 FEBRERO 2025", "JUEVES"), ("14 FEBRERO 2025", "VIERNES")
        ]
        fecha_col_map = {}
        columnas_fecha = ["ING", "TAR", "SALIDA", "EXT"]
        color_rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        color_encabezado = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        
        for fecha, dia in fechas_dias:
            fecha_iso = f"2025-02-{int(fecha[:2]):02d}"
            fecha_col_map[fecha_iso] = col
            ws.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col+3)
            ws.cell(row=8, column=col, value=fecha).alignment = Alignment(horizontal='center', vertical='center')
            ws.merge_cells(start_row=9, start_column=col, end_row=9, end_column=col+3)
            ws.cell(row=9, column=col, value=dia).alignment = Alignment(horizontal='center', vertical='center')
            for i, sub in enumerate(columnas_fecha):
                celda = ws.cell(row=10, column=col+i, value=sub)
                celda.alignment = Alignment(horizontal='center', vertical='center')
                if sub in ["TAR", "EXT"]:
                    celda.fill = color_rojo
                    celda.font = Font(color="FFFFFF", bold=True)
            col += 4

        for row in ws.iter_rows(min_row=8, max_row=10, min_col=1, max_col=col-1):
            for celda in row:
                if not celda.fill.start_color.index == "FF0000":
                    celda.fill = color_encabezado
                    celda.font = Font(color="FFFFFF", bold=True)

        fila_actual = 11

        empleados_validos = [e for e in empleados_data if isinstance(e, dict) and e.get("emp_code")]
        
        for idx, empleado in enumerate(empleados_validos, 1):
            try:
                # Número y DNI
                ws.cell(row=fila_actual, column=1, value=idx)
                ws.cell(row=fila_actual, column=2, value=empleado.get("emp_code", ""))
                

                first_name = empleado.get("first_name", "") or ""
                last_name = empleado.get("last_name", "") or ""
                nombre_completo = f"{first_name} {last_name}".strip()
                ws.cell(row=fila_actual, column=3, value=nombre_completo if nombre_completo else "-")
                
                if empleado.get("hire_date"):
                    try:
                        fecha_ingreso = datetime.strptime(empleado["hire_date"], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d/%m/%Y")
                        ws.cell(row=fila_actual, column=4, value=fecha_ingreso)
                    except (ValueError, TypeError):
                        ws.cell(row=fila_actual, column=4, value="-")
                else:
                    ws.cell(row=fila_actual, column=4, value="-")
                
                fecha_cese = None
                tiene_fecha_cese = False
                fecha_cese_str = "-"

                if empleado.get("fecha_cese"):
                    try:
                        fecha_cese = datetime.strptime(empleado["fecha_cese"], "%Y-%m-%dT%H:%M:%S.%fZ")
                        fecha_cese_str = fecha_cese.strftime("%d/%m/%Y")
                        tiene_fecha_cese = True
                    except (ValueError, TypeError):
                        pass
                        
                ws.cell(row=fila_actual, column=5, value=fecha_cese_str)
                
                ws.cell(row=fila_actual, column=6, value=empleado.get("position_name", "-"))
                
                dept_name = empleado.get("dept_name", "-")
                ws.cell(row=fila_actual, column=7, value=dept_name)
                
                ws.cell(row=fila_actual, column=8, value=empleado.get("gerencia", "-"))

                if tiene_fecha_cese:
                    estado = "Cesado"
                elif empleado.get("is_unactive", False):
                    estado = "Inactivo"
                else:
                    estado = "Activo"
                
                ws.cell(row=fila_actual, column=9, value=estado)
                
                ws.cell(row=fila_actual, column=10, value=empleado.get("registro", "-"))
                
                dias_labores = empleado.get("dias_labores", "-")
                if dias_labores == "lun-vier":
                    ws.cell(row=fila_actual, column=11, value="LUNES A VIERNES")
                else:
                    ws.cell(row=fila_actual, column=11, value=dias_labores.upper() if dias_labores else "-")
                
                dias_descanso = empleado.get("dias_descanso", "-")
                if dias_descanso == "sab-dom":
                    ws.cell(row=fila_actual, column=12, value="S Y D")
                else:
                    ws.cell(row=fila_actual, column=12, value=dias_descanso.upper() if dias_descanso else "-")
                
                if empleado.get("hora_ingreso") and empleado.get("hora_salida"):
                    horario = f"{empleado['hora_ingreso']}AM - {empleado['hora_salida']}PM"
                    ws.cell(row=fila_actual, column=13, value=horario)
                else:
                    ws.cell(row=fila_actual, column=13, value="-")

                dias_remoto = empleado.get("dias_remoto", [])
                ws.cell(row=fila_actual, column=14, value=formatear_dias_teletrabajo(dias_remoto))

                if "marcaciones" in empleado and isinstance(empleado["marcaciones"], list):
                    for marcacion in empleado["marcaciones"]:
                        try:
                            if not isinstance(marcacion, dict) or "fecha" not in marcacion:
                                continue
                                
                            fecha_marca = datetime.strptime(marcacion["fecha"], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%Y-%m-%d")
                            
                            if fecha_marca in fecha_col_map:
                                col_inicio = fecha_col_map[fecha_marca]
                                
                                ws.cell(row=fila_actual, column=col_inicio, 
                                       value=marcacion.get("hora_ingreso", "-"))
                                
                                if marcacion.get("ingreso_tarde", False) and "diferencia_ingreso" in marcacion:
                                    tar_valor = formato_tiempo(abs(marcacion['diferencia_ingreso']))
                                else:
                                    tar_valor = "0"
                                ws.cell(row=fila_actual, column=col_inicio+1, value=tar_valor)

                                ws.cell(row=fila_actual, column=col_inicio+2, value=marcacion.get("hora_salida", "-"))

                                if marcacion.get("salida_temprano", False) and "diferencia_salida" in marcacion:
                                    ext_valor = f"-{formato_tiempo(abs(marcacion['diferencia_salida']))}"
                                else:
                                    ext_valor = "0"
                                ws.cell(row=fila_actual, column=col_inicio+3, value=ext_valor)
                                
                                ws.cell(row=fila_actual, column=col_inicio+1).fill = color_rojo
                                ws.cell(row=fila_actual, column=col_inicio+1).font = Font(color="FFFFFF", bold=True)
                                ws.cell(row=fila_actual, column=col_inicio+3).fill = color_rojo
                                ws.cell(row=fila_actual, column=col_inicio+3).font = Font(color="FFFFFF", bold=True)
                        except Exception as e:
                            print(f"Error en marcación: {str(e)}")
                            continue
            
            except Exception as e:
                print(f"Error procesando empleado {idx}: {str(e)}")
            
            fila_actual += 1

        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        for row in ws.iter_rows(min_row=8, max_row=fila_actual-1, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border

        anchos_personalizados = {
            'A': 5, 'B': 12, 'C': 25, 'D': 15, 'E': 15, 'F': 20,
            'G': 15, 'H': 10, 'I': 12, 'J': 12, 'K': 18, 'L': 10,
            'M': 20, 'N': 20
        }
        
        for letra, ancho in anchos_personalizados.items():
            ws.column_dimensions[letra].width = ancho
            
        for idx in range(15, col):
            ws.column_dimensions[get_column_letter(idx)].width = 10

        archivo_salida = settings.EXCEL_OUTPUT_FILE
        wb.save(archivo_salida)
        
        if os.path.exists(archivo_salida) and os.path.getsize(archivo_salida) > 0:
            return archivo_salida, True
        else:
            return archivo_salida, False
            
    except Exception as e:
        print(f"Error al generar el Excel: {str(e)}")
        return settings.EXCEL_OUTPUT_FILE, False
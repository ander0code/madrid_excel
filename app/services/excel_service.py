import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import sys
import io
import os
from typing import List, Dict, Any, Optional

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.formatters import formatear_dias_teletrabajo
from config import settings


COLOR_ROJO = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
COLOR_AMARILLO = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
COLOR_VERDE = PatternFill(start_color="538D22", end_color="538D22", fill_type="solid")
COLOR_ENCABEZADO = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

MARGEN_TOLERANCIA = 5

async def generate_excel_report(empleados_data: List[Dict[str, Any]], fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None) -> bytes:
    """
    Genera un archivo Excel con las marcaciones de los empleados y lo devuelve como bytes.
    
    Args:
        empleados_data: Lista de datos de empleados con sus marcaciones
        fecha_inicio: Fecha inicial en formato YYYY-MM-DD
        fecha_fin: Fecha final en formato YYYY-MM-DD
        
    Returns:
        Bytes del archivo Excel generado
    """
    try:
        print(f"Generando Excel con {len(empleados_data)} empleados...")
        print(f"Rango de fechas: {fecha_inicio} a {fecha_fin}")
        print(f"Margen de tolerancia configurado: {MARGEN_TOLERANCIA} minutos")
        
        if fecha_inicio and fecha_fin:
            try:
                fecha_inicio_dt = datetime.strptime(fecha_inicio, "%Y-%m-%d")
                fecha_fin_dt = datetime.strptime(fecha_fin, "%Y-%m-%d")
                
                if fecha_fin_dt < fecha_inicio_dt:
                    print("Advertencia: Fecha de fin anterior a fecha de inicio. Invirtiendo el rango.")
                    fecha_inicio_dt, fecha_fin_dt = fecha_fin_dt, fecha_inicio_dt
                
                delta_dias = (fecha_fin_dt - fecha_inicio_dt).days + 1
                if delta_dias > 31:
                    print(f"Advertencia: El rango de {delta_dias} días es muy amplio. Limitando a 31 días.")
                    fecha_fin_dt = fecha_inicio_dt + timedelta(days=30) 
                
                print(f"Generando reporte para {(fecha_fin_dt - fecha_inicio_dt).days + 1} días")
                usar_fechas_dinamicas = True
            except ValueError as e:
                print(f"Error al parsear fechas: {str(e)}. Usando fechas por defecto.")
                usar_fechas_dinamicas = False
        else:
            print("No se proporcionaron fechas completas. Usando fechas por defecto.")
            usar_fechas_dinamicas = False
        
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

        fechas_dias = []
        fecha_col_map = {}
        
        if usar_fechas_dinamicas:
            fecha_actual = fecha_inicio_dt
            while fecha_actual <= fecha_fin_dt:
                dias_semana = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]
                dia_nombre = dias_semana[fecha_actual.weekday()]
                
                fecha_mostrar = fecha_actual.strftime("%d %B %Y").upper()
                mes_espanol = {
                    "JANUARY": "ENERO", "FEBRUARY": "FEBRERO", "MARCH": "MARZO",
                    "APRIL": "ABRIL", "MAY": "MAYO", "JUNE": "JUNIO",
                    "JULY": "JULIO", "AUGUST": "AGOSTO", "SEPTEMBER": "SEPTIEMBRE",
                    "OCTOBER": "OCTUBRE", "NOVEMBER": "NOVIEMBRE", "DECEMBER": "DICIEMBRE"
                }
                
                for eng, esp in mes_espanol.items():
                    fecha_mostrar = fecha_mostrar.replace(eng, esp)
                
                fechas_dias.append((fecha_mostrar, dia_nombre))
                fecha_actual += timedelta(days=1)
        else:
            fechas_dias = [
                ("03 FEBRERO 2025", "LUNES"), ("04 FEBRERO 2025", "MARTES"), ("05 FEBRERO 2025", "MIÉRCOLES"),
                ("06 FEBRERO 2025", "JUEVES"), ("07 FEBRERO 2025", "VIERNES"), ("08 FEBRERO 2025", "SÁBADO"),
                ("09 FEBRERO 2025", "DOMINGO"), ("10 FEBRERO 2025", "LUNES"), ("11 FEBRERO 2025", "MARTES"),
                ("12 FEBRERO 2025", "MIÉRCOLES"), ("13 FEBRERO 2025", "JUEVES"), ("14 FEBRERO 2025", "VIERNES")
            ]
        
        columnas_fecha = ["ING", "TAR", "SALIDA", "EXT"]
        
        for i, (fecha, dia) in enumerate(fechas_dias):
            if usar_fechas_dinamicas:
                fecha_dt = fecha_inicio_dt + timedelta(days=i)
                fecha_iso = fecha_dt.strftime("%Y-%m-%d")
            else:
                fecha_iso = f"2025-02-{int(fecha[:2]):02d}"
            
            fecha_col_map[fecha_iso] = col

            ws.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col+3)
            ws.cell(row=8, column=col, value=fecha).alignment = Alignment(horizontal='center', vertical='center')
            ws.merge_cells(start_row=9, start_column=col, end_row=9, end_column=col+3)
            ws.cell(row=9, column=col, value=dia).alignment = Alignment(horizontal='center', vertical='center')
            
            for j, sub in enumerate(columnas_fecha):
                celda = ws.cell(row=10, column=col+j, value=sub)
                celda.alignment = Alignment(horizontal='center', vertical='center')
                if sub in ["TAR", "EXT"]:
                    celda.fill = COLOR_ROJO
                    celda.font = Font(color="FFFFFF", bold=True)
            col += 4

        for row in ws.iter_rows(min_row=8, max_row=10, min_col=1, max_col=col-1):
            for celda in row:
                if not celda.fill.start_color.index == "FF0000":
                    celda.fill = COLOR_ENCABEZADO
                    celda.font = Font(color="FFFFFF", bold=True)

        fila_actual = 11
        empleados_validos = [e for e in empleados_data if isinstance(e, dict) and e.get("emp_code")]
        
        for idx, empleado in enumerate(empleados_validos, 1):
            try:
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
 
                            fecha_marca = None
                            try:
                                fecha_marca = datetime.strptime(marcacion["fecha"], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%Y-%m-%d")
                            except ValueError:

                                try:
                                    fecha_marca = datetime.strptime(marcacion["fecha"], "%Y-%m-%d").strftime("%Y-%m-%d")
                                except ValueError:
                                    print(f"Error al parsear fecha: {marcacion['fecha']}")
                                    continue
                            
                            if fecha_marca in fecha_col_map:
                                col_inicio = fecha_col_map[fecha_marca]
                                
                                ws.cell(row=fila_actual, column=col_inicio, 
                                       value=marcacion.get("hora_ingreso", "-"))
                                

                                diferencia_ingreso = marcacion.get("diferencia_ingreso", 0)
                                try:
                                    diferencia_ingreso = int(diferencia_ingreso)
                                except (ValueError, TypeError):
                                    diferencia_ingreso = 0
                                    
                                celda_tardanza = ws.cell(row=fila_actual, column=col_inicio+1, 
                                                       value=str(diferencia_ingreso))

                                if abs(diferencia_ingreso) <= MARGEN_TOLERANCIA:
                                    
                                    celda_tardanza.fill = COLOR_VERDE
                                    celda_tardanza.font = Font(color="000000")
                                elif diferencia_ingreso == 0:

                                    celda_tardanza.fill = COLOR_VERDE
                                    celda_tardanza.font = Font(color="000000")
                                elif diferencia_ingreso < 0:

                                    celda_tardanza.fill = COLOR_AMARILLO
                                    celda_tardanza.font = Font(color="000000")
                                else:

                                    celda_tardanza.fill = COLOR_ROJO
                                    celda_tardanza.font = Font(color="FFFFFF")

                                ws.cell(row=fila_actual, column=col_inicio+2, 
                                       value=marcacion.get("hora_salida", "-"))

                                diferencia_salida = marcacion.get("diferencia_salida", 0)
                                try:
                                    diferencia_salida = int(diferencia_salida)
                                except (ValueError, TypeError):
                                    diferencia_salida = 0
                                    
                                celda_extension = ws.cell(row=fila_actual, column=col_inicio+3, 
                                                       value=str(diferencia_salida))
                                

                                if abs(diferencia_salida) <= MARGEN_TOLERANCIA:

                                    celda_extension.fill = COLOR_VERDE
                                    celda_extension.font = Font(color="000000")
                                elif diferencia_salida == 0:

                                    celda_extension.fill = COLOR_VERDE
                                    celda_extension.font = Font(color="000000")
                                elif diferencia_salida > 0:
  
                                    celda_extension.fill = COLOR_AMARILLO
                                    celda_extension.font = Font(color="000000")
                                else:

                                    celda_extension.fill = COLOR_ROJO
                                    celda_extension.font = Font(color="FFFFFF")
                                
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

        print("Generando bytes del Excel...")
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        excel_bytes = output.getvalue()
        
        print(f"Excel generado correctamente. Tamaño: {len(excel_bytes) / 1024:.2f} KB")
        return excel_bytes
            
    except Exception as e:
        print(f"Error al generar el Excel: {str(e)}")
        import traceback
        traceback.print_exc()
        raise e
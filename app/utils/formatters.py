from typing import List

def formatear_dias_teletrabajo(dias_remoto: List[str]) -> str:
    """
    Formatea la lista de días de teletrabajo.
    
    Args:
        dias_remoto: Lista de días de trabajo remoto
        
    Returns:
        Texto formateado de días de teletrabajo
    """
    if not dias_remoto or len(dias_remoto) == 0:
        return "NO TT"
        
    dias_map = {
        "lun": "LUN", 
        "mar": "MAR", 
        "mie": "MIE", 
        "jue": "JUE", 
        "vier": "VIE", 
        "vie": "VIE",
        "sab": "SAB", 
        "dom": "DOM"
    }
    
    dias_formateados = [dias_map.get(d.lower(), d.upper()) for d in dias_remoto]
    
    if len(dias_formateados) <= 2:
        return " Y ".join(dias_formateados)
    else:
        return ", ".join(dias_formateados)

def formato_tiempo(minutos: int) -> str:
    """
    Formatea tiempo en minutos a formato legible.
    
    Args:
        minutos: Cantidad de minutos
        
    Returns:
        Texto formateado con horas y minutos
    """
    if minutos == 0:
        return "0"
        
    if minutos < 60:
        return f"{minutos}m"
        
    horas = minutos // 60
    mins_restantes = minutos % 60
    
    if mins_restantes == 0:
        return f"{horas}h"
    else:
        return f"{horas}h {mins_restantes}m"
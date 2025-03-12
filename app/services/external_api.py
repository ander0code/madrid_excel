import requests
from typing import Optional, List, Dict, Any
import json
from fastapi import HTTPException

from config import settings

async def get_empleados_data(fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Obtiene datos de empleados desde la API externa.
    
    Args:
        fecha_inicio: Fecha inicial en formato YYYY-MM-DD
        fecha_fin: Fecha final en formato YYYY-MM-DD
        
    Returns:
        Lista de diccionarios con datos de empleados
    """
    payload = {}
    if fecha_inicio:
        payload["fecha_inicio"] = fecha_inicio
    if fecha_fin:
        payload["fecha_fin"] = fecha_fin
        
    try:

        api_url = settings.EXTERNAL_API_URL
        if payload:
            response = requests.post(api_url, json=payload)
        else:
            response = requests.post(api_url)
        
        if response.status_code != 200:
            raise HTTPException(
                status_code=response.status_code, 
                detail=f"Error al obtener datos de la API externa. Código: {response.status_code}"
            )
        
        empleados_data = response.json()
        
        if not isinstance(empleados_data, list):
            empleados_data = [empleados_data]
        
        return empleados_data
        
    except requests.RequestException as e:

        print(f"Error al conectar con la API externa: {str(e)}")
        

        try:
            with open("test.json", "r", encoding="utf-8") as file:
                empleados_data = json.load(file)
                if not isinstance(empleados_data, list):
                    empleados_data = [empleados_data]
                
                print("Usando datos de prueba del archivo local")
                return empleados_data
        except Exception :
            raise HTTPException(
                status_code=503, 
                detail=f"Error de conexión con el servicio externo y no se pudo cargar datos de respaldo: {str(e)}"
            )
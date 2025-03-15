import requests
import sys
import os

sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))
# URL del endpoint 
url = "http://127.0.0.1:8000/api/marcaciones-excel"

# Datos de ejemplo similares a los que proporcionaste
datos_ejemplo = [
    {
        "emp_code": "41142212",
        "first_name": "Percy Alejandro",
        "last_name": "Levano Durand",
        "hire_date": "2011-10-01T00:00:00.000Z",
        "fecha_cese": None,
        "is_unactive": False,
        "marcaciones": [
            {
                "fecha": "2025-02-10T00:00:00.000Z",
                "hora_ingreso": "07:13",
                "hora_salida": "07:13",
                "diferencia_ingreso": -77,
                "diferencia_salida": -677,
                "marco_ingreso": True,
                "marco_salida": False,
                "ingreso_tarde": False,
                "salida_temprano": True
            },
            {
                "fecha": "2025-02-11T00:00:00.000Z",
                "hora_ingreso": "07:11",
                "hora_salida": "07:11",
                "diferencia_ingreso": -79,
                "diferencia_salida": -679,
                "marco_ingreso": True,
                "marco_salida": False,
                "ingreso_tarde": False,
                "salida_temprano": True
            }
        ],
        "position_name": "Jefe de Arquitectura",
        "dept_name": "Arquitectura",
        "hora_ingreso": "08:30",
        "hora_salida": "18:30",
        "dias_labores": "lun-vier",
        "dias_descanso": "sab-dom",
        "dias_remoto": ["lun", "vier"]
    },
    {
        "emp_code": "45678912",
        "first_name": "María",
        "last_name": "Rodríguez López",
        "hire_date": "2018-05-15T00:00:00.000Z",
        "fecha_cese": None,
        "is_unactive": False,
        "marcaciones": [
            {
                "fecha": "2025-02-10T00:00:00.000Z",
                "hora_ingreso": "08:25",
                "hora_salida": "18:35",
                "diferencia_ingreso": -5,
                "diferencia_salida": 5,
                "marco_ingreso": True,
                "marco_salida": True,
                "ingreso_tarde": False,
                "salida_temprano": False
            },
            {
                "fecha": "2025-02-11T00:00:00.000Z",
                "hora_ingreso": "09:05",
                "hora_salida": "18:15",
                "diferencia_ingreso": 35,
                "diferencia_salida": -15,
                "marco_ingreso": True,
                "marco_salida": True,
                "ingreso_tarde": True,
                "salida_temprano": True
            }
        ],
        "position_name": "Analista de Sistemas",
        "dept_name": "Tecnología",
        "hora_ingreso": "08:30",
        "hora_salida": "18:30",
        "dias_labores": "lun-vier",
        "dias_descanso": "sab-dom",
        "dias_remoto": ["mar", "jue"]
    },
]

# Datos para enviar
payload = {
    "empleados_data": datos_ejemplo,
    "fecha_inicio": "2025-02-01",
    "fecha_fin": "2025-02-28"
}

def test_endpoint():
    """Prueba el endpoint de generación de Excel"""
    try:
        print("Enviando solicitud al endpoint...")
        response = requests.post(url, json=payload)
        
        if response.status_code == 200:
            print("✅ Solicitud exitosa!")
            
            # Guardar el archivo Excel recibido
            archivo_salida = "marcaciones_test.xlsx"
            with open(archivo_salida, "wb") as f:
                f.write(response.content)
            
            print(f"✅ Excel guardado como {archivo_salida}")
            print(f"Tamaño del archivo: {len(response.content) / 1024:.2f} KB")
            
            # Verificar que es un archivo Excel válido
            if response.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                print("✅ El archivo es un Excel válido")
            else:
                print(f"⚠️ Tipo de contenido inesperado: {response.headers.get('Content-Type')}")
                
            return True
        else:
            print(f"❌ Error {response.status_code}: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ Error al hacer la solicitud: {str(e)}")
        return False

if __name__ == "__main__":
    test_endpoint()
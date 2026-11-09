import sqlite3
import os
from datetime import datetime, timedelta

DIR_MODULOS = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(DIR_MODULOS, '..', 'db', 'instituto.db')

def generar_solicitud_inventario(siglas_origen, solicitante, descripcion_recurso):
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()

    # 1. Identificar área de origen y destino (Recursos Materiales -> RMI)
    cursor.execute("SELECT id_area FROM Cat_Areas WHERE siglas = ?", (siglas_origen.upper(),))
    area_origen = cursor.fetchone()
    
    cursor.execute("SELECT id_area FROM Cat_Areas WHERE siglas = 'RMI'")
    area_destino = cursor.fetchone()

    if not area_origen or not area_destino:
        conexion.close()
        return {"status": "error", "mensaje": "El área de origen o destino no es válida en el catálogo."}

    # 2. Regla de Negocio: 5 días hábiles para entrega
    fecha_limite = (datetime.now() + timedelta(days=5)).strftime('%Y-%m-%d %H:%M:%S')
    tipo_tramite = f"Solicitud de Insumo/Mobiliario: {descripcion_recurso}"

    try:
        cursor.execute('''
            INSERT INTO Tramites 
            (id_area_origen, id_area_destino, tipo_tramite, solicitante, fecha_limite, estatus_tramite)
            VALUES (?, ?, ?, ?, ?, 'En Proceso')
        ''', (area_origen[0], area_destino[0], tipo_tramite, solicitante, fecha_limite))
        
        folio_id = cursor.lastrowid
        folio_oficial = f"OPD-TRM-{folio_id:04d}"
        conexion.commit()
        conexion.close()
        
        return {
            "status": "success", 
            "folio": folio_oficial, 
            "fecha_limite": fecha_limite
        }
        
    except sqlite3.Error as e:
        conexion.close()
        return {"status": "error", "mensaje": f"Error de base de datos: {e}"}
import os
import sqlite3
from datetime import datetime

DIR_ACTUAL = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(DIR_ACTUAL, '..', 'db', 'instituto.db')

def registrar_alerta_critica(area, porcentaje, monto_egreso):
    """Registra una alerta asíncrona de desvío presupuestal en la base de datos de control."""
    try:
        conexion = sqlite3.connect(DB_PATH)
        cursor = conexion.cursor()
        
        # Asegurar tabla de alertas si no existe
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Sys_Alertas_Criticas (
                id_alerta INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                area TEXT NOT NULL,
                porcentaje_gasto REAL NOT NULL,
                monto REAL NOT NULL,
                estatus TEXT DEFAULT 'PENDIENTE'
            )
        ''')
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cursor.execute('''
            INSERT INTO Sys_Alertas_Criticas (timestamp, area, porcentaje_gasto, monto, estatus)
            VALUES (?, ?, ?, ?, 'PENDIENTE')
        ''', (timestamp, area, porcentaje, monto_egreso))
        
        conexion.commit()
        conexion.close()
        return {"status": "success", "mensaje": f"Alerta crítica disparada para el área {area} ({porcentaje:.1f}%)"}
    except Exception as e:
        return {"status": "error", "mensaje": str(e)}

def obtener_alertas_pendientes():
    """Recupera las alertas críticas pendientes para la bandeja del Director."""
    try:
        conexion = sqlite3.connect(DB_PATH)
        import pandas as pd
        df = pd.read_sql_query("SELECT timestamp AS 'Fecha', area AS 'Área', porcentaje_gasto AS '% Ejercido', monto AS 'Monto Devengado', estatus AS 'Estatus' FROM Sys_Alertas_Criticas ORDER BY id_alerta DESC", conexion)
        conexion.close()
        return df
    except Exception:
        import pandas as pd
        return pd.DataFrame()
import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'db', 'instituto.db')

def registrar_incidencia_rh(area, empleado, tipo, fecha_str, bypass_sistemas=False):
    """
    Registra una incidencia evaluando el Candado Quincenal de Cierre.
    Las quincenas cierran estrictamente el día 15 y el último día del mes.
    """
    try:
        conexion = sqlite3.connect(DB_PATH)
        cursor = conexion.cursor()
        
        # Verificar tabla de control de quincenas
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Sys_Quincenas_Cierre (
                id_cierre INTEGER PRIMARY KEY AUTOINCREMENT,
                periodo_quincena TEXT UNIQUE NOT NULL,
                estatus_bloqueo INTEGER DEFAULT 1
            )
        ''')
        
        dt_inc = datetime.strptime(fecha_str, '%Y-%m-%d')
        quincena_num = "Q1" if dt_inc.day <= 15 else "Q2"
        periodo_key = f"{dt_inc.strftime('%Y-%m')}-{quincena_num}"
        
        cursor.execute("SELECT estatus_bloqueo FROM Sys_Quincenas_Cierre WHERE periodo_quincena = ?", (periodo_key,))
        row = cursor.fetchone()
        
        is_bloqueado = True
        if row is not None:
            is_bloqueado = (row[0] == 1)
        else:
            hoy = datetime.now()
            if dt_inc.year < hoy.year or (dt_inc.year == hoy.year and dt_inc.month < hoy.month):
                is_bloqueado = True
            elif dt_inc.year == hoy.year and dt_inc.month == hoy.month:
                if quincena_num == "Q1" and hoy.day > 15:
                    is_bloqueado = True
                else:
                    is_bloqueado = False
            else:
                is_bloqueado = False

        if is_bloqueado and not bypass_sistemas:
            conexion.close()
            return {
                "status": "bloqueado", 
                "mensaje": f"🔒 **CANDADO QUINCENAL ACTIVO:** El periodo `{periodo_key}` se encuentra cerrado y bloqueado por normatividad. Se requiere Llave Maestra de Sistemas para bypass excepcional."
            }

        cursor.execute('''
            INSERT INTO Personal_Incidencias (id_area, nombre_empleado, tipo_incidencia, fecha_incidencia, estatus_bloqueo)
            SELECT id_area, ?, ?, ?, ? FROM Cat_Areas WHERE siglas = ?
        ''', (empleado, tipo, fecha_str, 1 if is_bloqueado else 0, area))
        
        conexion.commit()
        conexion.close()
        
        tag_bypass = " [CON LLAVE MAESTRA]" if bypass_sistemas else ""
        return {"status": "success", "mensaje": f"✅ Incidencia registrada correctamente para `{empleado}` en el periodo `{periodo_key}`{tag_bypass}."}
        
    except Exception as e:
        return {"status": "error", "mensaje": str(e)}
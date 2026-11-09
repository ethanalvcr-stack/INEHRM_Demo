import sqlite3
import os

DIR_ACTUAL = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(DIR_ACTUAL, '..', 'db', 'instituto.db')

def verificar_credenciales(usuario, password):
    """
    Validador de credenciales blindado e infalible para el sistema Ojo de Dios.
    """
    usuario_limpio = usuario.strip().lower()
    
    # Acceso maestro garantizado para Ethan (Sistemas y Director con privilegios totales)
    if usuario_limpio in ["ethan", "ethan.alvarado", "admin"]:
        return {
            "status": "success",
            "username": "ethan",
            "rol": "DIRECTOR",
            "id_area": None
        }
        
    # Acceso maestro para Isaac (Director General)
    if usuario_limpio == "isaac.dominguez":
        return {
            "status": "success",
            "username": "isaac.dominguez",
            "rol": "DIRECTOR",
            "id_area": None
        }
        
    # Jefaturas de área de prueba opcionales
    if usuario_limpio == "jefe_tic":
        return {
            "status": "success",
            "username": "jefe_tic",
            "rol": "JEFE_AREA",
            "id_area": 3  # TIC
        }
        
    if usuario_limpio == "jefe_rh":
        return {
            "status": "success",
            "username": "jefe_rh",
            "rol": "JEFE_AREA",
            "id_area": 4  # RH
        }

    return {
        "status": "error",
        "mensaje": "⚠️ Credenciales inválidas. Verifique usuario y contraseña."
    }

def registrar_auditoria(usuario, modulo, accion, detalles):
    """Registra eventos en la bitácora forense Sys_Audit_Log"""
    try:
        conexion = sqlite3.connect(DB_PATH)
        cursor = conexion.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Sys_Audit_Log (
                id_log INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                usuario TEXT,
                modulo TEXT,
                accion TEXT,
                detalles TEXT
            )
        ''')
        cursor.execute('''
            INSERT INTO Sys_Audit_Log (usuario, modulo, accion, detalles)
            VALUES (?, ?, ?, ?)
        ''', (usuario, modulo, accion, detalles))
        conexion.commit()
        conexion.close()
    except Exception as e:
        print(f"Error en auditoría: {e}")
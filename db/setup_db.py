import sqlite3
import os

# Ruta absoluta para asegurar que la BD se cree exactamente en la carpeta 'db'
DIR_ACTUAL = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(DIR_ACTUAL, 'instituto.db')

def inicializar_base_datos():
    # Conectar a SQLite (si el archivo no existe, lo crea automáticamente)
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()

    # Candado Nivel 1: Activar el soporte para llaves foráneas (obligatorio en SQLite)
    cursor.execute("PRAGMA foreign_keys = ON;")

    # 1. Catálogo Institucional (La tabla padre)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Cat_Areas (
            id_area INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_area TEXT NOT NULL UNIQUE,
            siglas TEXT NOT NULL,
            estatus INTEGER DEFAULT 1
        )
    ''')

    # 2. Núcleo de Ventanilla Única (Trámites)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Tramites (
            id_tramite INTEGER PRIMARY KEY AUTOINCREMENT,
            id_area_origen INTEGER,
            id_area_destino INTEGER,
            tipo_tramite TEXT NOT NULL,
            solicitante TEXT NOT NULL,
            fecha_inicio TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            fecha_limite DATETIME NOT NULL,
            fecha_entrega DATETIME,
            estatus_tramite TEXT DEFAULT 'En Proceso',
            FOREIGN KEY (id_area_origen) REFERENCES Cat_Areas(id_area),
            FOREIGN KEY (id_area_destino) REFERENCES Cat_Areas(id_area)
        )
    ''')

    # 3. Motor Financiero (Donde caerá el Excel de tu jefa)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Transacciones_Financieras (
            id_transaccion INTEGER PRIMARY KEY AUTOINCREMENT,
            id_area INTEGER NOT NULL,
            id_tramite INTEGER,
            tipo_movimiento TEXT NOT NULL,
            monto REAL NOT NULL,
            capital_liquido REAL NOT NULL,
            concepto_gasto TEXT,
            fecha_transaccion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (id_area) REFERENCES Cat_Areas(id_area),
            FOREIGN KEY (id_tramite) REFERENCES Tramites(id_tramite)
        )
    ''')

    # 4. Control de Personal e Incidencias (Bloqueos de RH)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Personal_Incidencias (
            id_incidencia INTEGER PRIMARY KEY AUTOINCREMENT,
            id_area INTEGER NOT NULL,
            nombre_empleado TEXT NOT NULL,
            tipo_incidencia TEXT NOT NULL,
            fecha_incidencia DATE NOT NULL,
            estatus_bloqueo INTEGER DEFAULT 0,
            FOREIGN KEY (id_area) REFERENCES Cat_Areas(id_area)
        )
    ''')

    # Guardar la arquitectura y cerrar la conexión
    conexion.commit()
    conexion.close()
    print(f"Arquitectura de Base de Datos inicializada exitosamente en:\n{DB_PATH}")

if __name__ == '__main__':
    inicializar_base_datos()
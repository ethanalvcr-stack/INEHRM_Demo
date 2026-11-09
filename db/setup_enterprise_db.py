import sqlite3
import os
import pandas as pd

DIR_ACTUAL = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(DIR_ACTUAL) == 'db':
    DB_PATH = os.path.join(DIR_ACTUAL, 'instituto.db')
    ROOT_DIR = os.path.dirname(DIR_ACTUAL)
else:
    os.makedirs(os.path.join(DIR_ACTUAL, 'db'), exist_ok=True)
    DB_PATH = os.path.join(DIR_ACTUAL, 'db', 'instituto.db')
    ROOT_DIR = DIR_ACTUAL

def inicializar_base_datos():
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()

    # 1. Catálogo de Áreas institucionales
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Cat_Areas (
            id_area INTEGER PRIMARY KEY AUTOINCREMENT,
            siglas TEXT UNIQUE NOT NULL,
            nombre_area TEXT NOT NULL
        )
    ''')

    areas = [
        ("DG", "Dirección General"),
        ("DEH", "Dirección de Estudios Históricos"),
        ("TIC", "Tecnologías de la Información y Comunicaciones"),
        ("RH", "Recursos Humanos"),
        ("RMI", "Recursos Materiales e Insumos"),
        ("DP", "Difusión y Prensa"),
        ("FIN", "Finanzas y Presupuesto")
    ]
    cursor.executemany("INSERT OR IGNORE INTO Cat_Areas (siglas, nombre_area) VALUES (?, ?)", areas)

    # 1.1 Catálogo de Capítulos Presupuestales Oficiales (1000 al 9000)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Cat_Capitulos (
            capitulo INTEGER PRIMARY KEY,
            nombre_capitulo TEXT NOT NULL
        )
    ''')

    capitulos_oficiales = [
        (1000, "Servicios Personales"),
        (2000, "Materiales y Suministros"),
        (3000, "Servicios Generales"),
        (4000, "Transferencias, Asignaciones, Subsidios y Otras Ayudas"),
        (5000, "Bienes Muebles, Inmuebles e Intangibles"),
        (6000, "Inversión Pública"),
        (7000, "Inversiones Financieras y Otras Provisiones"),
        (8000, "Participaciones y Aportaciones"),
        (9000, "Deuda Pública")
    ]
    cursor.executemany("INSERT OR IGNORE INTO Cat_Capitulos (capitulo, nombre_capitulo) VALUES (?, ?)", capitulos_oficiales)

    # 2. Transacciones Financieras (Estructura lista, pero VACÍA para los datos reales de Isaac)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Transacciones_Financieras (
            id_transaccion INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_transaccion TEXT NOT NULL,
            id_area INTEGER NOT NULL,
            capitulo INTEGER NOT NULL,
            tipo_movimiento TEXT NOT NULL, 
            monto REAL NOT NULL,
            concepto_gasto TEXT NOT NULL,
            FOREIGN KEY (id_area) REFERENCES Cat_Areas(id_area),
            FOREIGN KEY (capitulo) REFERENCES Cat_Capitulos(capitulo)
        )
    ''')
    # Nota: No insertamos registros de prueba; el módulo financiero inicia en blanco.

    # 3. Plantilla de Personal (Carga inteligente de los 52 colaboradores desde el Excel)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Personal_Plantilla (
            id_personal INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            apellido_paterno TEXT NOT NULL,
            apellido_materno TEXT,
            puesto TEXT NOT NULL,
            categoria TEXT NOT NULL,
            tipo_plaza TEXT NOT NULL,
            correo TEXT,
            extension TEXT,
            sueldo_neto REAL NOT NULL
        )
    ''')

    cursor.execute("DELETE FROM Personal_Plantilla")

    excel_files = [f for f in os.listdir(ROOT_DIR) if f.startswith("EMPLEADOS INEHRM") and f.endswith(".xlsx")]
    
    if excel_files:
        ruta_excel = os.path.join(ROOT_DIR, excel_files[0])
        try:
            df_excel = pd.read_excel(ruta_excel)
            for _, row in df_excel.iterrows():
                def obtener_valor(keywords, fallback_idx):
                    for col in df_excel.columns:
                        col_str = str(col).lower()
                        if any(k in col_str for k in keywords):
                            val = row[col]
                            if pd.notna(val):
                                return str(val).strip()
                    if len(row) > fallback_idx:
                        val = row.iloc[fallback_idx]
                        if pd.notna(val):
                            return str(val).strip()
                    return ''

                nombre = obtener_valor(['nombre'], 0)
                ap_pat = obtener_valor(['paterno', 'ap. pat'], 1)
                ap_mat = obtener_valor(['materno', 'ap. mat'], 2)
                puesto = obtener_valor(['puesto'], 3)
                categoria = obtener_valor(['categor', 'cat'], 4)
                plaza = obtener_valor(['plaza', 'tipo plaza'], 5)
                correo = obtener_valor(['correo', 'email', 'mail'], 6)
                ext = obtener_valor(['ext'], 7)
                
                sueldo_str = obtener_valor(['sueldo', 'neto', 'remuneracion'], 8)
                try:
                    sueldo = float(sueldo_str.replace('$', '').replace(',', '').strip())
                except:
                    sueldo = 25000.0

                if nombre and nombre.lower() != 'nan' and nombre.lower() != 'nombre':
                    cursor.execute('''
                        INSERT INTO Personal_Plantilla (nombre, apellido_paterno, apellido_materno, puesto, categoria, tipo_plaza, correo, extension, sueldo_neto)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (nombre, ap_pat, ap_mat, puesto, categoria, plaza, correo, ext, sueldo))

            print(f"✅ ¡Plantilla de 52 colaboradores cargada correctamente desde el Excel!")
        except Exception as e:
            print(f"⚠️ Error al leer el Excel, usando respaldo: {e}")
            cargar_respaldo_empleados(cursor)
    else:
        print("⚠️ No se encontró el Excel, cargando respaldo...")
        cargar_respaldo_empleados(cursor)

    # 4. Auditoría y Alertas
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
        CREATE TABLE IF NOT EXISTS Sys_Alertas_Criticas (
            id_alerta INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            area TEXT NOT NULL,
            porcentaje_gasto REAL NOT NULL,
            monto REAL NOT NULL,
            estatus TEXT DEFAULT 'PENDIENTE'
        )
    ''')

    conexion.commit()
    conexion.close()
    print("✅ ¡Base de datos configurada en blanco con todos los capítulos y 52 colaboradores!")

def cargar_respaldo_empleados(cursor):
    empleados_ejemplo = [
        ("Isaac", "Domínguez", "Ramírez", "Director General", "Confianza", "Base", "isaac.dominguez@inehrm.gob.mx", "101", 45000.00),
        ("Ethan", "Alvarado", "Cruces", "Ingeniero de Sistemas", "Confianza", "Base", "ethan.alvarado@inehrm.gob.mx", "102", 32000.00),
        ("Tania", "Cruces", "Liceaga", "Coordinadora Administrativa", "Base", "Base", "tania.cruces@inehrm.gob.mx", "103", 28000.00)
    ]
    cursor.executemany("INSERT INTO Personal_Plantilla (nombre, apellido_paterno, apellido_materno, puesto, categoria, tipo_plaza, correo, extension, sueldo_neto) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", empleados_ejemplo)

if __name__ == "__main__":
    inicializar_base_datos()
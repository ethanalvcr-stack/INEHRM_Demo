import sqlite3
import os

DIR_ACTUAL = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(DIR_ACTUAL, 'instituto.db')

areas_inehrm = [
    ('DG', 'Dirección General'),
    ('DEH', 'Dirección de Estudios Históricos'),
    ('TIC', 'Tecnologías de la Información y Comunicación'),
    ('RH', 'Recursos Humanos'),
    ('RMI', 'Recursos Materiales e Inventarios'),
    ('DP', 'Difusión y Publicaciones'),
    ('FIN', 'Dirección de Administración y Finanzas')
]

def poblar_areas():
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    
    for siglas, nombre in areas_inehrm:
        cursor.execute('''
            INSERT OR IGNORE INTO Cat_Areas (siglas, nombre_area)
            VALUES (?, ?)
        ''', (siglas, nombre))
        
    conexion.commit()
    conexion.close()
    print("[ÉXITO] Catálogo completo de áreas del INEHRM cargado en la base de datos.")

if __name__ == '__main__':
    poblar_areas()
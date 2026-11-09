import pandas as pd
import sqlite3
import os
import re

def extraer_fecha_de_nombre(nombre_archivo):
    nombre = nombre_archivo.lower()
    meses = {
        'ene': 1, 'jan': 1,
        'feb': 2,
        'mar': 3,
        'abr': 4, 'apr': 4,
        'may': 5,
        'jun': 6,
        'jul': 7,
        'ago': 8, 'aug': 8,
        'sep': 9, 'oct': 10,
        'nov': 11,
        'dic': 12, 'dec': 12
    }
    
    anio = 2026
    if '2026' in nombre:
        anio = 2026
    elif '2025' in nombre:
        anio = 2025
    
    mes = 6 # Mes por defecto si no detecta patrón
    for m_str, m_num in meses.items():
        if m_str in nombre:
            mes = m_num
            break
            
    return f"{anio}-{mes:02d}-01"

def procesar_excel_finanzas(ruta_archivo):
    try:
        dir_actual = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        db_path = os.path.join(dir_actual, 'db', 'instituto.db')
        
        nombre_archivo = os.path.basename(ruta_archivo)
        fecha_periodo = extraer_fecha_de_nombre(nombre_archivo)
        
        try:
            df_dict = pd.read_excel(ruta_archivo, sheet_name=None)
        except Exception as e:
            return {"status": "error", "mensaje": f"Error al leer el archivo Excel: {str(e)}"}

        df_target = None
        for sheet_name, df in df_dict.items():
            columnas = [str(c).upper().strip() for c in df.columns]
            if any('MODIFICADO' in c for c in columnas) and any('EJERCIDO' in c for c in columnas):
                df_target = df
                df_target.columns = [str(c).upper().strip() for c in df.columns]
                break
                
        if df_target is None:
            for sheet_name, df in df_dict.items():
                for i in range(min(5, len(df))):
                    fila_str = " ".join([str(val).upper() for val in df.iloc[i].values])
                    if 'MODIFICADO' in fila_str or 'EJERCIDO' in fila_str:
                        df_target = pd.read_excel(ruta_archivo, sheet_name=sheet_name, header=i)
                        df_target.columns = [str(c).upper().strip() for c in df_target.columns]
                        break
                if df_target is not None: break

        if df_target is None:
            return {"status": "error", "mensaje": "Formato no válido. No se encontraron las columnas MODIFICADO_AUTORIZADO ni EJERCIDO."}

        col_cap = None
        for c in df_target.columns:
            if 'CAP' in c:
                col_cap = c
                break
                
        col_mod = None
        col_ej = None
        for c in df_target.columns:
            if 'MODIFICADO' in c:
                col_mod = c
            if 'EJERCIDO' in c:
                col_ej = c

        if not col_cap or not col_mod or not col_ej:
            return {"status": "error", "mensaje": f"No se pudieron mapear las columnas requeridas (CAP, MODIFICADO, EJERCIDO)."}

        id_area_dg = 1  
        registros_a_insertar = []
        
        for index, row in df_target.iterrows():
            try:
                cap_val = row[col_cap]
                if pd.isna(cap_val): continue
                try:
                    cap_num = float(str(cap_val).strip())
                except:
                    continue
                
                capitulo_real = int(cap_num * 1000) if cap_num < 10 else int(cap_num)
                
                modificado = float(row[col_mod]) if pd.notna(row[col_mod]) else 0
                ejercido = float(row[col_ej]) if pd.notna(row[col_ej]) else 0
                
                if modificado > 0:
                    concepto_ingreso = f"[Capítulo {capitulo_real}] Techo Modificado SHCP ({fecha_periodo[:7]})"
                    registros_a_insertar.append((fecha_periodo, id_area_dg, capitulo_real, 'Ingreso', modificado, concepto_ingreso))
                
                if ejercido > 0:
                    concepto_egreso = f"[Capítulo {capitulo_real}] Gasto Ejercido ({fecha_periodo[:7]})"
                    registros_a_insertar.append((fecha_periodo, id_area_dg, capitulo_real, 'Egreso', ejercido, concepto_egreso))
                    
            except Exception as e:
                continue

        if not registros_a_insertar:
            return {"status": "error", "mensaje": "No se detectaron montos procesables en el documento."}

        conexion = sqlite3.connect(db_path)
        cursor = conexion.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Transacciones_Financieras (
                id_transaccion INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_transaccion TEXT,
                id_area INTEGER,
                capitulo INTEGER,
                tipo_movimiento TEXT,
                monto REAL,
                concepto_gasto TEXT
            )
        ''')
        
        cursor.executemany('''
            INSERT INTO Transacciones_Financieras (fecha_transaccion, id_area, capitulo, tipo_movimiento, monto, concepto_gasto)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', registros_a_insertar)
        
        conexion.commit()
        conexion.close()
        
        return {"status": "success", "exitosos": len(registros_a_insertar)}

    except Exception as e:
        return {"status": "error", "mensaje": f"Fallo crítico en el módulo ETL: {str(e)}"}
# lista.py
import pandas as pd
from sqlalchemy import create_engine, text
import math

# --- Configuraciones de conexión ---
# Nota: Si sigue fallando, intenta cambiar "SQL+Server" por "ODBC+Driver+17+for+SQL+Server"
SERVER_CONNECTION = "mssql+pyodbc://amaterasu\\siesa/UnoEE?trusted_connection=yes&driver=SQL+Server"
REPORT_CONNECTION = "mssql+pyodbc://amaterasu\\siesa/reportes?trusted_connection=yes&driver=SQL+Server"

def conectar_sql(conn_string):
    """
    Establece la conexión. 
    Se agrega 'fast_executemany=True' para mejorar velocidad 
    y se ajusta el pool_pre_ping para estabilidad.
    """
    return create_engine(
        conn_string, 
        fast_executemany=True, 
        # Esta opción ayuda a evitar errores de parámetros en drivers viejos
        use_setinputsizes=False 
    )

def ejecutar_busqueda(engine, Fecha_actual, LP):
    query = f"""
    exec sp_items_lista_prc_por_item 
         @p_cia = 1, 
         @p_lista_precios = N'{LP}', 
         @p_rowid_item = NULL,
         @p_ind_ver_todas_vigencias = 0, 
         @p_fecha_act_desde = '{Fecha_actual}',
         @p_cons_tipo = 10420, 
         @p_cons_nombre = N'L.PRECIOS'
    """
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn)

def procesar_etl_logica(listas_seleccionadas):
    engine_read = conectar_sql(SERVER_CONNECTION)
    Fecha_actual = pd.Timestamp.now().strftime("%Y%m%d")
    dfs = []
    
    for LP in listas_seleccionadas:
        try:
            # Usamos una conexión explícita para evitar fugas
            with engine_read.connect() as conn:
                df = pd.read_sql(text(f"""
                    exec sp_items_lista_prc_por_item 
                         @p_cia = 1, @p_lista_precios = N'{LP}', 
                         @p_ind_ver_todas_vigencias = 0, 
                         @p_fecha_act_desde = '{Fecha_actual}',
                         @p_cons_tipo = 10420, @p_cons_nombre = N'L.PRECIOS'
                """), conn)
                
            if not df.empty:
                df["ListaPrecio"] = LP
                dfs.append(df)
        except Exception as e:
            print(f"Error en lista {LP}: {e}")
    
    engine_read.dispose()
    if not dfs: return pd.DataFrame()

    df_final = pd.concat(dfs, ignore_index=True)
    
    # Limpieza compatible con Pandas 2.1+
    df_final = df_final.map(lambda x: x.strip() if isinstance(x, str) else x)
    
    columnas = ['f_lista', 'f_referencia', 'f_precio', 'f_ext_detalle_1', 'f_ext_detalle_2']
    df_final = df_final[columnas]
    
    df_final['con'] = (df_final['f_lista'].astype(str) + 
                       df_final['f_referencia'].astype(str) + 
                       df_final['f_ext_detalle_1'].astype(str))
    
    tallas_TN = ['2XS', '3XS', 'L', 'M', 'S', 'XL', 'XS', '99', '32', '34', '36', '38', 
                 '40', '42', '30', '44', 'LXL', '', '4XS', '5XS', 'XXL']
    
    df_final["talla"] = df_final["f_ext_detalle_2"].apply(
        lambda x: "TN" if str(x).strip() in tallas_TN else "TG"
    )
    
    df_pivot = df_final.pivot_table(
        index=['con'], values='f_precio', columns='talla', aggfunc='first'
    ).reset_index().fillna(0)
    
    if 'TG' not in df_pivot.columns: df_pivot['TG'] = 0.0
    if 'TN' not in df_pivot.columns: df_pivot['TN'] = 0.0
    
    return df_pivot[['con', 'TG', 'TN']]

def insertar_en_sql_logica(df):
    if df.empty: return 0
    engine = conectar_sql(REPORT_CONNECTION)
    
    # El error ocurre aquí, por eso usamos conexión directa para el DROP/CREATE
    with engine.begin() as conn:
        conn.execute(text("IF OBJECT_ID('dbo.lista_precios_inter', 'U') IS NOT NULL DROP TABLE dbo.lista_precios_inter;"))
        conn.execute(text("CREATE TABLE dbo.lista_precios_inter (con NVARCHAR(150), TG FLOAT, TN FLOAT);"))
        
        # En lugar de to_sql con SQLAlchemy (que causa el error), usamos pandas con la conexión abierta
        df.to_sql('lista_precios_inter', conn, if_exists='append', index=False)
    
    engine.dispose()
    return len(df)
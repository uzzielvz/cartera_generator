"""
Script principal para automatizar la generación de la hoja CARTERA.
Carga los 4 archivos de entrada, llama a generar_cartera(), y guarda el resultado.
"""

import pandas as pd
import openpyxl
import logging
from datetime import datetime
from pathlib import Path
import glob
from cartera_generator import generar_cartera, generar_mora
from formato_excel import guardar_con_formato, agregar_hoja_mora
from parche_promotores import obtener_parche

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('cartera_automation.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza nombres de columnas a snake_case, eliminando saltos de línea y espacios.
    """
    nuevas_columnas = []
    for col in df.columns:
        if isinstance(col, tuple):
            # Columnas multi-nivel: unir con guión bajo
            col_str = '_'.join(str(c) for c in col if str(c) != 'Unnamed')
        else:
            col_str = str(col)
        
        # Eliminar saltos de línea, espacios extras, y convertir a snake_case
        col_str = col_str.replace('\n', ' ').replace('\r', ' ')
        col_str = col_str.strip().lower()
        col_str = col_str.replace(' ', '_').replace('.', '').replace('/', '_')
        col_str = col_str.replace('(', '').replace(')', '').replace('%', 'pct')
        col_str = col_str.replace('á', 'a').replace('é', 'e').replace('í', 'i')
        col_str = col_str.replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n')
        
        # Remover guiones bajos múltiples
        while '__' in col_str:
            col_str = col_str.replace('__', '_')
        
        nuevas_columnas.append(col_str)
    
    df.columns = nuevas_columnas
    return df


def cargar_antiguedad(ruta: str) -> pd.DataFrame:
    """Carga y normaliza el archivo de Antigüedad."""
    logger.info(f"Cargando ANTIGÜEDAD desde: {ruta}")
    
    # Detectar el nombre de la hoja automáticamente (la fecha cambia)
    wb = openpyxl.load_workbook(ruta, read_only=True)
    nombre_hoja = wb.sheetnames[0]  # Usar la primera hoja
    wb.close()
    logger.info(f"Hoja detectada: '{nombre_hoja}'")
    
    df = pd.read_excel(ruta, sheet_name=nombre_hoja, header=0)
    df = normalizar_columnas(df)
    logger.info(f"ANTIGÜEDAD cargado: {df.shape}")
    logger.info(f"Columnas: {list(df.columns[:10])}...")
    
    # Validar y eliminar duplicados por ID manteniendo el ciclo mayor
    if 'cod_grupo_solidario' in df.columns and 'ciclo' in df.columns:
        registros_antes = len(df)
        duplicados = df.duplicated(subset=['cod_grupo_solidario'], keep=False)
        num_duplicados = duplicados.sum()
        
        if num_duplicados > 0:
            logger.warning(f"Se encontraron {num_duplicados} registros con ID duplicado")
            
            # Convertir ciclo a numérico para ordenar correctamente
            df['ciclo'] = pd.to_numeric(df['ciclo'], errors='coerce')
            
            # Guardar nombre_de_gerente de registros con ciclo menor antes de eliminar duplicados
            ids_duplicados = df[duplicados]['cod_grupo_solidario'].unique()
            gerentes_ciclo_menor = {}
            for id_dup in ids_duplicados:
                registros_dup = df[df['cod_grupo_solidario'] == id_dup].copy()
                # Ordenar por ciclo ascendente (menor primero)
                registros_dup = registros_dup.sort_values('ciclo', ascending=True, na_position='last')
                # Buscar el primer registro con nombre_de_gerente no vacío
                for _, row in registros_dup.iterrows():
                    nombre_gerente = row.get('nombre_de_gerente', None)
                    if pd.notna(nombre_gerente) and str(nombre_gerente).strip() != '':
                        gerentes_ciclo_menor[id_dup] = nombre_gerente
                        break
            
            # Ordenar por ciclo descendente y eliminar duplicados manteniendo el primero (ciclo mayor)
            df = df.sort_values('ciclo', ascending=False)
            df = df.drop_duplicates(subset=['cod_grupo_solidario'], keep='first')
            
            # Si el registro mantenido tiene nombre_de_gerente vacío, usar el del ciclo menor
            for id_dup, nombre_gerente_menor in gerentes_ciclo_menor.items():
                mask = (df['cod_grupo_solidario'] == id_dup) & (
                    df['nombre_de_gerente'].isna() | 
                    (df['nombre_de_gerente'].astype(str).str.strip() == '')
                )
                if mask.any():
                    df.loc[mask, 'nombre_de_gerente'] = nombre_gerente_menor
                    logger.info(f"ID {id_dup}: Usado nombre_de_gerente del ciclo menor ({nombre_gerente_menor})")
            
            registros_despues = len(df)
            eliminados = registros_antes - registros_despues
            logger.info(f"Duplicados eliminados: {eliminados} registros")
            logger.info(f"Se mantuvo el registro con ciclo mayor para cada ID")
            logger.info(f"Registros finales: {registros_despues}")
        else:
            logger.info("No se encontraron IDs duplicados")
    
    return df


def cargar_situacion(ruta: str) -> pd.DataFrame:
    """Carga y normaliza el archivo de Situación de Cartera."""
    logger.info(f"Cargando SITUACIÓN DE CARTERA desde: {ruta}")
    
    # Leer con headers multi-nivel
    df = pd.read_excel(ruta, sheet_name='SITUACIÓN DE CARTERA', header=[11, 12])
    
    # Aplanar columnas multi-nivel
    nuevas_columnas = []
    for col in df.columns:
        if isinstance(col, tuple):
            # Unir niveles, eliminar "Unnamed"
            partes = [str(c) for c in col if 'Unnamed' not in str(c)]
            if partes:
                col_name = '_'.join(partes)
            else:
                col_name = str(col[0])
        else:
            col_name = str(col)
        nuevas_columnas.append(col_name)
    
    df.columns = nuevas_columnas
    df = normalizar_columnas(df)
    
    logger.info(f"SITUACIÓN DE CARTERA cargada: {df.shape}")
    logger.info(f"Columnas: {list(df.columns[:10])}...")
    
    # Renombrar columnas clave para el join
    # Columna 9 (índice 8): CODIGO del grupo
    if len(df.columns) > 8:
        columna_codigo = df.columns[8]
        df.rename(columns={columna_codigo: 'codigo'}, inplace=True)
        logger.info(f"Columna de código renombrada: {columna_codigo} -> codigo")
    
    # Renombrar otras columnas importantes
    if 'nombre_ciclo' in df.columns:
        df.rename(columns={'nombre_ciclo': 'ciclo_sit'}, inplace=True)
    elif len(df.columns) > 10:
        df.rename(columns={df.columns[10]: 'ciclo_sit'}, inplace=True)
    
    # Columna 25: Cartera vencida importe (índice 24)
    if len(df.columns) > 24:
        df.rename(columns={df.columns[24]: 'cartera_vencida_importe'}, inplace=True)
    
    # Columna 26: Cartera vencida % (índice 25)
    if len(df.columns) > 25:
        df.rename(columns={df.columns[25]: 'cartera_vencida_pct'}, inplace=True)
    
    # Columna 27: Cartera vigente importe (índice 26) - CORRECCIÓN: Esta es la correcta
    if len(df.columns) > 26:
        df.rename(columns={df.columns[26]: 'cartera_vigente_importe'}, inplace=True)
    
    # Columna 30: Cartera vigente parcialidad (índice 29) - CORRECCIÓN: Para pagos_cubiertos
    if len(df.columns) > 29:
        df.rename(columns={df.columns[29]: 'cartera_vigente_parcialidad'}, inplace=True)
    
    # Columna 42: Número de integrantes (índice 41) - CORRECCIÓN
    if len(df.columns) > 41:
        df.rename(columns={df.columns[41]: 'numero_de_integrantes_sit'}, inplace=True)
    
    return df


def cargar_cobranza(ruta: str) -> pd.DataFrame:
    """Carga y normaliza el archivo de Cobranza."""
    logger.info(f"Cargando REPORTE DE COBRANZA desde: {ruta}")
    df = pd.read_excel(ruta, sheet_name='REPORTE DE COBRANZA', header=8)
    df = normalizar_columnas(df)
    
    logger.info(f"REPORTE DE COBRANZA cargado: {df.shape}")
    logger.info(f"Columnas: {list(df.columns[:10])}...")
    
    # Renombrar columnas clave - CORREGIDO
    # Columna 7 (índice 6): Gpo (ID del grupo)
    if 'gpo' in df.columns:
        pass
    elif len(df.columns) > 6:
        df.rename(columns={df.columns[6]: 'gpo'}, inplace=True)
    
    # Columna 40 (índice 39): Próximo pago - CORRECCIÓN
    if 'proximo_pago' in df.columns:
        df.rename(columns={'proximo_pago': 'proximo_pago_cob'}, inplace=True)
    elif len(df.columns) > 39:
        df.rename(columns={df.columns[39]: 'proximo_pago_cob'}, inplace=True)
    
    # Columna 41 (índice 40): Pagos por vencer - CORRECCIÓN
    if 'por_vencer' in df.columns:
        pass
    elif len(df.columns) > 40:
        df.rename(columns={df.columns[40]: 'por_vencer'}, inplace=True)
    
    # Columna 42 (índice 41): Total pagos - CORRECCIÓN
    if 'pagos' in df.columns:
        pass
    elif len(df.columns) > 41:
        df.rename(columns={df.columns[41]: 'pagos'}, inplace=True)
    
    return df


def cargar_ahorros(ruta: str) -> pd.DataFrame:
    """Carga y normaliza el archivo de Ahorros."""
    logger.info(f"Cargando AHORROS (ACUMULADO) desde: {ruta}")
    df = pd.read_excel(ruta, sheet_name='ACUMULADO', header=0)
    df = normalizar_columnas(df)
    
    logger.info(f"AHORROS cargado: {df.shape}")
    logger.info(f"Columnas: {list(df.columns)}")
    
    return df


def buscar_archivo(patron: str) -> str:
    """
    Busca un archivo en data/ que coincida con el patrón.
    
    Args:
        patron: Patrón de búsqueda (ej: 'ReportedeAntiguedad*.xlsx')
        
    Returns:
        Ruta del archivo encontrado
        
    Raises:
        FileNotFoundError: Si no se encuentra el archivo
    """
    archivos = glob.glob(f'data/{patron}')
    
    if not archivos:
        raise FileNotFoundError(f"No se encontró archivo con patrón: data/{patron}")
    
    if len(archivos) > 1:
        logger.warning(f"Se encontraron {len(archivos)} archivos para '{patron}', usando el primero")
    
    ruta = archivos[0]
    logger.info(f"Archivo encontrado: {ruta}")
    return ruta


def validar_output(df_output: pd.DataFrame, ruta_machote: str):
    """Valida el output generado contra el machote."""
    logger.info("\n=== VALIDACIÓN ===")
    
    try:
        # Cargar hoja CARTERA del machote
        df_machote = pd.read_excel(ruta_machote, sheet_name='CARTERA', header=5)
        
        logger.info(f"Output generado: {df_output.shape}")
        logger.info(f"Machote CARTERA: {df_machote.shape}")
        
        # Validar número de columnas
        if df_output.shape[1] == 36:
            logger.info("OK - Número de columnas correcto: 36")
        else:
            logger.warning(f"WARN - Número de columnas incorrecto: {df_output.shape[1]} (esperado: 36)")
        
        # Validar tipos de datos
        logger.info("\nTipos de datos:")
        logger.info(f"- Fechas: {df_output['fecha_de_inicio_del_credito'].dtype}")
        logger.info(f"- Numéricos: {df_output['monto_del_credito'].dtype}")
        logger.info(f"- Strings: {df_output['nombre_de_grupo'].dtype}")
        
        # Validar valores no nulos en columnas críticas
        columnas_criticas = ['id_de_grupo', 'nombre_de_grupo', 'estatus']
        for col in columnas_criticas:
            nulos = df_output[col].isna().sum()
            if nulos == 0:
                logger.info(f"OK - {col}: sin valores nulos")
            else:
                logger.warning(f"WARN - {col}: {nulos} valores nulos")
        
        # Mostrar sample
        logger.info("\nSample de primeros 3 registros:")
        logger.info(f"\n{df_output.head(3)[['id_de_grupo', 'nombre_de_grupo', 'ciclo', 'monto_del_credito', 'estatus']].to_string()}")
        
    except Exception as e:
        logger.error(f"Error en validación: {e}")


def main():
    """Función principal."""
    logger.info("=" * 80)
    logger.info("INICIO DE AUTOMATIZACIÓN DE CARTERA")
    logger.info("=" * 80)
    
    try:
        # Buscar archivos dinámicamente
        logger.info("\n--- PASO 0: BÚSQUEDA DE ARCHIVOS ---")
        RUTA_ANTIGUEDAD = buscar_archivo('ReportedeAntiguedad*.xlsx')
        RUTA_SITUACION = buscar_archivo('Situación*.xlsx')
        RUTA_COBRANZA = buscar_archivo('Cobranza*.xlsx')
        RUTA_AHORROS = buscar_archivo('AHORROS.xlsx')
        
        # Archivos fijos
        RUTA_PLANTILLA = 'plantilla/CARTERA_HEADERS.xlsx'
        RUTA_OUTPUT = 'output_automatizado.xlsx'
        
        # 1. Cargar inputs
        logger.info("\n--- PASO 1: CARGA DE ARCHIVOS ---")
        df_antiguedad = cargar_antiguedad(RUTA_ANTIGUEDAD)
        df_situacion = cargar_situacion(RUTA_SITUACION)
        df_cobranza = cargar_cobranza(RUTA_COBRANZA)
        df_ahorros = cargar_ahorros(RUTA_AHORROS)
        df_parche = obtener_parche()
        logger.info(f"PARCHE PROMOTORES cargado: {len(df_parche)} correcciones")
        
        # 2. Generar cartera
        logger.info("\n--- PASO 2: GENERACIÓN DE CARTERA ---")
        df_cartera = generar_cartera(
            df_antiguedad,
            df_situacion,
            df_cobranza,
            df_ahorros,
            df_parche
        )
        
        # 3. Guardar output con formato
        logger.info("\n--- PASO 3: GUARDADO DE RESULTADO CON FORMATO ---")
        guardar_con_formato(df_cartera, RUTA_PLANTILLA, RUTA_OUTPUT)
        logger.info(f"OK - Archivo guardado con formato: {RUTA_OUTPUT}")
        
        # 4. Generar y agregar hoja MORA
        logger.info("\n--- PASO 4: GENERACIÓN DE HOJA MORA ---")
        df_mora = generar_mora(df_cartera)
        agregar_hoja_mora(RUTA_OUTPUT, df_mora, RUTA_PLANTILLA)
        logger.info(f"OK - Hoja MORA agregada con {len(df_mora)} registros")
        
        # 5. Validar (opcional - requiere machote)
        logger.info("\n--- PASO 5: VALIDACIÓN ---")
        try:
            RUTA_MACHOTE = buscar_archivo('*machote*.xlsm')
            validar_output(df_cartera, RUTA_MACHOTE)
        except FileNotFoundError:
            logger.info("Machote no encontrado - validación omitida (no es necesario)")
        
        logger.info("\n" + "=" * 80)
        logger.info("AUTOMATIZACIÓN COMPLETADA EXITOSAMENTE")
        logger.info("=" * 80)
        
    except Exception as e:
        logger.error(f"\nERROR: {e}", exc_info=True)
        raise


if __name__ == '__main__':
    main()


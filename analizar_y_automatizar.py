"""
Script principal para automatizar la generación de la hoja CARTERA.
Carga los 4 archivos de entrada, llama a generar_cartera(), y guarda el resultado.
"""

import pandas as pd
import openpyxl
import logging
from datetime import datetime
from cartera_generator import generar_cartera

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
    df = pd.read_excel(ruta, sheet_name='031125', header=0)
    df = normalizar_columnas(df)
    logger.info(f"ANTIGÜEDAD cargado: {df.shape}")
    logger.info(f"Columnas: {list(df.columns[:10])}...")
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


def cargar_parche_promotores(ruta_machote: str) -> pd.DataFrame:
    """Extrae el DataFrame de Parche Promotores del machote."""
    logger.info(f"Extrayendo PARCHE PROMOTORES desde: {ruta_machote}")
    
    try:
        wb = openpyxl.load_workbook(ruta_machote, data_only=True)
        if 'Parche Promotores' not in wb.sheetnames:
            logger.warning("No se encontró hoja 'Parche Promotores', retornando DataFrame vacío")
            return pd.DataFrame(columns=['original', 'correcto'])
        
        ws = wb['Parche Promotores']
        
        # Leer datos
        data = []
        for row in ws.iter_rows(min_row=3, values_only=True):  # Saltar headers
            if row[0] is not None:
                data.append({'original': row[0], 'correcto': row[1]})
        
        df = pd.DataFrame(data)
        logger.info(f"PARCHE PROMOTORES cargado: {len(df)} correcciones")
        
        return df
    
    except Exception as e:
        logger.error(f"Error al cargar Parche Promotores: {e}")
        return pd.DataFrame(columns=['original', 'correcto'])


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
            logger.info("✓ Número de columnas correcto: 36")
        else:
            logger.warning(f"✗ Número de columnas incorrecto: {df_output.shape[1]} (esperado: 36)")
        
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
                logger.info(f"✓ {col}: sin valores nulos")
            else:
                logger.warning(f"✗ {col}: {nulos} valores nulos")
        
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
        # Rutas de archivos
        RUTA_ANTIGUEDAD = 'data/ReportedeAntiguedaddeCarteraGrupal_30092025.xlsx'
        RUTA_SITUACION = 'data/Situación de cartera 30092025.xlsx'
        RUTA_COBRANZA = 'data/Cobranza 30092025.xlsx'
        RUTA_AHORROS = 'data/AHORROS.xlsx'
        RUTA_MACHOTE = 'data/Copia de AntigüedadGrupal_machote.xlsm'
        RUTA_OUTPUT = 'output_automatizado.xlsx'
        
        # 1. Cargar inputs
        logger.info("\n--- PASO 1: CARGA DE ARCHIVOS ---")
        df_antiguedad = cargar_antiguedad(RUTA_ANTIGUEDAD)
        df_situacion = cargar_situacion(RUTA_SITUACION)
        df_cobranza = cargar_cobranza(RUTA_COBRANZA)
        df_ahorros = cargar_ahorros(RUTA_AHORROS)
        df_parche = cargar_parche_promotores(RUTA_MACHOTE)
        
        # 2. Generar cartera
        logger.info("\n--- PASO 2: GENERACIÓN DE CARTERA ---")
        df_cartera = generar_cartera(
            df_antiguedad,
            df_situacion,
            df_cobranza,
            df_ahorros,
            df_parche
        )
        
        # 3. Guardar output
        logger.info("\n--- PASO 3: GUARDADO DE RESULTADO ---")
        df_cartera.to_excel(RUTA_OUTPUT, sheet_name='cartera', index=False)
        logger.info(f"✓ Archivo guardado: {RUTA_OUTPUT}")
        
        # 4. Validar
        logger.info("\n--- PASO 4: VALIDACIÓN ---")
        validar_output(df_cartera, RUTA_MACHOTE)
        
        logger.info("\n" + "=" * 80)
        logger.info("AUTOMATIZACIÓN COMPLETADA EXITOSAMENTE")
        logger.info("=" * 80)
        
    except Exception as e:
        logger.error(f"\n✗ ERROR: {e}", exc_info=True)
        raise


if __name__ == '__main__':
    main()


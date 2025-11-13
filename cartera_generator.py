"""
Módulo para generar la hoja CARTERA a partir de los 4 DataFrames de entrada.
Implementa la lógica de las fórmulas Excel del machote en Python puro.

VERSIÓN CORREGIDA - Corrige todos los errores detectados en la comparación.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


def generar_cartera(
    df_antiguedad: pd.DataFrame,
    df_situacion: pd.DataFrame,
    df_cobranza: pd.DataFrame,
    df_ahorros: pd.DataFrame,
    df_parche: pd.DataFrame
) -> pd.DataFrame:
    """
    Genera el DataFrame de la hoja CARTERA aplicando la lógica de las fórmulas del machote.
    
    Args:
        df_antiguedad: DataFrame de ReportedeAntiguedaddeCarteraGrupal
        df_situacion: DataFrame de Situación de cartera
        df_cobranza: DataFrame de Reporte de cobranza
        df_ahorros: DataFrame de AHORROS (hoja ACUMULADO)
        df_parche: DataFrame de Parche Promotores
        
    Returns:
        DataFrame con la estructura de la hoja CARTERA (36 columnas)
    """
    logger.info("Iniciando generación de cartera (VERSIÓN CORREGIDA)...")
    logger.info(f"Registros en antiguedad: {len(df_antiguedad)}")
    
    # Crear DataFrame base desde ANTIGÜEDAD
    df = df_antiguedad.copy()
    
    # ========== PASO 1: COLUMNAS BASE (extraídas directamente de ANTIGÜEDAD) ==========
    
    # C. ID de grupo - Formatear a 6 dígitos con ceros
    df['id_de_grupo'] = df['cod_grupo_solidario'].astype(str).str.zfill(6)
    logger.info(f"ID de grupo generados: {df['id_de_grupo'].head()}")
    
    # D. Nombre de grupo
    df['nombre_de_grupo'] = df['grupo_solidario']
    
    # F. Monto del crédito - CORRECCIÓN: Usar cantidad_prestada, si está NaN usar cantidad_entregada
    df['monto_del_credito'] = df['cantidad_prestada'].fillna(df['cantidad_entregada'])
    
    # G. Tipo de grupo
    df['tipo_de_grupo'] = df['tipo_de_grupo']
    
    # H. Fecha de inicio del crédito
    df['fecha_de_inicio_del_credito'] = pd.to_datetime(df['inicio_ciclo'], errors='coerce')
    
    # I. Plazo
    df['plazo'] = df['plazo_del_credito']
    
    # J. Día de reunión
    df['dia_de_reunion'] = df['dia_junta']
    
    # K. Hora de reunión
    df['hora_de_reunion'] = df['hora_junta']
    
    # L. Periodicidad
    df['periodicidad'] = df['periodicidad']
    
    # M. Pago Semanal
    df['pago_semanal'] = df['parcialidad_+_parcialidad_comision']
    
    # AF. Días de mora
    df['dias_de_mora'] = df['dias_de_mora']
    
    # ========== PASO 2: JOINS CON OTROS DATAFRAMES ==========
    
    # Preparar claves para joins
    df_situacion_join = df_situacion.copy()
    df_situacion_join['id_grupo_join'] = df_situacion_join['codigo'].astype(str).str.zfill(6)
    
    df_cobranza_join = df_cobranza.copy()
    df_cobranza_join['id_grupo_join'] = df_cobranza_join['gpo'].astype(str).str.zfill(6)
    
    df_ahorros_join = df_ahorros.copy()
    df_ahorros_join['id_grupo_join'] = df_ahorros_join['id'].astype(str).str.zfill(6)
    
    # Realizar joins
    # CORRECCIÓN: Agregar columnas adicionales de SITUACIÓN
    cols_sit = ['id_grupo_join', 'ciclo_sit', 'cartera_vencida_importe', 
                'cartera_vencida_pct', 'numero_de_integrantes_sit', 
                'cartera_vigente_importe', 'cartera_vigente_parcialidad']
    
    df = df.merge(
        df_situacion_join[cols_sit],
        left_on='id_de_grupo',
        right_on='id_grupo_join',
        how='left',
        suffixes=('', '_sit')
    )
    
    # CORRECCIÓN: Mapear columnas correctas de cobranza
    df = df.merge(
        df_cobranza_join[['id_grupo_join', 'proximo_pago_cob', 'por_vencer', 'pagos']],
        left_on='id_de_grupo',
        right_on='id_grupo_join',
        how='left',
        suffixes=('', '_cob')
    )
    
    df = df.merge(
        df_ahorros_join[['id_grupo_join', 'ahorro_acumulado']],
        left_on='id_de_grupo',
        right_on='id_grupo_join',
        how='left',
        suffixes=('', '_aho')
    )
    
    # Log de joins
    logger.info(f"Después de joins: {len(df)} registros")
    logger.info(f"NaN en situacion: {df['ciclo_sit'].isna().sum()}")
    logger.info(f"NaN en cobranza: {df['proximo_pago_cob'].isna().sum()}")
    logger.info(f"NaN en ahorros: {df['ahorro_acumulado'].isna().sum()}")
    
    # ========== PASO 3: COLUMNAS CON LÓGICA ESPECIAL ==========
    
    # A. Nombre del gerente (ya viene de ANTIGÜEDAD)
    df['nombre_del_gerente'] = df['nombre_de_gerente']
    
    # B. Nombre promotor - Aplicar parche
    df['nombre_promotor'] = df['nombre_promotor']
    if not df_parche.empty and 'original' in df_parche.columns and 'correcto' in df_parche.columns:
        parche_dict = dict(zip(df_parche['original'], df_parche['correcto']))
        df['nombre_promotor'] = df['nombre_promotor'].map(parche_dict).fillna(df['nombre_promotor'])
        logger.info(f"Parche aplicado: {len(parche_dict)} correcciones")
    
    # E. Ciclo - Con fallback
    df['ciclo'] = df['ciclo_sit'].fillna(df['ciclo'])
    
    # N. Próximo pago - CORRECCIÓN: Usar columna correcta de cobranza y formatear
    df['proximo_pago'] = df['proximo_pago_cob'].fillna("")
    
    # AG. Ahorro Acumulado
    df['ahorro_acumulado'] = df['ahorro_acumulado'].fillna(0)
    
    # ========== PASO 4: ESTATUS (crítico, se calcula antes) ==========
    
    # AI. Estatus
    def calcular_estatus(situacion_credito):
        if pd.isna(situacion_credito):
            return "Vigente"  # Por defecto, si no hay valor, asumir Vigente
        situacion_str = str(situacion_credito).strip()
        if situacion_str == "Entregado" or situacion_str == "Autorizado por cartera":
            return "Vigente"
        elif situacion_str == "Liquidado":
            return "Desertor sin mora"
        else:
            # Si es un valor desconocido, por defecto asumir Vigente
            logger.warning(f"Valor desconocido de situacion_credito: '{situacion_str}'. Asignando 'Vigente' por defecto")
            return "Vigente"
    
    df['estatus'] = df['situacion_credito'].apply(calcular_estatus)
    logger.info(f"Estatus calculados: {df['estatus'].value_counts().to_dict()}")
    
    # ========== PASO 5: COLUMNAS CONDICIONALES (dependen de Estatus) ==========
    
    # O. Cartera vigente sistema - CORRECCIÓN: Usar saldo_total de ANTIGÜEDAD
    # El valor esperado es directamente saldo_total de ANTIGÜEDAD
    # Ejemplos: 000089 -> 32,832.51, 000108 -> 106,395.49
    if 'saldo_total' in df.columns:
        df['cartera_vigente_sistema'] = np.where(
            df['estatus'] == "Desertor sin mora",
            0,
            df['saldo_total'].fillna(0)
        )
        logger.info("Cartera vigente sistema calculada como: saldo_total (de ANTIGÜEDAD)")
    else:
        logger.warning("Columna 'saldo_total' no encontrada; se utilizará cartera_vigente_importe")
        df['cartera_vigente_sistema'] = np.where(
            df['estatus'] == "Desertor sin mora",
            0,
            df['cartera_vigente_importe'].fillna(0)
        )
    
    # R. Cartera Insoluta - CORRECCIÓN: Usar cartera_vigente_importe (igual que O)
    df['cartera_insoluta'] = np.where(
        df['estatus'] == "Desertor sin mora",
        0,
        df['cartera_vigente_importe'].fillna(0)
    )
    
    # V. Cartera vencida Total
    df['cartera_vencida_total'] = np.where(
        df['estatus'] == "Desertor sin mora",
        0,
        df['cartera_vencida_importe'].fillna(0)
    )
    
    # W. % Mora
    df['pct_mora'] = np.where(
        df['estatus'] == "Desertor sin mora",
        0,
        df['cartera_vencida_pct'].fillna(0) / 100
    )
    
    # X. Saldo en riesgo
    df['saldo_en_riesgo'] = np.where(
        df['cartera_vencida_total'] > 0,
        df['cartera_vigente_importe'].fillna(0),
        0
    )
    
    # AA. Número de Integrantes - CORRECCIÓN: Invertir orden de prioridad
    df['numero_de_integrantes'] = np.where(
        df['estatus'] == "Desertor sin mora",
        df['numero_integrantes'].fillna(df['numero_de_integrantes_sit']),
        df['numero_de_integrantes_sit'].fillna(df['numero_integrantes'])
    )
    
    # Z. Monto promedio del grupo
    df['monto_promedio_del_grupo'] = np.where(
        df['estatus'] == "Desertor sin mora",
        df['cantidad_prestada'] / df['numero_de_integrantes'],
        df['monto_del_credito'] / df['numero_de_integrantes']
    )
    
    # AB. Semana - CORRECCIÓN: Usar columnas correctas
    today = pd.Timestamp.now()
    df['semana'] = np.where(
        df['estatus'] == "Desertor sin mora",
        ((today - df['fecha_de_inicio_del_credito']).dt.days / 7).fillna(0).astype(int),
        df['pagos'].fillna(0) - df['por_vencer'].fillna(0)
    )
    
    # AD. Pagos por vencer
    df['pagos_por_vencer'] = np.where(
        df['estatus'] == "Desertor sin mora",
        0,
        df['por_vencer'].fillna(0)
    )
    
    # AE. Total de pagos
    df['total_de_pagos'] = np.where(
        df['estatus'] == "Desertor sin mora",
        df['plazo'],
        df['pagos'].fillna(0)
    )
    
    # AC. Pagos cubiertos - CORRECCIÓN: Usar total_de_pagos - pagos_por_vencer
    # Similar a cómo se calculan las otras dos columnas (AD y AE)
    # Para Vigente: pagos (COBRANZA) - por_vencer (COBRANZA)
    # Para Desertor sin mora: plazo - 0 = plazo
    # Esta fórmula da el resultado correcto (ej: ID 000041 = 10.0)
    # vs la fórmula original (cartera_vigente_parcialidad / pago_semanal) que da 0.822361
    df['pagos_cubiertos'] = df['total_de_pagos'] - df['pagos_por_vencer']
    
    # ========== PASO 6: COLUMNAS CALCULADAS ==========
    
    # P. Cartera vigente inicial
    df['cartera_vigente_inicial'] = df['pago_semanal'] * 16
    
    # Q. Cartera vigente calculada
    calc_temp = -((df['semana'] * df['pago_semanal']) - df['cartera_vigente_inicial'])
    df['cartera_vigente_calculada'] = np.maximum(calc_temp, df['cartera_vigente_sistema'])
    
    # S. Diferencia Validación vigente
    df['diferencia_validacion_vigente'] = df['cartera_vigente_sistema'] - df['cartera_vigente_calculada']
    
    # T. Ahorro Consumido
    ahorro_mas_10pct = df['ahorro_acumulado'] + (df['monto_del_credito'] * 0.1)
    df['ahorro_consumido'] = np.where(
        df['cartera_vencida_total'] > 0,
        np.where(
            ahorro_mas_10pct > df['cartera_vencida_total'],
            df['cartera_vencida_total'],
            ahorro_mas_10pct
        ),
        0
    )
    
    # U. Cartera Vencida Estadística
    df['cartera_vencida_estadistica'] = df['cartera_vencida_total'] - df['ahorro_consumido']
    
    # Y. Saldo ahorro acumulado
    df['saldo_ahorro_acumulado'] = df['ahorro_acumulado']
    
    # AH. % de Ahorro
    denominador = df['pago_semanal'] * df['semana']
    df['pct_de_ahorro'] = np.where(
        denominador != 0,
        df['ahorro_acumulado'] / denominador,
        0
    )
    df['pct_de_ahorro'] = df['pct_de_ahorro'].replace([np.inf, -np.inf], 0).fillna(0)
    
    # AJ. Concepto Depósito
    df['concepto_deposito'] = (
        "0" + 
        df['id_de_grupo'].astype(str).str.zfill(6) + 
        df['ciclo'].astype(int).astype(str).str.zfill(2)
    )
    
    # ========== PASO 7: SELECCIONAR Y ORDENAR COLUMNAS FINALES ==========
    
    columnas_finales = [
        'nombre_del_gerente',           # A
        'nombre_promotor',              # B
        'id_de_grupo',                  # C
        'nombre_de_grupo',              # D
        'ciclo',                        # E
        'monto_del_credito',            # F
        'tipo_de_grupo',                # G
        'fecha_de_inicio_del_credito',  # H
        'plazo',                        # I
        'dia_de_reunion',               # J
        'hora_de_reunion',              # K
        'periodicidad',                 # L
        'pago_semanal',                 # M
        'proximo_pago',                 # N
        'cartera_vigente_sistema',      # O
        'cartera_vigente_inicial',      # P
        'cartera_vigente_calculada',    # Q
        'cartera_insoluta',             # R
        'diferencia_validacion_vigente', # S
        'ahorro_consumido',             # T
        'cartera_vencida_estadistica',  # U
        'cartera_vencida_total',        # V
        'pct_mora',                     # W
        'saldo_en_riesgo',              # X
        'saldo_ahorro_acumulado',       # Y
        'monto_promedio_del_grupo',     # Z
        'numero_de_integrantes',        # AA
        'semana',                       # AB
        'pagos_cubiertos',              # AC
        'pagos_por_vencer',             # AD
        'total_de_pagos',               # AE
        'dias_de_mora',                 # AF
        'ahorro_acumulado',             # AG
        'pct_de_ahorro',                # AH
        'estatus',                      # AI
        'concepto_deposito'             # AJ
    ]
    
    # Verificar que todas las columnas existen
    columnas_faltantes = [col for col in columnas_finales if col not in df.columns]
    if columnas_faltantes:
        logger.error(f"Columnas faltantes: {columnas_faltantes}")
        raise ValueError(f"Columnas faltantes en el DataFrame final: {columnas_faltantes}")
    
    df_final = df[columnas_finales].copy()
    
    logger.info(f"Cartera generada exitosamente: {len(df_final)} filas x {len(df_final.columns)} columnas")
    
    return df_final


def generar_mora(df_cartera: pd.DataFrame) -> pd.DataFrame:
    """
    Genera el DataFrame de la hoja MORA filtrando grupos con %mora > 5%.
    
    Args:
        df_cartera: DataFrame de CARTERA generado
        
    Returns:
        DataFrame con la estructura de la hoja MORA (14 columnas)
    """
    logger.info("\nIniciando generación de hoja MORA...")
    
    # Filtrar registros con %mora > 5%
    # La columna de %mora en CARTERA se llama 'pct_mora'
    df_mora = df_cartera[df_cartera['pct_mora'] > 0.05].copy()
    logger.info(f"Registros con %mora > 5%: {len(df_mora)} de {len(df_cartera)}")
    
    if len(df_mora) == 0:
        logger.warning("No hay registros con %mora > 5%")
        # Crear DataFrame vacío con las columnas esperadas
        return pd.DataFrame(columns=[
            'nombre_del_gerente',
            'nombre_promotor',
            'id_de_grupo',
            'nombre_de_grupo',
            'ciclo',
            'monto_del_credito',
            'semana',
            'pago_semanal',
            'cartera_vencida_total',
            'pct_mora',
            'saldo_en_riesgo',
            'dias_de_mora',
            'mora_potencial_mensual',
            'cartera_vencida_total_calculada'
        ])
    
    # Seleccionar y renombrar columnas según especificación
    df_mora_final = pd.DataFrame({
        # A. Nombre del gerente
        'nombre_del_gerente': df_mora['nombre_del_gerente'],
        
        # B. Nombre del promotor
        'nombre_promotor': df_mora['nombre_promotor'],
        
        # C. ID GRUPO
        'id_de_grupo': df_mora['id_de_grupo'],
        
        # D. Nombre de grupo
        'nombre_de_grupo': df_mora['nombre_de_grupo'],
        
        # E. Ciclo
        'ciclo': df_mora['ciclo'],
        
        # F. Monto del crédito
        'monto_del_credito': df_mora['monto_del_credito'],
        
        # G. Semana
        'semana': df_mora['semana'],
        
        # H. Pago semanal
        'pago_semanal': df_mora['pago_semanal'],
        
        # I. Cartera vencida total
        'cartera_vencida_total': df_mora['cartera_vencida_total'],
        
        # J. %mora (con relleno amarillo)
        'pct_mora': df_mora['pct_mora'],
        
        # K. Saldo en riesgo
        'saldo_en_riesgo': df_mora['saldo_en_riesgo'],
        
        # L. Días de mora (con relleno amarillo)
        'dias_de_mora': df_mora['dias_de_mora'],
    })
    
    # M. Mora potencial mensual = pago_semanal * 4
    df_mora_final['mora_potencial_mensual'] = np.where(
        df_mora_final['pago_semanal'].notna(),
        df_mora_final['pago_semanal'] * 4,
        np.nan
    )
    
    # N. Cartera vencida total calculada = pago_semanal * semana
    df_mora_final['cartera_vencida_total_calculada'] = np.where(
        (df_mora_final['pago_semanal'].notna()) & (df_mora_final['semana'].notna()),
        df_mora_final['pago_semanal'] * df_mora_final['semana'],
        np.nan
    )
    
    logger.info(f"Hoja MORA generada con {len(df_mora_final)} registros y {len(df_mora_final.columns)} columnas")
    
    return df_mora_final


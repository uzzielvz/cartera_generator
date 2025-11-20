"""
Parche de grupos: correcciones de nombre_promotor por ID de grupo.
Este diccionario se usa para corregir nombres de promotores para grupos específicos.
"""

import pandas as pd

# Mapeo de correcciones: ID de grupo -> Nombre correcto del promotor
CORRECCIONES_GRUPOS = {
    '000184': 'Garcia Herrera Jonathan',
    '000216': 'Olivares Morales Josue Edgar'
}


def aplicar_parche_grupos(df: pd.DataFrame, columna_id: str, columna_promotor: str) -> pd.DataFrame:
    """
    Aplica el parche de grupos a un DataFrame.
    Modifica el nombre_promotor para grupos específicos.
    
    Args:
        df: DataFrame a modificar
        columna_id: Nombre de la columna que contiene el ID del grupo (normalmente 'id_de_grupo')
        columna_promotor: Nombre de la columna que contiene el nombre del promotor (normalmente 'nombre_promotor')
        
    Returns:
        DataFrame con correcciones aplicadas
    """
    df = df.copy()
    
    # Normalizar IDs a string con ceros a la izquierda
    df[columna_id] = df[columna_id].astype(str).str.zfill(6)
    
    # Aplicar correcciones
    for id_grupo, nombre_correcto in CORRECCIONES_GRUPOS.items():
        mask = df[columna_id] == id_grupo
        if mask.any():
            df.loc[mask, columna_promotor] = nombre_correcto
    
    return df


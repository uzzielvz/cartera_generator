"""
Parche de promotores: correcciones de nombres.
Este diccionario se usa para corregir nombres de promotores mal escritos.
"""

import pandas as pd

# Mapeo de correcciones: Original -> Correcto
CORRECCIONES_PROMOTORES = {
    'Ponce Galindo': 'Contreras Martinez Jose Luis'
}


def obtener_parche() -> pd.DataFrame:
    """
    Retorna el DataFrame de parche promotores.
    
    Returns:
        DataFrame con columnas 'original' y 'correcto'
    """
    data = [
        {'original': k, 'correcto': v} 
        for k, v in CORRECCIONES_PROMOTORES.items()
    ]
    return pd.DataFrame(data)


def aplicar_parche(df: pd.DataFrame, columna: str) -> pd.DataFrame:
    """
    Aplica el parche de promotores a una columna del DataFrame.
    
    Args:
        df: DataFrame a modificar
        columna: Nombre de la columna a corregir
        
    Returns:
        DataFrame con correcciones aplicadas
    """
    df = df.copy()
    for original, correcto in CORRECCIONES_PROMOTORES.items():
        df[columna] = df[columna].replace(original, correcto)
    return df


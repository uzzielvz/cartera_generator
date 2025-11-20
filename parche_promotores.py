"""
Parche de promotores: correcciones de nombres.
Este diccionario se usa para corregir nombres de promotores mal escritos.
"""

import pandas as pd

# Mapeo de correcciones: Original -> Correcto
CORRECCIONES_PROMOTORES = {
    'Ponce Galindo': 'Contreras Martinez Jose Luis'
}

# Mapeo de correcciones para gerentes: Original -> Correcto
CORRECCIONES_GERENTES = {
    'JUAN EDMIUNDO': 'JUAN EDMUNDO'
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
    Usa coincidencia parcial (regex) para encontrar y reemplazar.
    
    Args:
        df: DataFrame a modificar
        columna: Nombre de la columna a corregir
        
    Returns:
        DataFrame con correcciones aplicadas
    """
    df = df.copy()
    for original, correcto in CORRECCIONES_PROMOTORES.items():
        # Usar regex para coincidencia parcial (ej: "Ponce Galindo Alicia" -> "Contreras Martinez Jose Luis")
        df[columna] = df[columna].astype(str).str.replace(
            rf'.*{original}.*', 
            correcto, 
            regex=True, 
            case=False
        )
    return df


def aplicar_parche_gerentes(df: pd.DataFrame, columna: str) -> pd.DataFrame:
    """
    Aplica el parche de gerentes a una columna del DataFrame.
    Usa coincidencia parcial (regex) para encontrar y reemplazar.
    
    Args:
        df: DataFrame a modificar
        columna: Nombre de la columna a corregir (normalmente 'nombre_de_gerente' o 'nombre_del_gerente')
        
    Returns:
        DataFrame con correcciones aplicadas
    """
    df = df.copy()
    for original, correcto in CORRECCIONES_GERENTES.items():
        # Usar regex para coincidencia parcial (ej: "JUAN EDMIUNDO LUNA" -> "JUAN EDMUNDO LUNA")
        df[columna] = df[columna].astype(str).str.replace(
            original, 
            correcto, 
            regex=False, 
            case=False
        )
    return df


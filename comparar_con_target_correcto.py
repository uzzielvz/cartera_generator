import pandas as pd
import openpyxl
import numpy as np

print("Comparando output con target (hoja CARTERA)...")
print("=" * 60)

# Leer archivos
ruta_output = 'output_automatizado.xlsx'
ruta_target = 'data/target/AntigüedadGrupal_171125.xlsm'

print(f"\n1. Leyendo archivos...")
print(f"   Output: {ruta_output}")
print(f"   Target: {ruta_target} (hoja CARTERA)")

# Leer output
df_output = pd.read_excel(ruta_output, sheet_name=0, header=None, skiprows=5)
df_output.columns = range(len(df_output.columns))

# Leer target - hoja CARTERA
df_target = pd.read_excel(ruta_target, sheet_name='CARTERA', header=None, skiprows=5, engine='openpyxl')
df_target.columns = range(len(df_target.columns))

print(f"\n2. Dimensiones:")
print(f"   Output: {len(df_output)} filas x {len(df_output.columns)} columnas")
print(f"   Target: {len(df_target)} filas x {len(df_target.columns)} columnas")

# Normalizar IDs - Columna C (índice 2) es ID de grupo
df_output['id'] = df_output[2].astype(str).str.zfill(6)
df_target['id'] = df_target[2].astype(str).str.zfill(6)

# Filtrar solo IDs válidos (numéricos de 6 dígitos)
df_output = df_output[df_output['id'].str.match(r'^\d{6}$')].copy()
df_target = df_target[df_target['id'].str.match(r'^\d{6}$')].copy()

print(f"\n3. IDs válidos:")
print(f"   Output: {len(df_output)} IDs")
print(f"   Target: {len(df_target)} IDs")

# Obtener IDs comunes
ids_output = set(df_output['id'].unique())
ids_target = set(df_target['id'].unique())
ids_comunes = ids_output & ids_target
ids_solo_output = ids_output - ids_target
ids_solo_target = ids_target - ids_output

print(f"\n4. Comparación de IDs:")
print(f"   IDs comunes: {len(ids_comunes)}")
print(f"   IDs solo en output: {len(ids_solo_output)}")
if ids_solo_output:
    print(f"      Ejemplos: {list(ids_solo_output)[:10]}")
print(f"   IDs solo en target: {len(ids_solo_target)}")
if ids_solo_target:
    print(f"      Ejemplos: {list(ids_solo_target)[:10]}")

# Comparar columnas clave
columnas_importantes = {
    0: 'Nombre del gerente',
    1: 'Nombre promotor',
    2: 'ID de grupo',
    3: 'Nombre de grupo',
    4: 'Ciclo',
    5: 'Monto del crédito',
    12: 'Pago Semanal',
    14: 'Cartera vigente sistema',
    15: 'Carera vigente inicial',
    16: 'Cartera vigente calculada',
    17: 'Cartera Insoluta',
    18: 'Diferencia Validación vigente',
    19: 'Ahorro Consumido',
    20: 'Cartera Vencida Estadistica',
    21: 'Cartera vencida Total',
}

print(f"\n5. Comparando valores por columna (solo IDs comunes):")
print("=" * 60)

diferencias_totales = 0
tolerancia = 0.01  # Para comparaciones numéricas

for col_idx, nombre_col in columnas_importantes.items():
    if col_idx >= len(df_output.columns) or col_idx >= len(df_target.columns):
        continue
    
    diferencias = 0
    diferencias_detalle = []
    
    for id_grupo in sorted(ids_comunes):
        row_output = df_output[df_output['id'] == id_grupo]
        row_target = df_target[df_target['id'] == id_grupo]
        
        if len(row_output) == 0 or len(row_target) == 0:
            continue
        
        val_output = row_output[col_idx].iloc[0]
        val_target = row_target[col_idx].iloc[0]
        
        # Manejar NaN
        if pd.isna(val_output):
            val_output = None
        if pd.isna(val_target):
            val_target = None
        
        # Comparar
        if val_output != val_target:
            # Si son numéricos, comparar con tolerancia
            try:
                if val_output is not None and val_target is not None:
                    num_output = float(val_output)
                    num_target = float(val_target)
                    if abs(num_output - num_target) > tolerancia:
                        diferencias += 1
                        if len(diferencias_detalle) < 5:  # Mostrar solo primeras 5
                            diferencias_detalle.append({
                                'id': id_grupo,
                                'output': val_output,
                                'target': val_target,
                                'diff': abs(num_output - num_target)
                            })
                elif val_output != val_target:
                    diferencias += 1
                    if len(diferencias_detalle) < 5:
                        diferencias_detalle.append({
                            'id': id_grupo,
                            'output': val_output,
                            'target': val_target
                        })
            except (ValueError, TypeError):
                # No son numéricos, comparar como strings
                if str(val_output).strip() != str(val_target).strip():
                    diferencias += 1
                    if len(diferencias_detalle) < 5:
                        diferencias_detalle.append({
                            'id': id_grupo,
                            'output': val_output,
                            'target': val_target
                        })
    
    if diferencias > 0:
        print(f"\n{nombre_col} (columna {col_idx}): {diferencias} diferencias")
        for det in diferencias_detalle:
            if 'diff' in det:
                print(f"   ID {det['id']}: Output={det['output']}, Target={det['target']}, Diff={det['diff']:.2f}")
            else:
                print(f"   ID {det['id']}: Output={det['output']}, Target={det['target']}")
        diferencias_totales += diferencias
    else:
        print(f"{nombre_col}: OK")

print(f"\n" + "=" * 60)
print(f"RESUMEN:")
print(f"   Diferencias totales encontradas: {diferencias_totales}")
if diferencias_totales == 0:
    print(f"   OK: TODAS LAS COLUMNAS COINCIDEN")
else:
    print(f"   ATENCION: Se encontraron {diferencias_totales} diferencias")


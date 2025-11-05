# Automatización de Reporte de Cartera

Sistema automatizado para generar la hoja CARTERA replicando las fórmulas del machote Excel en Python.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
python analizar_y_automatizar.py
```

## Archivos de Entrada

Coloca estos archivos en `/data`:

- `ReportedeAntiguedaddeCarteraGrupal_DDMMYYYY.xlsx`
- `Situación de cartera DDMMYYYY.xlsx`
- `Cobranza DDMMYYYY.xlsx`
- `AHORROS.xlsx`
- `Copia de AntigüedadGrupal_machote.xlsm`

## Salida

`output_automatizado.xlsx` - Hoja CARTERA con 36 columnas calculadas

## Estructura

```
analizar_y_automatizar.py  - Script principal
cartera_generator.py        - Lógica de generación
requirements.txt            - Dependencias
output_automatizado.xlsx    - Resultado
```

## Función Principal

```python
from cartera_generator import generar_cartera

df_cartera = generar_cartera(
    df_antiguedad,
    df_situacion,
    df_cobranza,
    df_ahorros,
    df_parche
)
```

## Características

- Replica fórmulas Excel (VLOOKUP, IF, IFERROR)
- Manejo automático de NaN y errores
- Logging detallado
- Validación automática
- Función testeable

## Notas

- Usar archivos de la misma fecha para máxima coincidencia
- El script genera log en `cartera_automation.log`
- Todas las fórmulas se calculan automáticamente

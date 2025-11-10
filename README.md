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

`output_automatizado.xlsx` con 2 hojas:

### Hoja CARTERA
- 36 columnas calculadas
- Tabla Excel con totales automáticos
- Formato idéntico al machote

### Hoja MORA
- Filtro automático: registros con %mora > 5%
- 14 columnas seleccionadas
- Columnas resaltadas en amarillo: %mora y Días de mora
- Fórmulas calculadas en Python:
  - Mora potencial mensual = Pago semanal × 4
  - Cartera vencida calculada = Pago semanal × Semana

## Estructura

```
analizar_y_automatizar.py     - Script principal
cartera_generator.py           - Lógica de generación
formato_excel.py               - Formato Excel con tablas y totales
crear_plantilla.py             - Generador de plantilla (ejecutar una vez)
plantilla/CARTERA_HEADERS.xlsx - Plantilla ligera (6.1 KB)
requirements.txt               - Dependencias
output_automatizado.xlsx       - Resultado con formato
```

## Primera Vez

Si no existe la plantilla:

```bash
python crear_plantilla.py
```

Esto crea `plantilla/CARTERA_HEADERS.xlsx` (6.1 KB) independiente del machote.

## Funciones Principales

### Generar Cartera
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

### Generar Mora
```python
from cartera_generator import generar_mora

df_mora = generar_mora(df_cartera)  # Filtra registros con %mora > 5%
```

## Características

- Replica fórmulas Excel (VLOOKUP, IF, IFERROR)
- Formato idéntico al machote
- Tabla Excel con totales automáticos (fila de totales con fórmulas SUBTOTAL)
- **Hoja MORA automática** con filtrado inteligente (%mora > 5%)
- Columnas con formato condicional (amarillo para alertas)
- Plantilla ligera independiente del machote (99.6% más pequeña)
- Manejo automático de NaN y errores
- Logging detallado
- Validación automática
- Función testeable

## Notas

- Usar archivos de la misma fecha para máxima coincidencia
- El script genera log en `cartera_automation.log`
- Todas las fórmulas se calculan automáticamente

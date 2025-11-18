# REPORTE DE COMPARACIÓN CON TARGET

**Fecha:** 18 de noviembre de 2025  
**Archivo target:** `AntigüedadGrupal_171125.xlsm` (hoja CARTERA)  
**Archivo output:** `output_automatizado.xlsx`

## RESUMEN GENERAL

- **IDs comunes:** 224
- **IDs solo en output:** 0
- **IDs solo en target:** 1 (ID 000000 - probablemente fila de totales)
- **Diferencias totales:** 21

## DIFERENCIAS ENCONTRADAS

### 1. Nombre del Gerente (4 diferencias)

**IDs afectados:** 000231, 000235, 000238, 000240

- **IDs 000231, 000235, 000238:**
  - Output: "Región Estado México" (tomado de `nombre_gerencia_regional` de SITUACIÓN)
  - Target: "JUDITH PATRICIA SERRANO HERRER"
  - **Causa:** Estos IDs tienen `nombre_de_gerente` vacío en ANTIGÜEDAD. Usamos `nombre_gerencia_regional` como fallback, pero el target usa el gerente más común de la misma coordinación "Ixtapaluca".

- **ID 000240:**
  - Output: "JUAN EDMIUNDO LUNA AGUILLON" (typo)
  - Target: "JUAN EDMUNDO LUNA AGUILLON"
  - **Causa:** Error de tipeo en los datos originales de ANTIGÜEDAD.

### 2. Nombre del Promotor (7 diferencias)

**IDs afectados:** 000009, 000022, 000048, 000067, 000073

- Output: "Ponce Galindo Alicia Berenice" (de ANTIGÜEDAD)
- Target: "Contreras Martinez Jose Luis"
- **Causa:** Los datos originales en ANTIGÜEDAD tienen un promotor diferente al que aparece en el target. Esto podría ser:
  - Un parche/corrección manual en el target
  - Datos actualizados en el target que no están en los archivos de entrada
  - Cambio de promotor que no se reflejó en ANTIGÜEDAD

### 3. Cartera Insoluta (10 diferencias)

**IDs afectados:** 000022, 000031, 000048, 000055, 000077, etc.

**Ejemplos:**
- ID 000022: Output=30500, Target=59140.66, Diff=28640.66
- ID 000031: Output=19250, Target=19270.62, Diff=20.62
- ID 000048: Output=43750, Target=66622.84, Diff=22872.84

**Causa:** 
- **Nuestro código actual:** `cartera_insoluta = cartera_vigente_importe` (cuando estatus != "Desertor sin mora")
- **Fórmula correcta (según investigación):** `cartera_insoluta = cartera_vigente_importe - cartera_vigente_parcialidad`

**Verificación:**
- ID 000022: cartera_vigente_importe=30500, cartera_vigente_parcialidad=7625 → Calculado=22875 (pero target tiene 59140.66)
- Esto sugiere que el target podría estar usando una fórmula diferente o datos diferentes.

## COLUMNAS QUE COINCIDEN PERFECTAMENTE

✅ ID de grupo  
✅ Nombre de grupo  
✅ Ciclo  
✅ Monto del crédito  
✅ Pago Semanal  
✅ Cartera vigente sistema  
✅ Carera vigente inicial  
✅ Cartera vigente calculada  
✅ Diferencia Validación vigente  
✅ Ahorro Consumido  
✅ Cartera Vencida Estadistica  
✅ Cartera vencida Total  

## RECOMENDACIONES

1. **Cartera Insoluta:** Verificar la fórmula correcta. El target tiene valores diferentes que no coinciden con `cartera_vigente_importe - cartera_vigente_parcialidad`.

2. **Nombre del Gerente:** Cuando `nombre_de_gerente` está vacío, usar el gerente más común de la misma coordinación en lugar de `nombre_gerencia_regional`.

3. **Nombre del Promotor:** Verificar si hay un parche o actualización de promotores que deba aplicarse.

4. **Typo en nombre:** Corregir "JUAN EDMIUNDO" a "JUAN EDMUNDO" en los datos de entrada o aplicar un parche.


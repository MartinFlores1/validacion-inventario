# Validación de inventario con Python

Cruza el inventario de almacén contra la base de costos y pesos, y genera un reporte en Excel con todo lo que hay que corregir **antes** de usar los datos para costear recetas o calcular márgenes.

Nace de un problema real del trabajo en restaurantes: el mismo producto se registra con nombres distintos, hay códigos repetidos, insumos sin costo y pesos en cero. Si nadie lo revisa, esos errores terminan en los reportes de gerencia.

> Los datos de este repositorio son **ficticios**. Los genera `generar_datos.py`.

## Qué revisa

| Revisión | Qué detecta |
|---|---|
| Códigos duplicados | Un mismo código usado para más de un registro |
| Faltantes | Productos del inventario sin costo registrado, y costos de productos que ya no están en inventario |
| Costos inválidos | Costo vacío o en cero |
| Pesos inválidos | Peso vacío, en cero o negativo |
| Nombres distintos | Mismo código con nombre diferente en cada archivo (después de normalizar) |
| Posibles duplicados | Distinto código, mismo producto: nombres casi iguales con la misma presentación |

Antes de comparar nombres, los normaliza: minúsculas, sin tildes y sin espacios repetidos. Así `AZUCAR RUBIA  x 1 kg` y `Azúcar rubia x 1 kg` cuentan como el mismo nombre y no generan falsos positivos.

## Cómo usarlo

```bash
pip install -r requirements.txt
python generar_datos.py        # crea data/inventario.xlsx y data/base_costos.xlsx
python validar_inventario.py   # crea salida/reporte_validacion.xlsx
```

Con tus propios archivos:

```bash
python validar_inventario.py --inventario mi_inventario.xlsx --costos mis_costos.xlsx
```

Columnas esperadas:
- **Inventario:** `codigo`, `nombre`, `categoria`, `unidad`, `stock`
- **Base de costos:** `codigo`, `nombre`, `costo_unitario`, `peso_kg`

## Resultado con los datos de ejemplo

```
Productos en inventario: 315 | en base de costos: 306
  Codigos duplicados       12 casos
  Faltantes                17 casos
  Costos invalidos          9 casos
  Pesos invalidos           7 casos
  Nombres distintos         3 casos
  Posibles duplicados       5 casos
```

El Excel de salida trae una hoja de **Resumen** y una hoja por cada revisión, con el ancho de columnas ajustado para leerlo directo.

## Tecnologías

Python · pandas · openpyxl · difflib

## Próximos pasos

- Leer directamente desde la base de datos (SQL Server) en lugar de Excel
- Sugerir automáticamente cuál de los duplicados conservar
- Programarlo para que corra cada semana y envíe el reporte por correo

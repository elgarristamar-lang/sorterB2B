# Sorter VDL B2B — Configurador de semanas especiales

Herramienta para generar el fichero `GRUPO_DESTINOS` del sorter VDL B2B en semanas con salidas canceladas o que cambian de día (festivos, Semana Santa, etc.), con visualizaciones para los equipos operativos.

## Requisitos

```bash
pip install openpyxl pandas
```

## Uso rápido — interfaz web

```bash
python app.py
```

Se abre automáticamente en `http://localhost:5001`. Sube los ficheros, pulsa **Generar configuración** y descarga los outputs.

## Uso por línea de comandos

```bash
# 1. Generar GRUPO_DESTINOS + resumen HTML
python process_parrilla.py <parrilla.xlsx> <grupo_destinos.xlsx> <ramp_capacity.csv> <hoja> [semana]

# 2. Gantt 1H visual  (requiere bloques_horarios.xlsx)
python gantt_1h.py <ramp_capacity.csv> <grupo_destinos.xlsx> <bloques_horarios.xlsx> <output.xlsx> [hoja]

# 3. Sorter Map por día  (requiere bloques_horarios.xlsx)
python sorter_map_por_dia.py <ramp_capacity.csv> <grupo_destinos.xlsx> <bloques_horarios.xlsx> <output.xlsx> [hoja]
```

## Ficheros de entrada

| Fichero | Obligatorio | Descripción |
|---|---|---|
| `parrilla_de_salidas.xlsx` | ✓ | Parrilla semanal. Columnas: `PLAYA`, `DIA_SALIDA`, `DIA_SALIDA_NEW`, `TIPO_SALIDA`, `ID_CLUSTER`. Debe incluir la hoja `Resumen Bloques`. |
| `GRUPO_DESTINOS.xlsx` | ✓ | Config actual del sorter. Acepta formato clásico o export DXC (con columna `Estado`). |
| `ramp_capacity.csv` | ✓ | Capacidad de cada sub-rampa. Columnas: `RAMP;PALLETS`. |
| `bloques_horarios.xlsx` | opcional | Para Gantt y Sorter Map. Columnas: `NUEVO BLOQUE`, `Día LIBERACIÓN BLOQUES`, `Hora LIBERACIÓN BLOQUES`, `Día DESACTIVACIÓN`, `Hora DESACTIVACIÓN`. |

### Tipos de salida en la parrilla

| Tipo | Acción |
|---|---|
| `HABITUAL` | Sin cambios — se mantiene en el GD |
| `CANCELADA` | Se elimina del GD esta semana |
| `ESPECIAL DIA CAMBIO` | Se reasigna a rampas libres en el nuevo día |
| `IRREGULAR` | Ignorada (no pasa por el GD semanal) |

## Outputs

| Fichero | Script | Descripción |
|---|---|---|
| `GRUPO_DESTINOS_{SEMANA}.xlsx` | `process_parrilla.py` | Listo para subir a DXC/MAR |
| `resumen_sorter_{SEMANA}.html` | `process_parrilla.py` | Informe interactivo: gráfico de ocupación por bloque, panel de detalle por playa, lista de canceladas |
| `gantt_1h_{SEMANA}.xlsx` | `gantt_1h.py` | Gantt visual hora a hora: rampas × tiempo, coloreado por bloque. Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO |
| `sorter_map_{SEMANA}.xlsx` | `sorter_map_por_dia.py` | Mapa de slots físicos: 1 pestaña por día, posiciones POSTEX coloreadas por bloque |

## Flujo completo

```
parrilla.xlsx ─┐
gd.xlsx        ├─► process_parrilla.py ──► GRUPO_DESTINOS_S14.xlsx
ramp_cap.csv ──┘                       └─► resumen_S14.html

bloques_horarios.xlsx ─┬─► gantt_1h.py          ──► gantt_1h_S14.xlsx
 + GRUPO_DESTINOS      └─► sorter_map_por_dia.py ──► sorter_map_S14.xlsx
```

## Lógica de asignación (`process_parrilla.py`)

- **Slots como unidad** — la unidad de asignación es el slot físico `(rampa, posición)`, que puede agrupar múltiples destinos (e.g. MEXICO empaqueta ~63 códigos de tienda por slot).
- **Sin colisiones** — acumula las posiciones asignadas en el mismo ciclo para que dos destinos que van al mismo día nuevo no se solapen dentro del mismo bloque.
- **Control horario** — dos destinos pueden compartir rampa si sus ventanas de bloque `[liberación, desactivación]` no se solapan en el tiempo.
- **Bloque correcto** — deriva el bloque del nuevo día desde el `ID_CLUSTER` del destino cuando la parrilla devuelve `#N/A`.
- **Rampas excluidas** — configurable al principio del script:

```python
EXCLUDED_RAMPAS = {
    'R03A', 'R03B', 'R03C', 'R03D',   # Rampa 3: manipulado exclusivo
}
```

## Estructura del repo

```
app.py                   # Servidor web local → http://localhost:5001
process_parrilla.py      # Genera GRUPO_DESTINOS y resumen HTML
gantt_1h.py              # Genera Gantt 1H visual
sorter_map_por_dia.py    # Genera Sorter Map por día
README.md
.gitignore
```

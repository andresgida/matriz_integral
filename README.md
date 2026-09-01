# Matriz Integral - Monitor de Tickets Zoho Analytics

Sistema de monitoreo diario de tickets descargados de Zoho Analytics.
Compara el estado actual contra el snapshot anterior y genera un reporte
Excel con los movimientos entre estados y una interfaz web opcional.

## Contenido

- `monitor.py` — script principal que compara los 3 Excel de entrada y
  genera el reporte con graficos, historial y hojas de detalle.
- `app.py` — interfaz grafica en Streamlit para subir los Excel sin
  preocuparse por el nombre y ejecutar el monitor con un clic.
- `requirements.txt` — dependencias Python.
- `ejecutar_monitor.bat` / `ejecutar_app.bat` — accesos rapidos en Windows.
- `assets/` — logos e iconos usados por la interfaz.

## Requisitos

- Python 3.11+ (probado con 3.12)
- Dependencias: ver `requirements.txt`

## Instalacion

```powershell
py -m pip install -r requirements.txt
```

## Uso rapido

1. Descargar de Zoho Analytics los 3 archivos Excel:
   - Tickets nuevos / abiertos
   - Tickets en revision
   - Tickets devueltos a servicios
2. Ejecutar la interfaz grafica:
   ```powershell
   py -m streamlit run app.py
   ```
   Abrir `http://localhost:8501` y arrastrar los 3 archivos. El sistema
   los renombra automaticamente y ejecuta el monitor.

O, si prefieres CLI:

```powershell
# Copia los 3 archivos a la carpeta entrada/ con los nombres:
#   tickets_nuevos.xlsx, tickets_revision.xlsx, tickets_devuelto.xlsx
py monitor.py
```

El reporte se guarda en `reportes/Reporte_Monitoreo_YYYY-MM-DD_HH-MM-SS.xlsx`.

## Carpetas

| Carpeta | Uso |
|---|---|
| `entrada/` | Excel descargados de Zoho (no se sube al repo) |
| `historico/` | Snapshots por ejecucion + `historial_movimientos.csv` (no se sube al repo) |
| `reportes/` | Reportes Excel generados (no se sube al repo) |
| `assets/` | Logos e iconos de la app |

## Automatizacion

Ver `docs/automatizacion.md` (proximamente) para configurar:

- Descarga automatica desde Zoho Analytics via API
- Ejecucion programada (Task Scheduler / GitHub Actions)
- Deploy en Streamlit Community Cloud

## Licencia

Uso interno.

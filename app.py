"""
app.py - Interfaz grafica (Streamlit) para el monitor de tickets
================================================================

Permite subir los 3 archivos Excel descargados de Zoho Analytics
sin importar como se llamen, y los renombra automaticamente a los
nombres que espera monitor.py:

    tickets_nuevos.xlsx
    tickets_revision.xlsx
    tickets_devuelto.xlsx

Luego ejecuta el monitor y muestra el resultado con boton de descarga.

Como se ejecuta:
    pip install streamlit pandas openpyxl xlsxwriter
    streamlit run app.py
"""
from __future__ import annotations

import io
import sys
import shutil
import subprocess
from pathlib import Path

import streamlit as st

# Reusamos las constantes ya definidas en monitor.py para no duplicar
# nombres de archivo, carpetas, etc.
from monitor import (
    ARCHIVOS,
    ENTRADA_DIR,
    REPORTES_DIR,
    HISTORICO_DIR,
    asegurar_carpetas,
)


# ---------------------------------------------------------------------------
# Configuracion de la pagina
# ---------------------------------------------------------------------------
BASE_DIR   = Path(__file__).resolve().parent
LOGO_PATH  = BASE_DIR / "assets" / "ofima-logo.webp"
FAVICON    = BASE_DIR / "assets" / "ofima-favicon.webp"

st.set_page_config(
    page_title="Monitor de Tickets - Zoho Analytics",
    page_icon=str(FAVICON) if FAVICON.exists() else None,
    layout="wide",
)

# Encabezado con el logo de Ofima en lugar del emoji anterior.
col_logo, col_txt = st.columns([1, 5], vertical_alignment="center")
with col_logo:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=140)
with col_txt:
    st.title("Monitor de Tickets - Zoho Analytics")
    st.caption(
        "Sube los 3 archivos de Excel descargados de Zoho. El sistema los "
        "renombra automaticamente y ejecuta el monitoreo."
    )

asegurar_carpetas()


# ---------------------------------------------------------------------------
# 1) Zona de carga de archivos
# ---------------------------------------------------------------------------
st.header("1. Cargar archivos de Zoho")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**Tickets Nuevos / Abiertos**")
    file_nuevos = st.file_uploader(
        "Archivo con los tickets en estado Nuevo o Abierto",
        type=["xlsx"],
        key="up_nuevos",
        label_visibility="collapsed",
    )
    st.caption(f"Se guardara como: `{ARCHIVOS['nuevos']}`")

with col2:
    st.markdown("**Tickets en Revision**")
    file_revision = st.file_uploader(
        "Archivo con los tickets en estado Revision",
        type=["xlsx"],
        key="up_revision",
        label_visibility="collapsed",
    )
    st.caption(f"Se guardara como: `{ARCHIVOS['revision']}`")

with col3:
    st.markdown("**Tickets Devueltos**")
    file_devuelto = st.file_uploader(
        "Archivo con los tickets en estado Devuelto a Servicios",
        type=["xlsx"],
        key="up_devuelto",
        label_visibility="collapsed",
    )
    st.caption(f"Se guardara como: `{ARCHIVOS['devuelto']}`")


# ---------------------------------------------------------------------------
# 2) Ejecucion del monitor
# ---------------------------------------------------------------------------
st.header("2. Ejecutar monitoreo")

listo = all([file_nuevos, file_revision, file_devuelto])

if not listo:
    st.info(
        "Sube los 3 archivos para habilitar el boton de ejecucion. "
        "El orden no importa: cada uploader corresponde a un estado."
    )

col_btn, col_info = st.columns([1, 3])
with col_btn:
    ejecutar = st.button(
        "▶ Ejecutar monitor",
        type="primary",
        disabled=not listo,
        use_container_width=True,
    )
with col_info:
    if listo:
        st.success("Archivos listos. Presiona el boton para procesar.")


def _guardar_archivo(uploaded_file, nombre_destino: str) -> Path:
    """Guarda el archivo subido en la carpeta 'entrada/' con el nombre
    que espera monitor.py (independientemente del nombre original)."""
    destino = ENTRADA_DIR / nombre_destino
    with open(destino, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return destino


def _ultimo_reporte() -> Path | None:
    """Devuelve el reporte mas reciente generado en 'reportes/', o None."""
    if not REPORTES_DIR.exists():
        return None
    reportes = sorted(
        REPORTES_DIR.glob("Reporte_Monitoreo_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return reportes[0] if reportes else None


if ejecutar and listo:
    # --- Renombrado (guardado con nombre canonico) ---
    with st.status("Preparando archivos...", expanded=True) as status:
        mapeo = [
            (file_nuevos,   ARCHIVOS["nuevos"]),
            (file_revision, ARCHIVOS["revision"]),
            (file_devuelto, ARCHIVOS["devuelto"]),
        ]
        for uf, nombre in mapeo:
            ruta = _guardar_archivo(uf, nombre)
            st.write(f"✔ `{uf.name}` → `{ruta.name}`")

        # --- Ejecucion del monitor como subproceso ---
        # Usamos el mismo interprete de Python que corre Streamlit para
        # asegurar que ve las mismas dependencias instaladas.
        status.update(label="Ejecutando monitor...", state="running")
        try:
            proc = subprocess.run(
                [sys.executable, "monitor.py"],
                cwd=str(Path(__file__).resolve().parent),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=300,
            )
        except subprocess.TimeoutExpired:
            status.update(label="Tiempo de ejecucion agotado", state="error")
            st.error("El monitor tardo mas de 5 minutos. Revisa los archivos.")
            st.stop()

        if proc.returncode != 0:
            status.update(label="Error al ejecutar el monitor", state="error")
            st.error("El monitor termino con errores. Revisa la salida abajo.")
        else:
            status.update(label="Monitoreo completado", state="complete")

        with st.expander("Ver salida del monitor", expanded=(proc.returncode != 0)):
            if proc.stdout:
                st.code(proc.stdout, language="text")
            if proc.stderr:
                st.markdown("**stderr:**")
                st.code(proc.stderr, language="text")

    # --- Descarga del reporte generado ---
    if proc.returncode == 0:
        ultimo = _ultimo_reporte()
        if ultimo is not None:
            st.success(f"Reporte generado: `{ultimo.name}`")
            with open(ultimo, "rb") as f:
                st.download_button(
                    label="⬇ Descargar reporte Excel",
                    data=f.read(),
                    file_name=ultimo.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
        else:
            st.warning("No se encontro reporte generado en la carpeta 'reportes/'.")


# ---------------------------------------------------------------------------
# 3) Informacion adicional
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Estado del entorno")
    st.markdown(f"**Carpeta base:** `{Path(__file__).resolve().parent}`")
    st.markdown(f"**Entrada:** `{ENTRADA_DIR}`")
    st.markdown(f"**Reportes:** `{REPORTES_DIR}`")
    st.markdown(f"**Historico:** `{HISTORICO_DIR}`")

    st.divider()
    st.subheader("Ultimos reportes")
    if REPORTES_DIR.exists():
        reportes = sorted(
            REPORTES_DIR.glob("Reporte_Monitoreo_*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )[:10]
        if reportes:
            for r in reportes:
                st.markdown(f"- `{r.name}`")
        else:
            st.caption("Aun no hay reportes generados.")
    else:
        st.caption("La carpeta de reportes aun no existe.")

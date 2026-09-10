"""
monitor.py - Monitoreo diario de tickets de Zoho Analytics
============================================================

Compara los tickets de HOY vs AYER a partir de 3 archivos Excel
descargados de Zoho Analytics y genera un reporte con el movimiento
de tickets entre estados ("nuevo/abierto" -> "revision" / "devuelto").

Estructura real de los archivos de Zoho (196 columnas), de las cuales
usamos las siguientes:

    ID de la solicitud                -> numero de ticket (ej: 40488)  [clave de comparacion]
    Asunto                            -> descripcion corta
    Estado                            -> Nuevo / Abierto / Revision / Devuelto
    Prioridad                         -> Alto / Medio / Bajo
    Responsable FDS                   -> nombre del asignado (ej: "Diana Salazar")
    Hora de creacion                  -> cuando entro el ticket como nuevo
    Hora de actualizacion del estado  -> cuando cambio al estado actual

Uso:
    python monitor.py
"""

# ---------------------------------------------------------------------------
# PASO 0: Importar librerias
# ---------------------------------------------------------------------------
# pandas    -> para leer/manipular tablas de Excel
# pathlib   -> para manejar rutas de archivos de forma portable
# datetime  -> para manejar fechas (hoy, ayer)
# shutil    -> para copiar archivos al historico
# xlsxwriter (via pandas ExcelWriter) -> para escribir el reporte con graficos
from pathlib import Path
from datetime import date, datetime, timedelta
import shutil
import sys

import pandas as pd


# ---------------------------------------------------------------------------
# PASO 1: Configuracion (el usuario final NO necesita tocar nada aqui)
# ---------------------------------------------------------------------------
BASE_DIR      = Path(__file__).resolve().parent
ENTRADA_DIR   = BASE_DIR / "entrada"
HISTORICO_DIR = BASE_DIR / "historico"
REPORTES_DIR  = BASE_DIR / "reportes"

# Archivo CSV que acumula, dia a dia, los conteos de movimientos.
# Sirve para ver la evolucion en el tiempo (grafico de lineas).
HISTORIAL_CSV = HISTORICO_DIR / "historial_movimientos.csv"

# Nombres esperados de los archivos de entrada (los que descarga el usuario)
ARCHIVOS = {
    "nuevos":   "tickets_nuevos.xlsx",
    "revision": "tickets_revision.xlsx",
    "devuelto": "tickets_devuelto.xlsx",
}

# Columna que identifica de forma UNICA a cada ticket (numero de ticket
# legible que asigna Zoho, ej: 40488). Este es el "numero del ticket"
# que ve el usuario final.
COL_ID = "ID de la solicitud"

# Columnas legibles que queremos mostrar en el reporte (si existen).
# Se toman directamente del Excel de Zoho y se renombran a etiquetas
# amigables mediante RENOMBRAR_COLUMNAS al momento de armar el reporte.
COLUMNAS_REPORTE = [
    "ID de la solicitud",
    "Asunto",
    "Estado",
    "Prioridad",
    "Nombre completo",
    "Responsable FDS",
    "Estado FDS",
    "Fecha Inicio FDS",
    "Fecha Fin FDS",
    "Hora de creación",
    "Hora de actualización del estado",
]

# Nombre visible en el reporte para cada columna del Excel original.
RENOMBRAR_COLUMNAS = {
    "Responsable FDS":                  "Asignado a",
    "Estado FDS":                       "Estado FDS",
    "Nombre completo":                  "Propietario",
    "Hora de creación":                 "Fecha de ingreso",
    "Hora de actualización del estado": "Fecha cambio de estado",
}

# Nombres de columnas del Excel de Zoho que usamos en el Resumen para saber
# quien esta desarrollando el ticket y en que estado FDS se encuentra.
COL_RESP_FDS = "Responsable FDS"
COL_EST_FDS  = "Estado FDS"


# ---------------------------------------------------------------------------
# PASO 2: Utilidades de archivos
# ---------------------------------------------------------------------------
def asegurar_carpetas():
    """Crea las carpetas de trabajo si aun no existen."""
    for d in (ENTRADA_DIR, HISTORICO_DIR, REPORTES_DIR):
        d.mkdir(parents=True, exist_ok=True)


def leer_excel(path: Path) -> pd.DataFrame:
    """Lee un archivo Excel de Zoho.

    Devuelve un DataFrame vacio si el archivo no existe. Esto permite que
    el primer dia (cuando aun no hay 'ayer' en historico) todo siga funcionando.
    """
    if not path.exists():
        return pd.DataFrame(columns=[COL_ID])
    # IMPORTANTE: leemos el ID como TEXTO. Los IDs de Zoho tienen 18 digitos
    # (ej: 937604000041682560) y no caben con precision en un float64.
    df = pd.read_excel(path, dtype={COL_ID: str})
    if COL_ID not in df.columns:
        raise ValueError(
            f"El archivo {path.name} no tiene la columna '{COL_ID}'. "
            f"Columnas encontradas: {list(df.columns)[:10]}..."
        )
    df[COL_ID] = df[COL_ID].astype(str).str.strip()
    return df


def ids_de(df: pd.DataFrame) -> set:
    """Extrae el conjunto de IDs de un DataFrame."""
    if df.empty or COL_ID not in df.columns:
        return set()
    return set(df[COL_ID].dropna().tolist())


# ---------------------------------------------------------------------------
# Gestion de SNAPSHOTS (una carpeta por ejecucion, no por dia)
# ---------------------------------------------------------------------------
# Cada vez que ejecutas el monitor se crea un snapshot con timestamp:
#     historico/2026-08-27_14-30-15/
# Esto permite hacer varias comparaciones el mismo dia (ej: al inicio y al
# final de la jornada) y comparar siempre contra el snapshot mas reciente
# anterior, sea del mismo dia o de dias previos.

FORMATO_SNAPSHOT = "%Y-%m-%d_%H-%M-%S"


def carpeta_snapshot(ts: datetime) -> Path:
    return HISTORICO_DIR / ts.strftime(FORMATO_SNAPSHOT)


def listar_snapshots() -> list[tuple[datetime, Path]]:
    """Devuelve todos los snapshots existentes ordenados de mas antiguo a
    mas reciente. Acepta tanto el formato nuevo con hora
    (AAAA-MM-DD_HH-MM-SS) como el antiguo solo con fecha (AAAA-MM-DD),
    para no romper historicos previos.
    """
    snaps: list[tuple[datetime, Path]] = []
    if not HISTORICO_DIR.exists():
        return snaps
    for d in HISTORICO_DIR.iterdir():
        if not d.is_dir():
            continue
        for fmt in (FORMATO_SNAPSHOT, "%Y-%m-%d"):
            try:
                ts = datetime.strptime(d.name, fmt)
                # Formato antiguo (solo fecha) -> lo colocamos al final del dia
                if fmt == "%Y-%m-%d":
                    ts = ts.replace(hour=23, minute=59, second=59)
                snaps.append((ts, d))
                break
            except ValueError:
                continue
    return sorted(snaps, key=lambda x: x[0])


def snapshot_anterior(ahora: datetime) -> tuple[datetime, Path] | None:
    """Devuelve (timestamp, carpeta) del snapshot mas reciente estrictamente
    anterior a 'ahora'. None si no hay ninguno previo (primera ejecucion)."""
    previos = [(t, p) for t, p in listar_snapshots() if t < ahora]
    return previos[-1] if previos else None


def clasificar_intervalo(actual: datetime, anterior: datetime) -> tuple[str, str]:
    """Devuelve (etiqueta_corta, descripcion) del intervalo entre dos snapshots.

    Ejemplos:
        (mismo dia -> "Mismo dia", "hace 5h 20min")
        (dia siguiente -> "Dia siguiente", "hace 1 dia 3h")
        (varios dias -> "Hace N dias", "hace 3 dias 12h")
    """
    delta = actual - anterior
    segundos = int(delta.total_seconds())
    dias  = segundos // 86400
    horas = (segundos % 86400) // 3600
    mins  = (segundos % 3600) // 60

    def fmt():
        if dias == 0 and horas == 0: return f"hace {mins} min"
        if dias == 0:                return f"hace {horas}h {mins:02d}min"
        if horas == 0:               return f"hace {dias} dia" + ("s" if dias != 1 else "")
        return f"hace {dias} dia" + ("s" if dias != 1 else "") + f" {horas}h"

    if actual.date() == anterior.date():
        etiqueta = "Mismo dia"
    elif (actual.date() - anterior.date()).days == 1:
        etiqueta = "Dia siguiente"
    else:
        etiqueta = f"Hace {(actual.date() - anterior.date()).days} dias"
    return etiqueta, fmt()


def actualizar_historial(datos: dict) -> pd.DataFrame:
    """Anade una fila al historial acumulado (una fila POR EJECUCION).

    Si en la misma marca de tiempo ya existiera un registro (mismo segundo),
    se reemplaza. Asi cada corrida queda registrada por separado.

    Estructura del CSV:
        fecha_hora, fecha_hora_anterior, tipo_comparacion, intervalo,
        total_nuevos, total_revision, total_devuelto,
        pasaron_a_revision, pasaron_a_devuelto, otros_estados, nuevos_ingresados
    """
    fila = {
        "fecha_hora":          datos["ts_actual"].strftime(FORMATO_SNAPSHOT),
        "fecha_hora_anterior": (datos["ts_anterior"].strftime(FORMATO_SNAPSHOT)
                                if datos["ts_anterior"] else ""),
        "tipo_comparacion":    datos["tipo_comparacion"],
        "intervalo":           datos["intervalo_desc"],
        # Excluye "repetidos" (Estado FDS = Devuelto) para que el historial
        # solo cuente los nuevos/abiertos realmente activos.
        "total_nuevos":        datos["total_nuevos_hoy_activos"],
        "total_revision":      datos["total_revision_hoy"],
        "total_devuelto":      datos["total_devuelto_hoy"],
        "pasaron_a_revision":  len(datos["ids_movidos_revision"]),
        "pasaron_a_devuelto":  len(datos["ids_movidos_devuelto"]),
        "otros_estados":       len(datos["ids_otros_estados"]),
        "nuevos_ingresados":   len(datos["ids_nuevos_ingresados"]),
        "reaperturas":         len(datos["ids_reaperturas"]),
        "salieron_de_revision": len(datos["ids_salieron_revision"]),
        "salieron_de_devuelto": len(datos["ids_salieron_devuelto"]),
    }
    if HISTORIAL_CSV.exists():
        hist = pd.read_csv(HISTORIAL_CSV)
        # Compatibilidad: si el CSV viejo tenia columna 'fecha' la renombramos.
        if "fecha" in hist.columns and "fecha_hora" not in hist.columns:
            hist = hist.rename(columns={"fecha": "fecha_hora"})
        hist = hist[hist["fecha_hora"] != fila["fecha_hora"]]
        hist = pd.concat([hist, pd.DataFrame([fila])], ignore_index=True)
    else:
        hist = pd.DataFrame([fila])
    hist = hist.sort_values("fecha_hora").reset_index(drop=True)
    HISTORICO_DIR.mkdir(parents=True, exist_ok=True)
    hist.to_csv(HISTORIAL_CSV, index=False)
    return hist


def archivar_snapshot(ts: datetime):
    """Copia los 3 Excel de 'entrada/' a 'historico/AAAA-MM-DD_HH-MM-SS/'.

    Cada ejecucion crea su propio snapshot, permitiendo comparaciones
    varias veces al mismo dia.
    """
    destino = carpeta_snapshot(ts)
    destino.mkdir(parents=True, exist_ok=True)
    for nombre in ARCHIVOS.values():
        origen = ENTRADA_DIR / nombre
        if origen.exists():
            shutil.copy2(origen, destino / nombre)


# ---------------------------------------------------------------------------
# PASO 3: Preparacion de datos para el reporte
# ---------------------------------------------------------------------------
def subset(df: pd.DataFrame, ids: set,
           estado_anterior: str | None = None,
           estado_anterior_por_id: dict | None = None) -> pd.DataFrame:
    """Devuelve solo las filas del DataFrame cuyos IDs estan en 'ids',
    quedandose unicamente con las COLUMNAS_REPORTE que existan.

    - 'estado_anterior' (str): agrega columna con el mismo valor para todos.
    - 'estado_anterior_por_id' (dict): agrega columna cuyo valor depende
      del ID de cada fila. Util para 'Reaperturas' donde unos venian
      de 'Revision' y otros de 'Devuelto a Servicios'.
    """
    cols = [c for c in COLUMNAS_REPORTE if c in df.columns]
    if df.empty or not ids:
        out = pd.DataFrame(columns=cols)
    else:
        out = df[df[COL_ID].isin(ids)][cols].reset_index(drop=True)
    if estado_anterior_por_id is not None and COL_ID in out.columns:
        out.insert(1, "Estado anterior",
                   out[COL_ID].map(estado_anterior_por_id).fillna("Desconocido"))
    elif estado_anterior is not None:
        out.insert(1 if COL_ID in out.columns else 0, "Estado anterior", estado_anterior)
    # Renombramos a etiquetas amigables (Asignado a, Fecha de ingreso, etc.)
    out = out.rename(columns=RENOMBRAR_COLUMNAS)
    return out


# ---------------------------------------------------------------------------
# PASO 4: Comparacion HOY vs AYER
# ---------------------------------------------------------------------------
def comparar(ts_actual: datetime) -> dict:
    """Calcula que tickets desaparecieron de 'nuevos' y a donde se movieron,
    comparando la entrada actual contra el snapshot anterior mas reciente
    (sea del mismo dia o de dias previos).
    """
    # --- Archivos ACTUALES (recien descargados a 'entrada/') ---
    df_hoy_nuevos = leer_excel(ENTRADA_DIR / ARCHIVOS["nuevos"])
    df_hoy_revi   = leer_excel(ENTRADA_DIR / ARCHIVOS["revision"])
    df_hoy_devu   = leer_excel(ENTRADA_DIR / ARCHIVOS["devuelto"])

    # --- Snapshot ANTERIOR (el mas reciente estrictamente < ts_actual) ---
    prev = snapshot_anterior(ts_actual)
    if prev is None:
        ts_anterior, ant_dir = None, None
        etiqueta_tipo, intervalo_desc = "Primer informe", "-"
        df_ayer_nuevos = leer_excel(Path("__no_existe__"))
        df_ayer_revi   = leer_excel(Path("__no_existe__"))
        df_ayer_devu   = leer_excel(Path("__no_existe__"))
    else:
        ts_anterior, ant_dir = prev
        etiqueta_tipo, intervalo_desc = clasificar_intervalo(ts_actual, ts_anterior)
        df_ayer_nuevos = leer_excel(ant_dir / ARCHIVOS["nuevos"])
        df_ayer_revi   = leer_excel(ant_dir / ARCHIVOS["revision"])
        df_ayer_devu   = leer_excel(ant_dir / ARCHIVOS["devuelto"])

    hoy_nuevos_ids  = ids_de(df_hoy_nuevos)
    hoy_revi_ids    = ids_de(df_hoy_revi)
    hoy_devu_ids    = ids_de(df_hoy_devu)
    ayer_nuevos_ids = ids_de(df_ayer_nuevos)
    ayer_revi_ids   = ids_de(df_ayer_revi)
    ayer_devu_ids   = ids_de(df_ayer_devu)

    # --- LOGICA PRINCIPAL ---
    # 1) Un ticket "cambio de estado" si estaba en NUEVOS ayer y ya no esta hoy.
    desaparecidos = ayer_nuevos_ids - hoy_nuevos_ids

    # 2) De esos desaparecidos, cuales aparecieron ahora en REVISION o DEVUELTO
    #    y NO estaban ya ayer en esa lista (para contar solo movimientos nuevos).
    movidos_a_revision = desaparecidos & (hoy_revi_ids - ayer_revi_ids)
    movidos_a_devuelto = desaparecidos & (hoy_devu_ids - ayer_devu_ids)

    # 3) Desaparecidos que no aparecen en ninguna lista monitoreada
    #    (probablemente cerrados, resueltos u otro estado no seguido).
    otros_estados = desaparecidos - movidos_a_revision - movidos_a_devuelto

    # 4) Tickets que aparecen en 'nuevos' HOY, separando:
    #    - REAPERTURAS: ya existian ayer pero en 'revision' o 'devuelto'
    #      (volvieron al estado nuevo/abierto).
    #    - NUEVOS INGRESADOS: no aparecian en NINGUN archivo ayer,
    #      son tickets verdaderamente nuevos.
    reaperturas       = hoy_nuevos_ids & (ayer_revi_ids | ayer_devu_ids)
    ayer_todos_ids    = ayer_nuevos_ids | ayer_revi_ids | ayer_devu_ids
    hoy_todos_ids     = hoy_nuevos_ids | hoy_revi_ids | hoy_devu_ids
    nuevos_ingresados = hoy_nuevos_ids - ayer_todos_ids

    # 5) Tickets que ya no aparecen en REVISION hoy y tampoco estan en
    #    ninguno de los otros 2 archivos = salieron a "otros estados"
    #    (probablemente cerrados/resueltos desde revision).
    salieron_de_revision = ayer_revi_ids - hoy_todos_ids

    # 6) Igual pero desde DEVUELTO A SERVICIOS.
    salieron_de_devuelto = ayer_devu_ids - hoy_todos_ids

    # 7) Tickets "excluidos": estan en Nuevos/Abiertos pero su columna
    #    "Estado FDS" tiene un valor que NO representa trabajo activo de
    #    desarrollo (Devuelto, Backlog Robot, Certificacion, Done). No
    #    deben sumar en Nuevos/Abiertos, se muestran en una hoja aparte.
    COL_EST_FDS_LOC = "Estado FDS"
    ESTADOS_FDS_EXCLUIR = {
        "devuelto",
        "backlog robot",
        "certificacion", "certificación",
        "done",
    }

    def _excluidos_fds(df: pd.DataFrame, todos_ids: set) -> set:
        """IDs cuyo 'Estado FDS' esta en ESTADOS_FDS_EXCLUIR (insensible a
        mayusculas y espacios). Solo dentro del universo 'todos_ids'."""
        if df.empty or COL_EST_FDS_LOC not in df.columns:
            return set()
        est = df[COL_EST_FDS_LOC].astype(str).str.strip().str.lower()
        ids = set(df.loc[est.isin(ESTADOS_FDS_EXCLUIR), COL_ID].astype(str))
        return ids & todos_ids

    hoy_excluidos_ids  = _excluidos_fds(df_hoy_nuevos,  hoy_nuevos_ids)
    ayer_excluidos_ids = _excluidos_fds(df_ayer_nuevos, ayer_nuevos_ids)

    # Los nuevos "activos" son los que quedan tras excluir los anteriores.
    hoy_nuevos_activos_ids  = hoy_nuevos_ids  - hoy_excluidos_ids
    ayer_nuevos_activos_ids = ayer_nuevos_ids - ayer_excluidos_ids

    return {
        # Marcas de tiempo del snapshot actual y del que se comparo
        "ts_actual":         ts_actual,
        "ts_anterior":       ts_anterior,
        "tipo_comparacion":  etiqueta_tipo,
        "intervalo_desc":    intervalo_desc,
        # Retrocompatibilidad con la logica anterior (algunos usos siguen
        # llamandolo "hubo_ayer")
        "fecha_hoy":  ts_actual.date(),
        "fecha_ayer": ts_anterior.date() if ts_anterior else None,
        "hubo_ayer":  ts_anterior is not None and bool(ayer_nuevos_ids or ayer_revi_ids or ayer_devu_ids),

        # Conteos
        "total_nuevos_hoy":    len(hoy_nuevos_ids),
        "total_nuevos_ayer":   len(ayer_nuevos_ids),
        "total_revision_hoy":  len(hoy_revi_ids),
        "total_revision_ayer": len(ayer_revi_ids),
        "total_devuelto_hoy":  len(hoy_devu_ids),
        "total_devuelto_ayer": len(ayer_devu_ids),

        # Sub-DataFrames con columnas legibles para el reporte:
        # - Los MOVIDOS los buscamos en el archivo DESTINO de HOY, porque
        #   alli aparecen ahora con toda la informacion actualizada.
        # - Los "otros_estados" los buscamos en el archivo de nuevos de AYER
        #   (ya no estan en ninguno de los 3 archivos de hoy).
        # - Los "nuevos_ingresados" los buscamos en el archivo de nuevos de HOY.
        "df_movidos_revision":  subset(df_hoy_revi,    movidos_a_revision, estado_anterior="Nuevo/Abierto"),
        "df_movidos_devuelto":  subset(df_hoy_devu,    movidos_a_devuelto, estado_anterior="Nuevo/Abierto"),
        "df_otros_estados":     subset(df_ayer_nuevos, otros_estados,      estado_anterior="Nuevo/Abierto"),
        "df_nuevos_ingresados": subset(df_hoy_nuevos,  nuevos_ingresados),
        # Reaperturas: para cada ID indicamos si venia de "Revision" o "Devuelto"
        "df_reaperturas":       subset(
            df_hoy_nuevos, reaperturas,
            estado_anterior_por_id={
                **{tid: "Revision"             for tid in ayer_revi_ids},
                **{tid: "Devuelto a Servicios" for tid in ayer_devu_ids},
            },
        ),
        # Salieron a otros estados desde REVISION / DEVUELTO
        # (probablemente fueron cerrados o resueltos).
        "df_salieron_revision": subset(df_ayer_revi, salieron_de_revision, estado_anterior="Revision"),
        "df_salieron_devuelto": subset(df_ayer_devu, salieron_de_devuelto, estado_anterior="Devuelto a Servicios"),

        # Listados COMPLETOS de HOY por estado (para hojas navegables desde
        # los numeros del resumen). "df_todos_nuevos_hoy" excluye los
        # tickets "repetidos" (Estado FDS = Devuelto) que van en su propia hoja.
        "df_todos_nuevos_hoy":   subset(df_hoy_nuevos, hoy_nuevos_activos_ids),
        "df_todos_revision_hoy": subset(df_hoy_revi,   hoy_revi_ids),
        "df_todos_devuelto_hoy": subset(df_hoy_devu,   hoy_devu_ids),

        # Hoja aparte: tickets nuevos cuya columna "Estado FDS" es uno
        # de: Devuelto, Backlog Robot, Certificacion, Done. No representan
        # trabajo activo de desarrollo, por eso no cuentan en Nuevos.
        "df_tickets_excluidos":         subset(df_hoy_nuevos, hoy_excluidos_ids),
        "total_tickets_excluidos_hoy":  len(hoy_excluidos_ids),
        "total_tickets_excluidos_ayer": len(ayer_excluidos_ids),
        "total_nuevos_hoy_activos":     len(hoy_nuevos_activos_ids),
        "total_nuevos_ayer_activos":    len(ayer_nuevos_activos_ids),

        # Tambien devolvemos los IDs por si se necesitan
        "ids_movidos_revision":  sorted(movidos_a_revision),
        "ids_movidos_devuelto":  sorted(movidos_a_devuelto),
        "ids_otros_estados":     sorted(otros_estados),
        "ids_nuevos_ingresados": sorted(nuevos_ingresados),
        "ids_reaperturas":       sorted(reaperturas),
        "ids_salieron_revision": sorted(salieron_de_revision),
        "ids_salieron_devuelto": sorted(salieron_de_devuelto),
    }


# ---------------------------------------------------------------------------
# PASO 5: Generar el reporte Excel con tablas y graficos
# ---------------------------------------------------------------------------
def _ruta_reporte_libre(ts: datetime) -> Path:
    """Devuelve una ruta unica para el reporte usando timestamp completo.
    Si por alguna razon ya existe (y esta abierta en Excel), agrega sufijo.
    """
    base = REPORTES_DIR / f"Reporte_Monitoreo_{ts.strftime(FORMATO_SNAPSHOT)}.xlsx"
    if not base.exists():
        return base
    try:
        with open(base, "a"):
            pass
        return base
    except PermissionError:
        i = 2
        while True:
            alt = REPORTES_DIR / f"Reporte_Monitoreo_{ts.strftime(FORMATO_SNAPSHOT)}_v{i}.xlsx"
            if not alt.exists():
                print(f"[AVISO] '{base.name}' esta abierto en Excel; guardando como '{alt.name}'.")
                return alt
            i += 1


def generar_reporte(datos: dict, historial: pd.DataFrame) -> Path:
    ts_actual = datos["ts_actual"]
    salida = _ruta_reporte_libre(ts_actual)

    # xlsxwriter permite insertar graficos nativos de Excel.
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        wb = writer.book

        # -------- Formatos visuales --------
        f_titulo = wb.add_format({
            "bold": True, "font_size": 14, "font_color": "white",
            "bg_color": "#1F4E78", "align": "center", "valign": "vcenter",
            "border": 1,
        })
        f_header = wb.add_format({
            "bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "center",
        })
        f_cell   = wb.add_format({"border": 1})
        f_num    = wb.add_format({"border": 1, "align": "center"})
        f_ok     = wb.add_format({"border": 1, "bg_color": "#C6EFCE", "align": "center"})
        f_warn   = wb.add_format({"border": 1, "bg_color": "#FFEB9C", "align": "center"})
        f_fecha  = wb.add_format({"border": 1, "num_format": "yyyy-mm-dd hh:mm"})
        # Formato para celdas que son hipervinculos internos (estilo "link").
        f_link   = wb.add_format({
            "border": 1, "font_color": "#0563C1", "underline": 1,
        })

        # =====================================================================
        # HOJA 1: Resumen
        # =====================================================================
        ws = wb.add_worksheet("Resumen")
        writer.sheets["Resumen"] = ws
        ws.set_column("A:A", 34)
        ws.set_column("B:D", 14)

        titulo_reporte = (
            f"Reporte de Monitoreo de Tickets - "
            f"{ts_actual.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        ws.merge_range("A1:E1", titulo_reporte, f_titulo)
        ws.set_row(0, 26)
        # Enlace al historial diario acumulado (visible siempre).
        ws.write_url("F1", "internal:'Historial_Diario'!A1", f_link,
                     "Ver historial diario >>")

        # --- Bloque "COMPARACION: actual vs anterior" ---
        # Deja explicito si es una comparacion "mismo dia" o "dia siguiente".
        f_tipo_mismo = wb.add_format({
            "bold": True, "border": 1, "align": "center",
            "bg_color": "#FFE699", "font_color": "#7F6000",  # amarillo
        })
        f_tipo_sig   = wb.add_format({
            "bold": True, "border": 1, "align": "center",
            "bg_color": "#C6EFCE", "font_color": "#375623",  # verde
        })
        f_tipo_otro  = wb.add_format({
            "bold": True, "border": 1, "align": "center",
            "bg_color": "#D9E1F2", "font_color": "#1F4E78",  # azul
        })
        if datos["ts_anterior"] is None:
            f_tipo = f_tipo_otro
        elif datos["tipo_comparacion"] == "Mismo dia":
            f_tipo = f_tipo_mismo
        elif datos["tipo_comparacion"] == "Dia siguiente":
            f_tipo = f_tipo_sig
        else:
            f_tipo = f_tipo_otro

        ts_ant_str = (datos["ts_anterior"].strftime("%Y-%m-%d %H:%M:%S")
                      if datos["ts_anterior"] else "(no hay snapshot previo)")
        ws.write("A2", "Comparando:", f_header)
        ws.merge_range("B2:C2", ts_ant_str + "  -->  " + ts_actual.strftime("%Y-%m-%d %H:%M:%S"), f_cell)
        ws.write("D2", datos["tipo_comparacion"], f_tipo)
        ws.write("E2", datos["intervalo_desc"], f_cell)
        ws.set_row(1, 20)

        primer_dia = not datos["hubo_ayer"]

        if primer_dia:
            ws.merge_range("A3:E3",
                "Primer dia de monitoreo: aun no hay datos de AYER para comparar. "
                "Vuelve a ejecutar mañana.", f_warn)
            # Totales clickeables (llevan a la hoja con la lista completa)
            ws.write_url("A5", "internal:'Todos_Nuevos_Hoy'!A1",   f_link, "Total NUEVOS/ABIERTOS hoy (activos):")
            ws.write("B5", datos["total_nuevos_hoy_activos"], f_num)
            ws.write_url("A6", "internal:'Todos_Revision_Hoy'!A1", f_link, "Total en REVISION hoy:")
            ws.write("B6", datos["total_revision_hoy"], f_num)
            ws.write_url("A7", "internal:'Todos_Devuelto_Hoy'!A1", f_link, "Total DEVUELTOS hoy:")
            ws.write("B7", datos["total_devuelto_hoy"], f_num)
            ws.write_url("A8", "internal:'Excluidos_de_Nuevos'!A1", f_link, "Excluidos de Nuevos (Devuelto / Backlog Robot / Certificacion / Done):")
            ws.write("B8", datos["total_tickets_excluidos_hoy"], f_num)

        if not primer_dia:
            # --- Tabla comparativa HOY vs AYER ---
            # La columna "Estado" es un HIPERVINCULO a la hoja con todos los
            # tickets de ese estado de HOY.
            headers = ["Estado", "Informe anterior", "Informe actual", "Diferencia"]
            for i, h in enumerate(headers):
                ws.write(2, i, h, f_header)

            filas = [
                ("Nuevos / Abiertos (activos)",
                    datos["total_nuevos_ayer_activos"], datos["total_nuevos_hoy_activos"], "Todos_Nuevos_Hoy"),
                ("En revision",
                    datos["total_revision_ayer"],      datos["total_revision_hoy"],       "Todos_Revision_Hoy"),
                ("Devueltos",
                    datos["total_devuelto_ayer"],      datos["total_devuelto_hoy"],       "Todos_Devuelto_Hoy"),
                ("Excluidos de Nuevos (Devuelto / Backlog Robot / Certificacion / Done)",
                    datos["total_tickets_excluidos_ayer"], datos["total_tickets_excluidos_hoy"], "Excluidos_de_Nuevos"),
            ]
            for i, (estado, ayer_v, hoy_v, hoja) in enumerate(filas, start=3):
                diff = hoy_v - ayer_v
                ws.write_url(i, 0, f"internal:'{hoja}'!A1", f_link, estado)
                ws.write(i, 1, ayer_v, f_num)
                ws.write(i, 2, hoy_v,  f_num)
                # Verde si baja o queda igual (bueno para "nuevos"), amarillo si sube.
                ws.write(i, 3, diff, f_ok if diff <= 0 else f_warn)

            # --- Tabla de movimientos ---
            # Ancho especial para la columna de "Numeros de ticket".
            ws.set_column("E:E", 55)
            ws.merge_range("A8:E8", "Movimiento desde 'Nuevos/Abiertos'", f_titulo)
            ws.set_row(7, 22)

            def _nums(df: pd.DataFrame) -> str:
                """Devuelve, uno por linea, cada ticket con su responsable y
                estado FDS actuales para mostrarlos directo en el Resumen:
                    <ID>  |  <Asignado a>  |  <Estado FDS>
                Asi se ve rapido quien esta desarrollando cada ticket y en
                que estado se encuentra sin necesidad de abrir la hoja de
                detalle."""
                if df.empty or COL_ID not in df.columns:
                    return ""
                # Nombres finales (post-rename hecho en subset()).
                col_resp = RENOMBRAR_COLUMNAS.get(COL_RESP_FDS, COL_RESP_FDS)
                col_est  = RENOMBRAR_COLUMNAS.get(COL_EST_FDS,  COL_EST_FDS)
                lineas = []
                for _, row in df.iterrows():
                    tid  = str(row.get(COL_ID, "")).strip()
                    resp = str(row.get(col_resp, "") or "").strip() or "(sin asignar)"
                    est  = str(row.get(col_est,  "") or "").strip() or "(sin estado FDS)"
                    lineas.append(f"{tid}  |  {resp}  |  {est}")
                return "\n".join(lineas)

            # La columna "Concepto" es un HIPERVINCULO a la hoja de detalle
            # correspondiente. Ademas mostramos los numeros de ticket
            # directamente en el resumen para que se vean sin hacer click.
            mov_filas = [
                # -- Desde Nuevos/Abiertos --
                ("Pasaron a REVISION",                    datos["df_movidos_revision"],  "Movidos_a_Revision"),
                ("Pasaron a DEVUELTO",                    datos["df_movidos_devuelto"],  "Movidos_a_Devuelto"),
                ("Cambiaron de propietario (distinto a FDS)", datos["df_otros_estados"],     "Cambio_Propietario_No_FDS"),
                # -- Ingresos y reaperturas --
                ("Nuevos ingresados HOY",                 datos["df_nuevos_ingresados"], "Nuevos_Ingresados"),
                ("Reaperturas (volvieron a Nuevo)",       datos["df_reaperturas"],       "Reaperturas"),
                # -- Desde Revision --
                ("Salieron de REVISION (otros estados)",  datos["df_salieron_revision"], "Salieron_de_Revision"),
                # -- Desde Devuelto --
                ("Salieron de DEVUELTO (otros estados)",  datos["df_salieron_devuelto"], "Salieron_de_Devuelto"),
            ]
            f_wrap = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
            ws.write(8, 0, "Concepto (click para ver detalle)", f_header)
            ws.write(8, 1, "Cantidad", f_header)
            ws.merge_range(8, 2, 8, 4,
                           "Tickets  |  Responsable FDS  |  Estado FDS",
                           f_header)
            # Ampliamos aun mas la columna de detalle para que quepan las 3
            # piezas de informacion por ticket.
            ws.set_column("C:E", 30)
            for i, (concepto, df_mov, hoja) in enumerate(mov_filas, start=9):
                ws.write_url(i, 0, f"internal:'{hoja}'!A1", f_link, concepto)
                ws.write(i, 1, len(df_mov), f_num)
                texto = _nums(df_mov)
                ws.merge_range(i, 2, i, 4, texto, f_wrap)
                # Alto de fila proporcional a la cantidad de tickets listados
                # (15 pt por linea, minimo 30 pt).
                n_lineas = max(texto.count("\n") + 1, 1) if texto else 1
                ws.set_row(i, max(30, n_lineas * 15))

            # --- Tabla: tickets en desarrollo (con Estado FDS y Responsable FDS) ---
            # Se muestra debajo del bloque de movimientos. Solo aparecen los
            # tickets del sistema (nuevos + revision + devuelto de HOY) que
            # TIENEN ambos campos completos y significativos.
            col_resp = RENOMBRAR_COLUMNAS.get(COL_RESP_FDS, COL_RESP_FDS)
            col_est  = RENOMBRAR_COLUMNAS.get(COL_EST_FDS,  COL_EST_FDS)
            VALORES_VACIOS = {"", "nan", "none", "--ninguna--", "sin clasificar"}

            # Solo consideramos los tickets que HOY estan en Nuevo/Abierto
            # (no Revision ni Devuelto), segun pedido del usuario.
            partes = []
            for key, estado_actual in [
                ("df_todos_nuevos_hoy",   "Nuevo/Abierto"),
            ]:
                d = datos.get(key)
                if d is None or d.empty:
                    continue
                if col_resp not in d.columns or col_est not in d.columns:
                    continue
                tmp = d[[COL_ID, col_resp, col_est]].copy()
                tmp["__estado_actual__"] = estado_actual
                partes.append(tmp)

            if partes:
                df_dev = pd.concat(partes, ignore_index=True)

                def _limpio(v) -> str:
                    return str(v if v is not None else "").strip()

                def _valido(v) -> bool:
                    return _limpio(v).lower() not in VALORES_VACIOS

                mask = df_dev[col_resp].apply(_valido) & df_dev[col_est].apply(_valido)
                df_dev = df_dev[mask].reset_index(drop=True)
            else:
                df_dev = pd.DataFrame(columns=[COL_ID, col_resp, col_est, "__estado_actual__"])

            # Fila donde arranca este nuevo bloque (justo despues de la tabla
            # de movimientos, que ocupa filas 9..9+len(mov_filas)-1).
            fila_dev = 9 + len(mov_filas) + 1  # deja una fila en blanco
            ws.merge_range(fila_dev, 0, fila_dev, 4,
                           f"Tickets en desarrollo (con Estado FDS y Responsable FDS) - Total: {len(df_dev)}",
                           f_titulo)
            ws.set_row(fila_dev, 22)

            headers_dev = ["ID de la solicitud", "Estado actual", "Responsable FDS", "Estado FDS"]
            for j, h in enumerate(headers_dev):
                ws.write(fila_dev + 1, j, h, f_header)
            # La 5ta columna la dejamos vacia como separador visual, pero
            # ajustamos anchos para las utiles.
            ws.set_column("A:A", 34)
            ws.set_column("B:B", 22)
            ws.set_column("C:C", 28)
            ws.set_column("D:D", 30)

            if df_dev.empty:
                ws.merge_range(fila_dev + 2, 0, fila_dev + 2, 4,
                               "(no hay tickets con Estado FDS y Responsable FDS asignados)",
                               f_cell)
            else:
                for k, row in df_dev.iterrows():
                    r = fila_dev + 2 + k
                    ws.write(r, 0, str(row[COL_ID]),          f_cell)
                    ws.write(r, 1, row["__estado_actual__"],  f_cell)
                    ws.write(r, 2, str(row[col_resp]),        f_cell)
                    ws.write(r, 3, str(row[col_est]),         f_cell)

            # --- Grafico 1: comparativo Hoy vs Ayer por estado ---
            chart1 = wb.add_chart({"type": "column"})
            chart1.add_series({
                "name":       "Informe anterior",
                "categories": ["Resumen", 3, 0, 5, 0],
                "values":     ["Resumen", 3, 1, 5, 1],
                "fill":       {"color": "#8FAADC"},
            })
            chart1.add_series({
                "name":       "Informe actual",
                "categories": ["Resumen", 3, 0, 5, 0],
                "values":     ["Resumen", 3, 2, 5, 2],
                "fill":       {"color": "#1F4E78"},
            })
            chart1.set_title({"name": "Tickets por estado: informe actual vs anterior"})
            chart1.set_y_axis({"name": "Cantidad"})
            chart1.set_size({"width": 520, "height": 300})
            ws.insert_chart("F3", chart1)

            # --- Grafico 2: torta con la distribucion de movimientos ---
            # Ahora incluye 4 categorias: revision, devuelto, otros, reaperturas
            # (no incluimos "nuevos ingresados" porque no son un movimiento
            # DESDE 'nuevos', sino un ingreso al sistema).
            chart2 = wb.add_chart({"type": "pie"})
            n_mov = len(mov_filas)
            chart2.add_series({
                "name":         "Movimientos",
                "categories":   ["Resumen", 9, 0, 9 + n_mov - 1, 0],
                "values":       ["Resumen", 9, 1, 9 + n_mov - 1, 1],
                "data_labels":  {"percentage": True},
            })
            chart2.set_title({"name": "A donde se movieron los tickets 'nuevos'"})
            chart2.set_size({"width": 520, "height": 300})
            ws.insert_chart("F20", chart2)

        # =====================================================================
        # HOJAS DE DETALLE: tickets que cambiaron con TODA la informacion
        # =====================================================================
        def escribir_detalle(nombre_hoja: str, titulo: str, df: pd.DataFrame):
            """Escribe una hoja con el detalle de un conjunto de tickets."""
            ws2 = wb.add_worksheet(nombre_hoja)
            writer.sheets[nombre_hoja] = ws2

            # Titulo
            ncols = max(len(df.columns), 1)
            ws2.merge_range(0, 0, 0, ncols - 1, titulo, f_titulo)
            ws2.set_row(0, 22)

            if df.empty:
                ws2.write(2, 0, "(ninguno)", f_cell)
                return

            # Encabezados en fila 3
            for j, col in enumerate(df.columns):
                ws2.write(2, j, col, f_header)

            # Filas de datos
            for i, row in enumerate(df.itertuples(index=False), start=3):
                for j, (col, val) in enumerate(zip(df.columns, row)):
                    if pd.isna(val):
                        ws2.write(i, j, "", f_cell)
                    elif col.startswith("Fecha"):
                        # Cualquier columna "Fecha ..." se escribe como fecha real de Excel
                        try:
                            ws2.write_datetime(i, j, pd.to_datetime(val).to_pydatetime(), f_fecha)
                        except Exception:
                            ws2.write(i, j, str(val), f_cell)
                    else:
                        ws2.write(i, j, val, f_cell)

            # Anchos de columna razonables
            anchos = {
                "ID de la solicitud": 14, "Estado anterior": 16,
                "Asunto": 60, "Estado": 20, "Prioridad": 10,
                "Asignado a": 22, "Fecha de ingreso": 18,
                "Fecha cambio de estado": 20,
            }
            for j, col in enumerate(df.columns):
                ws2.set_column(j, j, anchos.get(col, 18))
            ws2.freeze_panes(3, 0)
            ws2.autofilter(2, 0, 2 + len(df), len(df.columns) - 1)

        # --- Hojas con TODOS los tickets de HOY (destino de los links del resumen)
        escribir_detalle("Todos_Nuevos_Hoy",
                         "Todos los tickets NUEVOS / ABIERTOS de HOY (activos, excluye repetidos)",
                         datos["df_todos_nuevos_hoy"])
        escribir_detalle("Todos_Revision_Hoy",
                         "Todos los tickets en REVISION de HOY",
                         datos["df_todos_revision_hoy"])
        escribir_detalle("Todos_Devuelto_Hoy",
                         "Todos los tickets DEVUELTOS a Servicios de HOY",
                         datos["df_todos_devuelto_hoy"])
        escribir_detalle("Excluidos_de_Nuevos",
                         "Tickets en Nuevos/Abiertos con Estado FDS = Devuelto, Backlog Robot, Certificacion o Done (no cuentan en Nuevos)",
                         datos["df_tickets_excluidos"])

        # =====================================================================
        # HOJA: Historial diario acumulado (evolucion de movimientos)
        # =====================================================================
        def escribir_historial(hist: pd.DataFrame):
            wsh = wb.add_worksheet("Historial_Diario")
            writer.sheets["Historial_Diario"] = wsh
            wsh.set_column("A:A", 20)   # fecha_hora
            wsh.set_column("B:B", 20)   # fecha_hora_anterior
            wsh.set_column("C:C", 15)   # tipo
            wsh.set_column("D:D", 18)   # intervalo
            wsh.set_column("E:N", 14)
            wsh.merge_range("A1:N1",
                            "Historial: comparaciones realizadas (una fila por ejecucion)",
                            f_titulo)
            wsh.set_row(0, 22)

            columnas = [
                ("fecha_hora",           "Fecha y hora"),
                ("fecha_hora_anterior",  "Comparado contra"),
                ("tipo_comparacion",     "Tipo"),
                ("intervalo",            "Intervalo"),
                ("total_nuevos",         "Nuevos/Abiertos"),
                ("total_revision",       "En revision"),
                ("total_devuelto",       "Devueltos"),
                ("pasaron_a_revision",   "Pasaron a Revision"),
                ("pasaron_a_devuelto",   "Pasaron a Devuelto"),
                ("otros_estados",        "Cambio de propietario (no FDS)"),
                ("nuevos_ingresados",    "Nuevos ingresados"),
                ("reaperturas",          "Reaperturas"),
                ("salieron_de_revision", "Salieron de Revision"),
                ("salieron_de_devuelto", "Salieron de Devuelto"),
            ]
            for j, (_, label) in enumerate(columnas):
                wsh.write(2, j, label, f_header)

            # Formato distintivo para la columna 'Tipo'
            f_mismo = wb.add_format({"border": 1, "bg_color": "#FFE699",
                                     "align": "center", "font_color": "#7F6000"})
            f_sig   = wb.add_format({"border": 1, "bg_color": "#C6EFCE",
                                     "align": "center", "font_color": "#375623"})
            f_otro  = wb.add_format({"border": 1, "bg_color": "#D9E1F2",
                                     "align": "center", "font_color": "#1F4E78"})

            for i, row in enumerate(hist.itertuples(index=False), start=3):
                d = dict(zip(hist.columns, row))
                for j, (key, _) in enumerate(columnas):
                    val = d.get(key)
                    if key in ("fecha_hora", "fecha_hora_anterior", "intervalo"):
                        wsh.write(i, j, "" if pd.isna(val) else str(val), f_cell)
                    elif key == "tipo_comparacion":
                        v = "" if pd.isna(val) else str(val)
                        fmt = f_mismo if v == "Mismo dia" else (f_sig if v == "Dia siguiente" else f_otro)
                        wsh.write(i, j, v, fmt)
                    else:
                        wsh.write(i, j, int(val) if pd.notna(val) else 0, f_num)

            wsh.freeze_panes(3, 1)
            wsh.autofilter(2, 0, 2 + len(hist), len(columnas) - 1)

            # --- Grafico de lineas: evolucion de movimientos ---
            if len(hist) >= 1:
                n = len(hist)
                chart = wb.add_chart({"type": "line"})
                # Serie: Pasaron a Revision  (columna 8 = idx 7)
                chart.add_series({
                    "name":       "Pasaron a Revision",
                    "categories": ["Historial_Diario", 3, 0, 2 + n, 0],
                    "values":     ["Historial_Diario", 3, 7, 2 + n, 7],
                    "line":       {"color": "#1F4E78", "width": 2.25},
                    "marker":     {"type": "circle", "size": 6, "fill": {"color": "#1F4E78"}},
                })
                # Serie: Pasaron a Devuelto (columna 9 = idx 8)
                chart.add_series({
                    "name":       "Pasaron a Devuelto",
                    "categories": ["Historial_Diario", 3, 0, 2 + n, 0],
                    "values":     ["Historial_Diario", 3, 8, 2 + n, 8],
                    "line":       {"color": "#C00000", "width": 2.25},
                    "marker":     {"type": "square", "size": 6, "fill": {"color": "#C00000"}},
                })
                # Serie: Nuevos ingresados (columna 11 = idx 10)
                chart.add_series({
                    "name":       "Nuevos ingresados",
                    "categories": ["Historial_Diario", 3, 0, 2 + n, 0],
                    "values":     ["Historial_Diario", 3, 10, 2 + n, 10],
                    "line":       {"color": "#548235", "width": 2.25, "dash_type": "dash"},
                    "marker":     {"type": "diamond", "size": 6, "fill": {"color": "#548235"}},
                })
                # Serie: Reaperturas (columna 12 = idx 11)
                chart.add_series({
                    "name":       "Reaperturas",
                    "categories": ["Historial_Diario", 3, 0, 2 + n, 0],
                    "values":     ["Historial_Diario", 3, 11, 2 + n, 11],
                    "line":       {"color": "#BF8F00", "width": 2.25, "dash_type": "dash_dot"},
                    "marker":     {"type": "triangle", "size": 6, "fill": {"color": "#BF8F00"}},
                })
                # Serie: Salieron de Revision (idx 12)
                chart.add_series({
                    "name":       "Salieron de Revision",
                    "categories": ["Historial_Diario", 3, 0, 2 + n, 0],
                    "values":     ["Historial_Diario", 3, 12, 2 + n, 12],
                    "line":       {"color": "#7030A0", "width": 2.25},
                    "marker":     {"type": "x", "size": 7, "fill": {"color": "#7030A0"}},
                })
                # Serie: Salieron de Devuelto (idx 13)
                chart.add_series({
                    "name":       "Salieron de Devuelto",
                    "categories": ["Historial_Diario", 3, 0, 2 + n, 0],
                    "values":     ["Historial_Diario", 3, 13, 2 + n, 13],
                    "line":       {"color": "#ED7D31", "width": 2.25},
                    "marker":     {"type": "plus", "size": 7, "fill": {"color": "#ED7D31"}},
                })
                chart.set_title({"name": "Evolucion de movimientos entre ejecuciones"})
                chart.set_x_axis({"name": "Ejecucion (fecha y hora)", "num_font": {"rotation": -45}})
                chart.set_y_axis({"name": "Cantidad de tickets"})
                chart.set_size({"width": 820, "height": 400})
                wsh.insert_chart("P3", chart)

        escribir_historial(historial)

        # --- Hojas con los tickets que CAMBIARON (destino de los links de movimientos)
        # Solo si hay datos de ayer para comparar.
        if not primer_dia:
            escribir_detalle("Movidos_a_Revision",
                             "Tickets que pasaron a REVISION",
                             datos["df_movidos_revision"])
            escribir_detalle("Movidos_a_Devuelto",
                             "Tickets que pasaron a DEVUELTO a Servicios",
                             datos["df_movidos_devuelto"])
            # En esta hoja NO tiene sentido mostrar campos propios del flujo
            # FDS (Estado / Prioridad / Propietario / Asignado a / Estado FDS)
            # porque son tickets que justamente salieron del control de FDS
            # y sus valores son irrelevantes o repetitivos. Se dejan solo las
            # columnas utiles: ID, Asunto, Estado anterior y fechas.
            _cols_ocultar_otros = [
                "Estado", "Prioridad", "Propietario",
                "Asignado a", "Estado FDS",
            ]
            _df_otros = datos["df_otros_estados"].drop(
                columns=[c for c in _cols_ocultar_otros
                         if c in datos["df_otros_estados"].columns],
                errors="ignore",
            )
            escribir_detalle("Cambio_Propietario_No_FDS",
                             "Tickets nuevos/abiertos que cambiaron de propietario diferente a FDS",
                             _df_otros)
            escribir_detalle("Nuevos_Ingresados",
                             "Tickets nuevos que ingresaron HOY",
                             datos["df_nuevos_ingresados"])
            escribir_detalle("Reaperturas",
                             "Reaperturas: tickets que volvieron al estado Nuevo/Abierto",
                             datos["df_reaperturas"])
            escribir_detalle("Salieron_de_Revision",
                             "Tickets que salieron de REVISION a otros estados (cerrados / resueltos)",
                             datos["df_salieron_revision"])
            escribir_detalle("Salieron_de_Devuelto",
                             "Tickets que salieron de DEVUELTO a Servicios a otros estados",
                             datos["df_salieron_devuelto"])

    return salida


# ---------------------------------------------------------------------------
# PASO 6: Programa principal
# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print("  MONITOR DE TICKETS - ZOHO ANALYTICS")
    print("=" * 60)

    asegurar_carpetas()

    # Validar que existan los 3 archivos de entrada
    faltantes = [n for n in ARCHIVOS.values() if not (ENTRADA_DIR / n).exists()]
    if faltantes:
        print("\n[ERROR] Faltan archivos en la carpeta 'entrada/':")
        for f in faltantes:
            print(f"   - {f}")
        print(f"\nColocalos en: {ENTRADA_DIR}")
        print("y vuelve a ejecutar: python monitor.py")
        sys.exit(1)

    ahora = datetime.now().replace(microsecond=0)
    prev  = snapshot_anterior(ahora)

    print(f"\n> Snapshot actual:   {ahora.strftime('%Y-%m-%d %H:%M:%S')}")
    if prev is None:
        print(f"> Snapshot anterior: (ninguno - primera ejecucion)")
    else:
        etiqueta, desc = clasificar_intervalo(ahora, prev[0])
        print(f"> Snapshot anterior: {prev[0].strftime('%Y-%m-%d %H:%M:%S')}  [{etiqueta}, {desc}]")

    # Comparar contra el snapshot anterior mas reciente
    datos = comparar(ahora)

    # Actualizar el historial acumulativo (una fila por ejecucion)
    historial = actualizar_historial(datos)

    # Imprimir resumen en consola
    print("\n--- RESUMEN ---")
    lbl_ant, lbl_act = "Anterior", "Actual"
    print(f"  Nuevos/Abiertos:  {lbl_ant}={datos['total_nuevos_ayer']:>4}   {lbl_act}={datos['total_nuevos_hoy']:>4}")
    print(f"  En revision:      {lbl_ant}={datos['total_revision_ayer']:>4}   {lbl_act}={datos['total_revision_hoy']:>4}")
    print(f"  Devueltos:        {lbl_ant}={datos['total_devuelto_ayer']:>4}   {lbl_act}={datos['total_devuelto_hoy']:>4}")
    if datos["hubo_ayer"]:
        print(f"\n  Desde Nuevos/Abiertos:")
        print(f"    -> Pasaron a REVISION:      {len(datos['ids_movidos_revision'])}")
        print(f"    -> Pasaron a DEVUELTO:      {len(datos['ids_movidos_devuelto'])}")
        print(f"    -> Cambio de propietario:   {len(datos['ids_otros_estados'])}")
        print(f"\n  Ingresos y reaperturas:")
        print(f"    -> Nuevos ingresados:       {len(datos['ids_nuevos_ingresados'])}")
        print(f"    -> Reaperturas:             {len(datos['ids_reaperturas'])}")
        print(f"\n  Salidas desde otros estados:")
        print(f"    -> Salieron de REVISION:    {len(datos['ids_salieron_revision'])}")
        print(f"    -> Salieron de DEVUELTO:    {len(datos['ids_salieron_devuelto'])}")
    else:
        print("\n  (Primer informe: no hay uno anterior para comparar)")

    # Generar el reporte Excel
    reporte = generar_reporte(datos, historial)
    print(f"\n[OK] Reporte generado: {reporte}")

    # Archivar el snapshot actual
    archivar_snapshot(ahora)
    print(f"[OK] Snapshot guardado en: {carpeta_snapshot(ahora)}")
    print(f"[OK] Historial acumulado ({len(historial)} ejecuciones): {HISTORIAL_CSV}")

    print("\nListo. Puedes abrir el reporte en Excel.\n")


if __name__ == "__main__":
    main()

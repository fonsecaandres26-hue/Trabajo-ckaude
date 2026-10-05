"""Genera MATRIZ_CONTROL_LOGISTICO_v2.xlsx: plantilla del modelo descrito en
docs/rediseno-matriz-control-logistico.md (catálogos, bitácora de eventos,
tablas SAE / Remisión / Reembarque con KPI y semáforo calculados).

Uso:  python plantilla/generar_plantilla.py
"""
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = Path(__file__).with_name("MATRIZ_CONTROL_LOGISTICO_v2.xlsx")
N = 2000  # filas cubiertas por rangos de fórmulas y validaciones
PAD = 300  # filas preformuladas en cada tabla (las vacías quedan en blanco)

FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")          # celdas de captura
BLACK = Font(name=FONT, color="000000")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF")
BLOCK_FILL = {  # color por bloque lógico (prefijo)
    "ID": "1F3864", "CTL": "595959", "RUT": "2E75B6", "TRN": "2F5597",
    "CAR": "548235", "DOC": "BF8F00", "F": "7030A0", "EST": "C00000",
    "KPI": "833C0B", "FAC": "375623", "CAT": "404040",
}
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")

wb = Workbook()


def hdr_block(name):
    pre = name.split("_")[0]
    return BLOCK_FILL.get(pre, BLOCK_FILL["CAT"])


def build_sheet(ws, cols, rows, table_name, widths=None, formulas=None):
    """cols: lista de (nombre, es_captura). formulas: {col: plantilla con {r}}."""
    formulas = formulas or {}
    for j, (name, is_input) in enumerate(cols, 1):
        c = ws.cell(row=1, column=j, value=name)
        c.font = HDR_FONT
        c.fill = PatternFill("solid", fgColor=hdr_block(name))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(name, max(12, len(name) + 2))
    rows = list(rows) + [{}] * max(0, PAD - len(rows))
    for i, row in enumerate(rows, 2):
        for j, (name, is_input) in enumerate(cols, 1):
            c = ws.cell(row=i, column=j)
            if name in formulas:
                c.value = formulas[name].format(r=i)
                c.font = BLACK
            else:
                c.value = row.get(name)
                c.font = BLUE if is_input else BLACK
                if is_input:
                    c.fill = INPUT_FILL
            c.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    last = max(len(rows) + 1, 2)
    t = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(cols))}{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "B2"
    ws.row_dimensions[1].height = 32
    ws.sheet_view.zoomScale = 90


def col(ws_cols, name):
    return get_column_letter([c[0] for c in ws_cols].index(name) + 1)


def add_list(ws, ws_cols, name, source):
    dv = DataValidation(type="list", formula1=source, allow_blank=True,
                        showErrorMessage=True, errorTitle="Valor fuera de catálogo",
                        error="Seleccione un valor de la lista (hoja CATALOGOS).")
    L = col(ws_cols, name)
    dv.add(f"{L}2:{L}{N}")
    ws.add_data_validation(dv)


# ---------------------------------------------------------------- LEEME
ws = wb.active
ws.title = "LEEME"
leeme = [
    ("MATRIZ DE CONTROL LOGÍSTICO v2 — Reembarques SAE / DEX", True),
    ("", False),
    ("Cómo usarla", True),
    ("1. Registre cada hito en EVENTOS (una fila por hecho, con fecha-hora real). Es la ÚNICA hoja donde se escriben fechas.", False),
    ("2. Cree el reembarque en REEMBARQUE, el file en REMISION y cada SAE en SAE (solo columnas amarillas).", False),
    ("3. Fase, horas, cumplimiento de ANS, conciliación y semáforo se calculan solos (texto negro). No los sobrescriba.", False),
    ("4. Los umbrales de ANS y la fecha de corte del tablero están en CATALOGOS > Parámetros.", False),
    ("", False),
    ("Convención de colores", True),
    ("Fondo amarillo + texto azul = celda de captura.   Texto negro = fórmula.", False),
    ("Encabezado por bloque: ID (azul oscuro) · CTL (gris) · RUT/TRN (azul) · CAR (verde) · DOC (ocre) · F (morado) · EST (rojo) · KPI (marrón) · FAC (verde oscuro)", False),
    ("", False),
    ("Formatos", True),
    ("Fechas: yyyy-mm-dd hh:mm (hora Colombia). Duraciones: horas decimales. SAE y DEX: texto de 13 dígitos. País: ISO 3166 alfa-2. Puertos: UN/LOCODE. Contenedor: ISO 6346.", False),
    ("", False),
    ("Las filas que trae la plantilla son un EJEMPLO ficticio (reembarque BST3-0412). Bórrelas antes de empezar a operar.", True),
    ("Diseño completo: docs/rediseno-matriz-control-logistico.md", False),
]
for i, (t, b) in enumerate(leeme, 1):
    ws.cell(row=i, column=1, value=t).font = Font(name=FONT, bold=b, size=14 if i == 1 else 10)
ws.column_dimensions["A"].width = 140

# ---------------------------------------------------------------- CATALOGOS
cat = wb.create_sheet("CATALOGOS")


def cat_table(ws, top_left_col, title, headers, data, tname):
    c0 = top_left_col
    ws.cell(row=1, column=c0, value=title).font = Font(name=FONT, bold=True, size=11)
    for j, h in enumerate(headers):
        c = ws.cell(row=2, column=c0 + j, value=h)
        c.font = HDR_FONT
        c.fill = PatternFill("solid", fgColor=BLOCK_FILL["CAT"])
        ws.column_dimensions[get_column_letter(c0 + j)].width = max(10, min(40, max(len(str(h)), *(len(str(r[j])) for r in data)) + 2))
    for i, r in enumerate(data, 3):
        for j, v in enumerate(r):
            ws.cell(row=i, column=c0 + j, value=v).font = BLUE
    ref = f"{get_column_letter(c0)}2:{get_column_letter(c0 + len(headers) - 1)}{len(data) + 2}"
    t = Table(displayName=tname, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight8", showRowStripes=True)
    ws.add_table(t)
    return {h: f"CATALOGOS!${get_column_letter(c0 + j)}$3:${get_column_letter(c0 + j)}${len(data) + 2}"
            for j, h in enumerate(headers)}


# Parámetros (columnas A-C)
params = [
    ("FECHA_CORTE", datetime(2026, 10, 4, 8, 0), "Fecha-hora del tablero. En operación real reemplazar por =AHORA()"),
    ("ANS_ELAB_SAE_H", 4, "Horas recepción → borrador (fórmula actual TIME(4,0,0))"),
    ("ANS_TOTAL_SAE_H", 28, "Horas recepción → SAE autorizada (fórmula actual 1+TIME(4,0,0))"),
    ("ANS_CLASIF_H", 3, "Horas solicitud → respuesta clasificación"),
    ("ANS_DEX_H", 72, "Horas manifiesto → DEX. SUPUESTO: definir según norma vigente y contrato"),
    ("TOL_PESO_KG", 0.5, "Diferencia máxima peso SAE vs manifestado. SUPUESTO: definir"),
    ("UMBRAL_AMBAR", 0.75, "Fracción del ANS de fase a partir de la cual el semáforo pasa a ámbar"),
]
P = cat_table(cat, 1, "Parámetros", ["PARAMETRO", "VALOR", "NOTA"], params, "cat_Parametros")
cat["B3"].number_format = "yyyy-mm-dd hh:mm"
PARAM = {name: f"CATALOGOS!$B${i}" for i, (name, _, _) in enumerate(params, 3)}

hitos = [  # HITO_ID, ORDEN, PASO, FASE, NIVEL, DUEÑO, PRÓXIMA ACCIÓN, ANS_FASE_H
    ("H000", 0, 1, "F00 Remisión recibida", "REMISION", "AA KN", "Enviar borrador SAE", 4),
    ("H010", 10, 2, "F10 Clasificación solicitada", "REMISION", "CLIENTE", "Responder clasificación", 3),
    ("H015", 15, 2, "F10 Clasificación respondida", "REMISION", "AA KN", "Enviar borrador SAE", 4),
    ("H020", 20, 3, "F20 Borrador SAE enviado", "REMISION", "KN CDLI", "Aprobar SAE en bodega", 24),
    ("H030", 30, 4, "F30 Aprobado bodega", "REMISION", "AA KN", "Autorizar SAE", 4),
    ("H040", 40, 5, "F40 SAE autorizada", "SAE", "TERMINAL", "Ingreso a puerto", 48),
    ("H045", 45, 6, "F45 Inspección física", "SAE", "DIAN", "Liberar inspección", 48),
    ("H046", 46, 6, "F45 Inspección liberada", "SAE", "TERMINAL", "Autorizar embarque", 24),
    ("H050", 50, 6, "F50 Autoriza embarque", "REEMBARQUE", "NAVIERA", "Zarpe / manifiesto", 72),
    ("H060", 60, 7, "F60 Embarcado (manifiesto)", "REEMBARQUE", "AA KN", "Presentar DEX", 72),
    ("H070", 70, 8, "F70 DEX presentada", "SAE", "AA KN", "Cerrar carpeta documental", 120),
    ("H080", 80, 9, "F80 Cierre documental", "REMISION", "ADMIN", "Prefacturar", 240),
    ("H090", 90, 10, "F90 Facturado", "REMISION", "—", "—", 0),
]
H = cat_table(cat, 5, "cat_Hito (máquina de estados)",
              ["HITO_ID", "ORDEN", "PASO", "FASE", "NIVEL", "DUENO", "PROXIMA_ACCION", "ANS_FASE_H"], hitos, "cat_Hito")

causales = [
    ("REM01", "Remisión incompleta", "CLIENTE"), ("REM02", "Remisión con datos errados", "CLIENTE"),
    ("REM03", "Cambio de remisión", "CLIENTE"), ("REM04", "Recibida fuera de horario / fin de semana", "CLIENTE"),
    ("CLA01", "Model code sin clasificar", "AA KN"), ("CLA02", "Respuesta tardía del cliente a clasificación", "CLIENTE"),
    ("CLA03", "PA sugerida difiere", "CLIENTE"),
    ("SAE01", "Demora en elaboración SAE", "AA KN"), ("SAE02", "Demora en elaboración SAE por clasificación", "AA KN"),
    ("SAE03", "Demora en reserva / confirmación o cambio de transportador", "TERCERO"),
    ("SAE04", "Remisión para adelantar sin información completa", "KN CDLI"),
    ("SAE05", "Demora +3 h en aprobación de SAE por bodega", "KN CDLI"),
    ("SAE06", "Demora +6 h en aprobación de SAE por bodega", "KN CDLI"),
    ("SAE07", "Demora +24 h en aprobación de SAE por bodega", "KN CDLI"),
    ("SAE08", "Error o falta de información en la remisión", "KN CDLI"),
    ("SAE09", "Operación fin de semana", "KN CDLI"),
    ("SAE10", "Rechazo DIAN de SAE", "AA KN"), ("SAE11", "Error de pesos", "KN CDLI"),
    ("SAE12", "Error de bultos", "KN CDLI"), ("SAE13", "Error de valor FOB", "AA KN"),
    ("SAE14", "Error de ítems / cantidades", "AA KN"),
    ("PTO01", "Inspección física", "DIAN"), ("PTO02", "Cierre de terminal", "TERMINAL"),
    ("PTO03", "Roll-over de naviera", "NAVIERA"),
    ("DEX01", "DEX manual por contingencia", "DIAN"), ("DEX02", "Corrección de DEX", "AA KN"),
    ("DEX03", "Diferencia manifiesto vs SAE", "AA KN"),
]
C = cat_table(cat, 14, "cat_Causal", ["CAUSAL_ID", "DESCRIPCION", "RESPONSABLE_TIPO"], causales, "cat_Causal")

transp = [
    ("MAEU", "MAERSK", "CARRIER/FORWARDER", "MAERSK | \\xa0MAERSK"),
    ("HLCU", "HAPAG-LLOYD", "CARRIER", "HAPAG LLOYD | HAPAG LLOYD CONTAINER LINE"),
    ("ZIMU", "ZIM", "CARRIER", "ZIM ISRAEL NAVIGATION | ZIM ISRAELNAVIGATION"),
    ("ONEY", "OCEAN NETWORK EXPRESS", "CARRIER", "OCEAN NETWORKEXPRESS"),
    ("CMDU", "CMA CGM", "CARRIER/FORWARDER", "CMA CGM INC"),
    ("EGLV", "EVERGREEN", "CARRIER", "EVERGREEN"),
    ("KN", "KUEHNE+NAGEL", "FORWARDER", "KUEHNE NAGEL | KUEHNE+NAGEL"),
    ("L360", "LOGÍSTICA 360", "FORWARDER", "LOGISTICA 360 | LOGÍSTICA 360 | LOGISTICA 362"),
    ("DHL", "DHL", "FORWARDER", "DHL"),
    ("COPA", "COPA AIRLINES CARGO", "CARRIER AER", "COPA"),
]
T = cat_table(cat, 18, "cat_Transportista", ["TRANSP_ID", "NOMBRE", "TIPO", "ALIAS_CONSOLIDADO"], transp, "cat_Transportista")

paises = [("CL", "CHILE"), ("MX", "MÉXICO"), ("UY", "URUGUAY"), ("AR", "ARGENTINA"), ("PA", "PANAMÁ"),
          ("CR", "COSTA RICA"), ("PY", "PARAGUAY"), ("SV", "EL SALVADOR"), ("IE", "IRLANDA"),
          ("US", "ESTADOS UNIDOS"), ("CA", "CANADÁ"), ("CO", "COLOMBIA")]
PA = cat_table(cat, 23, "cat_Pais (ISO 3166-1)", ["PAIS_ID", "NOMBRE"], paises, "cat_Pais")

personas = [("P001", "Digitador 1", "DIGITADOR"), ("P002", "Digitador 2", "DIGITADOR"),
            ("P003", "Digitador 3", "DIGITADOR"), ("P005", "Responsable DEX", "RESPONSABLE_DEX"),
            ("CDLI", "Bodega KN CDLI", "BODEGA"), ("TERMINAL", "Terminal portuaria", "EXTERNO"),
            ("NAVIERA", "Naviera", "EXTERNO")]
PE = cat_table(cat, 26, "cat_Persona", ["PERSONA_ID", "NOMBRE", "ROL"], personas, "cat_Persona")

listas = [("CONTECAR", "MAR", "AUTORIZA_EMBARQUE", "AUTOMATICA", "EN_DIGITACION", "CARPETA_CREADA", "MANUAL", "S", "BOOSTER", "45G1"),
          ("SPRC", "AER", "INSPECCION_FISICA", "MANUAL", "EN_REVISION", "PDTE", "CORREO", "N", "RT", "42G1"),
          (None, "TER", "INSPECCION_DOCUMENTAL", "CORRECCION", "PDTE_VALORES", "FALTA_DEX", "MUISCA", None, "DOLLARCITY", "22G1"),
          (None, None, None, None, "CONFORME", "COMPLETO", "WMS", None, None, "45R1"),
          (None, None, None, None, None, None, "NAVIERA", None, None, None)]
hdrs = ["TERMINAL", "MODO", "CANAL_DIAN", "DEX_TIPO", "REVISION", "GESTOR_DOC", "FUENTE", "SI_NO", "PROYECTO", "TIPO_CONT"]
c0 = 30
cat.cell(row=1, column=c0, value="Listas cortas").font = Font(name=FONT, bold=True, size=11)
L = {}
for j, h in enumerate(hdrs):
    cc = cat.cell(row=2, column=c0 + j, value=h)
    cc.font = HDR_FONT
    cc.fill = PatternFill("solid", fgColor=BLOCK_FILL["CAT"])
    vals = [r[j] for r in listas if r[j] is not None]
    for i, v in enumerate(vals, 3):
        cat.cell(row=i, column=c0 + j, value=v).font = BLUE
    letter = get_column_letter(c0 + j)
    cat.column_dimensions[letter].width = max(12, max(len(v) for v in vals) + 2)
    L[h] = f"CATALOGOS!${letter}$3:${letter}${len(vals) + 2}"
cat.freeze_panes = "A3"

# ---------------------------------------------------------------- EVENTOS
ev = wb.create_sheet("EVENTOS")
ev_cols = [("ID_EVENTO", False), ("F_FECHA_HORA", True), ("CTL_NIVEL", True), ("ID_REEMBARQUE", True),
           ("ID_FILE_ADUANA", True), ("ID_SAE", True), ("EST_HITO_ID", True), ("EST_ORDEN", False),
           ("EST_FASE", False), ("CTL_RESPONSABLE_ID", True), ("CTL_FUENTE", True),
           ("EST_CAUSAL_ID", True), ("CTL_NOTA", True)]
R, F, S = "BST3-0412", "900000001-0223-412", "900000001-0223-413"
M, R2 = "900000001-0223-414", "BST3-0413"
d = lambda s: datetime.strptime(s, "%Y-%m-%d %H:%M")
ev_rows = [
    # file 412
    (d("2026-09-29 08:15"), "REMISION", R, F, None, "H000", "CDLI", "CORREO", None, "PL 2900230011"),
    (d("2026-09-29 10:51"), "REMISION", R, F, None, "H020", "P002", "CORREO", None, "2 SAE"),
    (d("2026-09-29 15:30"), "REMISION", R, F, None, "H030", "CDLI", "CORREO", None, None),
    (d("2026-09-30 05:40"), "SAE", R, F, "6027770000011", "H040", "P002", "MUISCA", None, None),
    (d("2026-09-30 05:40"), "SAE", R, F, "6027770000028", "H040", "P002", "MUISCA", None, None),
    # file 413
    (d("2026-09-29 09:00"), "REMISION", R, S, None, "H000", "CDLI", "CORREO", None, None),
    (d("2026-09-29 14:06"), "REMISION", R, S, None, "H020", "P001", "CORREO", "SAE08", "Remisión con bultos errados, se corrige"),
    (d("2026-09-29 20:00"), "REMISION", R, S, None, "H030", "CDLI", "CORREO", None, None),
    (d("2026-09-30 15:12"), "SAE", R, S, "6027770000035", "H040", "P001", "MUISCA", None, None),
    (d("2026-10-01 10:00"), "SAE", R, S, "6027770000035", "H045", "TERMINAL", "MUISCA", "PTO01", "Selección inspección física"),
    (d("2026-10-02 11:00"), "SAE", R, S, "6027770000035", "H046", "TERMINAL", "MUISCA", None, None),
    # reembarque 412 (aplica a todos sus files)
    (d("2026-10-01 09:10"), "REEMBARQUE", R, None, None, "H050", "TERMINAL", "MUISCA", None, "Canal: autoriza embarque"),
    (d("2026-10-03 22:40"), "REEMBARQUE", R, None, None, "H060", "NAVIERA", "NAVIERA", None, "Buque/viaje 640S"),
    (d("2026-10-03 22:40"), "SAE", R, F, "6027770000011", "H070", "P005", "MUISCA", None, None),
    (d("2026-10-03 22:40"), "SAE", R, F, "6027770000028", "H070", "P005", "MUISCA", None, None),
    # file 414 (otro reembarque)
    (d("2026-10-03 14:20"), "REMISION", R2, M, None, "H000", "CDLI", "CORREO", "REM04", "Recibida sábado"),
    (d("2026-10-03 17:44"), "REMISION", R2, M, None, "H020", "P003", "CORREO", None, None),
    (d("2026-10-04 04:48"), "REMISION", R2, M, None, "H030", "CDLI", "CORREO", None, None),
]
names = ["F_FECHA_HORA", "CTL_NIVEL", "ID_REEMBARQUE", "ID_FILE_ADUANA", "ID_SAE", "EST_HITO_ID",
         "CTL_RESPONSABLE_ID", "CTL_FUENTE", "EST_CAUSAL_ID", "CTL_NOTA"]
ev_dicts = [dict(zip(names, r)) for r in ev_rows]
ev_f = {
    "ID_EVENTO": "=IF(B{r}=\"\",\"\",ROW()-1)",
    "EST_ORDEN": f"=IFERROR(INDEX({H['ORDEN']},MATCH(G{{r}},{H['HITO_ID']},0)),\"\")",
    "EST_FASE": f"=IFERROR(INDEX({H['FASE']},MATCH(G{{r}},{H['HITO_ID']},0)),\"\")",
}
build_sheet(ev, ev_cols, ev_dicts, "tbl_Eventos", widths={"CTL_NOTA": 40, "EST_FASE": 28, "F_FECHA_HORA": 18}, formulas=ev_f)
for r in range(2, N + 1):
    ev[f"B{r}"].number_format = "yyyy-mm-dd hh:mm"
    ev[f"E{r}"].number_format = "@"
    ev[f"F{r}"].number_format = "@"
add_list(ev, ev_cols, "CTL_NIVEL", '"REEMBARQUE,REMISION,SAE"')
add_list(ev, ev_cols, "EST_HITO_ID", H["HITO_ID"])
add_list(ev, ev_cols, "CTL_RESPONSABLE_ID", PE["PERSONA_ID"])
add_list(ev, ev_cols, "CTL_FUENTE", L["FUENTE"])
add_list(ev, ev_cols, "EST_CAUSAL_ID", C["CAUSAL_ID"])
dt_dv = DataValidation(type="date", operator="between", formula1="DATE(2018,1,1)", formula2="TODAY()+180",
                       showErrorMessage=True, error="Fecha-hora real entre 2018 y hoy+180 días.")
dt_dv.add(f"B2:B{N}")
ev.add_data_validation(dt_dv)
ev["B1"].comment = Comment("Única columna de fechas del modelo. Siempre fecha-hora real del hecho (hora Colombia).", "Plantilla")

E_FECHA, E_REEMB, E_FILE, E_SAE, E_HITO, E_ORDEN = (f"EVENTOS!${c}$2:${c}${N}" for c in "BDEFGH")

# ---------------------------------------------------------------- REEMBARQUE
rb = wb.create_sheet("REEMBARQUE")
rb_cols = [("ID_REEMBARQUE", True), ("ID_PROYECTO", True), ("TRN_MODO", True), ("TRN_AGENTE_CARGA_ID", True),
           ("TRN_NAVIERA_ID", True), ("TRN_BOOKING", True), ("TRN_SEMANA_RESERVA", True), ("RUT_TERMINAL_ID", True),
           ("RUT_PUERTO_EMBARQUE", True), ("RUT_PUERTO_DESTINO", True), ("TRN_BUQUE_VIAJE", True),
           ("TRN_BL", True), ("TRN_CONTENEDOR", True), ("TRN_TIPO_CONTENEDOR", True), ("F_ETD", True),
           ("F_ZARPE_REAL", False), ("KPI_DESVIO_ZARPE_D", False), ("CAR_FILES", False), ("CAR_SAE", False),
           ("CAR_PESO_BRUTO_KG", False), ("DOC_POLIZA_NUM", True), ("EST_FASE", False)]
rb_rows = [
    dict(ID_REEMBARQUE=R, ID_PROYECTO="BOOSTER", TRN_MODO="MAR", TRN_AGENTE_CARGA_ID="MAEU", TRN_NAVIERA_ID="MAEU",
         TRN_BOOKING="268114502", TRN_SEMANA_RESERVA="2026-W40", RUT_TERMINAL_ID="CONTECAR",
         RUT_PUERTO_EMBARQUE="COCTG", RUT_PUERTO_DESTINO="CLVAP", TRN_BUQUE_VIAJE="EJEMPLO / 640S",
         TRN_BL="MAEU268114502", TRN_CONTENEDOR="MSKU7340215", TRN_TIPO_CONTENEDOR="45G1", F_ETD=datetime(2026, 10, 3)),
    dict(ID_REEMBARQUE=R2, ID_PROYECTO="BOOSTER", TRN_MODO="MAR", TRN_AGENTE_CARGA_ID="L360", TRN_NAVIERA_ID="HLCU",
         TRN_BOOKING="HLCU-CTG-88231", TRN_SEMANA_RESERVA="2026-W40", RUT_TERMINAL_ID="SPRC",
         RUT_PUERTO_EMBARQUE="COCTG", RUT_PUERTO_DESTINO="MXZLO", TRN_CONTENEDOR="HLXU8812043",
         TRN_TIPO_CONTENEDOR="45G1", F_ETD=datetime(2026, 10, 6)),
]
rb_f = {
    "F_ZARPE_REAL": f"=IF(A{{r}}=\"\",\"\",IF(_xlfn.MAXIFS({E_FECHA},{E_REEMB},A{{r}},{E_HITO},\"H060\")=0,\"\",_xlfn.MAXIFS({E_FECHA},{E_REEMB},A{{r}},{E_HITO},\"H060\")))",
    "KPI_DESVIO_ZARPE_D": "=IF(OR(P{r}=\"\",O{r}=\"\"),\"\",ROUND(INT(P{r})-O{r},0))",
    "CAR_FILES": "=IF(A{r}=\"\",\"\",COUNTIF(REMISION!$C$2:$C$" + str(N) + ",A{r}))",
    "CAR_SAE": "=IF(A{r}=\"\",\"\",SUMIF(REMISION!$C$2:$C$" + str(N) + ",A{r},REMISION!$H$2:$H$" + str(N) + "))",
    "CAR_PESO_BRUTO_KG": "=IF(A{r}=\"\",\"\",SUMIF(REMISION!$C$2:$C$" + str(N) + ",A{r},REMISION!$I$2:$I$" + str(N) + "))",
    "EST_FASE": f"=IF(A{{r}}=\"\",\"\",IFERROR(INDEX({H['FASE']},MATCH(_xlfn.MINIFS(REMISION!$L$2:$L${N},REMISION!$C$2:$C${N},A{{r}}),{H['ORDEN']},0)),\"\"))",
}
build_sheet(rb, rb_cols, rb_rows, "tbl_Reembarque", formulas=rb_f, widths={"EST_FASE": 28})
for r in range(2, N + 1):
    rb[f"O{r}"].number_format = "yyyy-mm-dd"
    rb[f"P{r}"].number_format = "yyyy-mm-dd hh:mm"
    rb[f"T{r}"].number_format = "#,##0.00"
    rb[f"F{r}"].number_format = "@"
add_list(rb, rb_cols, "ID_PROYECTO", L["PROYECTO"])
add_list(rb, rb_cols, "TRN_MODO", L["MODO"])
add_list(rb, rb_cols, "TRN_AGENTE_CARGA_ID", T["TRANSP_ID"])
add_list(rb, rb_cols, "TRN_NAVIERA_ID", T["TRANSP_ID"])
add_list(rb, rb_cols, "RUT_TERMINAL_ID", L["TERMINAL"])
add_list(rb, rb_cols, "TRN_TIPO_CONTENEDOR", L["TIPO_CONT"])
cont_dv = DataValidation(type="textLength", operator="equal", formula1="11", allow_blank=True,
                         showErrorMessage=True, error="ISO 6346: 4 letras + 7 dígitos (11 caracteres).")
cont_dv.add(f"M2:M{N}")
rb.add_data_validation(cont_dv)
rb["Q1"].comment = Comment("Días entre zarpe real y ETD. Positivo = zarpó tarde.", "Plantilla")
rb["V1"].comment = Comment("Fase del file MÁS ATRASADO del reembarque.", "Plantilla")

# ---------------------------------------------------------------- REMISION
rm = wb.create_sheet("REMISION")
rm_cols = [("ID_FILE_ADUANA", True), ("ID_REMISION", True), ("ID_REEMBARQUE", True), ("RUT_PAIS_DESTINO", True),
           ("RUT_DESTINATARIO", True), ("DOC_PACKING_LIST", True), ("CTL_DIGITADOR_ID", True),
           ("CAR_SAE", False), ("CAR_PESO_BRUTO_KG", False), ("CAR_VALOR_FOB_USD", False),
           ("F_RECEPCION", False), ("EST_ORDEN", False), ("EST_FASE", False), ("EST_PROGRESO", False),
           ("EST_REVISION", True), ("EST_GESTOR_DOC", True), ("EST_BLOQUEO", True), ("EST_BLOQUEO_CAUSAL_ID", True),
           ("KPI_SAE_NO_CONCILIADAS", False), ("EST_DUENO_ACTUAL", False), ("EST_PROXIMA_ACCION", False),
           ("F_ULTIMO_HITO", False), ("EST_HORAS_EN_FASE", False), ("EST_ANS_FASE_H", False), ("EST_SEMAFORO", False),
           ("FAC_PERIODO", True), ("CTL_OBSERVACION", True)]
rm_rows = [
    dict(ID_FILE_ADUANA=F, ID_REMISION="BOOST-412.1-CL", ID_REEMBARQUE=R, RUT_PAIS_DESTINO="CL",
         RUT_DESTINATARIO="DECATHLON CHILE SPA", DOC_PACKING_LIST="2900230011", CTL_DIGITADOR_ID="P002",
         EST_REVISION="CONFORME", EST_GESTOR_DOC="PDTE", EST_BLOQUEO="N"),
    dict(ID_FILE_ADUANA=S, ID_REMISION="BOOST-412.2-CL", ID_REEMBARQUE=R, RUT_PAIS_DESTINO="CL",
         RUT_DESTINATARIO="DECATHLON CHILE SPA", DOC_PACKING_LIST="2900230012", CTL_DIGITADOR_ID="P001",
         EST_REVISION="CONFORME", EST_GESTOR_DOC="FALTA_DEX", EST_BLOQUEO="N"),
    dict(ID_FILE_ADUANA=M, ID_REMISION="BOOST-413.1-MX", ID_REEMBARQUE=R2, RUT_PAIS_DESTINO="MX",
         RUT_DESTINATARIO="ARTICULOS DEPORTIVOS DECATHLON SA DE CV", DOC_PACKING_LIST="2900230020",
         CTL_DIGITADOR_ID="P003", EST_REVISION="EN_REVISION", EST_GESTOR_DOC="PDTE", EST_BLOQUEO="N"),
]
SA = lambda c: f"SAE!${c}$2:${c}${N}"
orden_file = f"_xlfn.MAXIFS({E_ORDEN},{E_FILE},A{{r}})"
orden_reemb = f"_xlfn.MAXIFS({E_ORDEN},{E_REEMB},C{{r}},{E_FILE},\"=\")"
fecha_file = f"_xlfn.MAXIFS({E_FECHA},{E_FILE},A{{r}})"
fecha_reemb = f"_xlfn.MAXIFS({E_FECHA},{E_REEMB},C{{r}},{E_FILE},\"=\")"
rm_f = {
    "CAR_SAE": f"=IF(A{{r}}=\"\",\"\",COUNTIF({SA('B')},A{{r}}))",
    "CAR_PESO_BRUTO_KG": f"=IF(A{{r}}=\"\",\"\",SUMIF({SA('B')},A{{r}},{SA('F')}))",
    "CAR_VALOR_FOB_USD": f"=IF(A{{r}}=\"\",\"\",SUMIF({SA('B')},A{{r}},{SA('D')}))",
    "F_RECEPCION": f"=IF(A{{r}}=\"\",\"\",IF(_xlfn.MAXIFS({E_FECHA},{E_FILE},A{{r}},{E_HITO},\"H000\")=0,\"\",_xlfn.MAXIFS({E_FECHA},{E_FILE},A{{r}},{E_HITO},\"H000\")))",
    "EST_ORDEN": f"=IF(A{{r}}=\"\",\"\",MAX({orden_file},{orden_reemb}))",
    "EST_FASE": f"=IF(A{{r}}=\"\",\"\",IFERROR(INDEX({H['FASE']},MATCH(L{{r}},{H['ORDEN']},0)),\"\"))",
    "EST_PROGRESO": f"=IF(A{{r}}=\"\",\"\",IFERROR(REPT(\"●\",INDEX({H['PASO']},MATCH(L{{r}},{H['ORDEN']},0)))&REPT(\"○\",10-INDEX({H['PASO']},MATCH(L{{r}},{H['ORDEN']},0))),\"\"))",
    "KPI_SAE_NO_CONCILIADAS": f"=IF(A{{r}}=\"\",\"\",COUNTIFS({SA('B')},A{{r}},{SA('K')},\"N\"))",
    "EST_DUENO_ACTUAL": f"=IF(A{{r}}=\"\",\"\",IFERROR(INDEX({H['DUENO']},MATCH(L{{r}},{H['ORDEN']},0)),\"\"))",
    "EST_PROXIMA_ACCION": f"=IF(A{{r}}=\"\",\"\",IFERROR(INDEX({H['PROXIMA_ACCION']},MATCH(L{{r}},{H['ORDEN']},0)),\"\"))",
    "F_ULTIMO_HITO": f"=IF(A{{r}}=\"\",\"\",MAX({fecha_file},{fecha_reemb}))",
    "EST_HORAS_EN_FASE": f"=IF(OR(A{{r}}=\"\",V{{r}}=0),\"\",ROUND(({PARAM['FECHA_CORTE']}-V{{r}})*24,1))",
    "EST_ANS_FASE_H": f"=IF(A{{r}}=\"\",\"\",IFERROR(INDEX({H['ANS_FASE_H']},MATCH(L{{r}},{H['ORDEN']},0)),\"\"))",
    "EST_SEMAFORO": (f"=IF(A{{r}}=\"\",\"\",IF(OR(Q{{r}}=\"S\",S{{r}}>0,L{{r}}=45),\"ROJO\",IF(L{{r}}>=80,\"CERRADO\","
                     f"IF(W{{r}}>X{{r}},\"ROJO\",IF(W{{r}}>X{{r}}*{PARAM['UMBRAL_AMBAR']},\"AMBAR\",\"VERDE\")))))"),
}
build_sheet(rm, rm_cols, rm_rows, "tbl_Remision", formulas=rm_f,
            widths={"EST_FASE": 28, "EST_PROGRESO": 16, "RUT_DESTINATARIO": 34, "EST_PROXIMA_ACCION": 26,
                    "CTL_OBSERVACION": 40, "ID_FILE_ADUANA": 22})
for r in range(2, N + 1):
    for c in "KV":
        rm[f"{c}{r}"].number_format = "yyyy-mm-dd hh:mm"
    rm[f"I{r}"].number_format = "#,##0.00"
    rm[f"J{r}"].number_format = "#,##0.00"
    rm[f"F{r}"].number_format = "@"
add_list(rm, rm_cols, "ID_REEMBARQUE", f"REEMBARQUE!$A$2:$A${N}")
add_list(rm, rm_cols, "RUT_PAIS_DESTINO", PA["PAIS_ID"])
add_list(rm, rm_cols, "CTL_DIGITADOR_ID", PE["PERSONA_ID"])
add_list(rm, rm_cols, "EST_REVISION", L["REVISION"])
add_list(rm, rm_cols, "EST_GESTOR_DOC", L["GESTOR_DOC"])
add_list(rm, rm_cols, "EST_BLOQUEO", L["SI_NO"])
add_list(rm, rm_cols, "EST_BLOQUEO_CAUSAL_ID", C["CAUSAL_ID"])
for color, val in (("F8696B", "ROJO"), ("FFD966", "AMBAR"), ("A9D08E", "VERDE"), ("D9D9D9", "CERRADO")):
    rm.conditional_formatting.add(f"Y2:Y{N}", FormulaRule(formula=[f'$Y2="{val}"'],
                                  fill=PatternFill("solid", fgColor=color), font=Font(name=FONT, bold=True)))
rm.conditional_formatting.add(f"A2:AA{N}", FormulaRule(formula=['$Y2="ROJO"'], font=Font(name=FONT, color="C00000")))
rm["Y1"].comment = Comment("Prioridad: 1) bloqueo / SAE no conciliada / inspección física = ROJO; 2) fase >= F80 = CERRADO; "
                           "3) horas en fase > ANS = ROJO; 4) > UMBRAL_AMBAR x ANS = AMBAR; 5) VERDE.", "Plantilla")
rm["L1"].comment = Comment("Último hito cumplido: máximo ORDEN entre eventos del file y eventos del reembarque sin file.", "Plantilla")

# ---------------------------------------------------------------- SAE
sae = wb.create_sheet("SAE")
sae_cols = [("ID_SAE", True), ("ID_FILE_ADUANA", True), ("CAR_CANTIDAD_UND", True), ("CAR_VALOR_FOB_USD", True),
            ("CAR_PESO_NETO_KG", True), ("CAR_PESO_BRUTO_KG", True), ("CAR_BULTOS", True),
            ("DOC_PESO_BRUTO_MANIF_KG", True), ("DOC_BULTOS_MANIF", True), ("KPI_DIF_PESO_KG", False),
            ("KPI_CONCILIADO", False), ("CAR_CANT_MODEL_CODE", True), ("CAR_CANT_PARTIDAS", True),
            ("DOC_CANAL_DIAN", True), ("DOC_DEX", True), ("DOC_DEX_TIPO", True), ("ID_SAE_REEMPLAZADA", True),
            ("KPI_H_ELAB", False), ("KPI_ANS_ELAB", False), ("KPI_H_TOTAL_SAE", False), ("KPI_ANS_SAE", False),
            ("EST_CAUSAL_ID", True), ("KPI_ALERTA_CAUSAL", False)]
sae_rows = [
    dict(ID_SAE="6027770000011", ID_FILE_ADUANA=F, CAR_CANTIDAD_UND=3647, CAR_VALOR_FOB_USD=76754.29,
         CAR_PESO_NETO_KG=2102.55, CAR_PESO_BRUTO_KG=2486.79, CAR_BULTOS=212, DOC_PESO_BRUTO_MANIF_KG=2486.79,
         DOC_BULTOS_MANIF=212, CAR_CANT_MODEL_CODE=310, CAR_CANT_PARTIDAS=74, DOC_CANAL_DIAN="AUTORIZA_EMBARQUE",
         DOC_DEX="6007770000101", DOC_DEX_TIPO="AUTOMATICA"),
    dict(ID_SAE="6027770000028", ID_FILE_ADUANA=F, CAR_CANTIDAD_UND=6692, CAR_VALOR_FOB_USD=121983.36,
         CAR_PESO_NETO_KG=4934.10, CAR_PESO_BRUTO_KG=5724.23, CAR_BULTOS=418, DOC_PESO_BRUTO_MANIF_KG=5724.23,
         DOC_BULTOS_MANIF=418, CAR_CANT_MODEL_CODE=453, CAR_CANT_PARTIDAS=81, DOC_CANAL_DIAN="AUTORIZA_EMBARQUE",
         DOC_DEX="6007770000118", DOC_DEX_TIPO="AUTOMATICA"),
    dict(ID_SAE="6027770000035", ID_FILE_ADUANA=S, CAR_CANTIDAD_UND=1347, CAR_VALOR_FOB_USD=33083.82,
         CAR_PESO_NETO_KG=1462.20, CAR_PESO_BRUTO_KG=1708.47, CAR_BULTOS=192, DOC_PESO_BRUTO_MANIF_KG=1702.10,
         DOC_BULTOS_MANIF=190, CAR_CANT_MODEL_CODE=74, CAR_CANT_PARTIDAS=69, DOC_CANAL_DIAN="INSPECCION_FISICA",
         EST_CAUSAL_ID="SAE08"),
    dict(ID_SAE="6027770000042", ID_FILE_ADUANA=M, CAR_CANTIDAD_UND=288, CAR_VALOR_FOB_USD=13878.52,
         CAR_PESO_NETO_KG=944.30, CAR_PESO_BRUTO_KG=1095.41, CAR_BULTOS=136, CAR_CANT_MODEL_CODE=68, CAR_CANT_PARTIDAS=73),
]
recep = f"_xlfn.MAXIFS({E_FECHA},{E_FILE},B{{r}},{E_HITO},\"H000\")"
borr = f"_xlfn.MAXIFS({E_FECHA},{E_FILE},B{{r}},{E_HITO},\"H020\")"
saeok = f"_xlfn.MAXIFS({E_FECHA},{E_SAE},A{{r}},{E_HITO},\"H040\")"
sae_f = {
    "KPI_DIF_PESO_KG": "=IF(OR(A{r}=\"\",H{r}=\"\"),\"\",ROUND(H{r}-F{r},2))",
    "KPI_CONCILIADO": f"=IF(OR(A{{r}}=\"\",H{{r}}=\"\"),\"\",IF(AND(ABS(J{{r}})<={PARAM['TOL_PESO_KG']},I{{r}}=G{{r}}),\"S\",\"N\"))",
    "KPI_H_ELAB": f"=IF(OR(A{{r}}=\"\",{recep}=0,{borr}=0),\"\",ROUND(({borr}-{recep})*24,1))",
    "KPI_ANS_ELAB": f"=IF(A{{r}}=\"\",\"\",IF(R{{r}}=\"\",\"PENDIENTE\",IF(R{{r}}<={PARAM['ANS_ELAB_SAE_H']},\"CUMPLE\",\"NO_CUMPLE\")))",
    "KPI_H_TOTAL_SAE": f"=IF(OR(A{{r}}=\"\",{recep}=0,{saeok}=0),\"\",ROUND(({saeok}-{recep})*24,1))",
    "KPI_ANS_SAE": f"=IF(A{{r}}=\"\",\"\",IF(T{{r}}=\"\",\"PENDIENTE\",IF(T{{r}}<={PARAM['ANS_TOTAL_SAE_H']},\"CUMPLE\",\"NO_CUMPLE\")))",
    "KPI_ALERTA_CAUSAL": "=IF(AND(OR(S{r}=\"NO_CUMPLE\",U{r}=\"NO_CUMPLE\",K{r}=\"N\"),V{r}=\"\"),\"FALTA CAUSAL\",\"\")",
}
build_sheet(sae, sae_cols, sae_rows, "tbl_SAE", formulas=sae_f, widths={"ID_FILE_ADUANA": 22, "DOC_CANAL_DIAN": 22})
for r in range(2, N + 1):
    for c in "AOQ":
        sae[f"{c}{r}"].number_format = "@"
    for c in "DEFHJ":
        sae[f"{c}{r}"].number_format = "#,##0.00"
    for c in "CGILM":
        sae[f"{c}{r}"].number_format = "#,##0"
    for c in "RT":
        sae[f"{c}{r}"].number_format = '0.0 "h"'
add_list(sae, sae_cols, "ID_FILE_ADUANA", f"REMISION!$A$2:$A${N}")
add_list(sae, sae_cols, "DOC_CANAL_DIAN", L["CANAL_DIAN"])
add_list(sae, sae_cols, "DOC_DEX_TIPO", L["DEX_TIPO"])
add_list(sae, sae_cols, "EST_CAUSAL_ID", C["CAUSAL_ID"])
for letter in ("A", "O"):
    dv = DataValidation(type="textLength", operator="equal", formula1="13", allow_blank=True,
                        showErrorMessage=True, error="SAE / DEX: 13 dígitos, capturar como texto.")
    dv.add(f"{letter}2:{letter}{N}")
    sae.add_data_validation(dv)
pos = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                     showErrorMessage=True, error="Debe ser un número >= 0.")
for c in "CDEFGHI":
    pos.add(f"{c}2:{c}{N}")
sae.add_data_validation(pos)
for c, cols_ in (("S", "S"), ("U", "U")):
    sae.conditional_formatting.add(f"{c}2:{c}{N}", FormulaRule(formula=[f'${c}2="NO_CUMPLE"'], fill=PatternFill("solid", fgColor="F8696B")))
    sae.conditional_formatting.add(f"{c}2:{c}{N}", FormulaRule(formula=[f'${c}2="CUMPLE"'], fill=PatternFill("solid", fgColor="A9D08E")))
sae.conditional_formatting.add(f"K2:K{N}", FormulaRule(formula=['$K2="N"'], fill=PatternFill("solid", fgColor="F8696B")))
sae.conditional_formatting.add(f"W2:W{N}", FormulaRule(formula=['$W2<>""'], font=Font(name=FONT, bold=True, color="C00000")))
sae["R1"].comment = Comment("Horas reales recepción (H000) → borrador (H020), leídas de EVENTOS. Nunca se digitan.", "Plantilla")
sae["T1"].comment = Comment("Horas reales recepción (H000) → SAE autorizada (H040). Si falta un evento queda PENDIENTE, no se estima.", "Plantilla")

# ---------------------------------------------------------------- RESUMEN
rs = wb.create_sheet("RESUMEN", 1)
rs["A1"] = "Tablero rápido"
rs["A1"].font = Font(name=FONT, bold=True, size=14)
rs["A2"] = "Fecha de corte:"
rs["B2"] = f"={PARAM['FECHA_CORTE']}"
rs["B2"].number_format = "yyyy-mm-dd hh:mm"
heads = ["Fase", "VERDE", "AMBAR", "ROJO", "CERRADO", "Total"]
for j, h in enumerate(heads, 1):
    c = rs.cell(row=4, column=j, value=h)
    c.font = HDR_FONT
    c.fill = PatternFill("solid", fgColor=BLOCK_FILL["EST"])
fases = list(dict.fromkeys(h[3] for h in hitos))
for i, f in enumerate(fases, 5):
    rs.cell(row=i, column=1, value=f).font = BLACK
    for j, s in enumerate(["VERDE", "AMBAR", "ROJO", "CERRADO"], 2):
        rs.cell(row=i, column=j, value=f'=COUNTIFS(REMISION!$M$2:$M${N},$A{i},REMISION!$Y$2:$Y${N},"{s}")')
    rs.cell(row=i, column=6, value=f"=SUM(B{i}:E{i})")
tr = 5 + len(fases)
rs.cell(row=tr, column=1, value="Total").font = Font(name=FONT, bold=True)
for j in range(2, 7):
    L_ = get_column_letter(j)
    rs.cell(row=tr, column=j, value=f"=SUM({L_}5:{L_}{tr - 1})").font = Font(name=FONT, bold=True)
k = tr + 2
kpis = [
    ("% ANS elaboración SAE (≤ ANS_ELAB_SAE_H)",
     f'=IFERROR(COUNTIF(SAE!$S$2:$S${N},"CUMPLE")/(COUNTIF(SAE!$S$2:$S${N},"CUMPLE")+COUNTIF(SAE!$S$2:$S${N},"NO_CUMPLE")),"")', "0.0%"),
    ("% ANS total SAE (≤ ANS_TOTAL_SAE_H)",
     f'=IFERROR(COUNTIF(SAE!$U$2:$U${N},"CUMPLE")/(COUNTIF(SAE!$U$2:$U${N},"CUMPLE")+COUNTIF(SAE!$U$2:$U${N},"NO_CUMPLE")),"")', "0.0%"),
    ("Mediana horas recepción → SAE", f'=IFERROR(MEDIAN(SAE!$T$2:$T${N}),"")', '0.0 "h"'),
    ("SAE no conciliadas (peso/bultos)", f'=COUNTIF(SAE!$K$2:$K${N},"N")', "0"),
    ("SAE con incumplimiento sin causal", f'=COUNTIF(SAE!$W$2:$W${N},"FALTA CAUSAL")', "0"),
    ("Files abiertos en ROJO", f'=COUNTIF(REMISION!$Y$2:$Y${N},"ROJO")', "0"),
]
for i, (lbl, fml, fmt) in enumerate(kpis, k):
    rs.cell(row=i, column=1, value=lbl).font = Font(name=FONT, bold=True)
    c = rs.cell(row=i, column=2, value=fml)
    c.number_format = fmt
rs.column_dimensions["A"].width = 42
for c in "BCDEF":
    rs.column_dimensions[c].width = 12
for i in range(5, tr):
    rs.conditional_formatting.add(f"D{i}", FormulaRule(formula=[f"$D{i}>0"], fill=PatternFill("solid", fgColor="F8696B")))

wb.save(OUT)
print("OK", OUT)

# Rediseño de la Matriz de Control Logístico — Operación de Reembarques (SAE / DEX)

> **Alcance.** Propuesta de arquitectura de datos para reemplazar el `CONSOLIDADO_GENERAL` (hojas `DATA`, `Operacion Dollarcity`, `Reporte`, `Check List.`, `Causales`, `Rep Mensual`, `Estadisticas`, `Facturacion 2025`, `Doc. Legales`, etc.) por un modelo único, normalizado y listo para Excel (Tablas + Power Query) y Power BI.
>
> Plantilla operativa que acompaña este documento: [`plantilla/MATRIZ_CONTROL_LOGISTICO_v2.xlsx`](../plantilla/MATRIZ_CONTROL_LOGISTICO_v2.xlsx).

---

## 0. Diagnóstico del consolidado actual (hoja `DATA`, 4.011 filas, 2024‑2026)

Antes de rediseñar, medí el archivo actual. Estos son los problemas que la nueva estructura tiene que resolver:

| # | Hallazgo | Evidencia en el archivo | Impacto |
|---|----------|-------------------------|---------|
| 1 | **El KPI de SAE no mide la realidad.** En 3.065 filas (76 %) `FECHA APROBACION SAE (KN CL)` es la fórmula `=FECHA RECEPCION + 2` y en 3.115 filas `FECHA SAE` es `=FECHA APROBACION + TIME(2,40,0)`. | El 75 % de las filas tiene exactamente 50 h 40 min entre recepción y SAE. `CUMPLIMIENTO SAE` da 82 % `OUT TIME` por construcción. | El indicador principal de ANS está sesgado: no refleja ni el desempeño de KN ni el de la bodega. |
| 2 | **La causal no discrimina.** `CAUSAL DE INCUMPLIMIENTO` = `SAE04` en 4.002 de 4.004 filas. | La hoja `Causales` tiene 9 códigos (SAE01‑SAE09), pero en la práctica solo se usa uno. | No hay análisis de causa raíz posible. |
| 3 | **Granularidad mezclada.** Cada fila es una **SAE**, pero repite datos de la **remisión** (≈2,3 SAE por remisión) y del **reembarque** (≈4,8 remisiones por reembarque). | 4.011 filas → 1.751 remisiones/files → 369 reembarques. | Contar “files” o “contenedores” exige `DISTINCTCOUNT`; si alguien suma, cuenta de más. |
| 4 | **Catálogos libres.** Hay espacios, `NBSP` (`\xa0`), tildes y variantes de nombre. | `MAERSK` aparece de 4 formas; `LOGISTICA 360` de 5 (incluye el error `LOGISTICA 362`); `ZIM` de 3; `LUGAR DE INGRESO` tiene **38** valores distintos para 2 terminales; `OK CONFORMIDAD` con y sin espacio final; país/ciudad en texto libre (`ARGENTINA/ BUENOS AIRES ` ≠ `ARGENTINA/ BUENOS AIRES`). | Los filtros, las tablas dinámicas y Power BI fragmentan los totales. |
| 5 | **Estados escritos en columnas de fecha o número.** | `FECHA MANIFIESTO` contiene `inspección física` (68), `Autoriza embarque` (66), `Solicita Embarque` (6) y fechas como texto (`13-02-2025\xa0`). `DEX` contiene `DEX MANUAL`, `Autoriza embarque` y `Selección para inspección física`. | Rompen el tipo de dato y esconden **estados reales** del proceso (canal DIAN y DEX manual) que hoy no tienen columna propia. |
| 6 | **Formatos de fecha mezclados.** | `FECHA FACTURACION`: 1.189 seriales de Excel + 251 textos (`27.01.2025`, `26,04,2025`). `FECHA ENVIO BORRADOR`: 90 textos ISO con 7 decimales (`2026-07-09T15:33:11.0000000`). | No se pueden calcular tiempos ni filtrar por período. |
| 7 | **Duraciones guardadas como fecha.** | `TIEMPO … D / HH:MM` se ve como `1900-01-02 02:40`. | Por encima de 24 h el formato `HH:MM` engaña y Power BI las importa como fechas de 1900. |
| 8 | **Campos multivalor.** | `DIGITADOR = "KATERINE RODRIGUEZ - ANDRES FONSECA"`; `Packing list = "2900222429 & 2900222430"`. | No se puede medir productividad por persona ni cruzar con el WMS. |
| 9 | **Claves con información incrustada y sucia.** | `REMISION NUMERO = "BOOST-438.7 MX "` mezcla proyecto, reembarque, secuencia, país y a veces modo (`… UY AEREO`), con espacios extra. | Para unir tablas hay que hacer `TRIM` manual y es fácil equivocarse. |
| 10 | **Conciliación sin control.** | 91 SAE con `PESO BRUTO MANIFESTADO` ≠ `PESO BRUTO KG` (diferencia > 0,5 kg) y 89 con bultos distintos. Ninguna columna lo marca. | Riesgo aduanero que no se ve (correcciones de DEX y SAE de reemplazo). |
| 11 | **Columnas redundantes.** | `CANT. MODEL CODE` = `CANT. ITEM CODE` en el 99,4 % de las filas. `AÑO`, `MES` y `SEMANA` se escriben a mano aunque se derivan de la fecha. | Más digitación y más errores. |
| 12 | **Dos calendarios de corte sin modelar.** | Facturación del 25 al 25 frente a operación del 1 al 30/31 (`Rep Mensual`). | Los reportes mensuales no cuadran entre sí. |

**Conclusión:** el problema no es de celdas; es de **modelo**. Hoy una sola hoja intenta ser bitácora de eventos, maestro de catálogos, tablero de KPI y base de facturación al mismo tiempo.

---

## 1. Arquitectura de datos optimizada

### 1.1 Principio rector: “un hecho, un lugar”

Se pasa de **una hoja ancha de 53‑70 columnas** a un **modelo estrella** con tres niveles de negocio y una bitácora de eventos:

```
                          ┌──────────────────┐
                          │  dim_Calendario  │  (día, semana ISO, mes operativo,
                          └────────┬─────────┘   período facturación 25‑25, hábil S/N)
                                   │
┌──────────────┐      ┌────────────┴───────────┐       ┌──────────────────┐
│ dim_Cliente  │◄─────┤   tbl_Remision (FILE)  ├──────►│ dim_Destinatario │
└──────────────┘      │  1 fila = 1 file aduana│       │  + dim_Pais ISO  │
                      └───┬───────────────┬────┘       └──────────────────┘
                          │1             N│
          ┌───────────────┘               └──────────────┐
          │N                                            1│
┌─────────┴──────────┐                       ┌───────────┴──────────┐
│  tbl_SAE (HECHOS)  │                       │   tbl_Reembarque     │
│ 1 fila = 1 SAE/DEX │                       │ 1 fila = 1 booking / │
│ pesos, bultos, USD │                       │ contenedor / zarpe   │
└─────────┬──────────┘                       └──────────┬───────────┘
          │1                                            │
          │N                                            │
┌─────────┴─────────────────────────────────────────────┴──────────┐
│  tbl_Eventos  (bitácora larga: ID_OBJETO, HITO, FECHA_HORA, ...)  │
│  Fuente única de TODAS las fechas → los tiempos y KPI se derivan  │
└──────────────────────────────────────────────────────────────────┘
   Catálogos: cat_Estado · cat_Hito · cat_Causal · cat_Transportista ·
              cat_Terminal · cat_Persona · cat_Pais · cat_TipoDocumento · cat_Parametros_ANS
```

**Por qué es innovador para esta operación:**

1. **Las fechas dejan de ser columnas y pasan a ser eventos.** Hoy, cada hito nuevo (por ejemplo “inspección física” o “SAE de reemplazo”) obliga a crear columnas o a escribir texto en una columna de fecha. En `tbl_Eventos` un hito nuevo es **una fila más con un código de catálogo**. La vista ancha se reconstruye sola con Power Query (Pivot) o con DAX.
2. **Los KPI no se digitan: se calculan** desde eventos reales con marca de tiempo. Así se elimina el problema #1 de raíz: si el evento “Aprobación bodega” no existe, el KPI queda *pendiente*, no *inventado*.
3. **Cada nivel cuenta lo suyo:** contenedores en `tbl_Reembarque`, files en `tbl_Remision` y SAE, partidas y valores en `tbl_SAE`. No más `DISTINCTCOUNT` para corregir duplicados.
4. **Estado = función de los eventos.** El estado macro se deriva del último hito cumplido (ver §3), así no se desincroniza con las fechas.

### 1.2 Bloques de columnas por tabla

Convención de nombres: `PREFIJO_NOMBRE` en MAYÚSCULAS, sin tildes ni espacios. El prefijo indica el bloque lógico. Esto permite que Power BI agrupe en carpetas de visualización y que las fórmulas no se rompan por un espacio final.

| Prefijo | Bloque lógico |
|---|---|
| `ID_` | Claves y referencias |
| `CTL_` | Control y auditoría (quién, cuándo, fuente) |
| `RUT_` | Origen, destino y ruta |
| `TRN_` | Transporte (agente, naviera, buque, BL, contenedor) |
| `CAR_` | Carga (pesos, bultos, valores, partidas) |
| `DOC_` | Documentación y aduanas (SAE, DEX, manifiesto, póliza) |
| `F_` | Fechas de hitos (en la vista ancha; en el modelo viven en `tbl_Eventos`) |
| `EST_` | Estados y alertas |
| `KPI_` | Métricas calculadas — **nunca se digitan** |
| `FAC_` | Facturación |

#### A) `tbl_Reembarque` — 1 fila por reembarque (booking / contenedor / zarpe)

| Bloque | Campo | Tipo | Regla / catálogo | Origen actual |
|---|---|---|---|---|
| ID | `ID_REEMBARQUE` | Texto | `BOOSTERIII-406` → normalizado `BST3-0406` | `# REEMBARQUE` |
| ID | `ID_PROYECTO` | Cat. | `BOOSTER`, `RT`, `DOLLARCITY` | prefijo del reembarque |
| TRN | `TRN_MODO` | Cat. | `MAR`, `AER`, `TER` | `Modo de Transporte` |
| TRN | `TRN_AGENTE_CARGA_ID` | Cat. | `cat_Transportista` (tipo = FORWARDER) | `AGENTE DE CARGA` |
| TRN | `TRN_NAVIERA_ID` | Cat. | `cat_Transportista` (tipo = CARRIER, con SCAC) | `LINEA NAVIERA` |
| TRN | `TRN_BOOKING` | Texto | n.º de reserva | `Asunto` (hoy viene del asunto del correo) |
| TRN | `TRN_SEMANA_RESERVA` | Texto | `2026-W37` (ISO 8601) | `… W37` dentro del asunto |
| TRN | `TRN_BUQUE_VIAJE` | Texto | `MAERSK XXX / 437N` | `(CTG) UVI ‑ MOTONAVE` |
| TRN | `TRN_BL` | Texto | sin espacios | `(CTG) B.L EXPORTACION` |
| TRN | `TRN_CONTENEDOR` | Texto | ISO 6346: `AAAU1234567` (4 letras + 7 dígitos, con dígito de control) | `(CTG) N° CONTENEDOR EXP` |
| TRN | `TRN_TIPO_CONTENEDOR` | Cat. | ISO 6346 size‑type: `22G1`, `42G1`, `45G1` | nuevo |
| RUT | `RUT_PUERTO_EMBARQUE` | Cat. | UN/LOCODE: `COCTG` | nuevo (implícito) |
| RUT | `RUT_TERMINAL_ID` | Cat. | `CONTECAR`, `SPRC` | `LUGAR DE INGRESO` (parte texto) |
| RUT | `RUT_PUERTO_DESTINO` | Cat. | UN/LOCODE: `CLVAP`, `MXZLO`, `UYMVD`… | `Asunto` (`COCTG - CLVAP`) |
| F | `F_ETD`, `F_ZARPE_REAL`, `F_ETA` | Fecha‑hora | ISO 8601 | `(CTG) FECHA DE ZARPE` |
| DOC | `DOC_POLIZA_NUM`, `DOC_POLIZA_RAD_ACEPTACION`, `DOC_POLIZA_RAD_CANCELACION`, `DOC_POLIZA_CERT_CANCELACION` | Texto | — | bloque `(CTG) … POLIZA` (Dollarcity) |
| CTL | `CTL_CREADO_POR`, `CTL_FECHA_CREACION`, `CTL_ULT_MODIF` | Auditoría | — | nuevo |

#### B) `tbl_Remision` — 1 fila por remisión / file de aduana

| Bloque | Campo | Tipo | Regla / catálogo | Origen actual |
|---|---|---|---|---|
| ID | `ID_FILE_ADUANA` | Texto (clave) | patrón `NNNNNNNNNN-0223-NNN` (3 dígitos con cero a la izquierda: `-030`, no `-30`) | `FILE ADUANA` |
| ID | `ID_REMISION` | Texto | `BOOST-438.7-MX`, sin espacios | `REMISION NUMERO` |
| ID | `ID_REEMBARQUE` | FK | → `tbl_Reembarque` | `# REEMBARQUE` |
| ID | `ID_REMISION_SECUENCIA` | Entero | `7` (de `438.7`) | parseado |
| RUT | `RUT_DESTINATARIO_ID` | Cat. | `cat_Destinatario` (razón social única + NIT/Tax ID) | `DESTINO` |
| RUT | `RUT_PAIS_DESTINO` | Cat. | ISO 3166‑1 alfa‑2: `CL`, `MX`, `UY`, `AR`, `PA`, `CR`, `PY`, `SV`, `IE`, `US`, `CA` | `PAIS / CIUDAD` |
| RUT | `RUT_CIUDAD_DESTINO` | Cat. | UN/LOCODE de ciudad: `CLSCL`, `MXMEX`… | `PAIS / CIUDAD` |
| DOC | `DOC_PACKING_LIST` | Texto | **uno por fila** en `tbl_RemisionPL` si hay varios | `Packing list` |
| DOC | `DOC_REQUIERE_CLASIFICACION` | S/N | — | `CANT. MODEL CODE PARA CLASIFICACION` > 0 |
| DOC | `DOC_CANT_MODEL_CLASIFICAR` | Entero | — | ídem |
| CTL | `CTL_DIGITADOR_ID` | Cat. | `cat_Persona` — **una persona**; el apoyo va en `CTL_APOYO_ID` | `DIGITADOR` |
| CTL | `CTL_RESPONSABLE_DEX_ID` | Cat. | `cat_Persona` | `RESPONSABLE` |
| EST | `EST_FASE` | Cat. (calculado) | ver §3 | `ESTADO` + `ESTADO GESTOR DOCUMENTAL` |
| EST | `EST_REVISION` | Cat. | `EN_DIGITACION`, `EN_REVISION`, `PDTE_VALORES`, `CONFORME` | `ESTADO` |
| EST | `EST_GESTOR_DOC` | Cat. | `CARPETA_CREADA`, `PDTE`, `FALTA_DEX`, `COMPLETO` | `ESTADO GESTOR DOCUMENTAL` |
| EST | `EST_BLOQUEO_CAUSAL_ID` | Cat. | `cat_Causal` | nuevo |
| FAC | `FAC_PERIODO` | Texto | `2026-09` (corte 25‑25, derivado de `dim_Calendario`) | `FECHA FACTURACION` |
| FAC | `FAC_NUM_PREFACTURA`, `FAC_VALOR_USD`, `FAC_TRM` | — | — | `Facturacion 2025` |
| CTL | `CTL_OBSERVACION` | Texto libre | **única** columna de texto libre; complementa la causal, no la reemplaza | `OBSERVACIONES BOG` |

#### C) `tbl_SAE` — 1 fila por SAE (tabla de hechos principal)

| Bloque | Campo | Tipo | Regla / catálogo | Origen actual |
|---|---|---|---|---|
| ID | `ID_SAE` | Texto 13 dígitos | `^602\d{10}$` | `SAE` |
| ID | `ID_FILE_ADUANA` | FK | → `tbl_Remision` | `FILE ADUANA` |
| ID | `ID_SAE_REEMPLAZADA` | Texto | SAE original si esta es de reemplazo | `SAE INICIAL CON REEMPLAZO` |
| CAR | `CAR_CANTIDAD_UND` | Entero | ≥ 0 | `CANTIDADES` |
| CAR | `CAR_VALOR_FOB_USD` | Decimal (2) | ≥ 0, moneda ISO 4217 en `CAR_MONEDA` | `VALORTOTAL` |
| CAR | `CAR_PESO_NETO_KG` / `CAR_PESO_BRUTO_KG` | Decimal (2) | neto ≤ bruto | `PESO NETO/BRUTO KG` |
| CAR | `CAR_BULTOS` | Entero | > 0 | `BULTOS` |
| CAR | `CAR_CANT_MODEL_CODE` | Entero | — | `CANT. MODEL CODE` (se elimina `ITEM CODE` si siempre es igual) |
| CAR | `CAR_CANT_PARTIDAS` | Entero | — | `CANT. PARTIDAS ARANCELARIAS` |
| DOC | `DOC_PESO_BRUTO_MANIF_KG`, `DOC_BULTOS_MANIF` | Decimal / Entero | — | `PESO/BULTOS MANIFESTADOS` |
| DOC | `DOC_CANAL_DIAN` | Cat. | `AUTORIZA_EMBARQUE`, `INSPECCION_FISICA`, `INSPECCION_DOCUMENTAL` | **hoy escrito como texto dentro de `FECHA MANIFIESTO` / `DEX`** |
| DOC | `DOC_DEX` | Texto 13 dígitos | `^600\d{10}$` | `DEX` |
| DOC | `DOC_DEX_TIPO` | Cat. | `AUTOMATICA`, `MANUAL`, `CORRECCION` | `DEX MANUAL`, `CORECCION DEX`, hoja `Dex Cierre Manual` |
| DOC | `DOC_DEX_CORRECCION` | Texto | — | `CORECCION DEX` |
| KPI | `KPI_DIF_PESO_KG`, `KPI_DIF_BULTOS`, `KPI_CONCILIADO` | Calc. | tolerancia en `cat_Parametros_ANS` | nuevo |
| KPI | `KPI_H_RECEPCION_BORRADOR`, `KPI_H_REVISION_BODEGA`, `KPI_H_TOTAL_SAE`, `KPI_D_MANIF_DEX` | Calc. (horas decimales) | desde `tbl_Eventos` | `TIEMPO … D / HH:MM` |
| KPI | `KPI_ANS_SAE`, `KPI_ANS_ELAB`, `KPI_ANS_DEX`, `KPI_ANS_CLASIF` | Calc. | `CUMPLE` / `NO_CUMPLE` / `PENDIENTE` | `CUMPLIMIENTO …` |
| EST | `EST_CAUSAL_ID` | Cat. | **obligatoria si `NO_CUMPLE`** | `CAUSAL DE INCUMPLIMIENTO` |

#### D) `tbl_Eventos` — bitácora de hitos (formato largo)

| Campo | Tipo | Ejemplo |
|---|---|---|
| `ID_EVENTO` | Autonumérico | 10234 |
| `NIVEL` | Cat. | `REEMBARQUE` / `REMISION` / `SAE` |
| `ID_OBJETO` | Texto | `6027757362143` |
| `HITO_ID` | Cat. (`cat_Hito`) | `H040_BORRADOR_ENVIADO` |
| `FECHA_HORA` | Fecha‑hora ISO 8601, zona America/Bogota | `2026-09-14 10:38` |
| `RESPONSABLE_ID` | Cat. | `P003` |
| `FUENTE` | Cat. | `MANUAL`, `CORREO`, `MUISCA`, `WMS`, `NAVIERA` |
| `CAUSAL_ID` | Cat. (opcional) | `SAE05` |
| `NOTA` | Texto | libre, corto |

> **Regla de oro:** ninguna fecha de la vista ancha se escribe a mano. Toda fecha es un evento; la vista ancha la lee con `MAXIFS` (Excel) o con Pivot (Power Query).

---

## 2. Normalización y estandarización

### 2.1 Formatos estándar

| Tipo de dato | Estándar | Formato de almacenamiento | Formato de visualización | Validación |
|---|---|---|---|---|
| Fecha‑hora de hito | ISO 8601, hora local America/Bogota (UTC‑5) | Serial de Excel / `datetime` | `yyyy-mm-dd hh:mm` | Validación de datos: *Fecha* entre `2018-01-01` y `HOY()+180` |
| Fecha (sin hora) | ISO 8601 | Serial | `yyyy-mm-dd` | ídem |
| Duración | **Horas decimales** (`50.67`) o días decimales | Número | `0.0 "h"` | ≥ 0; un valor negativo alerta de secuencia invertida |
| Semana | ISO 8601 | Derivada: `ISOWEEKNUM` | `2026-W37` | no editable |
| Mes operativo / período de facturación | Derivados de `dim_Calendario` | Texto `yyyy-mm` | — | no editable |
| Moneda | ISO 4217 (`USD`, `COP`) | Texto 3 | — | lista |
| Valores monetarios | Decimal, 2 posiciones; TRM en tabla aparte por fecha | Número | `#,##0.00` | ≥ 0 |
| Pesos | kg, 2 decimales | Número | `#,##0.00` | neto ≤ bruto |
| País | ISO 3166‑1 alfa‑2 | Texto 2 | nombre por relación | lista |
| Puertos / ciudades | UN/LOCODE (5 caracteres) | Texto 5 | nombre por relación | lista |
| Contenedor | ISO 6346 | Texto 11 | — | longitud 11 + patrón |
| Tipo de contenedor | ISO 6346 size‑type | Texto 4 | — | lista |
| Transportista marítimo | Código SCAC (`MAEU`, `HLCU`, `ZIMU`, `ONEY`, `CMDU`, `EGLV`) | Texto 4 | nombre por relación | lista |
| Personas | Código interno `P001…` | Texto | nombre por relación | lista |
| SAE / DEX | Numérico de 13 dígitos guardado como **texto** (para no perder ceros ni pasar a notación científica) | Texto | — | `LARGO = 13` |
| File aduana | `NIT-0223-NNN` | Texto | — | patrón con 3 dígitos |
| Booleanos | `S` / `N` | Texto 1 | ✓ / ✗ con formato condicional | lista |

### 2.2 Catálogos (listas cerradas)

Todos están en la plantilla (hoja `CATALOGOS`) como **Tablas de Excel con nombre**, para que las listas desplegables crezcan solas.

**`cat_Transportista`** — cierra las 18 variantes de agente y las 17 de naviera actuales:

| ID | NOMBRE_NORMALIZADO | TIPO | SCAC | Alias detectados en el consolidado |
|---|---|---|---|---|
| `MAEU` | MAERSK | CARRIER/FORWARDER | MAEU | `MAERSK`, `\xa0MAERSK`, `MAERSK\xa0`, `\xa0MAERSK\xa0` |
| `HLCU` | HAPAG‑LLOYD | CARRIER | HLCU | `HAPAG LLOYD`, `HAPAG LLOYD CONTAINER LINE`, `HAPAG LLOYD ` |
| `ZIMU` | ZIM | CARRIER | ZIMU | `ZIM ISRAEL NAVIGATION`, `ZIM ISRAELNAVIGATION`, `…\xa0` |
| `ONEY` | OCEAN NETWORK EXPRESS | CARRIER | ONEY | `OCEAN NETWORKEXPRESS`, `OCEAN NETWORK EXPRESS` |
| `CMDU` | CMA CGM | CARRIER/FORWARDER | CMDU | `CMA CGM`, `CMA CGM INC` |
| `EGLV` | EVERGREEN | CARRIER | EGLV | `EVERGREEN` |
| `KN` | KUEHNE+NAGEL | FORWARDER | — | `KUEHNE NAGEL`, `KUEHNE+NAGEL`, `\xa0KUEHNE NAGEL` |
| `L360` | LOGÍSTICA 360 | FORWARDER | — | `LOGISTICA 360`, `LOGÍSTICA 360`, `\xa0LOGÍSTICA 360`, `LOGISTICA 362` |
| `DHL` | DHL | FORWARDER | — | `DHL` |
| `COPA` | COPA AIRLINES CARGO | CARRIER (AER) | — | `COPA` |

> `PDTE` y `PANAMA` **no son transportistas**. `PDTE` se convierte en campo vacío + `EST_FASE` < reserva, y `PANAMA` es un error de digitación que hay que corregir en la fuente.

**`cat_Terminal`**: `CONTECAR`, `SPRC` (Sociedad Portuaria Regional de Cartagena). El sufijo actual (`SPRCTG 1`, `CONTECAR 2`, `SPRC14`, `… TX`, `… AEREO`) se separa en `RUT_TERMINAL_ID` + `TRN_SECUENCIA_INGRESO` (entero) + `TRN_MODO`. *Hay que validar con la operación qué representa exactamente el número (ingreso, vehículo o lote).*

**`cat_Causal`** — se conserva la codificación `SAE0x`, que ya conoce el equipo, y se amplía a todo el proceso. Cada causal tiene un **responsable tipificado**, lo que permite medir de quién es la demora:

| Familia | Códigos | Responsable tipo |
|---|---|---|
| Recepción de remisión | `REM01` remisión incompleta · `REM02` datos errados · `REM03` cambio de remisión · `REM04` recibida fuera de horario o en fin de semana | CLIENTE / CDLI |
| Clasificación | `CLA01` model code sin clasificar · `CLA02` respuesta tardía del cliente · `CLA03` PA sugerida difiere | CLIENTE / AA KN |
| Elaboración SAE (existentes) | `SAE01` … `SAE09` (tal cual están en la hoja `Causales`) | AA KN / KN CDLI / TERCERO |
| Rechazo / reemplazo SAE | `SAE10` rechazo DIAN · `SAE11` error de pesos · `SAE12` error de bultos · `SAE13` error de valor FOB · `SAE14` error de ítems | AA KN / CDLI |
| Puerto / DIAN | `PTO01` inspección física · `PTO02` cierre de terminal · `PTO03` roll‑over de naviera | DIAN / TERMINAL / NAVIERA |
| DEX | `DEX01` DEX manual por contingencia · `DEX02` corrección de DEX · `DEX03` diferencia manifiesto vs SAE | AA KN / DIAN |

**`cat_Persona`**: `ID`, nombre, rol (`DIGITADOR`, `RESPONSABLE_DEX`, `REVISOR`), activo S/N. Desaparecen los valores compuestos `A - B`.

**`cat_Destinatario`**: razón social única + país ISO + Tax ID. Unifica, por ejemplo, `BROSPIL S,A,` / `BROSPIL S.A.` y `ARTICULOS DEPORTIVOS DECATHLON SA DE CV` / `… MEXICO`.

**`cat_Parametros_ANS`**: los umbrales **no se escriben dentro de las fórmulas** (hoy están quemados en `TIME(4,0,0)` y `1 + TIME(4,0,0)`):

| Parámetro | Valor actual (inferido de las fórmulas) | Unidad |
|---|---|---|
| `ANS_ELAB_SAE_H` | 4 | horas, desde recepción hasta borrador |
| `ANS_TOTAL_SAE_H` | 28 | horas, desde recepción hasta SAE (`1 + TIME(4,0,0)`) |
| `ANS_CLASIF_H` | 3 | horas (columna `CUMPLIMIENTO CLASIFICACION (3 HRS)`) |
| `ANS_DEX_D` | *definir según la norma vigente y el contrato* | días, desde manifiesto hasta DEX |
| `TOL_PESO_KG` | *definir* | kg de diferencia SAE vs manifiesto |
| `MEDIR_EN_HORAS_HABILES` | S/N | aplica el calendario laboral, para que la causal SAE09 “fin de semana” no penalice |

### 2.3 Reglas de limpieza (Power Query, paso “Normalizar”)

1. `Text.Trim` + `Text.Clean` + reemplazar `#(00A0)` (NBSP) por espacio → `Text.Upper` → quitar tildes con una tabla de mapeo.
2. Mapear cada texto libre contra la columna **Alias** de su catálogo y, si no aparece, enviarlo a la hoja `ERRORES_CALIDAD`, no al modelo.
3. Las columnas de fecha que traen texto: si el texto está en `cat_Canal` (`inspección física`, `Autoriza embarque`, `Solicita Embarque`, `Selección para inspección física`), **moverlo** a `DOC_CANAL_DIAN` y dejar la fecha vacía (pendiente). Si es una fecha en texto (`13-02-2025`, `27.01.2025`, `26,04,2025`), convertirla con `Date.FromText` y la cultura `es-CO`.
4. `DEX MANUAL` en la columna `DEX` → `DOC_DEX_TIPO = MANUAL` y número vacío.
5. Valores compuestos (`A - B`, `X & Y`) → dividir en filas de una tabla puente.
6. Las fórmulas que inventan fechas (`=RECEPCION+2`, `+TIME(2,40,0)`) **se eliminan**. Esas filas se marcan con `CTL_FECHA_ESTIMADA = S` para excluirlas del KPI histórico o reportarlas por separado.

---

## 3. Secciones clave y estados operativos

### 3.1 Máquina de estados (fase macro)

Se reemplazan las dos columnas actuales (`ESTADO` y `ESTADO GESTOR DOCUMENTAL`), que mezclan la revisión interna con la documentación, por **una fase macro calculada** a partir de los hitos, más **sub‑estados** independientes:

```
 F00 ─► F10 ─► F20 ─► F30 ─► F40 ─► F50 ─► F60 ─► F70 ─► F80 ─► F90
 RECIB  CLASIF BORRAD APROB  SAE    PUERTO EMBARC DEX    CIERRE FACTUR
        (opc.)  SAE   BODEGA AUTORIZ CANAL  ZARPE  PRESENT DOC    ADO
                                   │
                                   └─► F45 INSPECCIÓN FÍSICA (desvío, vuelve a F50)
 Estados terminales de excepción:  X1 SAE REEMPLAZADA · X2 ANULADO
```

| Código | Fase | Hito que la abre (`cat_Hito`) | Criterio de salida (“Definition of Done”) | Dueño |
|---|---|---|---|---|
| `F00` | Remisión recibida | `H000_REMISION_RECIBIDA` | packing list + destinatario + reembarque | CDLI → AA KN |
| `F10` | Clasificación *(solo si aplica)* | `H010_CLASIF_SOLICITADA` | `H015_CLASIF_RESPONDIDA` | AA KN / Cliente |
| `F20` | Borrador SAE enviado | `H020_BORRADOR_ENVIADO` | aprobación de bodega | AA KN |
| `F30` | Aprobado por bodega | `H030_APROBACION_CDLI` | SAE autorizada | KN CDLI |
| `F40` | SAE autorizada | `H040_SAE_AUTORIZADA` | ingreso a puerto | AA KN |
| `F45` | Inspección física | `H045_INSPECCION_FISICA` | `H046_INSPECCION_LIBERADA` | DIAN / Terminal |
| `F50` | En puerto / autoriza embarque | `H050_AUTORIZA_EMBARQUE` | manifiesto | Terminal |
| `F60` | Embarcado (manifiesto / zarpe) | `H060_MANIFIESTO` | DEX presentada | Naviera |
| `F70` | DEX presentada | `H070_DEX_PRESENTADA` | gestor documental completo (+ póliza cancelada, si aplica) | AA KN |
| `F80` | Cierre documental | `H080_DOC_COMPLETOS` | prefactura | AA KN |
| `F90` | Facturado | `H090_FACTURADO` | — | Administración |

**Fórmula de fase:** el código del **último hito cumplido** en orden (en Excel: `MAXIFS` sobre el número de fase de los eventos del objeto; en DAX: `MAXX(RELATEDTABLE(tbl_Eventos), RELATED(cat_Hito[ORDEN]))`).

### 3.2 Sub‑estados ortogonales (no compiten con la fase)

| Dimensión | Valores | Uso |
|---|---|---|
| `EST_REVISION` | `EN_DIGITACION` · `EN_REVISION` · `PDTE_VALORES` · `CONFORME` | Control de calidad interno (equivale al `ESTADO` actual) |
| `EST_GESTOR_DOC` | `CARPETA_CREADA` · `PDTE` · `FALTA_DEX` · `COMPLETO` | Archivo documental |
| `EST_BLOQUEO` | `S`/`N` + `EST_BLOQUEO_CAUSAL_ID` | Detiene el semáforo en rojo, sin importar la fase |
| `DOC_CANAL_DIAN` | `AUTORIZA_EMBARQUE` · `INSPECCION_FISICA` · `INSPECCION_DOCUMENTAL` | Hoy está escondido en las columnas de fecha |
| `KPI_CONCILIADO` | `S`/`N` | Peso y bultos SAE = manifiesto (dentro de la tolerancia) |

### 3.3 Semáforo de un vistazo

Hay una sola columna `EST_SEMAFORO`, calculada con esta prioridad (gana la primera regla que se cumple):

| Prioridad | Condición | Semáforo |
|---|---|---|
| 1 | `EST_BLOQUEO = S` **o** `KPI_CONCILIADO = N` **o** fase `F45` | 🔴 **ROJO** |
| 2 | Fase ≥ `F80` | ⚪ **CERRADO** (gris) |
| 3 | Horas en la fase actual > ANS de la fase | 🔴 **ROJO** |
| 4 | Horas en la fase actual > 75 % del ANS | 🟡 **ÁMBAR** |
| 5 | Resto | 🟢 **VERDE** |

El ANS de cada fase vive en `cat_Hito[ANS_FASE_H]`, así que se ajusta sin tocar fórmulas.

Complementos de lectura rápida para la vista operativa:

- **`EST_PROGRESO`** — barra textual `●●●●●○○○○○` (`=REPT("●",n)&REPT("○",10-n)`), que muestra el avance del flujo sin abrir ningún filtro.
- **`EST_PROXIMA_ACCION`** y **`EST_DUENO_ACTUAL`** — se derivan de `cat_Hito` (qué sigue y quién lo hace). El tablero responde así “¿quién tiene la bola?”.
- **`EST_HORAS_EN_FASE`** — envejecimiento: el tiempo que lleva el ítem atascado.

### 3.4 Agrupaciones que más rinden

| Agrupación | Por qué |
|---|---|
| **Pesos, bultos y valores de SAE vs manifiesto, lado a lado** con `KPI_DIF_*` | Hace visible el riesgo de corrección de DEX (hoy hay 91 casos ocultos) |
| **Las cuatro fechas de cada ANS juntas** (inicio → fin → horas → cumple → causal) | Cada KPI se lee de izquierda a derecha como una oración |
| **Bloque `(CTG)` de póliza** en `tbl_Reembarque` | Hoy solo existe en la hoja Dollarcity; queda disponible para cualquier proyecto que la requiera |
| **Facturación separada** (`FAC_*`) con período 25‑25 | Concilia `Rep Mensual` (operativo) con `Facturacion 2025` sin dos hojas paralelas |
| **`Doc. Legales`** como tabla satélite `tbl_DocLegal` (documento, terminal, emisión, vencimiento, días para vencer, semáforo) | Hoy todos los documentos aparecen `VENCIDO` y nadie lo nota; con una alerta a 30 días se vuelve preventivo |

---

## 4. Ejemplo práctico — reembarque marítimo Cartagena → Valparaíso

**Escenario (ficticio, con la misma forma que la operación real):** reembarque `BST3-0412`, Maersk, terminal CONTECAR, contenedor 40' HC. Agrupa dos remisiones con destino Chile y una con destino México que salió en otro buque. Fecha de corte del tablero: **2026‑10‑04 08:00**.

### 4.1 `tbl_Reembarque`

| ID_REEMBARQUE | ID_PROYECTO | TRN_MODO | TRN_AGENTE | TRN_NAVIERA | TRN_BOOKING | TRN_SEMANA | RUT_TERMINAL | RUT_POL | RUT_POD | TRN_CONTENEDOR | TRN_TIPO | F_ETD | F_ZARPE_REAL |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| BST3-0412 | BOOSTER | MAR | MAEU | MAEU | 268114502 | 2026-W40 | CONTECAR | COCTG | CLVAP | MSKU7340215 | 45G1 | 2026-10-03 | 2026-10-03 22:40 |
| BST3-0413 | BOOSTER | MAR | L360 | HLCU | HLCU-CTG-88231 | 2026-W40 | SPRC | COCTG | MXZLO | HLXU8812043 | 45G1 | 2026-10-06 | — |

### 4.2 `tbl_Remision` (vista operativa, con estado y semáforo)

| ID_FILE_ADUANA | ID_REMISION | ID_REEMBARQUE | RUT_PAIS | RUT_DESTINATARIO | CTL_DIGITADOR | EST_FASE | EST_PROGRESO | EST_REVISION | EST_GESTOR_DOC | EST_DUENO_ACTUAL | EST_PROXIMA_ACCION | EST_HORAS_EN_FASE | EST_SEMAFORO |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 900000001-0223-412 | BOOST-412.1-CL | BST3-0412 | CL | DECATHLON CHILE SPA | P002 | F70 DEX presentada | ●●●●●●●●○○ | CONFORME | PDTE | AA KN | Cerrar carpeta | 9.3 | 🟢 VERDE |
| 900000001-0223-413 | BOOST-412.2-CL | BST3-0412 | CL | DECATHLON CHILE SPA | P001 | F60 Embarcado | ●●●●●●●○○○ | CONFORME | FALTA_DEX | AA KN | Conciliar y presentar DEX | 9.3 | 🔴 ROJO (no conciliado) |
| 900000001-0223-414 | BOOST-413.1-MX | BST3-0413 | MX | ART. DEP. DECATHLON SA DE CV | P003 | F30 Aprobado bodega | ●●●●○○○○○○ | EN_REVISION | PDTE | AA KN | Autorizar SAE | 3.2 | 🟡 ÁMBAR |

### 4.3 `tbl_SAE` (hechos + conciliación + ANS)

| ID_SAE | ID_FILE | CAR_BULTOS | CAR_PESO_BRUTO_KG | DOC_BULTOS_MANIF | DOC_PESO_MANIF_KG | KPI_DIF_PESO | KPI_CONCILIADO | DOC_CANAL_DIAN | DOC_DEX | DOC_DEX_TIPO | KPI_H_ELAB (≤4) | KPI_H_TOTAL_SAE (≤28) | KPI_ANS_SAE | EST_CAUSAL |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 6027770000011 | …-412 | 212 | 2,486.79 | 212 | 2,486.79 | 0.00 | S | AUTORIZA_EMBARQUE | 6007770000101 | AUTOMATICA | 2.6 | 21.4 | CUMPLE | — |
| 6027770000028 | …-412 | 418 | 5,724.23 | 418 | 5,724.23 | 0.00 | S | AUTORIZA_EMBARQUE | 6007770000118 | AUTOMATICA | 2.6 | 21.4 | CUMPLE | — |
| 6027770000035 | …-413 | 192 | 1,708.47 | 190 | 1,702.10 | **−6.37** | **N** | INSPECCION_FISICA | — | — | 5.1 | 30.2 | NO_CUMPLE | `SAE08` error en la remisión |
| 6027770000042 | …-414 | 136 | 1,095.41 | — | — | — | — | — | — | — | 3.4 | *pendiente* | PENDIENTE | — |

### 4.4 `tbl_Eventos` — trazabilidad completa del file `…-412`

| FECHA_HORA | NIVEL | ID_OBJETO | HITO | RESPONSABLE | FUENTE | CAUSAL | NOTA |
|---|---|---|---|---|---|---|---|
| 2026-09-29 08:15 | REMISION | 900000001-0223-412 | H000 Remisión recibida | CDLI | CORREO | — | PL 2900230011 |
| 2026-09-29 10:51 | REMISION | 900000001-0223-412 | H020 Borrador SAE enviado | P002 | CORREO | — | 2 SAE |
| 2026-09-29 15:30 | REMISION | 900000001-0223-412 | H030 Aprobación bodega CDLI | CDLI | CORREO | — | — |
| 2026-09-30 05:40 | SAE | 6027770000011 | H040 SAE autorizada | P002 | MUISCA | — | — |
| 2026-09-30 05:40 | SAE | 6027770000028 | H040 SAE autorizada | P002 | MUISCA | — | — |
| 2026-10-01 09:10 | REEMBARQUE | BST3-0412 | H050 Autoriza embarque | TERMINAL | MUISCA | — | Canal: autoriza embarque |
| 2026-10-03 22:40 | REEMBARQUE | BST3-0412 | H060 Manifiesto / zarpe | NAVIERA | NAVIERA | — | Buque/viaje 640S |
| 2026-10-03 22:40 | SAE | 6027770000011 | H070 DEX presentada | P005 | MUISCA | — | — |
| 2026-10-03 22:40 | SAE | 6027770000028 | H070 DEX presentada | P005 | MUISCA | — | — |

Con estos eventos, la fila de la remisión **se calcula sola**: fase `F70` (último hito), horas de elaboración = 10:51 − 08:15 = **2,6 h** (cumple ≤ 4 h) y horas totales hasta la SAE = 05:40 (+1 día) − 08:15 = **21,4 h** (cumple ≤ 28 h). No hay que escribir ningún “ON TIME” a mano.

### 4.5 Tablero que se obtiene (Power BI o una hoja `DASHBOARD` en Excel)

| Indicador | Medida (DAX de referencia) |
|---|---|
| % ANS SAE real | `DIVIDE(COUNTROWS(FILTER(tbl_SAE, [KPI_ANS_SAE]="CUMPLE")), COUNTROWS(FILTER(tbl_SAE, [KPI_ANS_SAE]<>"PENDIENTE")))` |
| Mediana de horas recepción → SAE | `MEDIANX(tbl_SAE, [KPI_H_TOTAL_SAE])` |
| Pareto de causales por responsable | conteo de `EST_CAUSAL_ID` × `cat_Causal[RESPONSABLE_TIPO]` |
| SAE no conciliadas | `CALCULATE(COUNTROWS(tbl_SAE), tbl_SAE[KPI_CONCILIADO]="N")` |
| Files abiertos por fase y semáforo | matriz `EST_FASE` × `EST_SEMAFORO` |
| Productividad por digitador | SAE, partidas y model codes por `CTL_DIGITADOR_ID` y semana ISO |
| Files facturables del período 25‑25 | `CALCULATE(DISTINCTCOUNT(tbl_Remision[ID_FILE_ADUANA]), dim_Calendario[PERIODO_FAC] = …)` |

---

## 5. Hoja de ruta de implementación

| Fase | Entregable | Duración estimada |
|---|---|---|
| 1. Catálogos | Cargar y validar `cat_*` con el equipo (sobre todo terminal/secuencia y las nuevas causales) | 1 semana |
| 2. Migración histórica | Power Query: `DATA` → `tbl_Reembarque` / `tbl_Remision` / `tbl_SAE` / `tbl_Eventos`, con hoja `ERRORES_CALIDAD` y la marca `CTL_FECHA_ESTIMADA` | 1‑2 semanas |
| 3. Operación en paralelo | El equipo registra eventos en la plantilla v2 mientras el consolidado viejo sigue vivo | 2‑4 semanas |
| 4. Tablero | Modelo Power BI conectado a las tablas (o a SharePoint/OneDrive) | 1 semana |
| 5. Automatización | Formulario de eventos (Microsoft Forms/Power Apps) o lectura de correos con asunto estandarizado → `tbl_Eventos` | evolutivo |

**Decisiones que el equipo debe confirmar:**

1. Qué significa el número en `SPRCTG 1` / `CONTECAR 2` / `SPRC14`.
2. El ANS de presentación de la DEX y la tolerancia de peso (según la norma vigente y el contrato).
3. Si los ANS se miden en horas calendario o en horas hábiles.
4. Cómo reportar el histórico con fechas estimadas: excluirlo o mostrarlo por separado.

# Matriz de Control Logístico v2 — Reembarques SAE / DEX

Rediseño del consolidado operativo (`CONSOLIDADO_GENERAL`) como un modelo de datos normalizado, listo para Excel (Tablas + Power Query) y Power BI.

| Archivo | Contenido |
|---|---|
| [`docs/rediseno-matriz-control-logistico.md`](docs/rediseno-matriz-control-logistico.md) | Diagnóstico del consolidado actual, arquitectura (modelo estrella + bitácora de eventos), normalización, máquina de estados, semáforo y ejemplo práctico |
| [`plantilla/MATRIZ_CONTROL_LOGISTICO_v2.xlsx`](plantilla/MATRIZ_CONTROL_LOGISTICO_v2.xlsx) | Plantilla operativa: catálogos, EVENTOS, REEMBARQUE, REMISION, SAE y RESUMEN, con KPI, conciliación y semáforo calculados |
| [`plantilla/generar_plantilla.py`](plantilla/generar_plantilla.py) | Script que regenera la plantilla (`python plantilla/generar_plantilla.py`) |

Los datos de ejemplo de la plantilla son ficticios.

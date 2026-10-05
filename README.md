# Control Logístico — app de escritorio para reembarques SAE / DEX

Aplicación local (Python + HTML) que reemplaza el `CONSOLIDADO_GENERAL.xlsx`. Se abre en su propia ventana, como un programa, y guarda los datos en una base SQLite dentro de la carpeta `app/datos/`.

## Cómo abrirla

**Windows:** instale Python 3.10 o superior (python.org, marcando *Add to PATH*) y haga doble clic en **`app/Iniciar_App.bat`**.
**macOS / Linux:** ejecute `./app/iniciar.sh`.

La primera vez se instala `openpyxl` (para leer y escribir Excel). La app se abre con Edge o Chrome en *modo aplicación* (ventana propia, sin barra de direcciones). Si instala `pip install pywebview`, se abre en una ventana nativa. Al cerrar la ventana, el servidor se detiene solo.

Opciones: `python app/logistica.py --navegador` (abre en una pestaña) · `--puerto 8800` · `--sin-ventana` (solo el servidor).

## Qué hace

| Pantalla | Para qué |
|---|---|
| **Tablero** | Files abiertos, files en rojo, % ANS de elaboración y de SAE (solo con tiempos reales), mediana de horas a SAE, SAE no conciliadas, incumplimientos sin causal, files por fase y semáforo, tendencia mensual, Pareto de causales, países y productividad por digitador. Filtro por período. |
| **Matriz operativa** | Un file por fila, ordenado por urgencia: semáforo, fase, barra de avance, a quién le toca, próxima acción y horas en la fase. Búsqueda y filtros. Al hacer clic se abre el detalle con la línea de pasos, el registro rápido de hitos, las SAE con conciliación y la bitácora completa. |
| **Registrar** | Formularios validados para hitos, reembarques, remisiones, SAE y personas (SAE y DEX de 13 dígitos, contenedor ISO 6346, neto ≤ bruto y ninguna fecha futura). |
| **Importar / Exportar** | Importa la hoja `DATA` del consolidado actual: unifica nombres, convierte fechas, separa los estados escritos en columnas de fecha y marca las fechas calculadas por fórmula. Exporta a Excel (Matriz, SAE, Eventos) para Power BI. |
| **Configuración** | Umbrales de ANS y tolerancias, fases y causales. |

**Regla central:** las fechas solo se capturan como *hitos* con su fecha y hora real. La fase, el semáforo y los KPI se calculan solos. El diseño completo está en [`docs/rediseno-matriz-control-logistico.md`](docs/rediseno-matriz-control-logistico.md).

## Crear un .exe (opcional)

```
pip install pyinstaller openpyxl
cd app
pyinstaller --onefile --noconsole --name ControlLogistico --add-data "web;web" logistica.py
```

El ejecutable queda en `app/dist/`. Los datos se guardan en una carpeta `datos/` junto al .exe.

## Estructura

```
app/
  logistica.py      servidor local + ventana
  motor.py          base de datos, reglas (fase, semáforo, KPI), importador y exportador
  web/index.html    interfaz
  Iniciar_App.bat   arranque con doble clic (Windows)
  iniciar.sh        arranque macOS / Linux
docs/               diseño del modelo de datos
plantilla/          versión anterior en Excel (referencia)
```

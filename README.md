# descarga_resultados_100X

Extractor concurrente de resultados agregados Saber 11 del ICFES. Procesa un
listado de códigos DANE con hasta 100 trabajadores y genera un Excel organizado
para análisis y Power BI.

## Contenido

- `descarga_resultados_100X.py`: programa principal.
- `ejecutar_descarga_100X.bat`: ejecución recomendada en Windows.
- `EstablecimientosCalendarioB_Grados_11_26.csv`: listado incluido de calendario B.
- `requirements.txt`: dependencias de Python.

Los Excel generados, checkpoints y logs no se almacenan en GitHub.

## Requisitos

- Windows 10/11 o Windows Server.
- Python 3.11 o superior disponible mediante el comando `python`.
- Conexión a Internet.

Instale las dependencias una sola vez:

```powershell
python -m pip install -r requirements.txt
```

## Ejecución

Haga doble clic en `ejecutar_descarga_100X.bat`, o ejecute:

```powershell
python descarga_resultados_100X.py `
  --entrada "EstablecimientosCalendarioB_Grados_11_26.csv" `
  --salida "resultados_icfes_2026_calendario_B_100_hilos.xlsx" `
  --hilos 100
```

La configuración predeterminada consulta la aplicación `2026-1`, grado 11. El
programa acepta entradas `.csv`, `.xls` y `.xlsx` con una columna DANE o Código.

## Archivos de salida

- `resultados_icfes_2026_calendario_B_100_hilos.xlsx`
- `resultados_icfes_2026_calendario_B_100_hilos.progreso.jsonl`
- `resultados_icfes_2026_calendario_B_100_hilos.progreso.jsonl.meta.json`
- `resultados_icfes_2026_calendario_B_100_hilos.log`

El checkpoint permite continuar después de una interrupción ejecutando nuevamente
el mismo comando. No use `--reiniciar` para continuar, porque esa opción comienza
el lote desde cero.

## Protección y consistencia

- Un único escritor guarda el progreso y evita conflictos entre trabajadores.
- Las respuestas 403, 429, CAPTCHA y fallos repetidos activan el cortacircuitos.
- Los colegios pendientes por red no se registran como rechazados.
- El Excel se construye en un archivo temporal, se reabre, se verifican sus filas
  y solo entonces se publica con el nombre definitivo.
- Si Excel falla, el checkpoint permanece disponible para reconstruir la salida
  con `--solo-exportar`.

## Modelo para Power BI

El libro contiene:

- `Dim_Instituciones`
- `Hecho_Poblacion`
- `Hecho_Puntajes`
- `Hecho_Niveles`
- `Hecho_Aprendizajes`
- `Control_Proceso`
- `Rechazados`
- `Resumen_Ejecucion`
- `Modelo_PowerBI`

Las tablas de hechos se relacionan con `Dim_Instituciones` mediante el campo
`dane`.

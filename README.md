# Caso Especialista GI — Validación, conciliación y análisis de colocación

Resolución del caso técnico para la posición de **Especialista de Gestión de Información**: limpieza y validación de datos, conciliación contra cifras de control, análisis de cumplimiento vs. metas, y revisión crítica de un borrador de conclusiones generado sin validar los datos.

## Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `Caso_de_uso.py` | Script único que resuelve los 6 pasos del caso |
| `Caso_Especialista_GI_CANDIDATO.xlsx` | Archivo original (hojas `Datos`, `Control`, `Metas`, `Conclusión preliminar`) |
| `Caso_Especialista_GI_CANDIDATO_LIMPIO.xlsx` | Salida del script: datos limpios + conciliación + análisis + crítica, todo en hojas separadas |
| `Caso_Especialista_GI_Presentacion.pptx` | Presentación con los resultados del caso |
| `reglas_aprendidas.json` | Reglas de normalización de categorías, aprendidas y reutilizadas entre corridas |

## Cómo correrlo

```bash
pip install pandas openpyxl
python Caso_de_uso.py
```

El `.py` debe estar en la misma carpeta que `Caso_Especialista_GI_CANDIDATO.xlsx`. Genera `Caso_Especialista_GI_CANDIDATO_LIMPIO.xlsx`, `Datos_Limpios.csv` y `reglas_aprendidas.json`.

**Capa generativa (opcional):** los pasos 2 y 3 pueden redactarse automáticamente con un modelo open source corriendo local vía [Ollama](https://ollama.com) (sin API keys, sin enviar datos a internet):

```bash
ollama pull llama3.1
python Caso_de_uso.py
```

Sin Ollama corriendo, el script igual calcula todas las cifras y tablas — solo deja el texto narrativo marcado como pendiente en vez de inventarlo.

## Qué resuelve cada paso

1. **Validar y limpiar** — perfilado automático de la hoja `Datos`, detección de formatos de fecha mezclados, categorías inconsistentes, montos con formato mixto, duplicados, valores faltantes y atípicos. Todo en Python/pandas, determinista.
2. **Conciliar** — puente de conciliación contra la hoja `Control`, cifra por cifra, con explicación de cada diferencia.
3. **Analizar** — % de cumplimiento vs. `Metas` por ejecutivo y región, y distribución de colocación por segmento, calculado con y sin los registros atípicos.
4. **Criterio vs. IA** — crítica punto por punto del borrador de la hoja `Conclusión preliminar` (generado sin validar datos) y su reescritura con cifras validadas.
5. **Mejora** — qué parte del proceso se automatizaría de forma recurrente y cómo.
6. **Bonus SQL** — consulta de Top 5 ejecutivos por colocación en la región Centro.

## Sobre el uso de IA

La redacción de texto narrativo (explicaciones, hallazgos, crítica y recomendación) se generó con un modelo de lenguaje open source corriendo local. Todo lo que es cálculo —totales, porcentajes, deduplicación, detección de outliers— se resuelve 100% en Python. Cada texto generado pasa por un verificador automático que compara toda cifra citada contra las tablas ya calculadas y marca lo que no cuadra, en vez de publicarlo a ciegas. El detalle completo está en la presentación (slide "Declaración de uso de IA").

## Autora

Claudia Anaya Morales

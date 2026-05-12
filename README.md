# Análisis Económico Minero JORC / NI 43-101

Aplicación web interactiva desarrollada con **Streamlit** para la evaluación de viabilidad económica de depósitos polimetálicos tipo skarn (Cu-Zn-Ag-Mo).

## Descripción

Esta herramienta permite realizar un análisis completo conforme a estándares **JORC** y **NI 43-101**, incluyendo:

- Cálculo determinístico de leyes de corte (COG), equivalentes de cobre (CuEq), márgenes operativos y factores de conversión recurso-reserva.
- Simulación de Monte Carlo para evaluación de incertidumbre y riesgo.
- Generación automática de 14 gráficos profesionales (histogramas, CDF, tornado, sensibilidad, scatter, comparativos recurso/reserva).
- Exportación de un reporte completo en formato Excel (.xlsx) con 5 hojas de trabajo:
  1. Datos de Entrada
  2. Interpretación de Resultados
  3. Cálculo paso a paso JORC/NI 43-101
  4. Estadísticas Monte Carlo
  5. Muestra Monte Carlo (500 iteraciones)

## Estructura del Proyecto

```
.
├── app.py              # Aplicación principal (Streamlit)
├── requirements.txt    # Dependencias de Python
└── README.md           # Este archivo
```

## Requisitos

- Python >= 3.9
- pip o conda para gestión de paquetes

## Instalación Local

1. Clona o descarga este repositorio en tu máquina:

```bash
git clone <URL_DEL_REPO>
cd <NOMBRE_DEL_REPO>
```

2. Crea un entorno virtual (recomendado):

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

3. Instala las dependencias:

```bash
pip install -r requirements.txt
```

## Ejecución Local

Una vez instaladas las dependencias, ejecuta:

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en tu navegador predeterminado en `http://localhost:8501`.

## Despliegue en Streamlit Cloud

1. Sube estos tres archivos (`app.py`, `requirements.txt`, `README.md`) a un repositorio público de **GitHub**.
2. Ve a [Streamlit Cloud](https://streamlit.io/cloud) e inicia sesión con tu cuenta de GitHub.
3. Haz clic en **"New app"** y selecciona tu repositorio.
4. Asegúrate de que el archivo principal sea `app.py`.
5. Haz clic en **"Deploy"**. La app estará lista en minutos.

## Uso de la Aplicación

1. En la **barra lateral izquierda** encontrarás todos los parámetros organizados por categorías:
   - Block Model
   - Grades
   - Costs
   - Metal Prices
   - Metallurgical Recovery
   - Dilution & Mining
   - Resources
   - Reserve Tonnages / Grades
   - Royalty
   - Monte Carlo

2. Puedes modificar cualquier valor numérico directamente. Todos los campos cargan con valores por defecto representativos de un depósito tipo Antamina.

3. Presiona el botón **"Ejecutar Análisis"** para calcular los resultados.

4. Navega por las pestañas:
   - **Resumen Ejecutivo:** Interpretaciones de riesgo con semáforos de color.
   - **Resultados Determinísticos:** Métricas clave en formato de tarjetas.
   - **Monte Carlo:** Estadísticas de salida, percentiles P5/P50/P90/P95 y correlaciones de Pearson.
   - **Gráficos:** 14 figuras profesionales generadas con matplotlib.
   - **Exportar Excel:** Descarga el reporte completo con un solo clic.

## Tecnologías Utilizadas

- [Streamlit](https://streamlit.io/) - Framework de aplicaciones web para Python
- [NumPy](https://numpy.org/) - Computación numérica
- [Matplotlib](https://matplotlib.org/) - Visualización de datos
- [OpenPyXL](https://openpyxl.readthedocs.io/) - Generación de archivos Excel
- [SciPy](https://scipy.org/) - Estadísticas y análisis científico

## Autor y Licencia

Desarrollado para fines de evaluación económica minera conforme a estándares internacionales JORC y NI 43-101.

---

**Nota:** Esta aplicación es una herramienta de soporte a decisiones. Los resultados deben ser validados por un competente persona (QP) antes de su uso en reportes técnicos oficiales.

# BCV Exchange Rates to Excel Exporter

Un script en **Python** diseñado para automatizar la extracción (*web scraping*) de las tasas de cambio oficiales publicadas por el Banco Central de Venezuela (BCV) y exportarlas de manera limpia y formateada a un archivo **Excel (.xlsx)**.

---

## Características
* **Extracción Multi-moneda:** Obtiene las cotizaciones de Dólar (USD), Euro (EUR), Yuan (CNY), Lira Turca (TRY) y Rublo (RUB).
* **Formato Profesional:** Genera un reporte en Excel listo para presentación con encabezados corporativos, ancho de columna ajustado, formato numérico explícito (`#,##0.00`) y marcas de fecha/hora.
* **Resiliencia HTTP:** Maneja tiempos de espera y bloqueos SSL de forma segura con `httpx`.

---

## Tecnologías Utilizadas
* **Python 3.10+**
* **[HTTPX](https://www.python-httpx.org/):** Para solicitudes HTTP asíncronas/síncronas eficientes.
* **[BeautifulSoup4](https://www.crummy.com/software/BeautifulSoup/):** Para el parseo del DOM HTML del sitio web del BCV.
* **[OpenPyXL](https://openpyxl.readthedocs.io/):** Para la creación y estilizado directo de hojas de cálculo Excel.

---

## Instalación y Uso

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/DaveroDev/bcv-excel-exporter.git](https://github.com/DaveroDev/bcv-excel-exporter.git)
   cd bcv-excel-exporter

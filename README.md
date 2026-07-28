# 🏭 Smart Factory ETL & OEE Automated Pipeline

Pipeline de procesamiento de datos (ETL) y automatización industrial desarrollado en Python para el cálculo de métricas de eficiencia global de equipos (**OEE** - Disponibilidad, Rendimiento y Calidad) y seguimiento de líneas de producción (Inyectoras y Ensamblaje).

---

## 🚀 ¿Qué hace este proyecto?
1. **Extracción (Extract):** Conecta de forma automatizada con múltiples fuentes de datos estructuradas en Google Sheets (registros operativos de planta, sensores LOT/FAL y tablas maestras de estándares).
2. **Transformación (Transform):** 
   - Limpieza y normalización de DataFrames con `pandas` y `numpy`.
   - Estandarización de nomenclaturas de máquinas y códigos de piezas.
   - Cálculo dinámico de tiempos netos, de proceso, paradas (con manejo de variables dummies) y restricciones operativas.
   - Cruce de datos maestros (`GPH`, `MUL`, `DIV`) mediante claves únicas compuestas (`CONC` e `ID_Inyeccion`).
3. **Carga (Load):** Sincroniza y actualiza automáticamente los informes procesados de vuelta en Google Sheets estructurados por planta.
4. **Automatización:** Se ejecuta de forma desatendida mediante **GitHub Actions** en días hábiles.

---

## 🛠️ Stack Tecnológico
* **Lenguaje:** Python 3.10
* **Manipulación y Análisis de Datos:** Pandas, NumPy
* **APIs & Cloud Integration:** Gspread, Google Cloud Service Accounts (IAM)
* **Automatización CI/CD:** GitHub Actions (cron triggers y ejecución programada)

---

## ⚙️ Arquitectura de Automatización (CI/CD)
El proyecto utiliza **GitHub Actions** para correr de manera autónoma sin intervención humana. Las credenciales de acceso a la API de Google se gestionan de forma segura mediante **GitHub Secrets**, inyectando el archivo JSON de la cuenta de servicio en tiempo de ejecución.
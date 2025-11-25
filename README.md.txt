📦 Automatización de Facturas Vencidas — Transportes Betancourt

Automatización completa del proceso de cobranza de facturas vencidas para las razones sociales:

Betancourt Hermanos

Transportes Claudio EIRL

Transportes Eduardo EIRL

Este sistema está en producción y ejecuta diariamente la detección de facturas vencidas, envío de correos, generación de reportes y trazabilidad completa de cobranza.

🚀 Funcionalidades Principales

✔ Lectura automática de planillas Excel
✔ Filtrado de facturas vencidas (>30 días)
✔ Consolidación de las tres razones sociales
✔ Envío interno de resumen
✔ Cruce con archivo CLIENTES.xlsx
✔ Envío automático de correos por cliente
✔ Generación de PDF con KPIs para jefatura
✔ Registro histórico en Excel
✔ Manejo de RUTs sin correo
✔ Logging detallado del proceso

📁 Estructura del Proyecto

automatizacion_facturas/
│
├── automatizacion_facturas.py     # Script principal
├── README.md                      # Documentación del proyecto
├── .gitignore                     # Archivos que no deben subirse
├── logs/                          # Logs generados
└── output/                        # Archivos generados (PDF, Excel)

▶️ Ejecución

Instalar dependencias:

pip install pandas reportlab openpyxl


Ejecutar el script:

python automatizacion_facturas.py

🔐 Seguridad

Este repositorio no incluye credenciales reales.

Para la ejecución en producción se usa un archivo .env:

SMTP_USER=facturacion@transportesbetancourt.cl
SMTP_PASS=CLAVE_DE_APLICACION


Y en el script:

import os
from dotenv import load_dotenv
load_dotenv()

SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

🧠 Próximas mejoras (para escalar a IAagent)

Migrar historial a SQLite

Automatizar flujos con n8n

Agente de IA para cobranza con CrewAI

Lectura automática de respuestas de clientes (IMAP + NLP)

Dashboard en Power BI para análisis de cobranza

👤 Autor

Jorge Vidal Larrondo
Ingeniero Comercial – Diplomado en Data Science 
Automatizaciones – Python – IA aplicada








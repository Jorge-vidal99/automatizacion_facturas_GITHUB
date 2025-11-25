# -*- coding: utf-8 -*-
import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.message import EmailMessage
import logging
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, HRFlowable
from reportlab.lib import colors

#%%
# === CONFIGURACIÓN DE LOGGING Y RUTA BASE ===
BASE_PATH = r"C:\Users\Pablo\Desktop\automatizacion_facturas"

log_folder = os.path.join(BASE_PATH, 'logs')
os.makedirs(log_folder, exist_ok=True)

log_filename = os.path.join(log_folder, f'log_{datetime.now().strftime("%Y%m%d")}.log')

logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logging.info("====== INICIO DEL SCRIPT ======")

# === RUTAS DE PLANILLAS ===
PLANILLAS = {
    "BETANCOURT HERMANOS": os.path.join(BASE_PATH, "FACTURAS_TBH_2025.xlsx"),
    "CLAUDIO EIRL": os.path.join(BASE_PATH, "FACTURAS_CLAUDIO_2025.xlsx"),
    "EDUARDO EIRL": os.path.join(BASE_PATH, "FACTURAS_EDUARDO_2025.xlsx")
}
#%%
# === PROCESAMIENTO Y FILTRO ===
facturas_consolidadas = []
fecha_limite = datetime.now() - timedelta(days=30)

for razon_social, ruta in PLANILLAS.items():
    if not os.path.exists(ruta):
        logging.error(f"No se encontró el archivo para {razon_social}: {ruta}")
        continue

    try:
        df = pd.read_excel(ruta)
        df.columns = df.columns.str.strip().str.upper()

        # --- BLOQUE DE MONTO ---
        col_monto = next((col for col in df.columns if "MONTO" in col.upper()), None)
        if col_monto:
            df.rename(columns={col_monto: "MONTO"}, inplace=True)
            if df["MONTO"].dtype == object:
                df["MONTO"] = (
                    df["MONTO"]
                    .astype(str)
                    .str.replace(r"[^\d,.-]", "", regex=True)
                    .str.replace(",", ".", regex=False)
                )
            df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce")
        else:
            logging.warning(f"No se encontró columna 'MONTO' en archivo: {ruta}")

        # Procesar fecha y estado
        df["FECHA EMISION"] = pd.to_datetime(df["FECHA EMISION"], errors="coerce")
        df = df.dropna(subset=["FECHA EMISION", "ESTADO"])

        # Filtro de facturas impagas y vencidas
        df_filtrado = df[
            (df["FECHA EMISION"] < fecha_limite)
            & (df["ESTADO"].astype(str).str.strip().str.upper() == "IMPAGA")
        ].copy()

        if not df_filtrado.empty:
            df_filtrado["RAZON SOCIAL"] = razon_social
            facturas_consolidadas.append(df_filtrado)
            logging.info(f"{razon_social}: {len(df_filtrado)} facturas vencidas encontradas.")
        else:
            logging.info(f"{razon_social}: Sin facturas vencidas.")

    except Exception as e:
        logging.error(f"Error al procesar {razon_social}: {e}")
        continue



#%%
# === GUARDAR ARCHIVO CONSOLIDADO ===
if facturas_consolidadas:
    df_final = pd.concat(facturas_consolidadas, ignore_index=True)

    # Normalizar nombres de columnas
    df_final.columns = df_final.columns.str.strip().str.upper()

    # Renombrar variantes comunes
    df_final.rename(columns={
        "Nº FACTURA": "N_FACTURA",
        "N° FACTURA": "N_FACTURA",
        "NUMERO FACTURA": "N_FACTURA",
        "FACTURA": "N_FACTURA",
        "MONTO ": "MONTO",
        "TOTAL": "MONTO"
    }, inplace=True)

    # Verificar columnas disponibles
    logging.info(f"Columnas finales en consolidado: {df_final.columns.tolist()}")

    # Definir columnas deseadas
    columnas_deseadas = ["N_FACTURA", "FECHA EMISION", "CLIENTE", "RUT", "MONTO", "ESTADO", "RAZON SOCIAL"]

    # Filtrar solo columnas que existen
    columnas_existentes = [c for c in columnas_deseadas if c in df_final.columns]
    df_final = df_final[columnas_existentes]

    # Guardar archivo
    archivo_salida = os.path.join(BASE_PATH, "FACTURAS_VENCIDAS_CONSOLIDADO.xlsx")
    df_final.to_excel(archivo_salida, index=False)
    logging.info(f"✅ Archivo consolidado guardado: {archivo_salida}")
    print(f"✅ Archivo generado: {archivo_salida}")

else:
    logging.info("❌ No se generó archivo: no se encontraron facturas vencidas.")
    print("No hay facturas vencidas para consolidar.")

    # Evita errores en sección de envío
    df_final = None

#%%
# === ENVÍO DE CORREO CON RESUMEN POR RAZÓN SOCIAL ===
def enviar_correo(df, archivo_adjunto):
    try:
        # Crear resumen por razón social
        resumen = df.groupby("RAZON SOCIAL").size().reset_index(name="CANTIDAD")
        resumen_texto = "\n".join(f"- {row['RAZON SOCIAL']}: {row['CANTIDAD']} factura(s) vencida(s)"
                                  for _, row in resumen.iterrows())

        # Crear el mensaje
        msg = EmailMessage()
        msg['Subject'] = '🔔 Facturas Vencidas (>30 días) 🔔 - Transportes Betancourt'
        msg['From'] = 'facturacion@transportesbetancourt.cl'
        msg['To'] = 'facturacion@transportesbetancourt.cl'
        msg['Cc'] = '@transportesbetancourt.cl', ''

        msg.set_content(f"""
Estimados,

Se adjunta el consolidado de facturas vencidas por más de 30 días correspondientes a las 3 razones sociales de Transportes Betancourt.

Resumen de facturas vencidas por razón social:
{resumen_texto}

Total general: {len(df)} facturas vencidas

"QUEDO ATENTO A CUALQUIER RECOMENDACIÓN O SUGERENCIA"

Saludos cordiales,
Sistema de Automatización de recordatorio de facturas
Jorge Vidal Larrondo
        """)

        # Adjuntar el archivo Excel
        with open(archivo_adjunto, 'rb') as f:
            contenido = f.read()
            msg.add_attachment(contenido,
                               maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=os.path.basename(archivo_adjunto))

        # Enviar el correo
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ejemplo@transportesbetancourt.cl', 'clave ejemplo')
            smtp.send_message(msg)

        print("✅ Correo enviado correctamente.")
        logging.info("Correo enviado correctamente.")

    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")
        logging.error(f"Error al enviar correo: {e}")

#%%
# === ENVIAR SI HAY ARCHIVO ===
if df_final is not None:
    enviar_correo(df_final, archivo_salida)
    
#%%

# === CONFIGURACIÓN ===

logging.info("====== INICIO DEL SCRIPT DE COBRANZA ======")

# === CARGA DE DATOS ===
ARCHIVO_FACTURAS = os.path.join(BASE_PATH, "FACTURAS_VENCIDAS_CONSOLIDADO.xlsx")
ARCHIVO_CLIENTES = r"C:\Users\Pablo\Desktop\automatizacion_facturas\CLIENTES.xlsx"
SALIDA_SIN_CORREO = os.path.join(BASE_PATH, "RUTS_SIN_CORREO.xlsx")

try:
    df_facturas = pd.read_excel(ARCHIVO_FACTURAS)
    df_clientes = pd.read_excel(ARCHIVO_CLIENTES)

    # Normalizar columnas
    df_clientes.columns = df_clientes.columns.str.strip().str.upper()
    df_facturas.columns = df_facturas.columns.str.strip().str.upper()

    # Renombrar columnas clave para evitar errores
    df_facturas.rename(columns={
        "Nº FACTURA": "N_FACTURA",
        "N° FACTURA": "N_FACTURA",
        "NUMERO FACTURA": "N_FACTURA",
        "MONTO ": "MONTO",
        "MONTO": "MONTO",
        "FECHA EMISIÓN": "FECHA EMISION",
    }, inplace=True)

    df_clientes['RUT'] = df_clientes['RUT'].astype(str).str.strip()
    
    # Asegurar que todas las facturas tengan el campo CLIENTE
    df_facturas['RUT'] = df_facturas['RUT'].astype(str).str.strip()
    df_clientes['RUT'] = df_clientes['RUT'].astype(str).str.strip()

# Guardamos la razon social original y renombramos la del cliente como CLIENTE
    df_facturas.rename(columns={"RAZON SOCIAL": "RAZON_SOCIAL_FACTURA"}, inplace=True)

    df_facturas = df_facturas.merge(
    df_clientes[['RUT', 'RAZON SOCIAL']], 
    how='left', 
    on='RUT'
).rename(columns={"RAZON SOCIAL": "CLIENTE"})

# Restauramos la columna de razón social original con un nombre más simple
    df_facturas.rename(columns={"RAZON_SOCIAL_FACTURA": "RAZON SOCIAL"}, inplace=True)

# Confirmar que no hay columnas duplicadas
    df_facturas = df_facturas.loc[:, ~df_facturas.columns.duplicated()]

    
except Exception as e:
    logging.error(f"Error cargando archivos: {e}")
    raise

# === PREPARACIÓN ===
ruts_sin_correo = []
clientes_procesados = 0
facturas_enviadas = []

# Agrupar facturas por RUT cliente
for rut_cliente, df_rut in df_facturas.groupby("RUT"):
    rut_cliente = str(rut_cliente).strip()
    info_cliente = df_clientes[df_clientes['RUT'] == rut_cliente]

    if info_cliente.empty:
        logging.warning(f"RUT {rut_cliente} no encontrado en archivo de clientes.")
        continue

    razon_social = info_cliente.iloc[0]['RAZON SOCIAL']
    correo_to = info_cliente.iloc[0]['CORREOS']
    correo_cc = info_cliente.iloc[0]['CORREOS 2']

    if pd.isna(correo_to) and pd.isna(correo_cc):
        logging.warning(f"{razon_social} (RUT: {rut_cliente}) no tiene correos disponibles.")
        ruts_sin_correo.append({"RAZON SOCIAL": razon_social, "RUT": rut_cliente})
        continue

    # Construir mensaje
    facturas_texto = "\n".join(
        f"- N°{row['N_FACTURA']} - Monto: ${row['MONTO']:,} - Emitida: {row['FECHA EMISION'].date()}"
        for _, row in df_rut.iterrows()
    )

    msg = EmailMessage()
    msg['Subject'] = f"🔔 Aviso: Facturas vencidas de {razon_social}"
    msg['From'] = 'facturacion@transportesbetancourt.cl'
    msg['To'] = correo_to if pd.notna(correo_to) else ''
    msg['Cc'] = ', '.join(filter(pd.notna, [correo_cc, 'administracion@transportesbetancourt.cl']))

    msg.set_content(f"""
Estimados {razon_social},

Nos dirigimos a ustedes para informar que se han vencido las siguientes facturas por más de 30 días:

{facturas_texto}

Les solicitamos revisar y gestionar su pago. Ante cualquier duda, quedamos atentos.

Saludos cordiales,
Transportes Betancourt
Jorge Ignacio Vidal Larrondo
""")

    # Envío
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ejemplo@transportesbetancourt.cl', 'clave ejemplo')
            smtp.send_message(msg)
        logging.info(f"Correo enviado a {razon_social} ({rut_cliente})")
        clientes_procesados += 1
        
        facturas_enviadas.extend(df_rut.to_dict(orient='records'))
        
    except Exception as e:
        logging.error(f"Error enviando correo a {razon_social} ({rut_cliente}): {e}")

# === GUARDAR RUTS SIN CORREO ===
if ruts_sin_correo:
    pd.DataFrame(ruts_sin_correo).to_excel(SALIDA_SIN_CORREO, index=False)
    logging.info(f"Se guardó archivo de RUTs sin correo: {SALIDA_SIN_CORREO}")

print(f"✅ Proceso finalizado. Correos enviados a {clientes_procesados} clientes.")

#%%
# === ENVÍO DE CORREO RESUMEN A JEFATURA CON PDF Y LOGO ===


def enviar_resumen_a_jefatura(facturas_enviadas, ruts_sin_correo, df_facturas):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image

        fecha_envio = datetime.now().strftime("%d de %B de %Y")
        nombre_pdf = f"resumen_jefatura_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        archivo_pdf = os.path.join(BASE_PATH, nombre_pdf)
        logo_path = r"C:\Users\Pablo\Desktop\automatizacion_facturas\LOGO_BETANCOURT.png"

        # === FACTURAS ENVIADAS ===
        resumen_por_razon = facturas_enviadas.groupby("RAZON SOCIAL").size().reset_index(name="FACTURAS ENVIADAS")
        detalle_envios = "\n".join(
            f"• {row['RAZON SOCIAL']}: {row['FACTURAS ENVIADAS']} factura(s)"
            for _, row in resumen_por_razon.iterrows()
        )

        # === FACTURAS NO COBRADAS ===
        ruts_cobrados = set(facturas_enviadas["RUT"].astype(str))
        ruts_totales = set(df_facturas["RUT"].astype(str))
        ruts_no_cobrados = ruts_totales - ruts_cobrados

        facturas_no_enviadas = df_facturas[df_facturas["RUT"].astype(str).isin(ruts_no_cobrados)]
        total_no_cobradas = len(facturas_no_enviadas)
        if not facturas_no_enviadas.empty:
           detalle_no_enviadas = ""
           for razon, grupo in facturas_no_enviadas.groupby("RAZON SOCIAL"):
               detalle_no_enviadas += f"\n➡ {razon}:\n"
               for _, row in grupo.iterrows():
                   detalle_no_enviadas += (
                f" - N°{row['N_FACTURA']} | ${row['MONTO']:,} | {row['FECHA EMISION'].date()}\n"
            )
        else:
           detalle_no_enviadas = "– Ninguna –"


        # === CREAR PDF ===
        doc = SimpleDocTemplate(archivo_pdf, pagesize=A4)
        styles = getSampleStyleSheet()
        flowables = []

        if os.path.exists(logo_path):
            img = Image(logo_path, width=6*cm, height=3*cm)
            img.hAlign = 'CENTER'
            flowables.append(img)
            flowables.append(Spacer(1, 12))

       # === BLOQUE PRINCIPAL CON KPI ===
        total_monto_cobradas = facturas_enviadas["MONTO"].sum()
        total_monto_no_cobradas = facturas_no_enviadas["MONTO"].sum()

        flowables.extend([
    Paragraph("<b>■ Resumen del Proceso de Cobranza Automatizada</b>", styles["Title"]),
    Spacer(1, 12),
    Paragraph(f"📅 <b> Fecha del resumen:</b> {fecha_envio}", styles["Normal"]),
    Spacer(1, 12),
    Paragraph("<hr width='100%'>", styles["Normal"]),

    Paragraph(f"✔️ <b>Total de facturas cobradas:</b> {len(facturas_enviadas)} — ${total_monto_cobradas:,.0f}", styles["Normal"]),
    Spacer(1, 6),
    Paragraph(f"❌ <b>Total de facturas NO cobradas:</b> {total_no_cobradas} — ${total_monto_no_cobradas:,.0f}", styles["Normal"]),
    Spacer(1, 12),
    Paragraph("<hr width='100%'>", styles["Normal"]),

    Paragraph(f"📁 <b>Detalle por razón social:</b>", styles["Normal"]),
    Paragraph(detalle_envios if detalle_envios else "- Sin registros -", styles["Normal"]),
    Spacer(1, 12),
    Paragraph("<hr width='100%'>", styles["Normal"]),
    Paragraph(f"🔍 <b>Detalle de facturas cobradas:</b>", styles["Normal"]),
    Spacer(1, 6),
])

# === DETALLE COBRADAS (en lista) ===
        flowables.append(Spacer(1, 12))
        flowables.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.grey))
        flowables.append(Spacer(1, 6))
        flowables.append(Paragraph("📋 <b>Detalle de facturas cobradas:</b>", styles["Normal"]))
        flowables.append(Spacer(1, 6))

        for razon, grupo in facturas_enviadas.groupby("RAZON SOCIAL"):
            flowables.append(Paragraph(f"🗂️ <b>{razon}:</b>", styles["Normal"]))
            for _, row in grupo.iterrows():
                detalle = (
                    f"- N°{row['N_FACTURA']} | Cliente: {row['CLIENTE']} | "
                    f"${row['MONTO']:,} | {row['FECHA EMISION'].date()}"
            )
                flowables.append(Paragraph(detalle, styles["Normal"]))
                flowables.append(Spacer(1, 6))

# === DETALLE NO COBRADAS (moved to bottom) ===
        flowables.append(Spacer(1, 12))
        flowables.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.grey))
        flowables.append(Spacer(1, 6))
        flowables.append(Paragraph("📋 <b>Detalle de facturas no cobradas:</b>", styles["Normal"]))
        flowables.append(Spacer(1, 6))

        for razon, grupo in facturas_no_enviadas.groupby("RAZON SOCIAL"):
            flowables.append(Paragraph(f"🗂️ <b>{razon}:</b>", styles["Normal"]))
            for _, row in grupo.iterrows():
                detalle = (
                    f"- N°{row['N_FACTURA']} | Cliente: {row['CLIENTE']} | "
                    f"${row['MONTO']:,} | {row['FECHA EMISION'].date()}"
                )   
                flowables.append(Paragraph(detalle, styles["Normal"]))
            flowables.append(Spacer(1, 6))


# === CIERRE ===
        flowables.append(Paragraph("<hr width='100%'/>", styles["Normal"]))
        flowables.extend([
    Spacer(1, 12),
    Paragraph("Saludos cordiales,", styles["Normal"]),
    Paragraph("Sistema de Cobranza Automatica de Facturas", styles["Normal"]),
    Paragraph("Jorge Vidal Larrondo", styles["Normal"]),
])

        doc.build(flowables)

        # ENVIAR
        msg = EmailMessage()
        msg['Subject'] = "📬 Resumen Diario De Cobranza Automatizada de Facturas "
        msg['From'] = 'facturacion@transportesbetancourt.cl'
        msg['To'] = 'administracion@transportesbetancourt.cl'
        msg['Cc'] = 'jvidal@transportesbetancourt.cl' 
        msg.set_content("Estimados,\n\nComparto con ustedes el resumen diario del proceso de cobranza automatizada de facturas.\n\nSaludos cordiales,\nJorge Ignacio Vidal Larrondo")

        with open(archivo_pdf, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=nombre_pdf)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ejemplo@transportesbetancourt.cl', 'clave ejemplo')
            smtp.send_message(msg)

        print("📩 Correo resumen enviado a jefatura.")
        logging.info("Correo resumen enviado a jefatura.")

    except Exception as e:
        print(f"❌ Error al enviar correo resumen a jefatura: {e}")
        logging.error(f"Error al enviar correo resumen a jefatura: {e}")


#%%

# === ENVIAR CORREO RESUMEN A JEFATURA ===

if facturas_enviadas:
    df_enviadas = pd.DataFrame(facturas_enviadas)
    enviar_resumen_a_jefatura(df_enviadas, ruts_sin_correo, df_facturas)

    
    
 #%%   
 
# === REGISTRO EN HISTORIAL DE COBRANZAS CON INTENTOS ===

historial_path = r"C:\Users\Pablo\Desktop\automatizacion_facturas\HISTORIAL_COBRANZAS.xlsx"
registros = []

if facturas_enviadas:
    df_enviadas = pd.DataFrame(facturas_enviadas)
    df_enviadas.columns = df_enviadas.columns.str.upper()

    for _, row in df_enviadas.iterrows():
        registros.append({
            "N_FACTURA": row.get("N_FACTURA"),
            "FECHA_EMISION": row.get("FECHA EMISION"),
            "RUT_CLIENTE": row.get("RUT"),
            "RAZON_SOCIAL": row.get("RAZON SOCIAL"),
            "MONTO": row.get("MONTO"),
            "FECHA_COBRO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "CORREO_ENVIADO": row.get("CORREOS") if "CORREOS" in row else "",
            "OBSERVACION": "Cobranza automática enviada",
            "INTENTOS": 1  # valor inicial
        })

    df_registros = pd.DataFrame(registros)

    try:
        if os.path.exists(historial_path):
            df_historial = pd.read_excel(historial_path)
            df_historial.columns = df_historial.columns.str.upper()
            df_registros.columns = df_registros.columns.str.upper()

            # Concatenar
            df_combinado = pd.concat([df_historial, df_registros], ignore_index=True)

            # Agrupar por factura y cliente y contar intentos
            df_combinado["INTENTOS"] = df_combinado.groupby(["N_FACTURA", "RUT_CLIENTE"]).cumcount() + 1

            # Eliminar duplicados exactos dejando el último intento
            df_combinado.drop_duplicates(
                subset=["N_FACTURA", "RUT_CLIENTE", "FECHA_COBRO", "MONTO"],
                keep="last",
                inplace=True
            )
        else:
            df_combinado = df_registros

        df_combinado.to_excel(historial_path, index=False)
        logging.info(f"✅ Registro actualizado en {historial_path} con campo INTENTOS")

    except Exception as e:
        logging.error(f"❌ Error al registrar en historial de cobranzas: {e}")
        print(f"❌ Error al guardar historial de cobranzas: {e}")

#%%

df_facturas.to_excel(os.path.join(BASE_PATH, "FACTURAS_CON_TRAZABILIDAD.xlsx"), index=False)

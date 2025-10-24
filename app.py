import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText

# Config icono y título (igual, si tienes)
st.set_page_config(page_title="Presupuestos Automáticos", page_icon="icono.png")

# Datos de la hoja "Tablas" (completo como en tu mensaje)
data_tablas = [
    {'Espesor': 2, 'CR Select': 0, 'Velocidad': 5, 'VDI': 29, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '2-0'},
    {'Espesor': 2, 'CR Select': 5, 'Velocidad': 5, 'VDI': 29, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '2-5'},
    {'Espesor': 2, 'CR Select': 6, 'Velocidad': 8, 'VDI': 24, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '2-6'},
    {'Espesor': 2, 'CR Select': 7, 'Velocidad': 19, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '2-7'},
    {'Espesor': 2, 'CR Select': 8, 'Velocidad': 5, 'VDI': 29, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '2-8'},
    {'Espesor': 2, 'CR Select': 9, 'Velocidad': 8, 'VDI': 24, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '2-9'},
    {'Espesor': 5, 'CR Select': 0, 'Velocidad': 3.5, 'VDI': 27, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '5-0'},
    {'Espesor': 5, 'CR Select': 5, 'Velocidad': 3.5, 'VDI': 27, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '5-5'},
    {'Espesor': 5, 'CR Select': 6, 'Velocidad': 8, 'VDI': 24, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '5-6'},
    {'Espesor': 5, 'CR Select': 7, 'Velocidad': 19, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '5-7'},
    {'Espesor': 5, 'CR Select': 8, 'Velocidad': 3.5, 'VDI': 27, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '5-8'},
    {'Espesor': 5, 'CR Select': 9, 'Velocidad': 8, 'VDI': 24, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '5-9'},
    {'Espesor': 10, 'CR Select': 0, 'Velocidad': 3, 'VDI': 26, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '10-0'},
    {'Espesor': 10, 'CR Select': 5, 'Velocidad': 3, 'VDI': 26, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '10-5'},
    {'Espesor': 10, 'CR Select': 6, 'Velocidad': 5.7, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '10-6'},
    {'Espesor': 10, 'CR Select': 7, 'Velocidad': 18.7, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '10-7'},
    {'Espesor': 10, 'CR Select': 8, 'Velocidad': 3, 'VDI': 26, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '10-8'},
    {'Espesor': 10, 'CR Select': 9, 'Velocidad': 5.7, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '10-9'},
    {'Espesor': 20, 'CR Select': 0, 'Velocidad': 2.8, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '20-0'},
    {'Espesor': 20, 'CR Select': 5, 'Velocidad': 2.8, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '20-5'},
    {'Espesor': 20, 'CR Select': 6, 'Velocidad': 5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '20-6'},
    {'Espesor': 20, 'CR Select': 7, 'Velocidad': 16.8, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '20-7'},
    {'Espesor': 20, 'CR Select': 8, 'Velocidad': 2.8, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '20-8'},
    {'Espesor': 20, 'CR Select': 9, 'Velocidad': 5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '20-9'},
    {'Espesor': 30, 'CR Select': 0, 'Velocidad': 2.5, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '30-0'},
    {'Espesor': 30, 'CR Select': 5, 'Velocidad': 2.5, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '30-5'},
    {'Espesor': 30, 'CR Select': 6, 'Velocidad': 4.4, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '30-6'},
    {'Espesor': 30, 'CR Select': 7, 'Velocidad': 15, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '30-7'},
    {'Espesor': 30, 'CR Select': 8, 'Velocidad': 2.5, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '30-8'},
    {'Espesor': 30, 'CR Select': 9, 'Velocidad': 4.4, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '30-9'},
    {'Espesor': 40, 'CR Select': 0, 'Velocidad': 1.5, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '40-0'},
    {'Espesor': 40, 'CR Select': 5, 'Velocidad': 1.5, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '40-5'},
    {'Espesor': 40, 'CR Select': 6, 'Velocidad': 4.1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '40-6'},
    {'Espesor': 40, 'CR Select': 7, 'Velocidad': 13.1, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '40-7'},
    {'Espesor': 40, 'CR Select': 8, 'Velocidad': 1.5, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '40-8'},
    {'Espesor': 40, 'CR Select': 9, 'Velocidad': 4.1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '40-9'},
    {'Espesor': 50, 'CR Select': 0, 'Velocidad': 1.2, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '50-0'},
    {'Espesor': 50, 'CR Select': 5, 'Velocidad': 1.2, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '50-5'},
    {'Espesor': 50, 'CR Select': 6, 'Velocidad': 3.5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '50-6'},
    {'Espesor': 50, 'CR Select': 7, 'Velocidad': 12, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '50-7'},
    {'Espesor': 50, 'CR Select': 8, 'Velocidad': 1.2, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '50-8'},
    {'Espesor': 50, 'CR Select': 9, 'Velocidad': 3.5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '50-9'},
    {'Espesor': 60, 'CR Select': 0, 'Velocidad': 1.1, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '60-0'},
    {'Espesor': 60, 'CR Select': 5, 'Velocidad': 1.1, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '60-5'},
    {'Espesor': 60, 'CR Select': 6, 'Velocidad': 3.5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '60-6'},
    {'Espesor': 60, 'CR Select': 7, 'Velocidad': 10.9, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '60-7'},
    {'Espesor': 60, 'CR Select': 8, 'Velocidad': 1.1, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '60-8'},
    {'Espesor': 60, 'CR Select': 9, 'Velocidad': 3.4, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '60-9'},
    {'Espesor': 70, 'CR Select': 0, 'Velocidad': 0.9, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '70-0'},
    {'Espesor': 70, 'CR Select': 5, 'Velocidad': 0.9, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '70-5'},
    {'Espesor': 70, 'CR Select': 6, 'Velocidad': 3.2, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '70-6'},
    {'Espesor': 70, 'CR Select': 7, 'Velocidad': 10.2, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '70-7'},
    {'Espesor': 70, 'CR Select': 8, 'Velocidad': 0.9, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '70-8'},
    {'Espesor': 70, 'CR Select': 9, 'Velocidad': 3.2, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '70-9'},
    {'Espesor': 80, 'CR Select': 0, 'Velocidad': 0.7, 'VDI': 29, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '80-0'},
    {'Espesor': 80, 'CR Select': 5, 'Velocidad': 0.7, 'VDI': 29, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '80-5'},
    {'Espesor': 80, 'CR Select': 6, 'Velocidad': 2.8, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '80-6'},
    {'Espesor': 80, 'CR Select': 7, 'Velocidad': 9.5, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '80-7'},
    {'Espesor': 80, 'CR Select': 8, 'Velocidad': 0.7, 'VDI': 29, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '80-8'},
    {'Espesor': 80, 'CR Select': 9, 'Velocidad': 2.8, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '80-9'},
    {'Espesor': 90, 'CR Select': 0, 'Velocidad': 0.6, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '90-0'},
    {'Espesor': 90, 'CR Select': 5, 'Velocidad': 0.6, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '90-5'},
    {'Espesor': 90, 'CR Select': 6, 'Velocidad': 2.5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '90-6'},
    {'Espesor': 90, 'CR Select': 7, 'Velocidad': 8.8, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '90-7'},
    {'Espesor': 90, 'CR Select': 8, 'Velocidad': 0.6, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '90-8'},
    {'Espesor': 90, 'CR Select': 9, 'Velocidad': 2.5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '90-9'},
    {'Espesor': 100, 'CR Select': 0, 'Velocidad': 0.5, 'VDI': 29, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '100-0'},
    {'Espesor': 100, 'CR Select': 5, 'Velocidad': 0.5, 'VDI': 29, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '100-5'},
    {'Espesor': 100, 'CR Select': 6, 'Velocidad': 2.1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '100-6'},
    {'Espesor': 100, 'CR Select': 7, 'Velocidad': 8, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '100-7'},
    {'Espesor': 100, 'CR Select': 8, 'Velocidad': 0.5, 'VDI': 29, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '100-8'},
    {'Espesor': 100, 'CR Select': 9, 'Velocidad': 2.1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '100-9'},
    {'Espesor': 125, 'CR Select': 0, 'Velocidad': 0.4, 'VDI': 29, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '125-0'},
    {'Espesor': 125, 'CR Select': 5, 'Velocidad': 0.4, 'VDI': 29, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '125-5'},
    {'Espesor': 125, 'CR Select': 6, 'Velocidad': 1.6, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '125-6'},
    {'Espesor': 125, 'CR Select': 7, 'Velocidad': 7, 'VDI': 16, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '125-7'},
    {'Espesor': 125, 'CR Select': 8, 'Velocidad': 0.4, 'VDI': 29, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '125-8'},
    {'Espesor': 125, 'CR Select': 9, 'Velocidad': 1.6, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '125-9'},
    {'Espesor': 150, 'CR Select': 0, 'Velocidad': 0.3, 'VDI': 29, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '150-0'},
    {'Espesor': 150, 'CR Select': 5, 'Velocidad': 0.3, 'VDI': 29, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '150-5'},
    {'Espesor': 150, 'CR Select': 6, 'Velocidad': 1.5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '150-6'},
    {'Espesor': 150, 'CR Select': 7, 'Velocidad': 5.8, 'VDI': 17, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '150-7'},
    {'Espesor': 150, 'CR Select': 8, 'Velocidad': 0.3, 'VDI': 29, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '150-8'},
    {'Espesor': 150, 'CR Select': 9, 'Velocidad': 1.5, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '150-9'},
    {'Espesor': 175, 'CR Select': 0, 'Velocidad': 0.2, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '175-0'},
    {'Espesor': 175, 'CR Select': 5, 'Velocidad': 0.2, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '175-5'},
    {'Espesor': 175, 'CR Select': 6, 'Velocidad': 1.1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '175-6'},
    {'Espesor': 175, 'CR Select': 7, 'Velocidad': 5.1, 'VDI': 17, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '175-7'},
    {'Espesor': 175, 'CR Select': 8, 'Velocidad': 0.2, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '175-8'},
    {'Espesor': 175, 'CR Select': 9, 'Velocidad': 1.1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '175-9'},
    {'Espesor': 200, 'CR Select': 0, 'Velocidad': 0.15, 'VDI': 28, 'Tipo de paso': 'Rough (Baja)', 'Búsqueda': '200-0'},
    {'Espesor': 200, 'CR Select': 5, 'Velocidad': 0.15, 'VDI': 28, 'Tipo de paso': 'Rough (Alta)', 'Búsqueda': '200-5'},
    {'Espesor': 200, 'CR Select': 6, 'Velocidad': 1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Alta)', 'Búsqueda': '200-6'},
    {'Espesor': 200, 'CR Select': 7, 'Velocidad': 4.4, 'VDI': 17, 'Tipo de paso': 'Repaso2 (Alta)', 'Búsqueda': '200-7'},
    {'Espesor': 200, 'CR Select': 8, 'Velocidad': 0.15, 'VDI': 28, 'Tipo de paso': 'Rough (Media)', 'Búsqueda': '200-8'},
    {'Espesor': 200, 'CR Select': 9, 'Velocidad': 1, 'VDI': 25, 'Tipo de paso': 'Repaso1 (Media)', 'Búsqueda': '200-9'},
]
df_tablas = pd.DataFrame(data_tablas)

# Tabla estándar VDI a Ra (µm) - basada en VDI 3400
vdi_to_ra = {
    0: 0.1, 1: 0.112, 2: 0.126, 3: 0.14, 4: 0.16, 5: 0.18, 6: 0.2, 7: 0.22, 8: 0.25, 9: 0.28,
    10: 0.32, 11: 0.35, 12: 0.4, 13: 0.45, 14: 0.5, 15: 0.56, 16: 0.63, 17: 0.7, 18: 0.8, 19: 0.9,
    20: 1.0, 21: 1.12, 22: 1.26, 23: 1.4, 24: 1.6, 25: 1.8, 26: 2.0, 27: 2.2, 28: 2.5, 29: 2.8,
    30: 3.2, 31: 3.5, 32: 4.0, 33: 4.5, 34: 5.0, 35: 5.6, 36: 6.3, 37: 7.0, 38: 8.0, 39: 9.0,
    40: 10.0, 41: 11.2, 42: 12.6, 43: 14.0, 44: 16.0, 45: 18.0
}
st.title("Solicitud de Presupuesto - Servicorte por Hilo")
# Inputs
material = "Acero" # Fijo
espesor = st.number_input("Espesor (mm)", min_value=2, max_value=200, value=100, step=1)
calidad = st.selectbox("Calidad", ["Baja", "Media", "Alta"])
longitud = st.number_input("Longitud corte (mm)", min_value=0.0, value=25.0)
email = st.text_input("Tu email (para enviar presupuesto)")
# Valores fijos internos
tasa = 30.0
costo_fijo = 4.0
# Determinar códigos según calidad (pega tu lógica aquí)
if calidad == "Baja":
    codes = [0]
elif calidad == "Media":
    codes = [8, 9]
elif calidad == "Alta":
    codes = [5, 6, 7]
# Encontrar espesor disponible (más cercano inferior)
available_espesores = sorted(df_tablas['Espesor'].unique())
if espesor not in available_espesores:
    espesor_use = max([e for e in available_espesores if e <= espesor], default=available_espesores[0])
    st.warning(f"Espesor {espesor} mm no exacto. Usando el más cercano: {espesor_use} mm.")
else:
    espesor_use = espesor
# Buscar velocidades y VDI (internamente, no se muestran)
velocities = [0.0] * 3
times = [0.0] * 3
vdi_final = 0
for i, code in enumerate(codes):
    busqueda = f"{espesor_use}-{code}"
    row = df_tablas[df_tablas['Búsqueda'] == busqueda]
    if not row.empty:
        vel = row['Velocidad'].values[0]
        vdi = row['VDI'].values[0]
        velocities[i] = vel
        if vel > 0:
            times[i] = longitud / vel
        vdi_final = vdi
# Cálculos
total_min = sum(times)
total_h = total_min / 60
costo = round(total_h * tasa + costo_fijo, 2)
ra_estimado = f"{vdi_to_ra.get(vdi_final, 'Desconocido')}"
# Botón para solicitar
if st.button("Solicitar Presupuesto"):
    if not email:
        st.error("Ingresa tu email.")
    else:
        # Conectar a Google Sheet usando secrets
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gsheets"], scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_id("TU_SHEET_ID").sheet1 # Reemplaza TU_SHEET_ID con el ID de tu Sheet
        emails = [row[0].lower() for row in sheet.get_all_values()[1:] if row]
        cuerpo = f"Presupuesto:\nMaterial: {material}\nEspesor: {espesor} mm\nCalidad: {calidad}\nLongitud: {longitud} mm\nCosto: {costo} €\nVDI: {vdi_final}\nRa: {ra_estimado}"
        if email.lower() in emails:
            # Enviar email al cliente
            msg = MIMEText(cuerpo)
            msg['Subject'] = "Tu Presupuesto Servicorte"
            msg['From'] = "servicorteporhilo@servicorteporhilo.es"
            msg['To'] = email
            server = smtplib.SMTP('smtp.nominalia.com', 587) # SMTP de Nominalia
            server.starttls()
            server.login(st.secrets["email"]["servicorteporhilo@servicorteporhilo.es"], st.secrets["email"]["47704349Aa"])
            server.send_message(msg)
            server.quit()
            st.success("Presupuesto enviado a tu email.")
        else:
            # Enviar notificación a ti
            msg = MIMEText(f"Nueva solicitud de {email}:\n{cuerpo}")
            msg['Subject'] = "Nueva Solicitud Presupuesto"
            msg['From'] = "servicorteporhilo@servicorteporhilo.es"
            msg['To'] = "servicorteporhilo@servicorteporhilo.es"
            server = smtplib.SMTP('smtp.nominalia.com', 587)
            server.starttls()
            server.login(st.secrets["email"]["servicorteporhilo@servicorteporhilo.es"], st.secrets["email"]["47704349Aa"])
            server.send_message(msg)
            server.quit()
            st.warning("Solicitud enviada. Te contactaremos si eres cliente registrado.")
# Resumen sin costo
calc_data = {
    'Descripción': [
        'Material', 'Espesor (mm)', 'Calidad', 'Longitud corte (mm)',
        'VDI final', 'Acabado estimado (µm Ra)'
    ],
    'Valor': [
        material, espesor, calidad, longitud,
        vdi_final, ra_estimado
    ]
}
df_calc = pd.DataFrame(calc_data)
st.header("Resumen (presupuesto por email)")
st.dataframe(df_calc)
# Descargar (opcional, sin costo)
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df_calc.to_excel(writer, sheet_name='Resumen', index=False)
st.download_button("Descargar resumen (sin precio)", output.getvalue(), "resumen.xlsx")

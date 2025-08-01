import pandas as pd
import win32com.client as win32
import os
from datetime import datetime

# Ruta del archivo Excel
ruta_excel = 'C:\Users\stamayo\Documents\Paz y Salvo\Correo Masivo.xlsx'  # <-- Modifica esta ruta

# Leer archivo
df = pd.read_excel(ruta_excel)

# Iniciar Outlook
outlook = win32.Dispatch('Outlook.Application')

# Lista para registrar estado
estado_envio = []

# Recorrer cada fila
for index, row in df.iterrows():
    correo = row[0]
    asunto = row[1]
    adjunto = row[2]

    if pd.isna(correo) or pd.isna(asunto) or pd.isna(adjunto):
        estado_envio.append("Datos incompletos")
        print(f"Fila {index + 2}: Datos incompletos, se omite.")
        continue

    if not os.path.isfile(adjunto):
        estado_envio.append("Archivo no encontrado")
        print(f"Fila {index + 2}: Archivo no encontrado: {adjunto}")
        continue

    try:
        mail = outlook.CreateItem(0)
        mail.To = correo
        mail.Subject = asunto
        mail.Body = "Estimado/a,\n\nAdjunto encontrará el archivo solicitado.\n\nSaludos cordiales."
        mail.Attachments.Add(adjunto)

        # *** Cambiar remitente ***
        mail.SentOnBehalfOfName = "info@yamahamotor-financiera.com.co"

        mail.Send()
        # mail.Display()  # <-- Descomenta si quieres revisar antes de enviar
        estado_envio.append("Enviado")
        print(f"Correo enviado a {correo}")
    except Exception as e:
        print(f"Error en fila {index + 2}: {e}")
        estado_envio.append(f"Error: {e}")

# Guardar con columna de estado y respaldo histórico
df["Estado"] = estado_envio
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
ruta_salida = ruta_excel.replace('.xlsx', f'_enviados_{timestamp}.xlsx')
df.to_excel(ruta_salida, index=False)

print(f"\nProceso finalizado. Se guardó el archivo con el histórico en:\n{ruta_salida}")

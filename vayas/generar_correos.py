import pandas as pd
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# Función para generar un correo a partir de la lista de distribuidores con facturas vencidas
def generar_correo(distribuidor, email, facturas):
    msg = MIMEMultipart()
    msg['From'] = "tu_email@example.com"
    msg['To'] = email
    msg['Subject'] = f"Recordatorio de pago de {len(facturas)} facturas vencidas"

    body = f"Estimado/a {distribuidor},\n\n"
    body += f"Hemos identificado que tiene {len(facturas)} facturas pendientes de pago que ya se encuentran vencidas. A continuación, le proporcionamos los detalles:\n\n"

    # Añadir al cuerpo del correo cada factura vencida
    for factura in facturas:
        body += f"- # CF: {factura['# CF']}, Banco: {factura['BANCO']}, Monto: {factura['MONTO']} {factura['MONEDA']}, Fecha de Vencimiento: {factura['FECVCTO'].strftime('%d/%m/%Y')}\n"

    body += "\nPor favor, regularice su situación lo antes posible para evitar inconvenientes futuros.\n\n"
    body += "Atentamente,\nSu equipo de finanzas"

    msg.attach(MIMEText(body, 'plain'))
    return msg

# Establecer las rutas
ruta_base = os.path.dirname(os.path.abspath(__file__))
ruta_excel = os.path.join(ruta_base, 'facturacion.xlsx')
ruta_correos = os.path.join(ruta_base, 'correos_generados')

# Crear la carpeta para los correos si no existe
os.makedirs(ruta_correos, exist_ok=True)

# Leer el archivo Excel
df = pd.read_excel(ruta_excel)

# Convertir la columna de fecha de vencimiento a datetime
df['FECVCTO'] = pd.to_datetime(df['FECVCTO'], format='%d/%m/%Y')

# Filtrar las facturas vencidas
hoy = datetime.datetime.today()
facturas_vencidas = df[df['FECVCTO'] < hoy]

# Crear un diccionario para almacenar las facturas por distribuidor
deudores = {}

# Agrupar las facturas por distribuidor y correo electrónico
for idx, row in facturas_vencidas.iterrows():
    distribuidor = row['Distribuidor']
    email = row['correos']
    
    # Si el distribuidor ya está en el diccionario, añadir la nueva factura a su lista
    if distribuidor in deudores:
        deudores[distribuidor]['facturas'].append(row)
    else:
        # Si es la primera vez que encontramos este distribuidor, crear su entrada en el diccionario
        deudores[distribuidor] = {
            'email': email,
            'facturas': [row]
        }

# Generar y guardar los correos para cada distribuidor en el diccionario de deudores
for distribuidor, data in deudores.items():
    correo = generar_correo(distribuidor, data['email'], data['facturas'])
    
    # Guardar cada correo en un archivo .eml en la carpeta correos_generados
    correo_filename = f"correo_{distribuidor.replace(' ', '_').replace('@', '_at_')}.eml"
    with open(os.path.join(ruta_correos, correo_filename), "w") as f:
        f.write(correo.as_string())

    print(f"Correo generado para {distribuidor} ({data['email']}) con {len(data['facturas'])} facturas vencidas.")


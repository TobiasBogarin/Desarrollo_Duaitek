import requests
import logging
import csv
from datetime import datetime, timedelta

# Configuración básica de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def obtener_token_mercadolibre(client_id, client_secret):
    url = 'https://api.mercadolibre.com/oauth/token'
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret
    }
    try:
        response = requests.post(url, data=data)
        response.raise_for_status()
        token = response.json().get('access_token')
        logger.info(f"Token de acceso para Client ID {client_id} obtenido con éxito.")
        return token
    except requests.RequestException as e:
        logger.error(f"Error al obtener token para Client ID {client_id}: {e}")
        return None

def consultar_ordenes_por_fecha(access_token, fecha_inicio, fecha_fin, seller_id, seen_shipment_ids):
    orders = []
    offset = 0
    limit = 50

    while True:
        url = (f"https://api.mercadolibre.com/orders/search?seller={seller_id}"
               f"&order.date_created.from={fecha_inicio}&order.date_created.to={fecha_fin}"
               f"&offset={offset}&limit={limit}&access_token={access_token}")

        headers = {'Authorization': f'Bearer {access_token}'}
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json().get('results', [])
            if not data:
                break

            for order in data:
                shipment_id = order.get('shipping', {}).get('id')
                if shipment_id and shipment_id not in seen_shipment_ids:
                    orders.append(order)
                    seen_shipment_ids.add(shipment_id)

            if len(data) < limit:
                break

            offset += limit
        except requests.RequestException as e:
            logger.error(f"Error al consultar órdenes para seller {seller_id}: {e}")
            break

    return orders

def extraer_shipment_ids_y_estados(ordenes):
    shipment_info = []
    for orden in ordenes:
        shipment_id = orden.get('shipping', {}).get('id')
        status = orden.get('status', 'N/A')
        shipment_info.append((shipment_id, status))
    return shipment_info

def guardar_shipment_info_en_csv(shipment_info, nombre_archivo):
    with open(nombre_archivo, mode='a', newline='') as file:
        writer = csv.writer(file)
        for shipment_id, status in shipment_info:
            writer.writerow([shipment_id, status])

def divide_fecha_en_semanas(fecha_inicio, fecha_fin):
    start = datetime.strptime(fecha_inicio, "%Y-%m-%dT%H:%M:%SZ")
    end = datetime.strptime(fecha_fin, "%Y-%m-%dT%H:%M:%SZ")
    while start < end:
        next_week = start + timedelta(days=7)
        yield start.strftime("%Y-%m-%dT%H:%M:%SZ"), min(next_week, end).strftime("%Y-%m-%dT%H:%M:%SZ")
        start = next_week

def main():
    cuentas = [
        {"client_id": "8545932337071872", "client_secret": "VtRNp52gxGKCDKLvR8YRfovPLL3ZBfIk", "seller_id": "264490310"},
        {"client_id": "8595972831111833", "client_secret": "fKyvOSCw3vhI2fDKqroEY8zdcd77ZN5L", "seller_id": "39384567"}
    ]
    fecha_inicio = '2024-08-01T00:00:00Z'
    fecha_fin = '2024-08-07T23:59:59Z'
    nombre_archivo = "shipment_info.csv"
    seen_shipment_ids = set()

    # Crear el archivo CSV y escribir los encabezados solo una vez
    with open(nombre_archivo, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["shipment_id", "status"])

    for cuenta in cuentas:
        access_token = obtener_token_mercadolibre(cuenta['client_id'], cuenta['client_secret'])
        if access_token:
            for start, end in divide_fecha_en_semanas(fecha_inicio, fecha_fin):
                ordenes = consultar_ordenes_por_fecha(access_token, start, end, cuenta['seller_id'], seen_shipment_ids)
                if ordenes:
                    shipment_info = extraer_shipment_ids_y_estados(ordenes)
                    guardar_shipment_info_en_csv(shipment_info, nombre_archivo)
                    logger.info(f"Datos guardados para la semana desde {start} hasta {end} para el seller {cuenta['seller_id']}.")
                else:
                    logger.warning(f"No se encontraron órdenes para la semana desde {start} hasta {end} para el seller {cuenta['seller_id']}.")
        else:
            logger.error("No se pudo obtener el token de acceso, verifique las credenciales y la conexión a Internet.")

if __name__ == "__main__":
    main()

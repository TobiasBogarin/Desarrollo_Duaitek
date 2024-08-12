import requests
import logging

# Configuración básica de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def obtener_token_mercadolibre(client_id, client_secret):
    """
    Obtiene un token de acceso de la API de Mercado Libre usando las credenciales proporcionadas.
    """
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
        logger.info("Token de acceso obtenido con éxito.")
        return token
    except requests.RequestException as e:
        logger.error(f"Error al obtener token: {e}")
        return None

def obtener_datos_de_envio(access_token, shipment_id):
    """
    Obtiene los datos de un envío en Mercado Libre usando las credenciales proporcionadas.
    """
    url = f"https://api.mercadolibre.com/shipments/{shipment_id}"
    headers = {'Authorization': f'Bearer {access_token}'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        logger.error(f"Error al obtener datos de envío para {shipment_id}: {e}")
        return {}

def main():
    client_id = '8595972831111833'
    client_secret = 'fKyvOSCw3vhI2fDKqroEY8zdcd77ZN5L'

    # ID del envío a consultar
    shipment_id = '43670404869'

    access_token = obtener_token_mercadolibre(client_id, client_secret)
    if access_token:
        envio_data = obtener_datos_de_envio(access_token, shipment_id)
        if envio_data:
            logger.info("Datos del envío obtenidos:")
            logger.info(envio_data)
        else:
            logger.warning("No se pudieron obtener los datos del envío.")
    else:
        logger.error("No se pudo obtener el token de acceso, verifique las credenciales y la conexión a Internet.")

if __name__ == "__main__":
    main()

import requests
from requests.auth import HTTPBasicAuth

# Configura la URL de tu tienda y la clave API
prestashop_url = 'http://todomicro.com.ar/api'
api_key = 'ARAW1J81WVJD22YSH2GQ58VNFUGQ9YRG'
order_id = '70360'  # Reemplaza con el ID de la orden que deseas consultar

# Configura los encabezados para aceptar JSON
headers = {
    'Accept': 'application/json'
}

# Realiza la solicitud GET a la API de PrestaShop
response = requests.get(
    f'{prestashop_url}/orders/{order_id}',
    headers=headers,
    auth=HTTPBasicAuth(api_key, '')
)

# Verifica el estado de la respuesta
if response.status_code == 200:
    order_data = response.json()
    print(order_data)
else:
    print(f'Error: {response.status_code}')
    print('Headers:', response.headers)
    print('Response:', response.text)

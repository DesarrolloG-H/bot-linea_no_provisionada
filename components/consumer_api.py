import requests

def getdata(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()  # Devuelve la respuesta en formato JSON
    except requests.exceptions.RequestException as e:
        print(f"Error en la solicitud: {e}")
        return None


import win32com.client
import win32com.client as win32
import os

def conectar_a_outlook():
    """
    Conecta a la aplicación de Outlook y devuelve el objeto Namespace.
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    print("Se conecto a outlook")
    return outlook

def obtener_cuenta_outlook(outlook,nombre_cuenta):
    """
    Obtiene una cuenta específica de Outlook por nombre.
    """
    try:
        # Conectar a Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Obtener todas las cuentas configuradas
        for i in range(1, outlook.Folders.Count + 1):
            cuenta = outlook.Folders.Item(i)
            if cuenta.Name == nombre_cuenta:
                print(f"Cuenta encontrada: {cuenta.Name}")
                return cuenta

        print(f"No se encontró la cuenta con el nombre: {nombre_cuenta}")
        return None
    except Exception as e:
        print(f"Error al obtener la cuenta de Outlook: {e}")
        return None
    

def enviar_correo_con_adjunto(destinatario, asunto, cuerpo,archivo_adjunto, cc=None, imagenes=None):
    """Envía un correo con archivo adjunto."""
    try:
        # Obtener el objeto Namespace para acceder a la cuenta
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        mail = outlook.CreateItem(0)  # 0 para correo

        # Configuración del correo
        mail.To = "; ".join(destinatario) if destinatario else ""
        mail.CC = "; ".join(cc) if cc else ""
        mail.Subject = asunto
        mail.HTMLBody = cuerpo

        # Adjuntar los archivos si son proporcionados
        if archivo_adjunto:
            mail.Attachments.Add(archivo_adjunto)
        
        # Enviar el correo
        mail.Send()
        print(f"Correo enviado a {destinatario}")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")

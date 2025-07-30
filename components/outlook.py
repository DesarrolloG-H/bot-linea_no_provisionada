import os
import psutil
import win32gui
import win32con
import time

def esta_outlook_abierto():
    """
    Verifica si Outlook ya está abierto.
    """
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name']:
            return True
    return False

def minimizar_outlook():
    """
    Minimizamos la ventana de Outlook si está abierta.
    """
    try:
        def enum_windows_callback(hwnd, results):
            # Callback para buscar la ventana de Outlook
            if 'Bandeja de entrada' in win32gui.GetWindowText(hwnd).lower():
                results.append(hwnd)

        results = []
        win32gui.EnumWindows(enum_windows_callback, results)
        if results:
            for hwnd in results:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            print("Outlook se ha minimizado.")
        else:
            print("No se encontró una ventana de Outlook para minimizar.")
    except Exception as e:
        print(f"Error al minimizar Outlook: {e}")

def abrir_outlook():
    """
    Abre Outlook si no está abierto y lo minimiza.
    """
    try:
        # Ruta del ejecutable de Outlook, ajusta según tu sistema
        outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"

        if esta_outlook_abierto():
            print("Outlook ya está abierto.")
        else:
            if os.path.exists(outlook_path):
                os.startfile(outlook_path)
                print("Outlook se ha abierto correctamente.")
            else:
                print(f"No se encontró Outlook en la ruta especificada: {outlook_path}")
        time.sleep(10)
        # Minimizar Outlook después de abrirlo
        minimizar_outlook()
    except Exception as e:
        print(f"Error al abrir Outlook: {e}")

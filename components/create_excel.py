import pandas as pd

# def create_file(data_list, filename):
#     with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
#         for data in data_list:
#             if data:
#                 print(data['TELEFONO'])
#                 # Extraer nombres de columnas y convertir la data a DataFrame
#                 columns = [col['TELEFONO'] for col in data]
#                 df = pd.DataFrame(data['TELEFONO'], columns=columns)
#                 df.to_excel(writer, sheet_name='Linea', index=False)



def create_file(data_list, filename):
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        lista_telefonos = []
        for data in data_list:
            if data:
                print(data['TELEFONO'])
                lista_telefonos.append(data['TELEFONO'])
                # data es una lista de diccionarios con clave 'TELEFONO'
        df = pd.DataFrame(lista_telefonos,columns=["TELEFONO"])  # convierte toda la lista en una tabla
        sheet_name = 'Linea'
        df.to_excel(writer, sheet_name=sheet_name, index=False)
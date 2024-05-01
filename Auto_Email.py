import win32com.client as client
import pandas as pd

try:
    excel_file = '/Desktop/Empleados.xlsx' #Archivo con la información sobre los empleados del departamento
    df_empleados= pd.read_excel(excel_file)
    outlook = client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    account = namespace.Folders['dirección de correo'] #Acá debe de ir la dirección del correo con el que se quiere trabajar
    inbox = account.Folders['Inbox']
    departamento_folder = inbox.Folders['Departamento'] #Este es el folder destino dentro del correo al cual iran los correos categorizados
    excel_file = 'Desktop/History.xlsx' #Archivo con los casos y su respectivo dueño
    df_Casos = pd.read_excel(excel_file) #La función lee el documento de excel y crea un data frame (como una tabla)


    nombres_Empl = [] #Nombre de la firma de los empleados (se utiliza en la sección de busqueda si en el subject viene la palabra "Investigation")
    nom_comp_Empl = [] #Nombre completo de los empleados (se utiliza para buscar el dueño del caso)
    categorias_Empl = [] #Categoría del correo de cada empleado

    # -------------------Llenar las tablas para Empleados---------------------------------------------

    for index in df_empleados.index:
        nombres.append(df_empleados['Nombre'][index])

    for index in df_empleados.index:
        categorias.append(df_empleados['Categoria'][index])

    for index in df_empleados.index:
        nom_comp.append(df_empleados['Nom_Comp'][index])

#-------------------------Llenar la lista de mensajes del inbox--------------------------
    mensajes = list(inbox.Items)

    # Lista para mantener un registro de los IDs de los mensajes procesados
    ids_procesados = []
#-----------------Revisar owner de closure para cartas de Investigación---------------
    for message in mensajes:
            checks = []
            #En esta sección el programa, para todos los correos que en el subject tenga la palabra "Investigation", buscará el nombre en la firma del correo y lo comparará con la lista de empleados
            #Si hace un match, el programa le dará al correo la categoría que pertenece a ese empleado y lo enviará a la carpeta del departamento respectivo
            if "Investigation" in message.Subject:
                for empleado in nombres:
                    if empleado in message.Body.lower():
                        checks.append(empleado)
                        categoria = categorias[nombres.index(checks[0])]
                        message.Categories = categoria
                        message.move(departamento_folder)
                    else: continue

            # -----------------Categorización del correo por número de caso---------------
            # En esta sección el programa, para todos los correos que en el subject NO tenga la palabra "Investigation", buscará en el archivo "History" el número de caso que viene en el subjecto (este número es un identificador único)
            # Si hace un match, el programa le dará al correo la categoría del empleado que es dueño de dicho caso y lo enviará a la carpeta del departamento correspondiente.
            for index in df_Casos.index:
                caso = str(df_Casos['Case Number'][index])

                found = False

                if caso in message.Subject and "Investigation" not in message.Subject:
                    owner = df_Casos['Case Owner'][index]
                    #print(owner)


                    # Verificar si el ID del mensaje ya está en la lista de IDs procesados, para evitar que el programa genere un error al tener un correo con dos o más números de caso
                    if message.EntryID in ids_procesados:
                        continue
                    else:
                        for nombre in nom_comp:
                            if owner == nombre:
                                posicion = nom_comp.index(nombre)
                                message.Categories = categorias[posicion]
                                message.move(departamento_folder)
                                ids_procesados.append(message.EntryID)  # Agregar el ID del mensaje a la lista de IDs procesados
                            else: continue

except Exception as e:
    print("An error occurred:", e)
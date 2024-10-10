import pandas as pd
import numpy as np
import requests as req
import os
import customtkinter as ctk
from tkinter import filedialog, ttk
from tkcalendar import DateEntry
from datetime import datetime
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

# ACCEDER A LA BASE DE DATOS DE SOFTLAND. HAY UNA TABLA QUE SE LLAMA
# GESTIÓN DE NÓMINA, TABLA DE MANTENIMIENTO, LOS EMPLEADOS INACTIVOS TIENEN ID- NOMRBE, NO ES PRECISO.
# COLUMNAS DE Código y Nombre. PEDIR EL EXCEL QUE SÍ SIRVE.
# URL to send the GET request
url = "http://10.10.100.50:8001/empleado"

# Parameters to include in the GET request
params = {'code': 'C01'}

# Initialize the customtkinter window
root = ctk.CTk()
root.title("Nóminas Concatenator")
root.geometry("800x600")

# Create a Tabview to manage different tabs
tabview = ctk.CTkTabview(root, width=800, height=600)
tabview.pack(fill="both", expand=True)

tab = {}
# Add two tabs: Satori and Hispizza
tab["Satori"] = tabview.add("Satori")
tab["Hispizza"] = tabview.add("Hispizza")

# Define the column headers manually
headers = ["EMPLEADO", "NOMINA", "CONCEPTO", "CANTIDAD", "MONTO"]
headersPropina = ['CODIGO', 'NOMBRE', 'MONTO PROPINA VOLUNTARIA', '% DESC. TC.', 'PROPINA NETO', 'columna6', 'MONTO',
                  'BARTENDER', 'columna9']

df_SatoriCoop = pd.DataFrame()
# formato_column = ["EMPLEADO","NOMINA","CONCEPTO","CANTIDAD","MONTO"]
df_out = {"Satori": [], "Hispizza": []}
df_individual = {"Satori": {}, "Hispizza": {}}

########## VARIABLES DE ESTADO DE CAMPOS
VarState = {"Satori": {}, "Hispizza": {}}
VarState["Satori"]["Cooperativa"] = True
VarState["Satori"]["Propina Voluntaria"] = True
VarState["Satori"]["Horas consolidado"] = True
VarState["Satori"]["Desc. Gym"] = False

VarState["Hispizza"]["Cooperativa"] = True
VarState["Hispizza"]["Fripick"] = True
VarState["Hispizza"]["Farmacia"] = True
VarState["Hispizza"]["Horas CAC"] = True
VarState["Hispizza"]["Horas PBI"] = True
VarState["Hispizza"]["Incentivos"] = True
VarState["Hispizza"]["Licencias Med"] = False

######### Conceptos
ConceptoPropina = "0001-B017"
ConceptoFripick = "0001-D006"
ConceptoFarmacia = "0001-D024"
ConceptoIncentivosCAC = "0001-B009"

concepto_mapping = {
    '0431': '0001-D009',  # Referente a 0431
    '0405': '0001-D010',  # Referente a 0405
    '0416': '0001-D011',  # Referente a 0416
    431: '0001-D009',  # Referente a 0431
    405: '0001-D010',  # Referente a 0405
    416: '0001-D011'  # Referente a 0416
}


# horasPBI_mapping = {
#     'Salario Quincenal': '0001-B001',
#     'Horas Extra 35%': '0001-B002',
#     'Nocturnidad 15%': '0001-B004',
#     'Horas Feriadas Laboradas': '0001-B014',
#     'Horas Extras Feriadas/ Extra Dia Libre': '0001-B016'
# }

# ConceptoHorasPBI = "0001-D024"
# Concepto0431 = "0001-D009" # referente a Desc. Ahorro Cooperativa
# Concepto0405 = "0001-D010" # referente a Desc. Préstamos Cooperativa
# Concepto0416 = "0001-D011" # referente a Desc. Servicios Cooperativa


def display_table_in_treeview(df, tree):
    # Clear existing data in the treeview
    for i in tree.get_children():
        tree.delete(i)

    # Create columns dynamically from DataFrame
    tree["column"] = list(df.columns)
    tree["show"] = "headings"

    # Format columns
    for col in tree["column"]:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor='center')

    # Add rows to Treeview
    for index, row in df.iterrows():
        tree.insert("", "end", values=list(row))


# Function to match names using fuzzy matching
def match_names(name, choices, threshold=80):
    # Ensure name is a string (convert float or NaN to empty string)
    if isinstance(name, str):
        result = process.extractOne(name, choices, scorer=fuzz.token_set_ratio, score_cutoff=threshold)
        if result:
            return result[0]  # Return the best match
    return None


# Function to get the first word from a column name
def get_first_word(column_name):
    try:
        return column_name.split()[0]
    except:
        return column_name.tostring()


# Adaptado para el archivo cooperativo de todas las empresas.
def load_and_concatenate_tables(path, grupo, date_entry):
    dataframes = []
    reorderCol = ["Empleados", "Nomina", "Concepto", "cant2", "Monto"]
    df_individual[grupo]["Cooperativa"] = []
    # Get the selected date from the date entry
    selected_date = date_entry.get_date().strftime('%d.%m.%y')

    # Loop through all .txt files in the specified folder
    for filename in os.listdir(path):
        if filename.endswith(".txt") or filename.endswith(".TXT"):
            file_path = os.path.join(path, filename)
            df = pd.read_csv(file_path, sep='\t', header=None, names=reorderCol)  # Adjust 'sep' as needed

            # Remove single quotes from all string values using replace
            df = df.replace("'", "", regex=True)

            # Convert 'Empleados' column to whole numbers (integers) and then back to text
            df['Empleados'] = df['Empleados'].astype(float).astype(int).astype(str)

            # Convert the 'Nomina' column to datetime with dayfirst=True
            df['Nomina'] = pd.to_datetime(df['Nomina'], dayfirst=True, errors='coerce')

            # Format the datetime object to the desired format 'dd.mm.yy'
            df['Nomina'] = df['Nomina'].dt.strftime('%d.%m.%y')

            # Filter rows based on the selected date
            df_filtered = df[df['Nomina'] == selected_date]
            if not df_filtered.empty:
                dataframes.append(df_filtered)

    if dataframes:
        df_individual[grupo]["Cooperativa"] = pd.concat(dataframes, ignore_index=True)
        # df_individual[grupo]["Cooperativa"].drop('cant2', axis=1, inplace=True)
        df_individual[grupo]["Cooperativa"]["cant2"] = 0.0
        df_individual[grupo]["Cooperativa"]["Concepto"] = df_individual[grupo]["Cooperativa"]["Concepto"].replace(
            concepto_mapping)

        # print(df_individual[grupo]["Cooperativa"])
        # df_individual[grupo]["Cooperativa"] = df_individual[grupo]["Cooperativa"][reorderCol]
        df_individual[grupo]["Cooperativa"] = df_individual[grupo]["Cooperativa"][reorderCol]
        df_individual[grupo]["Cooperativa"].columns = headers
        df_individual[grupo]["Cooperativa"]["NOMINA"] = "0001"

        # Display the DataFrame in the text box
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, df_individual[grupo]["Cooperativa"].to_string(index=False))
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, df_individual[grupo]["Cooperativa"].to_string(index=False))
        # pretty_print_with_colors(df_individual[grupo]["Cooperativa"], result_textbox)
        # pretty_print_with_colors(df_individual[grupo]["Cooperativa"], result_textbox2)

    else:
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, "No matching .txt files found for the selected date.")
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, "No matching .txt files found for the selected date.")
    print(VarState)


# SOLO DISEÑADO PARA SATORI, ESCALAR CUANDO SEA NECESARIO
# RENOMBRAR PARA PROPINA VOLUNTARIA.
def load_excel_by_date(date_entry, grupo):
    print(grupo)
    # Get the selected date from the DateEntry widget
    selected_date = date_entry.get_date().strftime('%d-%m-%y')

    # Define the folder and file path
    folder_path = "./Satori/Propina Voluntaria"
    excel_file = None

    # Find the Excel file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            excel_file = os.path.join(folder_path, filename)
            break

    if excel_file:
        try:
            # Load the Excel file to get all sheet names
            xl = pd.ExcelFile(excel_file)
            sheet_names = xl.sheet_names

            # Convert the sheet names to datetime objects
            sheet_dates = {}
            for sheet in sheet_names:
                try:
                    sheet_date = datetime.strptime(sheet, '%d-%m-%y')
                    sheet_dates[sheet_date] = sheet
                except ValueError:
                    continue

            # Find the sheet matching the selected date
            sheet_to_load = sheet_dates.get(datetime.strptime(selected_date, '%d-%m-%y'))

            if sheet_to_load:
                sheet_df = xl.parse(sheet_to_load)

                # Set the first row as the headers
                sheet_df.columns = sheet_df.iloc[0]
                sheet_df = sheet_df[1:].reset_index(drop=True)
                sheet_df.columns = headersPropina

                # Display the DataFrame in the text box
                result_textbox.delete(1.0, ctk.END)
                result_textbox.insert(ctk.END, sheet_df.to_string(index=False))

                # Define the columns for PropVolDF
                vol_columns = ['CODIGO', 'NOMBRE', 'MONTO PROPINA VOLUNTARIA', '% DESC. TC.', 'PROPINA NETO']

                # Separate the DataFrame into PropVolDF and PropBarDF
                global PropVolDF, PropBarDF
                sheet_df = sheet_df.dropna(subset=['NOMBRE'])
                PropVolDF = sheet_df[vol_columns]
                PropVolDF['CODIGO'] = PropVolDF['CODIGO'].astype(float).astype(int).astype(str)

                PropBarDF = sheet_df.drop(columns=vol_columns)
                PropVolDF.drop('NOMBRE', axis=1, inplace=True)
                PropVolDF.drop('MONTO PROPINA VOLUNTARIA', axis=1, inplace=True)
                PropVolDF.drop('% DESC. TC.', axis=1, inplace=True)

                PropVolDF['NOMINA'] = "0001"
                PropVolDF['CANTIDAD'] = 0.0
                PropVolDF['CONCEPTO'] = ConceptoPropina
                PropVolDF.rename(columns={'CODIGO': 'EMPLEADO'}, inplace=True)
                PropVolDF.rename(columns={'PROPINA NETO': 'MONTO'}, inplace=True)

                # PropVolDF.columns = headers

                PropVolDF = PropVolDF[headers]
                df_individual["Satori"]["Propina Voluntaria"] = PropVolDF

                # Display the PropVolDF DataFrame in the text box
                result_textbox.delete(1.0, ctk.END)
                result_textbox.insert(ctk.END, PropVolDF.to_string(index=False))
                result_textbox2.delete(1.0, ctk.END)
                result_textbox2.insert(ctk.END, PropVolDF.to_string(index=False))

            else:
                result_textbox.delete(1.0, ctk.END)
                result_textbox2.delete(1.0, ctk.END)
                result_textbox.insert(ctk.END, f"No sheet found for the date: {selected_date}")
                result_textbox2.insert(ctk.END, f"No sheet found for the date: {selected_date}")

        except ValueError:
            result_textbox.delete(1.0, ctk.END)
            result_textbox.insert(ctk.END, "Error reading Excel file or matching sheet.")
            result_textbox2.delete(1.0, ctk.END)
            result_textbox2.insert(ctk.END, "Error reading Excel file or matching sheet.")
    else:
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, "No Excel file found in the directory.")
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, "No Excel file found in the directory.")


def load_fripick(path, grupo, date_entry):
    # Adaptado para el archivo de fripick de todas las empresas.
    fripick_columns = ["CODIGO EMPLEADO", "NOMBRE", "MONTO A DESCONTAR"]
    xl = []
    df_individual[grupo]["Fripick"] = []

    # Define the folder and file path
    folder_path = os.path.join(path, date_entry.get_date().strftime('%d-%m-%y'))
    excel_file = None
    print(folder_path)
    # Find the Excel file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            excel_file = os.path.join(folder_path, filename)
            break
        print(filename)
    print(date_entry.get_date().strftime('%d-%m-%y'))

    if excel_file:
        try:
            # Load the Excel file to get all sheet names
            xl = pd.read_excel(excel_file)
            xl.columns = xl.columns.str.replace('  ', ' ')
            xl = xl.dropna(subset=['NOMBRE'])

            df_individual[grupo]["Fripick"] = xl[fripick_columns]
            df_individual[grupo]["Fripick"]['NOMINA'] = "0001"
            df_individual[grupo]["Fripick"]['CANTIDAD'] = 0.0
            df_individual[grupo]["Fripick"]['CONCEPTO'] = ConceptoPropina
            df_individual[grupo]["Fripick"].drop('NOMBRE', axis=1, inplace=True)

            # Specify the new column order
            new_order = ['CODIGO EMPLEADO', 'NOMINA', 'CONCEPTO', 'CANTIDAD', 'MONTO A DESCONTAR']

            # Reorder the DataFrame columns
            df_individual[grupo]["Fripick"] = df_individual[grupo]["Fripick"][new_order]
            df_individual[grupo]["Fripick"].columns = headers
            print(df_individual[grupo]["Fripick"])

            result_textbox.delete(1.0, ctk.END)
            result_textbox.insert(ctk.END, df_individual[grupo]["Fripick"].to_string(index=False))
            result_textbox2.delete(1.0, ctk.END)
            result_textbox2.insert(ctk.END, df_individual[grupo]["Fripick"].to_string(index=False))

        except ValueError:
            result_textbox.delete(1.0, ctk.END)
            result_textbox.insert(ctk.END, "Error reading Excel file or matching sheet.")
            result_textbox2.delete(1.0, ctk.END)
            result_textbox2.insert(ctk.END, "Error reading Excel file or matching sheet.")
    else:
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, "No Excel file found in the directory.")
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, "No Excel file found in the directory.")


def load_farmacia(path, grupo, date_entry):
    # NOTA: el mes de mayo entero se paga la cobertura en dos quincenas, por lo que el archivo del 15 y del 30 van a ser el mismo.
    # Sending the GET request
    response = req.get(url, params=params)

    # Checking if the request was successful (status code 200)
    if response.status_code == 200:
        # Print the content of the response (in JSON format if available)
        try:
            data = response.json()  # Assuming the response is in JSON format
            # print(data)
        except ValueError:
            print(response.text)  # If not JSON, print raw text
    else:
        print(f"Request failed with status code: {response.status_code}")

    # Convert to DataFrame
    empleados = pd.DataFrame(data)
    empleados = empleados[["EMPLEADO", "NOMBRE"]]
    # print(empleados["NOMBRE"])
    # print(path, grupo, date_entry)

    # Get the selected date from the DateEntry widget
    selected_date = date_entry.get_date().strftime('%d-%m-%y')
    folder_path = os.path.join(path, selected_date)
    # print(folder_path)
    # Define the folder and file path
    excel_file = None

    # Find the Excel file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith("Corporativa.xlsx"):
            excel_file = os.path.join(folder_path, filename)
            print(excel_file)
            break

    xl = pd.ExcelFile(excel_file)

    sheet_name = xl.sheet_names[0]
    # sheet_to_load = sheet_names.get()

    if sheet_name:
        sheet_df = xl.parse(sheet_name)
        sheet_df = sheet_df[["Código", "Cliente", "Cobertura"]]
        sheet_df = sheet_df.dropna(subset=['Cliente'])
        sheet_df.drop('Código', axis=1, inplace=True)

        # Remove extra spaces in both DataFrames
        sheet_df['Cliente'] = sheet_df['Cliente'].str.strip().str.replace(' +', ' ', regex=True)
        empleados['NOMBRE'] = empleados['NOMBRE'].str.strip().str.replace(' +', ' ', regex=True)
        sheet_df['Best Match'] = sheet_df['Cliente'].apply(
            lambda x: match_names(str(x) if isinstance(x, str) else '', empleados['NOMBRE'].tolist()))

        ## RELLENAR LOS CÓDIGOS QUE NO TIENEN VALOR CON LOS VALORES RESPECTIVOS DE LA TABLA DE EMPLEADOS Y PREPARAR FORMATO
        completeXL = pd.merge(sheet_df, empleados, left_on='Best Match', right_on='NOMBRE', how='left')
        print("tipos")

        # print(empleados)
        # print(sheet_df)
        completeXL["MONTO"] = completeXL["Cobertura"] / 2
        completeXL['NOMINA'] = "0001"
        completeXL['CANTIDAD'] = 0.0
        completeXL['CONCEPTO'] = ConceptoFarmacia
        # print(completeXL[["NOMBRE","EMPLEADO"]])
        completeXL.drop('Cliente', axis=1, inplace=True)
        completeXL.drop('Cobertura', axis=1, inplace=True)
        completeXL.drop('Best Match', axis=1, inplace=True)
        completeXL.drop('NOMBRE', axis=1, inplace=True)
        # completeXL = completeXL.rename(columns={'AAAAAA': 'EMPLEADO'})
        print(completeXL)

        ## REVISAR

        completeXL = completeXL[headers]

        print(completeXL.columns)

        # completeXL.rename(columns={'C': 'EMPLEADO'}, inplace=True)

    if isinstance(completeXL, pd.DataFrame):

        df_individual[grupo]["Farmacia"] = completeXL

        # Display the DataFrame in the text box
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, df_individual[grupo]["Farmacia"].to_string(index=False))
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, df_individual[grupo]["Farmacia"].to_string(index=False))
    else:
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, "No matching .txt files found for the selected date.")
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, "No matching .txt files found for the selected date.")


## REVISAR Y ARREGLAR
def load_horas_PBI(path, grupo, date_entry):
    print(path, grupo, date_entry)
    selected_date = date_entry.get_date().strftime('%d-%m-%y')
    folder_path = os.path.join(path, selected_date)
    print(folder_path)
    # Define the folder and file path
    excel_file = None
    df_out = pd.DataFrame()

    # Find the Excel file in the folder
    for filename in os.listdir(folder_path):
        print(filename)
        if filename.endswith(".xlsx"):
            excel_file = os.path.join(folder_path, filename)
            print(filename)

            # AGREGAR CONCATENACIÓN
            xl = pd.ExcelFile(excel_file)

            sheet_name = xl.sheet_names[0]
            # sheet_to_load = sheet_names.get()

            if sheet_name:
                sheet_df = xl.parse(sheet_name)
                sheet_df = sheet_df.dropna(subset=['Nomina'])
                sheet_df = sheet_df.drop(['Descripción', 'Puesto'], axis=1)
                sheet_df["Monto"] = 0.0
                sheet_df.columns = headers
                sheet_df['NOMINA'] = "0001"
                print(sheet_df)
                df_out = pd.concat([df_out, sheet_df])

    df_individual[grupo]["Horas PBI"] = df_out
    result_textbox.delete(1.0, ctk.END)
    result_textbox.insert(ctk.END, df_individual[grupo]["Horas PBI"].to_string(index=False))
    result_textbox2.delete(1.0, ctk.END)
    result_textbox2.insert(ctk.END, df_individual[grupo]["Horas PBI"].to_string(index=False))


def load_horas_CAC(path, grupo, date_entry):
    print(path, grupo, date_entry)
    selected_date = date_entry.get_date().strftime('%d-%m-%y')
    folder_path = os.path.join(path, selected_date)
    print(folder_path)
    # Define the folder and file path
    excel_file = None
    df_out = pd.DataFrame()

    # Find the Excel file in the folder
    for filename in os.listdir(folder_path):
        print(filename)
        if filename.endswith(".xlsx"):
            excel_file = os.path.join(folder_path, filename)
            print(filename)

            # AGREGAR CONCATENACIÓN
            xl = pd.ExcelFile(excel_file)
            if grupo == "Satori":
                sheet_df = xl.parse("data Cant.")
            else:
                sheet_df = xl.parse("data CAC")
            # Split DataFrame into chunks of 3 columns
            chunk_column_names = ["EMPLEADO", "CONCEPTO", "CANTIDAD"]
            # Split DataFrame into chunks of 3 columns and assign names to columns
            chunks = [sheet_df.iloc[:, i:i + 3].reset_index(drop=True) for i in range(0, sheet_df.shape[1], 3)]

            # Remove column names by setting them to None or empty string
            for chunk in chunks:
                chunk.columns = chunk_column_names  # or chunk.columns = [''] * len(chunk.columns)

            print(chunks)

            # Concatenate chunks into a single DataFrame
            sheet_df = pd.concat(chunks, ignore_index=True)
            sheet_df["MONTO"] = 0.0
            sheet_df["NOMINA"] = "0001"

            sheet_df = sheet_df[headers]
            # sheet_df.columns = headers
            print(sheet_df)
            df_out = sheet_df

    df_individual[grupo]["Horas CAC"] = df_out
    result_textbox.delete(1.0, ctk.END)
    result_textbox.insert(ctk.END, df_individual[grupo]["Horas CAC"].to_string(index=False))
    result_textbox2.delete(1.0, ctk.END)
    result_textbox2.insert(ctk.END, df_individual[grupo]["Horas CAC"].to_string(index=False))


def load_incentivos(path, grupo, date_entry):
    print(path, grupo, date_entry)
    selected_date = date_entry.get_date().strftime('%d-%m-%y')
    folder_path = os.path.join(path, selected_date)
    print(folder_path)
    # Define the folder and file path
    excel_file = None
    df_out = pd.DataFrame()

    # Find the Excel file in the folder
    for filename in os.listdir(folder_path):
        print(filename)
        if filename.endswith(".xlsx"):
            excel_file = os.path.join(folder_path, filename)
            print(filename)

            # AGREGAR CONCATENACIÓN
            xl = pd.ExcelFile(excel_file)

            sheet_names = xl.sheet_names
            # sheet_to_load = sheet_names.get()

            if "Archivo de carga" in sheet_names or "Archivo de carga " in sheet_names:  ## INCENTIVO DE REPARTO
                sheet_df = xl.parse("Archivo de carga")

                sheet_df.columns = sheet_df.iloc[0]
                sheet_df = sheet_df[1:].reset_index(drop=True)
                sheet_df = sheet_df.dropna(subset=['Empleado'])

                sheet_df.drop('Nombre', axis=1, inplace=True)
                sheet_df.drop('Descripción', axis=1, inplace=True)
                sheet_df.drop('Puesto', axis=1, inplace=True)
                sheet_df['CANTIDAD'] = 0.0

                # Rename columns to only show the first word
                sheet_df.columns = [get_first_word(col) for col in sheet_df.columns]

                sheet_df = sheet_df[["Empleado", "Nomina", "Concepto", "CANTIDAD", "Incentivo"]]

                sheet_df.columns = headers
                # sheet_df.columns = headers
                print(sheet_df)
                df_out = pd.concat([df_out, sheet_df])
            elif "Desc. horas" in sheet_names:  ## BONO GERENCIAL
                sheet_df = xl.parse("Desc. horas")

                sheet_df.columns = sheet_df.iloc[0]
                sheet_df = sheet_df[1:].reset_index(drop=True)
                sheet_df = sheet_df.dropna(subset=['Empleado'])
                sheet_df = sheet_df.dropna(axis=1, how='all')

                sheet_df.drop('Puesto', axis=1, inplace=True)
                sheet_df.drop('Nombre', axis=1, inplace=True)
                sheet_df.drop('Descripción', axis=1, inplace=True)

                sheet_df.columns = [get_first_word(col) for col in sheet_df.columns]
                sheet_df["CANTIDAD"] = 0.0
                sheet_df['Valor'] = sheet_df['Valor'].fillna(0)

                sheet_df = sheet_df[["Empleado", "Nomina", "Concepto", "CANTIDAD", "Valor"]]

                sheet_df.columns = headers
                sheet_df["NOMINA"] = "0001"

                print(sheet_df)
                df_out = pd.concat([df_out, sheet_df])
            elif "Summary" in sheet_names:  ## INCENTIVO CAC
                sheet_df = xl.parse("Summary")

                sheet_df = sheet_df.dropna(subset=['Nombre'])
                sheet_df = sheet_df[sheet_df['ID#'] != "ID"]

                sheet_df.drop('Nombre', axis=1, inplace=True)
                sheet_df.drop('Supervisor', axis=1, inplace=True)
                sheet_df["CANTIDAD"] = 0.0
                sheet_df["NOMINA"] = "0001"
                sheet_df["CONCEPTO"] = ConceptoIncentivosCAC

                sheet_df = sheet_df[["ID#", "NOMINA", "CONCEPTO", "CANTIDAD", "Payout"]]
                sheet_df.columns = headers

                print(sheet_df)
                df_out = pd.concat([df_out, sheet_df])

    df_individual[grupo]["Incentivos"] = df_out
    result_textbox.delete(1.0, ctk.END)
    result_textbox.insert(ctk.END, df_individual[grupo]["Incentivos"].to_string(index=False))
    result_textbox2.delete(1.0, ctk.END)
    result_textbox2.insert(ctk.END, df_individual[grupo]["Incentivos"].to_string(index=False))


###### REVISAR CARGA, CONCATENACIÓN Y ETL DE ARCHIVOS.
'''
def load_licenciasMed(path, grupo, date_entry):
    print(path, grupo, date_entry)
    selected_date = date_entry.get_date().strftime('%d-%m-%y')
    folder_path = path ## os.path.join(path, selected_date) 
    print(folder_path)
    # Define the folder and file path
    excel_file = None
    df_out = pd.DataFrame()

    # Find the Excel file in the folder
    for filename in os.listdir(folder_path):
        print(filename)
        if filename.endswith(".xlsx"):
            excel_file = os.path.join(folder_path, filename)
            print(filename)

            # AGREGAR CONCATENACIÓN
            xl = pd.ExcelFile(excel_file)

            sheet_name = xl.sheet_names[0]
            # sheet_to_load = sheet_names.get()

            if sheet_name:
                sheet_df = xl.parse(sheet_name)
                sheet_df = sheet_df.dropna(subset=['Nomina'])
                sheet_df = sheet_df.drop(['Descripción', 'Puesto'], axis=1)
                sheet_df["Monto"] = 0.0
                sheet_df.columns = headers
                print(sheet_df)
                df_out = pd.concat([df_out, sheet_df])

    df_individual[grupo]["Horas PBI"] = df_out
    result_textbox.delete(1.0, ctk.END)
    result_textbox.insert(ctk.END, df_individual[grupo]["Horas PBI"].to_string(index=False))
    result_textbox2.delete(1.0, ctk.END)
    result_textbox2.insert(ctk.END, df_individual[grupo]["Horas PBI"].to_string(index=False))
    '''


def generate_nomina(grupo):
    # global NominaDF
    df_out[grupo] = pd.DataFrame(pd.DataFrame(columns=headers))
    if isinstance(df_individual[grupo]["Cooperativa"], pd.DataFrame):
        # NominaDF = []
        # df_out[grupo].columns = headers

        print(VarState[grupo].keys)
        print(grupo)
        print(df_out[grupo])
        for key in list(VarState[grupo].keys()):
            print(key)
            if VarState[grupo][key] == True:
                print(key)
                df_out[grupo] = pd.concat([df_out[grupo], df_individual[grupo][key]])

        # Display the NominaDF DataFrame in the text box
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, df_out[grupo].to_string(index=False))
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, df_out[grupo].to_string(index=False))
    else:
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, "Please load both Cooperativa and Propina data first.")
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, "Please load both Cooperativa and Propina data first.")


def save_to_excel(grupo):
    if isinstance(df_out[grupo], pd.DataFrame):
        # Ask the user for the file name
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df_out[grupo]['EMPLEADO'] = df_out[grupo]['EMPLEADO'].astype(int)
                df_out[grupo]['EMPLEADO'] = df_out[grupo]['EMPLEADO'].apply(lambda x: f"{x:05}")
                df_out[grupo]['CANTIDAD'] = df_out[grupo]['CANTIDAD'].round(2)
                df_out[grupo]['MONTO'] = df_out[grupo]['MONTO'].round(2)
                df_out[grupo] = df_out[grupo][~((df_out[grupo]['MONTO'] == 0) & (df_out[grupo]['CANTIDAD'] == 0))]

                sheet_name = 'Nómina'
                df_out[grupo].to_excel(writer, sheet_name=sheet_name, index=False)

                # Access the XlsxWriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                # Apply number format with two decimal places
                format = workbook.add_format({'num_format': '0.00'})

                # Assuming 'MONTO' is the third column (adjust the range as needed)
                worksheet.set_column('D:E', None, format)  # C:C means third column

                writer.save()

            result_textbox.delete(1.0, ctk.END)
            result_textbox.insert(ctk.END, f"Excel file saved successfully at {save_path}")
            result_textbox2.delete(1.0, ctk.END)
            result_textbox2.insert(ctk.END, f"Excel file saved successfully at {save_path}")
    else:
        result_textbox.delete(1.0, ctk.END)
        result_textbox.insert(ctk.END, "Please generate Nómina first.")
        result_textbox2.delete(1.0, ctk.END)
        result_textbox2.insert(ctk.END, "Please generate Nómina first.")


########## SECCIÓN DE CHECKBOXES. Aquí se tiene una función reciclada que calculará el estado de varios checkboxes.
########## CAMBIAR TODAS ESTAS VARIABLES POR checkbox_var[grupo][concepto] e iterar
if True:
    # Function to handle tooltip display
    def create_tooltip(widget, text):
        tooltip = ctk.CTkToplevel(root)
        tooltip.overrideredirect(True)  # Remove window decorations
        tooltip.wm_withdraw()  # Hide the tooltip initially

        label = ctk.CTkLabel(tooltip, text=text, fg_color="transparent", corner_radius=0)
        label.pack(ipadx=0, ipady=0, padx=0, pady=0)  # Ensure no padding

        def on_enter(event):
            # Position the tooltip near the mouse cursor
            tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            tooltip.deiconify()  # Show the tooltip

        def on_leave(event):
            tooltip.withdraw()  # Hide the tooltip

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)


    # Function to handle the checkbox state changes
    def update_checkbox_value(checkbox_var1, grupo, campo):
        checkbox_value = checkbox_var1.get()  # 0 for unchecked, 1 for checked
        print(checkbox_value)
        if checkbox_value == 1:
            checkbox_state = True
        else:
            checkbox_state = False
            # print("feik")
        VarState[grupo][campo] = checkbox_state

        # print(f"Checkbox state: {checkbox_state}")


    checkbox_var = {"Satori": {}, "Hispizza": {}}
    checkbox_var["Satori"]["Cooperativa"] = ctk.IntVar(value=1)
    checkbox_var["Satori"]["Propina Voluntaria"] = ctk.IntVar(value=1)
    checkbox_var["Satori"]["Horas consolidado"] = ctk.IntVar(value=1)
    checkbox_var["Satori"]["Desc. Gym"] = ctk.IntVar(value=0)
    checkbox_var["Hispizza"]["Cooperativa"] = ctk.IntVar(value=1)
    checkbox_var["Hispizza"]["Fripick"] = ctk.IntVar(value=1)
    checkbox_var["Hispizza"]["Farmacia"] = ctk.IntVar(value=1)
    checkbox_var["Hispizza"]["Horas CAC"] = ctk.IntVar(value=1)
    checkbox_var["Hispizza"]["Horas PBI"] = ctk.IntVar(value=1)
    checkbox_var["Hispizza"]["Incentivos"] = ctk.IntVar(value=1)
    checkbox_var["Hispizza"]["Licencias Med"] = ctk.IntVar(value=0)
    for group in checkbox_var:  # Iterate through the outer dictionary keys
        for concept in checkbox_var[group]:  # Iterate through the inner dictionary keys
            # print(f"Group: {group}, Concept: {concept}, Value: {checkbox_var[group][concept].get()}")
            update_checkbox_value(checkbox_var[group][concept], group, concept)

    if True:
        print("")
        # checkbox_var1 = ctk.IntVar(value=1)
        # update_checkbox_value(checkbox_var1, "Satori", "Cooperativa")
        # checkbox_var2 = ctk.IntVar(value=1)
        # update_checkbox_value(checkbox_var2, "Satori", "Propina Voluntaria")
        # # checkbox_var3 = ctk.IntVar(value=1)
        # # update_checkbox_value(checkbox_var3, "Satori", "Desc. Gym")

        # checkbox_var21 = ctk.IntVar(value=1)
        # update_checkbox_value(checkbox_var21, "Hispizza", "Cooperativa")
        # checkbox_var22 = ctk.IntVar(value=1)
        # update_checkbox_value(checkbox_var22, "Hispizza", "Fripick")

        # Add checkboxes to the window, passing a unique identifier and state variable to the function
        # checkbox1 = ctk.CTkCheckBox(tab_satori, text="", variable=checkbox_var1,
        #                             command=lambda: update_checkbox_value(checkbox_var1, "Satori", "Cooperativa"))
        # checkbox1.grid(row=1, column=0, padx=10, pady=10, sticky="e")

        # checkbox2 = ctk.CTkCheckBox(tab_satori, text="", variable=checkbox_var2,
        #                             command=lambda: update_checkbox_value(checkbox_var2, "Satori", "Propina Voluntaria"))
        # checkbox2.grid(row=2, column=0, padx=10, pady=10, sticky="e")

    checkbox = {"Satori": {}, "Hispizza": {}}
    checkbox["Satori"]["Cooperativa"] = ctk.IntVar(value=1)
    checkbox["Satori"]["Propina Voluntaria"] = ctk.IntVar(value=1)
    checkbox["Satori"]["Horas consolidado"] = ctk.IntVar(value=1)
    checkbox["Hispizza"]["Cooperativa"] = ctk.IntVar(value=1)
    checkbox["Hispizza"]["Fripick"] = ctk.IntVar(value=1)
    checkbox["Hispizza"]["Farmacia"] = ctk.IntVar(value=1)
    checkbox["Hispizza"]["Horas CAC"] = ctk.IntVar(value=1)
    checkbox["Hispizza"]["Horas PBI"] = ctk.IntVar(value=1)
    checkbox["Hispizza"]["Incentivos"] = ctk.IntVar(value=1)
    # checkbox["Hispizza"]["Licencias Med"] = ctk.IntVar(value=0)

    for group in checkbox:  # Iterate through the outer dictionary keys
        i = 1
        for concept in checkbox[group]:  # Iterate through the inner dictionary keys
            # print(f"Group: {group}, Concept: {concept}, Value: {checkbox_var[group][concept].get()}")

            # Capture group and concept in the lambda using default argument trick
            checkbox[group][concept] = ctk.CTkCheckBox(
                tab[group],
                text="",
                variable=checkbox_var[group][concept],
                command=lambda g=group, c=concept: update_checkbox_value(checkbox_var[g][c], g, c)
            )
            checkbox[group][concept].grid(row=i, column=0, padx=10, pady=10, sticky="e")

            i += 1
            # print(checkbox[group][concept])

    # Apply tooltips to the checkboxes
    create_tooltip(checkbox["Satori"]["Cooperativa"], "Incluir Cooperativa")
    create_tooltip(checkbox["Satori"]["Propina Voluntaria"], "Incluir propina voluntaria")
    create_tooltip(checkbox["Satori"]["Horas consolidado"], "Incluir horas")

    # Apply tooltips to the checkboxes
    create_tooltip(checkbox["Hispizza"]["Cooperativa"], "Incluir Cooperativa")
    create_tooltip(checkbox["Hispizza"]["Fripick"], "Incluir Fripick")
    create_tooltip(checkbox["Hispizza"]["Farmacia"], "Incluir descuento de farmacia")
    create_tooltip(checkbox["Hispizza"]["Horas CAC"], "Incluir horas CAC")
    create_tooltip(checkbox["Hispizza"]["Horas PBI"], "Incluir horas PBI")
    create_tooltip(checkbox["Hispizza"]["Incentivos"], "Incluir incentivos")
    # create_tooltip(checkbox["Hispizza"]["Licencias Med"], "Incluir licencias médicas")

    ########## FIN DE SECCIÓN DE CHECKBOXES.


def run_etl_and_set_button_color(button, etl_function, *args):
    try:
        # Ejecuta la función ETL con los argumentos proporcionados
        result = etl_function(*args)

        # Verificación adicional de éxito, si la función retorna un resultado
        if isinstance(result, pd.DataFrame):
            if result.empty:
                raise ValueError("El DataFrame resultante está vacío")

        # Si la transformación fue exitosa, cambiar el color a verde
        button.configure(fg_color="green", hover_color="darkgreen")
    except Exception as e:
        # Si hubo algún error, mostrar en rojo y capturar el error
        button.configure(fg_color="red", hover_color="darkred")
        print(f"Error: {e}")


########## SATORI TAB ##########

# Add the DateEntry widget for date selection in the Satori tab
date_entry_satori = DateEntry(tab["Satori"], date_pattern='dd-mm-yy', width=12, background='darkblue',
                              foreground='white', borderwidth=2)
date_entry_satori.grid(row=0, column=2, pady=10)

# Create a button to each ETL in Satori

# Para el botón de cargar archivos de cooperativa en Satori
import_button_satori = ctk.CTkButton(
    tab["Satori"],
    text="Import Coop Files",
    command=lambda: run_etl_and_set_button_color(import_button_satori, load_and_concatenate_tables, "./Satori/Cooperativa", "Satori", date_entry_satori)
)
import_button_satori.grid(row=1, column=1, pady=10)

# Para el botón de cargar Propina Voluntaria en Satori
load_excel_button = ctk.CTkButton(
    tab["Satori"],
    text="Propina Voluntaria",
    command=lambda: run_etl_and_set_button_color(load_excel_button, load_excel_by_date, date_entry_satori, "Satori")
)
load_excel_button.grid(row=2, column=1, pady=10)

# Para el botón de cargar Horas consolidado en Satori
horasCAC_satori = ctk.CTkButton(
    tab["Satori"],
    text="Horas consolidado",
    command=lambda: run_etl_and_set_button_color(horasCAC_satori, load_horas_CAC, "./Satori/Horas", "Satori", date_entry_satori)
)
horasCAC_satori.grid(row=3, column=1, pady=10)

generate_nomina_button = ctk.CTkButton(tab["Satori"], text="Generar Nómina", command=lambda: generate_nomina("Satori"))
generate_nomina_button.grid(row=4, column=1, pady=10)

save_excel_button = ctk.CTkButton(tab["Satori"], text="Save to Excel", command=lambda: save_to_excel("Satori"))
save_excel_button.grid(row=5, column=1, pady=10)

result_textbox = ctk.CTkTextbox(tab["Satori"], height=20, width=80)
result_textbox.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Create a Treeview widget
# treeview = ttk.Treeview(tab["Satori"], height=6, show="tree")

########## HISPIZZA TAB ##########

# Add the DateEntry widget for date selection in the Hispizza tab
date_entry_hispizza = DateEntry(tab["Hispizza"], date_pattern='dd-mm-yy', width=12, background='darkblue',
                                foreground='white', borderwidth=2)
date_entry_hispizza.grid(row=0, column=2, pady=10)

# Create a button to each ETL in Hispizza
# Para el botón de cargar archivos de cooperativa en Hispizza
import_button_hispizza = ctk.CTkButton(
    tab["Hispizza"],
    text="Import Coop Files",
    command=lambda: run_etl_and_set_button_color(import_button_hispizza, load_and_concatenate_tables, "./Hispizza/Cooperativa", "Hispizza", date_entry_hispizza)
)
import_button_hispizza.grid(row=1, column=1, pady=10)

# Para el botón de cargar Fripick en Hispizza
fripick_hispizza = ctk.CTkButton(
    tab["Hispizza"],
    text="Load Fripick",
    command=lambda: run_etl_and_set_button_color(fripick_hispizza, load_fripick, "./Hispizza/Fripick", "Hispizza", date_entry_hispizza)
)
fripick_hispizza.grid(row=2, column=1, pady=10)

# Para el botón de cargar Farmacia en Hispizza
farmacia_hispizza = ctk.CTkButton(
    tab["Hispizza"],
    text="Farmacia",
    command=lambda: run_etl_and_set_button_color(farmacia_hispizza, load_farmacia, "./Hispizza/Desc. Farmacia", "Hispizza", date_entry_hispizza)
)
farmacia_hispizza.grid(row=3, column=1, pady=10)

# Para el botón de cargar Horas PBI en Hispizza
horasPBI_hispizza = ctk.CTkButton(
    tab["Hispizza"],
    text="Horas PBI",
    command=lambda: run_etl_and_set_button_color(horasPBI_hispizza, load_horas_PBI, "./Hispizza/Horas PBI", "Hispizza", date_entry_hispizza)
)
horasPBI_hispizza.grid(row=5, column=1, pady=10)

# Para el botón de cargar Horas CAC en Hispizza
horasCAC_hispizza = ctk.CTkButton(
    tab["Hispizza"],
    text="Horas CAC",
    command=lambda: run_etl_and_set_button_color(horasCAC_hispizza, load_horas_CAC, "./Hispizza/Horas CAC", "Hispizza", date_entry_hispizza)
)
horasCAC_hispizza.grid(row=4, column=1, pady=10)

# Para el botón de cargar Incentivos en Hispizza
incentivos_hispizza = ctk.CTkButton(
    tab["Hispizza"],
    text="Incentivos",
    command=lambda: run_etl_and_set_button_color(incentivos_hispizza, load_incentivos, "./Hispizza/Incentivos", "Hispizza", date_entry_hispizza)
)
incentivos_hispizza.grid(row=6, column=1, pady=10)

generate_nomina_button = ctk.CTkButton(tab["Hispizza"], text="Generar Nómina",
                                       command=lambda: generate_nomina("Hispizza"))
generate_nomina_button.grid(row=7, column=1, pady=10)

save_excel_button = ctk.CTkButton(tab["Hispizza"], text="Save to Excel", command=lambda: save_to_excel("Hispizza"))
save_excel_button.grid(row=8, column=1, pady=10)

# Licencias_hispizza = ctk.CTkButton(tab["Hispizza"], text="Licencia files",
#                                      command=lambda: load_licenciasMed("./Hispizza/Licencias Medicas", "Hispizza", date_entry_hispizza))
# Licencias_hispizza.grid(row=5, column=1, pady=10)

# Create a text widget to display the result
result_textbox2 = ctk.CTkTextbox(tab["Hispizza"], height=20, width=80)
result_textbox2.grid(row=9, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

########## LAYOUT THE LOS WIDGETS ##########
tab["Satori"].grid_rowconfigure(6, weight=1)
tab["Satori"].grid_columnconfigure(0, weight=1)
tab["Satori"].grid_columnconfigure(1, weight=1)
tab["Satori"].grid_columnconfigure(2, weight=1)

tab["Hispizza"].grid_rowconfigure(9, weight=1)
tab["Hispizza"].grid_columnconfigure(0, weight=1)
tab["Hispizza"].grid_columnconfigure(1, weight=1)
tab["Hispizza"].grid_columnconfigure(2, weight=1)


# Function to synchronize dates
def sync_dates(source, target):
    target.set_date(source.get_date())


# Bind the date change event
date_entry_satori.bind("<<DateEntrySelected>>", lambda event: sync_dates(date_entry_satori, date_entry_hispizza))
date_entry_hispizza.bind("<<DateEntrySelected>>", lambda event: sync_dates(date_entry_hispizza, date_entry_satori))

# Run the customtkinter main loop
root.mainloop()

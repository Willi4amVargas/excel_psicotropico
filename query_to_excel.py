import pandas as pd
import time
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
import os

def replace_dot_with_comma(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('.', ',')
    return df

def query_to_excel(conn, query, excel_file, sheet_name):
    try:
        print(f"Creando Archivo en la hoja {sheet_name}")
        df = pd.read_sql_query(query, conn)

        columns_to_replace = ['total_coin_dls', 'total_coin_bs']
        df = replace_dot_with_comma(df, columns_to_replace)

        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)


        libro = load_workbook(excel_file)
        hoja = libro[sheet_name]

        tab_range = f"A1:{chr(64 + len(df.columns))}{len(df) + 1}"

        tab = Table(displayName="Tabla1", ref=tab_range)

        style = TableStyleInfo(
            name="TableStyleMedium9", 
            showFirstColumn=False,
            showLastColumn=False, 
            showRowStripes=True, 
            showColumnStripes=True
        )
        tab.tableStyleInfo = style

        hoja.add_table(tab)

        if 'total_amount' in df.columns:
            for cell in hoja['C'][1:]:  
                cell.alignment = Alignment(horizontal='center')


        for column in hoja.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass

            hoja.column_dimensions[column_letter].width = max_length + 2

        libro.save(excel_file)
        print(f"Resultados guardados en {excel_file}")
        time.sleep(2)
    except Exception as e:
        print(f"Error al ejecutar la consulta y guardar en Excel: {e}")
        time.sleep(2)

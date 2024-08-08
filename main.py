from db import connect_to_db
from query_to_excel import query_to_excel
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
import time
import flet as ft
import threading

def execute_query_and_save(conn, query, excel_file, sheet_name):
    query_to_excel(conn, query, excel_file, sheet_name)


def main(page: ft.Page):
    # Configura el tamaño de la ventana
    width = 500  # Ancho deseado
    height = 200  # Alto deseado

    # Establece el tamaño de la ventana
    page.window_width = width
    page.window_height = height

    # Centro de la pantalla
    page.window.left = (1440 - width) / 2
    page.window.top = (900 - height) / 2

    page.title = "Generando Excel..."
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER

    # Mostrar imagen de carga
    container = ft.Column(
        controls=[
            ft.Text("Generando archivo Excel, por favor espere..."),
            ft.Image(src="loading.gif", width=100, height=100)
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=10
    )
    page.add(container)

    def run_task():
      desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
      folder_name = "PSICOTROPICOS"
      folder_path = os.path.join(desktop, folder_name)

      try:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
      # relacion_de_ventas_controlados_07_2024
      # relacion_de_compras_controlados_07_2024
      # Obtener la fecha actual
        fecha_actual = datetime.now()

        # Restar un mes a la fecha actual
        fecha_mes_anterior = fecha_actual - relativedelta(months=1)

        # Formatear la fecha si es necesario
        fecha_formateada = fecha_mes_anterior.strftime("%Y-%m")

        excel_file = f"relacion_de_ventas_controlados_{fecha_formateada}.xlsx"
        file_path = os.path.join(folder_path, excel_file)

        excel_file1 = f"relacion_de_compras_controlados_{fecha_formateada}.xlsx"
        file_path1 = os.path.join(folder_path, excel_file1)



        query1 = """
      SELECT 
      sales_operation_details.code_product as "Codigo",
      products.description as "Descripcion",
      SUM(CASE WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.amount ELSE -sales_operation_details.amount END) as "Monto Total",
      ROUND(SUM(CASE 
            WHEN sales_operation_details.coin_code = '01' THEN 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total 
                ELSE -sales_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total / sales_operation_coins.buy_aliquot 
                ELSE -sales_operation_details.total / sales_operation_coins.buy_aliquot 
              END 
          END)::DECIMAL,2) as "Total $",
      ROUND(SUM(CASE 
            WHEN sales_operation_details.coin_code = '02' THEN 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total 
                ELSE -sales_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total * sales_operation_coins.buy_aliquot 
                ELSE -sales_operation_details.total * sales_operation_coins.buy_aliquot 
              END 
          END)::DECIMAL,2) as "Total Bs"
    FROM 
      sales_operation_details
    INNER JOIN 
      sales_operation ON sales_operation.correlative = sales_operation_details.main_correlative
    INNER JOIN 
      products ON products.code = sales_operation_details.code_product
    INNER JOIN
      sales_operation_coins on sales_operation_coins.main_correlative = sales_operation_details.main_correlative
    WHERE 
    (sales_operation.emission_date >= (DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '1 month')
    AND 
      sales_operation.emission_date < DATE_TRUNC('month', CURRENT_DATE))
    AND
      (sales_operation.operation_type = 'BILL' OR sales_operation.operation_type = 'CREDITNOTE')
    AND
      sales_operation_coins.buy_aliquot != 1
    AND
      products.department='15'
    GROUP BY 
      sales_operation_details.code_product,
      products.description
    HAVING 
      SUM(CASE WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.amount ELSE -sales_operation_details.amount END) != 0 OR
      SUM(CASE 
            WHEN sales_operation_details.coin_code = '01' THEN 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total 
                ELSE -sales_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total / sales_operation_coins.buy_aliquot 
                ELSE -sales_operation_details.total / sales_operation_coins.buy_aliquot 
              END 
          END) != 0 OR
      SUM(CASE 
            WHEN sales_operation_details.coin_code = '02' THEN 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total 
                ELSE -sales_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.total * sales_operation_coins.buy_aliquot 
                ELSE -sales_operation_details.total * sales_operation_coins.buy_aliquot 
              END 
          END) != 0
    ORDER BY 
      products.description ASC;

        """
        query2="""
        SELECT 
      shopping_operation_details.code_product as "Codigo",
      products.description as "Descripcion",
      SUM(CASE WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.amount ELSE -shopping_operation_details.amount END) as "Monto Total",
      shopping_operation.document_no as "Numero de Documento",
      shopping_operation.emission_date as "Fecha de Emisión",
      shopping_operation.provider_name as "Proveedor",
      ROUND(SUM(CASE 
            WHEN shopping_operation_details.coin_code = '01' THEN 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total 
                ELSE -shopping_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total / shopping_operation_coins.buy_aliquot 
                ELSE -shopping_operation_details.total / shopping_operation_coins.buy_aliquot 
              END 
          END)::DECIMAL,2) as "Total $",
      ROUND(SUM(CASE 
            WHEN shopping_operation_details.coin_code = '02' THEN 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total 
                ELSE -shopping_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total * shopping_operation_coins.buy_aliquot 
                ELSE -shopping_operation_details.total * shopping_operation_coins.buy_aliquot 
              END 
          END)::DECIMAL,2) as "Total Bs"
    FROM 
      shopping_operation_details
    INNER JOIN 
      shopping_operation ON shopping_operation.correlative = shopping_operation_details.main_correlative
    INNER JOIN 
      products ON products.code = shopping_operation_details.code_product
    INNER JOIN
      shopping_operation_coins on shopping_operation_coins.main_correlative = shopping_operation_details.main_correlative
    WHERE 
    (shopping_operation.emission_date >= (DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '1 month')
    AND 
      shopping_operation.emission_date < DATE_TRUNC('month', CURRENT_DATE))
    AND
      (shopping_operation.operation_type = 'BILL' OR shopping_operation.operation_type = 'CREDITNOTE')
    AND
      shopping_operation_coins.buy_aliquot != 1
    AND
      products.department='15'
    GROUP BY 
      shopping_operation_details.code_product,
      products.description,
      shopping_operation.provider_name,
      shopping_operation.document_no,
      shopping_operation.emission_date
    HAVING 
      SUM(CASE WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.amount ELSE -shopping_operation_details.amount END) != 0 OR
      SUM(CASE 
            WHEN shopping_operation_details.coin_code = '01' THEN 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total 
                ELSE -shopping_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total / shopping_operation_coins.buy_aliquot 
                ELSE -shopping_operation_details.total / shopping_operation_coins.buy_aliquot 
              END 
          END) != 0 OR
      SUM(CASE 
            WHEN shopping_operation_details.coin_code = '02' THEN 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total 
                ELSE -shopping_operation_details.total 
              END
            ELSE 
              CASE 
                WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.total * shopping_operation_coins.buy_aliquot 
                ELSE -shopping_operation_details.total * shopping_operation_coins.buy_aliquot 
              END 
          END) != 0
    ORDER BY 
      products.description ASC;
        """

        conn = connect_to_db()

        if conn:
            execute_query_and_save(conn, query1, file_path, "Controlados")
            execute_query_and_save(conn, query2, file_path1, "Controlados")
            conn.close()
        else:
            print("No se pudo establecer la conexión a la base de datos. Verifica los parámetros de conexión.")
      except Exception as e:
         print(F"Error: {e}")
      finally:
        page.window.close()
    threading.Thread(target=run_task).start()

if __name__ == "__main__":
    ft.app(target=main)

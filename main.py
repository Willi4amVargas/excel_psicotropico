from db import connect_to_db
from query_to_excel import query_to_excel
from datetime import datetime
import os

def execute_query_and_save(conn, query, excel_file, sheet_name):
    query_to_excel(conn, query, excel_file, sheet_name)

def main():

    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    folder_name = "PSICOTROPICOS"
    folder_path = os.path.join(desktop, folder_name)

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    excel_file = "VENTAS.xlsx"
    file_path = os.path.join(folder_path, excel_file)

    excel_file1 = "PROVEEDORES.xlsx"
    file_path1 = os.path.join(folder_path, excel_file1)



    query1 = """
   SELECT 
  sales_operation_details.code_product,
  products.description,
  SUM(CASE WHEN sales_operation.operation_type = 'BILL' THEN sales_operation_details.amount ELSE -sales_operation_details.amount END) as total_amount,
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
      END)::DECIMAL,2) as total_coin_DLS,
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
      END)::DECIMAL,2) as total_coin_BS
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
  shopping_operation_details.code_product,
  products.description,
  shopping_operation.document_no,
  shopping_operation.provider_name,
  SUM(CASE WHEN shopping_operation.operation_type = 'BILL' THEN shopping_operation_details.amount ELSE -shopping_operation_details.amount END) as total_amount,
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
      END)::DECIMAL,2) as total_coin_DLS,
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
      END)::DECIMAL,2) as total_coin_BS
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
  shopping_operation.document_no
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
        execute_query_and_save(conn, query1, file_path, "Psicotropicos")
        execute_query_and_save(conn, query2, file_path1, "Psicotropicos")
        conn.close()
    else:
        print("No se pudo establecer la conexión a la base de datos. Verifica los parámetros de conexión.")

if __name__ == "__main__":
    main()

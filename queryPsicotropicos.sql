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
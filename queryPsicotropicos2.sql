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
# Inventory Analysis

This code performs an analysis of an inventory stored in an Excel file. It uses the `openpyxl` library to read and manipulate the Excel data.

## Code Explanation

1. Import the `openpyxl` library to work with Excel files.

   ```python
   import  openpyxl
   ```

2. Load the inventory file using `openpyxl.load_workbook()` and assign it to the `inv_file` variable.

   ```python
   inv_file = openpyxl.load_workbook("inventory.xlsx")
   ```

3. Get the "Sheet1" from the inventory file and assign it to the `product_list` variable.

   ```python
   product_list = inv_file["Sheet1"]
   ```

4. Initialize three dictionaries to store calculated data:

   - `product_per_supplier`: Number of products per supplier.
   - `total_value_per_supplier`: Total value of inventory per supplier.
   - `products_under_10_inv`: Products with inventory less than 10.

   ```python
   products_per_supplier = {}
   total_value_per_supplier = {}
   products_under_10_inv = {}
   ```

5. Iterate over each row of the product list starting from the second row (skipping the header).

   ```python
   for product_row in range(2, product_list.max_row + 1)
   ```

6. Extract relevant data from cells in the current row, such as supplier name, inventory count, price per unit, product number, and the cell for inventory price.

   ```python
   supplier_name = product_list.cell(product_row, 4).value
   inventory = product_list.cell(product_row, 2).value
   price = product_list.cell(product_row, 3).value
   product_num = product_list.cell(product_row, 1).value
   inventory_price = product_list.cell(product_row, 5)
   ```

7. Calculate the number of products per supplier by updating the `products_per_supplier` dictionary. If the supplier name is already present, increment the count; otherwise, add a new entry with a count of 1.

   ```python
   if supplier_name in products_per_supplier:
     current_num_products = products_per_supplier.get(supplier_name)
     products_per_supplier[supplier_name] = current_num_products + 1
   else:
     products_per_supplier[supplier_name] = 1
   ```

8. Calculate the total value of inventory per supplier by updating the `total_value_per_supplier` dictionary. If the supplier name is already present, add the product count and price to the current total value; otherwise, calculate the value for the current supplier.

   ```python
   if supplier_name in total_value_per_supplier:
     current_total_value = total_value_per_supplier.get(supplier_name)
     total_value_per_supplier[supplier_name] = inventory * price
   else:
     total_value_per_supplier[supplier_name] = inventory * price
   ```

9. Check if the inventory count is less than 10. If this, add the product number and inventory count to the `products_under_10_inv` dictionary.

   ```python
   if inventory < 10:
     products_under_10_inv[int(product_num)] = int(inventory)
   ```

10. Calculate the inventory price by multiplying the inventory count and price per unit, and update the corresponding cell in the Excel file.

    ```python
    inventory_price.value = inventory * price
    ```

11. Print the dictionaries containing the calculated data: `products_per_supplier`, `total_value_per_supplier`, and `products_under_10_inv`.

    ```python
    print(products_per_supplier)
    print(total_value_per_supplier)
    print(products_under_10_inv)
    ```

12. Save the modified.

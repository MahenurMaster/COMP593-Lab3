import os
import sys
import pandas as pd
from datetime import datetime

def validate_input_file(filepath):
    if not filepath:
        print("Error: No path to sales data CSV file provided.")
        sys.exit(1)
    if not os.path.isfile(filepath):
        print("Error: The provided path does not exist or is not a file.")
        sys.exit(1)

def create_orders_directory(base_dir):
    date_str = datetime.now().strftime("%Y-%m-%d")
    orders_dir = os.path.join(base_dir, f"Orders_{date_str}")
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)
    return orders_dir

def process_sales_data(filepath, orders_dir):
    sales_data = pd.read_csv(filepath)
    
    # Print the columns to check their names
    print("Columns in CSV file:", sales_data.columns)

    grouped_orders = sales_data.groupby('ORDER ID')
    
    for order_id, order_info in grouped_orders:
        order_info = order_info.sort_values(by='ITEM NUMBER')
        order_info['TOTAL PRICE'] = order_info['ITEM QUANTITY'] * order_info['ITEM PRICE']
        
        total_price = order_info['TOTAL PRICE'].sum()
        
        # Update the column names based on the CSV file
        order_info = order_info[['ORDER ID', 'ITEM NUMBER', 'PRODUCT LINE', 'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE']]
        
        order_filepath = os.path.join(orders_dir, f"Order_{order_id}.xlsx")
        with pd.ExcelWriter(order_filepath, engine='xlsxwriter') as writer:
            order_info.to_excel(writer, index=False, sheet_name=f'Order_{order_id}')
            
            workbook  = writer.book
            worksheet = writer.sheets[f'Order_{order_id}']
            
            money_format_variable = workbook.add_format({'num_format': '$#,##0.00'})
            worksheet.set_column('E:F', 18, money_format_variable)
            worksheet.set_column('A:D', 15)
            
            worksheet.write(len(order_info) + 1, 4, 'Grand Total')
            worksheet.write(len(order_info) + 1, 5, total_price, money_format_variable)
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 10)
            worksheet.set_column('C:C', 20)
            worksheet.set_column('D:D', 14)

def main():
    if len(sys.argv) < 2:
        print("Error: No path to sales data CSV file provided.")
        sys.exit(1)

    filepath = sys.argv[1]
    validate_input_file(filepath)
    
    base_directory = os.path.dirname(filepath)
    orders_directory = create_orders_directory(base_directory)
    
    process_sales_data(filepath, orders_directory)
    
    print(f"Excel files have been generated in {orders_directory}")

if __name__ == "__main__":
    main()

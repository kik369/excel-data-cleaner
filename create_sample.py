import csv

# Create a sample messy CSV to demonstrate the product
data = """   Name , Email ,  Phone,  Amount ,  Notes  
John  Smith,john@example.com,555-0101,100.00,Regular customer  
john  Smith,john@example.com,555-0101,150.00,Regular customer  
Jane Doe,Jane@Example.COM,555-0202,200.50,
  Bob  Jones,bob@test.com,,75.25,New  customer  
Alice  , ,555-0404,300.00,VIP  member  
,,,
Charlie Brown,charlie@mail.com,555-0606,50.75,  
"""

with open('/home/ref/.hermes/income/products/excel-data-cleaner/sample_data.csv', 'w', newline='') as f:
    f.write(data.strip())

print("Sample data created")

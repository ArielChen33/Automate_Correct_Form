import pandas as pd

# Get input
total = int(input("Please entter the total quantity: "))
unit = int(input("Please enter the quantity of each unit: "))

# Calculate how many rows
row_quantity = total // unit
if total % unit != 0: 
    row_quantity += 1


# Generate the excel file
data = []
for i in range(1, row_quantity + 1):
    if i < row_quantity or total % unit == 0: 
        quantity = unit
    else: 
        quantity = total % unit # remainder for the last row
    data.append({
        "Row": i, 
        "Quantity": quantity
    })

# Create the dataframe and export to Excel
df = pd.DataFrame(data)
output_filename = "output.xlsx"
df.to_excel(output_filename, index=False)

print(f"Excel file {output_filename} created with {row_quantity} rows")
    

# print("total: ", total)
# print("unit: ", total)
# print(f"There are going to be {row_quantity} rows in total")
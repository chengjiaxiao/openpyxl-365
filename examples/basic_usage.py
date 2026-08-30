from openpyxl365 import Excel365Workbook, inspect_excel365_dynamic_arrays

document = Excel365Workbook.create()
sheet = document.wb.active
sheet.title = "DynamicArrays"

for row, value in enumerate(("apple", "banana", "apple", "pear"), start=1):
    sheet.cell(row=row, column=1, value=value)

document.set_dynamic_formula("DynamicArrays", "C1", "=UNIQUE(A1:A4)")
output = document.save("dynamic-arrays.xlsx")

report = inspect_excel365_dynamic_arrays(output)
print(report)

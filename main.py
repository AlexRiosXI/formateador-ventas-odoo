from PyQt5.QtWidgets import QApplication, QFileDialog
import openpyxl
import sys
import re

output_file = input("Nombre del archivo de salida: ")

app = QApplication(sys.argv)
file, _ = QFileDialog.getOpenFileName(None, "Selecciona un archivo", "Archivos de datos (*.xlsx *.csv)")
if not file:
    exit()

file = openpyxl.load_workbook(file)

print("Verificando columnas...")
columns = file.active.iter_rows(min_row=1, max_row=1, values_only=True)
target_columns = ["date","orderlinesproduct", "orderlinesproductinternalreference", "orderlinesquantity", "orderlinesdiscount", "orderlinessubtotalwotax", "orderlinessubtotal", "salesteam"]
for column in columns:
    for n, col in enumerate(column):
        if type(col) == str:
            col = re.sub(r'[^a-zA-Z]', '', col).lower()
            if n >= len(target_columns):
                break
            if target_columns[n] != col:
                raise Exception(f"Columna {n+1} es {col} pero debería ser {target_columns[n]}")

print("Columnas correctas")
                
sale_index = 0
sale_date = None
sale_branch = None
rows = []
print("Creando archivo de salida")
with open(f"{output_file}.csv", "w") as f:
    f.write("NoVenta,Fecha,Sucursal,Producto,SKU,Cantidad,Descuento,SubtotalAntesImpuestos,Subtotal\n")
print("Operando filas...")
total_rows = file.active.max_row
for row_index, row in enumerate(file.active.iter_rows(min_row=2, values_only=True)):
    print(f"Procesando fila {row_index+1} de {total_rows}", end="\r")
    #Handle Sale ID
    if row[0]:
        sale_index += 1
        sale_date = row[0]
        sale_branch = row[7]
    sku = row[2]
    description = row[1]
    if sku:
        description = description.replace(f"[{sku}]", "")
    quantity = row[3]
    discount = row[4]
    subtotal_before_tax = row[5]
    subtotal = row[6]
    with open(f"{output_file}.csv", "a") as f:
        f.write(f"{sale_index},{sale_date},{sale_branch},{description},{sku},{quantity},{discount},{subtotal_before_tax},{subtotal}\n")

print(f"Archivo terminado con exito, {output_file}.csv")
    

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

ARCHIVO_DATOS = "gastos_guardados.xlsx"
registros = []
precio_dolar = 0.0

# --- Funciones ---
def validar_fecha(fecha):
    try:
        datetime.datetime.strptime(fecha, "%d/%m/%Y")
        return True
    except ValueError:
        return False

def set_dolar():
    global precio_dolar
    try:
        precio_dolar = float(entry_dolar.get())
        messagebox.showinfo("Éxito", f"Precio del dólar actualizado a Bs. {precio_dolar}")
    except ValueError:
        messagebox.showerror("Error", "Ingresa un valor numérico válido")

def guardar_datos():
    if registros:
        df = pd.DataFrame(registros)
        df.to_excel(ARCHIVO_DATOS, index=False)

def cargar_datos():
    if os.path.exists(ARCHIVO_DATOS):
        df = pd.read_excel(ARCHIVO_DATOS)
        for _, fila in df.iterrows():
            registros.append(fila.to_dict())
        actualizar_tabla()

def agregar_gasto():
    if precio_dolar == 0:
        messagebox.showwarning("Atención", "Primero ingresa el precio del dólar del día.")
        return
    # Validar campos vacíos
    campos = [entry_factura, entry_desc, entry_unidad, entry_cantidad, entry_precio_bs, entry_fecha, entry_proveedor]
    for c in campos:
        if not c.get().strip():
            messagebox.showerror("Error", "Todos los campos son obligatorios")
            return
    # Validar fecha
    if not validar_fecha(entry_fecha.get()):
        messagebox.showerror("Error", "Formato de fecha inválido. Use DD/MM/YYYY")
        return
    try:
        item = len(registros) + 1
        factura = entry_factura.get()
        desc = entry_desc.get()
        unidad = entry_unidad.get()
        cantidad = float(entry_cantidad.get())
        precio_bs = float(entry_precio_bs.get())
        precio_usd = round(precio_bs / precio_dolar, 2)
        fecha = entry_fecha.get()
        proveedor = entry_proveedor.get()
        total = round(precio_bs * cantidad, 2)

        registro = {
            "Item": item,
            "Factura": factura,
            "Descripción": desc,
            "Unidad": unidad,
            "Cantidad": cantidad,
            "Precio (Bs)": precio_bs,
            "Precio ($)": precio_usd,
            "Fecha": fecha,
            "Total": total,
            "Proveedor": proveedor
        }

        registros.append(registro)
        actualizar_tabla()
        limpiar_campos()
        guardar_datos()

    except ValueError:
        messagebox.showerror("Error", "Verifica los campos numéricos (cantidad y precio).")

def eliminar_gasto():
    selected = tabla.selection()
    if not selected:
        messagebox.showwarning("Atención", "Selecciona un registro para eliminar")
        return
    if messagebox.askyesno("Confirmar", "¿Deseas eliminar el registro seleccionado?"):
        for s in selected:
            item_val = tabla.item(s)['values'][0]  # columna Item
            registros[:] = [r for r in registros if r["Item"] != item_val]
        actualizar_tabla()
        guardar_datos()

def actualizar_tabla():
    for row in tabla.get_children():
        tabla.delete(row)
    for idx, r in enumerate(registros):
        r["Item"] = idx + 1  # actualizar numeración
        tabla.insert("", "end", values=list(r.values()))
    total_general = sum(r["Total"] for r in registros)
    label_total.config(text=f"TOTAL GENERAL: Bs. {round(total_general, 2)}")

def limpiar_campos():
    for entry in [entry_factura, entry_desc, entry_unidad, entry_cantidad, entry_precio_bs, entry_fecha, entry_proveedor]:
        entry.delete(0, tk.END)

def auto_guardar():
    guardar_datos()
    root.after(60000, auto_guardar)  # cada 60 segundos

def exportar_excel():
    if not registros:
        messagebox.showwarning("Atención", "No hay datos para exportar.")
        return

    archivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivo Excel", "*.xlsx")],
        title="Guardar reporte de gastos",
        initialfile="reporte_gastos.xlsx"
    )

    if not archivo:
        return

    # Guardar datos base
    df = pd.DataFrame(registros)
    df.to_excel(archivo, index=False)

    # --- Dar formato con openpyxl ---
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

    wb = load_workbook(archivo)
    ws = wb.active

    # Encabezado: fondo azul y texto blanco
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Bordes y alineación de todas las celdas
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Resaltar fila del total (última)
    total_row = ws.max_row
    for cell in ws[total_row]:
        cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        cell.font = Font(bold=True)

    # Ajustar ancho de columnas automáticamente
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 3

    wb.save(archivo)
    wb.close()

    messagebox.showinfo("Éxito", f"Datos exportados y formateados correctamente a:\n{archivo}")

# --- Interfaz gráfica ---
root = tk.Tk()
root.title("Control de Gastos - Terra Caliza")
root.geometry("1150x600")

# Precio dólar
frame_dolar = tk.Frame(root)
frame_dolar.pack(pady=10)
tk.Label(frame_dolar, text="Precio del dólar (Bs):").pack(side=tk.LEFT, padx=5)
entry_dolar = tk.Entry(frame_dolar, width=10)
entry_dolar.pack(side=tk.LEFT)
tk.Button(frame_dolar, text="Actualizar", command=set_dolar).pack(side=tk.LEFT, padx=5)

# Formulario
frame_form = tk.Frame(root)
frame_form.pack()
labels = ["Factura", "Descripción", "Unidad", "Cantidad", "Precio (Bs)", "Fecha (DD/MM/YYYY)", "Proveedor"]
entries = []
for i, label_text in enumerate(labels):
    tk.Label(frame_form, text=label_text).grid(row=0, column=i, padx=5, pady=5)
    entry = tk.Entry(frame_form, width=15)
    entry.grid(row=1, column=i, padx=5)
    entries.append(entry)

(entry_factura, entry_desc, entry_unidad, entry_cantidad, entry_precio_bs, entry_fecha, entry_proveedor) = entries

tk.Button(root, text="Agregar gasto", command=agregar_gasto, bg="#4CAF50", fg="white").pack(pady=10)
tk.Button(root, text="Eliminar gasto", command=eliminar_gasto, bg="#F44336", fg="white").pack(pady=5)

# Tabla
cols = ["Item", "Factura", "Descripción", "Unidad", "Cantidad", "Precio (Bs)", "Precio ($)", "Fecha", "Total", "Proveedor"]
tabla = ttk.Treeview(root, columns=cols, show="headings", height=15)
for col in cols:
    tabla.heading(col, text=col)
    tabla.column(col, width=100)
tabla.pack()

# Total
label_total = tk.Label(root, text="TOTAL GENERAL: Bs. 0.00", font=("Arial", 12, "bold"))
label_total.pack(pady=10)

# Exportar
tk.Button(root, text="Exportar a Excel", command=exportar_excel).pack(pady=5)

# Cargar datos y activar guardado automático
cargar_datos()
auto_guardar()
root.mainloop()
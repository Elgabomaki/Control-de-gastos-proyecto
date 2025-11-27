diff --git a/importtkinterastk.py b/importtkinterastk.py
index 23974d224f21ab0e6f5d81ae717ab557efca082c..e696e9f8e3b2c0509fc1d09aa0831c4252de76e4 100644
--- a/importtkinterastk.py
+++ b/importtkinterastk.py
@@ -1,232 +1,580 @@
-import tkinter as tk
-from tkinter import ttk, messagebox, filedialog
-import pandas as pd
-import os
-import datetime
-from openpyxl import Workbook
-from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
-
-ARCHIVO_DATOS = "gastos_guardados.xlsx"
-registros = []
-precio_dolar = 0.0
-
-# --- Funciones ---
-def validar_fecha(fecha):
-    try:
-        datetime.datetime.strptime(fecha, "%d/%m/%Y")
-        return True
-    except ValueError:
-        return False
-
-def set_dolar():
-    global precio_dolar
-    try:
-        precio_dolar = float(entry_dolar.get())
-        messagebox.showinfo("Éxito", f"Precio del dólar actualizado a Bs. {precio_dolar}")
-    except ValueError:
-        messagebox.showerror("Error", "Ingresa un valor numérico válido")
-
-def guardar_datos():
-    if registros:
-        df = pd.DataFrame(registros)
-        df.to_excel(ARCHIVO_DATOS, index=False)
-
-def cargar_datos():
-    if os.path.exists(ARCHIVO_DATOS):
-        df = pd.read_excel(ARCHIVO_DATOS)
-        for _, fila in df.iterrows():
-            registros.append(fila.to_dict())
-        actualizar_tabla()
-
-def agregar_gasto():
-    if precio_dolar == 0:
-        messagebox.showwarning("Atención", "Primero ingresa el precio del dólar del día.")
-        return
-    # Validar campos vacíos
-    campos = [entry_factura, entry_desc, entry_unidad, entry_cantidad, entry_precio_bs, entry_fecha, entry_proveedor]
-    for c in campos:
-        if not c.get().strip():
-            messagebox.showerror("Error", "Todos los campos son obligatorios")
-            return
-    # Validar fecha
-    if not validar_fecha(entry_fecha.get()):
-        messagebox.showerror("Error", "Formato de fecha inválido. Use DD/MM/YYYY")
-        return
-    try:
-        item = len(registros) + 1
-        factura = entry_factura.get()
-        desc = entry_desc.get()
-        unidad = entry_unidad.get()
-        cantidad = float(entry_cantidad.get())
-        precio_bs = float(entry_precio_bs.get())
-        precio_usd = round(precio_bs / precio_dolar, 2)
-        fecha = entry_fecha.get()
-        proveedor = entry_proveedor.get()
-        total = round(precio_bs * cantidad, 2)
-
-        registro = {
-            "Item": item,
-            "Factura": factura,
-            "Descripción": desc,
-            "Unidad": unidad,
-            "Cantidad": cantidad,
-            "Precio (Bs)": precio_bs,
-            "Precio ($)": precio_usd,
-            "Fecha": fecha,
-            "Total": total,
-            "Proveedor": proveedor
-        }
-
-        registros.append(registro)
-        actualizar_tabla()
-        limpiar_campos()
-        guardar_datos()
-
-    except ValueError:
-        messagebox.showerror("Error", "Verifica los campos numéricos (cantidad y precio).")
-
-def eliminar_gasto():
-    selected = tabla.selection()
-    if not selected:
-        messagebox.showwarning("Atención", "Selecciona un registro para eliminar")
-        return
-    if messagebox.askyesno("Confirmar", "¿Deseas eliminar el registro seleccionado?"):
-        for s in selected:
-            item_val = tabla.item(s)['values'][0]  # columna Item
-            registros[:] = [r for r in registros if r["Item"] != item_val]
-        actualizar_tabla()
-        guardar_datos()
-
-def actualizar_tabla():
-    for row in tabla.get_children():
-        tabla.delete(row)
-    for idx, r in enumerate(registros):
-        r["Item"] = idx + 1  # actualizar numeración
-        tabla.insert("", "end", values=list(r.values()))
-    total_general = sum(r["Total"] for r in registros)
-    label_total.config(text=f"TOTAL GENERAL: Bs. {round(total_general, 2)}")
-
-def limpiar_campos():
-    for entry in [entry_factura, entry_desc, entry_unidad, entry_cantidad, entry_precio_bs, entry_fecha, entry_proveedor]:
-        entry.delete(0, tk.END)
-
-def auto_guardar():
-    guardar_datos()
-    root.after(60000, auto_guardar)  # cada 60 segundos
-
-def exportar_excel():
-    if not registros:
-        messagebox.showwarning("Atención", "No hay datos para exportar.")
-        return
-
-    archivo = filedialog.asksaveasfilename(
-        defaultextension=".xlsx",
-        filetypes=[("Archivo Excel", "*.xlsx")],
-        title="Guardar reporte de gastos",
-        initialfile="reporte_gastos.xlsx"
-    )
-
-    if not archivo:
-        return
-
-    # Guardar datos base
-    df = pd.DataFrame(registros)
-    df.to_excel(archivo, index=False)
-
-    # --- Dar formato con openpyxl ---
-    from openpyxl import load_workbook
-    from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
-
-    wb = load_workbook(archivo)
-    ws = wb.active
-
-    # Encabezado: fondo azul y texto blanco
-    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
-    header_font = Font(color="FFFFFF", bold=True)
-
-    for cell in ws[1]:
-        cell.fill = header_fill
-        cell.font = header_font
-        cell.alignment = Alignment(horizontal="center", vertical="center")
-
-    # Bordes y alineación de todas las celdas
-    thin_border = Border(
-        left=Side(style="thin"),
-        right=Side(style="thin"),
-        top=Side(style="thin"),
-        bottom=Side(style="thin")
-    )
-
-    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
-        for cell in row:
-            cell.border = thin_border
-            cell.alignment = Alignment(horizontal="center", vertical="center")
-
-    # Resaltar fila del total (última)
-    total_row = ws.max_row
-    for cell in ws[total_row]:
-        cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
-        cell.font = Font(bold=True)
-
-    # Ajustar ancho de columnas automáticamente
-    for col in ws.columns:
-        max_length = 0
-        column = col[0].column_letter
-        for cell in col:
-            if cell.value:
-                max_length = max(max_length, len(str(cell.value)))
-        ws.column_dimensions[column].width = max_length + 3
-
-    wb.save(archivo)
-    wb.close()
-
-    messagebox.showinfo("Éxito", f"Datos exportados y formateados correctamente a:\n{archivo}")
-
-# --- Interfaz gráfica ---
-root = tk.Tk()
-root.title("Control de Gastos - Terra Caliza")
-root.geometry("1150x600")
-
-# Precio dólar
-frame_dolar = tk.Frame(root)
-frame_dolar.pack(pady=10)
-tk.Label(frame_dolar, text="Precio del dólar (Bs):").pack(side=tk.LEFT, padx=5)
-entry_dolar = tk.Entry(frame_dolar, width=10)
-entry_dolar.pack(side=tk.LEFT)
-tk.Button(frame_dolar, text="Actualizar", command=set_dolar).pack(side=tk.LEFT, padx=5)
-
-# Formulario
-frame_form = tk.Frame(root)
-frame_form.pack()
-labels = ["Factura", "Descripción", "Unidad", "Cantidad", "Precio (Bs)", "Fecha (DD/MM/YYYY)", "Proveedor"]
-entries = []
-for i, label_text in enumerate(labels):
-    tk.Label(frame_form, text=label_text).grid(row=0, column=i, padx=5, pady=5)
-    entry = tk.Entry(frame_form, width=15)
-    entry.grid(row=1, column=i, padx=5)
-    entries.append(entry)
-
-(entry_factura, entry_desc, entry_unidad, entry_cantidad, entry_precio_bs, entry_fecha, entry_proveedor) = entries
-
-tk.Button(root, text="Agregar gasto", command=agregar_gasto, bg="#4CAF50", fg="white").pack(pady=10)
-tk.Button(root, text="Eliminar gasto", command=eliminar_gasto, bg="#F44336", fg="white").pack(pady=5)
-
-# Tabla
-cols = ["Item", "Factura", "Descripción", "Unidad", "Cantidad", "Precio (Bs)", "Precio ($)", "Fecha", "Total", "Proveedor"]
-tabla = ttk.Treeview(root, columns=cols, show="headings", height=15)
-for col in cols:
-    tabla.heading(col, text=col)
-    tabla.column(col, width=100)
-tabla.pack()
-
-# Total
-label_total = tk.Label(root, text="TOTAL GENERAL: Bs. 0.00", font=("Arial", 12, "bold"))
-label_total.pack(pady=10)
-
-# Exportar
-tk.Button(root, text="Exportar a Excel", command=exportar_excel).pack(pady=5)
-
-# Cargar datos y activar guardado automático
-cargar_datos()
-auto_guardar()
-root.mainloop()
\ No newline at end of file
+"""Aplicación de control de gastos con Tkinter.
+
+Características principales:
+- Registro de gastos con cálculo automático en dólares.
+- Edición y eliminación desde la tabla.
+- Búsqueda y filtros por fecha, proveedor y montos.
+- Resumen de totales.
+- Guardado automático en un archivo Excel utilizado como base de datos.
+- Exportación con formato profesional mediante openpyxl.
+"""
+
+from __future__ import annotations
+
+import datetime as dt
+import os
+import tkinter as tk
+from dataclasses import dataclass, asdict
+from tkinter import filedialog, messagebox, ttk
+from typing import List, Optional
+
+import pandas as pd
+from openpyxl import Workbook
+from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
+
+ARCHIVO_DATOS = "gastos_guardados.xlsx"
+COLUMNAS = [
+    "Item",
+    "Factura",
+    "Descripción",
+    "Unidad",
+    "Cantidad",
+    "Precio (Bs)",
+    "Precio ($)",
+    "Fecha",
+    "Total",
+    "Proveedor",
+]
+
+
+@dataclass
+class RegistroGasto:
+    item: int
+    factura: str
+    descripcion: str
+    unidad: str
+    cantidad: float
+    precio_bs: float
+    precio_usd: float
+    fecha: str
+    total: float
+    proveedor: str
+
+    @classmethod
+    def from_dict(cls, data: dict[str, object]) -> "RegistroGasto":
+        return cls(
+            item=int(data.get("Item", 0)),
+            factura=str(data.get("Factura", "")),
+            descripcion=str(data.get("Descripción", "")),
+            unidad=str(data.get("Unidad", "")),
+            cantidad=float(data.get("Cantidad", 0)),
+            precio_bs=float(data.get("Precio (Bs)", 0)),
+            precio_usd=float(data.get("Precio ($)", 0)),
+            fecha=str(data.get("Fecha", "")),
+            total=float(data.get("Total", 0)),
+            proveedor=str(data.get("Proveedor", "")),
+        )
+
+    def to_row(self) -> List[object]:
+        return [
+            self.item,
+            self.factura,
+            self.descripcion,
+            self.unidad,
+            self.cantidad,
+            self.precio_bs,
+            self.precio_usd,
+            self.fecha,
+            self.total,
+            self.proveedor,
+        ]
+
+
+class GestorGastos:
+    def __init__(self) -> None:
+        self.registros: List[RegistroGasto] = []
+        self.precio_dolar: float = 0.0
+        self.registro_en_edicion: Optional[int] = None
+
+        self.root = tk.Tk()
+        self.root.title("Control de Gastos - Terra Caliza")
+        self.root.geometry("1280x720")
+        self.root.resizable(True, True)
+
+        self._configurar_estilos()
+        self._crear_componentes()
+        self._cargar_datos()
+        self._refrescar_tabla()
+
+    # --------------------------- Configuración UI ---------------------------
+    def _configurar_estilos(self) -> None:
+        style = ttk.Style(self.root)
+        style.theme_use("clam")
+        style.configure("TButton", padding=6, font=("Arial", 10))
+        style.configure("Primary.TButton", background="#1976D2", foreground="white")
+        style.map("Primary.TButton", background=[("active", "#1565C0")])
+        style.configure("Danger.TButton", background="#C62828", foreground="white")
+        style.map("Danger.TButton", background=[("active", "#B71C1C")])
+        style.configure("Success.TButton", background="#2E7D32", foreground="white")
+        style.map("Success.TButton", background=[("active", "#1B5E20")])
+        style.configure("TLabel", font=("Arial", 10))
+        style.configure("Header.TLabel", font=("Arial", 12, "bold"))
+
+    def _crear_componentes(self) -> None:
+        self._crear_frame_dolar()
+        self._crear_formulario()
+        self._crear_controles_tabla()
+        self._crear_tabla()
+        self._crear_resumen()
+        self._crear_exportacion()
+
+    def _crear_frame_dolar(self) -> None:
+        frame = ttk.LabelFrame(self.root, text="Precio del dólar del día")
+        frame.pack(fill="x", padx=10, pady=8)
+
+        ttk.Label(frame, text="Precio (Bs):").pack(side=tk.LEFT, padx=6)
+        self.entry_dolar = ttk.Entry(frame, width=10)
+        self.entry_dolar.pack(side=tk.LEFT)
+        ttk.Button(frame, text="Actualizar", style="Primary.TButton", command=self._set_dolar).pack(
+            side=tk.LEFT, padx=6
+        )
+
+    def _crear_formulario(self) -> None:
+        self.frame_form = ttk.LabelFrame(self.root, text="Datos del gasto")
+        self.frame_form.pack(fill="x", padx=10, pady=8)
+
+        labels = [
+            "Factura",
+            "Descripción",
+            "Unidad",
+            "Cantidad",
+            "Precio (Bs)",
+            "Fecha (DD/MM/YYYY)",
+            "Proveedor",
+        ]
+        self.entries: list[ttk.Entry] = []
+        for col, texto in enumerate(labels):
+            ttk.Label(self.frame_form, text=texto).grid(row=0, column=col, padx=6, pady=4, sticky="w")
+            entry = ttk.Entry(self.frame_form, width=18)
+            entry.grid(row=1, column=col, padx=6, pady=2)
+            self.entries.append(entry)
+
+        (
+            self.entry_factura,
+            self.entry_desc,
+            self.entry_unidad,
+            self.entry_cantidad,
+            self.entry_precio_bs,
+            self.entry_fecha,
+            self.entry_proveedor,
+        ) = self.entries
+
+        botones = ttk.Frame(self.root)
+        botones.pack(fill="x", padx=10, pady=4)
+        self.btn_agregar = ttk.Button(
+            botones, text="Agregar gasto", style="Success.TButton", command=self._agregar_o_actualizar
+        )
+        self.btn_agregar.pack(side=tk.LEFT, padx=4)
+        ttk.Button(botones, text="Cancelar edición", command=self._cancelar_edicion).pack(side=tk.LEFT, padx=4)
+        ttk.Button(botones, text="Eliminar seleccionado(s)", style="Danger.TButton", command=self._eliminar_gasto).pack(
+            side=tk.LEFT, padx=4
+        )
+        ttk.Button(botones, text="Limpiar lista", command=self._limpiar_lista).pack(side=tk.LEFT, padx=4)
+
+    def _crear_controles_tabla(self) -> None:
+        filtros = ttk.LabelFrame(self.root, text="Búsqueda y filtros")
+        filtros.pack(fill="x", padx=10, pady=8)
+
+        ttk.Label(filtros, text="Buscar (factura/proveedor/desc.):").grid(row=0, column=0, padx=4, pady=2, sticky="w")
+        self.entry_buscar = ttk.Entry(filtros, width=28)
+        self.entry_buscar.grid(row=1, column=0, padx=4, pady=2, sticky="w")
+
+        ttk.Label(filtros, text="Proveedor:").grid(row=0, column=1, padx=4, pady=2, sticky="w")
+        self.entry_filtro_proveedor = ttk.Entry(filtros, width=20)
+        self.entry_filtro_proveedor.grid(row=1, column=1, padx=4, pady=2, sticky="w")
+
+        ttk.Label(filtros, text="Fecha desde (DD/MM/YYYY):").grid(row=0, column=2, padx=4, pady=2, sticky="w")
+        self.entry_fecha_desde = ttk.Entry(filtros, width=14)
+        self.entry_fecha_desde.grid(row=1, column=2, padx=4, pady=2, sticky="w")
+
+        ttk.Label(filtros, text="Fecha hasta (DD/MM/YYYY):").grid(row=0, column=3, padx=4, pady=2, sticky="w")
+        self.entry_fecha_hasta = ttk.Entry(filtros, width=14)
+        self.entry_fecha_hasta.grid(row=1, column=3, padx=4, pady=2, sticky="w")
+
+        ttk.Label(filtros, text="Monto mínimo (Bs):").grid(row=0, column=4, padx=4, pady=2, sticky="w")
+        self.entry_monto_min = ttk.Entry(filtros, width=12)
+        self.entry_monto_min.grid(row=1, column=4, padx=4, pady=2, sticky="w")
+
+        ttk.Label(filtros, text="Monto máximo (Bs):").grid(row=0, column=5, padx=4, pady=2, sticky="w")
+        self.entry_monto_max = ttk.Entry(filtros, width=12)
+        self.entry_monto_max.grid(row=1, column=5, padx=4, pady=2, sticky="w")
+
+        ttk.Button(filtros, text="Aplicar filtros", command=self._aplicar_filtros).grid(row=1, column=6, padx=6)
+        ttk.Button(filtros, text="Limpiar filtros", command=self._limpiar_filtros).grid(row=1, column=7, padx=6)
+
+    def _crear_tabla(self) -> None:
+        frame = ttk.Frame(self.root)
+        frame.pack(fill="both", expand=True, padx=10, pady=8)
+
+        self.tabla = ttk.Treeview(frame, columns=COLUMNAS, show="headings", height=15)
+        for col in COLUMNAS:
+            self.tabla.heading(col, text=col)
+            ancho = 120 if col != "Descripción" else 180
+            self.tabla.column(col, width=ancho, anchor="center")
+
+        scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=self.tabla.yview)
+        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=self.tabla.xview)
+        self.tabla.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
+
+        self.tabla.grid(row=0, column=0, sticky="nsew")
+        scrollbar_y.grid(row=0, column=1, sticky="ns")
+        scrollbar_x.grid(row=1, column=0, sticky="ew")
+
+        frame.rowconfigure(0, weight=1)
+        frame.columnconfigure(0, weight=1)
+
+        self.tabla.bind("<Double-1>", self._iniciar_edicion)
+
+    def _crear_resumen(self) -> None:
+        frame = ttk.Frame(self.root)
+        frame.pack(fill="x", padx=10, pady=6)
+
+        self.label_total_bs = ttk.Label(frame, text="Total Bs: 0.00", style="Header.TLabel")
+        self.label_total_bs.pack(side=tk.LEFT, padx=8)
+
+        self.label_total_usd = ttk.Label(frame, text="Total $: 0.00", style="Header.TLabel")
+        self.label_total_usd.pack(side=tk.LEFT, padx=8)
+
+        self.label_cantidad = ttk.Label(frame, text="Registros: 0", style="Header.TLabel")
+        self.label_cantidad.pack(side=tk.LEFT, padx=8)
+
+    def _crear_exportacion(self) -> None:
+        frame = ttk.Frame(self.root)
+        frame.pack(fill="x", padx=10, pady=8)
+
+        ttk.Button(frame, text="Exportar a Excel", style="Primary.TButton", command=self._exportar_excel).pack(
+            side=tk.LEFT, padx=4
+        )
+
+    # ------------------------------ Utilidades ------------------------------
+    @staticmethod
+    def _validar_fecha(fecha: str) -> bool:
+        try:
+            dt.datetime.strptime(fecha, "%d/%m/%Y")
+            return True
+        except ValueError:
+            return False
+
+    def _mostrar_error(self, mensaje: str) -> None:
+        messagebox.showerror("Error", mensaje)
+
+    def _limpiar_campos(self) -> None:
+        for entry in self.entries:
+            entry.delete(0, tk.END)
+        self.registro_en_edicion = None
+        self.btn_agregar.config(text="Agregar gasto")
+
+    def _recalcular_items(self) -> None:
+        for idx, registro in enumerate(self.registros, start=1):
+            registro.item = idx
+
+    # --------------------------- Manejo de datos ---------------------------
+    def _cargar_datos(self) -> None:
+        if not os.path.exists(ARCHIVO_DATOS):
+            return
+        df = pd.read_excel(ARCHIVO_DATOS)
+        for _, fila in df.iterrows():
+            registro = RegistroGasto.from_dict(fila.to_dict())
+            self.registros.append(registro)
+        self._recalcular_items()
+
+    def _guardar_datos(self) -> None:
+        if not self.registros:
+            if os.path.exists(ARCHIVO_DATOS):
+                os.remove(ARCHIVO_DATOS)
+            return
+        df = pd.DataFrame([asdict(r) for r in self.registros])
+        df.rename(columns={
+            "item": "Item",
+            "factura": "Factura",
+            "descripcion": "Descripción",
+            "unidad": "Unidad",
+            "cantidad": "Cantidad",
+            "precio_bs": "Precio (Bs)",
+            "precio_usd": "Precio ($)",
+            "fecha": "Fecha",
+            "total": "Total",
+            "proveedor": "Proveedor",
+        }, inplace=True)
+        df.to_excel(ARCHIVO_DATOS, index=False)
+
+    # ------------------------------ Operaciones -----------------------------
+    def _set_dolar(self) -> None:
+        try:
+            self.precio_dolar = float(self.entry_dolar.get())
+            messagebox.showinfo("Éxito", f"Precio del dólar actualizado a Bs. {self.precio_dolar}")
+        except ValueError:
+            self._mostrar_error("Ingresa un valor numérico válido para el dólar.")
+
+    def _agregar_o_actualizar(self) -> None:
+        if self.precio_dolar <= 0:
+            messagebox.showwarning("Atención", "Primero ingresa el precio del dólar del día.")
+            return
+
+        campos = [
+            self.entry_factura,
+            self.entry_desc,
+            self.entry_unidad,
+            self.entry_cantidad,
+            self.entry_precio_bs,
+            self.entry_fecha,
+            self.entry_proveedor,
+        ]
+        if any(not c.get().strip() for c in campos):
+            self._mostrar_error("Todos los campos son obligatorios.")
+            return
+
+        if not self._validar_fecha(self.entry_fecha.get().strip()):
+            self._mostrar_error("Formato de fecha inválido. Use DD/MM/YYYY")
+            return
+
+        try:
+            cantidad = float(self.entry_cantidad.get())
+            precio_bs = float(self.entry_precio_bs.get())
+        except ValueError:
+            self._mostrar_error("Verifica los campos numéricos (cantidad y precio).")
+            return
+
+        precio_usd = round(precio_bs / self.precio_dolar, 2)
+        total_bs = round(precio_bs * cantidad, 2)
+
+        if self.registro_en_edicion is not None:
+            registro = self.registros[self.registro_en_edicion]
+            registro.factura = self.entry_factura.get().strip()
+            registro.descripcion = self.entry_desc.get().strip()
+            registro.unidad = self.entry_unidad.get().strip()
+            registro.cantidad = cantidad
+            registro.precio_bs = precio_bs
+            registro.precio_usd = precio_usd
+            registro.fecha = self.entry_fecha.get().strip()
+            registro.total = total_bs
+            registro.proveedor = self.entry_proveedor.get().strip()
+        else:
+            nuevo = RegistroGasto(
+                item=len(self.registros) + 1,
+                factura=self.entry_factura.get().strip(),
+                descripcion=self.entry_desc.get().strip(),
+                unidad=self.entry_unidad.get().strip(),
+                cantidad=cantidad,
+                precio_bs=precio_bs,
+                precio_usd=precio_usd,
+                fecha=self.entry_fecha.get().strip(),
+                total=total_bs,
+                proveedor=self.entry_proveedor.get().strip(),
+            )
+            self.registros.append(nuevo)
+
+        self._recalcular_items()
+        self._guardar_datos()
+        self._limpiar_campos()
+        self._refrescar_tabla()
+        messagebox.showinfo("Éxito", "Registro guardado correctamente.")
+
+    def _eliminar_gasto(self) -> None:
+        seleccion = self.tabla.selection()
+        if not seleccion:
+            messagebox.showwarning("Atención", "Selecciona uno o más registros para eliminar.")
+            return
+        if not messagebox.askyesno("Confirmar", "¿Deseas eliminar los registros seleccionados?"):
+            return
+        items_a_eliminar = {self.tabla.item(s)["values"][0] for s in seleccion}
+        self.registros = [r for r in self.registros if r.item not in items_a_eliminar]
+        self._recalcular_items()
+        self._guardar_datos()
+        self._limpiar_campos()
+        self._refrescar_tabla()
+
+    def _limpiar_lista(self) -> None:
+        if not self.registros:
+            return
+        if messagebox.askyesno("Confirmar", "Esto eliminará todos los registros. ¿Deseas continuar?"):
+            self.registros.clear()
+            self._guardar_datos()
+            self._limpiar_campos()
+            self._refrescar_tabla()
+
+    def _iniciar_edicion(self, event: tk.Event[tk.Misc]) -> None:
+        seleccionado = self.tabla.focus()
+        if not seleccionado:
+            return
+        valores = self.tabla.item(seleccionado)["values"]
+        if not valores:
+            return
+        item_id = int(valores[0])
+        indice = item_id - 1
+        if indice < 0 or indice >= len(self.registros):
+            return
+        registro = self.registros[indice]
+        self.registro_en_edicion = indice
+        self.entry_factura.delete(0, tk.END)
+        self.entry_factura.insert(0, registro.factura)
+        self.entry_desc.delete(0, tk.END)
+        self.entry_desc.insert(0, registro.descripcion)
+        self.entry_unidad.delete(0, tk.END)
+        self.entry_unidad.insert(0, registro.unidad)
+        self.entry_cantidad.delete(0, tk.END)
+        self.entry_cantidad.insert(0, str(registro.cantidad))
+        self.entry_precio_bs.delete(0, tk.END)
+        self.entry_precio_bs.insert(0, str(registro.precio_bs))
+        self.entry_fecha.delete(0, tk.END)
+        self.entry_fecha.insert(0, registro.fecha)
+        self.entry_proveedor.delete(0, tk.END)
+        self.entry_proveedor.insert(0, registro.proveedor)
+        self.btn_agregar.config(text="Actualizar gasto")
+
+    def _cancelar_edicion(self) -> None:
+        self._limpiar_campos()
+
+    # --------------------------- Filtros y tabla ---------------------------
+    def _aplicar_filtros(self) -> None:
+        texto_busqueda = self.entry_buscar.get().strip().lower()
+        proveedor = self.entry_filtro_proveedor.get().strip().lower()
+        fecha_desde = self.entry_fecha_desde.get().strip()
+        fecha_hasta = self.entry_fecha_hasta.get().strip()
+        monto_min = self.entry_monto_min.get().strip()
+        monto_max = self.entry_monto_max.get().strip()
+
+        fecha_inicio_dt = self._parse_fecha(fecha_desde) if fecha_desde else None
+        fecha_fin_dt = self._parse_fecha(fecha_hasta) if fecha_hasta else None
+        if (fecha_desde and not fecha_inicio_dt) or (fecha_hasta and not fecha_fin_dt):
+            self._mostrar_error("Fechas de filtro inválidas. Use DD/MM/YYYY")
+            return
+
+        try:
+            monto_min_val = float(monto_min) if monto_min else None
+            monto_max_val = float(monto_max) if monto_max else None
+        except ValueError:
+            self._mostrar_error("Montos de filtro inválidos.")
+            return
+
+        filtrados: list[RegistroGasto] = []
+        for registro in self.registros:
+            if texto_busqueda:
+                texto = f"{registro.factura} {registro.proveedor} {registro.descripcion}".lower()
+                if texto_busqueda not in texto:
+                    continue
+            if proveedor and proveedor not in registro.proveedor.lower():
+                continue
+            fecha_dt = self._parse_fecha(registro.fecha)
+            if fecha_inicio_dt and (not fecha_dt or fecha_dt < fecha_inicio_dt):
+                continue
+            if fecha_fin_dt and (not fecha_dt or fecha_dt > fecha_fin_dt):
+                continue
+            if monto_min_val is not None and registro.total < monto_min_val:
+                continue
+            if monto_max_val is not None and registro.total > monto_max_val:
+                continue
+            filtrados.append(registro)
+
+        self._refrescar_tabla(filtrados)
+
+    @staticmethod
+    def _parse_fecha(fecha: str) -> Optional[dt.datetime]:
+        try:
+            return dt.datetime.strptime(fecha, "%d/%m/%Y")
+        except ValueError:
+            return None
+
+    def _limpiar_filtros(self) -> None:
+        for entry in [
+            self.entry_buscar,
+            self.entry_filtro_proveedor,
+            self.entry_fecha_desde,
+            self.entry_fecha_hasta,
+            self.entry_monto_min,
+            self.entry_monto_max,
+        ]:
+            entry.delete(0, tk.END)
+        self._refrescar_tabla()
+
+    def _refrescar_tabla(self, registros: Optional[list[RegistroGasto]] = None) -> None:
+        registros = registros if registros is not None else self.registros
+        for row in self.tabla.get_children():
+            self.tabla.delete(row)
+        for registro in registros:
+            self.tabla.insert("", "end", values=registro.to_row())
+        self._actualizar_resumen(registros)
+
+    def _actualizar_resumen(self, registros: list[RegistroGasto]) -> None:
+        total_bs = sum(r.total for r in registros)
+        total_usd = sum(r.precio_usd * r.cantidad for r in registros)
+        self.label_total_bs.config(text=f"Total Bs: {total_bs:,.2f}")
+        self.label_total_usd.config(text=f"Total $: {total_usd:,.2f}")
+        self.label_cantidad.config(text=f"Registros: {len(registros)}")
+
+    # ------------------------------- Exportar -------------------------------
+    def _exportar_excel(self) -> None:
+        if not self.registros:
+            messagebox.showwarning("Atención", "No hay datos para exportar.")
+            return
+
+        archivo = filedialog.asksaveasfilename(
+            defaultextension=".xlsx",
+            filetypes=[("Archivo Excel", "*.xlsx")],
+            title="Guardar reporte de gastos",
+            initialfile="reporte_gastos.xlsx",
+        )
+        if not archivo:
+            return
+
+        wb = Workbook()
+        ws = wb.active
+
+        titulo = f"Reporte de gastos - {dt.datetime.now().strftime('%d/%m/%Y')}"
+        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNAS))
+        celda_titulo = ws.cell(row=1, column=1, value=titulo)
+        celda_titulo.font = Font(size=16, bold=True)
+        celda_titulo.alignment = Alignment(horizontal="center", vertical="center")
+
+        header_row = 3
+        for col_idx, encabezado in enumerate(COLUMNAS, start=1):
+            celda = ws.cell(row=header_row, column=col_idx, value=encabezado)
+            celda.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
+            celda.font = Font(color="FFFFFF", bold=True)
+            celda.alignment = Alignment(horizontal="center", vertical="center")
+
+        for row_idx, registro in enumerate(self.registros, start=header_row + 1):
+            for col_idx, valor in enumerate(registro.to_row(), start=1):
+                celda = ws.cell(row=row_idx, column=col_idx, value=valor)
+                celda.alignment = Alignment(horizontal="center", vertical="center")
+
+        total_row_idx = header_row + len(self.registros) + 1
+        ws.cell(row=total_row_idx, column=1, value="Totales").font = Font(bold=True)
+        ws.merge_cells(start_row=total_row_idx, start_column=1, end_row=total_row_idx, end_column=2)
+
+        total_bs = sum(r.total for r in self.registros)
+        total_usd = sum(r.precio_usd * r.cantidad for r in self.registros)
+        ws.cell(row=total_row_idx, column=7, value=total_usd).number_format = "#,##0.00"
+        ws.cell(row=total_row_idx, column=9, value=total_bs).number_format = "#,##0.00"
+
+        for col in range(1, len(COLUMNAS) + 1):
+            celda = ws.cell(row=total_row_idx, column=col)
+            celda.fill = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
+            celda.font = Font(bold=True)
+            celda.alignment = Alignment(horizontal="center", vertical="center")
+
+        thin = Border(
+            left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin")
+        )
+        for row in ws.iter_rows(min_row=header_row, max_row=total_row_idx, max_col=len(COLUMNAS)):
+            for celda in row:
+                celda.border = thin
+
+        for col_idx in range(1, len(COLUMNAS) + 1):
+            max_length = max(
+                len(str(ws.cell(row=row, column=col_idx).value or "")) for row in range(1, total_row_idx + 1)
+            )
+            ws.column_dimensions[ws.cell(row=header_row, column=col_idx).column_letter].width = max_length + 4
+
+        wb.save(archivo)
+        wb.close()
+        messagebox.showinfo("Éxito", f"Datos exportados y formateados correctamente a:\n{archivo}")
+
+    # ------------------------------- Ejecución ------------------------------
+    def ejecutar(self) -> None:
+        self.root.mainloop()
+
+
+if __name__ == "__main__":
+    app = GestorGastos()
+    app.ejecutar()

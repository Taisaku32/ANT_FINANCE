import sys as _sys
import os as _os

# ── PyInstaller frozen mode: re-spawned as Streamlit server ──────────────────
if len(_sys.argv) > 1 and _sys.argv[1] == "--run-dashboard":
    _dashboard_path = _sys.argv[2]
    _sys.argv = ["streamlit", "run", _dashboard_path,
                 "--server.headless=true", "--server.port=8501",
                 "--global.developmentMode=false"]
    try:
        from streamlit.web.cli import main as _st_main
    except ImportError:
        from streamlit.cli import main as _st_main
    _st_main()
    _sys.exit(0)
# ─────────────────────────────────────────────────────────────────────────────

import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
from datetime import date
import pandas as pd
import sqlite3
import os
import subprocess
import threading
import webbrowser
import time

try:
    from tkcalendar import DateEntry
    _HAS_DATE_PICKER = True
except ImportError:
    _HAS_DATE_PICKER = False

# --- Fijar directorio de trabajo al lado del .exe cuando está compilado ---
if getattr(_sys, "frozen", False):
    _APP_DIR = os.path.dirname(_sys.executable)
    os.chdir(_APP_DIR)
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Pedir nombre de usuario ---
root = tk.Tk()
root.withdraw()
usuario = simpledialog.askstring("Usuario", "Ingresa tu nombre de usuario:")
if not usuario:
    messagebox.showerror("Error", "Debes ingresar un nombre de usuario.")
    exit()
root.destroy()

# --- Crear carpeta por usuario ---
BASE_DIR = os.path.join(_APP_DIR, "users", usuario)
os.makedirs(BASE_DIR, exist_ok=True)

# 📁 Archivos por usuario
db_file = os.path.join(BASE_DIR, "finanzas.db")
excel_ingresos = os.path.join(BASE_DIR, "finanzas.xlsx")
excel_gastos   = os.path.join(BASE_DIR, "finanzas_gastos.xlsx")
excel_ahorros  = os.path.join(BASE_DIR, "finanzas_ahorros.xlsx")

# Categorías fijas
FIXED_INCOME_CAT = "Ingresos fijos mensuales"
FIXED_EXPENSE_CAT = "Gastos fijos mensuales"


# Crear base de datos y migrar
def crear_base_datos():
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # --- Tablas principales ---
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS ingresos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            monto REAL,
            categoria TEXT,
            fecha TEXT
        )"""
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS gastos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            monto REAL,
            categoria TEXT,
            fecha TEXT
        )"""
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS ahorros (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            monto          REAL    NOT NULL,
            tipo           TEXT    NOT NULL CHECK(tipo IN ('deposito', 'retiro')),
            categoria      TEXT,
            category_id    INTEGER REFERENCES categories(id),
            subcategory_id INTEGER REFERENCES subcategories(id),
            activity_name  TEXT,
            fecha          TEXT    NOT NULL
        )"""
    )

    # --- Tablas de soporte ---
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )"""
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS subcategories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER NOT NULL REFERENCES categories(id),
            name TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )"""
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS budgets (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT    NOT NULL UNIQUE,
            monto     REAL    NOT NULL CHECK(monto > 0)
        )"""
    )

    # --- Agregar columnas nuevas si faltan ---
    def _has_column(table, col):
        cursor.execute(f"PRAGMA table_info({table})")
        return any(row[1] == col for row in cursor.fetchall())

    for table in ("ingresos", "gastos"):
        if not _has_column(table, "category_id"):
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN category_id INTEGER")
        if not _has_column(table, "subcategory_id"):
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN subcategory_id INTEGER")
        if not _has_column(table, "activity_name"):
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN activity_name TEXT")

    # --- Sembrar categorías por defecto ---
    defaults = [FIXED_INCOME_CAT, FIXED_EXPENSE_CAT, "Sin categoría"]
    for name in defaults:
        cursor.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (name,))

    conn.commit()

    # --- Backfill: migrar categoria TEXT → category_id ---
    for table in ("ingresos", "gastos"):
        cursor.execute(
            f"SELECT DISTINCT categoria FROM {table} WHERE categoria IS NOT NULL AND category_id IS NULL"
        )
        legacy = [row[0] for row in cursor.fetchall()]
        for cat_name in legacy:
            # Crear categoría si no existe
            cursor.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (cat_name,))
            cursor.execute("SELECT id FROM categories WHERE name = ?", (cat_name,))
            cat_id = cursor.fetchone()[0]
            cursor.execute(
                f"UPDATE {table} SET category_id = ? WHERE categoria = ? AND category_id IS NULL",
                (cat_id, cat_name),
            )

    # Filas sin categoria → "Sin categoría"
    cursor.execute("SELECT id FROM categories WHERE name = 'Sin categoría'")
    fallback_id = cursor.fetchone()[0]
    for table in ("ingresos", "gastos"):
        cursor.execute(
            f"UPDATE {table} SET category_id = ? WHERE category_id IS NULL",
            (fallback_id,),
        )

    conn.commit()
    conn.close()


# --- Helpers de categorías ---
def obtener_categorias():
    """Retorna lista de (id, name) ordenada por nombre."""
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM categories ORDER BY name")
    rows = cursor.fetchall()
    conn.close()
    return rows


def obtener_subcategorias(category_id):
    """Retorna lista de (id, name) para la categoría dada."""
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY name",
        (category_id,),
    )
    rows = cursor.fetchall()
    conn.close()
    return rows


def obtener_saldo_ahorros():
    """Retorna el saldo acumulado de ahorros (depósitos - retiros)."""
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT COALESCE(SUM(CASE WHEN tipo='deposito' THEN monto ELSE 0 END), 0) - "
        "       COALESCE(SUM(CASE WHEN tipo='retiro'   THEN monto ELSE 0 END), 0) "
        "FROM ahorros"
    )
    saldo = cursor.fetchone()[0]
    conn.close()
    return saldo if saldo is not None else 0.0


# --- Guardar en Excel ---
def guardar_en_excel(file, monto, categoria, fecha, subcategoria=None, activity_name=None):
    nueva = pd.DataFrame([{
        "Monto": monto,
        "Categoría": categoria,
        "Fecha": fecha,
        "Subcategoría": subcategoria if subcategoria else "",
        "Nombre de la actividad": activity_name if activity_name else "",
    }])
    if os.path.exists(file):
        df = pd.read_excel(file, engine="openpyxl")
        # Agregar columnas faltantes para compatibilidad con archivos antiguos
        for col in ("Subcategoría", "Nombre de la actividad"):
            if col not in df.columns:
                df[col] = ""
        df = pd.concat([df, nueva], ignore_index=True)
    else:
        df = nueva
    df.to_excel(file, index=False, engine="openpyxl")


def guardar_en_excel_ahorros(monto, tipo, categoria, fecha, subcategoria=None, activity_name=None):
    nueva = pd.DataFrame([{
        "Monto": monto,
        "Tipo": tipo,
        "Categoría": categoria,
        "Subcategoría": subcategoria if subcategoria else "",
        "Nombre de la actividad": activity_name if activity_name else "",
        "Fecha": fecha,
    }])
    if os.path.exists(excel_ahorros):
        df = pd.read_excel(excel_ahorros, engine="openpyxl")
        for col in ("Subcategoría", "Nombre de la actividad"):
            if col not in df.columns:
                df[col] = ""
        df = pd.concat([df, nueva], ignore_index=True)
    else:
        df = nueva
    df.to_excel(excel_ahorros, index=False, engine="openpyxl")


# --- Guardar en SQLite ---
def guardar_en_sqlite(table, monto, categoria, fecha, category_id=None, subcategory_id=None, activity_name=None):
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    cursor.execute(
        f"INSERT INTO {table} (monto, categoria, fecha, category_id, subcategory_id, activity_name) "
        f"VALUES (?, ?, ?, ?, ?, ?)",
        (monto, categoria, fecha, category_id, subcategory_id, activity_name),
    )
    conn.commit()
    conn.close()


def guardar_en_sqlite_ahorros(monto, tipo, categoria, fecha, category_id=None, subcategory_id=None, activity_name=None):
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO ahorros (monto, tipo, categoria, fecha, category_id, subcategory_id, activity_name) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (monto, tipo, categoria, fecha, category_id, subcategory_id, activity_name),
    )
    conn.commit()
    conn.close()


# --- Ventana genérica de registro ---
def abrir_ventana(tipo, table, excel_file, preset_cat=None):
    win = tk.Toplevel()
    win.title(f"Registrar {tipo}")
    win.geometry("340x440")
    win.resizable(False, False)

    # Monto
    tk.Label(win, text="Monto:").pack(pady=(10, 2))
    entry_monto = tk.Entry(win, width=28)
    entry_monto.pack()

    # Cargar categorías para el combobox
    categorias = obtener_categorias()  # [(id, name), ...]
    cat_names = [c[1] for c in categorias]
    cat_id_map = {c[1]: c[0] for c in categorias}

    tk.Label(win, text="Categoría:").pack(pady=(10, 2))
    combo_cat = ttk.Combobox(win, values=cat_names, state="readonly", width=26)
    combo_cat.pack()

    # Pre-seleccionar categoría si viene preset
    if preset_cat and preset_cat in cat_names:
        combo_cat.set(preset_cat)

    # Subcategoría (se actualiza al cambiar categoría)
    tk.Label(win, text="Subcategoría (opcional):").pack(pady=(10, 2))
    combo_sub = ttk.Combobox(win, values=[], state="readonly", width=26)
    combo_sub.pack()

    sub_id_map = {}

    def actualizar_subcategorias(event=None):
        nombre_cat = combo_cat.get()
        combo_sub.set("")
        if nombre_cat and nombre_cat in cat_id_map:
            subs = obtener_subcategorias(cat_id_map[nombre_cat])
            sub_names = [s[1] for s in subs]
            sub_id_map.clear()
            sub_id_map.update({s[1]: s[0] for s in subs})
            combo_sub["values"] = sub_names
        else:
            combo_sub["values"] = []

    combo_cat.bind("<<ComboboxSelected>>", actualizar_subcategorias)

    # Pre-cargar subcategorías si hay preset
    if preset_cat:
        actualizar_subcategorias()

    # Nombre de la actividad
    tk.Label(win, text="Nombre de la actividad (opcional):").pack(pady=(10, 2))
    entry_actividad = tk.Entry(win, width=28)
    entry_actividad.pack()

    # Fecha
    tk.Label(win, text="Fecha:").pack(pady=(10, 2))
    if _HAS_DATE_PICKER:
        entry_fecha = DateEntry(win, width=26, date_pattern='yyyy-mm-dd',
                                year=date.today().year, month=date.today().month,
                                day=date.today().day)
    else:
        entry_fecha = tk.Entry(win, width=28)
        entry_fecha.insert(0, str(date.today()))
    entry_fecha.pack()

    def guardar():
        monto_str  = entry_monto.get().strip()
        cat_name   = combo_cat.get().strip()
        sub_name   = combo_sub.get().strip()
        actividad  = entry_actividad.get().strip()
        fecha      = entry_fecha.get().strip()

        if not monto_str:
            messagebox.showerror("Error", "El campo Monto es obligatorio.")
            return
        if not fecha:
            messagebox.showerror("Error", "El campo Fecha es obligatorio.")
            return
        if not cat_name:
            messagebox.showerror("Error", "Debes seleccionar una Categoría.")
            return

        try:
            monto = float(monto_str)
        except ValueError:
            messagebox.showerror("Error", "El monto debe ser numérico.")
            return

        category_id   = cat_id_map.get(cat_name)
        subcategory_id = sub_id_map.get(sub_name) if sub_name else None

        guardar_en_excel(
            excel_file, monto, cat_name, fecha,
            subcategoria=sub_name if sub_name else None,
            activity_name=actividad if actividad else None,
        )
        guardar_en_sqlite(
            table, monto, cat_name, fecha,
            category_id=category_id,
            subcategory_id=subcategory_id,
            activity_name=actividad if actividad else None,
        )
        messagebox.showinfo("Guardado", f"{tipo} registrado con éxito.")
        win.destroy()

    tk.Button(win, text="Guardar", command=guardar, width=15, bg="#d9edf7").pack(pady=15)


# Funciones específicas
def abrir_ventana_ingreso():
    abrir_ventana("Ingreso", "ingresos", excel_ingresos)

def abrir_ventana_gasto():
    abrir_ventana("Gasto", "gastos", excel_gastos)

def abrir_ventana_ingreso_fijo():
    abrir_ventana("Ingreso Fijo", "ingresos", excel_ingresos, preset_cat=FIXED_INCOME_CAT)

def abrir_ventana_gasto_fijo():
    abrir_ventana("Gasto Fijo", "gastos", excel_gastos, preset_cat=FIXED_EXPENSE_CAT)


def abrir_ventana_ahorro(tipo_ahorro):
    titulo = "Ahorro (Ingreso)" if tipo_ahorro == "deposito" else "Ahorro (Retiro)"
    bg_color = "#d9f0e8" if tipo_ahorro == "deposito" else "#f0e0d9"

    win = tk.Toplevel()
    win.title(titulo)
    win.geometry("340x440")
    win.resizable(False, False)

    # Monto
    tk.Label(win, text="Monto:").pack(pady=(10, 2))
    entry_monto = tk.Entry(win, width=28)
    entry_monto.pack()

    # Categorías
    categorias = obtener_categorias()
    cat_names = [c[1] for c in categorias]
    cat_id_map = {c[1]: c[0] for c in categorias}

    tk.Label(win, text="Categoría:").pack(pady=(10, 2))
    combo_cat = ttk.Combobox(win, values=cat_names, state="readonly", width=26)
    combo_cat.pack()

    # Subcategoría
    tk.Label(win, text="Subcategoría (opcional):").pack(pady=(10, 2))
    combo_sub = ttk.Combobox(win, values=[], state="readonly", width=26)
    combo_sub.pack()

    sub_id_map = {}

    def actualizar_subcategorias(event=None):
        nombre_cat = combo_cat.get()
        combo_sub.set("")
        if nombre_cat and nombre_cat in cat_id_map:
            subs = obtener_subcategorias(cat_id_map[nombre_cat])
            sub_names = [s[1] for s in subs]
            sub_id_map.clear()
            sub_id_map.update({s[1]: s[0] for s in subs})
            combo_sub["values"] = sub_names
        else:
            combo_sub["values"] = []

    combo_cat.bind("<<ComboboxSelected>>", actualizar_subcategorias)

    # Nombre de la actividad
    tk.Label(win, text="Nombre de la actividad (opcional):").pack(pady=(10, 2))
    entry_actividad = tk.Entry(win, width=28)
    entry_actividad.pack()

    # Fecha
    tk.Label(win, text="Fecha:").pack(pady=(10, 2))
    if _HAS_DATE_PICKER:
        entry_fecha = DateEntry(win, width=26, date_pattern='yyyy-mm-dd',
                                year=date.today().year, month=date.today().month,
                                day=date.today().day)
    else:
        entry_fecha = tk.Entry(win, width=28)
        entry_fecha.insert(0, str(date.today()))
    entry_fecha.pack()

    def guardar():
        monto_str = entry_monto.get().strip()
        cat_name  = combo_cat.get().strip()
        sub_name  = combo_sub.get().strip()
        actividad = entry_actividad.get().strip()
        fecha     = entry_fecha.get().strip()

        if not monto_str:
            messagebox.showerror("Error", "El campo Monto es obligatorio.")
            return
        if not fecha:
            messagebox.showerror("Error", "El campo Fecha es obligatorio.")
            return
        if not cat_name:
            messagebox.showerror("Error", "Debes seleccionar una Categoría.")
            return
        try:
            monto = float(monto_str)
        except ValueError:
            messagebox.showerror("Error", "El monto debe ser numérico.")
            return
        if monto <= 0:
            messagebox.showerror("Error", "El monto debe ser mayor a 0.")
            return
        if tipo_ahorro == "retiro":
            saldo_actual = obtener_saldo_ahorros()
            if monto > saldo_actual:
                messagebox.showerror(
                    "Saldo insuficiente",
                    f"No puedes retirar ${monto:,.2f}.\n"
                    f"Saldo actual de ahorros: ${saldo_actual:,.2f}"
                )
                return

        category_id   = cat_id_map.get(cat_name)
        subcategory_id = sub_id_map.get(sub_name) if sub_name else None

        guardar_en_excel_ahorros(
            monto, tipo_ahorro, cat_name, fecha,
            subcategoria=sub_name if sub_name else None,
            activity_name=actividad if actividad else None,
        )
        guardar_en_sqlite_ahorros(
            monto, tipo_ahorro, cat_name, fecha,
            category_id=category_id,
            subcategory_id=subcategory_id,
            activity_name=actividad if actividad else None,
        )
        messagebox.showinfo("Guardado", f"{titulo} registrado con éxito.")
        win.destroy()

    tk.Button(win, text="Guardar", command=guardar, width=15, bg=bg_color).pack(pady=15)


def abrir_ventana_ahorro_deposito():
    abrir_ventana_ahorro("deposito")


def abrir_ventana_ahorro_retiro():
    abrir_ventana_ahorro("retiro")


# --- Abrir Dashboard ---
_dashboard_proc = None

def _free_port_8501():
    """Mata cualquier proceso que esté escuchando en el puerto 8501."""
    try:
        result = subprocess.run(
            ['netstat', '-aon'], capture_output=True, text=True, timeout=5
        )
        for line in result.stdout.splitlines():
            if ':8501 ' in line and 'LISTENING' in line:
                pid = int(line.split()[-1])
                if pid > 0:
                    subprocess.run(['taskkill', '/F', '/PID', str(pid)],
                                   capture_output=True)
                break
        time.sleep(0.5)
    except Exception:
        pass


def abrir_dashboard():
    global _dashboard_proc

    # Terminar proceso previo si existe
    if _dashboard_proc is not None:
        try:
            _dashboard_proc.terminate()
            _dashboard_proc.wait(timeout=3)
        except Exception:
            pass
        _dashboard_proc = None

    # Liberar el puerto para que arranque con el código actualizado
    _free_port_8501()

    if getattr(_sys, "frozen", False):
        # Modo compilado
        script_path = os.path.join(_sys._MEIPASS, "dashboard.py")
        _dashboard_proc = subprocess.Popen(
            [_sys.executable, "--run-dashboard", script_path],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    else:
        # Modo desarrollo normal
        _dashboard_proc = subprocess.Popen(
            [_sys.executable, "-m", "streamlit", "run", "dashboard.py",
             "--server.headless=true", "--server.port=8501"],
            cwd=_APP_DIR,
            creationflags=subprocess.CREATE_NO_WINDOW if _sys.platform == "win32" else 0,
        )

    def _open_browser():
        time.sleep(5)
        webbrowser.open("http://localhost:8501")

    threading.Thread(target=_open_browser, daemon=True).start()


# Mostrar datos
def ver_tabla(table):
    conn = sqlite3.connect(db_file)
    df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
    conn.close()
    win = tk.Toplevel()
    win.title(f"{table.capitalize()}")
    win.geometry("700x300")
    cols = df.columns.tolist()
    tree = ttk.Treeview(win, columns=cols, show="headings")
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=100, anchor='center')
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))
    tree.pack(expand=True, fill='both')
    tk.Button(win, text="Cerrar", command=win.destroy).pack(pady=5)


# Funciones de ver
def ver_ingresos(): ver_tabla("ingresos")
def ver_gastos():   ver_tabla("gastos")


# --- Gestión de Categorías ---
def gestionar_categorias():
    win = tk.Toplevel()
    win.title("Gestionar Categorías")
    win.geometry("360x420")
    win.resizable(False, False)

    tk.Label(win, text="Categorías", font=("Helvetica", 13, "bold")).pack(pady=(10, 4))

    # Treeview
    frame_tree = tk.Frame(win)
    frame_tree.pack(fill="both", expand=True, padx=10)
    cols = ("ID", "Nombre")
    tree = ttk.Treeview(frame_tree, columns=cols, show="headings", height=10)
    tree.heading("ID",     text="ID");     tree.column("ID",     width=40,  anchor="center")
    tree.heading("Nombre", text="Nombre"); tree.column("Nombre", width=260, anchor="w")
    sb = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    def _recargar():
        tree.delete(*tree.get_children())
        for cat_id, cat_name in obtener_categorias():
            tree.insert("", "end", values=(cat_id, cat_name))

    _recargar()

    # Agregar
    frame_add = tk.Frame(win)
    frame_add.pack(pady=(8, 2), padx=10, fill="x")
    tk.Label(frame_add, text="Nueva categoría:").pack(side="left")
    entry_nueva = tk.Entry(frame_add, width=22)
    entry_nueva.pack(side="left", padx=(4, 0))

    def _agregar():
        nombre = entry_nueva.get().strip()
        if not nombre:
            messagebox.showerror("Error", "El nombre no puede estar vacío.", parent=win)
            return
        conn = sqlite3.connect(db_file)
        try:
            conn.execute("INSERT INTO categories (name) VALUES (?)", (nombre,))
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f'La categoría "{nombre}" ya existe.', parent=win)
        finally:
            conn.close()
        entry_nueva.delete(0, "end")
        _recargar()

    tk.Button(frame_add, text="Agregar", command=_agregar, bg="#dff0d8").pack(side="left", padx=(6, 0))

    # Eliminar
    def _eliminar():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona una categoría para eliminar.", parent=win)
            return
        cat_id, cat_name = tree.item(sel[0], "values")
        if not messagebox.askyesno("Confirmar", f'¿Eliminar la categoría "{cat_name}"?\n'
                                   'Las transacciones asociadas quedarán sin categoría vinculada.',
                                   parent=win):
            return
        conn = sqlite3.connect(db_file)
        conn.execute("DELETE FROM subcategories WHERE category_id = ?", (cat_id,))
        conn.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
        conn.commit()
        conn.close()
        _recargar()

    tk.Button(win, text="Eliminar seleccionada", command=_eliminar, bg="#f2dede", width=22).pack(pady=(4, 4))

    # Limpiar toda la lista
    def _limpiar_todo():
        if not messagebox.askyesno(
            "Confirmar",
            "¿Borrar TODAS las categorías y subcategorías?\n\n"
            "Las transacciones existentes NO se eliminarán.\n"
            "Se crearán solo las 3 categorías por defecto.",
            parent=win,
        ):
            return
        conn = sqlite3.connect(db_file)
        conn.execute("DELETE FROM subcategories")
        conn.execute("DELETE FROM categories")
        for name in (FIXED_INCOME_CAT, FIXED_EXPENSE_CAT, "Sin categoría"):
            conn.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (name,))
        conn.commit()
        conn.close()
        _recargar()

    tk.Button(win, text="Limpiar toda la lista", command=_limpiar_todo, bg="#fcf8e3", width=22).pack(pady=(0, 10))


# --- Gestión de Subcategorías ---
def gestionar_subcategorias():
    win = tk.Toplevel()
    win.title("Gestionar Subcategorías")
    win.geometry("460x460")
    win.resizable(False, False)

    tk.Label(win, text="Subcategorías", font=("Helvetica", 13, "bold")).pack(pady=(10, 4))

    # Treeview
    frame_tree = tk.Frame(win)
    frame_tree.pack(fill="both", expand=True, padx=10)
    cols = ("ID", "Categoría", "Nombre")
    tree = ttk.Treeview(frame_tree, columns=cols, show="headings", height=10)
    tree.heading("ID",        text="ID");        tree.column("ID",        width=40,  anchor="center")
    tree.heading("Categoría", text="Categoría"); tree.column("Categoría", width=170, anchor="w")
    tree.heading("Nombre",    text="Nombre");    tree.column("Nombre",    width=200, anchor="w")
    sb = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    def _recargar():
        tree.delete(*tree.get_children())
        conn = sqlite3.connect(db_file)
        rows = conn.execute(
            "SELECT s.id, c.name, s.name FROM subcategories s "
            "JOIN categories c ON c.id = s.category_id ORDER BY c.name, s.name"
        ).fetchall()
        conn.close()
        for row in rows:
            tree.insert("", "end", values=row)

    _recargar()

    # Agregar
    frame_add = tk.LabelFrame(win, text=" Nueva subcategoría ", padx=8, pady=6)
    frame_add.pack(pady=(8, 2), padx=10, fill="x")

    tk.Label(frame_add, text="Categoría padre:").grid(row=0, column=0, sticky="w")
    categorias = obtener_categorias()
    cat_names = [c[1] for c in categorias]
    cat_id_map = {c[1]: c[0] for c in categorias}
    combo_cat = ttk.Combobox(frame_add, values=cat_names, state="readonly", width=22)
    combo_cat.grid(row=0, column=1, padx=(4, 0), pady=2)

    tk.Label(frame_add, text="Nombre:").grid(row=1, column=0, sticky="w")
    entry_nombre = tk.Entry(frame_add, width=24)
    entry_nombre.grid(row=1, column=1, padx=(4, 0), pady=2)

    def _agregar():
        cat_name = combo_cat.get().strip()
        nombre   = entry_nombre.get().strip()
        if not cat_name:
            messagebox.showerror("Error", "Selecciona una categoría padre.", parent=win)
            return
        if not nombre:
            messagebox.showerror("Error", "El nombre no puede estar vacío.", parent=win)
            return
        cat_id = cat_id_map[cat_name]
        conn = sqlite3.connect(db_file)
        conn.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (cat_id, nombre))
        conn.commit()
        conn.close()
        entry_nombre.delete(0, "end")
        _recargar()

    tk.Button(frame_add, text="Agregar", command=_agregar, bg="#dff0d8").grid(
        row=2, column=0, columnspan=2, pady=(6, 0))

    # Eliminar
    def _eliminar():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona una subcategoría para eliminar.", parent=win)
            return
        sub_id, cat_name, sub_name = tree.item(sel[0], "values")
        if not messagebox.askyesno("Confirmar", f'¿Eliminar la subcategoría "{sub_name}" de "{cat_name}"?',
                                   parent=win):
            return
        conn = sqlite3.connect(db_file)
        conn.execute("DELETE FROM subcategories WHERE id = ?", (sub_id,))
        conn.commit()
        conn.close()
        _recargar()

    tk.Button(win, text="Eliminar seleccionada", command=_eliminar, bg="#f2dede", width=24).pack(pady=(4, 10))


# Ventana principal
if __name__ == "__main__":
    crear_base_datos()
    root = tk.Tk()
    root.title("Gestor de Finanzas Personales")
    root.geometry("420x760")
    root.resizable(False, False)

    tk.Label(root, text="Gestor de Finanzas", font=("Helvetica", 18, "bold")).pack(pady=20)
    tk.Button(root, text="Registrar Ingreso",      command=abrir_ventana_ingreso,      width=22, height=2, bg="#dff0d8").pack(pady=6)
    tk.Button(root, text="Registrar Gasto",        command=abrir_ventana_gasto,        width=22, height=2, bg="#f2dede").pack(pady=6)
    tk.Button(root, text="Registrar Ingreso Fijo", command=abrir_ventana_ingreso_fijo, width=22, height=2, bg="#d0e9c6").pack(pady=6)
    tk.Button(root, text="Registrar Gasto Fijo",   command=abrir_ventana_gasto_fijo,   width=22, height=2, bg="#f9d0d0").pack(pady=6)
    tk.Button(root, text="Ver Ingresos",           command=ver_ingresos,               width=22, height=2, bg="#d9edf7").pack(pady=6)
    tk.Button(root, text="Ver Gastos",             command=ver_gastos,                 width=22, height=2, bg="#d9edf7").pack(pady=6)
    ttk.Separator(root, orient="horizontal").pack(fill="x", padx=30, pady=8)
    tk.Button(root, text="Ahorro (Ingreso)", command=abrir_ventana_ahorro_deposito, width=22, height=2, bg="#d9f0e8").pack(pady=6)
    tk.Button(root, text="Ahorro (Retiro)",  command=abrir_ventana_ahorro_retiro,   width=22, height=2, bg="#f0e0d9").pack(pady=6)
    ttk.Separator(root, orient="horizontal").pack(fill="x", padx=30, pady=8)
    tk.Button(root, text="Gestionar Categorías",    command=gestionar_categorias,    width=22, height=2, bg="#fff3cd").pack(pady=4)
    tk.Button(root, text="Gestionar Subcategorías", command=gestionar_subcategorias, width=22, height=2, bg="#fff3cd").pack(pady=4)
    ttk.Separator(root, orient="horizontal").pack(fill="x", padx=30, pady=8)
    tk.Button(root, text="📊  Abrir Dashboard", command=abrir_dashboard,
              width=22, height=2, bg="#cce5ff", font=("Helvetica", 11, "bold")).pack(pady=4)

    root.mainloop()

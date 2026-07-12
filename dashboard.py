import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import hmac
from datetime import date, datetime
import plotly.express as px
import smtplib
from email.message import EmailMessage

from turso_db import TursoConnection


st.set_page_config(page_title='Dashboard Financiero', layout='wide')

# Categorías fijas y presupuesto
FIXED_INCOME_CAT = "Ingresos fijos mensuales"
FIXED_EXPENSE_CAT= "Gastos fijos mensuales"

# --- Modo nube (Turso) o modo local (SQLite) ---
# Si existen secretos [usuarios] (en Streamlit Cloud o .streamlit/secrets.toml),
# la app funciona 100% en la nube con login. Si no, usa las carpetas users/ locales.
def _config_usuarios_cloud():
    try:
        if "usuarios" in st.secrets:
            return st.secrets["usuarios"]
    except Exception:
        pass
    return None

USUARIOS_CLOUD = _config_usuarios_cloud()
IS_CLOUD = USUARIOS_CLOUD is not None


def _leer_secreto(nombre):
    try:
        return st.secrets.get(nombre)
    except Exception:
        return None


# --- Selección de usuario / login ---
if IS_CLOUD:
    if not st.session_state.get("usuario_autenticado"):
        st.title("🔐 Finanzas Personales")
        with st.form("form_login"):
            _user_in = st.text_input("Usuario")
            _pass_in = st.text_input("Contraseña", type="password")
            _login_ok = st.form_submit_button("Ingresar")
        if _login_ok:
            _cfg_u = USUARIOS_CLOUD.get(_user_in.strip())
            if _cfg_u is not None and hmac.compare_digest(
                str(_cfg_u.get("password", "")), _pass_in
            ):
                st.session_state["usuario_autenticado"] = _user_in.strip()
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        st.stop()

    usuario = st.session_state["usuario_autenticado"]
    st.sidebar.markdown(f"👤 **{usuario}**")
    if st.sidebar.button("Cerrar sesión", key="btn_logout"):
        del st.session_state["usuario_autenticado"]
        st.rerun()
    DB_FILE = None
else:
    USER_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users")
    usuarios = sorted(
        d for d in os.listdir(USER_ROOT) if os.path.isdir(os.path.join(USER_ROOT, d))
    ) if os.path.isdir(USER_ROOT) else []
    if not usuarios:
        st.error("No hay usuarios. Ejecuta main.py primero.")
        st.stop()
    usuario = st.sidebar.selectbox("Usuario", usuarios)

    # Rutas por usuario
    BASE_DIR = os.path.join(USER_ROOT, usuario)
    DB_FILE = os.path.join(BASE_DIR, "finanzas.db")
    EXCEL_INGRESOS = os.path.join(BASE_DIR, "finanzas.xlsx")
    EXCEL_GASTOS   = os.path.join(BASE_DIR, "finanzas_gastos.xlsx")
    EXCEL_AHORROS  = os.path.join(BASE_DIR, "finanzas_ahorros.xlsx")


def get_conn():
    """Conexión a la base de datos del usuario activo (Turso en la nube, SQLite en local)."""
    if IS_CLOUD:
        _cfg = USUARIOS_CLOUD[usuario]
        return TursoConnection(_cfg["db_url"], _cfg["auth_token"])
    return sqlite3.connect(DB_FILE)


def leer_df(query, params=(), parse_dates=None):
    """Ejecuta un SELECT y devuelve un DataFrame (reemplaza a pd.read_sql)."""
    conn = get_conn()
    try:
        cur = conn.execute(query, params)
        cols = [d[0] for d in cur.description] if cur.description else []
        df = pd.DataFrame(cur.fetchall(), columns=cols)
    finally:
        conn.close()
    if parse_dates:
        for c in parse_dates:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


# Crear tablas si no existen
def crear_tablas_dashboard():
    conn = get_conn()
    conn.execute('''CREATE TABLE IF NOT EXISTS ingresos(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        monto REAL, categoria TEXT, fecha TEXT,
        category_id INTEGER, subcategory_id INTEGER, activity_name TEXT)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS gastos(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        monto REAL, categoria TEXT, fecha TEXT,
        category_id INTEGER, subcategory_id INTEGER, activity_name TEXT)''')
    conn.execute('CREATE TABLE IF NOT EXISTS balances_mensuales(id INTEGER PRIMARY KEY AUTOINCREMENT, año INTEGER, mes INTEGER, balance REAL, fecha_creacion TEXT, UNIQUE(año, mes))')
    conn.execute('CREATE TABLE IF NOT EXISTS categories(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, created_at TEXT DEFAULT CURRENT_TIMESTAMP)')
    conn.execute('CREATE TABLE IF NOT EXISTS subcategories(id INTEGER PRIMARY KEY AUTOINCREMENT, category_id INTEGER NOT NULL REFERENCES categories(id), name TEXT NOT NULL, created_at TEXT DEFAULT CURRENT_TIMESTAMP)')
    conn.execute('''CREATE TABLE IF NOT EXISTS ahorros (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        monto          REAL    NOT NULL,
        tipo           TEXT    NOT NULL CHECK(tipo IN ('deposito', 'retiro')),
        categoria      TEXT,
        category_id    INTEGER REFERENCES categories(id),
        subcategory_id INTEGER REFERENCES subcategories(id),
        activity_name  TEXT,
        fecha          TEXT    NOT NULL
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS metas_ahorro (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        category_id    INTEGER NOT NULL REFERENCES categories(id),
        subcategory_id INTEGER REFERENCES subcategories(id),
        nombre         TEXT    NOT NULL,
        monto_objetivo REAL    NOT NULL CHECK(monto_objetivo > 0),
        created_at     TEXT    DEFAULT CURRENT_TIMESTAMP
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS budgets (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        categoria TEXT    NOT NULL UNIQUE,
        monto     REAL    NOT NULL CHECK(monto > 0)
    )''')
    conn.commit()

    # Migración segura: agregar columnas nuevas en bases antiguas (sin pérdida de datos)
    for _tabla in ("ingresos", "gastos"):
        for _col, _tipo in (("category_id", "INTEGER"),
                            ("subcategory_id", "INTEGER"),
                            ("activity_name", "TEXT")):
            try:
                conn.execute(f"SELECT {_col} FROM {_tabla} LIMIT 1")
            except Exception:
                try:
                    conn.execute(f"ALTER TABLE {_tabla} ADD COLUMN {_col} {_tipo}")
                except Exception:
                    pass

    # Sembrar categorías por defecto
    for _nombre in (FIXED_INCOME_CAT, FIXED_EXPENSE_CAT, "Sin categoría"):
        conn.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (_nombre,))
    conn.commit()
    conn.close()

# Funciones para manejar balances mensuales
def obtener_balance_mes_anterior(año, mes):
    """Obtiene el balance del mes anterior"""
    # Calcular mes y año anterior
    if mes == 1:
        mes_anterior = 12
        año_anterior = año - 1
    else:
        mes_anterior = mes - 1
        año_anterior = año

    conn = get_conn()
    result = conn.execute('SELECT balance FROM balances_mensuales WHERE año=? AND mes=?',
                          (año_anterior, mes_anterior)).fetchone()
    conn.close()
    return result[0] if result else 0.0

def guardar_balance_mensual(año, mes, balance):
    """Guarda o actualiza el balance de un mes específico"""
    # Asegurar que año y mes sean enteros
    año = int(año) if año is not None else date.today().year
    mes = int(mes) if mes is not None else date.today().month

    conn = get_conn()
    fecha_creacion = str(date.today())
    conn.execute('''
        INSERT OR REPLACE INTO balances_mensuales (año, mes, balance, fecha_creacion)
        VALUES (?, ?, ?, ?)
    ''', (año, mes, float(balance), fecha_creacion))
    conn.commit()
    conn.close()

def obtener_balances_guardados():
    """Obtiene todos los balances guardados"""
    df = leer_df('SELECT * FROM balances_mensuales ORDER BY año DESC, mes DESC')
    # Convertir tipos de datos explícitamente para evitar problemas con bytes
    if not df.empty:
        df['año'] = pd.to_numeric(df['año'], errors='coerce').astype('Int64')
        df['mes'] = pd.to_numeric(df['mes'], errors='coerce').astype('Int64')
        df['balance'] = pd.to_numeric(df['balance'], errors='coerce').astype('float64')
    return df

def cargar_datos():
    df_i = leer_df("SELECT * FROM ingresos", parse_dates=["fecha"])
    df_g = leer_df("SELECT * FROM gastos", parse_dates=["fecha"])
    df_i["monto"] = pd.to_numeric(df_i["monto"], errors="coerce")
    df_g["monto"] = pd.to_numeric(df_g["monto"], errors="coerce")
    return df_i, df_g


def cargar_ahorros():
    df_a = leer_df("SELECT * FROM ahorros", parse_dates=["fecha"])
    if not df_a.empty:
        df_a["monto"] = pd.to_numeric(df_a["monto"], errors="coerce")
    return df_a

def cargar_metas_ahorro():
    return leer_df("""
        SELECT m.id, m.nombre, m.monto_objetivo, m.category_id, m.subcategory_id,
               c.name AS categoria_nombre, s.name AS subcategoria_nombre, m.created_at
        FROM metas_ahorro m
        JOIN categories c ON c.id = m.category_id
        LEFT JOIN subcategories s ON s.id = m.subcategory_id
        ORDER BY m.created_at DESC
    """)

def calcular_ahorrado_meta(category_id: int, subcategory_id) -> float:
    sub_id = None if (subcategory_id is None or
                      (isinstance(subcategory_id, float) and pd.isna(subcategory_id))) \
             else int(subcategory_id)
    conn = get_conn()
    if sub_id is None:
        row = conn.execute(
            "SELECT COALESCE(SUM(monto),0.0) FROM ahorros WHERE tipo='deposito' AND category_id=?",
            (category_id,)).fetchone()
    else:
        row = conn.execute(
            "SELECT COALESCE(SUM(monto),0.0) FROM ahorros WHERE tipo='deposito' AND category_id=? AND subcategory_id=?",
            (category_id, sub_id)).fetchone()
    conn.close()
    return float(row[0]) if row else 0.0

def cargar_presupuestos() -> pd.DataFrame:
    try:
        df = leer_df('SELECT id, categoria, monto FROM budgets ORDER BY categoria')
    except Exception:
        df = pd.DataFrame(columns=['id', 'categoria', 'monto'])
    return df


def obtener_saldo_ahorros() -> float:
    """Saldo acumulado de ahorros (depósitos - retiros)."""
    conn = get_conn()
    row = conn.execute(
        "SELECT COALESCE(SUM(CASE WHEN tipo='deposito' THEN monto ELSE 0 END), 0) - "
        "       COALESCE(SUM(CASE WHEN tipo='retiro'   THEN monto ELSE 0 END), 0) "
        "FROM ahorros").fetchone()
    conn.close()
    return float(row[0]) if row and row[0] is not None else 0.0


# --- Guardado en Excel (solo modo local; en la nube se usa el botón de exportar) ---
def _guardar_excel_local(archivo, monto, categoria, fecha, subcategoria=None, actividad=None):
    nueva = pd.DataFrame([{
        "Monto": monto,
        "Categoría": categoria,
        "Fecha": fecha,
        "Subcategoría": subcategoria if subcategoria else "",
        "Nombre de la actividad": actividad if actividad else "",
    }])
    if os.path.exists(archivo):
        df = pd.read_excel(archivo, engine="openpyxl")
        for col in ("Subcategoría", "Nombre de la actividad"):
            if col not in df.columns:
                df[col] = ""
        df = pd.concat([df, nueva], ignore_index=True)
    else:
        df = nueva
    df.to_excel(archivo, index=False, engine="openpyxl")


def _guardar_excel_local_ahorros(monto, tipo, categoria, fecha, subcategoria=None, actividad=None):
    nueva = pd.DataFrame([{
        "Monto": monto,
        "Tipo": tipo,
        "Categoría": categoria,
        "Subcategoría": subcategoria if subcategoria else "",
        "Nombre de la actividad": actividad if actividad else "",
        "Fecha": fecha,
    }])
    if os.path.exists(EXCEL_AHORROS):
        df = pd.read_excel(EXCEL_AHORROS, engine="openpyxl")
        for col in ("Subcategoría", "Nombre de la actividad"):
            if col not in df.columns:
                df[col] = ""
        df = pd.concat([df, nueva], ignore_index=True)
    else:
        df = nueva
    df.to_excel(EXCEL_AHORROS, index=False, engine="openpyxl")


# --- Registro de movimientos ---
def _guardar_movimiento(tabla, monto, cat_name, cat_id, sub_name, sub_id, actividad, fecha):
    conn = get_conn()
    conn.execute(
        f"INSERT INTO {tabla} (monto, categoria, fecha, category_id, subcategory_id, activity_name) "
        f"VALUES (?, ?, ?, ?, ?, ?)",
        (monto, cat_name, str(fecha), cat_id, sub_id, actividad if actividad else None),
    )
    conn.commit()
    conn.close()
    if not IS_CLOUD:
        archivo = EXCEL_INGRESOS if tabla == 'ingresos' else EXCEL_GASTOS
        _guardar_excel_local(archivo, monto, cat_name, str(fecha),
                             subcategoria=sub_name, actividad=actividad)


def _formulario_registro(tabla, key, df_cats, df_subs, preset_cat=None):
    if df_cats.empty:
        st.info('No hay categorías disponibles. Crea una en "Gestionar categorías y subcategorías".')
        return
    nombres = df_cats['name'].tolist()
    idx_def = nombres.index(preset_cat) if (preset_cat in nombres) else 0
    monto = st.number_input('Monto', min_value=0.01, step=100.0, format='%.2f', key=f'{key}_monto')
    cat_sel = st.selectbox('Categoría', nombres, index=idx_def, key=f'{key}_cat')
    cat_id = int(df_cats[df_cats['name'] == cat_sel]['id'].iloc[0])
    subs = df_subs[df_subs['category_id'] == cat_id]
    sub_ops = ['(Sin subcategoría)'] + subs['name'].tolist()
    sub_sel = st.selectbox('Subcategoría (opcional)', sub_ops, key=f'{key}_sub')
    sub_id = int(subs[subs['name'] == sub_sel]['id'].iloc[0]) if sub_sel != '(Sin subcategoría)' else None
    actividad = st.text_input('Nombre de la actividad (opcional)', key=f'{key}_act')
    fecha = st.date_input('Fecha', value=date.today(), key=f'{key}_fecha')
    if st.button('💾 Guardar', key=f'{key}_btn'):
        _guardar_movimiento(tabla, float(monto), cat_sel, cat_id,
                            sub_sel if sub_id else None, sub_id,
                            actividad.strip(), fecha)
        st.success('Movimiento registrado con éxito.')


def _formulario_ahorro(df_cats, df_subs):
    if df_cats.empty:
        st.info('No hay categorías disponibles. Crea una en "Gestionar categorías y subcategorías".')
        return
    tipo_lbl = st.radio('Tipo', ['Depósito', 'Retiro'], horizontal=True, key='aho_tipo')
    tipo = 'deposito' if tipo_lbl == 'Depósito' else 'retiro'
    nombres = df_cats['name'].tolist()
    monto = st.number_input('Monto', min_value=0.01, step=100.0, format='%.2f', key='aho_monto')
    cat_sel = st.selectbox('Categoría', nombres, key='aho_cat')
    cat_id = int(df_cats[df_cats['name'] == cat_sel]['id'].iloc[0])
    subs = df_subs[df_subs['category_id'] == cat_id]
    sub_ops = ['(Sin subcategoría)'] + subs['name'].tolist()
    sub_sel = st.selectbox('Subcategoría (opcional)', sub_ops, key='aho_sub')
    sub_id = int(subs[subs['name'] == sub_sel]['id'].iloc[0]) if sub_sel != '(Sin subcategoría)' else None
    actividad = st.text_input('Nombre de la actividad (opcional)', key='aho_act')
    fecha = st.date_input('Fecha', value=date.today(), key='aho_fecha')
    if st.button('💾 Guardar', key='aho_btn'):
        if tipo == 'retiro':
            saldo_actual = obtener_saldo_ahorros()
            if float(monto) > saldo_actual:
                st.error(f'No puedes retirar ${float(monto):,.2f}. '
                         f'Saldo actual de ahorros: ${saldo_actual:,.2f}')
                return
        conn = get_conn()
        conn.execute(
            "INSERT INTO ahorros (monto, tipo, categoria, fecha, category_id, subcategory_id, activity_name) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (float(monto), tipo, cat_sel, str(fecha), cat_id, sub_id,
             actividad.strip() if actividad.strip() else None),
        )
        conn.commit()
        conn.close()
        if not IS_CLOUD:
            _guardar_excel_local_ahorros(float(monto), tipo, cat_sel, str(fecha),
                                         subcategoria=sub_sel if sub_id else None,
                                         actividad=actividad.strip())
        st.success(f'Ahorro ({tipo_lbl.lower()}) registrado con éxito.')


def _excel_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine='openpyxl')
    return buf.getvalue()


# Envío de correo (igual que main)
def enviar_reporte_email(df, recipient):
    user = os.getenv("EMAIL_USER") or _leer_secreto("EMAIL_USER")
    pwd  = os.getenv("EMAIL_PASS") or _leer_secreto("EMAIL_PASS")
    if not user or not pwd: st.error("Configura EMAIL_USER y EMAIL_PASS."); return
    csv=df.to_csv(index=False); msg=EmailMessage()
    msg['Subject']=f'Reporte Finanzas {date.today()}'; msg['From']=user; msg['To']=recipient
    msg.set_content('Adjunto reporte'); msg.add_attachment(csv,filename='reporte.csv',subtype='csv')
    try:
        with smtplib.SMTP('smtp.gmail.com',587) as smtp:
            smtp.starttls(); smtp.login(user,pwd); smtp.send_message(msg)
        st.success('Correo enviado')
    except Exception as e:
        st.error(f'Error envío: {e}')

# --- App Streamlit ---
_clave_tablas = f"tablas_ok::{usuario}"
if not st.session_state.get(_clave_tablas):
    crear_tablas_dashboard()
    st.session_state[_clave_tablas] = True

st.title('📊 Dashboard de Finanzas Personales')

# Categorías y subcategorías (compartidas por todas las secciones)
df_categorias = leer_df('SELECT id, name FROM categories ORDER BY name')
df_subcategorias = leer_df('SELECT id, category_id, name FROM subcategories ORDER BY name')
_cat_names_all  = df_categorias['name'].tolist()
_cat_ids_all    = df_categorias['id'].tolist()
_cat_id_map_all = dict(zip(_cat_names_all, _cat_ids_all))

# --- Registrar movimientos ---
with st.expander('➕ Registrar movimiento', expanded=False):
    tab_i, tab_g, tab_if, tab_gf, tab_a = st.tabs(
        ['Ingreso', 'Gasto', 'Ingreso Fijo', 'Gasto Fijo', 'Ahorro'])
    with tab_i:
        _formulario_registro('ingresos', 'reg_ing', df_categorias, df_subcategorias)
    with tab_g:
        _formulario_registro('gastos', 'reg_gas', df_categorias, df_subcategorias)
    with tab_if:
        _formulario_registro('ingresos', 'reg_ingf', df_categorias, df_subcategorias,
                             preset_cat=FIXED_INCOME_CAT)
    with tab_gf:
        _formulario_registro('gastos', 'reg_gasf', df_categorias, df_subcategorias,
                             preset_cat=FIXED_EXPENSE_CAT)
    with tab_a:
        _formulario_ahorro(df_categorias, df_subcategorias)

# --- Gestionar categorías y subcategorías ---
with st.expander('⚙️ Gestionar categorías y subcategorías'):
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.subheader('Categorías')
        nueva_cat = st.text_input('Nueva categoría', key='nueva_cat')
        if st.button('Agregar categoría', key='btn_add_cat'):
            if not nueva_cat.strip():
                st.error('El nombre no puede estar vacío.')
            else:
                conn = get_conn()
                try:
                    conn.execute('INSERT INTO categories (name) VALUES (?)', (nueva_cat.strip(),))
                    conn.commit()
                    conn.close()
                    st.rerun()
                except Exception:
                    conn.close()
                    st.error(f'La categoría "{nueva_cat.strip()}" ya existe.')
        if _cat_names_all:
            cat_del = st.selectbox('Eliminar categoría', _cat_names_all, key='cat_del_sel')
            st.caption('Las transacciones asociadas quedarán sin categoría vinculada.')
            if st.button('Eliminar categoría seleccionada', key='btn_del_cat'):
                _cat_del_id = int(_cat_id_map_all[cat_del])
                conn = get_conn()
                conn.execute('DELETE FROM subcategories WHERE category_id=?', (_cat_del_id,))
                conn.execute('DELETE FROM categories WHERE id=?', (_cat_del_id,))
                conn.commit()
                conn.close()
                st.rerun()
    with col_c2:
        st.subheader('Subcategorías')
        if not _cat_names_all:
            st.info('Primero crea una categoría.')
        else:
            cat_padre = st.selectbox('Categoría padre', _cat_names_all, key='sub_cat_padre')
            nueva_sub = st.text_input('Nueva subcategoría', key='nueva_sub')
            if st.button('Agregar subcategoría', key='btn_add_sub'):
                if not nueva_sub.strip():
                    st.error('El nombre no puede estar vacío.')
                else:
                    conn = get_conn()
                    conn.execute('INSERT INTO subcategories (category_id, name) VALUES (?, ?)',
                                 (int(_cat_id_map_all[cat_padre]), nueva_sub.strip()))
                    conn.commit()
                    conn.close()
                    st.rerun()
            if not df_subcategorias.empty:
                _df_subs_view = df_subcategorias.merge(
                    df_categorias.rename(columns={'id': 'category_id', 'name': 'categoria'}),
                    on='category_id', how='left')
                _sub_ids_view = _df_subs_view['id'].tolist()
                sub_del = st.selectbox(
                    'Eliminar subcategoría', _sub_ids_view,
                    format_func=lambda x: (
                        f"{_df_subs_view[_df_subs_view['id'] == x]['categoria'].iloc[0]} / "
                        f"{_df_subs_view[_df_subs_view['id'] == x]['name'].iloc[0]}"
                    ),
                    key='sub_del_sel')
                if st.button('Eliminar subcategoría seleccionada', key='btn_del_sub'):
                    conn = get_conn()
                    conn.execute('DELETE FROM subcategories WHERE id=?', (int(sub_del),))
                    conn.commit()
                    conn.close()
                    st.rerun()

df_ing, df_gas = cargar_datos()
df_ahorros = cargar_ahorros()
df_metas = cargar_metas_ahorro()

if st.button("🔄 Recargar datos"):
    st.rerun()

# Periodo y filtros
years=sorted(pd.concat([df_ing['fecha'].dt.year, df_gas['fecha'].dt.year]).dropna().unique())
if len(years) == 0:
    years = [date.today().year]
selected_year = int(st.sidebar.selectbox('Año',years,index=len(years)-1))
selected_month = int(st.sidebar.selectbox('Mes',list(range(1,13)),index=date.today().month-1))

# Filtrar
df_i=df_ing[(df_ing['fecha'].dt.year==selected_year)&(df_ing['fecha'].dt.month==selected_month)]
df_g=df_gas[(df_gas['fecha'].dt.year==selected_year)&(df_gas['fecha'].dt.month==selected_month)]

# KPIs
tot_i=df_i['monto'].sum(); tot_g=df_g['monto'].sum(); balance_mes_actual=tot_i-tot_g

# Obtener balance del mes anterior
balance_mes_anterior = obtener_balance_mes_anterior(selected_year, selected_month)
balance_total = balance_mes_actual + balance_mes_anterior

fixed_i=df_i[df_i['categoria'].str.startswith(FIXED_INCOME_CAT)]['monto'].sum()
fixed_g=df_g[df_g['categoria'].str.startswith(FIXED_EXPENSE_CAT)]['monto'].sum()

depositos_total = df_ahorros[df_ahorros["tipo"] == "deposito"]["monto"].sum() if not df_ahorros.empty else 0.0
retiros_total   = df_ahorros[df_ahorros["tipo"] == "retiro"]["monto"].sum()   if not df_ahorros.empty else 0.0
saldo_ahorros   = depositos_total - retiros_total
balance_disponible = balance_total - saldo_ahorros

col1,col2,col3,col4,col5,col6,col7=st.columns(7)
col1.metric('Ingresos Totales',f'${tot_i:,.2f}')
col2.metric('Gastos Totales',f'${tot_g:,.2f}')
col3.metric('Balance del Mes',f'${balance_mes_actual:,.2f}')
col4.metric('Balance Mes Anterior',f'${balance_mes_anterior:,.2f}')
col5.metric('Balance Total',f'${balance_total:,.2f}', delta=f'${balance_mes_actual:,.2f}')
col6.metric('Ingresos Fijos',f'${fixed_i:,.2f}')
col7.metric('Balance Disponible',f'${balance_disponible:,.2f}')

st.markdown('---')
col1_extra, col2_extra, col3_extra = st.columns(3)
col1_extra.metric('Gastos Fijos',f'${fixed_g:,.2f}')
col3_extra.metric('Saldo Ahorros', f'${saldo_ahorros:,.2f}',
                  delta=f'Dep: ${depositos_total:,.2f} | Ret: ${retiros_total:,.2f}',
                  delta_color="normal")

# Sección para guardar balance del mes actual
with col2_extra:
    st.subheader('💾 Guardar Balance')
    st.write(f'**Balance Total del mes actual ({selected_year}-{selected_month:02d}):** ${balance_total:,.2f}')
    if st.button('Guardar Balance Total del Mes Actual', key='guardar_balance'):
        guardar_balance_mensual(selected_year, selected_month, balance_total)
        st.success(f'Balance Total de {selected_year}-{selected_month:02d} guardado: ${balance_total:,.2f}')
        st.rerun()
st.markdown('---')

# Tablas fijos
st.header(f'📋 Ingresos fijos mensuales - {selected_year}-{selected_month:02d}')
_fi = df_i[df_i['categoria'].str.startswith(FIXED_INCOME_CAT)].copy()
if 'activity_name' in _fi.columns:
    _fi['Categoría'] = _fi['categoria'].str.cat(
        _fi['activity_name'].fillna('').astype(str).str.strip(), sep=' - '
    ).str.rstrip(' - ')
else:
    _fi['Categoría'] = _fi['categoria']
st.table(_fi[['id','monto','Categoría','fecha']].rename(columns={'id':'ID','monto':'Monto','fecha':'Fecha'}))
st.markdown('---')
st.header(f'📋 Gastos fijos mensuales - {selected_year}-{selected_month:02d}')
_fg = df_g[df_g['categoria'].str.startswith(FIXED_EXPENSE_CAT)].copy()
if 'activity_name' in _fg.columns:
    _fg['Categoría'] = _fg['categoria'].str.cat(
        _fg['activity_name'].fillna('').astype(str).str.strip(), sep=' - '
    ).str.rstrip(' - ')
else:
    _fg['Categoría'] = _fg['categoria']
st.table(_fg[['id','monto','Categoría','fecha']].rename(columns={'id':'ID','monto':'Monto','fecha':'Fecha'}))
st.markdown('---')

# Presupuesto vs Gasto
st.header(f'📊 Presupuesto vs Gasto - {selected_year}-{selected_month:02d}')
df_budgets = cargar_presupuestos()

with st.expander('⚙️ Gestionar presupuestos'):
    _cat_names_b = _cat_names_all
    with st.form('form_nuevo_presupuesto'):
        _col_b1, _col_b2 = st.columns(2)
        with _col_b1:
            _cat_pres = st.selectbox('Categoría', _cat_names_b, key='pres_cat')
        with _col_b2:
            _monto_pres = st.number_input('Monto ($)', min_value=0.01, step=1000.0, format='%.2f', key='pres_monto')
        if st.form_submit_button('Guardar presupuesto'):
            _conn_bp = get_conn()
            _conn_bp.execute('INSERT OR REPLACE INTO budgets (categoria, monto) VALUES (?, ?)', (_cat_pres, float(_monto_pres)))
            _conn_bp.commit(); _conn_bp.close()
            st.success(f'Presupuesto para "{_cat_pres}" guardado.')
            st.rerun()
    if not df_budgets.empty:
        st.dataframe(
            df_budgets.rename(columns={'id': 'ID', 'categoria': 'Categoría', 'monto': 'Presupuesto ($)'}),
            use_container_width=True
        )
        _id_del_b = st.selectbox(
            'Eliminar', df_budgets['id'].tolist(),
            format_func=lambda x: df_budgets[df_budgets['id'] == x]['categoria'].iloc[0],
            key='del_budget_id'
        )
        if st.button('Eliminar presupuesto seleccionado', key='btn_del_budget'):
            _conn_bd = get_conn()
            _conn_bd.execute('DELETE FROM budgets WHERE id=?', (int(_id_del_b),))
            _conn_bd.commit(); _conn_bd.close()
            st.rerun()

if df_budgets.empty:
    st.info('No hay presupuestos configurados. Ábrelos desde "Gestionar presupuestos" arriba.')
else:
    for _, _brow in df_budgets.iterrows():
        _cat  = _brow['categoria']
        _bud  = float(_brow['monto'])
        _spent = df_g[df_g['categoria'] == _cat]['monto'].sum()
        _pct  = _spent / _bud * 100 if _bud else 0
        st.write(f'**{_cat}**: ${_spent:,.2f} / ${_bud:,.2f} ({_pct:.1f}%)')
        st.progress(min(int(_pct), 100))
st.markdown('---')

# Benchmark histórico
st.subheader('¿Qué es Benchmark Histórico Mensual?')
st.write('Compara tus gastos mes a mes dentro del año seleccionado para identificar tendencias.')
st.header(f'📈 Evolución Gastos {selected_year}')
g_mes=df_gas[df_gas['fecha'].dt.year==selected_year].set_index('fecha').resample('ME')['monto'].sum()
g_mes.index=g_mes.index.month; st.bar_chart(g_mes)
st.markdown('---')

# Balance neto mensual
st.header(f'📈 Balance Neto por Mes — {selected_year}')
_meses_nombres = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
_i_año = df_ing[df_ing['fecha'].dt.year == selected_year]
_g_año = df_gas[df_gas['fecha'].dt.year == selected_year]
_netos = [
    _i_año[_i_año['fecha'].dt.month == m]['monto'].sum()
    - _g_año[_g_año['fecha'].dt.month == m]['monto'].sum()
    for m in range(1, 13)
]
_df_neto = pd.DataFrame({'Mes': _meses_nombres, 'Balance Neto': _netos})
_fig_neto = px.line(_df_neto, x='Mes', y='Balance Neto', markers=True,
                    color_discrete_sequence=['#2196F3'])
_fig_neto.add_hline(y=0, line_dash='dash', line_color='red', opacity=0.5)
_fig_neto.update_layout(yaxis_title='Balance Neto ($)', xaxis_title='')
st.plotly_chart(_fig_neto, use_container_width=True)
st.markdown('---')

# Pie charts
st.header(f'🍰 Distribución de Gastos {selected_year}-{selected_month:02d}')
if not df_g.empty:
    fig_g = px.pie(
        df_g,
        names='categoria',
        values='monto',
        hole=0
    )
    fig_g.update_traces(
        textinfo='label+percent',
        hovertemplate='%{label}: $%{value:,.2f} (%{percent})<extra></extra>'
    )
    fig_g.update_layout(
        showlegend=True,
        margin=dict(l=0, r=0, t=0, b=0)
    )
    # Configurar Plotly sin argumentos deprecados
    plotly_config = {
        'displayModeBar': False,
        'displaylogo': False
    }
    st.plotly_chart(fig_g, width='stretch', config=plotly_config, use_container_width=False)
else:
    st.info('No hay datos de gastos para este período')
st.markdown('---')
st.header(f'🍰 Distribución de Ingresos {selected_year}-{selected_month:02d}')
if not df_i.empty:
    fig_i = px.pie(
        df_i,
        names='categoria',
        values='monto',
        hole=0
    )
    fig_i.update_traces(
        textinfo='label+percent',
        hovertemplate='%{label}: $%{value:,.2f} (%{percent})<extra></extra>'
    )
    fig_i.update_layout(
        showlegend=True,
        margin=dict(l=0, r=0, t=0, b=0)
    )
    # Configurar Plotly sin argumentos deprecados
    plotly_config = {
        'displayModeBar': False,
        'displaylogo': False
    }
    st.plotly_chart(fig_i, width='stretch', config=plotly_config, use_container_width=False)
else:
    st.info('No hay datos de ingresos para este período')
st.markdown('---')

# --- Detalle y eliminación individual ---
st.header('📋 Detalle de Transacciones')
df_detail = pd.concat([df_i.assign(Tipo='Ingreso'),
                       df_g.assign(Tipo='Gasto')])
st.dataframe(df_detail.sort_values(by='fecha', ascending=False))

ids = st.multiselect('Eliminar IDs (Ingresos/Gastos)', df_detail['id'].tolist())
if st.button('Eliminar seleccionados', key='del_ids_dashboard'):
    conn = get_conn()
    cur  = conn.cursor()
    for i in ids:
        tabla = 'ingresos' if df_detail[df_detail['id']==i]['Tipo'].iloc[0]=='Ingreso' else 'gastos'
        cur.execute(f"DELETE FROM {tabla} WHERE id=?", (int(i),))
    conn.commit()
    conn.close()
    st.success('Registros eliminados')
    st.rerun()

st.markdown('---')

# --- Editar Transacción ---
st.header('✏️ Editar Transacción')
_tipo_edit = st.selectbox('Tipo', ['Ingreso', 'Gasto'], key='edit_tipo')
_df_edit_src = (df_ing if _tipo_edit == 'Ingreso' else df_gas) \
               .sort_values('fecha', ascending=False).reset_index(drop=True)
_tabla_edit = 'ingresos' if _tipo_edit == 'Ingreso' else 'gastos'

if _df_edit_src.empty:
    st.info('No hay transacciones registradas.')
else:
    def _edit_label(r):
        f = r['fecha'].strftime('%Y-%m-%d') if pd.notna(r['fecha']) else '?'
        return f"ID {int(r['id'])} — {f} — ${float(r['monto']):,.2f} — {r.get('categoria', '') or ''}"

    _opciones_edit = [_edit_label(_df_edit_src.iloc[i]) for i in range(len(_df_edit_src))]
    _idx_sel = st.selectbox('Transacción a editar', range(len(_opciones_edit)),
                             format_func=lambda i: _opciones_edit[i], key='edit_idx')
    _row = _df_edit_src.iloc[_idx_sel]

    _cat_names_edit  = _cat_names_all
    _cat_id_map_edit = _cat_id_map_all
    _cat_actual  = str(_row.get('categoria', '') or '')
    _cat_idx_def = _cat_names_edit.index(_cat_actual) if _cat_actual in _cat_names_edit else 0

    with st.form('form_editar'):
        _col_e1, _col_e2 = st.columns(2)
        with _col_e1:
            _nuevo_monto = st.number_input('Monto', value=float(_row['monto']),
                                           min_value=0.01, step=100.0, format='%.2f')
            _nueva_cat   = st.selectbox('Categoría', _cat_names_edit, index=_cat_idx_def)
        with _col_e2:
            _fecha_def   = _row['fecha'].date() if pd.notna(_row['fecha']) else date.today()
            _nueva_fecha = st.date_input('Fecha', value=_fecha_def)
            _act_val     = _row.get('activity_name', None)
            _nueva_act   = st.text_input(
                'Nombre de la actividad',
                value=str(_act_val) if _act_val is not None and pd.notna(_act_val) else ''
            )
        if st.form_submit_button('💾 Guardar cambios'):
            _nuevo_cat_id = _cat_id_map_edit.get(_nueva_cat)
            _conn_upd = get_conn()
            _conn_upd.execute(
                f'UPDATE {_tabla_edit} SET monto=?, categoria=?, fecha=?, category_id=?, activity_name=? WHERE id=?',
                (float(_nuevo_monto), _nueva_cat, str(_nueva_fecha),
                 int(_nuevo_cat_id) if _nuevo_cat_id is not None else None,
                 _nueva_act.strip() or None, int(_row['id']))
            )
            _conn_upd.commit(); _conn_upd.close()
            st.success(f'{_tipo_edit} ID {int(_row["id"])} actualizado.')
            st.rerun()

st.markdown('---')

# --- Historial de Ahorros ---
st.header('💰 Historial de Ahorros')
if df_ahorros.empty:
    st.info('No hay movimientos de ahorros registrados.')
else:
    cols_show = [c for c in ('id', 'monto', 'tipo', 'categoria', 'activity_name', 'fecha') if c in df_ahorros.columns]
    df_ahorros_display = df_ahorros[cols_show].sort_values('fecha', ascending=False).rename(columns={
        'id': 'ID', 'monto': 'Monto', 'tipo': 'Tipo',
        'categoria': 'Categoría', 'activity_name': 'Nombre de la actividad', 'fecha': 'Fecha'
    })
    st.dataframe(df_ahorros_display, use_container_width=True)

st.markdown('---')

# --- Metas de Ahorro ---
st.header('🎯 Metas de Ahorro')

st.subheader('Crear nueva meta')
cat_names  = _cat_names_all
cat_id_map = _cat_id_map_all

if not cat_names:
    st.info('No hay categorías disponibles. Crea una en "Gestionar categorías y subcategorías".')
else:
    with st.form('form_nueva_meta'):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            nombre_meta = st.text_input('Nombre de la meta', placeholder='Ej: Especialización')
            cat_sel = st.selectbox('Categoría', cat_names, key='meta_cat')
        with col_f2:
            monto_obj = st.number_input('Monto objetivo ($)', min_value=0.01, step=1000.0, format='%.2f')

        cat_id_sel = int(cat_id_map.get(cat_sel))
        _subs_meta = df_subcategorias[df_subcategorias['category_id'] == cat_id_sel]

        sub_names  = ['(Sin subcategoría)'] + _subs_meta['name'].tolist()
        sub_id_map = {'(Sin subcategoría)': None,
                      **dict(zip(_subs_meta['name'].tolist(), _subs_meta['id'].tolist()))}
        sub_sel    = st.selectbox('Subcategoría (opcional)', sub_names, key='meta_sub')

        if st.form_submit_button('Guardar meta'):
            if not nombre_meta.strip():
                st.error('El nombre de la meta no puede estar vacío.')
            elif monto_obj <= 0:
                st.error('El monto objetivo debe ser mayor a 0.')
            else:
                sub_id_val = sub_id_map.get(sub_sel)
                conn_ins = get_conn()
                conn_ins.execute(
                    'INSERT INTO metas_ahorro (category_id, subcategory_id, nombre, monto_objetivo) VALUES (?,?,?,?)',
                    (cat_id_sel, sub_id_val, nombre_meta.strip(), float(monto_obj)))
                conn_ins.commit(); conn_ins.close()
                st.success(f'Meta "{nombre_meta.strip()}" guardada.')
                st.rerun()

st.subheader('Progreso de metas')
df_metas = cargar_metas_ahorro()
if df_metas.empty:
    st.info('No hay metas registradas. Crea una arriba.')
else:
    for _, meta in df_metas.iterrows():
        ahorrado = calcular_ahorrado_meta(int(meta['category_id']), meta['subcategory_id'])
        objetivo = float(meta['monto_objetivo'])
        pct      = (ahorrado / objetivo * 100) if objetivo > 0 else 0.0
        label_sub = f" / {meta['subcategoria_nombre']}" \
                    if pd.notna(meta.get('subcategoria_nombre')) and meta['subcategoria_nombre'] else ''
        st.markdown(f"**{meta['nombre']}** — {meta['categoria_nombre']}{label_sub}")
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            st.progress(min(int(pct), 100))
        with col_p2:
            st.metric('Ahorrado / Objetivo', f'${ahorrado:,.2f}',
                      delta=f'Meta: ${objetivo:,.2f} ({pct:.1f}%)', delta_color='normal')
        if st.button(f'Eliminar meta #{int(meta["id"])}', key=f'del_meta_{int(meta["id"])}'):
            conn_del = get_conn()
            conn_del.execute('DELETE FROM metas_ahorro WHERE id=?', (int(meta['id']),))
            conn_del.commit(); conn_del.close()
            st.success(f'Meta "{meta["nombre"]}" eliminada.')
            st.rerun()
        st.markdown('---')

st.markdown('---')

# --- Eliminación por categoría ---
st.header('🗑️ Eliminar Transacciones por Categoría')
section = st.selectbox('Tipo', ['Ingreso','Gasto'], key='cat_type_dashboard')
df_sel = df_i if section=='Ingreso' else df_g
cat_list = sorted(df_sel['categoria'].unique())
cat_sel  = st.selectbox('Categoría', cat_list, key='cat_sel_dashboard')

filtered = df_sel[df_sel['categoria']==cat_sel]
st.dataframe(
    filtered[['id','monto','categoria','fecha']]
    .rename(columns={'id':'ID','monto':'Monto','categoria':'Categoría','fecha':'Fecha'})
)

ids_cat = st.multiselect('Seleccionar IDs a eliminar', filtered['id'].tolist(), key='cat_ids_dashboard')
if st.button('Eliminar seleccionados', key='del_cat_dashboard'):
    conn = get_conn()
    cur  = conn.cursor()
    tabla = 'ingresos' if section=='Ingreso' else 'gastos'
    for i in ids_cat:
        cur.execute(f"DELETE FROM {tabla} WHERE id=?", (int(i),))
    conn.commit()
    conn.close()
    st.success(f"Registros eliminados de la categoría {cat_sel}")
    st.rerun()

st.markdown('---')

# --- Gestión de Balances Mensuales ---
st.header('💰 Gestión de Balances Mensuales')
df_balances = obtener_balances_guardados()
if not df_balances.empty:
    st.subheader('Balances Guardados')
    # Asegurar que los tipos sean correctos antes de procesar
    df_balances['año'] = pd.to_numeric(df_balances['año'], errors='coerce').fillna(0).astype(int)
    df_balances['mes'] = pd.to_numeric(df_balances['mes'], errors='coerce').fillna(0).astype(int)
    df_balances['balance'] = pd.to_numeric(df_balances['balance'], errors='coerce').fillna(0.0).astype(float)

    df_display = df_balances[['año', 'mes', 'balance', 'fecha_creacion']].copy()
    df_display['mes'] = df_display['mes'].apply(lambda x: f'{int(x):02d}')
    df_display = df_display.rename(columns={
        'año': 'Año',
        'mes': 'Mes',
        'balance': 'Balance',
        'fecha_creacion': 'Fecha de Creación'
    })
    df_display['Balance'] = df_display['Balance'].apply(lambda x: f'${x:,.2f}')
    st.dataframe(df_display, width='stretch')

    # Opción para eliminar balances
    st.subheader('Eliminar Balance')
    if not df_balances.empty:
        # Los tipos ya están convertidos arriba, pero por seguridad los verificamos de nuevo
        # Crear opciones con índice para mejor identificación
        opciones_balance = []
        indices_balance = []
        for idx, row in df_balances.iterrows():
            año_val = int(row['año'])
            mes_val = int(row['mes'])
            balance_val = float(row['balance'])
            # Mostrar ID también para identificación única
            id_val = row.get('id', idx)
            opcion = f"ID {id_val}: {año_val}-{mes_val:02d} (${balance_val:,.2f})"
            opciones_balance.append(opcion)
            indices_balance.append((año_val, mes_val, id_val))

        balance_eliminar = st.selectbox('Seleccionar balance a eliminar', opciones_balance, key='eliminar_balance')

        if st.button('Eliminar Balance Seleccionado', key='btn_eliminar_balance'):
            # Encontrar el índice seleccionado
            idx_seleccionado = opciones_balance.index(balance_eliminar)
            año_elim, mes_elim, id_elim = indices_balance[idx_seleccionado]

            conn = get_conn()
            cur = conn.cursor()
            # Intentar eliminar por año y mes
            cur.execute('DELETE FROM balances_mensuales WHERE año=? AND mes=?', (año_elim, mes_elim))
            rows_deleted = cur.rowcount

            # Si no se eliminó, intentar por ID directamente
            if rows_deleted == 0:
                cur.execute('DELETE FROM balances_mensuales WHERE id=?', (int(id_elim),))
                rows_deleted = cur.rowcount

            # Si aún no funciona, intentar con CAST para manejar tipos de datos problemáticos
            if rows_deleted == 0:
                cur.execute('DELETE FROM balances_mensuales WHERE CAST(año AS INTEGER)=? AND CAST(mes AS INTEGER)=?', (año_elim, mes_elim))
                rows_deleted = cur.rowcount

            conn.commit()
            conn.close()

            if rows_deleted > 0:
                st.success(f'Balance de {año_elim}-{mes_elim:02d} (ID: {id_elim}) eliminado correctamente')
            else:
                st.error(f'No se pudo eliminar el balance de {año_elim}-{mes_elim:02d}. Intenta usar el botón de eliminar año 0.')
            st.rerun()

        # Opción adicional para eliminar registros con año 0 (inválidos)
        balances_año_0 = [idx for idx, row in df_balances.iterrows() if int(row['año']) == 0]
        if balances_año_0:
            st.warning('⚠️ Se detectaron balances con año 0 (inválidos)')
            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:
                if st.button('🗑️ Eliminar todos los balances con año 0', key='eliminar_año_0'):
                    conn = get_conn()
                    cur = conn.cursor()
                    rows_deleted = 0

                    # Intentar múltiples métodos de eliminación
                    try:
                        # Método 1: Por año=0
                        cur.execute('DELETE FROM balances_mensuales WHERE año=0')
                        rows_deleted += cur.rowcount
                    except:
                        pass

                    try:
                        # Método 2: Por año IS NULL
                        cur.execute('DELETE FROM balances_mensuales WHERE año IS NULL')
                        rows_deleted += cur.rowcount
                    except:
                        pass

                    try:
                        # Método 3: Por CAST
                        cur.execute("DELETE FROM balances_mensuales WHERE CAST(año AS INTEGER)=0")
                        rows_deleted += cur.rowcount
                    except:
                        pass

                    # Método 4: Eliminar por ID directamente si sabemos los IDs
                    for idx in balances_año_0:
                        try:
                            id_val = df_balances.loc[idx, 'id']
                            cur.execute('DELETE FROM balances_mensuales WHERE id=?', (int(id_val),))
                            rows_deleted += cur.rowcount
                        except:
                            pass

                    conn.commit()
                    conn.close()

                    if rows_deleted > 0:
                        st.success(f'Se eliminaron {rows_deleted} balance(s) con año inválido')
                    else:
                        st.error('No se pudo eliminar. Verifica manualmente en la base de datos.')
                    st.rerun()

            with col_btn2:
                # Botón para eliminar todos los registros (último recurso)
                if st.button('⚠️ Eliminar TODOS los balances (Cuidado)', key='eliminar_todos'):
                    conn = get_conn()
                    cur = conn.cursor()
                    cur.execute('SELECT COUNT(*) FROM balances_mensuales')
                    total = cur.fetchone()[0]
                    cur.execute('DELETE FROM balances_mensuales')
                    conn.commit()
                    conn.close()
                    st.success(f'Se eliminaron {total} balance(s) de la base de datos')
                    st.rerun()
else:
    st.info('No hay balances guardados aún. Usa el botón "Guardar Balance del Mes Actual" para guardar balances.')

st.markdown('---')

# Reportes
st.header('📑 Reportes')
csv_data=df_detail.to_csv(index=False).encode()
st.download_button('📥 Descargar CSV',csv_data,file_name=f'reporte_{selected_year}_{selected_month}.csv',mime='text/csv',key='download_dashboard')

with st.expander('📦 Exportar a Excel'):
    st.caption('Genera los archivos Excel con todos los movimientos registrados.')
    if st.button('Generar archivos Excel', key='gen_excel'):
        _sql_export = '''
            SELECT t.monto AS "Monto", t.categoria AS "Categoría", t.fecha AS "Fecha",
                   COALESCE(s.name, '') AS "Subcategoría",
                   COALESCE(t.activity_name, '') AS "Nombre de la actividad"
            FROM {tabla} t
            LEFT JOIN subcategories s ON s.id = t.subcategory_id
            ORDER BY t.id'''
        st.session_state['excel_ing'] = _excel_bytes(leer_df(_sql_export.format(tabla='ingresos')))
        st.session_state['excel_gas'] = _excel_bytes(leer_df(_sql_export.format(tabla='gastos')))
        st.session_state['excel_aho'] = _excel_bytes(leer_df('''
            SELECT a.monto AS "Monto", a.tipo AS "Tipo", a.categoria AS "Categoría",
                   COALESCE(s.name, '') AS "Subcategoría",
                   COALESCE(a.activity_name, '') AS "Nombre de la actividad", a.fecha AS "Fecha"
            FROM ahorros a
            LEFT JOIN subcategories s ON s.id = a.subcategory_id
            ORDER BY a.id'''))
    _mime_xlsx = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    if 'excel_ing' in st.session_state:
        _col_x1, _col_x2, _col_x3 = st.columns(3)
        _col_x1.download_button('📥 finanzas.xlsx', st.session_state['excel_ing'],
                                file_name='finanzas.xlsx', mime=_mime_xlsx, key='dl_xl_ing')
        _col_x2.download_button('📥 finanzas_gastos.xlsx', st.session_state['excel_gas'],
                                file_name='finanzas_gastos.xlsx', mime=_mime_xlsx, key='dl_xl_gas')
        _col_x3.download_button('📥 finanzas_ahorros.xlsx', st.session_state['excel_aho'],
                                file_name='finanzas_ahorros.xlsx', mime=_mime_xlsx, key='dl_xl_aho')

if st.button('✉️ Enviar por correo',key='send_email_dashboard'):
    enviar_reporte_email(df_detail,'jmohcm@gmail.com')

st.sidebar.caption('Ejecutar con: streamlit run dashboard.py')

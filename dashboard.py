import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import date, datetime
import plotly.express as px
import smtplib
from email.message import EmailMessage


st.set_page_config(page_title='Dashboard Financiero', layout='wide')
# --- Selección de usuario ---
USER_ROOT = os.path.join(os.getcwd(), "users")
usuarios = [d for d in os.listdir(USER_ROOT) if os.path.isdir(os.path.join(USER_ROOT, d))]
if not usuarios:
    st.error("No hay usuarios. Ejecuta main.py primero.")
    st.stop()
usuario = st.sidebar.selectbox("Usuario", usuarios)

# Rutas por usuario
BASE_DIR = os.path.join(USER_ROOT, usuario)
DB_FILE = os.path.join(BASE_DIR, "finanzas.db")
EXCEL_INGRESOS = os.path.join(BASE_DIR, "finanzas.xlsx")
EXCEL_GASTOS   = os.path.join(BASE_DIR, "finanzas_gastos.xlsx")

# Categorías fijas y presupuesto
FIXED_INCOME_CAT = "Ingresos fijos mensuales"
FIXED_EXPENSE_CAT= "Gastos fijos mensuales"
BUDGETS = {"comida":100000, "Transporte":50000, "Entretenimiento":100000}

# Crear tablas si no existen
def crear_tablas_dashboard():
    conn=sqlite3.connect(DB_FILE)
    cur=conn.cursor()
    cur.execute('CREATE TABLE IF NOT EXISTS ingresos(id INTEGER PRIMARY KEY, monto REAL, categoria TEXT, fecha TEXT)')
    cur.execute('CREATE TABLE IF NOT EXISTS gastos(id INTEGER PRIMARY KEY, monto REAL, categoria TEXT, fecha TEXT)')
    cur.execute('CREATE TABLE IF NOT EXISTS balances_mensuales(id INTEGER PRIMARY KEY AUTOINCREMENT, año INTEGER, mes INTEGER, balance REAL, fecha_creacion TEXT, UNIQUE(año, mes))')
    cur.execute('CREATE TABLE IF NOT EXISTS categories(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, created_at TEXT DEFAULT CURRENT_TIMESTAMP)')
    cur.execute('CREATE TABLE IF NOT EXISTS subcategories(id INTEGER PRIMARY KEY AUTOINCREMENT, category_id INTEGER NOT NULL REFERENCES categories(id), name TEXT NOT NULL, created_at TEXT DEFAULT CURRENT_TIMESTAMP)')
    cur.execute('''CREATE TABLE IF NOT EXISTS ahorros (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        monto          REAL    NOT NULL,
        tipo           TEXT    NOT NULL CHECK(tipo IN ('deposito', 'retiro')),
        categoria      TEXT,
        category_id    INTEGER REFERENCES categories(id),
        subcategory_id INTEGER REFERENCES subcategories(id),
        activity_name  TEXT,
        fecha          TEXT    NOT NULL
    )''')
    conn.commit();conn.close()

# Funciones para manejar balances mensuales
def obtener_balance_mes_anterior(año, mes):
    """Obtiene el balance del mes anterior"""
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    
    # Calcular mes y año anterior
    if mes == 1:
        mes_anterior = 12
        año_anterior = año - 1
    else:
        mes_anterior = mes - 1
        año_anterior = año
    
    cur.execute('SELECT balance FROM balances_mensuales WHERE año=? AND mes=?', (año_anterior, mes_anterior))
    result = cur.fetchone()
    conn.close()
    return result[0] if result else 0.0

def guardar_balance_mensual(año, mes, balance):
    """Guarda o actualiza el balance de un mes específico"""
    # Asegurar que año y mes sean enteros
    año = int(año) if año is not None else date.today().year
    mes = int(mes) if mes is not None else date.today().month
    
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    fecha_creacion = str(date.today())
    cur.execute('''
        INSERT OR REPLACE INTO balances_mensuales (año, mes, balance, fecha_creacion)
        VALUES (?, ?, ?, ?)
    ''', (año, mes, balance, fecha_creacion))
    conn.commit()
    conn.close()

def obtener_balances_guardados():
    """Obtiene todos los balances guardados"""
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql('SELECT * FROM balances_mensuales ORDER BY año DESC, mes DESC', conn)
    conn.close()
    # Convertir tipos de datos explícitamente para evitar problemas con bytes
    if not df.empty:
        df['año'] = pd.to_numeric(df['año'], errors='coerce').astype('Int64')
        df['mes'] = pd.to_numeric(df['mes'], errors='coerce').astype('Int64')
        df['balance'] = pd.to_numeric(df['balance'], errors='coerce').astype('float64')
    return df

def cargar_datos(db_path: str):
    conn = sqlite3.connect(db_path)
    df_i = pd.read_sql("SELECT * FROM ingresos", conn, parse_dates=["fecha"])
    df_g = pd.read_sql("SELECT * FROM gastos",  conn, parse_dates=["fecha"])
    conn.close()
    df_i["fecha"] = pd.to_datetime(df_i["fecha"], errors="coerce")
    df_g["fecha"] = pd.to_datetime(df_g["fecha"], errors="coerce")
    return df_i, df_g


def cargar_ahorros(db_path: str):
    conn = sqlite3.connect(db_path)
    df_a = pd.read_sql("SELECT * FROM ahorros", conn, parse_dates=["fecha"])
    conn.close()
    return df_a

# Envío de correo (igual que main)
def enviar_reporte_email(df, recipient):
    user=os.getenv("EMAIL_USER"); pwd=os.getenv("EMAIL_PASS")
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
crear_tablas_dashboard()

st.title('📊 Dashboard de Finanzas Personales')

df_ing, df_gas = cargar_datos(DB_FILE)
df_ahorros = cargar_ahorros(DB_FILE)

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

col1,col2,col3,col4,col5,col6=st.columns(6)
col1.metric('Ingresos Totales',f'${tot_i:,.2f}')
col2.metric('Gastos Totales',f'${tot_g:,.2f}')
col3.metric('Balance del Mes',f'${balance_mes_actual:,.2f}')
col4.metric('Balance Mes Anterior',f'${balance_mes_anterior:,.2f}')
col5.metric('Balance Total',f'${balance_total:,.2f}', delta=f'${balance_mes_actual:,.2f}')
col6.metric('Ingresos Fijos',f'${fixed_i:,.2f}')

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
for cat,bud in BUDGETS.items():
    spent=df_g[df_g['categoria']==cat]['monto'].sum(); pct=spent/bud*100 if bud else 0
    st.write(f'**{cat}**: ${spent:,.2f}/${bud:,.2f} ({pct:.1f}%)'); st.progress(min(int(pct),100))
st.markdown('---')

# Benchmark histórico
st.subheader('¿Qué es Benchmark Histórico Mensual?')
st.write('Compara tus gastos mes a mes dentro del año seleccionado para identificar tendencias.')
st.header(f'📈 Evolución Gastos {selected_year}')
g_mes=df_gas[df_gas['fecha'].dt.year==selected_year].set_index('fecha').resample('ME')['monto'].sum()
g_mes.index=g_mes.index.month; st.bar_chart(g_mes)
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
    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()
    for i in ids:
        tabla = 'ingresos' if df_detail[df_detail['id']==i]['Tipo'].iloc[0]=='Ingreso' else 'gastos'
        cur.execute(f"DELETE FROM {tabla} WHERE id=?", (i,))
    conn.commit()
    conn.close()
    st.success('Registros eliminados')
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
    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()
    tabla = 'ingresos' if section=='Ingreso' else 'gastos'
    for i in ids_cat:
        cur.execute(f"DELETE FROM {tabla} WHERE id=?", (i,))
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
            
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            # Intentar eliminar por año y mes
            cur.execute('DELETE FROM balances_mensuales WHERE año=? AND mes=?', (año_elim, mes_elim))
            rows_deleted = cur.rowcount
            
            # Si no se eliminó, intentar por ID directamente
            if rows_deleted == 0:
                cur.execute('DELETE FROM balances_mensuales WHERE id=?', (id_elim,))
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
                    conn = sqlite3.connect(DB_FILE)
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
                    conn = sqlite3.connect(DB_FILE)
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
if st.button('✉️ Enviar por correo',key='send_email_dashboard'):
    enviar_reporte_email(df_detail,'jmohcm@gmail.com')

st.sidebar.caption('Ejecutar con: streamlit run dashboard.py')
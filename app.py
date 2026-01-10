import streamlit as st
import pandas as pd
import io
import xlsxwriter
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN ROBUSTA
# ==========================================
st.set_page_config(page_title="Sistema Inventario PRO", page_icon="🏢", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    /* Estética profesional */
    html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
    
    /* El Panel de Control (Arriba) */
    .control-panel {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #d1d5db;
        margin-bottom: 20px;
    }
    
    /* Tabla de Solo Lectura (Abajo) - Más limpia */
    .stDataFrame { width: 100%; }
    
    /* Métricas grandes */
    div[data-testid="stMetricValue"] { font-size: 28px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. LÓGICA DE NEGOCIO
# ==========================================
def cargar_datos(file):
    try: df = pd.read_excel(file, header=None)
    except:
        file.seek(0)
        try: df = pd.read_csv(file, header=None, sep=None, engine='python', encoding='latin-1')
        except: return None
            
    start_row = -1
    for index, row in df.iterrows():
        s_row = str(row.values)
        if "Producto" in s_row and "Cantidad" in s_row:
            start_row = index
            break
    
    if start_row == -1: return None
    
    df_data = df.iloc[start_row + 1:].copy()
    try: df_clean = df_data.iloc[:, [0, 2, 8, 10]]
    except: df_clean = df_data.dropna(axis=1, how='all').iloc[:, :4]

    df_clean.columns = ['codigo', 'descripcion', 'requerido', 'unidad']
    df_clean['requerido'] = pd.to_numeric(df_clean['requerido'], errors='coerce').fillna(0)
    df_clean = df_clean[df_clean['requerido'] > 0]
    df_clean['descripcion'] = df_clean['descripcion'].astype(str).str.strip()
    df_clean['unidad'] = df_clean['unidad'].astype(str).str.strip().fillna('').str.upper()
    df_clean['fisico'] = 0.0
    df_clean['origen'] = 'SISTEMA'
    df_clean['fecha_conteo'] = None
    
    # Crear un campo único para búsqueda rápida
    df_clean['busqueda'] = df_clean['descripcion'] + " | " + df_clean['unidad']
    
    return df_clean.reset_index(drop=True)

# ==========================================
# 3. INTERFAZ TIPO "CAJERO / SCANNER"
# ==========================================
st.title("🏢 Inventario Rápido (Modo Scanner)")

if 'df_master' not in st.session_state: 
    st.session_state.df_master = None
    st.session_state.ultimo_editado = ""

# --- PASO 1: CARGA ---
if st.session_state.df_master is None:
    with st.expander("📂 INICIAR SESIÓN (Cargar Archivo)", expanded=True):
        archivo = st.file_uploader("Sube Excel", type=['xlsx','xls','csv'])
        if archivo:
            df_new = cargar_datos(archivo)
            if df_new is not None:
                st.session_state.df_master = df_new
                st.rerun()

# --- PASO 2: OPERACIÓN ---
else:
    df = st.session_state.df_master
    
    # Calcular métricas globales
    total_items = len(df)
    items_contados = len(df[df['fisico'] > 0])
    avance = int((items_contados / total_items) * 100) if total_items > 0 else 0
    
    # --- BARRA SUPERIOR DE ESTADO ---
    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
    c1.progress(avance, text=f"Progreso Global: {avance}%")
    c2.metric("Total Productos", total_items)
    c3.metric("Ya Contados", items_contados)
    
    pendientes = total_items - items_contados
    c4.metric("Pendientes", pendientes, delta_color="inverse")

    st.markdown("---")

    # ======================================================
    # ZONA DE TRABAJO (INPUTS ARRIBA - SIN LAG)
    # ======================================================
    col_input, col_info = st.columns([1, 1])

    with col_input:
        st.info("👇 **ZONA DE CONTEO**")
        
        # 1. Selector de Producto (Funciona como buscador)
        lista_productos = df['busqueda'].tolist()
        producto_selec = st.selectbox(
            "1. Busca el Producto:", 
            options=["Seleccionar..."] + lista_productos,
            index=0
        )
        
        # 2. Input de Cantidad
        # Usamos un formulario para agrupar la acción y evitar recargas parciales
        with st.form("panel_entrada", clear_on_submit=True):
            cantidad_input = st.number_input("2. Cantidad Física:", min_value=0.0, step=1.0)
            
            # Botones grandes
            cols_btn = st.columns(2)
            guardar = cols_btn[0].form_submit_button("✅ GUARDAR (Enter)", type="primary")
            nuevo_manual = cols_btn[1].form_submit_button("➕ CREAR NUEVO")
            
            if guardar and producto_selec != "Seleccionar...":
                # Lógica de actualización INSTANTÁNEA
                idx = df[df['busqueda'] == producto_selec].index[0]
                st.session_state.df_master.at[idx, 'fisico'] = cantidad_input
                st.session_state.df_master.at[idx, 'fecha_conteo'] = datetime.now().strftime("%H:%M:%S")
                st.session_state.ultimo_editado = f"{producto_selec.split('|')[0]} -> {cantidad_input}"
                st.rerun()
            
            if nuevo_manual:
                # Marcador para abrir modal o cambiar estado (simplificado aquí)
                st.session_state.modo_crear = True
                st.rerun()

    # ZONA DE INFORMACIÓN DEL PRODUCTO SELECCIONADO (VISTA PREVIA)
    with col_info:
        if producto_selec != "Seleccionar...":
            row = df[df['busqueda'] == producto_selec].iloc[0]
            
            st.success(f"**Producto:** {row['descripcion']}")
            
            col_i1, col_i2, col_i3 = st.columns(3)
            col_i1.metric("Sistema (Req)", row['requerido'])
            
            # Calculamos diferencia en tiempo real para visualización
            diff = cantidad_input - row['requerido'] # Pre-cálculo visual
            if diff == 0: color_diff = "normal"
            else: color_diff = "off"
            
            col_i2.metric("Unidad", row['unidad'])
            col_i3.metric("Estado Actual", 
                          "✅ OK" if row['fisico'] == row['requerido'] and row['fisico'] > 0 else 
                          ("🔻 FALTAN" if row['fisico'] < row['requerido'] else "🔺 SOBRAN"))

            if st.session_state.ultimo_editado:
                st.caption(f"Último guardado: {st.session_state.ultimo_editado}")
        else:
            st.warning("👈 Selecciona un producto para empezar a contar.")

    # --- AGREGAR PRODUCTO MANUAL (SI SE ACTIVÓ) ---
    if st.session_state.get('modo_crear', False):
        with st.expander("Crear Producto Nuevo", expanded=True):
            with st.form("crear_manual"):
                nm_nom = st.text_input("Nombre")
                nm_und = st.text_input("Unidad")
                nm_cant = st.number_input("Cantidad", 1.0)
                if st.form_submit_button("Guardar Nuevo"):
                    new_row = pd.DataFrame([{
                        'codigo':'MAN','descripcion':nm_nom,'requerido':0,'unidad':nm_und.upper(),
                        'fisico':nm_cant,'origen':'MANUAL', 'busqueda': f"{nm_nom} | {nm_und.upper()}",
                        'fecha_conteo': datetime.now().strftime("%H:%M:%S")
                    }])
                    st.session_state.df_master = pd.concat([st.session_state.df_master, new_row], ignore_index=True)
                    st.session_state.modo_crear = False
                    st.rerun()
            if st.button("Cancelar"):
                st.session_state.modo_crear = False
                st.rerun()

    st.write("---")

    # ======================================================
    # ZONA DE LISTADO (SOLO LECTURA - ACTUALIZACIÓN INSTANTÁNEA)
    # ======================================================
    # Filtros de visualización
    fc1, fc2, fc3 = st.columns([2, 1, 1])
    search_list = fc1.text_input("🔍 Filtrar lista abajo:", placeholder="Buscar...")
    ver_pendientes = fc2.checkbox("Ver solo pendientes", value=False)
    
    # Preparar vista
    df_view = st.session_state.df_master.copy()
    
    # Calcular Emojis para la tabla
    def get_status(r):
        if r['origen'] == 'MANUAL': return "🆕 NUEVO"
        if r['fisico'] == 0 and r['requerido'] > 0: return "⬜ PENDIENTE"
        if r['fisico'] == r['requerido']: return "✅ CORRECTO"
        if r['fisico'] < r['requerido']: return "🔻 FALTA"
        return "🔺 SOBRA"
    
    df_view['ESTADO'] = df_view.apply(get_status, axis=1)
    
    # Filtros
    if search_list:
        df_view = df_view[df_view['descripcion'].str.lower().str.contains(search_list.lower())]
    if ver_pendientes:
        df_view = df_view[df_view['fisico'] == 0]

    # Ordenar: Lo último contado arriba para ver confirmación visual
    df_view = df_view.sort_values(by=['fecha_conteo'], ascending=False)

    # Mostrar Tabla (SOLO LECTURA = CERO BUGS)
    st.dataframe(
        df_view[['descripcion', 'unidad', 'requerido', 'fisico', 'ESTADO', 'fecha_conteo']],
        use_container_width=True,
        hide_index=True,
        column_config={
            "descripcion": "Producto",
            "requerido": "Sistema",
            "fisico": "Físico",
            "ESTADO": "Estado",
            "fecha_conteo": "Hora"
        }
    )

    # --- DESCARGA ---
    st.write("---")
    if st.button("📥 Descargar Reporte Final"):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final = st.session_state.df_master.copy()
            df_final['Diferencia'] = df_final['fisico'] - df_final['requerido']
            df_final['Estado_Final'] = df_final.apply(get_status, axis=1)
            
            df_final.drop(columns=['busqueda']).to_excel(writer, sheet_name='INVENTARIO', index=False)
            
            # Formato condicional básico
            wb = writer.book; ws = writer.sheets['INVENTARIO']
            fmt_ok = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            ws.conditional_format(1, 6, len(df_final), 6, {'type':'text','criteria':'containing','value':'CORRECTO','format':fmt_ok})
            
        st.download_button("Guardar Excel", data=buffer.getvalue(), file_name="Inventario_Final.xlsx")

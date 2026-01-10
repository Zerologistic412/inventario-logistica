import streamlit as st
import pandas as pd
import io
import xlsxwriter
from datetime import datetime
import pdfplumber  # 🆕 NECESARIO PARA LEER PDFS

# ==========================================
# 1. CONFIGURACIÓN ROBUSTA & ESTÉTICA PRO
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
    
    /* Ajuste para inputs numéricos */
    input[type=number] { font-size: 20px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. LÓGICA DE NEGOCIO (MOTOR UNIVERSAL 5x5)
# ==========================================
def cargar_datos_universal(file):
    df = None
    filename = file.name.lower()
    
    # ---------------------------------------------------------
    # ESTRATEGIA A: ES UN PDF?
    # ---------------------------------------------------------
    if filename.endswith('.pdf'):
        try:
            with pdfplumber.open(file) as pdf:
                all_rows = []
                for page in pdf.pages:
                    # Extraer tablas de cada página
                    table = page.extract_table()
                    if table:
                        all_rows.extend(table)
                
                if all_rows:
                    df = pd.DataFrame(all_rows)
        except Exception as e:
            st.error(f"Error leyendo PDF: {e}")
            return None

    # ---------------------------------------------------------
    # ESTRATEGIA B: EXCEL / CSV / HTML / OTROS
    # ---------------------------------------------------------
    else:
        # 1. Intentar como Excel normal
        try:
            df = pd.read_excel(file, header=None)
        except:
            # 2. Intentar como HTML (Falsos Excel)
            try:
                file.seek(0)
                dfs = pd.read_html(file, header=None)
                if dfs: df = dfs[0]
            except:
                # 3. Intentar como CSV/Texto
                try:
                    file.seek(0)
                    df = pd.read_csv(file, header=None, sep=None, engine='python', encoding='latin-1')
                except: return None

    if df is None: return None

    # ---------------------------------------------------------
    # EL SABUESO: BUSCADOR DE CABECERAS INTELIGENTE
    # (Funciona igual para PDF que para Excel)
    # ---------------------------------------------------------
    fila_titulos = -1
    # Buscamos en las primeras 50 filas (los PDFs a veces tienen mucha basura arriba)
    for i, row in df.head(50).iterrows():
        fila_str = " ".join([str(val).lower() for val in row.values])
        # Buscamos palabras clave
        if "producto" in fila_str and ("cantidad" in fila_str or "stock" in fila_str or "saldo" in fila_str):
            fila_titulos = i
            break
    
    if fila_titulos == -1: 
        st.warning("⚠️ Leí el archivo, pero no encontré las columnas 'Producto' y 'Cantidad'.")
        return None

    # ---------------------------------------------------------
    # LIMPIEZA Y ESTRUCTURACIÓN
    # ---------------------------------------------------------
    try:
        df_data = df.iloc[fila_titulos+1:].copy()
        df_data.columns = df.iloc[fila_titulos]

        # Normalizar nombres de columnas
        cols = [str(c).lower().strip() for c in df_data.columns]
        df_data.columns = cols

        # Detectar columnas clave con lógica difusa
        col_desc = next((c for c in cols if "producto" in c or "descrip" in c or "item" in c), None)
        col_req = next((c for c in cols if "cantidad" in c or "stock" in c or "saldo" in c or "inv" in c), None)
        col_und = next((c for c in cols if "unidad" in c or "medida" in c or "u.m" in c), None)
        col_cod = next((c for c in cols if "codigo" in c or "sku" in c or "id" in c), None)

        if not col_desc or not col_req: return None

        df_clean = pd.DataFrame()
        df_clean['codigo'] = df_data[col_cod].astype(str).str.strip() if col_cod else "-"
        df_clean['descripcion'] = df_data[col_desc].astype(str).str.strip()
        
        # Limpieza nuclear de números (quita letras, símbolos de moneda, espacios)
        df_clean['requerido'] = pd.to_numeric(
            df_data[col_req].astype(str).str.replace(',', '.').str.replace(r'[^\d.]', '', regex=True), 
            errors='coerce'
        ).fillna(0)
        
        df_clean['unidad'] = df_data[col_und].astype(str).str.strip().fillna('').str.upper() if col_und else "UND"
        
        # Filtros: Eliminar filas vacías o con stock 0
        df_clean = df_clean[df_clean['descripcion'] != 'nan']
        df_clean = df_clean[df_clean['descripcion'] != '']
        df_clean = df_clean[df_clean['requerido'] > 0]
        
        df_clean['fisico'] = 0.0
        df_clean['origen'] = 'SISTEMA'
        df_clean['fecha_conteo'] = None
        
        # Campo de búsqueda
        df_clean['busqueda'] = df_clean['descripcion'] + " | " + df_clean['unidad']
        
        return df_clean.reset_index(drop=True)
    except Exception as e:
        st.error(f"Error procesando datos: {e}")
        return None

# ==========================================
# 3. INTERFAZ GRÁFICA
# ==========================================
st.title("🏢 Inventario Universal (PDF/Excel)")

if 'df_master' not in st.session_state: 
    st.session_state.df_master = None
    st.session_state.ultimo_editado = ""

# --- PASO 1: CARGA ---
if st.session_state.df_master is None:
    with st.expander("📂 CARGAR ARCHIVO", expanded=True):
        st.write("Sube tu inventario en PDF, Excel (nuevo o viejo) o CSV.")
        # Aceptamos casi cualquier cosa
        archivo = st.file_uploader("Arrastra aquí", type=['xlsx','xls','csv','pdf','txt'])
        
        if archivo:
            with st.spinner('Analizando documento...'):
                df_new = cargar_datos_universal(archivo)
                
                if df_new is not None:
                    if len(df_new) > 0:
                        st.session_state.df_master = df_new
                        st.success(f"✅ ¡Éxito! Se detectaron {len(df_new)} productos.")
                        st.rerun()
                    else:
                        st.error("El archivo se leyó, pero no quedaron productos (¿Quizás todos tienen Stock 0?).")
                else:
                    st.error("❌ No pude entender la tabla de este archivo. Asegúrate que tenga encabezados claros.")

# --- PASO 2: OPERACIÓN (Modo Scanner) ---
else:
    df = st.session_state.df_master
    
    # Calcular métricas
    total_items = len(df)
    items_contados = len(df[df['fisico'] > 0])
    avance = int((items_contados / total_items) * 100) if total_items > 0 else 0
    
    # Barra Superior
    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
    c1.progress(avance/100, text=f"Progreso Global: {avance}%")
    c2.metric("Total Productos", total_items)
    c3.metric("Ya Contados", items_contados)
    
    pendientes = total_items - items_contados
    c4.metric("Pendientes", pendientes, delta_color="inverse")

    st.markdown("---")

    # ZONA DE TRABAJO
    col_input, col_info = st.columns([1, 1])

    with col_input:
        st.info("👇 **ZONA DE ESCANEO / BÚSQUEDA**")
        
        lista_productos = df['busqueda'].tolist()
        producto_selec = st.selectbox(
            "1. Busca el Producto:", 
            options=["Seleccionar..."] + lista_productos,
            index=0
        )
        
        with st.form("panel_entrada", clear_on_submit=True):
            cantidad_input = st.number_input("2. Cantidad Física:", min_value=0.0, step=1.0)
            
            cols_btn = st.columns(2)
            guardar = cols_btn[0].form_submit_button("✅ GUARDAR", type="primary")
            nuevo_manual = cols_btn[1].form_submit_button("➕ CREAR NUEVO")
            
            if guardar and producto_selec != "Seleccionar...":
                idx = df[df['busqueda'] == producto_selec].index[0]
                st.session_state.df_master.at[idx, 'fisico'] = cantidad_input
                st.session_state.df_master.at[idx, 'fecha_conteo'] = datetime.now().strftime("%H:%M:%S")
                st.session_state.ultimo_editado = f"{producto_selec.split('|')[0]} -> {cantidad_input}"
                st.rerun()
            
            if nuevo_manual:
                st.session_state.modo_crear = True
                st.rerun()

    # INFORMACIÓN
    with col_info:
        if producto_selec != "Seleccionar...":
            row = df[df['busqueda'] == producto_selec].iloc[0]
            st.success(f"**{row['descripcion']}**")
            
            col_i1, col_i2, col_i3 = st.columns(3)
            col_i1.metric("Sistema", row['requerido'])
            col_i2.metric("Unidad", row['unidad'])
            col_i3.metric("Diferencia", row['fisico'] - row['requerido'] if row['fisico'] > 0 else 0)

            if st.session_state.ultimo_editado:
                st.caption(f"Último guardado: {st.session_state.ultimo_editado}")
        else:
            st.warning("👈 Selecciona un producto.")

    # MODAL CREAR NUEVO
    if st.session_state.get('modo_crear', False):
        with st.expander("Crear Producto Nuevo", expanded=True):
            with st.form("crear_manual"):
                nm_nom = st.text_input("Nombre")
                nm_und = st.text_input("Unidad")
                nm_cant = st.number_input("Cantidad", 1.0)
                if st.form_submit_button("Guardar"):
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

    # TABLA RESUMEN
    fc1, fc2 = st.columns([2, 1])
    search_list = fc1.text_input("🔎 Buscar en lista:", placeholder="Filtro...")
    ver_pendientes = fc2.checkbox("Ocultar completados")
    
    df_view = st.session_state.df_master.copy()
    
    def get_status(r):
        if r['origen'] == 'MANUAL': return "🆕 NUEVO"
        if r['fisico'] == 0: return "⬜ PENDIENTE"
        if r['fisico'] == r['requerido']: return "✅ EXACTO"
        return "⚠️ DIFERENCIA"
    
    df_view['ESTADO'] = df_view.apply(get_status, axis=1)
    
    if search_list:
        df_view = df_view[df_view['descripcion'].str.lower().str.contains(search_list.lower())]
    if ver_pendientes:
        df_view = df_view[df_view['fisico'] == 0]

    df_view = df_view.sort_values(by=['fecha_conteo'], ascending=False)

    st.dataframe(
        df_view[['descripcion', 'unidad', 'requerido', 'fisico', 'ESTADO', 'fecha_conteo']],
        use_container_width=True,
        hide_index=True
    )

    if st.button("📥 Descargar Excel"):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            st.session_state.df_master.to_excel(writer, index=False)
        st.download_button("Guardar", data=buffer.getvalue(), file_name="Inventario_Final.xlsx")

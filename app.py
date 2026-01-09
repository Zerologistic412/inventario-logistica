import streamlit as st
import pandas as pd
import io
import xlsxwriter
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN MOBILE-FIRST
# ==========================================
st.set_page_config(page_title="Inv. Móvil", page_icon="📱", layout="centered", initial_sidebar_state="collapsed")

# CSS TRUCADO PARA DEDOS (TOUCH FRIENDLY)
st.markdown("""
    <style>
    /* Botones más grandes y gordos para dedos */
    div.stButton > button {
        width: 100%;
        height: 60px;
        font-size: 20px;
        font-weight: bold;
        border-radius: 12px;
        margin-bottom: 10px;
    }
    
    /* Métricas grandes para ver de lejos */
    div[data-testid="stMetricValue"] { font-size: 24px; }
    
    /* Eliminar padding innecesario en móviles */
    .block-container { padding-top: 2rem; padding-bottom: 5rem; }
    
    /* Tarjeta de Producto Destacado */
    .product-card {
        background-color: #e8f4f8;
        padding: 15px;
        border-radius: 15px;
        border: 2px solid #0078D7;
        margin-bottom: 15px;
        text-align: center;
    }
    .product-title { font-size: 22px; font-weight: bold; color: #1f2937; }
    .product-code { font-size: 14px; color: #6b7280; font-family: monospace; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. LÓGICA (IGUAL PERO ROBUSTA)
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
    
    # Campo de búsqueda optimizado
    df_clean['busqueda'] = df_clean['descripcion'] 
    
    return df_clean.reset_index(drop=True)

# ==========================================
# 3. INTERFAZ TÁCTIL (TOUCH)
# ==========================================
if 'df_master' not in st.session_state: 
    st.session_state.df_master = None
    st.session_state.prod_index_actual = None # Para saber qué estamos editando

# --- PANTALLA 1: LOGIN / CARGA ---
if st.session_state.df_master is None:
    st.title("📱 Inventario Móvil")
    st.info("Sube el archivo desde tu celular o tablet.")
    archivo = st.file_uploader("Toca para subir Excel", type=['xlsx','csv'])
    if archivo:
        df_new = cargar_datos(archivo)
        if df_new is not None:
            st.session_state.df_master = df_new
            st.rerun()

# --- PANTALLA 2: OPERACIÓN ---
else:
    df = st.session_state.df_master
    
    # --- BARRA SUPERIOR FIJA ---
    col_a, col_b = st.columns([3, 1])
    progreso = len(df[df['fisico']>0])
    total = len(df)
    col_a.progress(progreso/total, text=f"Avance: {progreso}/{total}")
    if col_b.button("💾", help="Descargar"):
        st.session_state.descargar = True

    st.markdown("---")

    # --- SELECTOR DE PRODUCTO (Buscador) ---
    # Usamos selectbox porque en móviles el teclado nativo ayuda a filtrar
    lista_nombres = ["🔍 Toca para buscar..."] + df['busqueda'].tolist()
    
    # Mantenemos la selección en session state para que no se resetee
    sel_box = st.selectbox("", options=lista_nombres, label_visibility="collapsed")

    # --- TARJETA DE PRODUCTO ACTIVO ---
    if sel_box != "🔍 Toca para buscar...":
        # Encontrar índice
        idx = df[df['busqueda'] == sel_box].index[0]
        row = df.iloc[idx]
        val_actual = row['fisico']
        
        # Tarjeta Visual (HTML)
        st.markdown(f"""
        <div class="product-card">
            <div class="product-title">{row['descripcion']}</div>
            <div class="product-code">COD: {row['codigo']} | UNIDAD: {row['unidad']}</div>
            <div style="margin-top:10px; font-size:18px; color:#555;">
                Esperado: <b>{row['requerido']}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # --- CONTROLES TÁCTILES GIGANTES ---
        c_disp, c_val = st.columns([1, 2])
        
        # Display Gigante del valor actual
        with c_val:
            st.markdown(f"""
            <div style="text-align:center; font-size: 50px; font-weight:bold; color:#0078D7; border:1px solid #ddd; border-radius:10px;">
                {int(val_actual)}
            </div>
            """, unsafe_allow_html=True)

        st.write("###### 👇 Sumar/Restar (Toque rápido)")
        
        # Botonera Matemática (Evita usar teclado)
        b1, b2, b3, b4 = st.columns(4)
        
        # Definir acciones de botones
        def actualizar(cantidad):
            nueva_cantidad = max(0, val_actual + cantidad)
            st.session_state.df_master.at[idx, 'fisico'] = nueva_cantidad
            st.session_state.df_master.at[idx, 'fecha_conteo'] = datetime.now().strftime("%H:%M:%S")
            # No usamos rerun aquí para que sea fluido, Streamlit detecta el cambio de estado
        
        if b1.button("➖ 1"): actualizar(-1); st.rerun()
        if b2.button("➕ 1"): actualizar(1); st.rerun()
        if b3.button("➕ 5"): actualizar(5); st.rerun()
        if b4.button("➕ 10"): actualizar(10); st.rerun()

        # Entrada Manual (Solo si es necesario corregir un número grande)
        with st.expander("⌨️ Teclado Manual (Corrección)"):
            nuevo_manual = st.number_input("Escribir cantidad exacta:", value=float(val_actual), step=1.0)
            if st.button("Actualizar Manual"):
                st.session_state.df_master.at[idx, 'fisico'] = nuevo_manual
                st.session_state.df_master.at[idx, 'fecha_conteo'] = datetime.now().strftime("%H:%M:%S")
                st.rerun()
                
        # Feedback de estado
        if val_actual == row['requerido'] and val_actual > 0:
            st.success("✅ ¡CUADRADO PERFECTO!")
        elif val_actual > row['requerido']:
            st.warning(f"⚠️ Sobran {val_actual - row['requerido']}")
        elif val_actual > 0:
            st.error(f"🔻 Faltan {row['requerido'] - val_actual}")

    else:
        st.info("👆 Selecciona un producto arriba para empezar.")
        
        # Mostrar últimos contados (Historial para el Supervisor)
        st.write("---")
        st.caption("🕒 Últimos movimientos:")
        ultimos = df[df['fisico']>0].sort_values('fecha_conteo', ascending=False).head(3)
        if not ultimos.empty:
            for i, r in ultimos.iterrows():
                st.text(f"✅ {r['fisico']} - {r['descripcion'][:20]}...")

    # --- SECCIÓN DE DESCARGA (Solo aparece si se pidió) ---
    if st.session_state.get('descargar', False):
        st.write("---")
        st.write("### 📤 Exportar Datos")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final = st.session_state.df_master.copy()
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 TOCA AQUÍ PARA BAJAR EXCEL",
            data=buffer.getvalue(),
            file_name="Inventario_Movil.xlsx",
            mime="application/vnd.ms-excel",
            type="primary"
        )
        if st.button("Cerrar panel descarga"):
            st.session_state.descargar = False
            st.rerun()

import streamlit as st
import pandas as pd
import io
import xlsxwriter
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN
# ==========================================
st.set_page_config(page_title="Inv. Móvil", page_icon="📱", layout="centered", initial_sidebar_state="collapsed")

# CSS para móviles
st.markdown("""
    <style>
    div.stButton > button { width: 100%; height: 60px; font-size: 20px; font-weight: bold; border-radius: 12px; margin-bottom: 10px; }
    div[data-testid="stMetricValue"] { font-size: 24px; }
    .product-card { background-color: #e8f4f8; padding: 15px; border-radius: 15px; border: 2px solid #0078D7; margin-bottom: 15px; text-align: center; }
    .product-title { font-size: 22px; font-weight: bold; color: #1f2937; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. LÓGICA DE CARGA (MEJORADA)
# ==========================================
def cargar_datos(file):
    # Intentar leer Excel o CSV
    try:
        df = pd.read_excel(file, header=None)
    except:
        try:
            file.seek(0)
            df = pd.read_csv(file, header=None, sep=None, engine='python', encoding='latin-1')
        except Exception as e:
            st.error(f"❌ Error leyendo el archivo: {e}")
            return None

    # Buscar la fila de encabezados
    start_row = -1
    for index, row in df.iterrows():
        fila_texto = str(row.values).lower()
        # Buscamos palabras clave flexibles
        tiene_nombre = "producto" in fila_texto or "descripcion" in fila_texto or "descripción" in fila_texto
        tiene_cant = "cantidad" in fila_texto or "stock" in fila_texto or "saldo" in fila_texto or "requerido" in fila_texto
        
        if tiene_nombre and tiene_cant:
            start_row = index
            break
    
    if start_row == -1:
        st.error("⚠️ NO ENCUENTRO LOS DATOS. Tu Excel debe tener una fila con títulos que diga 'Producto' (o Descripción) y 'Cantidad' (o Stock).")
        st.write("Así veo tu archivo (primeras 5 filas):", df.head())
        return None
    
    # Procesar columnas
    try:
        df_data = df.iloc[start_row + 1:].copy()
        df_data.columns = df.iloc[start_row] # Asignar títulos encontrados
        
        # Limpiar nombres de columnas
        df_data.columns = [str(c).lower().strip() for c in df_data.columns]
        columnas = df_data.columns

        # Identificar columnas automáticamente
        col_desc = next((c for c in columnas if "producto" in c or "descrip" in c), None)
        col_req = next((c for c in columnas if "cantidad" in c or "stock" in c or "saldo" in c), None)
        col_cod = next((c for c in columnas if "codigo" in c or "sku" in c), None)
        col_und = next((c for c in columnas if "unidad" in c or "medida" in c), None)

        if not col_desc or not col_req:
            st.error(f"❌ Faltan columnas. Encontré: {list(columnas)}, pero necesito 'Producto' y 'Cantidad'.")
            return None

        # Crear DataFrame limpio
        df_clean = pd.DataFrame()
        df_clean['descripcion'] = df_data[col_desc].astype(str).str.strip()
        df_clean['requerido'] = pd.to_numeric(df_data[col_req], errors='coerce').fillna(0)
        
        if col_cod: df_clean['codigo'] = df_data[col_cod].astype(str).str.strip()
        else: df_clean['codigo'] = "-"
            
        if col_und: df_clean['unidad'] = df_data[col_und].astype(str).str.strip().upper()
        else: df_clean['unidad'] = "UND"

        # Filtrar vacíos
        df_clean = df_clean[df_clean['descripcion'] != 'nan']
        df_clean = df_clean[df_clean['requerido'] > 0]
        
        # Columnas de trabajo
        df_clean['fisico'] = 0.0
        df_clean['fecha_conteo'] = None
        df_clean['busqueda'] = df_clean['descripcion'] 
        
        return df_clean.reset_index(drop=True)

    except Exception as e:
        st.error(f"❌ Error procesando columnas: {e}")
        return None

# ==========================================
# 3. INTERFAZ
# ==========================================
if 'df_master' not in st.session_state: st.session_state.df_master = None

if st.session_state.df_master is None:
    st.title("📱 Inventario Móvil")
    st.info("Sube el archivo para empezar.")
    archivo = st.file_uploader("Toca para subir Excel", type=['xlsx','csv'])
    
    if archivo:
        with st.spinner("Leyendo archivo..."):
            df_new = cargar_datos(archivo)
            if df_new is not None:
                st.session_state.df_master = df_new
                st.rerun()

else:
    df = st.session_state.df_master
    
    # Barra Superior
    c1, c2 = st.columns([3,1])
    progreso = len(df[df['fisico']>0])
    c1.progress(progreso/len(df), text=f"Avance: {progreso}/{len(df)}")
    if c2.button("💾"): st.session_state.descargar = True

    # Buscador
    lista = ["🔍 Buscar..."] + df['busqueda'].tolist()
    sel = st.selectbox("", lista, label_visibility="collapsed")

    if sel != "🔍 Buscar...":
        idx = df[df['busqueda'] == sel].index[0]
        row = df.iloc[idx]
        val = row['fisico']
        
        st.markdown(f"""
        <div class="product-card">
            <div class="product-title">{row['descripcion']}</div>
            <div>COD: {row['codigo']} | {row['unidad']}</div>
            <h3>Esperado: {row['requerido']}</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Panel Conteo
        c_num, c_bot = st.columns([1,2])
        with c_num:
             st.markdown(f"<h1 style='text-align:center; color:#0078D7; font-size:50px;'>{int(val)}</h1>", unsafe_allow_html=True)
        
        # Botonera
        b1, b2, b3, b4 = st.columns(4)
        def act(n):
            st.session_state.df_master.at[idx, 'fisico'] = max(0, val + n)
            st.session_state.df_master.at[idx, 'fecha_conteo'] = datetime.now().strftime("%H:%M")
            st.rerun()

        if b1.button("➖1"): act(-1)
        if b2.button("➕1"): act(1)
        if b3.button("➕5"): act(5)
        if b4.button("➕10"): act(10)
        
        if val == row['requerido']: st.success("✅ ¡Completo!")
        
    else:
        st.info("👆 Selecciona un producto arriba.")

    # Descarga
    if st.session_state.get('descargar', False):
        st.write("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            st.session_state.df_master.to_excel(writer, index=False)
        st.download_button("📥 BAJAR EXCEL", buffer.getvalue(), "Conteo.xlsx", "application/vnd.ms-excel", type="primary")
        if st.button("Cerrar"): 
            st.session_state.descargar = False
            st.rerun()

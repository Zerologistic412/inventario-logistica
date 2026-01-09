import streamlit as st
import pandas as pd
import io
import xlsxwriter
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN E INTERFAZ MÓVIL (El "Look")
# ==========================================
st.set_page_config(page_title="Inv. Móvil", page_icon="📱", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    /* ESTILOS PARA DEDOS GORDOS (Móvil) */
    div.stButton > button {
        width: 100%;
        height: 60px;
        font-size: 20px;
        font-weight: bold;
        border-radius: 12px;
        margin-bottom: 10px;
        box-shadow: 0px 4px 6px rgba(0,0,0,0.1);
    }
    .product-card {
        background-color: #f0f9ff;
        padding: 20px;
        border-radius: 15px;
        border-left: 10px solid #0078D7;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .big-number {
        font-size: 50px;
        font-weight: 800;
        color: #0078D7;
        text-align: center;
        background: white;
        border-radius: 10px;
        border: 1px solid #ddd;
        padding: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. EL MOTOR "4x4" (Recuperado de la versión robusta)
# ==========================================
def normalizar_texto(texto):
    """Limpia textos para comparar (quita mayúsculas y espacios)"""
    if pd.isna(texto): return ""
    return str(texto).lower().strip().replace('.', '').replace(':', '')

def cargar_datos_robusto(file):
    """
    Esta es la función potente que usábamos antes.
    Lee Excels viejos (.xls), nuevos (.xlsx) y CSVs.
    Busca los títulos donde sea que estén.
    """
    df = None
    
    # 1. INTENTO DE LECTURA (Soporta .xls antiguo y .xlsx nuevo)
    try:
        # Intentamos leer con el motor por defecto
        df = pd.read_excel(file, header=None)
    except:
        try:
            # Si falla, intentamos como CSV
            file.seek(0)
            df = pd.read_csv(file, header=None, sep=None, engine='python', encoding='latin-1')
        except Exception as e:
            st.error(f"❌ No se pudo abrir el archivo. Asegúrate que no esté corrupto. Error: {e}")
            return None

    # 2. BUSCADOR DE CABECERAS (El "Sabueso")
    # Rastrea las primeras 20 filas buscando palabras clave
    fila_titulos = -1
    
    for i, row in df.head(20).iterrows():
        fila_str = " ".join([str(val).lower() for val in row.values])
        
        # Palabras clave flexibles (acepta sinónimos)
        tiene_prod = any(x in fila_str for x in ['producto', 'descrip', 'material', 'articulo', 'item'])
        tiene_cant = any(x in fila_str for x in ['cantidad', 'cant', 'stock', 'saldo', 'inventario', 'requerido'])
        
        if tiene_prod and tiene_cant:
            fila_titulos = i
            break
    
    if fila_titulos == -1:
        st.error("⚠️ No encuentro la fila de títulos. Busca que diga 'Producto' y 'Cantidad' en alguna fila.")
        return None

    # 3. LIMPIEZA Y ESTRUCTURACIÓN
    try:
        # Cortamos el Excel desde donde encontramos los títulos
        df_data = df.iloc[fila_titulos+1:].copy()
        df_data.columns = df.iloc[fila_titulos] # Asignamos los nombres correctos
        
        # Limpiamos nombres de columnas
        df_data.columns = [normalizar_texto(c) for c in df_data.columns]
        cols = df_data.columns
        
        # Identificamos columnas automáticamente
        col_desc = next((c for c in cols if any(x in c for x in ['producto', 'descrip', 'articulo'])), None)
        col_req = next((c for c in cols if any(x in c for x in ['cantidad', 'stock', 'saldo', 'requerido'])), None)
        col_cod = next((c for c in cols if any(x in c for x in ['codigo', 'sku', 'id'])), None)
        col_und = next((c for c in cols if any(x in c for x in ['unidad', 'medida', 'um'])), None)

        if not col_desc or not col_req:
            st.error(f"❌ Encontré la fila de títulos, pero no distingo cuál es Producto y cuál Cantidad. (Columnas detectadas: {list(cols)})")
            return None

        # Construimos la tabla limpia final
        df_final = pd.DataFrame()
        df_final['codigo'] = df_data[col_cod].astype(str).str.strip() if col_cod else "-"
        df_final['descripcion'] = df_data[col_desc].astype(str).str.strip()
        
        # Limpieza de números (convierte texto a numero, reemplaza comas, etc)
        df_final['requerido'] = pd.to_numeric(
            df_data[col_req].astype(str).str.replace(',', '.'), 
            errors='coerce'
        ).fillna(0)
        
        df_final['unidad'] = df_data[col_und].astype(str).str.strip().str.upper() if col_und else "UND"

        # Eliminar filas vacías o basura
        df_final = df_final[df_final['descripcion'] != 'nan']
        df_final = df_final[df_final['descripcion'] != '']
        df_final = df_final[df_final['requerido'] > 0] # Solo mostramos lo que tiene saldo > 0

        # Agregamos columnas de trabajo
        df_final['fisico'] = 0.0
        df_final['fecha'] = None
        df_final['busqueda'] = df_final['descripcion'] + " (" + df_final['codigo'] + ")"
        
        return df_final.reset_index(drop=True)

    except Exception as e:
        st.error(f"❌ Error procesando los datos: {e}")
        return None

# ==========================================
# 3. PANTALLAS DE LA APP
# ==========================================

# Inicializar sesión
if 'df_master' not in st.session_state: st.session_state.df_master = None
if 'descargar' not in st.session_state: st.session_state.descargar = False

# --- PANTALLA 1: CARGA DE ARCHIVO ---
if st.session_state.df_master is None:
    st.title("📱 Inventario 4x4")
    st.write("Versión robusta: Acepta Excels viejos y nuevos.")
    
    archivo = st.file_uploader("📂 Sube tu Excel aquí", type=['xlsx', 'xls', 'csv'])
    
    if archivo:
        with st.spinner("Analizando archivo a fondo..."):
            df_procesado = cargar_datos_robusto(archivo)
            if df_procesado is not None:
                st.session_state.df_master = df_procesado
                st.rerun()

# --- PANTALLA 2: OPERACIÓN MÓVIL ---
else:
    df = st.session_state.df_master
    
    # Barra de progreso superior
    c1, c2 = st.columns([3,1])
    conteo = len(df[df['fisico']>0])
    total = len(df)
    c1.progress(conteo/total, text=f"Progreso: {conteo} de {total}")
    
    if c2.button("💾"): 
        st.session_state.descargar = True
        st.rerun()

    st.write("---")

    # BUSCADOR
    opciones = ["🔍 Buscar producto..."] + df['busqueda'].tolist()
    seleccion = st.selectbox("", opciones, label_visibility="collapsed")

    # SI HAY SELECCIÓN, MOSTRAR TARJETA
    if seleccion != "🔍 Buscar producto...":
        idx = df[df['busqueda'] == seleccion].index[0]
        row = df.iloc[idx]
        val_actual = row['fisico']
        
        # Tarjeta de información
        st.markdown(f"""
        <div class="product-card">
            <h3 style="margin:0; color:#111;">{row['descripcion']}</h3>
            <p style="color:#666; margin:5px 0;">COD: {row['codigo']} | UNIDAD: {row['unidad']}</p>
            <div style="background:#ddd; height:2px; margin:10px 0;"></div>
            <p style="font-size:18px;">Sistema dice: <b>{row['requerido']}</b></p>
        </div>
        """, unsafe_allow_html=True)
        
        # Área de conteo
        c_num, c_btns = st.columns([1,2])
        
        with c_num:
            st.markdown(f'<div class="big-number">{int(val_actual)}</div>', unsafe_allow_html=True)
            if val_actual == row['requerido']:
                st.success("✅ Cuadrado")
            elif val_actual > row['requerido']:
                st.warning(f"⚠️ Sobran {val_actual - row['requerido']}")
        
        # Botones lógicos
        with c_btns:
            col_b1, col_b2 = st.columns(2)
            
            def sumar(n):
                st.session_state.df_master.at[idx, 'fisico'] = max(0, val_actual + n)
                st.session_state.df_master.at[idx, 'fecha'] = datetime.now().strftime("%H:%M")
                st.rerun()

            if col_b1.button("➖ 1"): sumar(-1)
            if col_b2.button("➕ 1"): sumar(1)
            
            col_b3, col_b4 = st.columns(2)
            if col_b3.button("➕ 5"): sumar(5)
            if col_b4.button("➕ 10"): sumar(10)
        
        # Corrección manual
        with st.expander("📝 Escribir número manualmente"):
            val_manual = st.number_input("Cantidad exacta", value=float(val_actual))
            if st.button("Guardar Manual"):
                st.session_state.df_master.at[idx, 'fisico'] = val_manual
                st.session_state.df_master.at[idx, 'fecha'] = datetime.now().strftime("%H:%M")
                st.rerun()

    else:
        st.info("👆 Usa el buscador de arriba para empezar a contar.")
        
        # Historial pequeño
        st.write("---")
        st.caption("Últimos contados:")
        st.dataframe(df[df['fisico']>0][['descripcion', 'fisico', 'fecha']].tail(3), hide_index=True)

    # --- ZONA DE DESCARGA ---
    if st.session_state.descargar:
        st.write("---")
        st.warning("⚠️ ¿Terminaste?")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            st.session_state.df_master.to_excel(writer, index=False)
        
        st.download_button("📥 DESCARGAR INVENTARIO FINAL", buffer.getvalue(), "Inventario_Final.xlsx", "application/vnd.ms-excel", type="primary")
        if st.button("Seguir contando"):
            st.session_state.descargar = False
            st.rerun()

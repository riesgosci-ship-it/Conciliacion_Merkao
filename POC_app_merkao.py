import streamlit as st
import smtplib
import base64
import os
import sys
import shutil
import re
import io
import zipfile
import pandas as pd
import difflib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl.styles import PatternFill, Font

# --- 1. CONFIGURACIÓN BÁSICA ---
st.set_page_config(page_title="Tesoriapp", page_icon="🏦", layout="wide", initial_sidebar_state="expanded")

# --- FUNCION PARA LEER IMAGEN LOCAL ---
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return ""

# =================================================================================
# ======================== MOTOR CONTABLE (PANDAS & LÓGICA) =======================
# =================================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIR_MAESTROS = os.path.join(BASE_DIR, "CONSOLIDADO MAESTROS")

# CORRECCIÓN DE RUTAS PARA LINUX (STREAMLIT CLOUD)
LOCALES_CONFIG = {
    957: {
        "codigos": [1051565, 1046395, 1051566, 1049627, 1052291, 1047968, 1052293],
        "plantilla_p1": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Ecommerce - 957.xlsx"),
        "plantilla_fin": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Final - 957.xlsx")
    },
    967: {
        "codigos": [1052311, 1052448], 
        "plantilla_p1": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Ecommerce - 967.xlsx"),
        "plantilla_fin": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Final - 967.xlsx")
    },
    972: {
        "codigos": [1052293, 1050522], 
        "plantilla_p1": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Ecommerce - 972.xlsx"),
        "plantilla_fin": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Final - 972.xlsx")
    },
    968: {
        "codigos": [1052310, 1052440], 
        "plantilla_p1": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Ecommerce - 968.xlsx"),
        "plantilla_fin": os.path.join(BASE_DIR, "PLANTILLA", "Plantilla Final - 968.xlsx")
    }
}

MESES_NOMBRES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

HEADERS_MAESTROS = {
    "JSATELITE": ["ORDEN_MERKAO", "COD_ENVIADO_CT2", "BOLETA", "TOTAL", "ID_TRANSACCION", "CAJA", "LOCAL", "FECHA", "Fecha_carga"],
    "LIQUIDADOR": ["FECHA", "ZONA", "RUTA", "CODIGO DE DESPACHO", "ESTADO BEETRACK", "DNI - CLIENTE", "TIPO DE PAGO BEETRACK", "MONTO DE PEDIDO", "MONTO A COBRAR", "MONTO COBRADO", "DIF", "MONTO PARCIAL", "OBSERVACIONES", "CODIGO DE COMERCIO", "N. COMERCIAL", "ESTADO FINAL", "TIPO DE POS", "TIPO DE PAGO", "COD AUTORIZACION", "Fecha_carga","FECHA DE PAGO"], 
    "MONITOR": ["N/V o Porforma", "Numero de viaje", "Fecha", "Metodo de pago", "Estado", "Tipo Doc. Cliente", "Nro. Doc. Cliente", "Centro de Distrib.", "Total", "Cantidad", "Monto Global", "Vtex ID", "Dispatch Number", "Tipo de facturacion", "Motivo", "Fecha_carga"],
    "IZIPAY": ["Codigo", "Producto", "Tipo_Mov", "Fecha_Proceso", "Fecha_Lote", "Lote_Manual", "Lote_Pos", "Terminal", "Voucher", "Autorizacion", "Cuotas", "Tarjeta", "Origen", "Transaccion", "Fecha_Consumo", "Importe", "Status", "Comision", "Comision_Afecta", "IGV", "Neto_Parcial", "Neto_Total", "Fecha_Abono", "Fecha_Abono_8Dig", "Observaciones", "ExtraComision", "Comis_Standar", "Comis_%", "Nro_ID", "Tpo_Comprob", "Nro_Comprob", "Fecha_carga"],
    "BBR": ["Fecha Trx", "N° Trx", "Grupo Medio Pago", "Medio Pago", "SubTipo", "Cód Cajero", "Cajero", "Caja", "N Cuenta", "Monto", "Fecha_carga"]
}

# --- FUNCIONES DE APOYO CONTABLE ---
def limpiar_id_texto(serie):
    return serie.astype(str).str.replace(r'\.0$', '', regex=True).str.strip().replace('nan', '')

def formatear_fecha_dt(serie, mes_objetivo=None):
    s_limpia = serie.astype(str).str.split('T').str[0].str.split(' ').str[0]
    def parse_smart_date(val):
        val = str(val).strip().replace('/', '-')
        if pd.isna(val) or val in ['nan', '', 'NaT', 'None']: return pd.NaT
        dt = pd.NaT
        if len(val) == 8 and val.isdigit(): dt = pd.to_datetime(val, format='%Y%m%d', errors='coerce')
        elif re.match(r'^\d{2}-\d{2}-\d{4}$', val): dt = pd.to_datetime(val, format='%d-%m-%Y', errors='coerce')
        elif re.match(r'^\d{4}-\d{2}-\d{2}$', val): dt = pd.to_datetime(val, format='%Y-%m-%d', errors='coerce')
        else: dt = pd.to_datetime(val, errors='coerce', dayfirst=True)
        if pd.notnull(dt) and mes_objetivo is not None:
            if dt.month != mes_objetivo and dt.day == mes_objetivo:
                try: dt = dt.replace(month=dt.day, day=dt.month)
                except ValueError: pass
        return dt
    return s_limpia.apply(parse_smart_date).dt.strftime('%Y-%m-%d')

def extraer_montos_regex(texto):
    if pd.isna(texto) or str(texto).strip() == "": return []
    texto = str(texto).replace(',', '')
    matches = re.findall(r'S/\s*(\d+(?:\.\d+)?)', texto, re.IGNORECASE)
    if not matches: matches = re.findall(r'(?<!\d)(\d+\.\d+)(?!\d)', texto)
    return matches

def extraer_codigos_autorizacion(texto):
    if pd.isna(texto) or str(texto).strip() == "": return []
    return [p.strip() for p in re.split(r'\s*[Yy/,.-]\s*', str(texto)) if p.strip()]

def obtener_tipo_nombre(n_tarjeta):
    n = str(n_tarjeta).strip()
    if n == 'nan' or not n: return ""
    if n.startswith('45733'): return "AGORA OFFLINE"
    if n.startswith(('2','3', '5')): return "TARJETA MASTER OFFLINE"
    return "MULTIMARCA"

def clasificar_tipo_codigo(n_tarjeta):
    n = str(n_tarjeta).strip()
    if n == 'nan' or not n: return ""
    if n.startswith('45733'): return 13
    if n.startswith(('2','3', '5')): return 21
    return 88

def leer_archivos_web(uploaded_files, columnas_esperadas=None):
    """Adaptación para leer archivos directamente de Streamlit (BytesIO)"""
    if columnas_esperadas is None: columnas_esperadas = []
    dfs = []
    for file in uploaded_files:
        try:
            nombre_archivo = file.name
            def procesar_df(df_temp):
                if df_temp.empty: return df_temp
                if columnas_esperadas:
                    renames = {}
                    for col in df_temp.columns:
                        col_str = str(col).strip()
                        if col_str in columnas_esperadas: continue
                        matches = difflib.get_close_matches(col_str, columnas_esperadas, n=1, cutoff=0.70)
                        if matches: renames[col] = matches[0] 
                    if renames: df_temp = df_temp.rename(columns=renames)
                df_temp['ARCHIVO_ORIGEN'] = nombre_archivo
                return df_temp

            if nombre_archivo.lower().endswith('.csv'):
                try: df = pd.read_csv(file, dtype=str, sep=',')
                except UnicodeDecodeError: 
                    file.seek(0)
                    df = pd.read_csv(file, dtype=str, sep=',', encoding='latin1')
                if not df.empty: dfs.append(procesar_df(df))
            else:
                try:
                    hojas_dict = pd.read_excel(file, sheet_name=None, dtype=str)
                    for nombre_hoja, df in hojas_dict.items():
                        if not df.empty: dfs.append(procesar_df(df))
                except Exception as e:
                    # Alternativa HTML para BBR en caso de error
                    file.seek(0)
                    df_list = pd.read_html(file, decimal='.', thousands=None)
                    if df_list: dfs.append(procesar_df(df_list[0].astype(str)))
        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")
            
    return pd.concat(dfs, ignore_index=True).drop_duplicates() if dfs else pd.DataFrame()


# =================================================================================
# ======================== ZONA A: PANTALLA DE LOGIN ==============================
# =================================================================================
if "usuario_autenticado" not in st.session_state:
    st.session_state["usuario_autenticado"] = False
if "vista_actual" not in st.session_state:
    st.session_state["vista_actual"] = "login"
if "usuario_nombre" not in st.session_state:
    st.session_state["usuario_nombre"] = ""

def enviar_correo_recuperacion(destinatario):
    try:
        remitente = st.secrets["email"]["usuario"]
        password = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario
        msg['Subject'] = "Recuperación de Acceso - Tesoriapp Merkao"
        cuerpo_html = f"""
        <html>
            <body style="font-family: Arial, sans-serif; color: #1E293B;">
                <h2 style="color: #1A237E;">Tesoriapp - Módulo de Conciliación</h2>
                <p>Tu contraseña corporativa provisional es: <strong>Spsa2026</strong></p>
            </body>
        </html>
        """
        msg.attach(MIMEText(cuerpo_html, 'html'))
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        return True, f"✅ Se ha enviado un correo a {destinatario}."
    except Exception as e:
        return False, f"❌ Error del sistema: Verifica credenciales en Streamlit Secrets."

if not st.session_state["usuario_autenticado"]:
    
    st.markdown("""
    <style>
        #MainMenu, header, footer, [data-testid="stToolbar"] {visibility: hidden;}
        .st-emotion-cache-1r6slb0 { padding: 0 !important; }
        [data-testid="stAppViewContainer"] { background: radial-gradient(circle, #283593 0%, #1A237E 70%, #121858 100%); }
        div[data-testid="stColumn"]:nth-of-type(2) > div {
            background-color: white; border-radius: 20px; padding: 50px 40px; box-shadow: 0 20px 40px rgba(0,0,0,0.5); max-width: 450px; margin: auto; 
        }
        div[data-testid="stColumn"]:nth-of-type(2) > div label, div[data-testid="stColumn"]:nth-of-type(2) > div p { color: #000000 !important; }
        .stTextInput label p { font-size: 13.5px !important; color: #000000 !important; font-weight: 600 !important; margin-bottom: 2px !important; }
        .stTextInput>div>div>input {
            background-color: #FFFFFF !important; border: 1.5px solid #CBD5E1 !important; color: #000000 !important; font-weight: 500 !important;
            border-radius: 8px !important; padding: 14px 16px !important; font-size: 15px;
        }
        .stTextInput>div>div>input::placeholder { color: #64748B !important; opacity: 1 !important; }
        .stButton { display: flex; justify-content: center; width: 100%; margin-top: 10px; }
        .stButton>button[kind="primary"] {
            background-color: #FFC20E !important; color: #1A237E !important; font-weight: 800 !important; font-size: 15px !important;
            border-radius: 25px !important; width: 100% !important; padding: 12px !important; white-space: nowrap !important; text-transform: uppercase;
        }
        .stButton>button[kind="primary"]:hover { background-color: #FFD54F !important; transform: translateY(-2px); box-shadow: 0 8px 15px rgba(255, 194, 14, 0.4) !important; }
        .boton-olvido { display: flex; justify-content: center !important; margin: 5px 0 10px 0; }
        .stButton>button[kind="secondary"] { color: #64748B !important; border: none !important; font-weight: 600 !important; }
        [data-testid="stCheckbox"] { display: flex; justify-content: center; margin-bottom: 15px; }
        .stCheckbox label span p { color: #000000 !important; }
    </style>
    """, unsafe_allow_html=True)

    c1, card_col, c3 = st.columns([1, 3, 1])
    with card_col:
        st.write("<br>", unsafe_allow_html=True) 
        if st.session_state["vista_actual"] == "login":
            img_b64 = get_base64_image("logo_merkao.png")
            if img_b64: st.markdown(f"<div style='display: flex; justify-content: center; align-items: center; gap: 20px; margin-bottom: 35px;'><img src='data:image/png;base64,{img_b64}' style='width: 100px; border-radius: 12px;'><h1 style='color: #000000 !important; margin: 0; padding: 0; font-size: 42px; font-weight: 900; letter-spacing: -1px;'>Tesoriapp</h1></div>", unsafe_allow_html=True)
            else: st.markdown("<h1 style='text-align: center; color: #000000 !important; font-weight: 900; font-size: 46px; margin-bottom: 30px;'>Tesoriapp</h1>", unsafe_allow_html=True)
            
            st.text_input("Usuario", placeholder="ej. MAYHELA.SIMON", key="usuario_input")
            st.text_input("Contraseña", type="password", placeholder="••••••••", key="password_input")
            st.markdown("<div class='boton-olvido'>", unsafe_allow_html=True)
            if st.button("¿Olvidaste tu contraseña?", type="secondary", key="btn_ir_recuperar"): st.session_state["vista_actual"] = "recuperar"; st.rerun() 
            st.markdown("</div>", unsafe_allow_html=True)
            st.checkbox("Recordar en este dispositivo", value=True)
            if "error_login" in st.session_state: st.error(st.session_state["error_login"])

            def verificar_login():
                usu = st.session_state.usuario_input.strip().upper() 
                pwd = st.session_state.password_input
                try:
                    if st.secrets["passwords"].get(usu) == pwd:
                        st.session_state["usuario_autenticado"] = True
                        st.session_state["usuario_nombre"] = usu
                        if "error_login" in st.session_state: del st.session_state["error_login"]
                    else: st.session_state["error_login"] = "❌ Usuario o contraseña incorrectos."
                except Exception: st.session_state["error_login"] = "❌ Error crítico: Bóveda secreta no configurada."

            st.button("INICIAR SESIÓN", on_click=verificar_login, type="primary")

        elif st.session_state["vista_actual"] == "recuperar":
            img_b64 = get_base64_image("logo_merkao.png")
            if img_b64: st.markdown(f"<div style='text-align: center; margin-bottom: 20px;'><img src='data:image/png;base64,{img_b64}' style='width: 100px; border-radius: 12px;'></div>", unsafe_allow_html=True)
            st.markdown("<h3 style='text-align: center; color: #000000 !important; font-weight: 800;'>Recuperar Acceso</h3>", unsafe_allow_html=True)
            st.text_input("Correo corporativo", placeholder="ejemplo@spsa.pe", key="correo_recuperar_input")
            st.markdown("<div class='boton-olvido'>", unsafe_allow_html=True)
            if st.button("Cancelar y Volver", type="secondary", key="btn_volver_login"):
                st.session_state["vista_actual"] = "login"; 
                if "pwd_recuperada_msg" in st.session_state: del st.session_state["pwd_recuperada_msg"]
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            if "pwd_recuperada_msg" in st.session_state:
                if "✅" in st.session_state["pwd_recuperada_msg"]: st.success(st.session_state["pwd_recuperada_msg"])
                else: st.error(st.session_state["pwd_recuperada_msg"])

            def intentar_recuperacion():
                correo = st.session_state.correo_recuperar_input.strip()
                if not correo or "@" not in correo: st.session_state["pwd_recuperada_msg"] = "⚠️ Ingresa un correo corporativo válido."
                else: _, st.session_state["pwd_recuperada_msg"] = enviar_correo_recuperacion(correo)
            
            st.button("ENVIAR INSTRUCCIONES", on_click=intentar_recuperacion, type="primary")


# =================================================================================
# ======================== ZONA B: APLICACIÓN PRINCIPAL (DASHBOARD) ===============
# =================================================================================
else:
    st.markdown("""
        <style>
            [data-testid="stAppViewContainer"] { background: #F4F7F9; }
            #MainMenu, footer {visibility: hidden;}
            header { background: transparent !important; }
            .stMainBlockContainer h1, .stMainBlockContainer h2, .stMainBlockContainer h3, .stMainBlockContainer p, .stMainBlockContainer label { color: #1E293B !important; }
            .top-bar { background-color: white; padding: 15px 30px 15px 60px; border-bottom: 2px solid #E2E8F0; display: flex; align-items: center; gap: 15px; margin-top: -80px; margin-bottom: 30px; margin-left: -50px; margin-right: -50px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
            [data-testid="stSidebar"] { background-color: #2C3E50 !important; border-right: none; }
            [data-testid="stSidebar"] * { color: #FFFFFF !important; }
            div.row-widget.stRadio > div { gap: 15px; padding: 10px 0; }
            div.row-widget.stRadio > div > label { background-color: transparent; padding: 10px 15px; border-radius: 8px; transition: 0.3s; cursor: pointer; }
            div.row-widget.stRadio > div > label:hover { background-color: rgba(255,255,255,0.1); }
            div.row-widget.stRadio > div > label > div:first-child { display: none; }
            div.row-widget.stRadio > div > label > div:last-child p { font-size: 16px; font-weight: 600; margin: 0; }
            [data-testid="stSidebar"] button[kind="secondary"] { background-color: rgba(220, 38, 38, 0.1) !important; border: 1px solid #DC2626 !important; color: #FCA5A5 !important; width: 100%; }
            [data-testid="stSidebar"] button[kind="secondary"]:hover { background-color: #DC2626 !important; color: white !important; }
        </style>
    """, unsafe_allow_html=True)

    img_b64 = get_base64_image("logo_merkao.png")
    logo_html = f"<img src='data:image/png;base64,{img_b64}' width='45' style='border-radius: 6px;'>" if img_b64 else "🏦"
    st.markdown(f"""
        <div class="top-bar">
            {logo_html}
            <h1 style="color: #1E293B !important; font-size: 28px !important; font-weight: 800 !important; margin: 0 !important;">Tesoriapp - Módulo de Conciliación</h1>
        </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("<h2 style='text-align: center; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 20px; margin-bottom: 20px; color: white !important;'>Menú Principal</h2>", unsafe_allow_html=True)
        opcion_menu = st.radio("Navegación", ["MAESTROS", "ARCHIVOS", "CONCILIACION"], label_visibility="collapsed")
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        st.markdown(f"**Bienvenid@**<br>{st.session_state['usuario_nombre']}", unsafe_allow_html=True)
        st.markdown("---")
        if st.button("Cerrar Sesión", use_container_width=True):
            st.session_state["usuario_autenticado"] = False
            st.rerun()

    if opcion_menu == "MAESTROS":
        st.markdown("<h2 style='color: #1E293B !important;'>📁 Base de Datos: Maestros SharePoint</h2>", unsafe_allow_html=True)
        st.write("Visualización de las listas maestras sincronizadas en la nube.")
        maestro_sel = st.selectbox("Seleccione el Maestro a visualizar:", ["JSATELITE", "IZIPAY", "LIQUIDADOR", "MONITOR", "BBR"])
        st.info(f"Mostrando estructura de la lista **{maestro_sel}** alojada en SharePoint:")
        
        # Esquemas conectados listos para SharePoint
        schema_maestros = {
            "JSATELITE": ["ORDEN_MERKAO", "COD_ENVIADO_CT2", "BOLETA", "TOTAL", "ID_TRANSACCION", "CAJA", "LOCAL", "FECHA", "Fecha_carga", "N_Trx"],
            "IZIPAY": ["Codigo", "Producto", "Tipo_Mov", "Fecha_Proceso", "Fecha_Lote", "Lote_Manual", "Lote_Pos", "Terminal", "Voucher", "Autorizacion", "Cuotas", "Tarjeta", "Origen", "Transaccion", "Fecha_Consumo", "Importe", "Status", "Comision", "Comision_Afecta", "IGV", "Neto_Parcial", "Neto_Total", "Fecha_Abono", "Fecha_Abono_8Dig", "Observaciones", "ExtraComision", "Comis_Standar", "Comis_%", "Nro_ID", "Tpo_Comprob", "Nro_Comprob", "Fecha_carga"],
            "LIQUIDADOR": ["FECHA", "ZONA", "RUTA", "CODIGO_DE_DESPACHO", "ESTADO_BEETRACK", "TIPO_DE_PAGO_BEETRACK", "MONTO_DE_PEDIDO", "MONTO_A_COBRAR", "MONTO_COBRADO", "DIF", "MONTO_PARCIAL", "OBSERVACIONES", "CODIGO_DE_COMERCIO", "N_COMERCIAL", "ESTADO_FINAL", "TIPO_DE_POS", "TIPO_DE_PAGO", "COD_AUTORIZACION", "FECHA_DE_PAGO"],
            "MONITOR": ["N/V_o_Porforma", "Numero_de_viaje", "Metodo_de_pago", "Estado", "Tipo_Doc_Cliente", "Nro_Doc_Cliente", "Centro_de_Distrib", "Total", "Cantidad", "Monto_Global", "Vtex_ID", "Dispatch_Number", "Tipo_de_facturacion", "Motivo"],
            "BBR": ["Fecha_Trx", "Medio_Pago", "SubTipo", "Cod_Cajero", "Cajero", "Caja", "N_Cuenta", "Monto"]
        }
        df_esquema = pd.DataFrame(columns=schema_maestros[maestro_sel])
        st.dataframe(df_esquema, use_container_width=True, hide_index=True)

    elif opcion_menu == "ARCHIVOS":
        st.markdown("<h2 style='color: #1E293B !important;'>📋 Historial de Archivos</h2>", unsafe_allow_html=True)
        st.write("Esta sección servirá para visualizar el log de archivos que ya han sido procesados y cargados al sistema histórico.")
        st.info("Para subir nuevos archivos y ejecutar el cruce contable, dirígete a la pestaña **CONCILIACIÓN**.")

    elif opcion_menu == "CONCILIACION":
        st.markdown("<h2 style='color: #1E293B !important;'>⚙️ Motor de Conciliación Automática</h2>", unsafe_allow_html=True)
        st.write("1. Selecciona los parámetros. 2. Sube los archivos del mes. 3. Ejecuta el cruce.")
        
        # Configuraciones de Ejecución
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            mes_sel_nombre = st.selectbox("📅 Mes a Conciliar", list(MESES_NOMBRES.values()), index=datetime.now().month - 1)
            mes_proceso = list(MESES_NOMBRES.keys())[list(MESES_NOMBRES.values()).index(mes_sel_nombre)]
        with col_c2:
            locales_seleccionados = st.multiselect("🏢 Locales a Procesar", list(LOCALES_CONFIG.keys()), default=list(LOCALES_CONFIG.keys()))

        st.markdown("---")
        
        # Área Unificada de Carga (Garantiza que no se borren al hacer clic)
        col_izq, col_der = st.columns(2)
        with col_izq:
            st.markdown("<h4 style='color: #1E293B !important;'>Fuentes de Ventas</h4>", unsafe_allow_html=True)
            f_jsat = st.file_uploader("Ventas JSatellite", type=['xlsx', 'csv'], accept_multiple_files=True)
            f_mon = st.file_uploader("Monitor Pedidos", type=['xlsx', 'csv'], accept_multiple_files=True)
            f_liq = st.file_uploader("Liquidador Beetrack", type=['xlsx', 'csv'], accept_multiple_files=True)
        with col_der:
            st.markdown("<h4 style='color: #1E293B !important;'>Pasarelas y Bancos</h4>", unsafe_allow_html=True)
            f_izi = st.file_uploader("Extracto IziPay", type=['xlsx', 'csv'], accept_multiple_files=True)
            f_bbr = st.file_uploader("Extracto BBR", type=['xlsx', 'csv'], accept_multiple_files=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Botón de Ejecución Principal
        if st.button("▶️ INICIAR PROCESO DE CONCILIACIÓN", type="primary", use_container_width=True):
            if not locales_seleccionados:
                st.warning("⚠️ Seleccione al menos un local para procesar.")
            else:
                with st.spinner("⏳ Procesando Inteligencia Contable (Esto puede tomar unos minutos)..."):
                    log_container = st.empty()
                    progress_bar = st.progress(0)
                    
                    try:
                        # Crear directorio temporal para guardar resultados a descargar
                        output_dir = os.path.join(BASE_DIR, f"OUTPUTS_{mes_sel_nombre}_{datetime.now().strftime('%H%M%S')}")
                        os.makedirs(output_dir, exist_ok=True)
                        
                        log_container.info("1/3 Consolidando y limpiando archivos subidos...")
                        df_jsat = leer_archivos_web(f_jsat, HEADERS_MAESTROS["JSATELITE"])
                        df_mon = leer_archivos_web(f_mon, HEADERS_MAESTROS["MONITOR"])
                        df_liq = leer_archivos_web(f_liq, HEADERS_MAESTROS["LIQUIDADOR"])
                        df_bbr = leer_archivos_web(f_bbr, HEADERS_MAESTROS["BBR"])
                        df_izi = leer_archivos_web(f_izi, HEADERS_MAESTROS["IZIPAY"])
                        
                        # Limpieza de fechas (Fase 2 adaptada)
                        def sanar_fechas(df, cols_fecha, mes):
                            if not df.empty:
                                for c in cols_fecha:
                                    if c in df.columns: df[c] = formatear_fecha_dt(df[c], mes)
                            return df
                            
                        df_jsat = sanar_fechas(df_jsat, ["FECHA"], mes_proceso)
                        df_mon = sanar_fechas(df_mon, ["Fecha"], mes_proceso)
                        df_liq = sanar_fechas(df_liq, ["FECHA", "FECHA DE PAGO"], mes_proceso)
                        df_bbr = sanar_fechas(df_bbr, ["Fecha Trx"], mes_proceso)
                        df_izi = sanar_fechas(df_izi, ["Fecha_Proceso", "Fecha_Lote", "Fecha_Consumo", "Fecha_Abono", "Fecha_Abono_8Dig"], mes_proceso)
                        
                        progress_bar.progress(30)
                        
                        # Fase 2 Normalización
                        log_container.info("2/3 Normalizando Identificadores de cruce (Vtex_ID, Códigos IziPay)...")
                        if not df_jsat.empty:
                            df_jsat['LOCAL_CLEAN'] = limpiar_id_texto(df_jsat.get('LOCAL', pd.Series(dtype=str)))
                            df_jsat['ID_M'] = limpiar_id_texto(df_jsat.get('ID_TRANSACCION', pd.Series(dtype=str)))
                            df_jsat['ORDEN_MERKAO_STR'] = limpiar_id_texto(df_jsat.get('ORDEN_MERKAO', pd.Series(dtype=str)))
                            df_jsat['FECHA_JS_DT'] = df_jsat.get('FECHA', pd.Series(dtype=str))
                        if not df_liq.empty:
                            df_liq['LIQ_M'] = limpiar_id_texto(df_liq.get('CODIGO DE DESPACHO', pd.Series(dtype=str))).str.split('-').str[0]
                            df_liq = df_liq.drop_duplicates(subset=['LIQ_M'], keep='last')
                            df_liq['FECHA_ENT_DT'] = df_liq.get('FECHA DE PAGO', pd.Series(dtype=str))
                            df_liq['COD_COM_L'] = limpiar_id_texto(df_liq.get('CODIGO DE COMERCIO', pd.Series(dtype=str)))
                        if not df_mon.empty:
                            df_mon['VTEX_M'] = limpiar_id_texto(df_mon.get('Vtex ID', pd.Series(dtype=str)))
                            df_mon = df_mon.drop_duplicates(subset=['VTEX_M'], keep='last')
                        if not df_bbr.empty:
                            df_bbr['BBR_M'] = limpiar_id_texto(df_bbr.get('N° Trx', pd.Series(dtype=str)))
                            df_bbr['FECHA_BBR_DT'] = df_bbr.get('Fecha Trx', pd.Series(dtype=str))
                        if not df_izi.empty:
                            df_izi['IZI_COD_M'] = limpiar_id_texto(df_izi.get('Codigo', pd.Series(dtype=str))).str.lstrip('0')
                            df_izi['FECHA_CONS_DT'] = df_izi.get('Fecha_Consumo', pd.Series(dtype=str))

                        progress_bar.progress(50)
                        
                        # Fase 3: Cruce por Local
                        log_container.info("3/3 Ejecutando algoritmos de cruce financiero (Match Automático)...")
                        
                        archivos_generados = []
                        total_locales = len(locales_seleccionados)
                        
                        for i, local_id in enumerate(locales_seleccionados):
                            config_local = LOCALES_CONFIG[local_id]
                            ruta_fin = os.path.join(output_dir, f"Rosa {mes_sel_nombre} - {local_id}.xlsx")
                            
                            # Si las plantillas están cargadas en GitHub, usamos esas, si no, creamos un Excel vacío de soporte
                            try: shutil.copyfile(config_local["plantilla_fin"], ruta_fin)
                            except: pd.DataFrame().to_excel(ruta_fin)

                            codigos_izipay = config_local["codigos"]
                            df_f = df_jsat[(df_jsat['LOCAL_CLEAN'] == str(local_id))].copy() if not df_jsat.empty else pd.DataFrame()
                            
                            if not df_f.empty:
                                # Cruce Monitor
                                if not df_mon.empty:
                                    df_f = pd.merge(df_f, df_mon[['VTEX_M', 'N/V o Porforma']], left_on='ORDEN_MERKAO_STR', right_on='VTEX_M', how='left')
                                    df_f['PED_M'] = limpiar_id_texto(df_f.get('N/V o Porforma', pd.Series(dtype=str)))
                                else: df_f['PED_M'] = ""
                                
                                # Cruce Liquidador
                                if not df_liq.empty:
                                    cols_liq_req = ['LIQ_M', 'FECHA_ENT_DT', 'ESTADO BEETRACK', 'MONTO A COBRAR', 'MONTO COBRADO', 'TIPO DE PAGO', 'CODIGO DE COMERCIO', 'OBSERVACIONES', 'COD AUTORIZACION']
                                    for col in cols_liq_req:
                                        if col not in df_liq.columns: df_liq[col] = "" 
                                    df_f = pd.merge(df_f, df_liq[cols_liq_req], left_on='PED_M', right_on='LIQ_M', how='left')
                                else:
                                    for col in ['FECHA_ENT_DT', 'ESTADO BEETRACK', 'MONTO A COBRAR', 'MONTO COBRADO', 'TIPO DE PAGO', 'CODIGO DE COMERCIO', 'OBSERVACIONES', 'COD AUTORIZACION']: df_f[col] = ""
                                df_f = df_f.drop_duplicates(subset=['ORDEN_MERKAO_STR'], keep='first')

                            # Importes BBR
                            importes_bbr_p1 = []
                            if not df_bbr.empty and 'Monto' in df_bbr.columns: df_bbr['Monto_Limpio'] = df_bbr['Monto'].astype(str).str.replace(',', '.')
                            else: df_bbr['Monto_Limpio'] = '0'
                            
                            for _, r in df_f.iterrows():
                                if not df_bbr.empty and pd.notna(r.get('ID_M')) and str(r.get('ID_M')).strip() != '':
                                    suma = pd.to_numeric(df_bbr[df_bbr['BBR_M'] == str(r['ID_M']).strip()]['Monto_Limpio'], errors='coerce').sum()
                                    importes_bbr_p1.append(suma if suma > 0 else 0)
                                else: importes_bbr_p1.append(0)
                            df_f['IMPORTE_BBR'] = pd.Series(importes_bbr_p1)

                            # Cálculos Diferenciales
                            cobrado_num = pd.to_numeric(df_f.get('MONTO COBRADO', pd.Series(0, index=df_f.index)), errors='coerce').fillna(0)
                            col_a_cobrar = pd.to_numeric(df_f.get('MONTO A COBRAR', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            col_total = pd.to_numeric(df_f.get('TOTAL', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            df_f['DIFF_O'] = col_a_cobrar - cobrado_num
                            df_f['DIFF_P'] = col_total - cobrado_num
                            
                            # Filtro Izipay
                            if not df_izi.empty:
                                codigos_str = [str(c).strip().lstrip('0') for c in codigos_izipay]
                                df_izi_f = df_izi[df_izi['IZI_COD_M'].isin(codigos_str)].copy() if codigos_izipay else df_izi.copy()
                            else: df_izi_f = pd.DataFrame()
                            df_izi_f['A_FECHA'], df_izi_f['B_TIPO'] = "", ""
                            usados_izi = set()

                            # Match Engine Izipay
                            for _, r_p1 in df_f.iterrows():
                                f_ent, tipo = r_p1.get('FECHA_JS_DT', ''), str(r_p1.get('TIPO DE PAGO', '')).upper()
                                montos = extraer_montos_regex(r_p1.get('OBSERVACIONES', '')) if "MÚLTIPLE" in tipo or "MULTIPLE" in tipo else [str(r_p1.get('MONTO COBRADO', ''))]
                                codigos_auth = extraer_codigos_autorizacion(r_p1.get('COD AUTORIZACION', ''))
                                cod_comercio = str(r_p1.get('CODIGO DE COMERCIO', '')).strip().lstrip('0') 
                                
                                for m in montos:
                                    try:
                                        mask = (df_izi_f['IZI_COD_M'] == cod_comercio) & (pd.to_numeric(df_izi_f['Importe'], errors='coerce') == float(m)) & (~df_izi_f.index.isin(usados_izi))
                                        candidatos = df_izi_f[mask].copy()
                                        if not candidatos.empty:
                                            match_fuerte = candidatos[candidatos['Autorizacion'].astype(str).str.strip().isin(codigos_auth)]
                                            if not match_fuerte.empty: idx = match_fuerte.index[0]
                                            else: idx = candidatos['FECHA_CONS_DT'].apply(lambda d: abs((pd.to_datetime(d) - pd.to_datetime(f_ent)).days) if pd.notnull(d) and pd.notnull(f_ent) else 999).idxmin()
                                            df_izi_f.at[idx, 'A_FECHA'], df_izi_f.at[idx, 'B_TIPO'] = f_ent, tipo
                                            usados_izi.add(idx)
                                    except: continue

                            # Matriz final Tesoreria
                            df_p1_u = df_f[df_f.get('TIPO DE PAGO', pd.Series(dtype=str)).astype(str).str.contains("ÚNICO|UNICO", case=False, na=False, regex=True)].copy()
                            df_p2_u = df_izi_f[df_izi_f.get('B_TIPO', pd.Series(dtype=str)).astype(str).str.contains("ÚNICO|UNICO", case=False, na=False, regex=True)].copy()
                            
                            df_p1_u['MONTO COBRADO'] = pd.to_numeric(df_p1_u.get('MONTO COBRADO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            df_p2_u['Importe'] = pd.to_numeric(df_p2_u.get('Importe', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            
                            matched_rows, unmatched_p1, used_p2_idx = [], [], set()
                            for _, r1 in df_p1_u.iterrows():
                                monto, cod = r1['MONTO COBRADO'], str(r1.get('CODIGO DE COMERCIO', '')).strip().lstrip('0') 
                                mask = (df_p2_u['Importe'] == monto) & (df_p2_u.get('Codigo', pd.Series(dtype=str)).astype(str).str.strip().str.lstrip('0') == cod) & (~df_p2_u.index.isin(used_p2_idx)) 
                                cands = df_p2_u[mask]
                                if not cands.empty:
                                    b_idx = cands.index[0]
                                    used_p2_idx.add(b_idx)
                                    combo = r1.to_dict(); combo.update(cands.loc[b_idx].to_dict())
                                    matched_rows.append(combo)
                                else: unmatched_p1.append(r1.to_dict())
                                    
                            unmatched_p2 = df_p2_u[~df_p2_u.index.isin(used_p2_idx)].to_dict('records')
                            df_final = pd.DataFrame(matched_rows + unmatched_p1 + unmatched_p2) 

                            bbr_match_final = []
                            for _, r in df_final.iterrows():
                                if pd.notna(r.get('ID_M')) and str(r.get('ID_M')).strip() != '' and not df_bbr.empty:
                                    m = df_bbr[df_bbr['BBR_M'] == str(r['ID_M']).strip()]
                                    bbr_match_final.append({'CTA': m.iloc[0]['N Cuenta'] if not m.empty else 0, 'MTO': pd.to_numeric(m.iloc[0]['Monto_Limpio'], errors='coerce') if not m.empty else 0})
                                else: bbr_match_final.append({'CTA': 0, 'MTO': 0})
                                    
                            df_bbr_res = pd.DataFrame(bbr_match_final)
                            df_final['N_CTA_BBR'] = df_bbr_res['CTA'] if not df_bbr_res.empty else 0
                            df_final['MTO_BBR'] = df_bbr_res['MTO'] if not df_bbr_res.empty else 0
                            df_final['DIF_Q'] = pd.to_numeric(df_final.get('Importe', pd.Series(dtype=float)), errors='coerce').fillna(0) - pd.to_numeric(df_final['MTO_BBR'], errors='coerce').fillna(0) 
                            df_final['T_NOM'] = df_final.get('Tarjeta', pd.Series(dtype=str)).apply(obtener_tipo_nombre)
                            df_final['T_CODE'] = df_final.get('Tarjeta', pd.Series(dtype=str)).apply(clasificar_tipo_codigo)

                            def det_cambio(row):
                                if pd.to_numeric(row.get('MONTO COBRADO', 0), errors='coerce') == 0 or pd.isna(row.get('MONTO COBRADO', 0)): return "REVISAR"
                                if row.get('DIF_Q', 0) != 0: return "MANUAL"
                                if row.get('DIF_Q', 0) == 0: return "NO" if row.get('T_CODE', '') == 88 else "SI"
                                return "MANUAL"

                            df_final['CAMBIO_R'] = df_final.apply(det_cambio, axis=1) if not df_final.empty else pd.Series(dtype=str)
                            
                            # Guardado Local en Output_dir
                            with pd.ExcelWriter(ruta_fin, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer_fin:
                                if not df_final.empty:
                                    cols_p1 = [c for c in ['ID_TRANSACCION', 'FECHA_JS_DT', 'MONTO COBRADO', 'CODIGO DE COMERCIO'] if c in df_final.columns]
                                    cols_p2 = [c for c in ['A_FECHA', 'Autorizacion', 'Tarjeta', 'Importe', 'Nro_ID', 'Codigo'] if c in df_final.columns]
                                    cols_res = [c for c in ['T_NOM', 'T_CODE', 'N_CTA_BBR', 'MTO_BBR', 'DIF_Q', 'CAMBIO_R'] if c in df_final.columns]
                                    df_final[cols_p1].to_excel(writer_fin, sheet_name=str(local_id), index=False, header=False, startrow=1)
                                    df_final[cols_p2].to_excel(writer_fin, sheet_name=str(local_id), index=False, header=False, startrow=1, startcol=6)
                                    df_final[cols_res].to_excel(writer_fin, sheet_name=str(local_id), index=False, header=False, startrow=1, startcol=12)

                                    df_rosa = df_final[df_final['CAMBIO_R'] == "SI"].copy()
                                    if not df_rosa.empty:
                                        df_rosa['LOCAL_COL'], df_rosa['CAJERO_COL'] = local_id, 993
                                        cols_rosa = [c for c in ['FECHA_JS_DT', 'LOCAL_COL', 'CAJERO_COL', 'ID_TRANSACCION', 'Tarjeta', 'Autorizacion', 'MONTO COBRADO', 'Nro_ID', 'T_CODE'] if c in df_rosa.columns]
                                        df_rosa[cols_rosa].to_excel(writer_fin, sheet_name="Rosa", index=False, header=False, startrow=1)
                                else:
                                    pd.DataFrame(columns=["Sin Datos Cruzados"]).to_excel(writer_fin, sheet_name=str(local_id), index=False)
                            
                            archivos_generados.append(ruta_fin)
                            progress_bar.progress(50 + int((i + 1) / total_locales * 45))

                        # Empaquetar todo en un archivo ZIP para descarga web
                        zip_path = os.path.join(output_dir, f"Conciliacion_Merkao_{mes_sel_nombre}.zip")
                        with zipfile.ZipFile(zip_path, 'w') as zipf:
                            for file in archivos_generados:
                                zipf.write(file, os.path.basename(file))
                                
                        progress_bar.progress(100)
                        log_container.success("🎉 ¡Conciliación procesada exitosamente!")
                        
                        # Generar el botón de descarga del ZIP
                        with open(zip_path, "rb") as fp:
                            st.download_button(
                                label="📦 DESCARGAR RESULTADOS (ZIP)",
                                data=fp,
                                file_name=f"Reportes_Rosa_{mes_sel_nombre}.zip",
                                mime="application/zip",
                                type="primary"
                            )

                    except Exception as e:
                        log_container.error(f"❌ Ocurrió un error en el cruce de datos: {str(e)}")

import streamlit as st
import smtplib
import base64
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# --- 1. CONFIGURACIÓN BÁSICA ---
st.set_page_config(page_title="Tesoriapp", page_icon="🏦", layout="wide", initial_sidebar_state="expanded")

# --- FUNCION PARA LEER IMAGEN LOCAL ---
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return ""

# --- 2. ESQUEMAS DE SHAREPOINT (METADATA) ---
SCHEMA_MAESTROS = {
    "JSATELITE": ["ORDEN_MERKAO", "COD_ENVIADO_CT2", "BOLETA", "TOTAL", "ID_TRANSACCION", "CAJA", "LOCAL", "FECHA", "Fecha_carga", "N_Trx"],
    "IZIPAY": ["Codigo", "Producto", "Tipo_Mov", "Fecha_Proceso", "Fecha_Lote", "Lote_Manual", "Lote_Pos", "Terminal", "Voucher", "Autorizacion", "Cuotas", "Tarjeta", "Origen", "Transaccion", "Fecha_Consumo", "Importe", "Status", "Comision", "Comision_Afecta", "IGV", "Neto_Parcial", "Neto_Total", "Fecha_Abono", "Fecha_Abono_8Dig", "Observaciones", "ExtraComision", "Comis_Standar", "Comis_%", "Nro_ID", "Tpo_Comprob", "Nro_Comprob", "Fecha_carga"],
    "LIQUIDADOR": ["FECHA", "ZONA", "RUTA", "CODIGO_DE_DESPACHO", "ESTADO_BEETRACK", "TIPO_DE_PAGO_BEETRACK", "MONTO_DE_PEDIDO", "MONTO_A_COBRAR", "MONTO_COBRADO", "DIF", "MONTO_PARCIAL", "OBSERVACIONES", "CODIGO_DE_COMERCIO", "N_COMERCIAL", "ESTADO_FINAL", "TIPO_DE_POS", "TIPO_DE_PAGO", "COD_AUTORIZACION", "FECHA_DE_PAGO"],
    "MONITOR": ["N/V_o_Porforma", "Numero_de_viaje", "Metodo_de_pago", "Estado", "Tipo_Doc_Cliente", "Nro_Doc_Cliente", "Centro_de_Distrib", "Total", "Cantidad", "Monto_Global", "Vtex_ID", "Dispatch_Number", "Tipo_de_facturacion", "Motivo"],
    "BBR": ["Fecha_Trx", "Medio_Pago", "SubTipo", "Cod_Cajero", "Cajero", "Caja", "N_Cuenta", "Monto"]
}

# --- 3. INICIALIZACIÓN DE MEMORIA DE SESIÓN ---
if "usuario_autenticado" not in st.session_state:
    st.session_state["usuario_autenticado"] = False
if "vista_actual" not in st.session_state:
    st.session_state["vista_actual"] = "login"
if "usuario_nombre" not in st.session_state:
    st.session_state["usuario_nombre"] = ""

# --- FUNCIÓN DE ENVÍO DE CORREO REAL ---
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
        return True, f"✅ Se ha enviado un correo con instrucciones a {destinatario}."
    except Exception as e:
        return False, f"❌ Error del sistema: Verifica credenciales en Streamlit Secrets."


# =================================================================================
# ======================== ZONA A: PANTALLA DE LOGIN ==============================
# =================================================================================
if not st.session_state["usuario_autenticado"]:
    
    st.markdown("""
    <style>
        #MainMenu, header, footer, [data-testid="stToolbar"] {visibility: hidden;}
        .st-emotion-cache-1r6slb0 { padding: 0 !important; }
        
        [data-testid="stAppViewContainer"] {
            background: radial-gradient(circle, #283593 0%, #1A237E 70%, #121858 100%);
        }
        
        div[data-testid="stColumn"]:nth-of-type(2) > div {
            background-color: white; 
            border-radius: 20px; 
            padding: 50px 40px; 
            box-shadow: 0 20px 40px rgba(0,0,0,0.5); 
            max-width: 450px; 
            margin: auto; 
        }
        
        /* Ajuste de Inputs para que el texto sea negro y legible */
        .stTextInput>div>div>input {
            background-color: #FFFFFF !important; 
            border: 1.5px solid #CBD5E1 !important;
            color: #000000 !important; /* TEXTO NEGRO */
            font-weight: 500 !important;
            border-radius: 8px !important; 
            padding: 14px 16px !important;
        }
        .stTextInput>div>div>input::placeholder {
            color: #64748B !important; 
            opacity: 1 !important;
        }
        
        /* Ajuste de Botones: Una sola línea, anchos y centrados */
        .stButton { 
            display: flex; 
            justify-content: center; 
            width: 100%; 
            margin-top: 10px; 
        }
        .stButton>button[kind="primary"] {
            background-color: #FFC20E !important; 
            color: #1A237E !important; 
            font-weight: 800 !important;
            border-radius: 8px !important; 
            width: 100% !important; /* Estirado al máximo del cuadro */
            padding: 12px !important;
            white-space: nowrap !important; /* Fuerza una sola línea */
        }
        
        .boton-olvido { 
            display: flex; 
            justify-content: center !important; 
            margin: 5px 0 10px 0; 
        }
        .stButton>button[kind="secondary"] { 
            color: #64748B !important; 
            border: none !important; 
        }
        
        [data-testid="stCheckbox"] { 
            display: flex; 
            justify-content: center; 
            margin-bottom: 15px; 
        }
        [data-testid="stCheckbox"] p { 
            color: #1E293B !important; 
            font-weight: 500 !important; 
        }
    </style>
    """, unsafe_allow_html=True)

    c1, card_col, c3 = st.columns([1, 3, 1])
    with card_col:
        st.write("<br>", unsafe_allow_html=True) 
        
        # --- VISTA 1: INICIO DE SESIÓN ---
        if st.session_state["vista_actual"] == "login":
            
            # Logo y Tesoriapp más grandes
            img_b64 = get_base64_image("logo_merkao.png")
            if img_b64:
                st.markdown(f"""
                    <div style='display: flex; justify-content: center; align-items: center; gap: 20px; margin-bottom: 35px;'>
                        <img src='data:image/png;base64,{img_b64}' style='width: 100px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                        <h1 style='color: #1A237E !important; margin: 0; font-size: 42px; font-weight: 900; letter-spacing: -1px;'>Tesoriapp</h1>
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("<h1 style='text-align: center; color: #1A237E !important; font-weight: 900; font-size: 46px; margin-bottom: 30px;'>Tesoriapp</h1>", unsafe_allow_html=True)
            
            st.text_input("Usuario", placeholder="ej. MAYHELA.SIMON", key="usuario_input")
            st.text_input("Contraseña", type="password", placeholder="••••••••", key="password_input")
            
            st.markdown("<div class='boton-olvido'>", unsafe_allow_html=True)
            if st.button("¿Olvidaste tu contraseña?", type="secondary"):
                st.session_state["vista_actual"] = "recuperar"; st.rerun() 
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

        # --- VISTA 2: RECUPERAR CONTRASEÑA ---
        elif st.session_state["vista_actual"] == "recuperar":
            
            img_b64 = get_base64_image("logo_merkao.png")
            if img_b64:
                st.markdown(f"<div style='text-align: center; margin-bottom: 20px;'><img src='data:image/png;base64,{img_b64}' style='width: 100px; border-radius: 12px;'></div>", unsafe_allow_html=True)

            st.markdown("<h3 style='text-align: center; color: #1A237E !important; font-weight: 800;'>Recuperar Acceso</h3>", unsafe_allow_html=True)
            st.text_input("Correo corporativo", placeholder="ejemplo@spsa.pe", key="correo_recuperar_input")
            
            st.markdown("<div class='boton-olvido'>", unsafe_allow_html=True)
            if st.button("Cancelar y Volver", type="secondary"):
                st.session_state["vista_actual"] = "login"; 
                if "pwd_recuperada_msg" in st.session_state: del st.session_state["pwd_recuperada_msg"]
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
                
            # Corrección del cuadro negro inmenso (Bug de Streamlit Magic)
            if "pwd_recuperada_msg" in st.session_state:
                if "✅" in st.session_state["pwd_recuperada_msg"]:
                    st.success(st.session_state["pwd_recuperada_msg"])
                else:
                    st.error(st.session_state["pwd_recuperada_msg"])

            def intentar_recuperacion():
                correo = st.session_state.correo_recuperar_input.strip()
                if not correo or "@" not in correo: st.session_state["pwd_recuperada_msg"] = "⚠️ Ingresa un correo válido."
                else: _, st.session_state["pwd_recuperada_msg"] = enviar_correo_recuperacion(correo)
            
            st.button("ENVIAR INSTRUCCIONES", on_click=intentar_recuperacion, type="primary")


# =================================================================================
# ======================== ZONA B: APLICACIÓN PRINCIPAL (DASHBOARD) ===============
# =================================================================================
else:
    # --- CSS ESPECÍFICO PARA EL DASHBOARD ---
    st.markdown("""
        <style>
            /* Fondo general claro */
            [data-testid="stAppViewContainer"] { background: #F4F7F9; }
            #MainMenu, footer {visibility: hidden;}
            
            /* Hacer la cabecera transparente para no ocultar el menú de hamburguesa */
            header { background: transparent !important; }
            
            /* Títulos forzados a color oscuro para que no se pierdan en el fondo blanco */
            .stMainBlockContainer h1, 
            .stMainBlockContainer h2, 
            .stMainBlockContainer h3, 
            .stMainBlockContainer p, 
            .stMainBlockContainer label { 
                color: #1E293B !important; 
            }
            
            /* TOP BAR PERSONALIZADA */
            .top-bar {
                background-color: white;
                padding: 15px 30px 15px 60px; /* Margen izquierdo extra para la hamburguesa */
                border-bottom: 2px solid #E2E8F0;
                display: flex;
                align-items: center;
                gap: 15px;
                margin-top: -80px;
                margin-bottom: 30px;
                margin-left: -50px;
                margin-right: -50px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            }

            /* ESTILIZAR EL SIDEBAR (Color Azul Oscuro) */
            [data-testid="stSidebar"] {
                background-color: #2C3E50 !important; 
                border-right: none;
            }
            [data-testid="stSidebar"] * {
                color: #FFFFFF !important; 
            }
            
            /* Estilizar botones del menú */
            div.row-widget.stRadio > div { gap: 15px; padding: 10px 0; }
            div.row-widget.stRadio > div > label {
                background-color: transparent; padding: 10px 15px; border-radius: 8px; transition: 0.3s; cursor: pointer;
            }
            div.row-widget.stRadio > div > label:hover { background-color: rgba(255,255,255,0.1); }
            div.row-widget.stRadio > div > label > div:first-child { display: none; }
            div.row-widget.stRadio > div > label > div:last-child p { font-size: 16px; font-weight: 600; margin: 0; }

            /* Botón de Salir en el Sidebar */
            [data-testid="stSidebar"] button[kind="secondary"] {
                background-color: rgba(220, 38, 38, 0.1) !important; border: 1px solid #DC2626 !important;
                color: #FCA5A5 !important; width: 100%;
            }
            [data-testid="stSidebar"] button[kind="secondary"]:hover {
                background-color: #DC2626 !important; color: white !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # --- TOP BAR (BARRA SUPERIOR BLANCA) ---
    img_b64 = get_base64_image("logo_merkao.png")
    logo_html = f"<img src='data:image/png;base64,{img_b64}' width='45' style='border-radius: 6px;'>" if img_b64 else "🏦"
    st.markdown(f"""
        <div class="top-bar">
            {logo_html}
            <h1 style="color: #1E293B !important; font-size: 28px !important; font-weight: 800 !important; margin: 0 !important;">Tesoriapp - Módulo de Conciliación</h1>
        </div>
    """, unsafe_allow_html=True)

    # --- SIDEBAR (MENÚ LATERAL AZUL OSCURO) ---
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 20px; margin-bottom: 20px; color: white !important;'>Menú Principal</h2>", unsafe_allow_html=True)
        
        opcion_menu = st.radio(
            "Navegación",
            ["MAESTROS", "ARCHIVOS", "CONCILIACION"],
            label_visibility="collapsed"
        )
        
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        st.markdown(f"**Bienvenid@**<br>{st.session_state['usuario_nombre']}", unsafe_allow_html=True)
        st.markdown("---")
        
        if st.button("Cerrar Sesión", use_container_width=True):
            st.session_state["usuario_autenticado"] = False
            st.rerun()

    # --- CONTENIDO DE LAS PÁGINAS ---

    if opcion_menu == "MAESTROS":
        st.markdown("<h2 style='color: #1E293B !important;'>📁 Base de Datos: Maestros SharePoint</h2>", unsafe_allow_html=True)
        st.write("Visualización de las listas maestras sincronizadas en la nube de Super Food Holding.")
        
        maestro_sel = st.selectbox("Seleccione el Maestro a visualizar:", ["JSATELITE", "IZIPAY", "LIQUIDADOR", "MONITOR", "BBR"])
        
        st.info(f"Mostrando estructura de la lista **{maestro_sel}** alojada en SharePoint:")
        
        df_esquema = pd.DataFrame(columns=SCHEMA_MAESTROS[maestro_sel])
        st.dataframe(df_esquema, use_container_width=True, hide_index=True)
        st.caption(f"📍 Rutas destino conectadas: OUTPUTS CONCILIACIÓN | PLANTILLA")

    elif opcion_menu == "ARCHIVOS":
        st.markdown("<h2 style='color: #1E293B !important;'>📥 Carga de Archivos Temporales</h2>", unsafe_allow_html=True)
        st.write("Cargue los extractos descargados de los portales para alimentar el motor de cruce.")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<h4 style='color: #1E293B !important;'>Sistemas Internos (Ventas/Logística)</h4>", unsafe_allow_html=True)
            f_jsat = st.file_uploader("Ventas JSatellite", type=['xlsx', 'csv'])
            f_mon = st.file_uploader("Monitor Pedidos", type=['xlsx', 'csv'])
            f_liq = st.file_uploader("Liquidador Beetrack", type=['xlsx', 'csv'])
        with col2:
            st.markdown("<h4 style='color: #1E293B !important;'>Pasarelas y Bancos</h4>", unsafe_allow_html=True)
            f_izi = st.file_uploader("Extracto IziPay", type=['xlsx', 'csv'])
            f_bbr = st.file_uploader("Extracto BBR", type=['xlsx', 'csv'])

    elif opcion_menu == "CONCILIACION":
        st.markdown("<h2 style='color: #1E293B !important;'>⚙️ Motor de Conciliación Automática</h2>", unsafe_allow_html=True)
        st.write("Este proceso cruzará los archivos cargados contra los maestros históricos de SharePoint y generará el reporte final (Archivo Rosa).")
        
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            st.button("▶️ INICIAR PROCESO DE CONCILIACIÓN", type="primary", use_container_width=True)
            st.caption("Al finalizar, el reporte se guardará automáticamente en la carpeta OUTPUTS CONCILIACIÓN.")

import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# --- 1. CONFIGURACIÓN BÁSICA ---
st.set_page_config(page_title="Tesoriapp - Acceso", page_icon="🏦", layout="centered")

# --- 2. DISEÑO UI AVANZADO (CSS UX/UI PROFESIONAL - TESORIAPP STYLE) ---
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden;}
    
    .st-emotion-cache-1r6slb0 { padding: 0 !important; }

    [data-testid="stAppViewContainer"] {
        background: radial-gradient(circle, #283593 0%, #1A237E 70%, #121858 100%);
        display: flex;
        justify-content: center;
        align-items: center;
    }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }

    div[data-testid="stColumn"]:nth-of-type(2) > div {
        background-color: white;
        border-radius: 20px; 
        padding: 60px 50px; 
        box-shadow: 0 15px 35px rgba(0,0,0,0.4); 
        border: 1px solid #EAEAEA;
        max-width: 450px; 
        margin: auto; 
    }

    div[data-testid="stColumn"]:nth-of-type(2) label, 
    div[data-testid="stColumn"]:nth-of-type(2) p, 
    div[data-testid="stColumn"]:nth-of-type(2) h1, 
    div[data-testid="stColumn"]:nth-of-type(2) h2, 
    div[data-testid="stColumn"]:nth-of-type(2) h3 {
        color: #1E293B !important; 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        text-align: left;
    }
    
    .titulo-recuperar {
        text-align: center !important;
        margin-bottom: 25px;
        color: #1A237E !important;
    }
    
    .stTextInput>div>div>input {
        background-color: #FBFDFF !important;
        border: 1px solid #CBD5E1 !important;
        color: #1E293B !important;
        border-radius: 10px !important;
        padding: 12px 18px !important;
        font-size: 15px;
    }
    .stTextInput>div>div>input:focus {
        border-color: #1A237E !important;
        box-shadow: 0 0 0 3px rgba(26, 35, 126, 0.1) !important;
    }

    .stButton>button[kind="primary"] {
        background-color: #FFC20E !important; 
        color: #1A237E !important; 
        font-weight: 800 !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        border: none !important;
        width: 100% !important; 
        padding: 14px !important;
        margin-top: 30px !important;
        box-shadow: 0 5px 10px rgba(0,0,0,0.15) !important;
        transition: all 0.3s ease !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #FFD54F !important; 
        transform: translateY(-3px); 
        box-shadow: 0 8px 15px rgba(0,0,0,0.2) !important;
    }

    .stButton>button[kind="secondary"] {
        background-color: transparent !important;
        color: #475569 !important; 
        border: none !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        padding: 0 !important;
        margin-top: 5px !important;
        transition: color 0.3s ease;
    }
    .stButton>button[kind="secondary"]:hover {
        color: #1A237E !important; 
        text-decoration: underline !important;
    }
    
    .boton-olvido {
        text-align: right !important;
        display: flex;
        justify-content: flex-end;
        margin-top: -15px; 
        margin-bottom: 20px;
    }

    .stCheckbox label {
        font-size: 14px !important;
        color: #64748B !important;
        margin-top: 10px;
    }

    .titulo-elegante {
        text-align: center;
        color: #1A237E !important;
        font-weight: 800;
        font-size: 36px;
        margin-bottom: 25px;
    }
</style>
""", unsafe_allow_html=True)

# --- 3. INICIALIZACIÓN DE MEMORIA DE SESIÓN ---
if "usuario_autenticado" not in st.session_state:
    st.session_state["usuario_autenticado"] = False

if "vista_actual" not in st.session_state:
    st.session_state["vista_actual"] = "login" 

# --- FUNCIÓN DE ENVÍO DE CORREO REAL (OFFICE 365) ---
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
                <p>Hola,</p>
                <p>Se ha solicitado la recuperación de credenciales para tu usuario.</p>
                <p>Tu contraseña corporativa provisional es: <strong>Spsa2026</strong></p>
                <br>
                <p>Por favor, ingresa al portal con estas credenciales.</p>
                <p>Saludos,<br>Equipo de Riesgos CI / Tesorería</p>
            </body>
        </html>
        """
        msg.attach(MIMEText(cuerpo_html, 'html'))
        
        # Conexión al servidor SMTP de Microsoft
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        
        return True, f"✅ Se ha enviado un correo con instrucciones a {destinatario}."
    except smtplib.SMTPAuthenticationError:
        return False, "❌ Autenticación fallida: Microsoft bloqueó el acceso. Pide a TI una 'Contraseña de Aplicación' para saltar el MFA."
    except Exception as e:
        return False, f"❌ Error del sistema de correo: {str(e)}"

# --- 4. PANTALLA DE ACCESO (DISEÑO UX/UI DE TARJETA) ---
if not st.session_state["usuario_autenticado"]:
    
    c1, card_col, c3 = st.columns([1, 3, 1])
    
    with card_col:
        st.write("<br><br>", unsafe_allow_html=True) 
        st.image("logo_merkao.png", width=130)
        
        # === VISTA 1: INICIO DE SESIÓN ===
        if st.session_state["vista_actual"] == "login":
            st.text_input("Usuario", placeholder="NOMBRE.APELLIDO", key="usuario_input")
            st.text_input("Contraseña", type="password", placeholder="••••••••", key="password_input")
            
            st.markdown("<div class='boton-olvido'>", unsafe_allow_html=True)
            if st.button("¿Olvidaste tu contraseña?", kind="secondary", key="btn_ir_recuperar"):
                st.session_state["vista_actual"] = "recuperar"
                st.rerun() 
            st.markdown("</div>", unsafe_allow_html=True)
            
            st.checkbox("Recordar en este dispositivo", value=True)
            
            if "error_login" in st.session_state:
                st.error(st.session_state["error_login"])

            def verificar_login():
                usu = st.session_state.usuario_input.strip().upper() 
                pwd = st.session_state.password_input
                
                try:
                    if st.secrets["passwords"].get(usu) == pwd:
                        st.session_state["usuario_autenticado"] = True
                        if "error_login" in st.session_state: del st.session_state["error_login"]
                    else:
                        st.session_state["error_login"] = "❌ Usuario o contraseña incorrectos."
                except Exception:
                    st.session_state["error_login"] = "❌ Error crítico: La bóveda de seguridad no está configurada."

            st.button("Iniciar Sesión", on_click=verificar_login, kind="primary")
            
        # === VISTA 2: RECUPERAR ACCESO ===
        elif st.session_state["vista_actual"] == "recuperar":
            
            st.markdown("<h3 class='titulo-recuperar'>Recuperar Acceso</h3>", unsafe_allow_html=True)
            st.markdown("<p style='font-size: 14px; color: #475569 !important; margin-bottom: 20px;'>Escribe tu correo corporativo y te enviaremos las instrucciones de acceso.</p>", unsafe_allow_html=True)
            
            st.text_input("Correo Electrónico", placeholder="ejemplo@spsa.pe", key="correo_recuperar_input")
            
            if st.button("Volver al Login", kind="secondary", key="btn_volver_login"):
                st.session_state["vista_actual"] = "login"
                if "pwd_recuperada_msg" in st.session_state: del st.session_state["pwd_recuperada_msg"]
                st.rerun()
                
            if "pwd_recuperada_msg" in st.session_state:
                if "❌" in st.session_state["pwd_recuperada_msg"]:
                    st.error(st.session_state["pwd_recuperada_msg"])
                else:
                    st.success(st.session_state["pwd_recuperada_msg"])

            def intentar_recuperacion():
                correo = st.session_state.correo_recuperar_input.strip()
                if not correo or "@" not in correo:
                    st.session_state["pwd_recuperada_msg"] = "⚠️ Ingresa un correo válido."
                else:
                    # Llamamos a la función real de envío de correos
                    exito, mensaje = enviar_correo_recuperacion(correo)
                    st.session_state["pwd_recuperada_msg"] = mensaje
            
            st.button("Enviar Instrucciones", on_click=intentar_recuperacion, kind="primary")

# === 5. PANTALLA PRINCIPAL (MOTOR CONTABLE - OCULTO) ===
else:
    st.markdown("""
        <style>
            [data-testid="stAppViewContainer"] { background: white; }
            h1, h2, h3, p, label { color: #1E293B !important; text-align: left; }
        </style>
    """, unsafe_allow_html=True)

    c_logo, c_titulo, c_logout = st.columns([1, 4, 1])
    with c_logo:
        try: st.image("logo_merkao.png", width=60)
        except: st.write("🏦")
    with c_titulo:
        st.markdown("<h2>Módulo de Conciliación - Tesoriapp</h2>", unsafe_allow_html=True)
    with c_logout:
        if st.button("Salir", key="btn_logout"):
            st.session_state["usuario_autenticado"] = False
            st.rerun()
        
    st.markdown("---")
    
    st.write("### 📁 Carga de Archivos Maestros para Procesamiento")
    st.info("💡 Sube los archivos consolidados del mes. El motor procesará los cruces automáticamente.")
    
    col_izq, col_der = st.columns(2)
    
    with col_izq:
        st.markdown("#### Fuentes de Ventas")
        file_jsat = st.file_uploader("JSatellite", type=['xlsx', 'csv'])
        file_mon = st.file_uploader("Monitor (Pedidos)", type=['xlsx', 'csv'])
        file_liq = st.file_uploader("Liquidador (Entregas)", type=['xlsx', 'csv'])

    with col_der:
        st.markdown("#### Fuentes Bancarias / Pasarelas")
        file_izi = st.file_uploader("IziPay", type=['xlsx', 'csv'])
        file_bbr = st.file_uploader("BBR", type=['xlsx', 'csv'])
    
    st.markdown("---")
    st.button("▶️ INICIAR CONCILIACIÓN", type="primary", key="btn_procesar_web")

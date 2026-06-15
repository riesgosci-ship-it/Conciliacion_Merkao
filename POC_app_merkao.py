import streamlit as st
import smtplib
import base64
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# --- 1. CONFIGURACIÓN BÁSICA ---
st.set_page_config(page_title="Tesoriapp - Acceso", page_icon="🏦", layout="centered")

# --- FUNCION PARA LEER IMAGEN LOCAL ---
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return ""

# --- 2. DISEÑO UI AVANZADO (CSS UX/UI PROFESIONAL - TESORIAPP STYLE) ---
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden;}
    
    .st-emotion-cache-1r6slb0 { padding: 0 !important; }

    /* Fondo degradado */
    [data-testid="stAppViewContainer"] {
        background: radial-gradient(circle, #283593 0%, #1A237E 70%, #121858 100%);
        display: flex;
        justify-content: center;
        align-items: center;
    }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }

    /* Tarjeta Blanca */
    div[data-testid="stColumn"]:nth-of-type(2) > div {
        background-color: white;
        border-radius: 24px; 
        padding: 50px 50px 40px 50px; 
        box-shadow: 0 20px 40px rgba(0,0,0,0.5); 
        border: 1px solid #EAEAEA;
        max-width: 420px; 
        margin: auto; 
    }

    /* Textos generales */
    div[data-testid="stColumn"]:nth-of-type(2) label, 
    div[data-testid="stColumn"]:nth-of-type(2) p {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Estilo de los Inputs (Cajas de texto más bonitas) */
    .stTextInput label p {
        font-size: 13.5px !important;
        color: #64748B !important;
        font-weight: 600 !important;
        margin-bottom: 2px !important;
    }
    .stTextInput>div>div>input {
        background-color: #FFFFFF !important;
        border: 1.5px solid #CBD5E1 !important;
        color: #1E293B !important;
        border-radius: 8px !important;
        padding: 14px 16px !important;
        font-size: 15px;
        box-shadow: inset 0 2px 4px rgba(0,0,0,0.02) !important;
        transition: all 0.2s ease;
    }
    .stTextInput>div>div>input:focus {
        border-color: #1A237E !important;
        box-shadow: 0 0 0 3px rgba(26, 35, 126, 0.15) !important;
    }

    /* Botón Principal Amarillo - Centrado y ajustado */
    .stButton {
        display: flex;
        justify-content: center; /* Fuerza el centrado del botón */
        width: 100%;
        margin-top: 10px;
    }
    .stButton>button[type="primary"] {
        background-color: #FFC20E !important; 
        color: #1A237E !important; 
        font-weight: 800 !important;
        font-size: 15px !important;
        border-radius: 25px !important; /* Forma tipo píldora */
        border: none !important;
        width: 85% !important; /* No ocupa todo el ancho, se ve más elegante */
        padding: 12px !important;
        box-shadow: 0 6px 12px rgba(255, 194, 14, 0.2) !important;
        transition: all 0.3s ease !important;
        letter-spacing: 0.5px;
    }
    .stButton>button[type="primary"]:hover {
        background-color: #FFD54F !important; 
        transform: translateY(-2px); 
        box-shadow: 0 8px 15px rgba(255, 194, 14, 0.4) !important;
    }

    /* Enlace Olvidaste tu contraseña - Centrado */
    .boton-olvido {
        display: flex;
        justify-content: center !important;
        margin-top: 5px; 
        margin-bottom: 10px;
    }
    .stButton>button[type="secondary"] {
        background-color: transparent !important;
        color: #64748B !important; 
        border: none !important;
        font-weight: 500 !important;
        font-size: 13.5px !important;
        padding: 0 !important;
        transition: color 0.3s ease;
    }
    .stButton>button[type="secondary"]:hover {
        color: #1A237E !important; 
        text-decoration: underline !important;
    }

    /* Checkbox Recordar - Centrado */
    [data-testid="stCheckbox"] {
        display: flex;
        justify-content: center; /* Centra el checkbox */
        margin-bottom: 15px;
    }
    .stCheckbox label {
        font-size: 13.5px !important;
        color: #64748B !important;
    }

    /* Título de Recuperar centrado */
    .titulo-recuperar {
        text-align: center !important;
        margin-bottom: 20px;
        color: #1A237E !important;
        font-weight: 800;
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
        
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        
        return True, f"✅ Se ha enviado un correo con instrucciones a {destinatario}."
    except smtplib.SMTPAuthenticationError:
        return False, "❌ Autenticación fallida: Microsoft bloqueó el acceso. Verifica la 'Contraseña de Aplicación' en la bóveda."
    except KeyError:
         return False, "❌ La bóveda de seguridad [email] no está configurada correctamente en Streamlit."
    except Exception as e:
        return False, f"❌ Error del sistema de correo: {str(e)}"

# --- 4. PANTALLA DE ACCESO (DISEÑO UX/UI DE TARJETA) ---
if not st.session_state["usuario_autenticado"]:
    
    c1, card_col, c3 = st.columns([1, 3, 1])
    
    with card_col:
        st.write("<br>", unsafe_allow_html=True) 
        
        # === VISTA 1: INICIO DE SESIÓN ===
        if st.session_state["vista_actual"] == "login":
            
            # --- Renderizar Logo y Título centrados horizontalmente ---
            img_b64 = get_base64_image("logo_merkao.png")
            if img_b64:
                st.markdown(f"""
                    <div style='display: flex; justify-content: center; align-items: center; gap: 15px; margin-bottom: 30px;'>
                        <img src='data:image/png;base64,{img_b64}' style='width: 75px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                        <h1 style='color: #1A237E; margin: 0; padding: 0; font-size: 34px; font-weight: 800; letter-spacing: -1px;'>Tesoriapp</h1>
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("<h1 style='text-align: center; color: #1A237E; font-weight: 800; font-size: 36px; margin-bottom: 30px;'>Tesoriapp</h1>", unsafe_allow_html=True)
            
            # --- Formulario ---
            st.text_input("Usuario", placeholder="ej. MAYHELA.SIMON", key="usuario_input")
            st.text_input("Contraseña", type="password", placeholder="••••••••", key="password_input")
            
            st.markdown("<div class='boton-olvido'>", unsafe_allow_html=True)
            if st.button("¿Olvidaste tu contraseña?", type="secondary", key="btn_ir_recuperar"):
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
                except KeyError:
                    st.session_state["error_login"] = "❌ Error crítico: La bóveda secreta no está configurada."
                except Exception as e:
                    st.session_state["error_login"] = f"❌ Error: {e}"

            st.button("INICIAR SESIÓN", on_click=verificar_login, type="primary")
            
        # === VISTA 2: RECUPERAR ACCESO ===
        elif st.session_state["vista_actual"] == "recuperar":
            
            # Renderizar solo el logo centrado
            img_b64 = get_base64_image("logo_merkao.png")
            if img_b64:
                st.markdown(f"<div style='text-align: center; margin-bottom: 20px;'><img src='data:image/png;base64,{img_b64}' style='width: 80px; border-radius: 12px;'></div>", unsafe_allow_html=True)
                
            st.markdown("<h3 class='titulo-recuperar'>¿Olvidaste tu contraseña?</h3>", unsafe_allow_html=True)
            st.markdown("<p style='font-size: 14px; text-align: center; color: #475569 !important; margin-bottom: 20px;'>Escribe tu correo corporativo para recibir las instrucciones de acceso.</p>", unsafe_allow_html=True)
            
            st.text_input("Correo corporativo", placeholder="ejemplo@spsa.pe", key="correo_recuperar_input")
            
            st.markdown("<div class='boton-olvido'>", unsafe_allow_html=True)
            if st.button("Cancelar y Volver", type="secondary", key="btn_volver_login"):
                st.session_state["vista_actual"] = "login"
                if "pwd_recuperada_msg" in st.session_state: del st.session_state["pwd_recuperada_msg"]
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
                
            if "pwd_recuperada_msg" in st.session_state:
                if "❌" in st.session_state["pwd_recuperada_msg"]:
                    st.error(st.session_state["pwd_recuperada_msg"])
                else:
                    st.success(st.session_state["pwd_recuperada_msg"])

            def intentar_recuperacion():
                correo = st.session_state.correo_recuperar_input.strip()
                if not correo or "@" not in correo:
                    st.session_state["pwd_recuperada_msg"] = "⚠️ Ingresa un correo corporativo válido."
                else:
                    exito, mensaje = enviar_correo_recuperacion(correo)
                    st.session_state["pwd_recuperada_msg"] = mensaje
            
            st.button("ENVIAR INSTRUCCIONES", on_click=intentar_recuperacion, type="primary")

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

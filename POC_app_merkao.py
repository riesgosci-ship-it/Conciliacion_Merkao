import streamlit as st
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# --- 1. CONFIGURACIÓN BÁSICA ---
st.set_page_config(page_title="Tesoriapp - Acceso", page_icon="🏦", layout="centered")

# --- 2. DISEÑO UI AVANZADO (CSS UX/UI PROFESIONAL) ---
# Usamos CSS para forzar el diseño de 'tarjeta' blanca y el degradado de fondo.
st.markdown("""
<style>
    /* Ocultar elementos nativos de Streamlit para Marca Blanca */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden;}

    /* 1. FONDO DE LA APP CON DEGRADADO (Radial Gradient) */
    [data-testid="stAppViewContainer"] {
        background: radial-gradient(circle, #283593 0%, #1A237E 70%, #121858 100%);
        display: flex;
        justify-content: center;
        align-items: center;
    }
    
    /* Asegurar que el contenedor principal sea transparente para ver el fondo */
    [data-testid="stHeader"] {
        background-color: rgba(0,0,0,0);
    }

    /* 2. DISEÑO DEL CUADRO BLANCO (Login Card) */
    /* Targeteamos el contenedor de la columna central para volverlo una tarjeta */
    div.stVerticalBlock > div[data-testid="stColumn"] > div > div.stVerticalBlock {
        background-color: white;
        border-radius: 15px;
        padding: 50px 40px; /* Espaciado interno generoso (aire) */
        box-shadow: 0 10px 25px rgba(0,0,0,0.3); /* Sombra suave para elevación */
        border: 1px solid #E0E0E0;
    }

    /* 3. ESTILOS DE TEXTO E INPUTS DENTRO DE LA TARJETA */
    /* Forzar textos oscuros dentro del cuadro blanco para contraste */
    [data-testid="stColumn"] label, 
    [data-testid="stColumn"] p, 
    [data-testid="stColumn"] h1, 
    [data-testid="stColumn"] h2 {
        color: #1E293B !important; /* Gris oscuro profesional */
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Estilo de los campos de texto input */
    .stTextInput>div>div>input {
        background-color: #F8FAFC !important;
        border: 1px solid #CBD5E1 !important;
        color: #1E293B !important;
        border-radius: 8px !important;
        padding: 10px 15px !important;
    }

    /* 4. BOTÓN AMARILLO CORPORATIVO (Tesoriapp Style) */
    .stButton>button {
        background-color: #FFC20E !important; /* Amarillo Merkao */
        color: #1A237E !important; /* Texto Azul oscuro */
        font-weight: bold !important;
        font-size: 16px !important;
        border-radius: 8px !important;
        border: none !important;
        width: 100% !important; /* Ancho completo del cuadro */
        padding: 12px !important;
        margin-top: 25px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        background-color: #FFD54F !important; /* Amarillo más claro en hover */
        transform: translateY(-2px); /* Pequeño salto hacia arriba */
        box-shadow: 0 6px 12px rgba(0,0,0,0.15) !important;
    }

    /* 5. ESTILOS DE ENLACES (Olvidaste contraseña) */
    .stButton>button.st-emotion-cache-1r5wkyy { /* Botón tipo link secundario */
        background-color: transparent !important;
        color: #1A237E !important;
        border: none !important;
        font-weight: normal !important;
        font-size: 14px !important;
        margin-top: 5px !important;
        padding: 0 !important;
        text-align: right !important;
        text-decoration: underline !important;
    }

    /* Ajuste para el checkbox */
    .stCheckbox label {
        font-size: 14px !important;
        color: #475569 !important;
    }

    /* Ajustar espacio del logo */
    [data-testid="stImage"] {
        display: flex;
        justify-content: center;
        margin-bottom: 10px;
    }
    
    /* Título Tesoriapp elegante si no hay logo */
    .titulo-elegante {
        text-align: center;
        color: #1A237E !important;
        font-weight: 800;
        font-size: 32px;
        margin-bottom: 20px;
    }

</style>
""", unsafe_allow_html=True)

# --- 3. LÓGICA DE MEMORIA DE SESIÓN (LOGIN PROFESIONAL) ---
# Usamos st.secrets para no exponer contraseñas en GitHub corporativo.
if "usuario_autenticado" not in st.session_state:
    st.session_state["usuario_autenticado"] = False

def verificar_login():
    usu = st.session_state.usuario_input.strip().upper() # Forzar mayúsculas
    pwd = st.session_state.password_input
    
    try:
        # Validar contra la Bóveda Secreta de Streamlit
        if st.secrets["passwords"].get(usu) == pwd:
            st.session_state["usuario_autenticado"] = True
            # Limpiar mensajes de error si existen
            if "error_login" in st.session_state: del st.session_state["error_login"]
        else:
            st.session_state["error_login"] = "❌ Usuario o contraseña incorrectos."
    except Exception:
        st.session_state["error_login"] = "❌ Error de configuración en la bóveda de seguridad."

# --- 4. PANTALLA DE LOGIN RE-DISEÑADA (TIPO CARD / BINNACLE) ---
if not st.session_state["usuario_autenticado"]:
    
    # Columnas para centrar la tarjeta blanca (diseño centered)
    # Ajustamos pesos [0.5, 3, 0.5] para que la tarjeta no sea gigante en pantallas grandes
    c1, card_col, c3 = st.columns([0.5, 3, 0.5])
    
    with card_col:
        st.write("<br><br>", unsafe_allow_html=True) # Espacio superior para centrar verticalmente
        
        # 4.1 LOGO O TÍTULO TESORIAPP
        # Intenta cargar logo corporativo. Debe estar en la raíz de GitHub.
        try:
            st.image("logo_merkao.png", width=120)
        except:
            # Fallback si no hay imagen
            st.markdown("<h1 class='titulo-elegante'>Tesoriapp</h1>", unsafe_allow_html=True)
        
        # 4.2 FORMULARIO DE INGRESO
        st.text_input("Usuario", placeholder="NOMBRE.APELLIDO", key="usuario_input")
        st.text_input("Contraseña", type="password", placeholder="••••••••", key="password_input")
        
        # Mostrar error de login si existe dentro de la tarjeta
        if "error_login" in st.session_state:
            st.error(st.session_state["error_login"])

        # 4.3 OLVIDASTE CONTRASEÑA (INTERACTIVO)
        # Usamos un expander invisible o un botón link para UX.
        # Aquí usaremos un expander estilizado para mostrar instrucciones sin salir de la app.
        with st.expander("¿Olvidaste tu contraseña?", expanded=False):
            st.markdown("""
                <p style='color: #475569 !important; font-size: 13px; text-align: center;'>
                Por favor, comunícate con el área de Tesorería o TI para restablecer tus credenciales corporativas de Tesoriapp.
                </p>
            """, unsafe_allow_html=True)
            
        # 4.4 RECORDAR (CHECKBOX)
        st.checkbox("Recordar en este dispositivo", value=True)
        
        # 4.5 BOTÓN DE ACCIÓN PRINCIPAL (AMARILLO)
        st.button("Iniciar Sesión", on_click=verificar_login)
        
        st.write("<br>", unsafe_allow_html=True) # Espacio inferior del cuadro

# --- 5. PANTALLA PRINCIPAL (MÓDULO DE TRABAJO - OCULTO) ---
# Esta sección permanece funcional, se puede estilizar luego.
else:
    # Restablecer el estilo para la zona de trabajo (fondo blanco estándar)
    st.markdown("""
        <style>
            [data-testid="stAppViewContainer"] { background: white; }
            h1, h2, h3, p, label { color: #1E293B !important; }
        </style>
    """, unsafe_allow_html=True)

    # Header de la aplicación
    c_logo, c_titulo, c_logout = st.columns([1, 4, 1])
    with c_logo:
        try: st.image("logo_merkao.png", width=60)
        except: st.write("🏦")
    with c_titulo:
        st.markdown("<h2>Módulo de Conciliación - Tesoriapp</h2>", unsafe_allow_html=True)
    with c_logout:
        # Botón secundario para cerrar sesión
        if st.button("Salir"):
            st.session_state["usuario_autenticado"] = False
            st.rerun()
        
    st.markdown("---")
    
    # INTEGRAMOS EL SCRIPT BASE DE CONCILIACIÓN
    st.write("### 📁 Carga de Archivos Maestros para Procesamiento")
    st.info("💡 Por favor, sube los archivos correspondientes al mes que deseas conciliar.")
    
    # Reemplazamos los diálogos de Tkinter por st.file_uploader de Streamlit (Entorno Web)
    col_izq, col_der = st.columns(2)
    
    with col_izq:
        st.markdown("#### Fuentes de Ventas")
        file_jsat = st.file_uploader("Ventas (JSatellite)", type=['xlsx', 'xls', 'csv'], help="Sube los archivos consolidado de ventas.")
        file_mon = st.file_uploader("Pedidos (Monitor)", type=['xlsx', 'xls', 'csv'])
        file_liq = st.file_uploader("Entregas (Liquidador)", type=['xlsx', 'xls', 'csv'])

    with col_der:
        st.markdown("#### Fuentes Bancarias")
        file_izi = st.file_uploader("IziPay", type=['xlsx', 'xls', 'csv'])
        file_bbr = st.file_uploader("BBR", type=['xlsx', 'xls', 'csv'])
    
    st.markdown("---")
    # Botón principal de acción
    st.button("▶️ INICIAR PROCESAMIENTO", type="primary")

    st.write("*(El motor contable de Pandas procesará los archivos cargados arriba)*")

import streamlit as st

# --- 1. CONFIGURACIÓN BÁSICA ---
st.set_page_config(page_title="Tesoriapp", page_icon="🏦", layout="centered")

# --- 2. DISEÑO UI (CSS CORPORATIVO MARCA BLANCA) ---
st.markdown("""
<style>
    /* Ocultar barra superior de Streamlit, menú de GitHub y footer */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden;}

    /* Fondo azul corporativo */
    [data-testid="stAppViewContainer"] {
        background-color: #1A237E; /* Tono azul oscuro elegante */
    }

    /* Color de texto blanco para los títulos y labels */
    h1, h2, h3, p, label, .stMarkdown {
        color: white !important;
    }

    /* Botón amarillo corporativo ocupando todo el ancho */
    .stButton>button {
        background-color: #FFC20E !important; /* Amarillo Merkao */
        color: #1A237E !important; /* Texto azul */
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        width: 100% !important;
        padding: 10px !important;
        margin-top: 10px !important;
    }
    .stButton>button:hover {
        background-color: #FFD54F !important;
        transform: scale(1.02);
        transition: 0.2s;
    }
    
    /* Centrar textos de ayuda (Olvidaste contraseña) */
    .centrar-texto {
        text-align: center;
        margin-top: 15px;
        font-size: 14px;
        color: #90CAF9 !important;
        cursor: pointer;
    }
</style>
""", unsafe_allow_html=True)

# --- 3. LÓGICA DE MEMORIA DE SESIÓN (LOGIN) ---
if "usuario_autenticado" not in st.session_state:
    st.session_state["usuario_autenticado"] = False

def verificar_login():
    usu = st.session_state.usuario_input.strip()
    pwd = st.session_state.password_input
    
    # Validar contra la Bóveda Secreta
    try:
        if st.secrets["passwords"].get(usu) == pwd:
            st.session_state["usuario_autenticado"] = True
        else:
            st.error("❌ Contraseña o usuario incorrectos.")
    except Exception:
        st.error("❌ Error conectando con la bóveda de seguridad.")

# --- 4. PANTALLA DE LOGIN (TIPO BINNACLE) ---
if not st.session_state["usuario_autenticado"]:
    
    # Usar columnas para centrar y simular una "tarjeta" más angosta
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.write("<br><br>", unsafe_allow_html=True) # Espacio superior
        
        # Intentar cargar el logo (debes subir un archivo 'logo_merkao.png' a GitHub)
        try:
            st.image("logo_merkao.png", use_container_width=True)
        except:
            st.markdown("<h1 style='text-align: center;'>Tesoriapp</h1>", unsafe_allow_html=True)
        
        st.write("<br>", unsafe_allow_html=True)
        
        # Formulario
        st.text_input("Usuario", placeholder="Ej. NOMBRE.APELLIDO", key="usuario_input")
        st.text_input("Contraseña", type="password", placeholder="Ingresa tu contraseña", key="password_input")
        
        st.button("Iniciar Sesión", on_click=verificar_login)
        
        # Enlaces inferiores solicitados
        st.markdown("<p class='centrar-texto'>¿Olvidaste tu contraseña?</p>", unsafe_allow_html=True)
        st.checkbox("Recuérdame en este dispositivo")

# --- 5. PANTALLA PRINCIPAL (MOTOR CONTABLE - OCULTO HASTA EL LOGIN) ---
else:
    # Header de la aplicación ya logueada
    c_logo, c_titulo, c_logout = st.columns([1, 3, 1])
    with c_titulo:
        st.markdown("<h2>Módulo de Conciliación</h2>", unsafe_allow_html=True)
    with c_logout:
        st.button("Cerrar Sesión", on_click=lambda: st.session_state.update(usuario_autenticado=False))
        
    st.markdown("---")
    
    # Aquí es donde vivirá la lógica de Pandas. 
    # Ya no usaremos Tkinter, usaremos zonas de "Arrastrar y Soltar"
    st.write("### 📁 Carga de Archivos Maestros")
    
    col_izq, col_der = st.columns(2)
    with col_izq:
        file_jsat = st.file_uploader("📥 Subir Ventas (JSatellite)", type=['xlsx', 'csv'])
        file_mon = st.file_uploader("📥 Subir Pedidos (Monitor)", type=['xlsx', 'csv'])
    with col_der:
        file_liq = st.file_uploader("📥 Subir Entregas (Liquidador)", type=['xlsx', 'csv'])
        file_izi = st.file_uploader("📥 Subir IziPay", type=['xlsx', 'csv'])
    
    st.markdown("---")
    st.write("*(El motor de cruce de Pandas se conectará aquí en el siguiente paso)*")

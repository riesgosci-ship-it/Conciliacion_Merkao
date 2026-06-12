import streamlit as st
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# --- CONFIGURACIÓN BÁSICA ---
st.set_page_config(page_title="Sistemas Merkao", page_icon="☁️", layout="centered")

# --- OCULTAR ELEMENTOS DE STREAMLIT (UI LIMPIA MARCA BLANCA) ---
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            header {visibility: hidden;}
            footer {visibility: hidden;}
            [data-testid="stToolbar"] {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# --- LÓGICA DE MEMORIA DE SESIÓN (LOGIN SEGURO) ---
if "usuario_autenticado" not in st.session_state:
    st.session_state["usuario_autenticado"] = False

def verificar_login():
    usu = st.session_state.usuario_input.strip()
    pwd = st.session_state.password_input
    
    # Validar contra la Bóveda Secreta de Streamlit (No hay contraseñas en el código)
    try:
        if st.secrets["passwords"].get(usu) == pwd:
            st.session_state["usuario_autenticado"] = True
        else:
            st.error("❌ Contraseña incorrecta.")
    except Exception:
        st.error("❌ Error de acceso o usuario no registrado.")

# --- PANTALLA DE LOGIN ---
if not st.session_state["usuario_autenticado"]:
    st.title("🔐 Acceso al Sistema Merkao")
    st.write("Por favor, ingresa tus credenciales corporativas.")
    
    st.text_input("Usuario (Ej. NOMBRE.APELLIDO)", key="usuario_input")
    st.text_input("Contraseña", type="password", key="password_input")
    
    st.button("Ingresar al Sistema", on_click=verificar_login, type="primary")

# --- PANTALLA PRINCIPAL DE LA APLICACIÓN (OCULTA) ---
else:
    # Botón para cerrar sesión
    col1, col2 = st.columns([0.8, 0.2])
    with col1:
        st.title("☁️ Test de Conexión SharePoint")
    with col2:
        st.button("Cerrar Sesión", on_click=lambda: st.session_state.update(usuario_autenticado=False))
        
    st.write("Validador de credenciales de aplicación para migración web.")
    st.markdown("---")

    # --- FORMULARIO DE CREDENCIALES SHAREPOINT ---
    url_sitio = st.text_input(
        "URL del sitio de SharePoint", 
        value="https://intercorpretail.sharepoint.com/sites/ConciliacinTesoreraMerkao"
    )

    client_id = st.text_input("Client ID", type="password")
    client_secret = st.text_input("Client Secret", type="password")

    st.markdown("---")

    if st.button("🚀 Probar Conexión", type="primary"):
        if not url_sitio or not client_id or not client_secret:
            st.warning("⚠️ Por favor, completa tu Client ID y Client Secret primero.")
        else:
            try:
                with st.spinner("Autenticando de forma segura con Microsoft Azure..."):
                    credenciales = ClientCredential(client_id, client_secret)
                    ctx = ClientContext(url_sitio).with_credentials(credenciales)
                    
                    web = ctx.web
                    ctx.load(web)
                    ctx.execute_query()
                    
                    st.success("✅ ¡Conexión Exitosa con Intercorp Retail!")
                    st.info(f"El script logró entrar como administrador al sitio: **{web.properties['Title']}**")
                    
                    listas = ctx.web.lists
                    ctx.load(listas)
                    ctx.execute_query()
                    
                    st.write("### 📁 Listas y Carpetas detectadas en el servidor:")
                    encontrado = False
                    for lista in listas:
                        titulo_lista = lista.properties['Title']
                        if "Satelite" in titulo_lista or "Satelite" in titulo_lista.lower() or titulo_lista == "Maestro_JSatelite":
                            st.write(f"- 🟢 **{titulo_lista}** (¡Lista detectada correctamente!)")
                            encontrado = True
                        else:
                            st.write(f"- ⚪ {titulo_lista}")
                    
                    if not encontrado:
                        st.warning("Pude entrar al sitio, pero no logré visualizar la lista 'Maestro_JSatelite'. Revisa si la lista tiene permisos heredados del sitio.")
                        
            except Exception as e:
                st.error("❌ Falló la conexión. Verifica que el Client ID y el Secret estén vigentes y tengan permisos para este sitio.")
                st.code(str(e))

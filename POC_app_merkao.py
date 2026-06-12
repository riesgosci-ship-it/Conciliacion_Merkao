import streamlit as st
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# Configuración básica de la página web
st.set_page_config(page_title="Prueba SharePoint Merkao", page_icon="☁️", layout="centered")

st.title("☁️ Test de Conexión a SharePoint - Merkao")
st.write("Validador de credenciales de aplicación (Client ID / Secret) para migración web.")
st.markdown("---")

# --- FORMULARIO DE CREDENCIALES ---
# La URL ya está precargada con la raíz exacta de tu sitio corporativo
url_sitio = st.text_input(
    "URL del sitio de SharePoint", 
    value="https://intercorpretail.sharepoint.com/sites/ConciliacinTesoreraMerkao"
)

client_id = st.text_input("Client ID", type="password", help="eab2f548-a996-4183-9cc3-40400c3a7d91")
client_secret = st.text_input("Client Secret", type="password", help="bXN3OFF+YllKd0VzeXNGOGJDd3VNdTNBQ3ZjdnVaY082NEkyZ2RBLg==")

st.markdown("---")

if st.button("🚀 Probar Conexión", type="primary"):
    if not url_sitio or not client_id or not client_secret:
        st.warning("⚠️ Por favor, completa tu Client ID y Client Secret primero.")
    else:
        try:
            with st.spinner("Autenticando de forma segura con Microsoft Azure..."):
                # 1. Inicializar credenciales de aplicación (El "Pase VIP")
                credenciales = ClientCredential(client_id, client_secret)
                ctx = ClientContext(url_sitio).with_credentials(credenciales)
                
                # 2. Intentar leer las propiedades básicas de la web para validar acceso
                web = ctx.web
                ctx.load(web)
                ctx.execute_query()
                
                st.success("✅ ¡Conexión Exitosa con Intercorp Retail!")
                st.info(f"El script logró entrar como administrador al sitio: **{web.properties['Title']}**")
                
                # 3. Buscar la lista de Maestro_JSatelite que creaste
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
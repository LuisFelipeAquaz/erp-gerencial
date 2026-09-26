import streamlit as st


def verificar_autenticacion():
    """
    Sistema de seguridad.
    Muestra el formulario de acceso si el usuario no está autenticado
    y detiene la ejecución del resto de la app (st.stop()).
    """
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        st.title("🔒 Acceso al ERP Gerencial")
        usuario = st.text_input("Usuario")
        clave = st.text_input("Contraseña", type="password")

        if st.button("Entrar"):
            if usuario in st.secrets["passwords"] and st.secrets["passwords"][usuario] == clave:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("😕 Usuario o contraseña incorrectos")

        st.stop()

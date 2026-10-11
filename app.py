import streamlit as st
import sqlite3
import datetime as dt
from openpyxl import Workbook
import io

# Configuración básica de la página
st.set_page_config(page_title="Registro de Visitas", layout="centered")

# Inicializar Base de Datos en memoria para la demostración
conn = sqlite3.connect(":memory:", check_same_thread=False)
cursor = conn.cursor()
cursor.execute("""
    CREATE TABLE IF NOT EXISTS visitas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT,
        distrito TEXT,
        zona TEXT,
        centro TEXT,
        fecha TEXT
    )
""")
conn.commit()

# Datos de prueba para la estructura del centro sanitario
ESTRUCTURA = {
    "Distrito Sanitario Sevilla": {
        "Zona Centro": ["C.S. San Luis", "C.S. Alameda"],
        "Zona Norte": ["C.S. Pino Montano", "C.S. Macarena"]
    },
    "Hospital Virgen del Rocío": {
        "Área General": ["Urgencias", "Consultas Externas"],
        "Área Infantil": ["Planta 2 Pediatría", "UCCU Infantil"]
    }
}

# Control de Sesión simplificado
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario = ""

st.title("🏥 Registro de Visitas Sanitarias")

# --- PANTALLA DE LOGIN ---
if not st.session_state.autenticado:
    st.subheader("Acceso al Sistema")
    u_input = st.text_input("Usuario (Prueba: Administrador)")
    p_input = st.text_input("Contraseña", type="password")
    
    if st.button("Iniciar Sesión"):
        if u_input == "Administrador" and p_input == "Finidi.2026":
            st.session_state.autenticado = True
            st.session_state.usuario = "Administrador"
            st.rerun()
        elif u_input and p_input == "Prueba.2026":
            st.session_state.autenticado = True
            st.session_state.usuario = u_input
            st.rerun()
        else:
            st.error("Credenciales incorrectas para la demostración.")

# --- INTERFAZ PRINCIPAL ---
else:
    st.sidebar.write(f"Conectado como: **{st.session_state.usuario}**")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.usuario = ""
        st.rerun()
        
    menu = st.sidebar.radio("Menú de Opciones", ["Nueva Visita", "Historial de Visitas", "Exportar Reportes Excel"])
    
    # OPCIÓN 1: NUEVA VISITA
    if menu == "Nueva Visita":
        st.subheader("Registrar Nueva Visita Sanitaria")
        
        distrito_sel = st.selectbox("Seleccione Distrito / Hospital", list(ESTRUCTURA.keys()))
        zona_sel = st.selectbox("Seleccione Zona", list(ESTRUCTURA[distrito_sel].keys()))
        centro_sel = st.selectbox("Seleccione Centro / Unidad", ESTRUCTURA[distrito_sel][zona_sel])
        fecha_sel = st.date_input("Fecha de la Visita", dt.date.today())
        
        if st.button("Guardar Registro"):
            cursor.execute("INSERT INTO visitas (usuario, distrito, zona, centro, fecha) VALUES (?, ?, ?, ?, ?)",
                           (st.session_state.usuario, distrito_sel, zona_sel, centro_sel, str(fecha_sel)))
            conn.commit()
            st.success(f"¡Registro guardado con éxito para {centro_sel}!")

    # OPCIÓN 2: HISTORIAL
    elif menu == "Historial de Visitas":
        st.subheader("Historial Local de Visitas")
        cursor.execute("SELECT usuario, distrito, zona, centro, fecha FROM visitas ORDER BY id DESC")
        datos = cursor.fetchall()
        
        if datos:
            for d in datos:
                st.info(f"📅 **{d[4]}** — {d[0]} visitó **{d[3]}** ({d[2]} - {d[1]})")
        else:
            st.warning("No hay visitas registradas en esta demostración todavía.")
            
    # OPCIÓN 3: EXPORTACIONES EXCEL
    elif menu == "Exportar Reportes Excel":
        st.subheader("Generación de Reportes del Distrito")
        st.write("Genera dinámicamente hojas de cálculo automatizadas usando la librería `openpyxl`.")
        
        cursor.execute("SELECT usuario, distrito, zona, centro, fecha FROM visitas")
        todas = cursor.fetchall()
        
        if todas:
            # Crear archivo Excel en memoria mediante openpyxl
            wb = Workbook()
            ws = wb.active
            ws.title = "VISITAS"
            ws.append(["Trabajador", "Distrito/Hospital", "Zona", "Centro Sanitario", "Fecha"])
            
            for fila in todas:
                ws.append(fila)
                
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label="📥 Descargar REGISTRO_VISITAS_DISTRITO_2026.xlsx",
                data=buffer,
                file_name="REGISTRO_VISITAS_DISTRITO_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Registra al menos una visita para poder activar la descarga del Excel.")

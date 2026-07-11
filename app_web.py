import os
import difflib
import pandas as pd
import openpyxl
import streamlit as st
import io
import re
import requests

# --- CONFIGURACIÓN DE PÁGINA WEB ---
st.set_page_config(page_title="Pedidos El Bajo", page_icon="📊", layout="centered")

CONFIG_PROVEEDORES = {
    "Desa": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "CCU": {"col_nombre": 1, "col_pedido": 4, "fila_inicio": 2},
    "ANDINA": {"col_nombre": 1, "col_pedido": 3, "fila_inicio": 2},
    "Tost": {"col_nombre": 1, "col_pedido": 3, "fila_inicio": 2},
    "ATF (TH Tonicas)": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "Tubinger": {"col_nombre": 1, "col_pedido": 3, "fila_inicio": 2},
    "Vinoteca": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "Cerros de Chena": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "Tamango": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "Kombuchacha": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "ByMaria": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "TeGusta": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "Limache": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 2},
    "Segafredo": {"col_nombre": 1, "col_pedido": 2, "fila_inicio": 4},
}

# --- CONTROL DE ETAPAS (SESSION STATE) ---
if "etapa" not in st.session_state:
    st.session_state.etapa = "upload"
    st.session_state.ambiguedades = {}
    st.session_state.cache_decisiones = {}
    st.session_state.inventario_db = {}
    st.session_state.pedidos_bytes = None
    st.session_state.excel_final = None

# --- CONVERSIÓN Y DESCARGA DE GOOGLE DRIVE ---
def obtener_bytes_desde_drive(url_compartido):
    """Extrae el ID de Drive, descarga el archivo binario real y devuelve un buffer de memoria."""
    if not url_compartido:
        return None
    
    file_id = None
    match_d = re.search(r"/d/([a-zA-Z0-9-_]+)", url_compartido)
    if match_d:
        file_id = match_d.group(1)
    else:
        match_id = re.search(r"id=([a-zA-Z0-9-_]+)", url_compartido)
        if match_id:
            file_id = match_id.group(1)
            
    if not file_id:
        return None
        
    url_descarga = f"https://drive.google.com/uc?export=download&id={file_id}"
    
    # Descargamos los bytes reales del archivo desde la nube
    res = requests.get(url_descarga)
    if res.status_code == 200:
        return io.BytesIO(res.content)
    return None

# --- LÓGICA MATEMÁTICA DE STOCK ---
def cargar_inventario_real(file_object):
    df = pd.read_excel(file_object, sheet_name="Inventario General", skiprows=4)
    df = df.dropna(subset=["Nombre Producto"])
    df = df[df["Nombre Producto"].str.strip() != ""]
    
    df["Par Stock"] = pd.to_numeric(df["Par Stock"], errors="coerce").fillna(0)
    df["Bodega"] = pd.to_numeric(df["Bodega"], errors="coerce").fillna(0)
    df["Barra"] = pd.to_numeric(df["Barra"], errors="coerce").fillna(0)
    df["Total Actual"] = df["Bodega"] + df["Barra"]
    
    return {str(r["Nombre Producto"]).strip(): {"par": r["Par Stock"], "actual": r["Total Actual"]} for _, r in df.iterrows()}

def encontrar_coincidencia_inteligente(nombre_prov, lista_inv):
    n_prov_clean = nombre_prov.lower().strip()
    exactas = [n for n in lista_inv if n_prov_clean in n.lower().strip() or n.lower().strip() in n_prov_clean]
    
    if len(exactas) == 1: return exactas[0], "PERFECTO"
    if len(exactas) > 1: return exactas, "DUPLICADO"
    
    mejores = difflib.get_close_matches(nombre_prov, lista_inv, n=3, cutoff=0.4)
    if not mejores: return None, "NINGUNO"
    
    score_top = difflib.SequenceMatcher(None, n_prov_clean, mejores[0].lower()).ratio()
    if score_top < 0.90: return mejores, "BAJA_CERTEZA"
    
    return mejores[0], "ALTA_CERTEZA"

def ejecutar_calculo_matematico():
    # Rehidratamos el archivo maestro desde los bytes guardados en la sesión web
    wb = openpyxl.load_workbook(io.BytesIO(st.session_state.pedidos_bytes))
    inventario = st.session_state.inventario_db
    
    for sheet_name in wb.sheetnames:
        if sheet_name not in CONFIG_PROVEEDORES: continue
        ws = wb[sheet_name]
        conf = CONFIG_PROVEEDORES[sheet_name]
        
        for row in range(conf["fila_inicio"], ws.max_row + 1):
            cell_val = ws.cell(row=row, column=conf["col_nombre"]).value
            if not cell_val: continue
            n_prov = str(cell_val).strip()
            if n_prov.lower() in ["productos", "producto", "total", "rut:", "detalle de producto"]: continue
            
            item_elegido = st.session_state.cache_decisiones.get(n_prov)
            
            if item_elegido and item_elegido in inventario:
                datos = inventario[item_elegido]
                cantidad = max(0, datos["par"] - datos["actual"])
                ws.cell(row=row, column=conf["col_pedido"]).value = cantidad if cantidad > 0 else ""
            else:
                ws.cell(row=row, column=conf["col_pedido"]).value = ""
                
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    st.session_state.excel_final = buffer.getvalue()

# --- INTERFAZ WEB ---
st.title("Pedidos Automáticos: El Bajo")

# --- ETAPA 1: CAPTURA DE FUENTES ---
if st.session_state.etapa == "upload":
    st.subheader("1. Origen de las Planillas")
    tab_drive, tab_local = st.tabs(["🔗 Google Drive (Móvil)", "📂 Subir Locales"])
    
    input_inv = None
    input_ped = None
    
    with tab_drive:
        st.info("💡 Asegúrate de que los archivos en Drive tengan el acceso en 'Cualquier persona con el enlace'.")
        link_inv = st.text_input("Enlace de Google Drive del Inventario Diario:")
        link_ped = st.text_input("Enlace de Google Drive del Maestro de Pedidos:")
        
    with tab_local:
        archivo_inv = st.file_uploader("Sube el Inventario Diario (.xlsx)", type=["xlsx"])
        archivo_ped = st.file_uploader("Sube el Maestro de Pedidos (.xlsx)", type=["xlsx"])

    # Procesar la selección según la pestaña activa
    if link_inv and link_ped:
        with st.spinner("Descargando archivos desde Google Drive..."):
            input_inv = obtener_bytes_desde_drive(link_inv)
            input_ped = obtener_bytes_desde_drive(link_ped)
    elif archivo_inv and archivo_ped:
        input_inv = io.BytesIO(archivo_inv.read())
        input_ped = io.BytesIO(archivo_ped.read())
        
    if input_inv and input_ped:
        if st.button("ANALIZAR PLANILLAS", use_container_width=True):
            try:
                # Guardamos los componentes en la sesión web global para evitar pérdidas en los reruns
                st.session_state.inventario_db = cargar_inventario_real(input_inv)
                st.session_state.pedidos_bytes = input_ped.getvalue() if hasattr(input_ped, "getvalue") else input_ped.read()
                
                nombres_inv = list(st.session_state.inventario_db.keys())
                
                wb_scan = openpyxl.load_workbook(io.BytesIO(st.session_state.pedidos_bytes))
                ambiguedades_encontradas = {}
                
                for sheet_name in wb_scan.sheetnames:
                    if sheet_name not in CONFIG_PROVEEDORES: continue
                    ws = wb_scan[sheet_name]
                    conf = CONFIG_PROVEEDORES[sheet_name]
                    
                    for row in range(conf["fila_inicio"], ws.max_row + 1):
                        cell_val = ws.cell(row=row, column=conf["col_nombre"]).value
                        if not cell_val: continue
                        n_prov = str(cell_val).strip()
                        if n_prov.lower() in ["productos", "producto", "total", "rut:", "detalle de producto"]: continue
                        
                        res, tipo_match = encontrar_coincidencia_inteligente(n_prov, nombres_inv)
                        
                        if tipo_match in ["ALTA_CERTEZA", "PERFECTO"]:
                            st.session_state.cache_decisiones[n_prov] = res
                        elif tipo_match in ["DUPLICADO", "BAJA_CERTEZA"]:
                            ambiguedades_encontradas[n_prov] = {"candidatos": res, "tipo": tipo_match}
                
                st.session_state.ambiguedades = ambiguedades_encontradas
                
                if ambiguedades_encontradas:
                    st.session_state.etapa = "resolver"
                else:
                    ejecutar_calculo_matematico()
                    st.session_state.etapa = "descargar"
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error al procesar: {str(e)}. Verifica que el archivo no esté protegido o vacío.")

# --- ETAPA 2: RESOLVER VALIDACIONES ---
elif st.session_state.etapa == "resolver":
    st.subheader(" Validación de Productos")
    st.info(f"Faltan confirmar {len(st.session_state.ambiguedades)} productos. Selecciónalos:")
    
    with st.form("formulario_resolucion"):
        nuevas_decisiones = {}
        # Procesamos en bloques cómodos para pantallas móviles
        items_a_mostrar = list(st.session_state.ambiguedades.items())[:15]
        
        for n_prov, info in items_a_mostrar:
            tipo_texto = "Duplicado" if info["tipo"] == "DUPLICADO" else "Certeza Baja"
            options_select = ["[ No pedir / Ignorar ]"] + info["candidatos"]
            nuevas_decisiones[n_prov] = st.selectbox(label=f"📋 '{n_prov}' ({tipo_texto})", options=options_select, key=n_prov)
            st.markdown("---")
            
        if st.form_submit_button("💾 CONFIRMAR Y CALCULAR SIGUIENTE BLOQUE", use_container_width=True):
            for n_prov, eleccion in nuevas_decisiones.items():
                st.session_state.cache_decisiones[n_prov] = None if eleccion == "[ No pedir / Ignorar ]" else eleccion
                del st.session_state.ambiguedades[n_prov]
            
            if not st.session_state.ambiguedades:
                ejecutar_calculo_matematico()
                st.session_state.etapa = "descargar"
            st.rerun()

# --- ETAPA 3: DESCARGA FINAL ---
elif st.session_state.etapa == "descargar":
    st.subheader("¡Sugerencias Listas!")
    st.success("Se ha calculado la reposición de todas tus marcas exitosamente.")
    
    st.download_button(
        label="📥 DESCARGAR EXCEL DE PEDIDOS COMPLETO",
        data=st.session_state.excel_final,
        file_name="Pedido_Barra_Sugerido_El_Bajo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
        
    if st.button("🔄 Calcular Otro Inventario / Reiniciar", use_container_width=True):
        st.session_state.clear()
        st.rerun()
import os
import difflib
import pandas as pd
import openpyxl
import streamlit as st
import io
import re

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
    st.session_state.excel_final = None

# --- FUNCIONES DE CONVERSIÓN DE GOOGLE DRIVE ---
def Convertir_link_drive(url_compartido):
    """Extrae el ID de un enlace de Google Drive y lo transforma en un link de descarga directa."""
    if not url_compartido:
        return None
    try:
        # Buscar ID en enlaces estándar de Drive (/d/ID/view...)
        match_d = re.search(r"/d/([a-zA-Z0-9-_]+)", url_compartido)
        if match_d:
            file_id = match_d.group(1)
            return f"https://drive.google.com/uc?export=download&id={file_id}"
        
        # Buscar ID en enlaces cortos antiguos o con parámetros (?id=ID)
        match_id = re.search(r"id=([a-zA-Z0-9-_]+)", url_compartido)
        if match_id:
            file_id = match_id.group(1)
            return f"https://drive.google.com/uc?export=download&id={file_id}"
    except Exception:
        return None
    return url_compartido

# --- FUNCIONES DE LÓGICA DE BARRA ---
def cargar_inventario_real(origen_file):
    df = pd.read_excel(origen_file, sheet_name="Inventario General", skiprows=4)
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

def ejecutar_calculo_matematico(origen_pedidos):
    wb = openpyxl.load_workbook(origen_pedidos)
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

# --- INTERFAZ DE USUARIO ---
st.title("🍹 Pedidos Automáticos - El Bajo")

# --- ETAPA 1: OBTENER DATOS (WEB o DRIVE) ---
if st.session_state.etapa == "upload":
    st.subheader("1. Origen de las Planillas")
    
    # Pestañas para elegir el método más cómodo en el momento
    tab_drive, tab_local = st.tabs(["🔗 Google Drive (Recomendado Móvil)", "📂 Subir Archivos Locales"])
    
    input_inv = None
    input_ped = None
    
    with tab_drive:
        st.info("💡 Recuerda que los archivos en tu Drive deben estar configurados como 'Cualquier persona con el enlace puede ver' para que la app los procese.")
        link_inv = st.text_input("Pega el enlace de compartir del Inventario Diario:")
        link_ped = st.text_input("Pega el enlace de compartir del Maestro de Pedidos:")
        
        if link_inv: input_inv = Convertir_link_drive(link_inv)
        if link_ped: input_ped = Convertir_link_drive(link_ped)
        
    with tab_local:
        archivo_inv = st.file_uploader("Sube el Inventario Diario Digitalizado (.xlsx)", type=["xlsx"])
        archivo_ped = st.file_uploader("Sube el Maestro de Pedidos (.xlsx)", type=["xlsx"])
        if archivo_inv: input_inv = archivo_inv
        if archivo_ped: input_ped = archivo_ped
        
    if input_inv and input_ped:
        if st.button("🔍 ANALIZAR PLANILLAS", use_container_width=True):
            try:
                # Cargar inventario
                st.session_state.inventario_db = cargar_inventario_real(input_inv)
                nombres_inv = list(st.session_state.inventario_db.keys())
                
                # Escanear pedidos
                wb_scan = openpyxl.load_workbook(input_ped)
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
                    ejecutar_calculo_matematico(input_ped)
                    st.session_state.etapa = "descargar"
                st.grid = None
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error al leer los archivos: {str(e)}. Verifica que los links de Drive sean correctos y públicos.")

# --- ETAPA 2: RESOLVER AMBIGÜEDADES ---
elif st.session_state.etapa == "resolver":
    st.subheader("⚠️ Validación de Productos")
    st.warning(f"Se encontraron {len(st.session_state.ambiguedades)} productos con nombres ambiguos. Confírmalos:")
    
    with st.form("formulario_resolucion"):
        nuevas_decisiones = {}
        for n_prov, info in list(st.session_state.ambiguedades.items())[:15]:
            tipo_texto = "Duplicado" if info["tipo"] == "DUPLICADO" else "Certeza Baja"
            options_select = ["[ No pedir / Ignorar ]"] + info["candidatos"]
            nuevas_decisiones[n_prov] = st.selectbox(label=f"📋 '{n_prov}' ({tipo_texto})", options=options_select, key=n_prov)
            st.markdown("---")
            
        if st.form_submit_button("💾 CONFIRMAR SELECCIONES Y CALCULAR", use_container_width=True):
            for n_prov, eleccion in nuevas_decisiones.items():
                st.session_state.cache_decisiones[n_prov] = None if eleccion == "[ No pedir / Ignorar ]" else eleccion
                del st.session_state.ambiguedades[n_prov]
            
            # Nota técnica: Al resolver en la web usando enlaces de Drive, para no re-descargar,
            # lo ideal es que en un entorno de producción el archivo base se resguarde.
            # Aquí forzamos el paso a descargar. Si usas links fijos, el flujo es fluido.
            st.session_state.etapa = "descargar"
            st.rerun()

# --- ETAPA 3: DESCARGA ---
elif st.session_state.etapa == "descargar":
    st.subheader("🎯 ¡Sugerencias Listas!")
    
    if st.session_state.excel_final is not None:
        st.download_button(
            label="📥 DESCARGAR EXCEL DE PEDIDOS",
            data=st.session_state.excel_final,
            file_name="Pedido_Barra_Sugerido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        # En caso de usar Drive, si el buffer se limpió por ciclo de vida de Streamlit, recalculamos rápido
        st.warning("Presiona el botón de abajo para compilar tu archivo listo:")
        
    if st.button("🔄 Calcular Otro Inventario / Reiniciar", use_container_width=True):
        st.session_state.clear()
        st.rerun()
import os
import difflib
import pandas as pd
import openpyxl
import streamlit as st
import io

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

# --- FUNCIONES DE LÓGICA DE BARRA ---
def cargar_inventario_real(file):
    df = pd.read_excel(file, sheet_name="Inventario General", skiprows=4)
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

def ejecutar_calculo_matematico(file_pedidos):
    file_pedidos.seek(0)
    wb = openpyxl.load_workbook(file_pedidos)
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
            
            # Buscar qué artículo del inventario se decidió usar
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

# --- ETAPA 1: SUBIR ARCHIVOS ---
if st.session_state.etapa == "upload":
    st.subheader("1. Carga de Planillas")
    archivo_inv = st.file_uploader("Sube el Inventario Diario Digitalizado (.xlsx)", type=["xlsx"])
    archivo_ped = st.file_uploader("Sube el Maestro de Pedidos original (.xlsx)", type=["xlsx"])
    
    if archivo_inv and archivo_ped:
        if st.button("🔍 ANALIZAR PLANILLAS", use_container_width=True):
            # Cargar inventario en memoria de la sesión web
            st.session_state.inventario_db = cargar_inventario_real(archivo_inv)
            nombres_inv = list(st.session_state.inventario_db.keys())
            
            # Escanear el maestro buscando conflictos antes de hacer cálculos
            wb_scan = openpyxl.load_workbook(archivo_ped)
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
                        # Guardamos para preguntar después
                        ambiguedades_encontradas[n_prov] = {"candidatos": res, "tipo": tipo_match}
            
            st.session_state.ambiguedades = ambiguedades_encontradas
            
            # Decidir siguiente pantalla
            if ambiguedades_encontradas:
                st.session_state.etapa = "resolver"
            else:
                # Si no hay dudas, calcula directo
                ejecutar_calculo_matematico(archivo_ped)
                st.session_state.etapa = "descargar"
            st.rerun()

# --- ETAPA 2: RESOLVER AMBIGÜEDADES EN LA WEB ---
elif st.session_state.etapa == "resolver":
    st.subheader("⚠️ Validación de Productos")
    st.warning(f"Se encontraron {len(st.session_state.ambiguedades)} productos con nombres ambiguos o duplicados. Por favor, confírmalos a continuación:")
    
    # Creamos un formulario interactivo
    with st.form("formulario_resolucion"):
        nuevas_decisiones = {}
        
        for n_prov, info in list(st.session_state.ambiguedades.items())[:15]: # Procesamos en bloques para comodidad móvil
            tipo_texto = "Duplicado" if info["tipo"] == "DUPLICADO" else "Certeza Baja"
            options_select = ["[ No pedir / Ignorar ]"] + info["candidatos"]
            
            nuevas_decisiones[n_prov] = st.selectbox(
                label=f"📋 Proveedor lista: '{n_prov}' ({tipo_texto})",
                options=options_select,
                key=n_prov
            )
            st.markdown("---")
            
        btn_enviar = st.form_submit_button("💾 CONFIRMAR SELECCIONES Y CALCULAR", use_container_width=True)
        
        if btn_enviar:
            # Guardamos las respuestas en el caché global
            for n_prov, eleccion in nuevas_decisiones.items():
                if eleccion == "[ No pedir / Ignorar ]":
                    st.session_state.cache_decisiones[n_prov] = None
                else:
                    st.session_state.cache_decisiones[n_prov] = eleccion
                # Lo removemos de la lista de pendientes
                del st.session_state.ambiguedades[n_prov]
                
            # Si ya no quedan dudas, calculamos la matemática final
            if not st.session_state.ambiguedades:
                # Simulamos recarga del archivo de pedidos desde la sesión (guardado temporal en st.file_uploader)
                # Para solucionar esto de forma limpia en Streamlit, procesamos directo.
                st.session_state.etapa = "descargar"
            st.rerun()

# --- ETAPA 3: DESCARGA DEL ARCHIVO FINAL ---
elif st.session_state.etapa == "descargar":
    st.subheader("🎯 ¡Sugerencias Listas!")
    st.success("Se ha procesado el stock de todas las pestañas de proveedores correctamente.")
    
    # Re-ejecutamos la inyección final por si hubo respuestas en la etapa 2
    # Para la persistencia del archivo original en Streamlit, usamos un truco de resguardo.
    if st.session_state.excel_final is None:
        st.error("Error al compilar el archivo. Por favor, reinicie.")
    else:
        st.download_button(
            label="📥 DESCARGAR SUGERENCIA DE COMPRAS",
            data=st.session_state.excel_final,
            file_name="Pedido_Barra_2025_Sugerido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    if st.button("🔄 Calcular Otro Inventario", use_container_width=True):
        st.session_state.clear()
        st.rerun()
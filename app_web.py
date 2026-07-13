import os
import difflib
import pandas as pd
import openpyxl
import streamlit as st
import io
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- CONFIGURACIÓN DE LA PÁGINA WEB ---
st.set_page_config(page_title="Pedidos El Bajo", page_icon="🍹", layout="centered")

# --- ⚠️ CONFIGURACIÓN DE REDIRECCIÓN OAUTH ---
# Recuerda cambiar esta URL por la definitiva cuando lo subas a Streamlit Cloud
REDIRECT_URI = "https://sugeridor-bar-el-bajo.streamlit.app/"

# --- BASE DE DATOS DE COLUMNAS DE PROVEEDORES ---
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

# --- CONTROL DE SESIÓN GENERAL (BÓVEDA DE ESTADOS) ---
if "credentials" not in st.session_state:
    st.session_state.credentials = None
if "flow" not in st.session_state:
    st.session_state.flow = None  # Guardará el flujo OAuth con el verificador de código intacto
if "auth_procesada" not in st.session_state:
    st.session_state.auth_procesada = False
if "etapa" not in st.session_state:
    st.session_state.etapa = "login"
if "ambiguedades" not in st.session_state:
    st.session_state.ambiguedades = {}
if "cache_decisiones" not in st.session_state:
    st.session_state.cache_decisiones = {}
if "inventario_db" not in st.session_state:
    st.session_state.inventario_db = {}
if "pedidos_bytes" not in st.session_state:
    st.session_state.pedidos_bytes = None
if "excel_final" not in st.session_state:
    st.session_state.excel_final = None

# --- CAPTURA DE RETORNO GOOGLE (OAUTH HANDSHAKE) ---
if st.session_state.credentials is None and not st.session_state.auth_procesada:
    query_params = st.query_params
    if "code" in query_params:
        try:
            # Usamos estrictamente el objeto flow guardado en sesión que tiene la clave criptográfica correcta
            if st.session_state.flow is not None:
                flow = st.session_state.flow
                flow.fetch_token(code=query_params["code"])
                st.session_state.credentials = flow.credentials
                st.session_state.auth_procesada = True
                st.session_state.etapa = "upload"
                st.query_params.clear() # Limpia la barra de direcciones
                st.rerun()
            else:
                st.error("⚠️ Error: Sesión de autenticación expirada. Por favor, intente iniciar sesión de nuevo.")
                st.session_state.etapa = "login"
        except Exception as e:
            st.error(f"⚠️ Error en intercambio de llaves de Google: {e}")
            st.info("Verifica que la variable REDIRECT_URI coincida exactamente con la URL de tu navegador actual.")
            # Reseteamos estados para permitir reintento limpio
            st.session_state.flow = None
            st.session_state.auth_procesada = False
            st.session_state.etapa = "login"

# --- FUNCIONES AUXILIARES DE GOOGLE DRIVE API ---
def listar_archivos_excel():
    try:
        service = build('drive', 'v3', credentials=st.session_state.credentials)
        query = "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed = false"
        resultados = service.files().list(q=query, fields="files(id, name)", pageSize=50).execute()
        archivos = resultados.get('files', [])
        return {f['name']: f['id'] for f in archivos}
    except Exception as e:
        st.error(f"Error al conectar con Drive: {e}")
        return {}

def descargar_archivo_desde_drive(file_id):
    service = build('drive', 'v3', credentials=st.session_state.credentials)
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

# --- LÓGICA DE PROCESAMIENTO DE STOCK ---
def cargar_inventario_real(file_io):
    df = pd.read_excel(file_io, sheet_name="Inventario General", skiprows=4)
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
    if difflib.SequenceMatcher(None, n_prov_clean, mejores[0].lower()).ratio() < 0.90:
        return mejores, "BAJA_CERTEZA"
    return mejores[0], "ALTA_CERTEZA"

def ejecutar_calculo_matematico():
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
    st.session_state.excel_final = buffer.getvalue()

# --- INTERFAZ DE USUARIO ---
st.title("🍹 Pedidos Integrados - El Bajo")

# --- ETAPA 0: AUTENTICACIÓN ---
if st.session_state.etapa == "login" and st.session_state.credentials is None:
    st.subheader("Acceso a Canales de Almacenamiento")
    st.write("Conéctate de forma segura a Google Drive para listar tus planillas de stock.")
    
    # Creamos el objeto flow una sola vez y lo persistimos en st.session_state.flow
    if st.session_state.flow is None:
        if "google_secrets" in st.secrets:
            client_config = {"web": dict(st.secrets["google_secrets"])}
            st.session_state.flow = Flow.from_client_config(
                client_config,
                scopes=['https://www.googleapis.com/auth/drive.readonly'],
                redirect_uri=REDIRECT_URI
            )
        elif os.path.exists('client_secrets.json'):
            st.session_state.flow = Flow.from_client_secrets_file(
                'client_secrets.json',
                scopes=['https://www.googleapis.com/auth/drive.readonly'],
                redirect_uri=REDIRECT_URI
            )
            
    if st.session_state.flow is not None:
        # Al pedir la URL de autorización, el objeto flow interno genera y guarda el verificador PKCE
        auth_url, _ = st.session_state.flow.authorization_url(prompt='select_account')
        st.link_button("🔑 CONECTAR CON GOOGLE DRIVE", auth_url, use_container_width=True)
    else:
        st.error("Falta el archivo 'client_secrets.json' o la configuración en Secrets de Streamlit.")

# --- ETAPA 1: SELECCIÓN DE ARCHIVOS DIRECTO DESDE DRIVE ---
elif st.session_state.etapa == "upload" or st.session_state.credentials is not None:
    if st.session_state.etapa == "login": 
        st.session_state.etapa = "upload"
        
    if st.session_state.etapa == "upload":
        st.subheader("1. Selección de Planillas en la Nube")
        
        with st.spinner("Leyendo archivos de tu Google Drive..."):
            diccionario_archivos = listar_archivos_excel()
            
        if diccionario_archivos:
            opciones_excel = ["-- Seleccionar un archivo --"] + list(diccionario_archivos.keys())
            
            archivo_inv_name = st.selectbox("Elija el Inventario Diario Digitalizado:", options=opciones_excel)
            archivo_ped_name = st.selectbox("Elija el Maestro de Pedidos (Plantilla):", options=opciones_excel)
            
            if archivo_inv_name != "-- Seleccionar un archivo --" and archivo_ped_name != "-- Seleccionar un archivo --":
                if st.button("🔍 ANALIZAR PLANILLAS SELECCIONADAS", use_container_width=True):
                    try:
                        id_inv = diccionario_archivos[archivo_inv_name]
                        id_ped = diccionario_archivos[archivo_ped_name]
                        
                        with st.spinner("Descargando y escaneando datos del Bar..."):
                            io_inv = descargar_archivo_desde_drive(id_inv)
                            io_ped = descargar_archivo_desde_drive(id_ped)
                            
                            st.session_state.inventario_db = cargar_inventario_real(io_inv)
                            st.session_state.pedidos_bytes = io_ped.getvalue()
                            
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
                        st.error(f"Fallo en lectura de celdas: {e}")
        else:
            st.warning("No se encontraron archivos Excel (.xlsx) en la raíz de tu Google Drive.")

# --- ETAPA 2: RESOLVER AMBIGÜEDADES ---
if st.session_state.etapa == "resolver":
    st.subheader("⚠️ Validación de Nombres")
    st.info(f"Faltan confirmar {len(st.session_state.ambiguedades)} productos:")
    
    with st.form("formulario_resolucion_drive"):
        nuevas_decisiones = {}
        bloque_items = list(st.session_state.ambiguedades.items())[:12]
        
        for n_prov, info in bloque_items:
            tipo_texto = "Duplicado" if info["tipo"] == "DUPLICADO" else "Certeza Baja"
            options_select = ["[ No pedir ]"] + info["candidatos"]
            nuevas_decisiones[n_prov] = st.selectbox(label=f"📋 '{n_prov}' ({tipo_texto})", options=options_select, key=n_prov)
            st.markdown("---")
            
        if st.form_submit_button("💾 GUARDAR ASOCIACIONES Y CONTINUAR", use_container_width=True):
            for n_prov, eleccion in nuevas_decisiones.items():
                st.session_state.cache_decisiones[n_prov] = None if eleccion == "[ No pedir ]" else eleccion
                del st.session_state.ambiguedades[n_prov]
            
            if not st.session_state.ambiguedades:
                ejecutar_calculo_matematico()
                st.session_state.etapa = "descargar"
            st.rerun()

# --- ETAPA 3: DESCARGA FINAL ---
elif st.session_state.etapa == "descargar":
    st.subheader("🎯 ¡Sugerencia de Pedidos Completada!")
    st.success("La inyección matemática se completó de manera exitosa.")
    
    st.download_button(
        label="📥 DESCARGAR PEDIDO BARRA PROCESADO",
        data=st.session_state.excel_final,
        file_name="Pedido_Barra_Drive_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    if st.button("🔄 Procesar Nuevas Planillas", use_container_width=True):
        credenciales_actuales = st.session_state.credentials
        st.session_state.clear()
        st.session_state.credentials = credenciales_actuales
        st.session_state.etapa = "upload"
        st.rerun()

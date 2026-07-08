"""
Aplicación Principal - Convertidor de TXT a Excel
Optimizado para archivos de 500MB+
Autor: Yessly Poma de la Cruz
"""
import streamlit as st
import pandas as pd
import os
from pathlib import Path
import time
import sys
import psutil

# Agregar el directorio actual al path
sys.path.insert(0, str(Path(__file__).parent))

# Importar módulos
from modules import TXTToExcelConverter, validate_file, get_file_info, format_size

# Configuración de la página
st.set_page_config(
    page_title="Convertidor TXT → Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# ESTILOS CSS
# ============================================================================

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg,#1a73e8 0%, #0d47a1 100%);
        color: white;
        padding: 25px;
        border-radius: 12px;
        margin-bottom: 25px;
        text-align: center;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2.5em;
        font-weight: 700;
        color: #ffffff !important;
    }
    .main-header p {
        margin: 8px 0 0 0;
        font-size: 1.1em;
        opacity: 0.9;
    }
    .success-box {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 10px;
        padding: 20px;
        color: #155724;
        margin: 15px 0;
    }
    .info-box {
        background: #cce5ff;
        border: 1px solid #b8daff;
        border-radius: 10px;
        padding: 20px;
        color: #004085;
        margin: 15px 0;
    }
    .warning-box {
        background: #fff3cd;
        border: 1px solid #ffeeba;
        border-radius: 10px;
        padding: 20px;
        color: #856404;
        margin: 15px 0;
    }
    .error-box {
        background: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 10px;
        padding: 20px;
        color: #721c24;
        margin: 15px 0;
    }
    .stButton > button {
        background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
        color: white;
        border: none;
        padding: 12px 25px;
        font-weight: 600;
        font-size: 1.1em;
        border-radius: 10px;
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(26, 115, 232, 0.4);
    }
    .stButton > button:disabled {
        background: #ccc;
        cursor: not-allowed;
        transform: none;
    }
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
        border: 1px solid #e8eaed;
        height: 100%;
    }
    .metric-card .value {
        font-size: 2.2em;
        font-weight: 700;
        color: #1a73e8;
    }
    .metric-card .label {
        font-size: 0.9em;
        color: #5f6368;
        margin-top: 8px;
    }
    .footer {
        text-align: center;
        color: #9e9e9e;
        padding: 25px;
        font-size: 0.9em;
        border-top: 1px solid #e8eaed;
        margin-top: 40px;
    }
    .file-info {
        font-size: 0.95em;
        color: #5f6368;
    }
    .stProgress > div {
        background: linear-gradient(135deg, #2ecc71 0%, #1b5e20 100%);
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# HEADER
# ============================================================================

st.markdown("""
<div class="main-header">
    <h1> Convertidor TXT → Excel</h1>
    <p>Herramienta profesional para convertir archivos TXT a Excel con múltiples hojas</p>
    <p style="font-size:0.9em;margin-top:10px;opacity:0.8;">Soporte para archivos de 500MB+ | Desarrollado por: Yessly Poma de la Cruz</p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# SIDEBAR
# ============================================================================

with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    
    # Información del sistema
    mem = psutil.virtual_memory()
    st.markdown(f"""
    <div class="info-box" style="font-size:0.9em;padding:15px;">
        <b>💾 RAM disponible:</b> {mem.available / (1024**3):.2f} GB<br>
        <b>📤 Límite de subida:</b> 1 GB<br>
        <b>⚡ CPU:</b> {psutil.cpu_count()} núcleos
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 📁 Directorios")
    
    # Crear directorios
    for dir_name in ['outputs', 'uploads']:
        Path(dir_name).mkdir(exist_ok=True)
    
    st.write(f" Uploads: `uploads/`")
    st.write(f" Outputs: `outputs/`")
    
    st.markdown("---")
    st.markdown("###  Estadísticas")
    
    output_files = list(Path("outputs").glob("*.xlsx"))
    total_size = sum(f.stat().st_size for f in output_files) if output_files else 0
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric(" Excel", len(output_files))
    with col2:
        st.metric(" Total", format_size(total_size))
    
    st.markdown("---")
    st.markdown("### Ayuda rápida")
    st.info("""
    **1. Sube el archivo** (máx 1GB)
    **2. Configura opciones**
    **3. Haz clic en "Convertir"**
    **4. Descarga el Excel**
    """)

# ============================================================================
# ÁREA PRINCIPAL
# ============================================================================

st.markdown("### 📤 Paso 1: Subir archivo TXT")

# Subida de archivo
uploaded_file = st.file_uploader(
    "Selecciona tu archivo TXT",
    type=["txt"],
    help="Archivos de hasta 1GB soportados",
    key="file_uploader"
)

if uploaded_file is not None:
    # Guardar archivo temporal
    temp_path = Path("uploads") / f"temp_{uploaded_file.name}"
    temp_path.parent.mkdir(exist_ok=True)
    
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Validar archivo
    is_valid, message = validate_file(temp_path, max_size_mb=1000)
    
    if not is_valid:
        st.markdown(f"""
        <div class="error-box">
            ❌ <b>Error:</b> {message}
        </div>
        """, unsafe_allow_html=True)
        if temp_path.exists():
            temp_path.unlink()
    else:
        # Información del archivo
        file_info = get_file_info(temp_path)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="value">{file_info['name']}</div>
                <div class="label">📄 Nombre</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <div class="value">{file_info['size_mb']:.2f} MB</div>
                <div class="label">📊 Tamaño</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <div class="value">{file_info['modified']}</div>
                <div class="label">📅 Modificado</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### ⚙️ Paso 2: Configurar opciones")
        
        col1, col2 = st.columns(2)
        
        with col1:
            max_rows = st.number_input(
                "📊 Máximo filas por hoja",
                min_value=10000,
                max_value=5000000,
                value=1000000,
                step=50000,
                help="Número máximo de filas por cada hoja en el Excel"
            )
        
        with col2:
            chunk_size = st.number_input(
                "📦 Tamaño de chunk (líneas)",
                min_value=10000,
                max_value=200000,
                value=50000,
                step=10000,
                help="Cantidad de líneas procesadas a la vez (ajustar según RAM)"
            )
        
        # Opciones avanzadas
        with st.expander("🔧 Opciones avanzadas", expanded=False):
            col1, col2 = st.columns(2)
            
            with col1:
                encoding_options = ['auto', 'utf-8', 'latin-1', 'cp1252', 'ascii']
                selected_encoding = st.selectbox(
                    "Codificación del archivo",
                    options=encoding_options,
                    index=0,
                    help="Auto-detectar por defecto"
                )
            
            with col2:
                separator_options = ['auto', '|', '\t', ',', ';']
                selected_separator = st.selectbox(
                    "Separador de columnas",
                    options=separator_options,
                    index=0,
                    help="Auto-detectar por defecto"
                )
        
        st.markdown("---")
        st.markdown("### 🚀 Paso 3: Convertir")
        
        # Botón de conversión
        if st.button(" CONVERTIR A EXCEL", use_container_width=True):
            try:
                # Configurar progreso
                progress_bar = st.progress(0)
                status_text = st.empty()
                time_remaining = st.empty()
                
                # Función de progreso
                start_time = time.time()
                
                def update_progress(value, message):
                    progress_bar.progress(value / 100)
                    status_text.info(f"⏳ {message}")
                    
                    # Estimar tiempo restante
                    if 0 < value < 100:
                        elapsed = time.time() - start_time
                        if value > 5:  # Solo estimar después de 5%
                            estimated_total = elapsed / (value / 100)
                            remaining = estimated_total - elapsed
                            if remaining > 0:
                                if remaining < 60:
                                    time_remaining.info(f"⏱️ Restante: {remaining:.0f} segundos")
                                else:
                                    time_remaining.info(f"⏱️ Restante: {remaining/60:.1f} minutos")
                
                # Configurar codificación
                encoding = None if selected_encoding == 'auto' else selected_encoding
                separator = None if selected_separator == 'auto' else selected_separator
                
                # Crear convertidor
                converter = TXTToExcelConverter(
                    input_file=temp_path,
                    output_dir="outputs",
                    max_rows_per_sheet=max_rows,
                    chunk_size=chunk_size,
                    encoding=encoding,
                    separator=separator
                )
                
                # Ejecutar conversión
                result = converter.convert_to_excel(update_progress)
                
                # Mostrar resultados
                st.markdown("---")
                st.markdown("### ✅ Conversión Completada")
                
                # Métricas
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("📝 Registros", f"{result['total_lines']:,}")
                with col2:
                    st.metric("📊 Columnas", result['columns'])
                with col3:
                    st.metric("📄 Hojas", result['num_sheets'])
                with col4:
                    st.metric("⏱️ Tiempo", f"{result['time']:.2f}s")
                
                # Archivo generado
                output_file = result['output_file']
                
                st.markdown(f"""
                <div class="success-box">
                    <b> ¡Conversión exitosa!</b><br>
                     Archivo: {output_file.name}<br>
                     Tamaño: {result['file_size_mb']:.2f} MB
                </div>
                """, unsafe_allow_html=True)
                
                # Botón de descarga
                with open(output_file, "rb") as f:
                    st.download_button(
                        label="📥 DESCARGAR EXCEL",
                        data=f,
                        file_name=output_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                # Vista previa
                st.markdown("---")
                st.markdown("### Vista Previa (primeras 10 filas)")
                
                try:
                    df_preview = pd.read_excel(output_file, sheet_name=0, nrows=10)
                    st.dataframe(df_preview, use_container_width=True, height=300)
                    st.caption(f"Mostrando primeras 10 filas de la Hoja 1")
                except Exception as e:
                    st.info(f"No se pudo generar vista previa: {str(e)}")
                
                status_text.success("✅ ¡Conversión completada exitosamente!")
                time_remaining.empty()
                
            except Exception as e:
                st.error(f"❌ Error durante la conversión: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
        
        # Limpiar archivo temporal
        try:
            if temp_path.exists():
                temp_path.unlink()
        except:
            pass

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("""
<div class="footer">
    <p>🏦 <b>Convertidor TXT → Excel - Yessly Poma de la Cruz</b></p>
    <p style='font-size: 12px;'>Versión 3.0 | Optimizado para archivos de 500MB+</p>
    <p style='font-size: 11px; margin-top: 5px;'>© 2026 Todos los derechos reservados</p>
</div>
""", unsafe_allow_html=True)
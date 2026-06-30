"""
Utilidades para el procesamiento de archivos
"""
import os
import re
from pathlib import Path
from datetime import datetime
import psutil

def validate_file(file_path, max_size_mb=1000):
    """
    Valida el archivo antes de procesarlo
    
    Args:
        file_path: Ruta del archivo
        max_size_mb: Tamaño máximo permitido en MB
    
    Returns:
        tuple: (is_valid, message)
    """
    file_path = Path(file_path)
    
    # Verificar que existe
    if not file_path.exists():
        return False, f"El archivo {file_path.name} no existe"
    
    # Verificar extensión
    if file_path.suffix.lower() != '.txt':
        return False, f"El archivo debe ser .txt (extensión actual: {file_path.suffix})"
    
    # Verificar tamaño
    size_mb = file_path.stat().st_size / (1024 * 1024)
    if size_mb > max_size_mb:
        return False, f"El archivo es muy grande ({size_mb:.2f} MB). Límite: {max_size_mb} MB"
    
    # Verificar espacio en disco (corregido)
    try:
        # Obtener la unidad del archivo
        drive = str(file_path.drive)
        if not drive:
            # Si no tiene drive (ruta relativa), usar el directorio actual
            drive = os.getcwd()
        
        # Obtener espacio libre
        free_space = psutil.disk_usage(drive).free / (1024**3)
        if free_space < 1:  # Menos de 1GB libre
            return False, f"Espacio en disco insuficiente. Se necesita al menos 1GB libre (solo {free_space:.2f}GB disponible)"
    except Exception as e:
        # Si no se puede verificar el espacio, continuar
        print(f"⚠️ No se pudo verificar espacio en disco: {e}")
        pass
    
    return True, f"✅ Archivo válido ({size_mb:.2f} MB)"

def get_file_info(file_path):
    """
    Obtiene información detallada del archivo
    """
    file_path = Path(file_path)
    stats = file_path.stat()
    
    return {
        'name': file_path.name,
        'size_mb': stats.st_size / (1024 * 1024),
        'size_bytes': stats.st_size,
        'modified': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
        'extension': file_path.suffix,
        'parent': str(file_path.parent)
    }

def clean_filename(filename):
    """
    Limpia el nombre del archivo para evitar problemas
    """
    # Eliminar caracteres no permitidos
    clean = re.sub(r'[<>:"/\\|?*]', '_', filename)
    # Eliminar espacios al inicio y final
    clean = clean.strip()
    return clean

def format_size(size_bytes):
    """
    Formatea el tamaño en bytes a una cadena legible
    """
    if size_bytes == 0:
        return "0 B"
    
    size_bytes = float(size_bytes)
    units = ['B', 'KB', 'MB', 'GB', 'TB']
    
    for unit in units:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    
    return f"{size_bytes:.2f} PB"

def estimate_processing_time(file_size_mb, lines_per_second=50000):
    """
    Estima el tiempo de procesamiento
    """
    # Estimación: 50,000 líneas por segundo (promedio)
    estimated_seconds = (file_size_mb * 1024 * 1024) / (lines_per_second * 100)
    if estimated_seconds < 60:
        return f"{estimated_seconds:.0f} segundos"
    elif estimated_seconds < 3600:
        return f"{estimated_seconds/60:.1f} minutos"
    else:
        return f"{estimated_seconds/3600:.1f} horas"
"""
Módulo para convertir archivos TXT a Excel con múltiples hojas
Optimizado para archivos de 500MB+
"""
import pandas as pd
import numpy as np
import re
from pathlib import Path
import chardet
import time
from datetime import datetime
import gc
import warnings
warnings.filterwarnings('ignore')

class TXTToExcelConverter:
    """
    Clase para convertir archivos TXT a Excel con múltiples hojas
    Optimizado para archivos grandes (>500MB)
    """
    
    def __init__(self, input_file, output_dir="outputs", max_rows_per_sheet=1000000, 
                 chunk_size=50000, encoding=None, separator=None):
        self.input_file = Path(input_file)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.max_rows_per_sheet = max_rows_per_sheet
        self.chunk_size = chunk_size
        self.encoding = encoding
        self.separator = separator
        self.columns = None
        self.has_header = False
        self.total_lines = 0
        self.file_size_mb = self.input_file.stat().st_size / (1024**2)
        
    def clean_value(self, value):
        """Limpia caracteres no válidos para Excel"""
        if pd.isna(value) or value is None:
            return ''
        
        value = str(value)
        # Eliminar caracteres de control
        value = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', value)
        # Reemplazar caracteres problemáticos
        value = value.replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
        # Eliminar espacios múltiples
        value = re.sub(r'\s+', ' ', value)
        return value.strip()
    
    def detect_encoding(self):
        """Detecta la codificación del archivo"""
        print("🔍 Detectando codificación...")
        with open(self.input_file, 'rb') as f:
            raw_data = f.read(10000)
            result = chardet.detect(raw_data)
            self.encoding = result['encoding'] or 'utf-8'
            confidence = result['confidence']
            print(f"✅ Codificación: {self.encoding} (confianza: {confidence:.2%})")
            return self.encoding
    
    def detect_separator(self):
        """Detecta el separador de columnas"""
        print("🔍 Detectando separador...")
        with open(self.input_file, 'r', encoding=self.encoding or 'utf-8') as f:
            sample = [next(f) for _ in range(min(10, 20))]
        
        separators = ['|', '\t', ',', ';', ' ']
        best_sep = '|'
        best_count = 0
        best_consistency = 0
        
        for sep in separators:
            counts = []
            for line in sample:
                if line.strip():
                    counts.append(len(line.split(sep)))
            
            if counts:
                # Verificar consistencia (todas las líneas tienen el mismo número de columnas)
                consistency = counts.count(counts[0]) / len(counts) if counts else 0
                avg_count = sum(counts) / len(counts) if counts else 0
                
                if consistency > best_consistency and avg_count > 1:
                    best_consistency = consistency
                    best_sep = sep
                    best_count = int(avg_count)
        
        self.separator = best_sep
        print(f"✅ Separador: '{best_sep}' ({best_count} columnas)")
        return self.separator
    
    def detect_header(self):
        """Detecta si el archivo tiene encabezado"""
        print("🔍 Detectando encabezado...")
        
        with open(self.input_file, 'r', encoding=self.encoding or 'utf-8') as f:
            first_line = f.readline().strip()
            second_line = f.readline().strip() if f else ''
            f.seek(0)
            third_line = f.readline().strip() if f else ''
        
        # Limpiar líneas
        first_line = self.clean_value(first_line)
        second_line = self.clean_value(second_line)
        third_line = self.clean_value(third_line)
        
        parts1 = first_line.split(self.separator) if first_line else []
        parts2 = second_line.split(self.separator) if second_line else []
        parts3 = third_line.split(self.separator) if third_line else []
        
        # Si la primera línea tiene pocas columnas comparada con las demás, es encabezado
        if len(parts1) < len(parts2) and len(parts2) == len(parts3):
            self.has_header = True
            self.columns = [self.clean_value(col) or f'Col_{i+1}' for i, col in enumerate(parts1)]
            print(f"✅ Encabezado detectado: {len(self.columns)} columnas")
            return self.has_header, self.columns
        
        # Verificar si la primera línea es numérica
        try:
            numeric_count = 0
            for p in parts1[:min(5, len(parts1))]:
                if p:
                    try:
                        float(p.replace(',', '.'))
                        numeric_count += 1
                    except:
                        pass
            
            # Si más del 80% son numéricos, no es encabezado
            if len(parts1) > 0 and numeric_count / len(parts1) > 0.8:
                self.has_header = False
                self.columns = [f'Col_{i+1}' for i in range(len(parts1))]
                print(f"✅ Sin encabezado: {len(self.columns)} columnas generadas")
                return self.has_header, self.columns
        except:
            pass
        
        # Por defecto, asumir que tiene encabezado
        self.has_header = True
        self.columns = [self.clean_value(col) or f'Col_{i+1}' for i, col in enumerate(parts1)]
        print(f"✅ Asumiendo encabezado: {len(self.columns)} columnas")
        return self.has_header, self.columns
    
    def count_lines_fast(self):
        """Cuenta las líneas del archivo de manera eficiente"""
        print("📊 Contando líneas...")
        lines = 0
        buffer_size = 1024 * 1024  # 1MB buffer
        
        with open(self.input_file, 'rb') as f:
            while True:
                chunk = f.read(buffer_size)
                if not chunk:
                    break
                lines += chunk.count(b'\n')
        
        self.total_lines = lines
        print(f"📝 Total de líneas: {lines:,}")
        return lines
    
    def process_chunk(self, chunk_data, sheet_name):
        """Procesa un chunk de datos y retorna un DataFrame"""
        try:
            # Limpiar datos
            cleaned_data = []
            for row in chunk_data:
                cleaned_row = [self.clean_value(cell) for cell in row]
                cleaned_data.append(cleaned_row)
            
            # Crear DataFrame
            df = pd.DataFrame(cleaned_data, columns=self.columns)
            
            # Limpiar nombres de columnas
            df.columns = [self.clean_value(col) or f'Col_{i+1}' for i, col in enumerate(df.columns)]
            
            # Reemplazar vacíos por NaN
            df = df.replace('', np.nan)
            df = df.replace(r'^\s*$', np.nan, regex=True)
            
            return df
            
        except Exception as e:
            print(f"❌ Error procesando chunk: {str(e)}")
            return None
    
    def convert_to_excel(self, progress_callback=None):
        """
        Convierte el archivo TXT a Excel con múltiples hojas
        """
        start_time = time.time()
        
        print("=" * 70)
        print("🚀 CONVIRTIENDO TXT A EXCEL")
        print("=" * 70)
        print(f"📁 Archivo: {self.input_file.name} ({self.file_size_mb:.2f} MB)")
        print(f"📊 Tamaño máximo por hoja: {self.max_rows_per_sheet:,} filas")
        
        # Detectar codificación
        if not self.encoding:
            self.detect_encoding()
        
        # Detectar separador
        if not self.separator:
            self.detect_separator()
        
        # Detectar encabezado
        self.detect_header()
        
        # Contar líneas
        if self.total_lines == 0:
            self.count_lines_fast()
        
        # Calcular número de hojas
        num_sheets = max(1, (self.total_lines + self.max_rows_per_sheet - 1) // self.max_rows_per_sheet)
        lines_per_sheet = max(1, self.total_lines // num_sheets)
        
        # Nombre del archivo de salida
        base_name = self.input_file.stem
        output_file = self.output_dir / f'CONVERTIDO_{base_name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        print(f"\n📊 Generando {num_sheets} hojas")
        print(f"📊 {lines_per_sheet:,} líneas por hoja")
        print(f"📊 Procesando en chunks de {self.chunk_size:,} líneas\n")
        
        # Procesar y guardar
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            current_sheet = 1
            chunk_data = []
            rows_written = 0
            
            with open(self.input_file, 'r', encoding=self.encoding, errors='ignore') as f:
                # Saltar encabezado si existe
                if self.has_header:
                    f.readline()
                    if progress_callback:
                        progress_callback(5, '📋 Encabezado omitido')
                
                # Procesar líneas
                for line_num, line in enumerate(f, 1):
                    # Limpiar línea
                    line = self.clean_value(line)
                    if not line:
                        continue
                    
                    parts_line = line.split(self.separator)
                    
                    # Asegurar número correcto de columnas
                    if len(parts_line) < len(self.columns):
                        parts_line.extend([''] * (len(self.columns) - len(parts_line)))
                    elif len(parts_line) > len(self.columns):
                        parts_line = parts_line[:len(self.columns)]
                    
                    chunk_data.append(parts_line)
                    rows_written += 1
                    
                    # Procesar cuando se alcanza el chunk_size
                    if len(chunk_data) >= self.chunk_size:
                        df = self.process_chunk(chunk_data, f'Hoja_{current_sheet}')
                        if df is not None:
                            sheet_name = f'Hoja_{current_sheet}'
                            
                            # Si es la primera hoja, escribir con encabezado
                            if current_sheet == 1 and not writer.sheets:
                                df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep='')
                            else:
                                # Si la hoja ya existe, agregar sin encabezado
                                if sheet_name in writer.sheets:
                                    startrow = writer.sheets[sheet_name].max_row
                                    df.to_excel(writer, sheet_name=sheet_name, index=False, 
                                              na_rep='', startrow=startrow, header=False)
                                else:
                                    df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep='')
                            
                            if progress_callback:
                                progress = min(95, (rows_written / self.total_lines) * 100)
                                progress_callback(progress, f'📊 Hoja {current_sheet}: {rows_written:,} filas')
                        
                        chunk_data = []
                        gc.collect()
                    
                    # Cambiar de hoja
                    if rows_written >= lines_per_sheet and line_num < self.total_lines:
                        current_sheet += 1
                        rows_written = 0
                        if progress_callback:
                            progress = min(95, (current_sheet / num_sheets) * 100)
                            progress_callback(progress, f'📄 Cambiando a Hoja {current_sheet}')
                
                # Procesar datos restantes
                if chunk_data:
                    df = self.process_chunk(chunk_data, f'Hoja_{current_sheet}')
                    if df is not None:
                        sheet_name = f'Hoja_{current_sheet}'
                        if sheet_name in writer.sheets:
                            startrow = writer.sheets[sheet_name].max_row
                            df.to_excel(writer, sheet_name=sheet_name, index=False, 
                                      na_rep='', startrow=startrow, header=False)
                        else:
                            df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep='')
                    chunk_data = []
                    gc.collect()
        
        elapsed_time = time.time() - start_time
        
        # Resultados
        result = {
            'output_file': output_file,
            'total_lines': self.total_lines,
            'num_sheets': len(writer.sheets) if hasattr(writer, 'sheets') else current_sheet,
            'columns': len(self.columns),
            'time': elapsed_time,
            'file_size_mb': output_file.stat().st_size / (1024**2) if output_file.exists() else 0
        }
        
        print("\n" + "=" * 70)
        print("✅ ¡CONVERSIÓN COMPLETADA!")
        print("=" * 70)
        print(f"📁 Archivo: {output_file.name}")
        print(f"📊 Hojas generadas: {result['num_sheets']}")
        print(f"📝 Total registros: {result['total_lines']:,}")
        print(f"⏱️ Tiempo: {elapsed_time:.2f} segundos")
        print(f"💾 Tamaño: {result['file_size_mb']:.2f} MB")
        print("=" * 70)
        
        if progress_callback:
            progress_callback(100, '✅ ¡Conversión completada!')
        
        return result
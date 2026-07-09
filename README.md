# 📊 Convertidor TXT → Excel

Una herramienta profesional y robusta para convertir archivos de texto (TXT) a formato Excel, optimizada para procesar archivos de gran tamaño (500MB+).

## 🎯 Características Principales

- ✅ **Soporte para archivos grandes**: Procesa archivos de hasta 1GB sin problemas
- ✅ **Conversión inteligente**: Detecta automáticamente:
  - Codificación del archivo (UTF-8, Latin-1, CP1252, ASCII, etc.)
  - Separador de columnas (|, tabulación, coma, punto y coma, espacio)
  - Presencia de encabezados en el archivo
- ✅ **Múltiples hojas**: Distribuye automáticamente los datos en hojas Excel cuando es necesario
- ✅ **Interfaz web amigable**: Utiliza Streamlit para una experiencia de usuario intuitiva
- ✅ **Limpieza de datos**: Elimina caracteres inválidos y normaliza formatos
- ✅ **Indicador de progreso**: Muestra progreso en tiempo real con estimación de tiempo restante
- ✅ **Información del sistema**: Monitorea RAM disponible y espacio en disco
- ✅ **Vista previa**: Visualiza las primeras 10 filas del archivo convertido

## 📋 Requisitos

### Sistema Operativo
- Windows, macOS o Linux
- 4GB RAM mínimo (recomendado 8GB+)
- 2GB espacio libre en disco

### Requisitos de Software
- Python 3.8 o superior
- pip (gestor de paquetes de Python)

## 🚀 Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/yessLU11/ConvertidorTXTaEXCEL.git
cd ConvertidorTXTaEXCEL
```

### 2. Crear un entorno virtual (recomendado)

```bash
# En Windows
python -m venv venv
venv\Scripts\activate

# En macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### Uso rápido en Windows

Ejecuta el archivo batch proporcionado:

```bash
install.bat    # Instala las dependencias
run.bat         # Inicia la aplicación
```

## 💻 Uso

### Opción 1: Interfaz Web (Recomendado)

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en tu navegador (http://localhost:8501)

### Uso desde la interfaz web

1. **Subir archivo**: Selecciona tu archivo TXT (máximo 1GB)
2. **Configurar opciones**: 
   - Máximo de filas por hoja (por defecto: 1,000,000)
   - Tamaño del chunk procesado (por defecto: 50,000 líneas)
   - Opciones avanzadas: codificación y separador personalizados
3. **Convertir**: Haz clic en "CONVERTIR A EXCEL"
4. **Descargar**: Descarga el archivo Excel generado
5. **Vista previa**: Visualiza un adelanto de los datos

### Opción 2: Uso programático

```python
from modules import TXTToExcelConverter

# Crear convertidor
converter = TXTToExcelConverter(
    input_file="archivo.txt",
    output_dir="outputs",
    max_rows_per_sheet=1000000,
    chunk_size=50000
)

# Convertir
result = converter.convert_to_excel()

# Resultados
print(f"Hojas generadas: {result['num_sheets']}")
print(f"Registros totales: {result['total_lines']}")
print(f"Tiempo: {result['time']:.2f} segundos")
```

## 📦 Dependencias

| Paquete | Versión | Descripción |
|---------|---------|-------------|
| streamlit | 1.28.0 | Framework para la interfaz web |
| pandas | 2.2.2 | Procesamiento de datos |
| openpyxl | 3.1.2 | Manejo de archivos Excel |
| chardet | 5.2.0 | Detección de codificación |
| psutil | 5.9.8 | Información del sistema |
| numpy | 1.26.4 | Operaciones numéricas |

## 📂 Estructura del Proyecto

```
ConvertidorTXTaEXCEL/
├── app.py                  # Aplicación principal (Streamlit)
├── requirements.txt        # Dependencias del proyecto
├── install.bat            # Script de instalación (Windows)
├── run.bat                # Script para ejecutar (Windows)
├── modules/
│   ├── __init__.py        # Inicializador del módulo
│   ├── converter.py       # Lógica de conversión
│   └── utils.py           # Funciones auxiliares
├── uploads/               # Directorio de archivos subidos
├── outputs/               # Directorio de archivos generados
└── README.md              # Este archivo
```

## 🔧 Características Avanzadas

### Detección Automática de Configuración

El convertidor detecta automáticamente:

- **Codificación**: Analiza los primeros 10KB del archivo
- **Separador**: Prueba múltiples separadores (|, tabulación, coma, punto y coma, espacio)
- **Encabezado**: Identifica si la primera fila contiene encabezados
- **Dimensiones**: Calcula automáticamente el número de hojas necesarias

### Limpieza de Datos

- Elimina caracteres de control inválidos para Excel
- Reemplaza saltos de línea y tabulaciones
- Normaliza espacios múltiples
- Valida el número de columnas en cada fila

### Optimización de Memoria

- Procesa archivos en chunks (bloques) configurables
- Libera memoria automáticamente entre chunks
- Utiliza streaming para archivos muy grandes
- No carga el archivo completo en memoria

## 📊 Ejemplos de Uso

### Ejemplo 1: Convertir archivo simple

```bash
streamlit run app.py
# 1. Selecciona archivo.txt
# 2. Haz clic en "CONVERTIR A EXCEL"
# 3. Descarga el resultado
```

### Ejemplo 2: Archivo con separador personalizado

```bash
streamlit run app.py
# 1. Sube archivo.txt (separado por |)
# 2. En "Opciones avanzadas", selecciona separador |
# 3. Haz clic en "CONVERTIR A EXCEL"
```

### Ejemplo 3: Archivo grande (500MB+)

```python
from modules import TXTToExcelConverter

converter = TXTToExcelConverter(
    input_file="archivo_grande.txt",
    max_rows_per_sheet=500000,  # Reduce filas por hoja
    chunk_size=100000            # Aumenta chunk para mejor rendimiento
)

result = converter.convert_to_excel()
```

## 🎨 Interfaz de Usuario

La aplicación incluye:

- **Encabezado profesional** con branding
- **Barra lateral** con información del sistema
- **Indicador de progreso** visual en tiempo real
- **Estimación de tiempo restante**
- **Métricas** del archivo procesado
- **Vista previa** de datos
- **Mensajes informativos** claros

## ⚙️ Configuración Recomendada

| Parámetro | Valor | Descripción |
|-----------|-------|-------------|
| Máximo filas/hoja | 1,000,000 | Evita problemas con Excel |
| Chunk size | 50,000 | Balance entre memoria y velocidad |
| Encoding | auto | Detecta automáticamente |
| Separador | auto | Detecta automáticamente |

## 🐛 Resolución de Problemas

### "El archivo es muy grande"
- Reduce el tamaño del archivo o aumenta el límite en la configuración
- Comprime el archivo antes de subirlo

### "Error en la detección de codificación"
- Especifica manualmente la codificación en "Opciones avanzadas"
- Prueba UTF-8, Latin-1 o CP1252

### "El Excel tiene pocas columnas"
- Verifica que el separador sea correcto
- Usa "Opciones avanzadas" para especificar el separador

### "Error de memoria"
- Reduce el tamaño del chunk
- Aumenta el límite de filas por hoja
- Cierra otras aplicaciones

### "Archivos no se generan"
- Verifica que tengas permisos de escritura en la carpeta `outputs/`
- Asegúrate de tener al menos 1GB de espacio libre

## 📈 Rendimiento

Velocidades aproximadas (dependen del hardware):

| Tamaño Archivo | RAM | Tiempo |
|---|---|---|
| 10 MB | 4GB | < 2 segundos |
| 100 MB | 4GB | 5-10 segundos |
| 500 MB | 8GB | 30-60 segundos |
| 1 GB | 16GB | 60-120 segundos |

## 🔒 Privacidad y Seguridad

- ✅ Los archivos se procesan localmente
- ✅ No se envían datos a servidores externos
- ✅ Los archivos temporales se limpian automáticamente
- ✅ No se almacenan datos sensibles

## 📝 Limitaciones Conocidas

- Máximo 1GB de tamaño de archivo
- Excel tiene límite de ~1 millón de filas por hoja (configurado a 1M para seguridad)
- Los nombres de archivo muy largos se truncan automáticamente

## 🤝 Contribuir

Las contribuciones son bienvenidas. Para contribuir:

1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

## 👨‍💼 Autor

**Yessly Poma de la Cruz**

- GitHub: [@yessLU11](https://github.com/yessLU11)
- Email: [Tu email si deseas]

## 📞 Soporte

¿Tienes preguntas o encontraste un bug?

- Abre un issue en [GitHub Issues](https://github.com/yessLU11/ConvertidorTXTaEXCEL/issues)
- Revisa la documentación en este README
- Consulta los ejemplos proporcionados

## 🎉 Agradecimientos

- Streamlit por la excelente framework
- Pandas y OpenPyXL por las librerías de datos
- La comunidad de código abierto

---

**Versión**: 3.0  
**Última actualización**: 2026  
**Estado**: ✅ Activo y mantenido

## 🚀 Hoja de Ruta (Roadmap)

Características planeadas para futuras versiones:

- [ ] Soporte para archivos CSV de salida
- [ ] Validación y corrección automática de datos
- [ ] Compresión automática de Excel
- [ ] Historial de conversiones
- [ ] Configuración guardada por usuario
- [ ] API REST
- [ ] Aplicación de escritorio
- [ ] Soporte para más formatos de entrada

---

**¡Gracias por usar Convertidor TXT → Excel! 🎊**

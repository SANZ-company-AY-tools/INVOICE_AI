# Extractor de Facturas con IA

Aplicación web para extraer datos de facturas (PDF e imágenes) usando **Claude AI** y generar un Excel con la información estructurada.

## Características

- **Extracción con IA**: Usa Claude Vision para entender facturas de cualquier formato
- Interfaz web con drag & drop para subir facturas
- Soporte para facturas españolas e internacionales
- Procesamiento de múltiples archivos simultáneamente
- Formatos soportados: PDF, PNG, JPG, JPEG, BMP, TIFF, GIF
- Extracción automática de:
  - Nombre de empresa/emisor
  - NIF/CIF/VAT
  - Número de factura
  - Fecha
  - Base imponible
  - % IVA/Tax
  - Importe IVA/Tax
  - IRPF (facturas españolas)
  - Total
- Generación de Excel profesional con totales y resumen

## Requisitos

- Python 3.9+
- Poppler (para procesar PDFs)
- **API Key de Anthropic** (Claude)

## Configuración de API Key

1. Obtén tu API key en [console.anthropic.com](https://console.anthropic.com)
2. Crea un archivo `.env` en la carpeta del proyecto:

```bash
cp .env.example .env
```

3. Edita `.env` y añade tu API key:

```
ANTHROPIC_API_KEY=sk-ant-api03-tu-api-key-aqui
```

## Instalación

### Opción 1: Docker (Recomendada)

```bash
cd invoice-extractor

# Copia y configura el archivo .env
cp .env.example .env
# Edita .env con tu API key de Anthropic

# Construir y ejecutar con Docker Compose
docker-compose up -d

# La aplicación estará disponible en http://localhost:5000
```

### Opción 2: Instalación Manual

#### 1. Instalar Poppler (para procesar PDFs)

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install poppler-utils
```

**macOS:**
```bash
brew install poppler
```

**Windows:**
- Descargar Poppler: https://github.com/oschwartz10612/poppler-windows/releases
- Añadir al PATH del sistema

#### 2. Instalar dependencias de Python

```bash
cd invoice-extractor
pip install -r requirements.txt
```

#### 3. Configurar API Key

```bash
cp .env.example .env
# Edita .env con tu API key de Anthropic
```

#### 4. Ejecutar la aplicación

```bash
python app.py
```

La aplicación estará disponible en `http://localhost:5000`

## Uso

1. Abre el navegador en `http://localhost:5000`
2. Arrastra tus facturas al área de carga o haz clic para seleccionarlas
3. Haz clic en "Procesar Facturas"
4. Revisa los resultados en la tabla
5. Descarga el Excel con el botón "Descargar Excel"

## Costes de API

El procesamiento usa Claude Sonnet con visión. Coste aproximado:
- **~$0.01-0.03 por factura** (dependiendo del tamaño del documento)
- PDFs de múltiples páginas costarán más

## Estructura del Proyecto

```
invoice-extractor/
├── app.py              # Aplicación Flask principal
├── extractor.py        # Módulo de extracción con Claude AI
├── excel_generator.py  # Generador de Excel
├── requirements.txt    # Dependencias Python
├── .env.example        # Ejemplo de configuración
├── Dockerfile          # Configuración Docker
├── docker-compose.yml  # Orquestación Docker
├── templates/
│   └── index.html      # Interfaz web
├── uploads/            # Archivos subidos (temporal)
└── outputs/            # Archivos Excel generados
```

## Variables de Entorno

| Variable | Descripción | Requerido |
|----------|-------------|-----------|
| `ANTHROPIC_API_KEY` | API key de Anthropic | **Sí** |
| `FLASK_ENV` | Modo de ejecución | No |
| `MAX_CONTENT_LENGTH` | Tamaño máximo archivo | No |

## API Endpoints

- `GET /` - Interfaz web principal
- `POST /upload` - Subir y procesar facturas
- `GET /download/<filename>` - Descargar Excel generado
- `GET /health` - Health check

## Ventajas vs OCR tradicional

| Característica | OCR + Regex | Claude AI |
|----------------|-------------|-----------|
| Precisión | Variable | Alta |
| Formatos nuevos | Requiere ajustes | Automático |
| Facturas escaneadas | Problemático | Funciona bien |
| Layouts complejos | Falla frecuente | Entiende contexto |
| Idiomas | Configuración manual | Multiidioma nativo |

## Licencia

MIT License

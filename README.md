# Procesador de CxP - Aplicación Web

Aplicación web Flask para procesar archivos CSV/Excel de Cuentas por Pagar (CxP).

## 🚀 Características

- **Carga de archivos**: Soporta CSV, XLS y XLSX
- **Procesamiento automático**: Limpia datos y organiza por proveedor y mes
- **Reporte Excel**: Genera archivo XLSX formateado con totales
- **Interfaz moderna**: Drag & Drop, responsive, fácil de usar

## 📁 Estructura del Proyecto

```
PYTHON_TEC/
├── app.py                  # Aplicación Flask principal
├── report_logic.py         # Lógica de procesamiento (tu código original)
├── main.py                 # Script local (opcional, no se usa en web)
├── requirements.txt        # Dependencias Python
├── .gitignore             # Archivos a ignorar en Git
├── README.md              # Este archivo
├── templates/
│   └── index.html         # Interfaz web
├── input/                 # Carpeta temporal para uploads
└── output/                # Carpeta temporal para archivos procesados
```

## 🔧 Instalación Local

1. **Clonar o descargar el proyecto**

2. **Crear entorno virtual** (recomendado):
```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
```

3. **Instalar dependencias**:
```bash
pip install -r requirements.txt
```

4. **Ejecutar la aplicación**:
```bash
python app.py
```

5. **Abrir en el navegador**:
```
http://localhost:5000
```

## 🌐 Desplegar en Render

### Opción 1: Desde GitHub (Recomendado)

1. **Sube tu código a GitHub**:
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
git push -u origin main
```

2. **En Render** (https://render.com):
   - Crea una cuenta o inicia sesión
   - Click en "New +" → "Web Service"
   - Conecta tu repositorio de GitHub
   - Configura:
     - **Name**: procesador-cxp (o el nombre que quieras)
     - **Environment**: Python 3
     - **Build Command**: `pip install -r requirements.txt`
     - **Start Command**: `gunicorn app:app`
   - Click en "Create Web Service"

3. **Espera el despliegue** (2-5 minutos)

4. **¡Listo!** Tu app estará disponible en: `https://procesador-cxp.onrender.com`

### Opción 2: Desde Git Local

Si no quieres usar GitHub:

1. **En Render**:
   - New + → Web Service
   - "Public Git repository"
   - Pega la URL de tu repo público

2. **Configuración igual que Opción 1**

## ⚙️ Variables de Entorno (Opcional)

En Render, puedes agregar en la sección "Environment":

```
SECRET_KEY=tu_clave_secreta_aqui
```

Esto mejora la seguridad de las sesiones de Flask.

## 📝 Uso de la Aplicación

1. Abre la aplicación en tu navegador
2. Arrastra tu archivo CSV/Excel o haz clic para seleccionarlo
3. Click en "Procesar Archivo"
4. El reporte se descargará automáticamente

## 🔒 Seguridad

- Límite de 16MB por archivo
- Solo acepta archivos CSV, XLS, XLSX
- Los archivos se eliminan después de procesarse
- Validación de extensiones en cliente y servidor

## 🐛 Troubleshooting

### Error: "Application failed to respond"
- Revisa que `gunicorn` esté en requirements.txt
- Verifica que el comando start sea: `gunicorn app:app`

### Error al procesar archivo
- Verifica que el CSV tenga las columnas: FECHAVTO, NOFAC, PROVEEDOR, SALDO
- Asegúrate que las fechas estén en formato DD/MM/YYYY

### La app está lenta en Render (plan gratuito)
- El plan gratuito de Render hiberna después de 15 min de inactividad
- El primer request después de hibernar puede tardar ~30 segundos

## 📦 Dependencias

- **Flask**: Framework web
- **pandas**: Procesamiento de datos
- **openpyxl**: Manipulación de archivos Excel
- **gunicorn**: Servidor WSGI para producción

## 🤝 Contribuciones

Si encuentras bugs o tienes sugerencias, siéntete libre de crear un issue o pull request.

## 📄 Licencia

Proyecto de uso interno/personal.

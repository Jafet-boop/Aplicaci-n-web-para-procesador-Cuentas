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

## 📄 Licencia

Proyecto de uso interno/personal.

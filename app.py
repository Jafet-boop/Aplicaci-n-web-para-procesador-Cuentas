from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from report_logic import run_report
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'tu_clave_secreta_aqui_cambiala'  # Cambiar en producción
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.config['UPLOAD_FOLDER'] = 'input'
app.config['OUTPUT_FOLDER'] = 'output'

ALLOWED_EXTENSIONS = {'dbf','csv', 'xlsx', 'xls'}

# Crear carpetas si no existen
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        try:
            # Guardar archivo subido
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename_with_timestamp = f"{timestamp}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_with_timestamp)
            file.save(filepath)
            
            # Procesar archivo
            today = datetime.today().strftime("%Y%m%d")
            base_name = "ReporteCXP"
            xlsx_name = f"{base_name}{today}.xlsx"
            
            run_report(filepath, base_name=base_name, output_dir=app.config['OUTPUT_FOLDER'])
            
            # Limpiar archivo de entrada después de procesar
            os.remove(filepath)
            
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], xlsx_name)
            
            return send_file(
                output_path,
                as_attachment=True,
                download_name=xlsx_name,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
        except Exception as e:
            flash(f'Error al procesar el archivo: {str(e)}', 'error')
            # Limpiar archivos en caso de error
            if os.path.exists(filepath):
                os.remove(filepath)
            return redirect(url_for('index'))
    else:
        flash('Tipo de archivo no permitido. Solo se aceptan DBF ,CSV, XLS y XLSX', 'error')
        return redirect(url_for('index'))

@app.route('/health')
def health():
    return {'status': 'ok'}, 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

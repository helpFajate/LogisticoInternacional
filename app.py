from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
from io import BytesIO
import os
from a_extr import generar_pdf_empaque # Importamos tu lógica de PDF

app = Flask(__name__)
app.secret_key = 'clave_secreta_vivell' # Necesario para mostrar mensajes de error (flash)

@app.route('/')
def index():
    return render_template('dashboard.html')

@app.route('/lista-empaque')
def lista_empaque():
    # Esta es la página donde está el formulario para generar el PDF
    return render_template('packing.html') 

@app.route('/procesar-packing', methods=['POST'])
def procesar_packing():
    # 1. Capturamos los datos enviados desde el HTML
    cia = request.form.get('cia')
    tipo = request.form.get('tipo_docto')
    consec_ini = request.form.get('consec_ini')
    consec_fin = request.form.get('consec_fin')

    # 2. Validamos que no falten datos
    if not all([cia, tipo, consec_ini, consec_fin]):
        error_msg = "Todos los campos son obligatorios"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'error': error_msg}), 400
        flash(error_msg)
        return redirect(url_for('lista_empaque'))

    # 3. Llamamos a tu función de backend (genera PDF en memoria)
    pdf_bytes, nombre_archivo = generar_pdf_empaque(cia, tipo, consec_ini, consec_fin)
    app.logger.info(f"Procesando packing: cia={cia}, tipo={tipo}, ini={consec_ini}, fin={consec_fin}")

    if pdf_bytes:
        # 4. Enviamos el PDF desde memoria al navegador del usuario
        pdf_stream = BytesIO(pdf_bytes)
        return send_file(pdf_stream, as_attachment=True, download_name=nombre_archivo, mimetype='application/pdf')
    else:
        error_msg = "Error: No se encontraron datos en SIESA o no fue posible generar el reporte."
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'error': error_msg}), 500
        flash(error_msg)
        return redirect(url_for('lista_empaque'))

if __name__ == '__main__':
    app.run(debug=True)
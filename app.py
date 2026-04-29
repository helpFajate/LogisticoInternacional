from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
from io import BytesIO
import os
from a_extr import generar_pdf_empaque, generar_excel_empaque

app = Flask(__name__)
app.secret_key = 'clave_secreta_vivell'

@app.route('/')
def index():
    return render_template('dashboard.html')

@app.route('/lista-empaque')
def lista_empaque():
    return render_template('packing.html')

@app.route('/procesar-packing', methods=['POST'])
def procesar_packing():
    cia        = request.form.get('cia')
    tipo       = request.form.get('tipo_docto')
    consec_ini = request.form.get('consec_ini')
    consec_fin = request.form.get('consec_fin')
    formato    = request.form.get('formato', 'pdf')   # 'pdf' o 'excel'

    if not all([cia, tipo, consec_ini, consec_fin]):
        error_msg = "Todos los campos son obligatorios"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'error': error_msg}), 400
        flash(error_msg)
        return redirect(url_for('lista_empaque'))

    app.logger.info(f"Procesando packing: cia={cia}, tipo={tipo}, ini={consec_ini}, fin={consec_fin}, fmt={formato}")

    if formato == 'excel':
        datos, nombre = generar_excel_empaque(cia, tipo, consec_ini, consec_fin)
        mimetype      = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    else:
        datos, nombre = generar_pdf_empaque(cia, tipo, consec_ini, consec_fin)
        mimetype      = 'application/pdf'

    if datos:
        return send_file(BytesIO(datos), as_attachment=True,
                         download_name=nombre, mimetype=mimetype)
    else:
        error_msg = "Error: No se encontraron datos en SIESA o no fue posible generar el reporte."
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'error': error_msg}), 500
        flash(error_msg)
        return redirect(url_for('lista_empaque'))

if __name__ == '__main__':
    app.run(debug=True)
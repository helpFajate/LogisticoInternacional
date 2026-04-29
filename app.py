from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
from io import BytesIO
import pandas as pd
import os
from a_extr import generar_pdf_empaque, generar_excel_empaque
from lista import procesar_etl_logica, insertar_en_sql_logica

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
    formato    = request.form.get('formato', 'pdf')

    if not all([cia, tipo, consec_ini, consec_fin]):
        error_msg = "Todos los campos son obligatorios"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'error': error_msg}), 400
        flash(error_msg)
        return redirect(url_for('lista_empaque'))

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
        error_msg = "Error: No se encontraron datos en SIESA."
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'error': error_msg}), 500
        flash(error_msg)
        return redirect(url_for('lista_empaque'))

@app.route('/sincronizar-precios')
def sincronizar_precios():
    return render_template('sincronizar.html')

@app.route('/ejecutar-sincronizacion', methods=['POST'])
def ejecutar_sincronizacion():
    listas = request.form.getlist('listas')
    
    if not listas:
        flash("Selecciona al menos una lista de precios.")
        return redirect(url_for('sincronizar_precios'))
    
    try:
        # 1. Ejecutar el proceso ETL
        df = procesar_etl_logica(listas)
        
        if df.empty:
            flash("No se encontraron datos para las listas seleccionadas.")
            return redirect(url_for('sincronizar_precios'))

        # 2. Guardar en SQL automáticamente
        insertar_en_sql_logica(df)
        
        # 3. Preparar el Excel para descarga inmediata
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Precios')
        output.seek(0)
        
        # Enviamos el archivo (esto es lo que el navegador recibirá)
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Reporte_Precios_Siesa.xlsx"
        )

    except Exception as e:
        flash(f"Ocurrió un error: {str(e)}")
        return redirect(url_for('sincronizar_precios'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
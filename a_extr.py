import pyodbc
import pandas as pd
from fpdf import FPDF
import warnings
import os

# Ocultar advertencias de Pandas
warnings.filterwarnings('ignore', category=UserWarning)

# ---------------------------------------------------------
# CLASE PDF PERSONALIZADA CON ENCABEZADO Y LOGOS
# ---------------------------------------------------------
class ReporteEmpaque(FPDF):
    def __init__(self, datos_encabezado, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.datos_envio = datos_encabezado
        self.ruta_base_img = r"C:\Users\mesa_ayuda\Desktop\Internacional\imagenes"

    def header(self):
        if self.page_no() == 1:
            y_inicial = 10
            alto_celda_cuadro = 6
            alto_total_cuadro = alto_celda_cuadro * 3
            alto_imagenes = alto_total_cuadro * 0.95
            
            # --- 1. CUADRO PACKING LIST ---
            x_cuadro = 130
            ancho_cuadro = 70
            self.set_y(y_inicial)
            self.set_x(x_cuadro)
            self.set_font("helvetica", "", 10)
            self.set_text_color(0, 0, 0)
            self.cell(ancho_cuadro, alto_celda_cuadro, "PACKING LIST/LISTA DE EMPAQUE", border=1, align="C", ln=1)
            self.set_x(x_cuadro)
            factura = self.datos_envio.get('factura', 'N/A')
            self.cell(ancho_cuadro, alto_celda_cuadro, f"Fact. N°: {factura}", border=1, align="C", ln=1)
            self.set_x(x_cuadro)
            self.cell(ancho_cuadro, alto_celda_cuadro, f"Pagina N°: {self.page_no()}", border=1, align="C", ln=1)

            # --- 2. IMÁGENES ---
            y_img = y_inicial + 0.45 
            img_vivell = os.path.join(self.ruta_base_img, "vivell.png")
            img_invima = os.path.join(self.ruta_base_img, "INVIMA.png")
            img_operador = os.path.join(self.ruta_base_img, "OPERADOR.png")

            if os.path.exists(img_vivell): self.image(img_vivell, x=10, y=y_img, h=alto_imagenes)
            if os.path.exists(img_invima): self.image(img_invima, x=60, y=y_img + 3, w=30)
            if os.path.exists(img_operador): self.image(img_operador, x=95, y=y_img + 3, w=25)

            self.ln(12)

            # --- 3. TABLA DE CLIENTE Y VENDEDOR ---
            self.set_font("helvetica", "", 8)
            w_cols = [35, 60, 35, 60] 
            alto_fila = 6 
            
            self.cell(w_cols[0] + w_cols[1], 5, "Customer/Cliente", border=1, align="C")
            self.cell(w_cols[2] + w_cols[3], 5, "Vendor/Vendedor", border=1, align="C", ln=1)
            
            def fila_datos(label1, val1, label2, val2):
                curr_x, curr_y = self.get_x(), self.get_y()
                self.cell(w_cols[0], alto_fila, "", border=1)
                self.cell(w_cols[1], alto_fila, "", border=1)
                self.cell(w_cols[2], alto_fila, "", border=1)
                self.cell(w_cols[3], alto_fila, "", border=1, ln=1)
                self.set_xy(curr_x, curr_y + 0.5)
                self.multi_cell(w_cols[0], 3, label1, align="L")
                self.set_xy(curr_x + w_cols[0], curr_y + 0.5)
                self.multi_cell(w_cols[1], 3, str(val1), align="L")
                self.set_xy(curr_x + w_cols[0] + w_cols[1], curr_y + 0.5)
                self.multi_cell(w_cols[2], 3, label2, align="L")
                self.set_xy(curr_x + w_cols[0] + w_cols[1] + w_cols[2], curr_y + 0.5)
                self.multi_cell(w_cols[3], 3, str(val2), align="L")
                self.set_xy(curr_x, curr_y + alto_fila)

            d = self.datos_envio
            fila_datos("Name/Nombre", d.get('cliente_nombre',''), "Name/Nombre", d.get('cia_nombre',''))
            fila_datos("ID/Nit", d.get('cliente_nit',''), "ID/NIT", d.get('cia_nit',''))
            fila_datos("Address/Dirección", d.get('cliente_dir',''), "Address/Dirección", d.get('cia_dir',''))
            fila_datos("City/Ciudad", d.get('cliente_ciu',''), "City/Ciudad", d.get('cia_ciu',''))
            fila_datos("State/Depto", d.get('cliente_dep',''), "State/Depto", d.get('cia_dep',''))
            fila_datos("Country/País", d.get('cliente_pais',''), "Country/País", d.get('cia_pais',''))
            fila_datos("Phone/Teléfono", d.get('cliente_tel',''), "Phone/Teléfono", d.get('cia_tel',''))
            fila_datos("Contact/Contacto", d.get('cliente_cont',''), "Vendor/Vendedor", d.get('vendedor',''))
            self.ln(5)
        else:
            self.set_y(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("helvetica", "I", 8)
        self.set_text_color(0, 0, 0)
        self.cell(0, 10, f"Página {self.page_no()}/{{nb}}", align="C")

    def check_page_break(self, h):
        if self.get_y() + h > self.page_break_trigger:
            self.add_page()

# ---------------------------------------------------------
# FUNCIÓN PRINCIPAL
# ---------------------------------------------------------
def generar_pdf_empaque(cia, tipo_docto, consec_inicial, consec_final):
    conn_str = (r'DRIVER={ODBC Driver 17 for SQL Server};SERVER=AMATERASU\SIESA;DATABASE=Reportes;Trusted_Connection=yes;')
    try:
        conn = pyodbc.connect(conn_str)
        query = "EXEC sp_ti_listaEmpaque @v_cia=?, @v_idco='063', @v_idtipodocto=?, @v_consecinicial=?, @v_consecfinal=?"
        df = pd.read_sql(query, conn, params=(cia, tipo_docto, consec_inicial, consec_final))
        conn.close()
    except Exception as e:
        print(f"Error conexión: {e}")
        return None, None

    if df.empty:
        return None, None

    f = df.iloc[0]
    datos_h = {
        'factura': f"{f['f_prefijo']}-{f['f_consec_docto']}",
        'cliente_nombre': f['f_cliente_razon_soc'], 'cliente_nit': f['f_cliente_nit'],
        'cliente_dir': f['f_direccion1_cliente'], 'cliente_ciu': f['f_ciudad_cliente'],
        'cliente_dep': f['f_depto_cliente'], 'cliente_pais': f['f_pais_cliente'],
        'cliente_tel': f['f_telefono_cliente'], 'cliente_cont': f['f_contacto_cliente'],
        'cia_nombre': f['f_razon_social_cia'], 'cia_nit': f['f_nit_cia'],
        'cia_dir': f['f_direccion1_co'], 'cia_ciu': f['f_ciudad_co'],
        'cia_dep': f['f_depto_co'], 'cia_pais': f['f_pais_co'],
        'cia_tel': f['f_telefono_co'], 'vendedor': f['f_vendedor_razon_social']
    }

    w = [18, 12, 12, 64, 64, 20] 

    pdf = ReporteEmpaque(datos_encabezado=datos_h)
    pdf.alias_nb_pages()
    pdf.add_page()

    total_gral_cant = 0
    total_gral_pb = 0.0
    grupos = df.groupby('Caja', sort=True)

    for caja, items in grupos:
        val_pb = float(items['f2_ent_peso_bruto'].iloc[0]) if pd.notna(items['f2_ent_peso_bruto'].iloc[0]) else 0.0
        val_pn = float(items['f2_ent_peso_neto'].iloc[0]) if pd.notna(items['f2_ent_peso_neto'].iloc[0]) else 0.0
        
        pdf.check_page_break(20)
        pdf.set_font("helvetica", "B", 9)
        pdf.set_fill_color(240, 240, 240)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(sum(w), 8, f" Box/Caja N°: {caja}  - Remision/Remission: {items['f2_remision'].iloc[0]}", border=1, fill=True, ln=1)

        # --- CABECERA DE COLUMNAS CON AUTOAJUSTE MULTI-CELL ---
        pdf.set_font("helvetica", "B", 8)
        pdf.set_fill_color(220, 220, 220)
        titulos = ["Reference", "Size/Talla", "Colour/Color", "Description", "Composition", "Qty/Cant"]
        
        h_encabezado = 8 # Alto fijo de la fila de títulos
        curr_x, curr_y = pdf.get_x(), pdf.get_y()
        
        for i, t in enumerate(titulos):
            # 1. Dibujamos el cuadro con fondo gris y borde
            pdf.set_xy(curr_x, curr_y)
            pdf.cell(w[i], h_encabezado, "", border=1, fill=True)
            
            # 2. Si la columna es estrecha (<= 20mm), cambiamos '/' por '/\n' para forzar 2 líneas
            texto_titulo = t.replace("/", "/\n") if w[i] <= 20 else t
            
            # 3. Ajustamos el margen Y para que quede centrado verticalmente
            y_margen = curr_y + 1 if '\n' in texto_titulo else curr_y + 2.5
            
            # 4. Imprimimos el texto
            pdf.set_xy(curr_x, y_margen)
            pdf.multi_cell(w[i], 3, texto_titulo, align="C")
            
            curr_x += w[i]
            
        pdf.set_xy(10, curr_y + h_encabezado) # Regresamos el cursor abajo de la fila
        # ------------------------------------------------------

        pdf.set_font("helvetica", "", 7)
        cant_caja = 0.0
        for _, row in items.iterrows():
            txt_desc, txt_comp = str(row['f2_desc_crit1']), str(row['f2_desc_crit2'])
            
            l_desc = (pdf.get_string_width(txt_desc) // w[3]) + 1
            l_comp = (pdf.get_string_width(txt_comp) // w[4]) + 1
            h_fila = max(7, max(l_desc, l_comp) * 4)

            pdf.check_page_break(h_fila)
            curr_y = pdf.get_y()
            
            pdf.cell(w[0], h_fila, str(row['f2_referencia']), border=1, align="C")
            pdf.cell(w[1], h_fila, str(row['f2_ext2_det']), border=1, align="C")
            pdf.cell(w[2], h_fila, str(row['f2_id_ext1_det']), border=1, align="C")
            
            x_desc = pdf.get_x()
            pdf.multi_cell(w[3], h_fila/l_desc if l_desc > 0 else h_fila, txt_desc, border=1)
            pdf.set_xy(x_desc + w[3], curr_y)
            pdf.multi_cell(w[4], h_fila/l_comp if l_comp > 0 else h_fila, txt_comp, border=1)
            
            pdf.set_xy(x_desc + w[3] + w[4], curr_y)
            v_can = float(row['f2_cantidad'])
            pdf.cell(w[5], h_fila, f"{v_can:.0f}", border=1, align="C", ln=1)
            cant_caja += v_can

        pdf.set_font("helvetica", "B", 8)
        pdf.cell(w[0]+w[1]+w[2], 7, f"Total BOX/CAJA: {caja}", border="LTB")
        pdf.cell(w[3], 7, f"Gross Weight/Peso Bruto: {val_pb:.2f}", border="TB")
        pdf.cell(w[4], 7, f"Net Weight/Peso Neto: {val_pn:.2f}", border="TB", align="R")
        pdf.cell(w[5], 7, f"{cant_caja:.0f}", border="RTB", align="C", ln=1)
        pdf.ln(4)

        total_gral_cant += cant_caja
        total_gral_pb += val_pb

    pdf.check_page_break(15)
    pdf.set_font("helvetica", "B", 10)
    pdf.set_fill_color(200, 200, 200)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(sum(w[:4]), 10, "  TOTAL GENERAL", border=1, fill=True)
    pdf.cell(sum(w[4:]), 10, f"{total_gral_cant:.0f} UNIDADES / {total_gral_pb:.2f} KG", border=1, fill=True, align="C")

    # Generar PDF en memoria (sin guardar a disco)
    try:
        pdf_bytes = pdf.output()
        nombre_archivo = f"Packing_List_{consec_inicial}.pdf"
        print(f"Éxito: PDF generado en memoria ({len(pdf_bytes)} bytes)")
        return pdf_bytes, nombre_archivo
    except Exception as e:
        print(f"Error al generar PDF: {e}")
        return None, None

if __name__ == "__main__":
    pdf_bytes, nombre_archivo = generar_pdf_empaque(1, 'FEE', '1412', '1412')
    if pdf_bytes:
        print(f"PDF generado exitosamente: {nombre_archivo} ({len(pdf_bytes)} bytes)")
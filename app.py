import os
import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
from openpyxl import Workbook
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ParseError
from urllib.parse import quote
import subprocess
import platform

# Namespaces para los XML CFDI
NAMESPACES = {
    'cfdi': 'http://www.sat.gob.mx/cfd/4',
    'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
}

def seleccionar_archivos():
    global archivos_seleccionados
    archivos_seleccionados = filedialog.askopenfilenames(
        filetypes=[("Archivos XML", "*.xml"), ("Todos los archivos", "*.*")],
        title="Seleccionar archivos XML (CFDI)"
    )
    if archivos_seleccionados:
        label_archivos.config(text=f"{len(archivos_seleccionados)} archivos seleccionados")
        boton_procesar.config(state=tk.NORMAL)

def generar_url_verificacion(uuid, rfc_emisor, rfc_receptor, total, sello):
    """Genera URL de verificación compatible exactamente con el formato del QR del SAT"""
    if uuid == 'N/A' or not sello or len(sello) < 8:
        return 'N/A'
    
    try:
        # Formatear total como lo hace el SAT (6 decimales con ceros)
        formatted_total = f"{float(total):016.6f}"
        
        # Extraer los últimos 8 caracteres del sello sin codificar
        fe = sello[-8:]
        
        # Parámetros en el orden y formato exacto que usa el SAT
        params = [
            f"id={uuid.strip()}",
            f"re={rfc_emisor.strip()}",
            f"rr={rfc_receptor.strip()}",
            f"tt={formatted_total}",
            f"fe={fe}"
        ]
        
        return f"https://verificacfdi.facturaelectronica.sat.gob.mx/default.aspx?{'&'.join(params)}"
    
    except Exception as e:
        print(f"Error generando URL: {str(e)}")
        return 'N/A'

def abrir_archivo(ruta):
    """Abre un archivo usando el comando apropiado para cada sistema operativo"""
    try:
        if platform.system() == 'Windows':
            os.startfile(ruta)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open', ruta])
        else:  # Linux
            subprocess.call(['xdg-open', ruta])
    except Exception as e:
        messagebox.showwarning("Advertencia", f"No se pudo abrir el archivo:\n{str(e)}")

def procesar_facturas():
    if not archivos_seleccionados:
        messagebox.showwarning("Advertencia", "No hay archivos seleccionados")
        return
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Facturas"
    
    # Configurar anchos de columna
    column_widths = {
        'A': 15, 'B': 15, 'C': 20, 'D': 15, 'E': 15, 'F': 15,
        'G': 12, 'H': 12, 'I': 12, 'J': 10, 'K': 12, 'L': 15,
        'M': 30, 'N': 15, 'O': 15, 'P': 30, 'Q': 15, 'R': 15,
        'S': 10, 'T': 40, 'U': 20, 'V': 70, 'W': 40, 'X': 15,
        'Y': 12, 'Z': 12
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    headers = [
        "Serie", "Folio", "Fecha", "TipoComprobante", "FormaPago", "MetodoPago",
        "SubTotal", "Descuento", "Total", "Moneda", "TipoCambio", 
        "EmisorRFC", "EmisorNombre", "EmisorRegimen",
        "ReceptorRFC", "ReceptorNombre", "ReceptorUsoCFDI", "ReceptorRegimen", "ReceptorCP",
        "UUID", "FechaTimbrado", "URLVerificacion",
        "Conceptos", "ImpuestosTrasladados", "ContieneCafe", "ContieneCerveza"
    ]
    ws.append(headers)
    
    # Formato para números
    number_format = '0.00'
    for col in ['G', 'H', 'I', 'X']:
        for cell in ws[col]:
            cell.number_format = number_format

    for archivo in archivos_seleccionados:
        try:
            tree = ET.parse(archivo)
            root = tree.getroot()

            # Datos generales del comprobante
            comprobante = root.attrib
            serie = comprobante.get('Serie', 'N/A')
            folio = comprobante.get('Folio', 'N/A')
            fecha = comprobante.get('Fecha', 'N/A')
            tipo_comprobante = comprobante.get('TipoDeComprobante', 'N/A')
            forma_pago = comprobante.get('FormaPago', 'N/A')
            metodo_pago = comprobante.get('MetodoPago', 'N/A')
            sub_total = comprobante.get('SubTotal', '0')
            descuento = comprobante.get('Descuento', '0')
            total = comprobante.get('Total', '0')
            moneda = comprobante.get('Moneda', 'N/A')
            tipo_cambio = comprobante.get('TipoCambio', '1')

            # Datos del emisor
            emisor = root.find('cfdi:Emisor', NAMESPACES)
            emisor_rfc = emisor.get('Rfc', 'N/A') if emisor is not None else 'N/A'
            emisor_nombre = emisor.get('Nombre', 'N/A') if emisor is not None else 'N/A'
            emisor_regimen = emisor.get('RegimenFiscal', 'N/A') if emisor is not None else 'N/A'

            # Datos del receptor
            receptor = root.find('cfdi:Receptor', NAMESPACES)
            receptor_rfc = receptor.get('Rfc', 'N/A') if receptor is not None else 'N/A'
            receptor_nombre = receptor.get('Nombre', 'N/A') if receptor is not None else 'N/A'
            receptor_uso = receptor.get('UsoCFDI', 'N/A') if receptor is not None else 'N/A'
            receptor_regimen = receptor.get('RegimenFiscalReceptor', 'N/A') if receptor is not None else 'N/A'
            receptor_cp = receptor.get('DomicilioFiscalReceptor', 'N/A') if receptor is not None else 'N/A'

            # Timbre fiscal (UUID)
            complemento = root.find('cfdi:Complemento', NAMESPACES)
            timbre = complemento.find('tfd:TimbreFiscalDigital', NAMESPACES) if complemento is not None else None
            uuid = timbre.get('UUID', 'N/A') if timbre is not None else 'N/A'
            fecha_timbrado = timbre.get('FechaTimbrado', 'N/A') if timbre is not None else 'N/A'
            sello_cfd = timbre.get('SelloCFD', '') if timbre is not None else ''

            # Impuestos
            impuestos = root.find('cfdi:Impuestos', NAMESPACES)
            total_impuestos = impuestos.get('TotalImpuestosTrasladados', '0') if impuestos is not None else '0'

            # URL de verificación con todos los parámetros necesarios
            url_verificacion = generar_url_verificacion(
                uuid, emisor_rfc, receptor_rfc, total, sello_cfd
            )

            # Conceptos
            conceptos = root.find('cfdi:Conceptos', NAMESPACES)
            descripciones = []
            contiene_cafe = False
            contiene_cerveza = False
            
            if conceptos is not None:
                for concepto in conceptos.findall('cfdi:Concepto', NAMESPACES):
                    descripcion = concepto.get('Descripcion', '')
                    descripciones.append(descripcion)
                    
                    # Buscar productos
                    desc_lower = descripcion.lower()
                    if 'café' in desc_lower or 'coffee' in desc_lower:
                        contiene_cafe = True
                    if 'cerveza' in desc_lower or 'beer' in desc_lower:
                        contiene_cerveza = True
            
            conceptos_str = ", ".join(descripciones) if descripciones else "N/A"

            # Construir fila para Excel
            fila = [
                serie, folio, fecha, tipo_comprobante, forma_pago, metodo_pago,
                float(sub_total), float(descuento), float(total), moneda, tipo_cambio,
                emisor_rfc, emisor_nombre, emisor_regimen,
                receptor_rfc, receptor_nombre, receptor_uso, receptor_regimen, receptor_cp,
                uuid, fecha_timbrado, url_verificacion,
                conceptos_str, float(total_impuestos),
                "Sí" if contiene_cafe else "No",
                "Sí" if contiene_cerveza else "No"
            ]
            ws.append(fila)

            # Abrir URL de verificación (opcional)
            if url_verificacion != 'N/A' and chk_abrir_web.get():
                webbrowser.open_new_tab(url_verificacion)

        except ParseError as e:
            messagebox.showerror("Error", f"El archivo {os.path.basename(archivo)} no es un XML válido.\nError: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar {os.path.basename(archivo)}\nError: {str(e)}")

    # Guardar Excel
    if len(archivos_seleccionados) > 0:
        ruta_salida = os.path.join(os.path.dirname(archivos_seleccionados[0]), "facturas.xlsx")
        try:
            wb.save(ruta_salida)
            messagebox.showinfo(
                "Éxito", 
                f"Facturas procesadas correctamente.\n\n"
                f"Archivo guardado en:\n{ruta_salida}\n\n"
                f"Total de facturas procesadas: {len(archivos_seleccionados)}"
            )
            
            # Preguntar si desea abrir el archivo Excel generado
            if messagebox.askyesno("Abrir archivo", "¿Desea abrir el archivo Excel generado?"):
                abrir_archivo(ruta_salida)
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo Excel:\n{str(e)}")

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Procesador de Facturas XML (CFDI) v2024")
root.geometry("600x350")

# Estilo
font_style = ('Arial', 10)
button_style = {'font': font_style, 'bg': '#4CAF50', 'fg': 'white', 'padx': 10, 'pady': 5}

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill=tk.BOTH)

label_titulo = tk.Label(
    frame, 
    text="Procesador de Facturas XML (CFDI)",
    font=('Arial', 12, 'bold'),
    pady=10
)
label_titulo.pack()

# Frame para botones
frame_botones = tk.Frame(frame)
frame_botones.pack(pady=10)

boton_seleccionar = tk.Button(
    frame_botones, 
    text="Seleccionar Facturas XML", 
    command=seleccionar_archivos,
    **button_style
)
boton_seleccionar.pack(side=tk.LEFT, padx=5)

boton_procesar = tk.Button(
    frame_botones, 
    text="Procesar Facturas", 
    command=procesar_facturas, 
    state=tk.DISABLED,
    **button_style
)
boton_procesar.pack(side=tk.LEFT, padx=5)

label_archivos = tk.Label(
    frame, 
    text="No hay archivos seleccionados",
    font=font_style,
    fg='gray'
)
label_archivos.pack(pady=5)

# Checkbox para abrir páginas web
chk_abrir_web = tk.BooleanVar(value=True)
checkbox_web = tk.Checkbutton(
    frame,
    text="Abrir páginas de verificación del SAT automáticamente",
    variable=chk_abrir_web,
    font=font_style
)
checkbox_web.pack(pady=5)

# Footer
label_footer = tk.Label(
    frame, 
    text="Sistema compatible con CFDI 4.0 | Versión 2024",
    font=('Arial', 8),
    fg='gray'
)
label_footer.pack(side=tk.BOTTOM, pady=5)

root.mainloop()
import streamlit as st
import openpyxl
from collections import Counter
import base64

# Función para procesar un archivo
def process_file(selected_file, ws, ordenes_procesadas):
    # Variables fijas
    sender_name = "Crystal"
    sender_phone = "57 318 522 9083"
    pickup_address = "Carrera 48 # 52 sur - 81, Sabaneta"
    delivery_type = "sdd"
    corp_client_id = "2d542aad62f54d6c98baec419c5ecdc1"
    return_address = "Carrera 48 # 52 sur - 81, Sabaneta"
    cargo_options = ""
    item_cost = "0"
    quantity = "1"
    sender_comment = "Sin Comentario"

    order_no = recipient_name = recipient_phone = delivery_address = recipient_comment = purchase_order = ""

    # Leer el contenido del archivo binario
    file_content = selected_file.read().decode("utf-8").splitlines()

    for line in file_content:
        parts = line.strip().split('|')

        if parts[0] == "ED" and parts[1] == "RE":
            order_no = parts[2]
            ordenes_procesadas.append(order_no)
        elif parts[0] == "DG" and parts[1] == "57":
            recipient_name = parts[4]
            recipient_phone = parts[5]
            delivery_address = parts[6]
            recipient_comment = delivery_address
        elif parts[0] == "VA":
            # cash_on_delivery y auto_accept ahora siempre están vacíos
            cash_on_delivery = ""
            auto_accept = ""
            if order_no and recipient_name and recipient_phone and delivery_address:
                recipient_comment += f" Orden de compra numero: {purchase_order}"
                row = [order_no, sender_name, sender_phone, pickup_address, recipient_name, recipient_phone,
                       delivery_address, delivery_type, '', '', '', cash_on_delivery, auto_accept,
                       sender_comment, recipient_comment, corp_client_id, return_address, cargo_options,
                       item_cost, quantity]
                ws.append(row)
                order_no = recipient_name = recipient_phone = delivery_address = recipient_comment = purchase_order = ""

def main():
    st.title("File Selector")

    # Widget para cargar archivos
    selected_files = st.file_uploader("Seleccionar Archivos", type=["IFT"], accept_multiple_files=True)

    if selected_files:
        # Botón para procesar archivos
        if st.button("Procesar"):
            # Crear un nuevo libro de Excel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Datos Consolidados"
            
            # Función para agregar encabezados
            def agregar_encabezados(ws):
                headers = ["Orden no.", "Nombre del remitente", "Teléfono del remitente", "Dirección de recogida",
                           "Nombre del destinatario", "Teléfono del destinatario", "Dirección de entrega",
                           "Tipo de entrega", "cargo_inn", "article", "vat_code", "cash_on_delivery", "auto_accept",
                           "Comentario del remitente", "Comentario para el destinatario",
                           "corp_client_id", "return_address", "cargo_options", "Costo del artículo", "Cantidad"]
                ws.append(headers)

            agregar_encabezados(ws)

            # Para almacenar los números de orden
            ordenes_procesadas = []

            for selected_file in selected_files:
                # Procesar cada archivo seleccionado y agregar sus datos a la hoja de Excel
                process_file(selected_file, ws, ordenes_procesadas)

            # Buscar números de orden duplicados
            duplicados = [num for num, count in Counter(ordenes_procesadas).items() if count > 1]

            # Guardar el archivo de Excel consolidado
            output_excel = "output_consolidado.xlsx"
            wb.save(output_excel)

            # Mensaje final con información sobre duplicados
            if duplicados:
                mensaje = f"Archivo Excel '{output_excel}' creado con éxito. \nNúmeros de orden duplicados: {', '.join(duplicados)}"
            else:
                mensaje = f"Archivo Excel '{output_excel}' creado con éxito sin duplicados."
            st.info(mensaje)

            # Leer el archivo binario y codificarlo en base64
            with open(output_excel, "rb") as f:
                bytes_data = f.read()
                base64_data = base64.b64encode(bytes_data).decode()

            # Generar el enlace de descarga
            href = f'<a href="data:application/octet-stream;base64,{base64_data}" download="{output_excel}">Descargar {output_excel}</a>'
            st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

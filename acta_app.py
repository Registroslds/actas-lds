import streamlit as st
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pandas as pd
from datetime import datetime

# Configuración de Email (EDITA ESTO)
EMAIL_FROM = "tu_correo@gmail.com"  # Tu email
EMAIL_PASSWORD = "tu_app_password"  # Contraseña de app de Gmail
EMAIL_LIST = ["destinatario1@empresa.com", "destinatario2@empresa.com"]  # Lista de receptores
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

def generate_pdf(data_general, data_part, data_acuerdos):
    """Genera PDF con estructura del acta."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    
    # Título
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=16, spaceAfter=20, alignment=1)
    story.append(Paragraph("Acta de Reunión", title_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Información General (como tabla)
    general_data = [[k, v] for k, v in data_general.items()]
    general_table = Table(general_data, colWidths=[2*inch, 4*inch])
    general_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(general_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Participantes
    part_headers = [['#', 'Nombre y Apellido', 'Empresa', 'Rol']]
    part_data = [[i+1] + row for i, row in enumerate(zip(data_part['nombre'], data_part['empresa'], data_part['rol']))]
    part_table = Table(part_headers + part_data, colWidths=[0.5*inch, 2.5*inch, 1.5*inch, 2*inch])
    part_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(Paragraph("2. Participantes", styles['Heading2']))
    story.append(part_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Acuerdos
    acuerdos_headers = [['#', 'Acuerdo', 'Responsable', 'Fecha Inicio', 'Fecha Final', 'Avance']]
    acuerdos_data = [[i+1] + row for i, row in enumerate(zip(data_acuerdos['acuerdo'], data_acuerdos['responsable'], 
                                                           data_acuerdos['fecha_inicio'], data_acuerdos['fecha_final'], 
                                                           data_acuerdos['avance']))]
    acuerdos_table = Table(acuerdos_headers + acuerdos_data, colWidths=[0.5*inch, 3*inch, 1.5*inch, 1*inch, 1*inch, 1*inch])
    acuerdos_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
    ]))
    story.append(Paragraph("3. Acuerdos", styles['Heading2']))
    story.append(acuerdos_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def send_email(pdf_bytes, subject="Acta de Reunión Generada"):
    """Envía PDF por email a la lista."""
    msg = MIMEMultipart()
    msg['From'] = EMAIL_FROM
    msg['To'] = ", ".join(EMAIL_LIST)
    msg['Subject'] = subject
    
    body = f"Adjunto el PDF de la acta de reunión generada el {datetime.now().strftime('%d/%m/%Y %H:%M')}."
    msg.attach(MIMEText(body, 'plain'))
    
    # Adjuntar PDF
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= Acta_Reunion_{datetime.now().strftime('%d%m%Y')}.pdf")
    msg.attach(part)
    
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        text = msg.as_string()
        server.sendmail(EMAIL_FROM, EMAIL_LIST, text)
        server.quit()
        return True, "Email enviado exitosamente."
    except Exception as e:
        return False, f"Error enviando email: {str(e)}"

# Interfaz Streamlit
st.title("🗂️ Generador de Actas de Reunión")
st.markdown("Ingresa los datos y genera el PDF automáticamente. Se enviará a la lista configurada.")

# Formulario General
with st.form("general_form"):
    st.subheader("1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        id_acta = st.text_input("ID")
        num_acta = st.text_input("Número de Acta")
        fecha = st.date_input("Fecha de la Reunión", value=datetime(2025, 10, 9))
        hora_inicio = st.time_input("Hora de Inicio", value=datetime.strptime("16:30", "%H:%M").time())
    with col2:
        hora_fin = st.time_input("Hora de Finalización", value=datetime.strptime("17:30", "%H:%M").time())
        lugar = st.text_input("Lugar", value="Reunión virtual Teams")
        cliente = st.text_input("Cliente", value="Luz Del Sur")
        proyecto = st.text_input("Nombre del Proyecto", value="Servicio de Mesa de Ayuda")
        objetivo = st.text_area("Objetivo", value="Informe Semanal de Mesa de Ayuda")
    
    submitted_general = st.form_submit_button("Siguiente: Participantes")

# Participantes Dinámicos
if submitted_general:
    data_part = {"nombre": [], "empresa": [], "rol": []}
    st.subheader("2. Participantes")
    num_part = st.number_input("Número de Participantes", min_value=1, value=3)
    for i in range(num_part):
        with st.expander(f"Participante {i+1}"):
            nombre = st.text_input(f"Nombre {i+1}", value=["Richard Perez", "Leonardo Valdizan", "Katherine Navarrete"][i] if i < 3 else "")
            empresa = st.text_input(f"Empresa {i+1}", value=["LDS", "Stefanini", "Stefanini"][i] if i < 3 else "")
            rol = st.text_input(f"Rol {i+1}", value=["Soporte TI", "Help Desk Lead", "Delivery Lead"][i] if i < 3 else "")
            if nombre:  # Solo agregar si hay nombre
                data_part["nombre"].append(nombre)
                data_part["empresa"].append(empresa)
                data_part["rol"].append(rol)
    
    # Acuerdos Dinámicos
    st.subheader("3. Acuerdos")
    num_acuerdos = st.number_input("Número de Acuerdos", min_value=1, value=5)
    data_acuerdos = {"acuerdo": [], "responsable": [], "fecha_inicio": [], "fecha_final": [], "avance": []}
    for i in range(num_acuerdos):
        with st.expander(f"Acuerdo {i+1}"):
            acuerdo = st.text_area(f"Descripción {i+1}", value=[
                "Agregar categorías en el catálogo de servicios (Incidentes).",
                "Agregar plantilla en la descripción de los tickets (Incidentes y Solicitudes).",
                "Informar al equipo MDA ante solicitudes de creación de PST no se procederá con dicha atención.",
                "Informar al equipo MDA sobre agregar en la descripción de los tickets, si la atención es para un usuario de Inland",
                "Revisar plan de renovación de equipos"
            ][i] if i < 5 else "")
            resp = st.text_input(f"Responsable {i+1}", value=[
                "Stefanini / Leonardo Valdizan / Carlos Rivera",
                "Stefanini / Leonardo Valdizan / Carlos Rivera",
                "Stefanini / Leonardo Valdizan / Carlos Rivera",
                "Stefanini / Leonardo Valdizan",
                "Stefanini / Leonardo Valdizan – LDS / Richard Perez"
            ][i] if i < 5 else "")
            f_inicio = st.date_input(f"Fecha Inicio {i+1}", value=datetime(2025, 6, 19))
            f_final = st.date_input(f"Fecha Final {i+1}", value=datetime(2025, 6, 26))
            avance = st.selectbox(f"Avance {i+1}", ["0%", "25%", "50%", "75%", "100%"], index=4)
            if acuerdo:  # Solo agregar si hay acuerdo
                data_acuerdos["acuerdo"].append(acuerdo)
                data_acuerdos["responsable"].append(resp)
                data_acuerdos["fecha_inicio"].append(f_inicio.strftime("%d/%m/%y"))
                data_acuerdos["fecha_final"].append(f_final.strftime("%d/%m/%y"))
                data_acuerdos["avance"].append(avance)
    
    # Datos General consolidados
    data_general = {
        "ID": id_acta,
        "Número de acta": num_acta,
        "Fecha de la reunión": fecha.strftime("%d/%m/%Y"),
        "Hora de Inicio": str(hora_inicio),
        "Hora de finalización": str(hora_fin),
        "Lugar de la reunión": lugar,
        "Cliente": cliente,
        "Nombre del proyecto": proyecto,
        "Objetivo de la reunión": objetivo
    }
    
    submitted = st.form_submit_button("Generar PDF y Enviar por Email")
    
    if submitted:
        with st.spinner("Generando PDF..."):
            pdf_bytes = generate_pdf(data_general, data_part, data_acuerdos)
        
        st.success("¡PDF generado!")
        st.download_button("Descargar PDF", pdf_bytes, "Acta_Reunion.pdf", "application/pdf")
        
        if st.button("Enviar por Email"):
            success, msg = send_email(pdf_bytes)
            if success:
                st.success(msg)
            else:
                st.error(msg)

if __name__ == "__main__":
    pass  # Streamlit lo maneja

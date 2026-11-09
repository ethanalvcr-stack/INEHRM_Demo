import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def generar_pdf_tramite(folio, area_solicitante, responsable, recurso_pedido, fecha_limite, tipo_gestion="Recursos Materiales"):
    # Ruta raíz del proyecto (sube un nivel desde 'modulos')
    dir_modulos = os.path.dirname(os.path.abspath(__file__))
    dir_raiz = os.path.dirname(dir_modulos)
    reportes_dir = os.path.join(dir_raiz, 'reportes_pdf')
    os.makedirs(reportes_dir, exist_ok=True)
    
    nombre_archivo = f"Acuse_{folio.replace('/', '_')}.pdf"
    ruta_pdf = os.path.join(reportes_dir, nombre_archivo)
    
    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=letter,
        rightMargin=36, leftMargin=36,
        topMargin=36, bottomMargin=36
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    COLOR_NEGRO = colors.HexColor("#000000")
    COLOR_GRIS = colors.HexColor("#444444")
    
    estilo_secretaria = ParagraphStyle('SecretariaEstilo', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=0, textColor=COLOR_NEGRO)
    estilo_instituto = ParagraphStyle('InstitutoEstilo', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=11, alignment=0, textColor=COLOR_GRIS)
    estilo_sistema = ParagraphStyle('SistemaEstilo', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9, alignment=0, textColor=colors.HexColor("#666666"))
    estilo_normal = ParagraphStyle('TextoNormal', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11, textColor=COLOR_NEGRO)
    estilo_bold = ParagraphStyle('TextoBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=11, textColor=COLOR_NEGRO)

    # --- ENCABEZADO CON LOGOTIPO PROPORCIONAL Y PROFESIONAL ---
    logo_path = os.path.join(dir_raiz, 'images', 'Logo.png')
    header_elements = []
    
    header_elements.append(Paragraph("<b>SECRETARÍA DE CULTURA</b>", estilo_secretaria))
    header_elements.append(Spacer(1, 2))
    header_elements.append(Paragraph("INSTITUTO NACIONAL DE ESTUDIOS HISTÓRICOS DE LAS REVOLUCIONES DE MÉXICO", estilo_instituto))
    header_elements.append(Spacer(1, 2))
    header_elements.append(Paragraph(f"SISTEMA DE ADMINISTRACIÓN INTEGRAL — ACUSE OFICIAL DE {tipo_gestion.upper()}", estilo_sistema))

    if os.path.exists(logo_path):
        # Se fija únicamente la altura para conservar las proporciones naturales y evitar el efecto compacto
        logo_img = Image(logo_path, height=45)
        logo_img.hAlign = 'LEFT'
        tabla_encabezado = Table([[logo_img, header_elements]], colWidths=[70, 470])
    else:
        tabla_encabezado = Table([[header_elements]], colWidths=[540])

    tabla_encabezado.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))

    story.append(tabla_encabezado)
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=1.2, color=COLOR_NEGRO, spaceAfter=12))
    
    data_tabla = [
        [Paragraph("FOLIO OFICIAL:", estilo_bold), Paragraph(str(folio), estilo_normal)],
        [Paragraph("Fecha de Emisión:", estilo_bold), Paragraph(datetime.now().strftime('%Y-%m-%d %H:%M:%S'), estilo_normal)],
        [Paragraph("Área Solicitante:", estilo_bold), Paragraph(str(area_solicitante), estilo_normal)],
        [Paragraph("Responsable:", estilo_bold), Paragraph(str(responsable), estilo_normal)],
        [Paragraph("Área Destino:", estilo_bold), Paragraph(f"Coordinación de {tipo_gestion}", estilo_normal)],
    ]
    
    t = Table(data_tabla, colWidths=[130, 410])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#fdfdfd")),
        ('BOX', (0,0), (-1,-1), 1, COLOR_NEGRO),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    
    story.append(t)
    story.append(Spacer(1, 15))
    story.append(Paragraph("<i>Detalle del Trámite o Insumo Solicitado:</i>", estilo_bold))
    story.append(Spacer(1, 4))
    
    story.append(Table(
        [[Paragraph(str(recurso_pedido), estilo_normal)]],
        colWidths=[540],
        style=[
            ('BOX', (0,0), (-1,-1), 1, COLOR_NEGRO),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#ffffff")),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]
    ))
    
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>Condiciones y Trazabilidad:</b>", estilo_bold))
    story.append(Paragraph("• El presente documento sirve como comprobante oficial digital de la solicitud ingresada en la Ventanilla Única.", estilo_normal))
    story.append(Paragraph("• El plazo normativo de atención y seguimiento es de 5 días hábiles a partir de la emisión de este acuse.", estilo_normal))
    
    story.append(Spacer(1, 45))
    
    firmas_data = [
        [Paragraph("________________________________________", estilo_normal), Paragraph("________________________________________", estilo_normal)],
        [Paragraph("<b>Firma del Solicitante</b>", estilo_normal), Paragraph(f"<b>Vo. Bo. {tipo_gestion}</b>", estilo_normal)]
    ]
    
    t_firmas = Table(firmas_data, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    
    story.append(t_firmas)
    doc.build(story)
    return ruta_pdf

def generar_pdf_incidencia_rh(folio_incidencia, area, empleado, tipo_incidencia, fecha_incidencia, estatus_quincena):
    dir_modulos = os.path.dirname(os.path.abspath(__file__))
    dir_raiz = os.path.dirname(dir_modulos)
    reportes_dir = os.path.join(dir_raiz, 'reportes_pdf')
    os.makedirs(reportes_dir, exist_ok=True)
    
    nombre_archivo = f"Incidencia_{folio_incidencia.replace('/', '_')}.pdf"
    ruta_pdf = os.path.join(reportes_dir, nombre_archivo)
    
    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=letter,
        rightMargin=36, leftMargin=36,
        topMargin=36, bottomMargin=36
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    COLOR_NEGRO = colors.HexColor("#000000")
    COLOR_GRIS = colors.HexColor("#444444")
    
    estilo_secretaria = ParagraphStyle('Sec', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=0, textColor=COLOR_NEGRO)
    estilo_instituto = ParagraphStyle('Inst', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=11, alignment=0, textColor=COLOR_GRIS)
    estilo_sistema = ParagraphStyle('Sis', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9, alignment=0, textColor=colors.HexColor("#666666"))
    estilo_normal = ParagraphStyle('Norm', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11, textColor=COLOR_NEGRO)
    estilo_bold = ParagraphStyle('Bld', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=11, textColor=COLOR_NEGRO)

    logo_path = os.path.join(dir_raiz, 'images', 'Logo.png')
    header_elements = []
    
    header_elements.append(Paragraph("<b>SECRETARÍA DE CULTURA</b>", estilo_secretaria))
    header_elements.append(Spacer(1, 2))
    header_elements.append(Paragraph("INSTITUTO NACIONAL DE ESTUDIOS HISTÓRICOS DE LAS REVOLUCIONES DE MÉXICO", estilo_instituto))
    header_elements.append(Spacer(1, 2))
    header_elements.append(Paragraph("DIRECCIÓN DE ADMINISTRACIÓN — SUBDIRECCIÓN DE RECURSOS HUMANOS", estilo_sistema))

    if os.path.exists(logo_path):
        # Altura proporcional para mantener la estética corporativa
        logo_img = Image(logo_path, height=45)
        logo_img.hAlign = 'LEFT'
        tabla_encabezado = Table([[logo_img, header_elements]], colWidths=[70, 470])
    else:
        tabla_encabezado = Table([[header_elements]], colWidths=[540])

    tabla_encabezado.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))

    story.append(tabla_encabezado)
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=1.2, color=COLOR_NEGRO, spaceAfter=12))
    
    story.append(Paragraph(f"<b>CÉDULA DE REGISTRO DE INCIDENCIA:</b> {folio_incidencia}", ParagraphStyle('Folio', parent=estilo_bold, fontSize=10, textColor=COLOR_NEGRO)))
    story.append(Paragraph(f"<b>FECHA DE EMISIÓN:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", estilo_normal))
    story.append(Spacer(1, 8))
    
    data_tabla = [
        [Paragraph("Área de Adscripción:", estilo_bold), Paragraph(str(area), estilo_normal)],
        [Paragraph("Nombre del Colaborador:", estilo_bold), Paragraph(str(empleado), estilo_normal)],
        [Paragraph("Tipo de Incidencia:", estilo_bold), Paragraph(str(tipo_incidencia), estilo_normal)],
        [Paragraph("Fecha del Evento:", estilo_bold), Paragraph(str(fecha_incidencia), estilo_normal)],
        [Paragraph("Estatus Quincenal:", estilo_bold), Paragraph(str(estatus_quincena), estilo_normal)],
    ]
    
    t = Table(data_tabla, colWidths=[140, 400])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#fdfdfd")),
        ('BOX', (0,0), (-1,-1), 1, COLOR_NEGRO),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    
    story.append(t)
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>Disposiciones Normativas y Trazabilidad:</b>", estilo_bold))
    story.append(Paragraph("• El presente registro se integra al expediente laboral conforme a las Condiciones Generales de Trabajo.", estilo_normal))
    story.append(Paragraph("• Su validación se encuentra supeditada al candado quincenal de cierre normativo de nómina.", estilo_normal))
    
    story.append(Spacer(1, 45))
    
    firmas_data = [
        [Paragraph("________________________________________", estilo_normal), Paragraph("________________________________________", estilo_normal)],
        [Paragraph("<b>Firma del Servidor Público</b>", estilo_normal), Paragraph("<b>Vo. Bo. Subdirección de Recursos Humanos</b>", estilo_normal)]
    ]
    
    t_firmas = Table(firmas_data, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    
    story.append(t_firmas)
    doc.build(story)
    return ruta_pdf
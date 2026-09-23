import io
import json
import time
import re
import streamlit as st
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv
from google import genai
from google.genai import types

# ReportLab para la generación del PDF profesional
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage

load_dotenv()

st.set_page_config(
    page_title="Taller IA — Sistema Integral de Siniestros",
    page_icon="🚗",
    layout="wide"
)

st.markdown("""
    <style>
    .titulo-pericial { font-size: 26px; font-weight: 800; color: #1F497D; margin-bottom: 2px; }
    .sub-pericial { font-size: 14px; color: #555555; margin-bottom: 18px; }
    .card-vehiculo { background-color: #F1F4F8; padding: 14px 20px; border-radius: 6px; border-left: 5px solid #1F497D; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="titulo-pericial">Taller IA — Peritaje, Baremo y Emisión de Presupuestos</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-pericial">Detección visual, deducción mecánica, baremo configurable y exportación profesional en PDF y Excel</div>', unsafe_allow_html=True)

# -------------------------------------------------------------
# FUNCIÓN DE LLAMADA SEGURA CON REINTENTOS (ANTI-ERROR 503)
# -------------------------------------------------------------
def llamar_gemini_con_reintentos(client, model, contents, config, max_reintentos=4):
    for intento in range(max_reintentos):
        try:
            return client.models.generate_content(
                model=model,
                contents=contents,
                config=config
            )
        except Exception as e:
            error_str = str(e).lower()
            if ("503" in error_str or "high demand" in error_str or "429" in error_str or "unavailable" in error_str) and intento < max_reintentos - 1:
                tiempo_espera = (intento + 1) * 3
                time.sleep(tiempo_espera)
                continue
            raise e

# -------------------------------------------------------------
# 1. BARRA LATERAL: IDENTIDAD DEL TALLER Y CLIENTE
# -------------------------------------------------------------
with st.sidebar:
    st.subheader("1. Identidad del Taller")
    logo_subido = st.file_uploader("Subir Logo del Taller (PNG / JPG):", type=["png", "jpg", "jpeg"])
    nombre_taller = st.text_input("Nombre / Razón Social:", value="MAXIAUTOMOTORES")
    dir_taller = st.text_input("Dirección:", value="Av Juan Bautista Justo 7214 CABA")
    fiscal_taller = st.text_input("Condición Fiscal:", value="Responsable Monotributista")
    cuit_taller = st.text_input("CUIT:", value="20-34151842-4")
    tel_taller = st.text_input("Teléfono:", value="15-5581-9975")
    mail_taller = st.text_input("Email:", value="maxiautomotores@gmail.com")

    st.markdown("---")
    st.subheader("2. Datos del Presupuesto y Cliente")
    nro_presupuesto = st.text_input("N° Presupuesto:", value="0002591")
    nombre_titular = st.text_input("Titular / Asegurado:", value="Vargas Hernandez Alvaro Paul")
    dni_titular = st.text_input("DNI / CUIT Titular:", value="96.012.998")
    domicilio_titular = st.text_input("Domicilio Titular:", value="Av Belgrano 2725, CABA")
    tel_titular = st.text_input("Teléfono Titular:", value="")

    st.markdown("---")
    st.subheader("3. Baremos y Tarifas")
    tarifa_pano = st.number_input("Valor por paño de pintura ($ ARS):", min_value=0, value=140000, step=5000)
    tarifa_dia_chapa = st.number_input("Valor día chapa pesada / banco ($ ARS):", min_value=0, value=110000, step=5000)
    tarifa_mecanica_base = st.number_input("M.O. Mecánica / Arme y Desarme ($ ARS):", min_value=0, value=300000, step=10000)
    costo_materiales = st.number_input("Materiales pintura y selladores ($ ARS):", min_value=0, value=150000, step=10000)

if "peritaje_listo" not in st.session_state:
    st.session_state.peritaje_listo = False
if "datos_peritaje" not in st.session_state:
    st.session_state.datos_peritaje = None
if "repuestos_cotizados" not in st.session_state:
    st.session_state.repuestos_cotizados = []

# -------------------------------------------------------------
# 2. CARGA DE FOTOS DEL SINIESTRO
# -------------------------------------------------------------
archivos_subidos = st.file_uploader(
    "Subí las fotos del siniestro (frente, laterales, vano motor) y/o la cédula:",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if archivos_subidos:
    cols = st.columns(min(len(archivos_subidos), 4))
    for i, arch in enumerate(archivos_subidos[:4]):
        with cols[i]:
            st.image(Image.open(arch), caption=f"Foto {i+1}", use_container_width=True)

boton_iniciar = st.button(
    "Iniciar Peritaje Técnico Integral",
    type="primary",
    use_container_width=True,
    disabled=(not archivos_subidos)
)

# -------------------------------------------------------------
# 3. MOTOR PERICIAL INTELIGENTE
# -------------------------------------------------------------
if boton_iniciar and archivos_subidos:
    client = genai.Client()
    imagenes_pil = []
    for a in archivos_subidos:
        im = Image.open(a).convert("RGB")
        im.thumbnail((1024, 1024))
        imagenes_pil.append(im)

    with st.spinner("Analizando cinemática de impacto, deduciendo piezas y cotizando a valores de reposición..."):
        prompt = """
        Sos un perito liquidador senior de siniestros automotores en Argentina y jefe técnico de taller.
        Analizá minuciosamente el lote de imágenes subidas.

        TAREAS TÉCNICAS:
        1. IDENTIFICACIÓN VEHICULAR:
           - Si hay cédula: leé exactamente Marca, Modelo, Versión, Año, Motor y Patente.
           - Si no hay cédula: deducí Marca, Modelo, Versión/Año y Patente visible.
        2. DICTAMEN TÉCNICO:
           - Redactá un dictamen pericial técnico describiendo el vector de impacto y deformación estructural.
        3. DESPIECE EXHAUSTIVO Y COTIZACIÓN DE REPOSICIÓN:
           En choques frontales o asimétricos, deducí OBLIGATORIAMENTE repuestos de:
           - Carrocería exterior (Capot, Paragolpes, Parrilla con emblema, Ópticas, Guardabarros, Patente duplicado).
           - Estructura oculta (Marco porta radiadores, Alma de paragolpes, Puntera de chasis, Pasarruedas).
           - Mecánica y tren delantero (Patas de motor izq/der, Pata de caja, Semieje del lado dañado, Radiador de agua, Condensador A/A, Electroventilador, Depósito lavaparabrisas con bomba).
           Cotizá cada pieza al valor de reposición original OEM de concesionario en pesos argentinos (ARS).
        4. BAREMOS:
           - 'panos_pintura': Paños necesarios según piezas a intervenir (decimal).
           - 'dias_chapa_banco': Días de banco de estiramiento para punteras (decimal).
           - 'requiere_mecanica_pesada': true/false.

        Respondé ÚNICAMENTE en formato JSON:
        {
          "vehiculo_detectado": "Volkswagen Gol Trend G7 (2018)",
          "patente_detectada": "AC 051 ID",
          "fuente_identificacion": "Peritaje visual de carrocería y chapa patente",
          "diagnostico_cinematica": "Dictamen técnico pericial...",
          "panos_pintura": 5.5,
          "dias_chapa_banco": 4.0,
          "requiere_mecanica_pesada": true,
          "repuestos": [
            {"categoria": "Carrocería Exterior", "pieza": "Capot Original", "precio_oem": 1619460.0},
            {"categoria": "Carrocería Exterior", "pieza": "Paragolpes Delantero Original", "precio_oem": 1096250.0},
            {"categoria": "Carrocería Exterior", "pieza": "Parrilla Superior de Radiador con Emblema Original", "precio_oem": 616750.0},
            {"categoria": "Carrocería Exterior", "pieza": "Óptica Delantera Izquierda Original", "precio_oem": 475710.0},
            {"categoria": "Carrocería Exterior", "pieza": "Guardabarros Delantero Izquierdo Original", "precio_oem": 471960.0},
            {"categoria": "Carrocería Exterior", "pieza": "Duplicado Legal de Chapa Patente Mercosur", "precio_oem": 45000.0},
            {"categoria": "Estructura Oculta", "pieza": "Panel Frente Porta Radiadores", "precio_oem": 240910.0},
            {"categoria": "Estructura Oculta", "pieza": "Alma / Travesaño de Paragolpes Delantero", "precio_oem": 210000.0},
            {"categoria": "Estructura Oculta", "pieza": "Pasarruedas Interior Izquierdo", "precio_oem": 185000.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Pata / Soporte de Motor Izquierdo Original", "precio_oem": 600970.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Pata / Soporte de Motor Derecho Original", "precio_oem": 242410.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Soporte / Pata de Caja de Cambios Trasera", "precio_oem": 82850.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Semieje Delantero Izquierdo Completo", "precio_oem": 317950.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Radiador de Agua de Motor Original", "precio_oem": 324020.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Condensador de Aire Acondicionado", "precio_oem": 389400.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Electroventilador Completo c/ Encauzador", "precio_oem": 155760.0},
            {"categoria": "Mecánica y Tren Delantero", "pieza": "Depósito Líquido Lavaparabrisas c/ Bomba Original", "precio_oem": 136370.0}
          ]
        }
        """

        try:
            resp = llamar_gemini_con_reintentos(
                client=client,
                model='gemini-3.6-flash',
                contents=[prompt] + imagenes_pil,
                config=types.GenerateContentConfig(response_mime_type="application/json"),
                max_reintentos=4
            )
            st.session_state.datos_peritaje = json.loads(resp.text)
            st.session_state.peritaje_listo = True
        except Exception as e:
            st.error(f"Error en el procesamiento: {e}")
            st.stop()

    if st.session_state.peritaje_listo and st.session_state.datos_peritaje:
        datos = st.session_state.datos_peritaje
        auto_str = datos.get("vehiculo_detectado", "Auto")
        st.session_state.repuestos_cotizados = []

        for item in datos.get("repuestos", []):
            slug = re.sub(r'\s+', '-', re.sub(r'[^a-zA-Z0-9\s]', ' ', f"{item['pieza']} {auto_str}").lower().strip())
            st.session_state.repuestos_cotizados.append({
                "categoria": item.get("categoria", "Carrocería Exterior"),
                "pieza": item["pieza"],
                "precio": float(item["precio_oem"]),
                "link": f"https://listado.mercadolibre.com.ar/{slug}_Condicion_Nuevo"
            })
        st.rerun()

# -------------------------------------------------------------
# 4. RESULTADOS, EDICIÓN Y EXPORTACIÓN DUAL (PDF + EXCEL)
# -------------------------------------------------------------
if st.session_state.peritaje_listo and st.session_state.datos_peritaje:
    datos = st.session_state.datos_peritaje

    st.markdown(f"""
        <div class="card-vehiculo">
            <h4 style="margin:0; color:#1F497D;">Vehículo: {datos.get('vehiculo_detectado')} - {datos.get('patente_detectada')}</h4>
            <p style="margin:4px 0 0 0; font-size:13px; color:#333;">
                <strong>Identificación:</strong> {datos.get('fuente_identificacion')} | 
                <strong>Línea de Reposición:</strong> Original OEM Concesionario
            </p>
        </div>
    """, unsafe_allow_html=True)

    with st.expander("Ver Dictamen Técnico Pericial", expanded=False):
        st.write(datos.get("diagnostico_cinematica"))

    # BAREMOS
    st.subheader("1. Mano de Obra y Baremos de Reparación")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        panos = st.number_input("Paños de Pintura:", value=float(datos.get("panos_pintura", 5.5)), step=0.5)
        sub_pintura = panos * tarifa_pano
        st.caption(f"Pintura: **${sub_pintura:,.2f}**")

    with c2:
        dias_chapa = st.number_input("Días Chapa Pesada / Banco:", value=float(datos.get("dias_chapa_banco", 4.0)), step=0.5)
        sub_chapa = dias_chapa * tarifa_dia_chapa
        st.caption(f"Chapa: **${sub_chapa:,.2f}**")

    with c3:
        req_mec = datos.get("requiere_mecanica_pesada", True)
        aplica_mec = st.checkbox("M.O. Mecánica y A/A", value=req_mec)
        sub_mec = tarifa_mecanica_base if aplica_mec else 0.0
        st.caption(f"Mecánica: **${sub_mec:,.2f}**")

    with c4:
        aplica_mat = st.checkbox("Materiales y Selladores", value=True)
        sub_mat = costo_materiales if aplica_mat else 0.0
        st.caption(f"Materiales: **${sub_mat:,.2f}**")

    total_mo = sub_pintura + sub_chapa + sub_mec + sub_mat

    # REPUESTOS
    st.subheader("2. Detalle de Repuestos Reclamados")
    piezas_finales = []
    for idx, r in enumerate(st.session_state.repuestos_cotizados):
        col_c, col_p, col_l = st.columns([5, 2, 2])
        with col_c:
            activa = st.checkbox(f"**{r['pieza']}**", value=True, key=f"r_{idx}")
        with col_p:
            pr = st.number_input("Precio ($ ARS)", value=float(r["precio"]), step=5000.0, key=f"pr_{idx}", label_visibility="collapsed")
        with col_l:
            st.markdown(f"[Ver en Mercado Libre]({r['link']})")

        if activa:
            piezas_finales.append({
                "categoria": r["categoria"],
                "pieza": r["pieza"],
                "precio": pr,
                "link": r["link"]
            })

    total_repuestos = sum(p["precio"] for p in piezas_finales)
    total_general = total_repuestos + total_mo

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Subtotal Mano de Obra", f"${total_mo:,.2f}")
    m2.metric("Subtotal Repuestos", f"${total_repuestos:,.2f}")
    m3.metric("TOTAL PRESUPUESTADO", f"${total_general:,.2f}")

    # -------------------------------------------------------------
    # GENERADOR DE PDF PROFESIONAL CON LOGO PERSONALIZADO
    # -------------------------------------------------------------
    def generar_pdf():
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            leftMargin=25,
            rightMargin=25,
            topMargin=25,
            bottomMargin=25
        )

        elementos = []
        styles = getSampleStyleSheet()

        style_taller_titulo = ParagraphStyle('TallerTit', fontName='Helvetica-Bold', fontSize=14, leading=16, textColor=colors.HexColor("#1A2B4C"))
        style_taller_sub = ParagraphStyle('TallerSub', fontName='Helvetica', fontSize=8, leading=10.5, textColor=colors.HexColor("#333333"))
        style_pres_titulo = ParagraphStyle('PresTit', fontName='Helvetica-Bold', fontSize=11, leading=13, alignment=2, textColor=colors.HexColor("#1A2B4C"))
        style_pres_sub = ParagraphStyle('PresSub', fontName='Helvetica', fontSize=8, leading=10.5, alignment=2, textColor=colors.HexColor("#333333"))
        style_sec_title = ParagraphStyle('SecTit', fontName='Helvetica-Bold', fontSize=9.5, leading=11.5, textColor=colors.HexColor("#1A2B4C"))
        style_celda = ParagraphStyle('Celda', fontName='Helvetica', fontSize=7.5, leading=9.5)
        style_celda_bold = ParagraphStyle('CeldaB', fontName='Helvetica-Bold', fontSize=7.5, leading=9.5)
        style_celda_num = ParagraphStyle('CeldaN', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=2)
        style_celda_num_b = ParagraphStyle('CeldaNB', fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=2)

        # Encabezado: Taller y Logo a la Izquierda / Presupuesto a la Derecha
        taller_info = []
        if logo_subido is not None:
            try:
                logo_bytes = io.BytesIO(logo_subido.getvalue())
                rl_img = RLImage(logo_bytes, width=110, height=45)
                taller_info.append(rl_img)
                taller_info.append(Spacer(1, 4))
            except Exception:
                taller_info.append(Paragraph(f"<b>{nombre_taller.upper()}</b>", style_taller_titulo))
        else:
            taller_info.append(Paragraph(f"<b>{nombre_taller.upper()}</b>", style_taller_titulo))

        taller_info.extend([
            Paragraph(dir_taller, style_taller_sub),
            Paragraph(fiscal_taller, style_taller_sub),
            Paragraph(f"CUIT: {cuit_taller}", style_taller_sub),
            Paragraph(f"Tel: {tel_taller}", style_taller_sub),
            Paragraph(f"Email: {mail_taller}", style_taller_sub),
        ])

        fecha_str = time.strftime("%d de Septiembre de %Y")
        hora_str = time.strftime("%H:%M hs")

        doc_info = [
            Paragraph(f"<b>Presupuesto - Nro: {nro_presupuesto}</b>", style_pres_titulo),
            Paragraph(f"<b>{nombre_taller.upper()}</b>", style_pres_sub),
            Spacer(1, 4),
            Paragraph(f"Fecha: {fecha_str}", style_pres_sub),
            Paragraph(f"Hora: {hora_str}", style_pres_sub),
        ]

        t_header = Table([[taller_info, doc_info]], colWidths=[280, 265])
        t_header.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        elementos.append(t_header)
        elementos.append(Spacer(1, 8))

        # Cuadro de Cliente y Vehículo
        vehiculo_str = f"{datos.get('vehiculo_detectado')} - {datos.get('patente_detectada')}"
        info_cliente = [
            [Paragraph(f"<b>Vehículo:</b> {vehiculo_str}", style_celda_bold), Paragraph(f"<b>Titular:</b> {nombre_titular}", style_celda)],
            [Paragraph(f"<b>DNI / CUIT:</b> {dni_titular}", style_celda), Paragraph(f"<b>Teléfono:</b> {tel_titular if tel_titular else 'S/D'}", style_celda)],
            [Paragraph(f"<b>Domicilio:</b> {domicilio_titular}", style_celda), Paragraph("<b>Localidad:</b> CABA, Buenos Aires", style_celda)]
        ]
        t_cli = Table(info_cliente, colWidths=[270, 275])
        t_cli.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#B0C0D0")),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        elementos.append(t_cli)
        elementos.append(Spacer(1, 8))

        # Tareas / Mano de Obra
        elementos.append(Paragraph("<b>Detalles de tareas (Mano de Obra)</b>", style_sec_title))
        elementos.append(Spacer(1, 3))
        tareas_data = [
            [Paragraph("<b>Descripción de la Tarea</b>", style_celda_bold), Paragraph("<b>Detalle de Baremo</b>", style_celda_bold), Paragraph("<b>Subtotal</b>", style_celda_num_b)],
            [Paragraph("M.O. CHAPA PESADA Y BANCO", style_celda), Paragraph(f"{dias_chapa} días de chapa / estiramiento", style_celda), Paragraph(f"$ {sub_chapa:,.2f}", style_celda_num)],
            [Paragraph("M.O. PINTURA EN CABINA", style_celda), Paragraph(f"{panos} paños de pintura", style_celda), Paragraph(f"$ {sub_pintura:,.2f}", style_celda_num)],
            [Paragraph("M.O. ARME, DESARME Y MECÁNICA", style_celda), Paragraph("Desarme de trompa, radiadores y mecánica", style_celda), Paragraph(f"$ {sub_mec:,.2f}", style_celda_num)],
            [Paragraph("MATERIALES DE PINTURA Y SELLADORES", style_celda), Paragraph("Insumos de pintura y sellado oficial", style_celda), Paragraph(f"$ {sub_mat:,.2f}", style_celda_num)],
            [Paragraph("<b>TOTAL MANO DE OBRA</b>", style_celda_bold), Paragraph("", style_celda), Paragraph(f"<b>$ {total_mo:,.2f}</b>", style_celda_num_b)]
        ]
        t_tar = Table(tareas_data, colWidths=[240, 185, 120])
        t_tar.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EAEFF5")),
            ('LINEBELOW', (0,0), (-1,0), 1, colors.HexColor("#1A2B4C")),
            ('LINEBELOW', (0,1), (-1,-2), 0.5, colors.HexColor("#E2E8F0")),
            ('LINEABOVE', (0,-1), (-1,-1), 1, colors.HexColor("#1A2B4C")),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#F1F4F8")),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ]))
        elementos.append(t_tar)
        elementos.append(Spacer(1, 8))

        # Dictamen técnico
        elementos.append(Paragraph("<b>Informe técnico:</b>", style_sec_title))
        elementos.append(Spacer(1, 2))
        style_dic = ParagraphStyle('Dic', fontName='Helvetica-Oblique', fontSize=7, leading=9, textColor=colors.HexColor("#444444"))
        elementos.append(Paragraph(datos.get("diagnostico_cinematica", "Sin dictamen"), style_dic))
        elementos.append(Spacer(1, 8))

        # Repuestos
        elementos.append(Paragraph("<b>Detalle repuestos utilizados</b>", style_sec_title))
        elementos.append(Spacer(1, 3))
        rep_data = [
            [Paragraph("<b>Cant.</b>", style_celda_bold), Paragraph("<b>Descripción</b>", style_celda_bold), Paragraph("<b>Importe unitario</b>", style_celda_num_b), Paragraph("<b>Importe total</b>", style_celda_num_b)]
        ]
        for p in piezas_finales:
            rep_data.append([
                Paragraph("1", style_celda),
                Paragraph(p["pieza"], style_celda),
                Paragraph(f"$ {p['precio']:,.2f}", style_celda_num),
                Paragraph(f"$ {p['precio']:,.2f}", style_celda_num)
            ])

        rep_data.append([
            Paragraph("", style_celda),
            Paragraph("<b>SUBTOTAL REPUESTOS</b>", style_celda_bold),
            Paragraph("", style_celda),
            Paragraph(f"<b>$ {total_repuestos:,.2f}</b>", style_celda_num_b)
        ])
        rep_data.append([
            Paragraph("", style_celda),
            Paragraph("<b>TOTAL PRESUPUESTADO</b>", style_sec_title),
            Paragraph("", style_celda),
            Paragraph(f"<b>$ {total_general:,.2f}</b>", ParagraphStyle('TotF', parent=style_celda_num_b, fontSize=9.5, textColor=colors.HexColor("#9C0000")))
        ])

        t_rep = Table(rep_data, colWidths=[35, 270, 120, 120])
        t_rep.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EAEFF5")),
            ('LINEBELOW', (0,0), (-1,0), 1, colors.HexColor("#1A2B4C")),
            ('LINEBELOW', (0,1), (-1,-3), 0.5, colors.HexColor("#E2E8F0")),
            ('LINEABOVE', (0,-2), (-1,-2), 1, colors.HexColor("#1A2B4C")),
            ('LINEABOVE', (0,-1), (-1,-1), 1.5, colors.HexColor("#9C0000")),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#FDF2F2")),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        elementos.append(t_rep)
        elementos.append(Spacer(1, 8))

        style_pie = ParagraphStyle('Pie', fontName='Helvetica-Bold', fontSize=7.5, alignment=1, textColor=colors.HexColor("#555555"))
        elementos.append(Paragraph("Presupuesto válido por 30 días", style_pie))

        doc.build(elementos)
        buffer.seek(0)
        return buffer.getvalue()

    # -------------------------------------------------------------
    # GENERADOR DE EXCEL OFICIAL (.XLSX)
    # -------------------------------------------------------------
    def generar_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Presupuesto Oficial"
        ws.views.sheetView[0].showGridLines = True

        azul = "1F497D"
        gris = "595959"
        fondo_total = "EBF1F5"

        borde = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        borde_doble_abajo = Border(
            top=Side(style='thin', color='1F497D'),
            bottom=Side(style='double', color='1F497D')
        )

        ws["A1"] = f"{nombre_taller.upper()} — PRESUPUESTO OFICIAL N° {nro_presupuesto}"
        ws["A1"].font = Font(size=13, bold=True, color=azul)
        ws["A2"] = f"Vehículo: {datos.get('vehiculo_detectado')} | Dominio: {datos.get('patente_detectada')} | Titular: {nombre_titular}"
        ws["A2"].font = Font(size=9.5, italic=True)

        fila = 4
        ws.cell(row=fila, column=1, value="I. DETALLE DE REPUESTOS RECLAMADOS").font = Font(bold=True, color=azul)
        fila += 1

        for c_idx, h in enumerate(["Ítem", "Rubro / Sistema", "Repuesto Reclamado", "Importe Unitario (ARS)", "Importe Total (ARS)", "Enlace Testigo"], 1):
            c = ws.cell(row=fila, column=c_idx, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color=azul, end_color=azul, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")

        fila += 1
        for idx, p in enumerate(piezas_finales, 1):
            ws.cell(row=fila, column=1, value=idx).alignment = Alignment(horizontal="center")
            ws.cell(row=fila, column=2, value=p["categoria"])
            ws.cell(row=fila, column=3, value=p["pieza"])
            
            c_u = ws.cell(row=fila, column=4, value=float(p["precio"]))
            c_u.number_format = "$#,##0.00"
            c_t = ws.cell(row=fila, column=5, value=float(p["precio"]))
            c_t.number_format = "$#,##0.00"

            ws.cell(row=fila, column=6, value=p["link"])
            for col in range(1, 7):
                ws.cell(row=fila, column=col).border = borde
            fila += 1

        ws.cell(row=fila, column=4, value="SUBTOTAL REPUESTOS:").font = Font(bold=True, color=azul)
        ws.cell(row=fila, column=4).alignment = Alignment(horizontal="right")
        c_tot_rep = ws.cell(row=fila, column=5, value=float(total_repuestos))
        c_tot_rep.font = Font(bold=True, color=azul)
        c_tot_rep.number_format = "$#,##0.00"

        fila += 2
        ws.cell(row=fila, column=1, value="II. MANO DE OBRA Y BAREMOS DE TALLER").font = Font(bold=True, color=azul)
        fila += 1

        for c_idx, h in enumerate(["Ítem", "Concepto de Tarea", "Detalle de Baremo", "Subtotal (ARS)"], 1):
            c = ws.cell(row=fila, column=c_idx, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color=gris, end_color=gris, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")

        fila += 1
        tareas = [
            ("M.O. Chapa pesada y banco", f"{dias_chapa} días @ ${tarifa_dia_chapa:,.2f}", sub_chapa),
            ("M.O. Pintura en cabina", f"{panos} paños @ ${tarifa_pano:,.2f}", sub_pintura),
            ("M.O. Arme, desarme y mecánica", "Desarme frontal y radiadores", sub_mec),
            ("Materiales de pintura y selladores", "Insumos oficiales de sellado", sub_mat)
        ]
        for i, t in enumerate(tareas, 1):
            ws.cell(row=fila, column=1, value=i).alignment = Alignment(horizontal="center")
            ws.cell(row=fila, column=2, value=t[0])
            ws.cell(row=fila, column=3, value=t[1])
            c_mo = ws.cell(row=fila, column=4, value=float(t[2]))
            c_mo.number_format = "$#,##0.00"
            for col in range(1, 5):
                ws.cell(row=fila, column=col).border = borde
            fila += 1

        ws.cell(row=fila, column=3, value="SUBTOTAL MANO DE OBRA:").font = Font(bold=True, color=gris)
        ws.cell(row=fila, column=3).alignment = Alignment(horizontal="right")
        c_tot_mo = ws.cell(row=fila, column=4, value=float(total_mo))
        c_tot_mo.font = Font(bold=True, color=gris)
        c_tot_mo.number_format = "$#,##0.00"

        fila += 2
        ws.cell(row=fila, column=3, value="TOTAL GENERAL PRESUPUESTADO:").font = Font(size=11, bold=True, color=azul)
        ws.cell(row=fila, column=3).alignment = Alignment(horizontal="right")
        tot = ws.cell(row=fila, column=4, value=float(total_general))
        tot.font = Font(size=12, bold=True, color="B00000")
        tot.number_format = "$#,##0.00"
        tot.fill = PatternFill(start_color=fondo_total, end_color=fondo_total, fill_type="solid")
        tot.border = borde_doble_abajo

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 38
        ws.column_dimensions["D"].width = 24
        ws.column_dimensions["E"].width = 24
        ws.column_dimensions["F"].width = 40

        buff = io.BytesIO()
        wb.save(buff)
        return buff.getvalue()

    # BOTONES DE DESCARGA DUAL LADO A LADO
    col_pdf, col_xlsx = st.columns(2)
    with col_pdf:
        st.download_button(
            label="📄 Descargar Presupuesto en PDF (Cliente / Aseguradora)",
            data=generar_pdf(),
            file_name=f"Presupuesto_{datos.get('patente_detectada','AUTO').replace(' ','_')}_{nro_presupuesto}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    with col_xlsx:
        st.download_button(
            label="📊 Descargar Planilla de Control en Excel (.xlsx)",
            data=generar_excel(),
            file_name=f"Presupuesto_{datos.get('patente_detectada','AUTO').replace(' ','_')}_{nro_presupuesto}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
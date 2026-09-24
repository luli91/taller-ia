import io
import json
import time
import streamlit as st
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv
from google import genai
from google.genai import types

# ReportLab para emisión del PDF formal
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage

load_dotenv()

st.set_page_config(
    page_title="Taller IA — Sistema Pericial y Liquidación Oficial",
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

st.markdown('<div class="titulo-pericial">Taller IA — Peritaje Integral de Siniestros</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-pericial">Liquidación pericial a valores oficiales de reposición | Autos y Motos | Emisión en PDF y Excel</div>', unsafe_allow_html=True)

# -------------------------------------------------------------
# LLAMADA ROBUSTA CON REINTENTOS (ANTI-ERROR 503 / 429)
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
                time.sleep((intento + 1) * 3)
                continue
            raise e

# -------------------------------------------------------------
# 1. BARRA LATERAL: IDENTIDAD DEL TALLER, CLIENTE Y BAREMOS
# -------------------------------------------------------------
with st.sidebar:
    st.subheader("1. Identidad del Taller")
    logo_subido = st.file_uploader("Subir Logo del Taller:", type=["png", "jpg", "jpeg"])
    nombre_taller = st.text_input("Nombre / Razón Social:", value="MAXIAUTOMOTORES")
    dir_taller = st.text_input("Dirección:", value="Av Juan Bautista Justo 7214 CABA")
    fiscal_taller = st.text_input("Condición Fiscal:", value="Responsable Monotributista")
    cuit_taller = st.text_input("CUIT:", value="20-34151842-4")
    tel_taller = st.text_input("Teléfono:", value="15-5581-9975")
    mail_taller = st.text_input("Email:", value="maxiautomotores@gmail.com")

    st.markdown("---")
    st.subheader("2. Datos del Presupuesto y Cliente")
    nro_presupuesto = st.text_input("N° Presupuesto:", value="0002594")
    nombre_titular = st.text_input("Titular / Asegurado:", value="Pirri Daniel Rodrigo")
    dni_titular = st.text_input("DNI / CUIT:", value="38.319.206")
    domicilio_titular = st.text_input("Domicilio:", value="Ortiz de Rozas 875, Moron")
    tel_titular = st.text_input("Teléfono Titular:", value="11-6917-6032")

    st.markdown("---")
    st.subheader("3. Baremos y Tarifas Base")
    tipo_vehiculo = st.radio("Tipo de Vehículo:", ["Automotor / Utilitario", "Motovehículo / Moto"], index=0)
    tarifa_pano = st.number_input("Valor por paño de pintura (ARS):", min_value=0, value=140000, step=5000)
    tarifa_dia_chapa = st.number_input("Valor día chapa pesada / banco (ARS):", min_value=0, value=110000, step=5000)
    tarifa_mecanica = st.number_input("M.O. Tren Delantero / Desarme (ARS):", min_value=0, value=300000, step=10000)
    tarifa_motor_alineacion = st.number_input("M.O. Motor / Cárter / Alineación (ARS):", min_value=0, value=240000, step=10000)
    tarifa_service_fluidos = st.number_input("Kit Service Aceite y Filtros (ARS):", min_value=0, value=316000, step=5000)
    costo_materiales = st.number_input("Materiales e insumos selladores (ARS):", min_value=0, value=150000, step=10000)

if "peritaje_listo" not in st.session_state:
    st.session_state.peritaje_listo = False
if "datos_peritaje" not in st.session_state:
    st.session_state.datos_peritaje = None
if "repuestos_cotizados" not in st.session_state:
    st.session_state.repuestos_cotizados = []

# -------------------------------------------------------------
# 2. CARGA DE ARCHIVOS (FOTOS O EXPEDIENTES EN PDF)
# -------------------------------------------------------------
archivos_subidos = st.file_uploader(
    "Subí fotos del siniestro (JPG/PNG) o carpetas periciales en PDF enviadas por estudios jurídicos:",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

if archivos_subidos:
    st.success(f"{len(archivos_subidos)} archivo(s) listo(s) para procesar.")

boton_iniciar = st.button(
    "Iniciar Peritaje y Liquidación Oficial",
    type="primary",
    use_container_width=True,
    disabled=(not archivos_subidos)
)

# -------------------------------------------------------------
# 3. MOTOR PERICIAL INTELIGENTE MULTIZONA
# -------------------------------------------------------------
if boton_iniciar and archivos_subidos:
    client = genai.Client()
    contenidos_gemini = []

    for a in archivos_subidos:
        bytes_data = a.read()
        if a.name.lower().endswith(".pdf"):
            contenidos_gemini.append(types.Part.from_bytes(data=bytes_data, mime_type="application/pdf"))
        else:
            im = Image.open(io.BytesIO(bytes_data)).convert("RGB")
            im.thumbnail((1024, 1024))
            buf = io.BytesIO()
            im.save(buf, format="JPEG")
            contenidos_gemini.append(types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"))

    with st.spinner("Analizando cinemática pericial, deduciendo piezas y cotizando a valores de reposición..."):
        prompt = f"""
        Sos un perito liquidador senior de siniestros viales en Argentina y jefe técnico de taller.
        Analizá minuciosamente el material fotográfico y/o documental en PDF subido.
        Vehículo configurado: {tipo_vehiculo}.

        TAREAS OBLIGATORIAS:

        1. IDENTIFICACIÓN VEHICULAR:
           - Marca, Modelo exacto, Generación/Año y Patente (visible o deducida de emblemas/documentos).

        2. CLASIFICACIÓN DE LA CINEMÁTICA Y ZONA DE DAÑO:
           - 'POZO_VIAL': Bache, pozo, cordoneo o desnivel vial con rotura de tren rodante y bajos de motor.
           - 'CHOQUE_FRONTAL': Impacto en trompa, capot, frente porta radiadores, mecánica delantera.
           - 'CHOQUE_LATERAL': Impacto sobre guardabarros soldados, puertas, ventiletes y banco de estiramiento.
           - 'CHOQUE_POSTERIOR': Impacto de cola, tapa baúl, panel de cola y piso.
           - 'MOTOVEHICULO': Siniestro de moto (horquilla, manillar, carenados plásticos, escape).

        3. MATRIZ DE LIQUIDACIÓN TÉCNICA (PRECIOS DE REPOSICIÓN ORIGINAL OEM EN ARS):

           A) SI ES 'POZO_VIAL':
              - Chapa y pintura en $0 (panos_pintura=0, dias_chapa_banco=0).
              - Repuestos Obligatorios: Llanta de aleación original rota/abollada, Neumático nuevo de reposición, Parrilla de suspensión delantera, Rótula de suspensión partida, Kit x2 amortiguadores delanteros (par de eje), Kit extremos de dirección y axiales.
              - Bajos de motor: Cárter de motor de aluminio original si hay golpe de panza o fuga de aceite.
              - Mano de Obra: Activar 'requiere_tren_delantero': true, 'requiere_motor_alineacion': true, 'incluye_service_aceite': true.

           B) SI ES 'CHOQUE_FRONTAL':
              - Capot, paragolpes, parrilla, ópticas, panel porta radiadores, alma, puntera, patas de motor/caja, semieje, radiadores. Chapa (3-5 días), Pintura (5-7 paños), Tren delantero: true.

           C) SI ES 'CHOQUE_LATERAL':
              - Lateral soldado (10-14 días chapa), desmonte de ventilete/cristal, llanta de aleación si hubo contacto, eje trasero.

           D) SI ES 'MOTOVEHICULO':
              - Horquilla, barrales, cristo, manillar, palancas, espejos, pedalines, escape, carenados plásticos. Chapa=0, Pintura plásticos (2-4 paños).

        Respondé ÚNICAMENTE en JSON con esta estructura exacta:
        {{
          "vehiculo_detectado": "Volkswagen Gol Trend 1.6",
          "patente_detectada": "IMA 532",
          "fuente_identificacion": "Peritaje visual de carrocería y componentes mecánicos",
          "tipologia_siniestro": "POZO_VIAL",
          "diagnostico_cinematica": "Dictamen pericial técnico detallado...",
          "panos_pintura": 0.0,
          "dias_chapa_banco": 0.0,
          "requiere_tren_delantero": true,
          "requiere_motor_alineacion": true,
          "incluye_service_aceite": true,
          "repuestos": [
            {{"categoria": "Tren Rodante", "pieza": "Llanta de Aleación Rodado 15 Original", "precio_oem": 756120.0}},
            {{"categoria": "Motor y Bajos", "pieza": "Cárter de Motor Aluminio Superior 1.6 8V Original", "precio_oem": 423760.0}},
            {{"categoria": "Suspensión y Dirección", "pieza": "Parrilla de Suspensión Delantera Izquierda Original", "precio_oem": 390080.0}},
            {{"categoria": "Tren Rodante", "pieza": "Neumático 195/55 R15 Nuevo", "precio_oem": 371920.0}},
            {{"categoria": "Suspensión y Dirección", "pieza": "Kit x2 Amortiguadores Delanteros (Par de Eje)", "precio_oem": 330930.0}},
            {{"categoria": "Suspensión y Dirección", "pieza": "Rótula de Suspensión Delantera Izquierda", "precio_oem": 91780.0}},
            {{"categoria": "Suspensión y Dirección", "pieza": "Kit Extremos de Dirección, Axiales y Fuelles Delanteros", "precio_oem": 140000.0}}
          ]
        }}
        """

        try:
            resp = llamar_gemini_con_reintentos(
                client=client,
                model='gemini-3.6-flash',
                contents=[prompt] + contenidos_gemini,
                config=types.GenerateContentConfig(response_mime_type="application/json"),
                max_reintentos=4
            )
            st.session_state.datos_peritaje = json.loads(resp.text)
            datos = st.session_state.datos_peritaje
            st.session_state.repuestos_cotizados = [
                {
                    "categoria": item.get("categoria", "Repuestos"),
                    "pieza": item["pieza"],
                    "precio": float(item["precio_oem"])
                }
                for item in datos.get("repuestos", [])
            ]
            st.session_state.peritaje_listo = True
        except Exception as e:
            st.error(f"Error en el procesamiento del siniestro: {e}")
            st.stop()
        st.rerun()

# -------------------------------------------------------------
# 4. VISUALIZACIÓN, EDICIÓN Y EXPORTACIÓN DUAL
# -------------------------------------------------------------
if st.session_state.peritaje_listo and st.session_state.datos_peritaje:
    datos = st.session_state.datos_peritaje

    st.markdown(f"""
        <div class="card-vehiculo">
            <h4 style="margin:0; color:#1F497D;">Vehículo: {datos.get('vehiculo_detectado')} - {datos.get('patente_detectada')}</h4>
            <p style="margin:4px 0 0 0; font-size:13px; color:#333;">
                <strong>Tipología Detectada:</strong> {datos.get('tipologia_siniestro')} | 
                <strong>Origen de datos:</strong> {datos.get('fuente_identificacion')}
            </p>
        </div>
    """, unsafe_allow_html=True)

    with st.expander("Ver Dictamen Técnico Pericial", expanded=False):
        st.write(datos.get("diagnostico_cinematica"))

    # BAREMOS DE MANO DE OBRA
    st.subheader("1. Mano de Obra y Baremos de Reparación")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        panos = st.number_input("Paños de Pintura:", value=float(datos.get("panos_pintura", 0.0)), step=0.5)
        sub_pintura = panos * tarifa_pano
        st.write(f"Pintura: **${sub_pintura:,.2f}**")

    with c2:
        dias_chapa = st.number_input("Días Banco / Chapa:", value=float(datos.get("dias_chapa_banco", 0.0)), step=0.5)
        sub_chapa = dias_chapa * tarifa_dia_chapa
        st.write(f"Chapa: **${sub_chapa:,.2f}**")

    with c3:
        req_tren = datos.get("requiere_tren_delantero", True)
        aplica_tren = st.checkbox("M.O. Tren Delantero", value=req_tren)
        sub_tren = tarifa_mecanica if aplica_tren else 0.0

        req_motor = datos.get("requiere_motor_alineacion", True)
        aplica_motor = st.checkbox("M.O. Motor y Alineación", value=req_motor)
        sub_motor = tarifa_motor_alineacion if aplica_motor else 0.0

        st.write(f"Tren Del.: **${sub_tren:,.2f}**")
        st.write(f"Motor/Alin.: **${sub_motor:,.2f}**")

    with c4:
        inc_service = datos.get("incluye_service_aceite", False)
        aplica_service = st.checkbox("Kit Service Aceite y Filtros", value=inc_service)
        sub_serv = tarifa_service_fluidos if aplica_service else 0.0
        st.write(f"Fluidos/Service: **${sub_serv:,.2f}**")

    total_mo = sub_pintura + sub_chapa + sub_tren + sub_motor + sub_serv

    # REPUESTOS LIMPIOS (SIN LINKS ROTOS)
    st.subheader(f"2. Detalle de Repuestos Reclamados ({len(st.session_state.repuestos_cotizados)} detectados)")
    piezas_finales = []
    for idx, r in enumerate(st.session_state.repuestos_cotizados):
        col_c, col_p = st.columns([7, 3])
        with col_c:
            activa = st.checkbox(f"**{r['pieza']}**", value=True, key=f"r_{idx}")
        with col_p:
            pr = st.number_input("Precio ARS", value=float(r["precio"]), step=5000.0, key=f"pr_{idx}", label_visibility="collapsed")

        if activa:
            piezas_finales.append({
                "categoria": r["categoria"],
                "pieza": r["pieza"],
                "precio": pr
            })

    total_repuestos = sum(p["precio"] for p in piezas_finales)
    total_general = total_repuestos + total_mo

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Subtotal Mano de Obra y Servicios", f"${total_mo:,.2f}")
    m2.metric("Subtotal Repuestos", f"${total_repuestos:,.2f}")
    m3.metric("TOTAL PRESUPUESTADO", f"${total_general:,.2f}")

    # -------------------------------------------------------------
    # GENERADOR DE PDF OFICIAL (FORMATO FORMAL)
    # -------------------------------------------------------------
    def generar_pdf():
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
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

        elementos.append(Table([[taller_info, doc_info]], colWidths=[280, 265], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
        elementos.append(Spacer(1, 8))

        vehiculo_str = f"{datos.get('vehiculo_detectado')} - {datos.get('patente_detectada')}"
        info_cliente = [
            [Paragraph(f"<b>Vehículo:</b> {vehiculo_str}", style_celda_bold), Paragraph(f"<b>Titular:</b> {nombre_titular}", style_celda)],
            [Paragraph(f"<b>DNI / CUIT:</b> {dni_titular}", style_celda), Paragraph(f"<b>Teléfono:</b> {tel_titular if tel_titular else 'S/D'}", style_celda)],
            [Paragraph(f"<b>Domicilio:</b> {domicilio_titular}", style_celda), Paragraph("<b>Localidad:</b> Morón, Buenos Aires", style_celda)]
        ]
        elementos.append(Table(info_cliente, colWidths=[270, 275], style=[
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#B0C0D0")),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        elementos.append(Spacer(1, 8))

        elementos.append(Paragraph("<b>Detalles de tareas y servicios</b>", style_sec_title))
        elementos.append(Spacer(1, 3))
        tareas_data = [
            [Paragraph("<b>Descripción de la Tarea</b>", style_celda_bold), Paragraph("<b>Detalle</b>", style_celda_bold), Paragraph("<b>Subtotal</b>", style_celda_num_b)]
        ]
        if sub_chapa > 0:
            tareas_data.append([Paragraph("M.O. CHAPA PESADA Y BANCO", style_celda), Paragraph(f"{dias_chapa} días", style_celda), Paragraph(f"$ {sub_chapa:,.2f}", style_celda_num)])
        if sub_pintura > 0:
            tareas_data.append([Paragraph("M.O. PINTURA EN CABINA", style_celda), Paragraph(f"{panos} paños", style_celda), Paragraph(f"$ {sub_pintura:,.2f}", style_celda_num)])
        if sub_tren > 0:
            tareas_data.append([Paragraph("M.O. REPARACIÓN DE TREN DELANTERO", style_celda), Paragraph("Desarme, reemplazo y montaje", style_celda), Paragraph(f"$ {sub_tren:,.2f}", style_celda_num)])
        if sub_motor > 0:
            tareas_data.append([Paragraph("M.O. MECÁNICA DE MOTOR, ARMADO Y ALINEACIÓN", style_celda), Paragraph("Cárter y alineación computarizada", style_celda), Paragraph(f"$ {sub_motor:,.2f}", style_celda_num)])
        if sub_serv > 0:
            tareas_data.append([Paragraph("KIT SERVICE DE ACEITE SINTETICO 5W40 (4L) + FILTROS", style_celda), Paragraph("Fluidos oficiales y filtrado completo", style_celda), Paragraph(f"$ {sub_serv:,.2f}", style_celda_num)])

        tareas_data.append([Paragraph("<b>TOTAL MANO DE OBRA Y SERVICIOS</b>", style_celda_bold), Paragraph("", style_celda), Paragraph(f"<b>$ {total_mo:,.2f}</b>", style_celda_num_b)])

        elementos.append(Table(tareas_data, colWidths=[240, 185, 120], style=[
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EAEFF5")),
            ('LINEBELOW', (0,0), (-1,0), 1, colors.HexColor("#1A2B4C")),
            ('LINEABOVE', (0,-1), (-1,-1), 1, colors.HexColor("#1A2B4C")),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#F1F4F8")),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ]))
        elementos.append(Spacer(1, 8))

        elementos.append(Paragraph("<b>Informe técnico:</b>", style_sec_title))
        elementos.append(Spacer(1, 2))
        elementos.append(Paragraph(datos.get("diagnostico_cinematica", "Sin dictamen"), ParagraphStyle('Dic', fontName='Helvetica-Oblique', fontSize=7, leading=9, textColor=colors.HexColor("#444444"))))
        elementos.append(Spacer(1, 8))

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

        elementos.append(Table(rep_data, colWidths=[35, 270, 120, 120], style=[
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EAEFF5")),
            ('LINEBELOW', (0,0), (-1,0), 1, colors.HexColor("#1A2B4C")),
            ('LINEABOVE', (0,-2), (-1,-2), 1, colors.HexColor("#1A2B4C")),
            ('LINEABOVE', (0,-1), (-1,-1), 1.5, colors.HexColor("#9C0000")),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#FDF2F2")),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        elementos.append(Spacer(1, 8))
        elementos.append(Paragraph("Presupuesto válido por 30 días", ParagraphStyle('Pie', fontName='Helvetica-Bold', fontSize=7.5, alignment=1, textColor=colors.HexColor("#555555"))))

        doc.build(elementos)
        buffer.seek(0)
        return buffer.getvalue()

    # -------------------------------------------------------------
    # GENERADOR DE EXCEL FORMAL (SIN LINKS ROTOS)
    # -------------------------------------------------------------
    def generar_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Presupuesto Oficial"
        ws.views.sheetView[0].showGridLines = True
        azul = "1F497D"
        gris = "595959"
        fondo_total = "EBF1F5"

        borde = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        borde_doble_abajo = Border(top=Side(style='thin', color='1F497D'), bottom=Side(style='double', color='1F497D'))

        ws["A1"] = f"{nombre_taller.upper()} — PRESUPUESTO OFICIAL N° {nro_presupuesto}"
        ws["A1"].font = Font(size=13, bold=True, color=azul)
        ws["A2"] = f"Vehículo: {datos.get('vehiculo_detectado')} | Dominio: {datos.get('patente_detectada')} | Tipología: {datos.get('tipologia_siniestro')}"
        ws["A2"].font = Font(size=9.5, italic=True)

        fila = 4
        ws.cell(row=fila, column=1, value="I. DETALLE DE REPUESTOS RECLAMADOS").font = Font(bold=True, color=azul)
        fila += 1

        for c_idx, h in enumerate(["Ítem", "Rubro / Sistema", "Repuesto Reclamado", "Importe Unitario (ARS)", "Importe Total (ARS)"], 1):
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
            for col in range(1, 6):
                ws.cell(row=fila, column=col).border = borde
            fila += 1

        ws.cell(row=fila, column=4, value="SUBTOTAL REPUESTOS:").font = Font(bold=True, color=azul)
        ws.cell(row=fila, column=4).alignment = Alignment(horizontal="right")
        c_tot_rep = ws.cell(row=fila, column=5, value=float(total_repuestos))
        c_tot_rep.font = Font(bold=True, color=azul)
        c_tot_rep.number_format = "$#,##0.00"

        fila += 2
        ws.cell(row=fila, column=1, value="II. MANO DE OBRA Y SERVICIOS").font = Font(bold=True, color=azul)
        fila += 1

        for c_idx, h in enumerate(["Ítem", "Concepto de Tarea", "Detalle", "Subtotal (ARS)"], 1):
            c = ws.cell(row=fila, column=c_idx, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color=gris, end_color=gris, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")

        fila += 1
        tareas_excel = []
        if sub_chapa > 0:
            tareas_excel.append(("M.O. Chapa pesada y banco", f"{dias_chapa} días", sub_chapa))
        if sub_pintura > 0:
            tareas_excel.append(("M.O. Pintura en cabina", f"{panos} paños", sub_pintura))
        if sub_tren > 0:
            tareas_excel.append(("M.O. Reparación de Tren Delantero", "Desarme, reemplazo y montaje", sub_tren))
        if sub_motor > 0:
            tareas_excel.append(("M.O. Mecánica de Motor, Armado y Alineación", "Cárter y alineación", sub_motor))
        if sub_serv > 0:
            tareas_excel.append(("Kit Service Aceite Sintético 5W40 + Filtros", "Fluidos oficiales", sub_serv))

        for i, t in enumerate(tareas_excel, 1):
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
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 46
        ws.column_dimensions["D"].width = 24
        ws.column_dimensions["E"].width = 24

        buff = io.BytesIO()
        wb.save(buff)
        return buff.getvalue()

    col_pdf, col_xlsx = st.columns(2)
    with col_pdf:
        st.download_button(
            label="📄 Descargar Presupuesto Oficial en PDF (Listo para Enviar)",
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
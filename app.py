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

# ReportLab para la generación del PDF profesional
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image as RLImage

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
st.markdown('<div class="sub-pericial">Liquidación pericial con baremos de cabina, relato pericial y entrega formal en PDF y Excel</div>', unsafe_allow_html=True)

# -------------------------------------------------------------
# LLAMADA ROBUSTA CON REINTENTOS AUTOMÁTICOS
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
# 1. BARRA LATERAL: TALLER, CLIENTE Y BAREMOS
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
    nro_presupuesto = st.text_input("N° Presupuesto:", value="0002596")
    nombre_titular = st.text_input("Titular / Asegurado:", value="Sicoli Agostina Maria")
    dni_titular = st.text_input("DNI / CUIT:", value="36.872.370")
    domicilio_titular = st.text_input("Domicilio:", value="Neuquen 1970, CABA")
    tel_titular = st.text_input("Teléfono Titular:", value="")

    st.markdown("---")
    st.subheader("3. Baremos y Tarifas Base")
    tipo_vehiculo = st.radio("Tipo de Vehículo:", ["Automotor / Utilitario", "Motovehículo / Moto"], index=0)
    tarifa_pano = st.number_input("Valor por paño de pintura (ARS):", min_value=0, value=200000, step=5000)
    tarifa_dia_chapa = st.number_input("Valor día chapa pesada / banco (ARS):", min_value=0, value=110000, step=5000)
    tarifa_desarme = st.number_input("M.O. Armado y Desarmado / Tren Delantero (ARS):", min_value=0, value=180000, step=10000)
    tarifa_motor_alineacion = st.number_input("M.O. Mecánica / Cárter / Alineación (ARS):", min_value=0, value=240000, step=10000)
    tarifa_service_fluidos = st.number_input("Kit Service Aceite y Filtros (ARS):", min_value=0, value=316000, step=5000)

if "peritaje_listo" not in st.session_state:
    st.session_state.peritaje_listo = False
if "datos_peritaje" not in st.session_state:
    st.session_state.datos_peritaje = None
if "repuestos_cotizados" not in st.session_state:
    st.session_state.repuestos_cotizados = []

# -------------------------------------------------------------
# 2. CARGA DE DOCUMENTOS Y RELATO DEL HECHO
# -------------------------------------------------------------
archivos_subidos = st.file_uploader(
    "Subí fotos del siniestro (JPG/PNG) o carpetas periciales en PDF (podés seleccionar múltiples archivos):",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

if archivos_subidos:
    st.success(f"{len(archivos_subidos)} archivo(s) seleccionado(s).")
    imgs_para_mostrar = [a for a in archivos_subidos if not a.name.lower().endswith(".pdf")]
    if imgs_para_mostrar:
        cols = st.columns(min(len(imgs_para_mostrar), 5))
        for i, arch in enumerate(imgs_para_mostrar):
            with cols[i % 5]:
                st.image(arch.getvalue(), caption=f"Foto {i+1}", use_container_width=True)

relato_siniestro = st.text_area(
    "Relato del hecho / Observaciones del perito (daños ocultos, mecánica o aclaraciones de taller):",
    placeholder="Ej: Encerrona en avenida: raspón corrido de punta a punta afectando paragolpes delantero, espejo izquierdo, guardabarros trasero y paragolpes trasero con ojo de gato. El guardabarros trasero se repara en chapa y pintura, no se cambia..."
)

boton_iniciar = st.button(
    "Iniciar Peritaje Integral y Liquidación Oficial",
    type="primary",
    use_container_width=True,
    disabled=(not archivos_subidos)
)

# -------------------------------------------------------------
# 3. MOTOR PERICIAL INTELIGENTE CON REGLAS DE BAREMO
# -------------------------------------------------------------
if boton_iniciar and archivos_subidos:
    client = genai.Client()
    contenidos_gemini = []

    for a in archivos_subidos:
        bytes_data = a.getvalue()
        if not bytes_data:
            continue

        if a.name.lower().endswith(".pdf"):
            contenidos_gemini.append(types.Part.from_bytes(data=bytes_data, mime_type="application/pdf"))
        else:
            try:
                im = Image.open(io.BytesIO(bytes_data)).convert("RGB")
                im.thumbnail((1024, 1024))
                buf = io.BytesIO()
                im.save(buf, format="JPEG")
                contenidos_gemini.append(types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"))
            except Exception as e_img:
                st.warning(f"No se pudo decodificar la imagen {a.name}: {e_img}")

    with st.spinner("Analizando imágenes, cotejando relato del hecho y computando baremos periciales..."):
        prompt = f"""
        Sos un perito liquidador senior de siniestros viales en Argentina y jefe técnico de taller.
        Analizá el material gráfico subido en conjunto con este relato técnico del taller:

        RELATO DEL HECHO / OBSERVACIONES:
        "{relato_siniestro if relato_siniestro.strip() else 'Sin relato provisto. Deducir puramente de las imágenes.'}"

        TIPO DE VEHÍCULO: {tipo_vehiculo}.

        REGLAS PERICIALES Y BAREMOS HOMOLOGADOS (CESVI / CONCESIONARIOS OEM):

        1. IDENTIFICACIÓN VEHICULAR:
           - Marca, Modelo exacto, Generación/Año y Patente (visible o deducida).

        2. REGLA ESTRUCTURAL DEL GUARDABARROS TRASERO (MONOCASCO):
           - El guardabarros trasero / lateral es pieza soldada al monocasco. Salvo destrucción total irreparable, NO SE PONE REPUESTO DE CHAPA: se liquida como M.O. CHAPA PESADA (días de estiramiento en banco) + PAÑOS DE PINTURA.

        3. BAREMO TÉCNICO DE PINTURA (MÁXIMO PERICIAL):
           - Un automóvil completo entero tiene entre 14 y 16 paños en total.
           - Si el siniestro es un ROCE LATERAL CORRIDO (roce de trompa a cola que afecta paragolpes delantero, espejo, puertas, lateral soldado, zócalo y paragolpes trasero):
             * 'panos_pintura': Asignar exactamente 11.0 paños (cubre las superficies rozadas más los difuminados técnicos indispensables de cabina). NUNCA superar 11.0 paños por un solo costado.
           - Si es choque frontal común: entre 5.0 y 7.0 paños.
           - Si es impacto trasero puro: entre 3.0 y 4.5 paños.
           - Si es pozo vial: 0.0 paños.

        4. MATRIZ DE REPUESTOS TÉCNICOS (PRECIOS DE REPOSICIÓN ORIGINAL OEM EN ARS):
           - Si hay roce corrido delantero-lateral-trasero:
             * Paragolpes trasero original con alojamiento de radar/sensores si corresponde.
             * Faro trasero exterior del lado afectado (inspeccionar anclajes y fisuras).
             * Paragolpes delantero original.
             * Moldura cromada / rejilla delantera.
             * Espejo retrovisor exterior completo.
             * Ojo de gato / reflectante de paragolpes trasero.
           - Si es pozo vial:
             * Llanta original, neumático nuevo, parrilla de suspensión, kit amortiguadores par, rótula, extremos, cárter de aluminio y activar 'incluye_service_aceite': true. Chapa y pintura en $0.
           - Si es choque frontal:
             * Capot, paragolpes, ópticas, panel porta radiador, alma, radiador, patas de motor/caja, semiejes.
           - Si es motovehículo:
             * Horquilla, barrales, cristo, manillar, manetas, pedalines, escape, plásticos/carenados. Chapa=0, Pintura plásticos (2-4 paños).

        Respondé ÚNICAMENTE en JSON con esta estructura exacta:
        {{
          "vehiculo_detectado": "Renault Sandero PH2 Zen 1.6",
          "patente_detectada": "AE 941 XW",
          "fuente_identificacion": "Peritaje visual de carrocería y chapa patente",
          "tipologia_siniestro": "ROCE_LATERAL_CORRIDO",
          "diagnostico_cinematica": "Dictamen pericial técnico detallado...",
          "panos_pintura": 11.0,
          "dias_chapa_banco": 5.0,
          "requiere_arme_desarme": true,
          "requiere_mecanica_motor": false,
          "incluye_service_aceite": false,
          "repuestos": [
            {{"categoria": "Carrocería Exterior", "pieza": "Paragolpes Trasero (c/ alojamiento radar/sensores) Original", "precio_oem": 1147340.0}},
            {{"categoria": "Iluminación", "pieza": "Faro Trasero Exterior Izquierdo Original", "precio_oem": 659000.0}},
            {{"categoria": "Carrocería Exterior", "pieza": "Paragolpes Delantero Original", "precio_oem": 555980.0}},
            {{"categoria": "Carrocería Exterior", "pieza": "Moldura Cromada Parrilla Delantera Original", "precio_oem": 405700.0}},
            {{"categoria": "Carrocería Exterior", "pieza": "Espejo Retrovisor Izquierdo Completo", "precio_oem": 230000.0}},
            {{"categoria": "Iluminación / Accesorios", "pieza": "Ojo de Gato Paragolpes Trasero Izquierdo", "precio_oem": 114030.0}}
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
# 4. RESULTADOS, EDICIÓN Y EXPORTACIÓN DUAL
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
        dias_chapa = st.number_input("Días Chapa Pesada / Banco:", value=float(datos.get("dias_chapa_banco", 0.0)), step=0.5)
        sub_chapa = dias_chapa * tarifa_dia_chapa
        st.write(f"Chapa: **${sub_chapa:,.2f}**")

    with c3:
        req_desarme = datos.get("requiere_arme_desarme", True)
        aplica_desarme = st.checkbox("M.O. Armado y Desarmado", value=req_desarme)
        sub_desarme = tarifa_desarme if aplica_desarme else 0.0

        req_motor = datos.get("requiere_mecanica_motor", False)
        aplica_motor = st.checkbox("M.O. Mecánica / Tren Delantero", value=req_motor)
        sub_motor = tarifa_motor_alineacion if aplica_motor else 0.0

        st.write(f"Arme/Desarme: **${sub_desarme:,.2f}**")
        if aplica_motor:
            st.write(f"Mecánica: **${sub_motor:,.2f}**")

    with c4:
        inc_service = datos.get("incluye_service_aceite", False)
        aplica_service = st.checkbox("Kit Service Aceite y Filtros", value=inc_service)
        sub_serv = tarifa_service_fluidos if aplica_service else 0.0
        if aplica_service:
            st.write(f"Service: **${sub_serv:,.2f}**")

    total_mo = sub_pintura + sub_chapa + sub_desarme + sub_motor + sub_serv

    # REPUESTOS LIMPIOS Y EDICIÓN RÁPIDA
    st.subheader(f"2. Detalle de Repuestos Reclamados ({len(st.session_state.repuestos_cotizados)} ítems)")
    piezas_finales = []
    for idx, r in enumerate(st.session_state.repuestos_cotizados):
        col_c, col_p = st.columns([7, 3])
        with col_c:
            activa = st.checkbox(f"**{r['pieza']}**", value=True, key=f"r_{idx}")
        with col_p:
            pr = st.number_input("Precio ARS", value=float(r["precio"]), step=5000.0, key=f"pr_{idx}", label_visibility="collapsed")

        if activa:
            piezas_finales.append({
                "categoria": r.get("categoria", "Repuestos"),
                "pieza": r["pieza"],
                "precio": pr
            })

    # FORMULARIO MANUAL PARA SUMAR PIEZAS INTERNAS
    with st.expander("➕ Agregar repuesto manualmente (daño oculto o traba partida)"):
        col_nom, col_pre, col_btn = st.columns([5, 3, 2])
        with col_nom:
            nuevo_nombre = st.text_input("Descripción de la pieza:", key="input_nuevo_nombre", placeholder="Ej: Faro Trasero Exterior Izquierdo Original")
        with col_pre:
            nuevo_precio = st.number_input("Precio oficial ARS:", min_value=0.0, step=5000.0, value=650000.0, key="input_nuevo_precio")
        with col_btn:
            st.write("")
            st.write("")
            if st.button("Agregar a la lista", use_container_width=True):
                if nuevo_nombre.strip():
                    st.session_state.repuestos_cotizados.append({
                        "categoria": "Repuestos Adicionales",
                        "pieza": nuevo_nombre.strip(),
                        "precio": float(nuevo_precio)
                    })
                    st.rerun()

    total_repuestos = sum(p["precio"] for p in piezas_finales)
    total_general = total_repuestos + total_mo

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Subtotal Mano de Obra", f"${total_mo:,.2f}")
    m2.metric("Subtotal Repuestos", f"${total_repuestos:,.2f}")
    m3.metric("TOTAL PRESUPUESTADO", f"${total_general:,.2f}")

    # -------------------------------------------------------------
    # GENERADOR DE PDF COMERCIAL LIMPIO
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
            [Paragraph(f"<b>Domicilio:</b> {domicilio_titular}", style_celda), Paragraph("<b>Localidad:</b> CABA, Buenos Aires", style_celda)]
        ]
        elementos.append(Table(info_cliente, colWidths=[270, 275], style=[
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#B0C0D0")),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        elementos.append(Spacer(1, 8))

        elementos.append(Paragraph("<b>Detalles de tareas (Mano de Obra)</b>", style_sec_title))
        elementos.append(Spacer(1, 3))
        tareas_data = [
            [Paragraph("<b>Descripción de la Tarea</b>", style_celda_bold), Paragraph("<b>Detalle</b>", style_celda_bold), Paragraph("<b>Subtotal</b>", style_celda_num_b)]
        ]
        if sub_chapa > 0:
            tareas_data.append([Paragraph("M.O. CHAPA PESADA", style_celda), Paragraph(f"{dias_chapa} días de chapa / estiramiento", style_celda), Paragraph(f"$ {sub_chapa:,.2f}", style_celda_num)])
        if sub_pintura > 0:
            tareas_data.append([Paragraph("M.O. PINTURA", style_celda), Paragraph(f"{panos} paños de pintura en cabina", style_celda), Paragraph(f"$ {sub_pintura:,.2f}", style_celda_num)])
        if sub_desarme > 0:
            tareas_data.append([Paragraph("ARMADO Y DESARMADO", style_celda), Paragraph("Desarme y ensamble de carrocería", style_celda), Paragraph(f"$ {sub_desarme:,.2f}", style_celda_num)])
        if sub_motor > 0:
            tareas_data.append([Paragraph("M.O. MECÁNICA Y TREN DELANTERO", style_celda), Paragraph("Mecánica y alineación computarizada", style_celda), Paragraph(f"$ {sub_motor:,.2f}", style_celda_num)])
        if sub_serv > 0:
            tareas_data.append([Paragraph("KIT SERVICE DE ACEITE Y FILTROS", style_celda), Paragraph("Fluidos oficiales y filtrado completo", style_celda), Paragraph(f"$ {sub_serv:,.2f}", style_celda_num)])

        tareas_data.append([Paragraph("<b>TOTAL MANO DE OBRA</b>", style_celda_bold), Paragraph("", style_celda), Paragraph(f"<b>$ {total_mo:,.2f}</b>", style_celda_num_b)])

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
    # GENERADOR DE PLANILLA EXCEL (.XLSX)
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
            tareas_excel.append(("M.O. Chapa pesada", f"{dias_chapa} días", sub_chapa))
        if sub_pintura > 0:
            tareas_excel.append(("M.O. Pintura", f"{panos} paños", sub_pintura))
        if sub_desarme > 0:
            tareas_excel.append(("Armado y desarmado", "Carrocería", sub_desarme))
        if sub_motor > 0:
            tareas_excel.append(("M.O. Mecánica y tren delantero", "Alineación", sub_motor))
        if sub_serv > 0:
            tareas_excel.append(("Kit Service Aceite Sintético + Filtros", "Fluidos", sub_serv))

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
            label="📄 Descargar Presupuesto Oficial en PDF",
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
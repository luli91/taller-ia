import io
import re
import json
import time
import streamlit as st
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv
from google import genai
from google.genai import types

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image as RLImage

load_dotenv()

st.set_page_config(
    page_title="Taller IA — Sistema Pericial Completo por Zona de Impacto",
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

st.markdown('<div class="titulo-pericial">Taller IA — Peritaje Integral con Criterio de Taller</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-pericial">Detección Exhaustiva por Zona de Siniestro | Precios Máximos en Vivo (Mercado Libre / OEM) | Destildado Rápido</div>', unsafe_allow_html=True)

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

def extraer_json_seguro(texto):
    texto_limpio = texto.strip()
    if texto_limpio.startswith("```"):
        texto_limpio = re.sub(r"^```(?:json)?", "", texto_limpio)
        texto_limpio = re.sub(r"```$", "", texto_limpio).strip()
    try:
        return json.loads(texto_limpio)
    except Exception:
        match = re.search(r"\{.*\}", texto_limpio, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise ValueError("No se pudo decodificar la respuesta JSON del motor.")

# -------------------------------------------------------------
# 1. BARRA LATERAL: DATOS DEL TALLER Y TARIFAS BASE
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
    tipo_vehiculo = st.radio("Categoría de Vehículo:", ["Automotor / Pick-up / Utilitario", "Motovehículo / Moto"], index=0)
    es_moto = (tipo_vehiculo == "Motovehículo / Moto")

    nro_presupuesto = st.text_input("N° Presupuesto:", value="0002605")
    nombre_titular = st.text_input("Titular / Asegurado:", value="Cliente / Asegurado")
    dni_titular = st.text_input("DNI / CUIT:", value="")
    domicilio_titular = st.text_input("Domicilio:", value="")
    tel_titular = st.text_input("Teléfono Titular:", value="")

    st.markdown("---")
    st.subheader("3. Tarifas de Mano de Obra del Taller")
    if es_moto:
        tarifa_mo_moto = st.number_input("M.O. Mecánica, Alineación y Armado Moto (ARS):", min_value=0, value=190000, step=10000)
        tarifa_pano = tarifa_dia_chapa = tarifa_desarme = tarifa_mecanica_elec = tarifa_materiales = tarifa_tren_delantero = tarifa_motor_alineacion = tarifa_service_fluidos = 0
    else:
        tarifa_mo_moto = 0
        tarifa_pano = st.number_input("Valor por Paño de Pintura (ARS):", min_value=0, value=200000, step=5000)
        tarifa_dia_chapa = st.number_input("Valor por Día de Chapa (ARS):", min_value=0, value=180000, step=5000)
        tarifa_desarme = st.number_input("M.O. Arme y Desarme (ARS):", min_value=0, value=250000, step=10000)
        tarifa_mecanica_elec = st.number_input("M.O. Mecánica y Electricidad (ARS):", min_value=0, value=150000, step=10000)
        tarifa_materiales = st.number_input("Materiales de Pintura y Selladores (ARS):", min_value=0, value=250000, step=10000)
        tarifa_tren_delantero = st.number_input("M.O. Tren Delantero (ARS):", min_value=0, value=300000, step=10000)
        tarifa_motor_alineacion = st.number_input("M.O. Motor y Alineación (ARS):", min_value=0, value=240000, step=10000)
        tarifa_service_fluidos = st.number_input("Kit Service Aceite y Filtros (ARS):", min_value=0, value=316000, step=5000)

if "peritaje_listo" not in st.session_state:
    st.session_state.peritaje_listo = False
if "datos_peritaje" not in st.session_state:
    st.session_state.datos_peritaje = None
if "repuestos_cotizados" not in st.session_state:
    st.session_state.repuestos_cotizados = []

# -------------------------------------------------------------
# 2. CARGA DE FOTOS Y DOCUMENTOS
# -------------------------------------------------------------
archivos_subidos = st.file_uploader(
    "Subí todas las fotos del siniestro (JPG/PNG) o carpetas en PDF:",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

if archivos_subidos:
    st.success(f"{len(archivos_subidos)} archivo(s) cargado(s).")
    imgs_para_mostrar = [a for a in archivos_subidos if not a.name.lower().endswith(".pdf")]
    if imgs_para_mostrar:
        cols = st.columns(min(len(imgs_para_mostrar), 6))
        for i, arch in enumerate(imgs_para_mostrar):
            with cols[i % 6]:
                st.image(arch.getvalue(), caption=f"Foto {i+1}", use_container_width=True)

relato_siniestro = st.text_area(
    "Observaciones del taller o daños ocultos (opcional):",
    placeholder="Podés dejarlo vacío: el sistema detectará todo el despiece de la zona afectada y podrás destildar lo que no quieras incluir...",
    height=68
)

boton_iniciar = st.button(
    "🔍 Analizar Zona del Siniestro y Cotizar Repuestos al Precio Más Alto",
    type="primary",
    use_container_width=True,
    disabled=(not archivos_subidos)
)

# -------------------------------------------------------------
# 3. MOTOR EN 2 ETAPAS: EXHAUSTIVO POR ZONA + PRECIOS MÁXIMOS ML
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
                im.thumbnail((1280, 1280))
                buf = io.BytesIO()
                im.save(buf, format="JPEG", quality=90)
                contenidos_gemini.append(types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"))
            except Exception as e_img:
                st.warning(f"No se pudo leer {a.name}: {e_img}")

    # =========================================================================
    # ETAPA 1: PERITAJE CON EXPERIENCIA DE CHAPISTA Y MECÁNICO (20 AÑOS)
    # =========================================================================
    with st.spinner("Etapa 1/2: Detectando todas las piezas visibles e internas de la zona del impacto..."):
        prompt_etapa_1 = f"""
        Actuás como un maestro chapista, mecánico y perito liquidador con 20 años de experiencia en Argentina.
        Analizá las fotos y documentos subidos junto con este comentario: "{relato_siniestro}".
        CATEGORÍA DE VEHÍCULO: {tipo_vehiculo}.

        ================================================================================
        FILOSOFÍA DEL SISTEMA (REGLA DE ORO):
        "Es mejor incluir TODAS las piezas afectadas y conectadas a la zona exacta del siniestro (para que el tallerista pueda destildar lo que no quiera), que olvidarse de un repuesto".
        PERO ATENCIÓN: SOLO debés incluir piezas que pertenezcan FÍSICAMENTE A LA ZONA DEL GOLPE. Prohibido inventar piezas de lados o sectores que no sufrieron impacto, para que el seguro jamás rebote el presupuesto.
        ================================================================================

        REGLAS DE DETECCIÓN EXHAUSTIVA POR ZONA DE IMPACTO:

        1. ÓPTICAS Y FAROS CONTIGUOS AL GOLPE (INCLUIRLOS SIEMPRE):
           - Si hay un choque, hundimiento o descalce en el paragolpes (delantero o trasero) o guardabarros de un lado (ej. zona izquierda), AUNQUE EL ACRÍLICO DE LA ÓPTICA O FARO DE ESE LADO SE VEA SANO POR FUERA, DEBÉS INCLUIR ESA ÓPTICA / FARO EN LA LISTA DE REPUESTOS.
           - Motivo técnico: El empuje del paragolpes quiebra las patitas/orejas plásticas de fijación interna de la óptica o faro contiguo.

        2. DESPIECE COMPLETO SEGÚN LA ZONA AFECTADA (AUTOS Y UTILITARIOS):
           - SI EL GOLPE AFECTA PARAGOLPES DELANTERO (esquina o centro):
             * 'Paragolpes Delantero Original'
             * 'Soporte Guía de Paragolpes Delantero (Izquierdo y/o Derecho según zona del golpe)'
             * 'Óptica Delantera (Izquierda y/o Derecha según zona del golpe)' (por impacto directo o rotura de anclajes internos)
             * 'Marco / Moldura de Faro Auxiliar (Izquierdo y/o Derecho)' y 'Faro Auxiliar' si el golpe es en la parte baja de esa esquina.
             * 'Parrilla Central / Molduras Cromadas de Parrilla' si el golpe llega hacia el centro o capot.
             * 'Alma de Acero de Paragolpes Delantero' y 'Frente Porta Radiadores' si el impacto es frontal o profundo.
           - SI EL GOLPE AFECTA GUARDABARROS DELANTERO:
             * 'Guardabarros Delantero (Izquierdo o Derecho) Original'
             * 'Guardaplast / Pasarruedas Interno Delantero (Izquierdo o Derecho)'
           - SI EL GOLPE AFECTA PUERTAS / LATERAL:
             * Si la puerta tiene abolladura, pliegue o raspón fuerte -> 'Puerta Delantera Original' y/o 'Puerta Trasera Original'.
             * 'Moldura / Bagueta de Puerta Delantera y/o Trasera' (obligatorio porque no se reutilizan).
             * 'Espejo Retrovisor Completo' o 'Cacha de Espejo Retrovisor' si el roce llegó adelante.
             * 'Zócalo Lateral de Chapa' si el golpe afecta la zona baja.
             * Guardabarros trasero (panel soldado): NO va como repuesto de chapa, va a Días de Chapa + Paños de Pintura.
           - SI EL GOLPE AFECTA COLA / BAÚL / PARAGOLPES TRASERO:
             * 'Paragolpes Trasero Original'
             * 'Soporte Guía de Paragolpes Trasero (Izquierdo y/o Derecho)'
             * 'Faro Trasero Exterior y/o Interior de Baúl (del lado afectado o juego)'
             * 'Ojo de Gato / Reflectante de Paragolpes Trasero' (si lleva en esa zona)
             * Si el golpe afectó Tapa de Baúl / Portón -> 'Tapa de Baúl / Portón Trasero Original', 'Panel de Cola Trasero Chapa Original', 'Cerradura / Actuador Eléctrico de Baúl Original' y 'Moldura / Aplique Portapatente Trasero Original'.
           - SI ES POZO VIAL / BACHE / TREN DELANTERO:
             * 'Llanta de Aleación Original', 'Neumático Nuevo', 'Parrilla de Suspensión', 'Rótula', 'Kit x2 Amortiguadores Delanteros', 'Extremos y Axiales de Dirección', 'Cárter de Motor'.

        3. DESPIECE COMPLETO PARA MOTOS (CHAPA = 0, PINTURA = 0, TODO REPUESTO NUEVO OEM):
           - Incluir TODOS los plásticos/carenados rozados o golpeados ('Cacha Cubre Tanque Lateral', 'Toma de Aire / Tobera', 'Guardabarros Delantero', 'Protector Cubre Silenciador de Escape', 'Emblema / Gráfica 3D Original', 'Colín') + TODOS los componentes de dirección y apoyo de ese lado ('Manubrio de Dirección Original', 'Bomba de Freno Delantero', 'Manija de Freno/Embrague', 'Espejo Retrovisor', 'Contrapeso de Manillar', 'Faro de Giro', 'Pedalín de Apoyo / Palanca').

        4. CÁLCULO CERTERO DE MANO DE OBRA:
           - Contá los paños de pintura reales según las piezas involucradas (ej. golpe esquinero de Paragolpes + 1 Guardabarros = 2.0 a 3.0 paños y 2.0 días de chapa; lateral con puertas o cola completa = 8.0 paños y 3.0 días de chapa; roce corrido de punta a punta = 11.0 paños y 4.0 a 5.0 días de chapa).

        En cada repuesto incluí el campo "tipo_deteccion" indicando:
        - "🔴 Daño directo visible en fotos"
        - "🟡 Por proximidad de impacto (anclajes/soportes internos — destildar si está sano)"

        Respondé ÚNICAMENTE en JSON válido con esta estructura:
        {{
          "vehiculo_detectado": "Marca Modelo Versión (Ej: Chevrolet Cruze 1.8 LTZ)",
          "patente_detectada": "MAW 684",
          "titular_detectado": "",
          "dni_detectado": "",
          "fuente_identificacion": "Inspección de carrocería, insignias y patente",
          "tipologia_siniestro": "IMPACTO_DELANTERO_IZQUIERDO",
          "diagnostico_cinematica": "Informe técnico pericial detallado...",
          "panos_pintura": 2.0,
          "dias_chapa_banco": 2.0,
          "requiere_arme_desarme": false,
          "requiere_mecanica_electricidad": false,
          "incluye_materiales_pintura": false,
          "requiere_tren_delantero": false,
          "requiere_motor_alineacion": false,
          "incluye_service_aceite": false,
          "requiere_mecanica_moto": false,
          "repuestos": [
            {{
              "categoria": "Carrocería",
              "pieza": "Paragolpes Delantero Original",
              "tipo_deteccion": "🔴 Daño directo visible en fotos",
              "terminos_busqueda_ml": "Paragolpes Delantero Original Chevrolet Cruze 1.8",
              "precio_estimado_respaldo": 2880000.0
            }}
          ]
        }}
        """

        try:
            resp_1 = llamar_gemini_con_reintentos(
                client=client,
                model='gemini-3.8-flash',
                contents=[prompt_etapa_1] + contenidos_gemini,
                config=types.GenerateContentConfig(response_mime_type="application/json"),
                max_reintentos=4
            )
            datos_etapa_1 = extraer_json_seguro(resp_1.text)
        except Exception as e:
            st.error(f"Error en la Etapa 1 (Análisis Visual): {e}")
            st.stop()

    # =========================================================================
    # ETAPA 2: BÚSQUEDA DEL PRECIO ORIGINAL MÁS CARO EN MERCADO LIBRE / PLAZA
    # =========================================================================
    vehiculo_str = datos_etapa_1.get("vehiculo_detectado", "Vehículo")
    lista_piezas_etapa_1 = datos_etapa_1.get("repuestos", [])

    with st.spinner(f"Etapa 2/2: Buscando los precios originales más altos vigentes en Argentina (Mercado Libre / Concesionarios) para {vehiculo_str}..."):
        prompt_etapa_2 = f"""
        Sos un auditor de precios de repuestos de seguros en Argentina (Septiembre 2026).
        Usá Google Search para consultar los precios actuales en Pesos Argentinos (ARS) en Mercado Libre Argentina y concesionarias oficiales para:

        VEHÍCULO: {vehiculo_str}

        REGLAS ESTRICTAS DE PRECIOS PARA NO QUEDAR CORTO:
        1. Buscá el precio de cada repuesto NUEVO Y ORIGINAL LEGÍTIMO (OEM). Descartá repuestos alternativos chinos baratos o usados.
        2. ELEGÍ SIEMPRE EL PRECIO ORIGINAL MÁS CARO VIGENTE EN PLAZA (el techo de mercado en Mercado Libre Argentina / Concesionario Oficial) para que el tallerista y el cliente estén 100% cubiertos ante aumentos de precios o inflación.
        3. Respetá la gama del vehículo: en vehículos como Chevrolet Cruze 1.8 (coreano), Honda Civic, Toyota Corolla, VW Suran/Vento o Pick-ups, los paragolpes, puertas, capots y ópticas originales tienen valores altos de mostrador oficial.

        LISTA DE REPUESTOS A COTIZAR:
        {json.dumps(lista_piezas_etapa_1, ensure_ascii=False, indent=2)}

        Respondé ÚNICAMENTE con un objeto JSON válido con este formato exacto:
        {{
          "repuestos_actualizados": [
            {{
              "categoria": "Categoría",
              "pieza": "Nombre formal del repuesto",
              "tipo_deteccion": "🔴 Daño directo visible en fotos",
              "precio_oem": 2883000.0
            }}
          ]
        }}
        """

        repuestos_finales_ia = []
        try:
            resp_2 = llamar_gemini_con_reintentos(
                client=client,
                model='gemini-3.8-flash',
                contents=prompt_etapa_2,
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())]
                ),
                max_reintentos=3
            )
            datos_etapa_2 = extraer_json_seguro(resp_2.text)
            repuestos_finales_ia = datos_etapa_2.get("repuestos_actualizados", [])
        except Exception:
            repuestos_finales_ia = [
                {
                    "categoria": r.get("categoria", "Repuestos"),
                    "pieza": r.get("pieza", "Pieza"),
                    "tipo_deteccion": r.get("tipo_deteccion", "🔴 Relacionado al siniestro"),
                    "precio_oem": float(r.get("precio_estimado_respaldo", 250000.0))
                }
                for r in lista_piezas_etapa_1
            ]

        st.session_state.datos_peritaje = datos_etapa_1
        st.session_state.repuestos_cotizados = [
            {
                "categoria": item.get("categoria", "Repuestos"),
                "pieza": item["pieza"],
                "tipo_deteccion": item.get("tipo_deteccion", "🔴 Relacionado al siniestro"),
                "precio": float(item.get("precio_oem", item.get("precio_estimado_respaldo", 150000.0)))
            }
            for item in repuestos_finales_ia
        ]
        st.session_state.peritaje_listo = True
        st.rerun()

# -------------------------------------------------------------
# 4. PANTALLA DE CONTROL DEL TALLERISTA (DESTILDAR EN 1 CLIC)
# -------------------------------------------------------------
if st.session_state.peritaje_listo and st.session_state.datos_peritaje:
    datos = st.session_state.datos_peritaje

    st.markdown(f"""
        <div class="card-vehiculo">
            <h4 style="margin:0; color:#1F497D;">Vehículo: {datos.get('vehiculo_detectado')} — Dominio: {datos.get('patente_detectada')}</h4>
            <p style="margin:4px 0 0 0; font-size:13px; color:#333;">
                <strong>Zona / Tipología Detectada:</strong> {datos.get('tipologia_siniestro')} | 
                <strong>Fuente:</strong> {datos.get('fuente_identificacion')}
            </p>
        </div>
    """, unsafe_allow_html=True)

    with st.expander("📋 Ver Dictamen Técnico Pericial", expanded=False):
        st.write(datos.get("diagnostico_cinematica"))

    st.subheader("1. Mano de Obra y Tareas (Podés tildar o destildar lo que quieras cobrar)")
    if es_moto:
        req_moto = datos.get("requiere_mecanica_moto", True)
        aplica_moto = st.checkbox("M.O. Mecánica, Alineación y Armado", value=req_moto)
        total_mo = tarifa_mo_moto if aplica_moto else 0.0
        st.write(f"Subtotal M.O. Moto: **${total_mo:,.2f}**")
        sub_pintura = sub_chapa = sub_desarme = sub_mec_elec = sub_materiales = sub_tren = sub_motor = sub_serv = 0.0
        panos = dias_chapa = 0.0
    else:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            panos = st.number_input("Paños de Pintura:", value=float(datos.get("panos_pintura", 2.0)), step=0.5)
            sub_pintura = panos * tarifa_pano
            st.write(f"Pintura: **${sub_pintura:,.2f}**")

            inc_mat = datos.get("incluye_materiales_pintura", False)
            aplica_materiales = st.checkbox("Materiales Pintura y Selladores", value=inc_mat)
            sub_materiales = tarifa_materiales if aplica_materiales else 0.0
            if aplica_materiales:
                st.write(f"Materiales: **${sub_materiales:,.2f}**")

        with c2:
            dias_chapa = st.number_input("Días de Chapa:", value=float(datos.get("dias_chapa_banco", 2.0)), step=0.5)
            sub_chapa = round(dias_chapa * tarifa_dia_chapa, -3) if dias_chapa > 0 else 0.0
            st.write(f"Chapa: **${sub_chapa:,.2f}**")

        with c3:
            req_desarme = datos.get("requiere_arme_desarme", False)
            aplica_desarme = st.checkbox("M.O. Arme y Desarme", value=req_desarme)
            sub_desarme = tarifa_desarme if aplica_desarme else 0.0

            req_mec_elec = datos.get("requiere_mecanica_electricidad", False)
            aplica_mec_elec = st.checkbox("M.O. Mecánica y Electricidad", value=req_mec_elec)
            sub_mec_elec = tarifa_mecanica_elec if aplica_mec_elec else 0.0

            req_tren = datos.get("requiere_tren_delantero", False)
            aplica_tren = st.checkbox("M.O. Tren Delantero", value=req_tren)
            sub_tren = tarifa_tren_delantero if aplica_tren else 0.0

            if aplica_desarme:
                st.write(f"Arme/Desarme: **${sub_desarme:,.2f}**")
            if aplica_mec_elec:
                st.write(f"Mec. y Elec.: **${sub_mec_elec:,.2f}**")
            if aplica_tren:
                st.write(f"Tren Del.: **${sub_tren:,.2f}**")

        with c4:
            req_motor = datos.get("requiere_motor_alineacion", False)
            aplica_motor = st.checkbox("M.O. Motor y Alineación", value=req_motor)
            sub_motor = tarifa_motor_alineacion if aplica_motor else 0.0

            inc_service = datos.get("incluye_service_aceite", False)
            aplica_service = st.checkbox("Kit Service Aceite y Filtros", value=inc_service)
            sub_serv = tarifa_service_fluidos if aplica_service else 0.0

            if aplica_motor:
                st.write(f"Motor/Alin.: **${sub_motor:,.2f}**")
            if aplica_service:
                st.write(f"Service: **${sub_serv:,.2f}**")

        total_mo = sub_pintura + sub_chapa + sub_desarme + sub_mec_elec + sub_materiales + sub_tren + sub_motor + sub_serv

    st.subheader(f"2. Repuestos de la Zona del Impacto ({len(st.session_state.repuestos_cotizados)} detectados — Destildá lo que no quieras incluir)")
    st.caption("💡 El sistema incluye todas las piezas visibles + las piezas conectadas de esa zona (como ópticas contiguas o soportes internos). Destildá en 1 clic lo que no corresponda.")

    piezas_finales = []
    for idx, r in enumerate(st.session_state.repuestos_cotizados):
        col_c, col_p = st.columns([7, 3])
        with col_c:
            activa = st.checkbox(
                f"**{r['pieza']}**  \n*{r.get('tipo_deteccion', '🔴 Relacionado al siniestro')}*",
                value=True,
                key=f"r_{idx}"
            )
        with col_p:
            pr = st.number_input("Precio ARS", value=float(r["precio"]), step=5000.0, key=f"pr_{idx}", label_visibility="collapsed")

        if activa:
            piezas_finales.append({
                "categoria": r.get("categoria", "Repuestos"),
                "pieza": r["pieza"],
                "precio": pr
            })

    with st.expander("➕ Agregar otro repuesto manualmente"):
        col_nom, col_pre, col_btn = st.columns([5, 3, 2])
        with col_nom:
            nuevo_nombre = st.text_input("Descripción de la pieza:", key="input_nuevo_nombre", placeholder="Ej: Guardaplast Delantero Izquierdo")
        with col_pre:
            nuevo_precio = st.number_input("Precio oficial ARS:", min_value=0.0, step=5000.0, value=95000.0, key="input_nuevo_precio")
        with col_btn:
            st.write("")
            st.write("")
            if st.button("Agregar a la lista", use_container_width=True):
                if nuevo_nombre.strip():
                    st.session_state.repuestos_cotizados.append({
                        "categoria": "Repuestos Adicionales",
                        "pieza": nuevo_nombre.strip(),
                        "tipo_deteccion": "🟢 Agregado manualmente por el taller",
                        "precio": float(nuevo_precio)
                    })
                    st.rerun()

    total_repuestos = sum(p["precio"] for p in piezas_finales)
    total_general = total_repuestos + total_mo

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Subtotal Mano de Obra", f"${total_mo:,.2f}")
    m2.metric(f"Subtotal Repuestos ({len(piezas_finales)} activos)", f"${total_repuestos:,.2f}")
    m3.metric("TOTAL PRESUPUESTADO", f"${total_general:,.2f}")

    # -------------------------------------------------------------
    # GENERACIÓN DE PDF OFICIAL
    # -------------------------------------------------------------
    def generar_pdf():
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
        elementos = []

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

        fecha_str = time.strftime("%d/%m/%Y")
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

        vehiculo_str_pdf = f"{datos.get('vehiculo_detectado')} - {datos.get('patente_detectada')}"
        titular_final = datos.get("titular_detectado") or nombre_titular
        dni_final = datos.get("dni_detectado") or dni_titular

        info_cliente = [
            [Paragraph(f"<b>Vehículo:</b> {vehiculo_str_pdf}", style_celda_bold), Paragraph(f"<b>Titular:</b> {titular_final}", style_celda)],
            [Paragraph(f"<b>DNI / CUIT:</b> {dni_final if dni_final else 'S/D'}", style_celda), Paragraph(f"<b>Teléfono:</b> {tel_titular if tel_titular else 'S/D'}", style_celda)],
            [Paragraph(f"<b>Domicilio:</b> {domicilio_titular if domicilio_titular else 'S/D'}", style_celda), Paragraph("<b>Localidad:</b> Buenos Aires", style_celda)]
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
        if es_moto:
            tareas_data.append([Paragraph("M.O. MECÁNICA, ALINEACIÓN Y ARMADO", style_celda), Paragraph("Desarme, ciclística y montaje completo", style_celda), Paragraph(f"$ {total_mo:,.2f}", style_celda_num)])
        else:
            if sub_mec_elec > 0:
                tareas_data.append([Paragraph("M.O. MECÁNICA Y ELECTRICIDAD", style_celda), Paragraph("Instalación y cierres", style_celda), Paragraph(f"$ {sub_mec_elec:,.2f}", style_celda_num)])
            if sub_chapa > 0:
                tareas_data.append([Paragraph(f"{dias_chapa:g} DÍAS DE CHAPA", style_celda), Paragraph(f"Valor unitario: $ {tarifa_dia_chapa:,.2f}", style_celda), Paragraph(f"$ {sub_chapa:,.2f}", style_celda_num)])
            if sub_pintura > 0:
                tareas_data.append([Paragraph(f"{panos:g} PAÑOS DE PINTURA", style_celda), Paragraph(f"Valor unitario: $ {tarifa_pano:,.2f}", style_celda), Paragraph(f"$ {sub_pintura:,.2f}", style_celda_num)])
            if sub_materiales > 0:
                tareas_data.append([Paragraph("MATERIALES DE PINTURA, FONDOS Y SELLADORES", style_celda), Paragraph("Insumos de cabina y selladores", style_celda), Paragraph(f"$ {sub_materiales:,.2f}", style_celda_num)])
            if sub_desarme > 0:
                tareas_data.append([Paragraph("M.O. ARME Y DESARME", style_celda), Paragraph("Desarme y ensamble de carrocería", style_celda), Paragraph(f"$ {sub_desarme:,.2f}", style_celda_num)])
            if sub_tren > 0:
                tareas_data.append([Paragraph("M.O. REPARACIÓN DE TREN DELANTERO", style_celda), Paragraph("Suspensión y tren rodante", style_celda), Paragraph(f"$ {sub_tren:,.2f}", style_celda_num)])
            if sub_motor > 0:
                tareas_data.append([Paragraph("M.O. MECÁNICA DE MOTOR Y ALINEACIÓN", style_celda), Paragraph("Cárter y alineación", style_celda), Paragraph(f"$ {sub_motor:,.2f}", style_celda_num)])
            if sub_serv > 0:
                tareas_data.append([Paragraph("KIT SERVICE DE ACEITE SINTÉTICO + FILTROS", style_celda), Paragraph("Fluidos y filtros oficiales", style_celda), Paragraph(f"$ {sub_serv:,.2f}", style_celda_num)])

        tareas_data.append([Paragraph("<b>Subtotal Mano de Obra</b>", style_celda_bold), Paragraph("", style_celda), Paragraph(f"<b>$ {total_mo:,.2f}</b>", style_celda_num_b)])

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
            Paragraph("<b>TOTAL</b>", style_sec_title),
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
    # GENERACIÓN DE EXCEL OFICIAL (.XLSX)
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
        if es_moto:
            ws.cell(row=fila, column=1, value=1).alignment = Alignment(horizontal="center")
            ws.cell(row=fila, column=2, value="M.O. Mecánica, Alineación y Armado")
            ws.cell(row=fila, column=3, value="Desarme, ciclística y montaje completo")
            c_mo = ws.cell(row=fila, column=4, value=float(total_mo))
            c_mo.number_format = "$#,##0.00"
            for col in range(1, 5):
                ws.cell(row=fila, column=col).border = borde
            fila += 1
        else:
            tareas_excel = []
            if sub_mec_elec > 0:
                tareas_excel.append(("M.O. Mecánica y Electricidad", "Instalación y cierres", sub_mec_elec))
            if sub_chapa > 0:
                tareas_excel.append((f"{dias_chapa:g} Días de Chapa", f"Unitario: ${tarifa_dia_chapa:,.2f}", sub_chapa))
            if sub_pintura > 0:
                tareas_excel.append((f"{panos:g} Paños de Pintura", f"Unitario: ${tarifa_pano:,.2f}", sub_pintura))
            if sub_materiales > 0:
                tareas_excel.append(("Materiales de Pintura y Selladores", "Insumos y selladores", sub_materiales))
            if sub_desarme > 0:
                tareas_excel.append(("M.O. Arme y Desarme", "Carrocería", sub_desarme))
            if sub_tren > 0:
                tareas_excel.append(("M.O. Reparación Tren Delantero", "Suspensión", sub_tren))
            if sub_motor > 0:
                tareas_excel.append(("M.O. Mecánica Motor y Alineación", "Cárter y alineación", sub_motor))
            if sub_serv > 0:
                tareas_excel.append(("Kit Service Aceite + Filtros", "Fluidos oficiales", sub_serv))
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
            file_name=f"Presupuesto_{datos.get('patente_detectada','VEHICULO').replace(' ','_')}_{nro_presupuesto}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    with col_xlsx:
        st.download_button(
            label="📊 Descargar Planilla de Control en Excel (.xlsx)",
            data=generar_excel(),
            file_name=f"Presupuesto_{datos.get('patente_detectada','VEHICULO').replace(' ','_')}_{nro_presupuesto}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
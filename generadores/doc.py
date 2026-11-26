import sys
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = sys.argv[1]
NAVY = RGBColor(0x14, 0x21, 0x3D)
ORANGE = RGBColor(0xE8, 0x77, 0x2E)
SLATE = RGBColor(0x5B, 0x6B, 0x7F)

doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(27.94), Cm(21.59)  # carta
for side in ("left_margin", "right_margin"):
    setattr(sec, side, Cm(2.5))
sec.top_margin = sec.bottom_margin = Cm(2.5)

st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(11.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15
for name, size in (("Heading 1", 15), ("Heading 2", 12.5)):
    h = doc.styles[name]
    h.font.name = "Cambria"; h.font.size = Pt(size); h.font.bold = True
    h.font.color.rgb = NAVY
    h.element.rPr.rFonts.set(qn("w:asciiTheme"), "") if False else None
    rf = h.element.rPr.rFonts
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Cambria")
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    h.paragraph_format.space_before = Pt(16 if name == "Heading 1" else 10)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True


def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hex_)
    tcPr.append(sh)


def P(text="", bold_lead=None, align=None, size=None, color=None, italic=False, after=None):
    p = doc.add_paragraph()
    if bold_lead:
        r = p.add_run(bold_lead); r.bold = True
    r = p.add_run(text); r.italic = italic
    if size: r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    if align: p.alignment = align
    else: p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p


def B(items, style="List Bullet"):
    for n, it in enumerate(items, 1):
        if style == "List Number":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.9); p.paragraph_format.first_line_indent = Cm(-0.6)
            p.add_run(f"{n}.\t")
            p.paragraph_format.tab_stops.add_tab_stop(Cm(0.9))
        else:
            p = doc.add_paragraph(style=style)
        if isinstance(it, tuple):
            r = p.add_run(it[0]); r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)
        p.paragraph_format.space_after = Pt(3)


def caption(text):
    p = doc.add_paragraph(); r = p.add_run(text); r.bold = True; r.font.size = Pt(10); r.font.color.rgb = SLATE
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(3); p.paragraph_format.keep_with_next = True


def T(header, rows, widths, cap=None, total_row=False):
    if cap: caption(cap)
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""
        r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(255, 255, 255)
        shade(c, "14213D")
    for k, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v)); r.font.size = Pt(10)
            if i == 0 or (total_row and k == len(rows) - 1): r.bold = True
            if k % 2 == 0: shade(cells[i], "EEF2F5")
    t.autofit = False
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[i] * 567)))
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    hdr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); hdr.append(th)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


# ---------------- Portada ----------------
for _ in range(5): doc.add_paragraph()
P("PROPUESTA", align=WD_ALIGN_PARAGRAPH.LEFT, size=12, color=ORANGE).runs[0].bold = True
p = P("Creación del Capítulo IEEE Computer Society", align=WD_ALIGN_PARAGRAPH.LEFT, size=28, color=NAVY)
p.runs[0].bold = True; p.runs[0].font.name = "Cambria"
P("Rama Estudiantil IEEE USANTOTO", align=WD_ALIGN_PARAGRAPH.LEFT, size=15, color=SLATE)
for _ in range(6): doc.add_paragraph()
info = [
    ("Presentada ante: ", "Consejo de la Facultad de Ingeniería de Sistemas, Universidad Santo Tomás"),
    ("Proponente: ", "Santiago Alejandro Zúñiga Melo"),
    ("Instancia de articulación: ", "Rama Estudiantil IEEE USANTOTO"),
    ("Fecha: ", "Noviembre de 2026"),
]
for a, b in info:
    P(b, bold_lead=a, align=WD_ALIGN_PARAGRAPH.LEFT, after=2)
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# ---------------- 1 ----------------
doc.add_heading("1. Resumen de la propuesta", 1)
P("Se propone la creación del Capítulo IEEE Computer Society como cuarto capítulo técnico de la Rama Estudiantil IEEE USANTOTO, con origen en la Facultad de Ingeniería de Sistemas y abierto a la participación de estudiantes de los distintos programas de ingeniería de la Universidad.")
P("El capítulo se constituye con un grupo fundador de seis (6) estudiantes y un (1) profesional, todos con membresía vigente del IEEE y de la IEEE Computer Society. El costo de arranque asciende a un máximo de 453 USD antes de descuentos y se cubre con las membresías individuales de los afiliados; la propuesta no exige asignación presupuestal obligatoria de la Facultad.")
P("La creación se sustenta en un plan de actividades a doce (12) meses, una junta de cinco cargos y un régimen de reuniones periódicas que demuestran la continuidad exigida por el IEEE.")
P("aval de la Facultad de Ingeniería de Sistemas para conformar el grupo fundador y radicar la solicitud de creación ante el IEEE, junto con los apoyos institucionales descritos en la sección 10.", bold_lead="Decisión solicitada: ")

# ---------------- 2 ----------------
doc.add_heading("2. Antecedentes", 1)
P("La Rama Estudiantil IEEE USANTOTO opera desde la Facultad de Ingeniería Electrónica y reúne cerca de diecisiete (17) integrantes con membresía formalizada y afiliación a capítulo. El número de inscritos del presente año supera el de periodos anteriores. La rama cuenta con el Ing. Camilo Pardo como consejero.")
T(["Capítulo", "Área", "Estado actual"], [
    ["AESS", "Aeronáutica (Aerospace and Electronic Systems Society)", "Activo"],
    ["RAS", "Robótica", "Activo. Desarrolló la actividad «El Trepador» en la Semana Tomasina"],
    ["EDS", "Electrónica", "Menor actividad; talleres en colegios con retraso por permisos"],
], [2.5, 6.5, 7.5], cap="Tabla 1. Capítulos técnicos vigentes de la Rama Estudiantil IEEE USANTOTO")
P("Además de los capítulos técnicos, los integrantes pueden afiliarse al capítulo WIE (Mujer en la Ingeniería) y a un capítulo de enfoque social. Ninguno de los capítulos existentes atiende el área de computación e ingeniería de software.")
P("La propuesta se fundamenta en la reunión sostenida con el SAC y con la directiva de la rama el 18 de septiembre de 2026, en la cual se precisaron los requisitos de membresía, la documentación exigida y la ruta de trabajo para la creación del capítulo.")

# ---------------- 3 ----------------
doc.add_heading("3. Justificación e importancia", 1)
B([
    ("Cubre un vacío temático. ", "La rama no cuenta con un capítulo en computación e ingeniería de software, área central de la Facultad de Ingeniería de Sistemas."),
    ("Canaliza una demanda existente. ", "Las jornadas de divulgación de la rama han generado interés en estudiantes de otros programas de ingeniería, entre ellos bioingeniería, que hoy no encuentran un capítulo afín."),
    ("Fortalece la rama. ", "Un cuarto capítulo incrementa el número de afiliados, la oferta de actividades y el reconocimiento institucional de la rama."),
    ("Proyecta a la Facultad. ", "Vincula a estudiantes y docentes con la IEEE Computer Society, la sociedad técnica del IEEE dedicada a la computación, y con la sección IEEE correspondiente."),
])
P("El capítulo opera bajo la identidad única de la Rama Estudiantil IEEE USANTOTO y no como una iniciativa separada por facultad: expone su programación en las reuniones de rama, promueve la inscripción de los miembros de la rama en sus actividades abiertas e incorpora actividades interfacultades en su cronograma anual.")

# ---------------- 4 ----------------
doc.add_heading("4. Objetivos", 1)
doc.add_heading("4.1 Objetivo general", 2)
P("Constituir el Capítulo IEEE Computer Society de la Rama Estudiantil IEEE USANTOTO como un espacio permanente de formación, divulgación y proyección en computación e ingeniería de software, articulado con la rama y abierto a estudiantes de distintos programas de ingeniería.")
doc.add_heading("4.2 Objetivos específicos", 2)
B([
    "Ofrecer formación técnica complementaria en áreas de competencia de la IEEE Computer Society mediante talleres, charlas y actividades prácticas de acceso abierto.",
    "Vincular a estudiantes de programas distintos al de origen, aprovechando el interés ya manifestado en las jornadas de la rama.",
    "Fortalecer la Rama Estudiantil IEEE USANTOTO mediante el aumento de su base de afiliados, su oferta de actividades y su reconocimiento institucional.",
    "Garantizar la continuidad del capítulo mediante una junta definida, reuniones periódicas y un mecanismo de relevo anual.",
], style="List Number")

# ---------------- 5 ----------------
doc.add_heading("5. Beneficios", 1)
doc.add_heading("5.1 Para los estudiantes", 2)
B([
    ("Formación complementaria ", "en desarrollo de software, inteligencia artificial, ciberseguridad y computación en la nube."),
    ("Experiencia en liderazgo y gestión ", "a través de los cinco cargos de junta, con funciones definidas y relevo anual."),
    ("Acceso a recursos de la IEEE Computer Society ", "(publicaciones, eventos y programas para miembros, según la oferta vigente al momento de la afiliación)."),
    ("Red profesional ", "con profesionales, otras ramas estudiantiles y la sección IEEE."),
    ("Trabajo interdisciplinario ", "mediante actividades conjuntas con otros programas de ingeniería."),
    ("Proyección social ", "con actividades de divulgación en instituciones de educación media, sujetas a los permisos institucionales."),
])
doc.add_heading("5.2 Para la Facultad de Ingeniería de Sistemas", 2)
B([
    "Espacio extracurricular permanente, alineado con las áreas de formación del programa.",
    "Evidencia documentada de actividades de formación complementaria y proyección (actas, balances semestrales e informe anual), útil como soporte en procesos de autoevaluación.",
    "Visibilidad del programa ante el IEEE y su sección.",
    "Formación de líderes estudiantiles mediante una junta con relevo anual.",
])
doc.add_heading("5.3 Para la Universidad", 2)
B([
    "Trabajo interfacultades bajo una sola identidad institucional: la Rama Estudiantil IEEE USANTOTO.",
    "Participación en la Semana Tomasina y en las actividades conjuntas de la rama.",
    "Mayor base de afiliados y mayor reconocimiento de la rama.",
    "Vínculo con instituciones de educación media a través de actividades de divulgación.",
])

# ---------------- 6 ----------------
doc.add_heading("6. Requisitos para la creación del capítulo", 1)
doc.add_heading("6.1 Requisitos de membresía", 2)
P("El IEEE exige seis (6) estudiantes y un (1) profesional de grado Member o superior, todos con membresía vigente del IEEE y de la IEEE Computer Society. Solo se contabilizan como miembros del capítulo quienes acrediten ambas membresías; los integrantes de la junta deben cumplir igualmente este requisito.")
doc.add_heading("6.2 Requisitos documentales", 2)
P("Reunidos los siete (7) miembros, se diligencia el formulario de creación disponible en la página del IEEE, cuyo enlace remite el SAC. El formulario exige un documento descriptivo con cinco componentes:")
T(["Componente", "Contenido exigido"], [
    ["Propósito", "Objeto y actividades que pretende desarrollar el capítulo"],
    ["Cronograma a doce meses", "Actividades del año que evidencien continuidad y no una iniciativa temporal"],
    ["Reuniones programadas", "Reuniones de la junta del capítulo y reuniones con la junta de la rama"],
    ["Actividades interfacultades", "Acciones conjuntas con otras facultades"],
    ["Junta del capítulo", "Líder, vicepresidente, secretario, tesorero y webmaster, con las membresías requeridas"],
], [5.0, 11.5], cap="Tabla 2. Componentes del documento descriptivo exigido por el IEEE")
doc.add_heading("6.3 Avales", 2)
B([
    "Aval de la Facultad de Ingeniería de Sistemas, objeto de la presente propuesta.",
    "Revisión de la iniciativa por el consejero de la rama, Ing. Camilo Pardo.",
    "Aval de la sección IEEE.",
])
P("El criterio de fondo del IEEE es la perdurabilidad: el capítulo debe demostrar programación sostenida y liderazgo definido.", italic=True)

# ---------------- 7 ----------------
doc.add_heading("7. Estructura organizativa", 1)
P("El capítulo se integra a una estructura de tres niveles: la junta del capítulo, la junta de la rama (con su consejero) y la sección IEEE, que otorga el aval al nuevo capítulo.")
T(["Cargo", "Funciones principales"], [
    ["Líder del capítulo", "Representa al capítulo ante la junta de la rama y el consejero; convoca las reuniones de junta; responde por la ejecución del plan anual"],
    ["Vicepresidente", "Apoya la dirección, asume las funciones del líder en su ausencia y coordina la línea de formación técnica"],
    ["Secretario", "Lleva las actas de las reuniones, mantiene el registro de asistencia y custodia la documentación"],
    ["Tesorero", "Administra los recursos, hace seguimiento al estado de las membresías y gestiona descuentos y rebates"],
    ["Webmaster", "Administra los medios de divulgación y apoya la difusión de actividades en coordinación con la rama"],
], [4.0, 12.5], cap="Tabla 3. Cargos de la junta del capítulo")
P("El periodo de la junta es de un (1) año, con relevo al cierre del ciclo. La designación inicial se efectúa una vez conformado el grupo fundador.")

# ---------------- 8 ----------------
doc.add_heading("8. Plan de trabajo a doce meses", 1)
doc.add_heading("8.1 Líneas de trabajo", 2)
B([
    ("Formación técnica. ", "Línea principal: talleres, charlas y prácticas en desarrollo de software, inteligencia artificial, ciberseguridad y computación en la nube."),
    ("Proyección y divulgación. ", "Actividades para otros programas de ingeniería y, sujeto a permisos, para instituciones de educación media."),
    ("Articulación institucional. ", "Presentación de la programación en las reuniones de rama, participación en la Semana Tomasina y en los espacios de la sección."),
])
doc.add_heading("8.2 Régimen de reuniones", 2)
B([
    ("Junta del capítulo: ", "reunión mensual durante los doce meses, incluidos los recesos, con acta a cargo del secretario."),
    ("Junta de la rama: ", "asistencia del líder a las reuniones ordinarias de la rama para presentar el estado del plan."),
    ("Balances semestrales: ", "al cierre de cada periodo académico, con reporte de actividades y estado de membresías."),
])
doc.add_heading("8.3 Cronograma", 2)
T(["Periodo", "Meses", "Actividades previstas"], [
    ["Instalación", "1", "Instalación de la junta; presentación ante la junta de la rama; plan operativo del primer semestre; apertura de medios de divulgación"],
    ["Primer semestre", "2 a 5", "Jornada de vinculación de afiliados; una actividad de formación técnica; una actividad interfacultades; gestión anticipada de permisos"],
    ["Balance y receso", "6", "Balance semestral ante la rama; planeación del segundo semestre; depuración del registro de membresías"],
    ["Segundo semestre", "7 a 10", "Dos actividades de formación técnica; una actividad interfacultades; Semana Tomasina; proyección externa sujeta a permisos"],
    ["Cierre del ciclo", "11 a 12", "Informe anual; verificación de membresías; designación de la nueva junta; empalme y entrega documental"],
], [3.5, 1.8, 11.2], cap="Tabla 4. Cronograma de actividades (meses contados desde la aprobación del capítulo)")
P("La distribución constituye un mínimo. Los temas de las actividades de formación los define la junta al inicio de cada semestre según los intereses de los afiliados y la disponibilidad de ponentes.")
doc.add_heading("8.4 Actividades interfacultades", 2)
P("Se prevén como mínimo dos (2) actividades interfacultades, una por semestre, en tres modalidades: actividad técnica conjunta con un capítulo o semillero de otro programa; participación cruzada en actividades de otros capítulos de la rama; y actividades conjuntas de toda la rama, como las de la Semana Tomasina. Los programas con interés ya manifestado, entre ellos bioingeniería, son el punto de partida de la primera actividad.")

# ---------------- 9 ----------------
doc.add_heading("9. Costos y financiación", 1)
T(["Perfil", "Cantidad mínima", "Membresía IEEE (USD)", "Membresía Computer Society (USD)", "Costo por persona (USD)", "Subtotal (USD)"], [
    ["Estudiante", "6", "27", "8", "35", "210"],
    ["Profesional (Member o superior)", "1", "183", "60", "243", "243"],
    ["Total grupo fundador", "7", "", "", "", "453"],
], [4.2, 2.0, 2.4, 3.0, 2.6, 2.3], cap="Tabla 5. Costo de arranque del grupo fundador, sin descuentos", total_row=True)
P("Los integrantes actuales de la rama ya cuentan con membresía IEEE vigente y solo deben adquirir la de Computer Society (8 USD). Si los seis estudiantes fundadores provienen de la rama, el costo del grupo se reduce a 291 USD (6 × 8 USD + 243 USD).")
P("Dos mecanismos permiten reducir adicionalmente el costo:")
B([
    ("Descuento institucional ", "sobre la primera membresía, tramitada conforme a los espacios acordados con la Universidad."),
    ("Rebates del IEEE, ", "cuya gestión adelanta la rama para cubrir las primeras membresías de los nuevos integrantes."),
])
P("Las tarifas corresponden a las indicadas en la reunión con el SAC, se actualizan por periodo y deben verificarse al momento de la inscripción. El porcentaje del descuento y las condiciones de los rebates están por confirmar.", italic=True)

# ---------------- 10 ----------------
doc.add_heading("10. Apoyo solicitado a la Facultad", 1)
P("La propuesta no requiere asignación presupuestal obligatoria de la Facultad, dado que las membresías son individuales. Se solicitan los siguientes apoyos institucionales:")
T(["Apoyo", "Descripción"], [
    ["Aval institucional", "Respaldo de la Facultad para la creación del capítulo y su presentación ante la rama y la sección IEEE"],
    ["Profesional fundador", "Apoyo para vincular a un docente o egresado con membresía profesional IEEE y Computer Society"],
    ["Espacios", "Salones o laboratorios para las reuniones de junta y las actividades del capítulo"],
    ["Divulgación", "Difusión de la convocatoria y de las actividades por los canales de la Facultad"],
    ["Permisos", "Agilización de los permisos para actividades externas y con instituciones de educación media"],
    ["Apoyo opcional", "Si la Facultad lo considera viable, cubrir la membresía del profesional fundador (243 USD)"],
], [4.0, 12.5], cap="Tabla 6. Apoyos solicitados a la Facultad")

# ---------------- 11 ----------------
doc.add_heading("11. Ruta de implementación", 1)
T(["Fase", "Acciones", "Responsable propuesto", "Criterio de cumplimiento"], [
    ["1. Conformación", "Convocar a 20 aspirantes y 3 candidatos a profesional; sesión informativa; pago de membresías con descuento y rebates", "Promotores del capítulo, con apoyo de la rama", "7 membresías IEEE y Computer Society vigentes"],
    ["2. Documentación", "Redactar propósito, cronograma, plan de reuniones y estructura de junta", "Junta propuesta del capítulo", "Documento descriptivo completo con los 5 cargos asignados"],
    ["3. Aval y registro", "Revisión con el consejero, aval de la sección y diligenciamiento del formulario", "Líder del capítulo, con el consejero", "Aval de la sección y formulario radicado"],
    ["4. Operación", "Ejecutar el cronograma: reuniones de junta, reuniones con la rama y actividades interfacultades", "Junta del capítulo", "Cronograma cumplido durante 12 meses"],
], [3.0, 6.0, 3.7, 3.8], cap="Tabla 7. Fases de implementación")
P("La fase de conformación es la de mayor duración y condiciona las demás. Por la deserción esperable entre la manifestación de interés y el pago, la convocatoria se dimensiona con margen: veinte (20) estudiantes convocados para obtener seis (6) afiliaciones y tres (3) candidatos para obtener un (1) profesional. La convocatoria avanza por hitos y no por fechas de cierre, y se mantiene abierta hasta completar el grupo.")

# ---------------- 12 ----------------
doc.add_heading("12. Riesgos y medidas de tratamiento", 1)
T(["Riesgo", "Efecto", "Medida de tratamiento"], [
    ["No reunir los 7 miembros con ambas membresías", "El formulario no puede radicarse", "Convocatoria con margen y refuerzo con otros programas; acompañamiento de la rama en el trámite"],
    ["Costo de las membresías", "Desistimiento de aspirantes", "Comunicar el valor con descuento desde el primer contacto; priorizar a integrantes de la rama (8 USD); gestionar rebates"],
    ["Ausencia del profesional", "Capítulo incompleto", "Abordar 3 candidatos en paralelo, con apoyo del consejero y de la Facultad"],
    ["Demora en permisos institucionales", "Retraso de actividades", "Tramitar al inicio de cada semestre; mantener alternativas internas"],
    ["Vencimiento de membresías", "Pérdida del mínimo exigido", "Seguimiento del tesorero y verificación en cada balance semestral"],
    ["Falta de continuidad", "Inactividad o pérdida del aval", "Actas mensuales, balances semestrales, informe anual y relevo de junta"],
], [4.5, 3.8, 8.2], cap="Tabla 8. Riesgos y medidas de tratamiento")

# ---------------- 13 ----------------
doc.add_heading("13. Indicadores de seguimiento", 1)
T(["Indicador", "Meta"], [
    ["Membresías IEEE y Computer Society vigentes", "Mínimo 7 durante todo el periodo"],
    ["Actas de reunión de junta", "12 (una por mes)"],
    ["Balances semestrales ante la rama", "2"],
    ["Actividades de formación técnica", "Mínimo 3 en el año"],
    ["Actividades interfacultades", "Mínimo 2 (una por semestre)"],
    ["Informe anual entregado a la rama y a la Facultad", "1, al cierre del ciclo"],
], [9.5, 7.0], cap="Tabla 9. Indicadores y metas del primer año")

# ---------------- 14 ----------------
doc.add_heading("14. Aspectos por confirmar", 1)
B([
    "Número mínimo de miembros: la rama mencionó ocho por capítulo, mientras que el SAC confirmó seis estudiantes y un profesional.",
    "Tarifas vigentes del IEEE y de la IEEE Computer Society al momento de la inscripción.",
    "Porcentaje del descuento institucional y condiciones de los rebates.",
    "Fecha de vencimiento del periodo de membresía.",
    "Periodicidad de las reuniones de la junta de la rama.",
    "Rol de María Paula en la evaluación del trámite.",
    "Envío del formulario de creación por parte del SAC.",
])

# ---------------- 15 ----------------
doc.add_heading("15. Solicitud", 1)
P("Con base en lo expuesto, se solicita al Consejo de la Facultad de Ingeniería de Sistemas:")
B([
    "Otorgar el aval de la Facultad para la creación del Capítulo IEEE Computer Society de la Rama Estudiantil IEEE USANTOTO.",
    "Respaldar la vinculación de un docente o egresado como profesional fundador.",
    "Facilitar espacios, canales de divulgación y la gestión de permisos para las actividades del capítulo.",
], style="List Number")
P("Otorgado el aval, se inicia de inmediato la convocatoria del grupo fundador y, una vez completado, la radicación de la solicitud ante el IEEE.")
for _ in range(3): doc.add_paragraph()
P("Santiago Alejandro Zúñiga Melo", align=WD_ALIGN_PARAGRAPH.LEFT, after=0).runs[0].bold = True
P("Proponente · Rama Estudiantil IEEE USANTOTO", align=WD_ALIGN_PARAGRAPH.LEFT)

# footer with page number
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = fp.add_run("Propuesta de creación del Capítulo IEEE Computer Society · "); r.font.size = Pt(9); r.font.color.rgb = SLATE
r = fp.add_run(); r.font.size = Pt(9)
for t, v in (("begin", None), (None, "PAGE"), ("end", None)):
    if t:
        e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), t); r._r.append(e)
    else:
        e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = v; r._r.append(e)
sec.different_first_page_header_footer = True

doc.core_properties.title = "Propuesta de creación del Capítulo IEEE Computer Society"
doc.core_properties.author = "Santiago Alejandro Zúñiga Melo"
doc.save(OUT)
print("ok", OUT)

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "IEEE CS USANTOTO",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B1F24", lt1: "FFFFFF", dk2: "14213D", lt2: "EEF2F5",
    accent1: "E8772E", accent2: "0F7C82", accent3: "5B6B7F",
    accent4: "F6C9A6", accent5: "BFE0E2", accent6: "B23A2E",
    hlink: "0F7C82", folHlink: "5B6B7F",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Propuesta de creación del Capítulo IEEE Computer Society";
pres.author = "Santiago Alejandro Zúñiga Melo";
pres.subject = "Rama Estudiantil IEEE USANTOTO";
const C = pres.SchemeColor;

const W = 13.333, M = 0.6;
const FOOT = "Propuesta de creación del Capítulo IEEE Computer Society  ·  Rama Estudiantil IEEE USANTOTO";

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "PORTADA",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 1.5, w: 9, h: 0.5, fontSize: 16, bold: true, color: C.accent1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.05, w: 9.6, h: 1.9, fontSize: 44, bold: true, color: C.background1, fontFace: "Cambria", valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.15, w: 9.6, h: 1.0, fontSize: 18, color: C.accent5, valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "author", type: "body", x: M, y: 5.9, w: 9, h: 0.8, fontSize: 14, color: C.background1, valign: "top", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.7, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.4, w: W - 2 * M, h: 0.85, fontSize: 32, bold: true, color: C.text2, fontFace: "Cambria", valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: M, y: 7.0, w: 9, h: 0.3, fontSize: 10, color: C.accent3, margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent3, align: "right", margin: 0 },
});

pres.defineSlideMaster({
  title: "CIERRE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.5, w: W - 2 * M, h: 0.9, fontSize: 36, bold: true, color: C.background1, fontFace: "Cambria", valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right", margin: 0 },
});

// ---------- Helpers ----------
async function iconPng(name, hex) {
  const el = React.createElement(fa[name], { color: "#" + hex, size: 256 });
  const svg = ReactDOMServer.renderToStaticMarkup(el);
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function iconCircle(slide, name, x, y, d, fill = C.accent1, iconHex = "FFFFFF", id = "icono") {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: id + "-circulo" });
  const p = d * 0.25;
  slide.addImage({ data: await iconPng(name, iconHex), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: id, objectName: id });
}
function card(slide, x, y, w, h, fill = C.background2, id = "tarjeta") {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: id });
}
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", fontSize: 15, color: C.text1 }, opts));
}
function bullets(items, size = 15, color = C.text1) {
  return items.map((t, i) => ({ text: t, options: { bullet: { indent: 16 }, breakLine: i < items.length - 1, fontSize: size, color, paraSpaceAfter: 6 } }));
}
function content(section, title) {
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
}

(async () => {
  // ===== 1. Portada =====
  pres.addSection({ title: "Apertura" });
  {
    const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Apertura" });
    s.addText("PROPUESTA · FACULTAD DE INGENIERÍA DE SISTEMAS", { placeholder: "kicker" });
    s.addText("Creación del Capítulo IEEE Computer Society", { placeholder: "title" });
    s.addText("Cuarto capítulo técnico de la Rama Estudiantil IEEE USANTOTO, Universidad Santo Tomás", { placeholder: "body" });
    s.addText([
      { text: "Santiago Alejandro Zúñiga Melo", options: { bold: true, breakLine: true } },
      { text: "Noviembre de 2026" },
    ], { placeholder: "author" });
    // motif: large icon in circle on the right
    await iconCircle(s, "FaLaptopCode", 10.0, 2.1, 2.6, C.accent1, "FFFFFF", "icono-portada");
    s.addNotes("Presentación de la propuesta de creación del Capítulo IEEE Computer Society en la Rama Estudiantil IEEE USANTOTO, con origen en la Facultad de Ingeniería de Sistemas.");
  }

  // ===== 2. Síntesis =====
  {
    const s = content("Apertura", "La propuesta en síntesis");
    txt(s, "Crear el Capítulo IEEE Computer Society como cuarto capítulo técnico de la Rama Estudiantil IEEE USANTOTO, con origen en la Facultad de Ingeniería de Sistemas y abierto a estudiantes de todos los programas de ingeniería.",
      { x: M, y: 1.45, w: W - 2 * M, h: 0.9, fontSize: 18 });
    const stats = [
      ["7", "miembros fundadores", "6 estudiantes y 1 profesional con membresía IEEE y Computer Society"],
      ["453 USD", "costo de arranque", "Valor máximo, antes de descuentos, cubierto con las membresías individuales"],
      ["12 meses", "plan de actividades", "Reuniones mensuales, actividades interfacultades y relevo de junta"],
    ];
    const cw = (W - 2 * M - 2 * 0.4) / 3;
    stats.forEach(([n, l, d], i) => {
      const x = M + i * (cw + 0.4);
      card(s, x, 2.6, cw, 2.6, C.background2, "dato-" + (i + 1));
      txt(s, n, { x: x + 0.35, y: 2.85, w: cw - 0.7, h: 0.95, fontSize: 48, bold: true, color: C.accent1, fontFace: "Cambria", valign: "middle" });
      txt(s, l, { x: x + 0.35, y: 3.85, w: cw - 0.7, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
      txt(s, d, { x: x + 0.35, y: 4.3, w: cw - 0.7, h: 0.8, fontSize: 14, color: C.accent3 });
    });
    card(s, M, 5.55, W - 2 * M, 1.05, C.text2, "decision");
    txt(s, [
      { text: "Decisión solicitada: ", options: { bold: true, color: C.accent1 } },
      { text: "aval de la Facultad para conformar el grupo fundador y radicar la solicitud de creación ante el IEEE.", options: { color: C.background1 } },
    ], { x: M + 0.35, y: 5.55, w: W - 2 * M - 0.7, h: 1.05, fontSize: 17, valign: "middle" });
    s.addNotes("Mensaje central: la creación del capítulo no exige asignación presupuestal de la Facultad; los costos se cubren con membresías individuales. Lo que se solicita es el aval y un conjunto acotado de apoyos institucionales.");
  }

  // ===== 3. Contexto =====
  pres.addSection({ title: "Importancia y beneficios" });
  {
    const s = content("Importancia y beneficios", "La rama hoy: ningún capítulo dedicado a computación");
    const caps = [
      ["AESS", "Aeronáutica", "Activo", C.accent2],
      ["RAS", "Robótica", "Activo · actividad «El Trepador» en Semana Tomasina", C.accent2],
      ["EDS", "Electrónica", "Talleres en colegios, con retraso por permisos", C.accent3],
      ["Computer Society", "Computación y software", "Propuesto", C.accent1],
    ];
    const gap = 0.3, cw = (W - 2 * M - 3 * gap) / 4;
    caps.forEach(([n, a, e, col], i) => {
      const x = M + i * (cw + gap);
      card(s, x, 1.55, cw, 2.75, i === 3 ? C.accent4 : C.background2, "capitulo-" + n);
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 1.85, w: 0.32, h: 0.32, fill: { color: col }, line: { type: "none" }, objectName: "estado-" + n });
      txt(s, n, { x: x + 0.3, y: 2.3, w: cw - 0.6, h: 0.75, fontSize: 20, bold: true, color: C.text2, fontFace: "Cambria", valign: "bottom" });
      txt(s, a, { x: x + 0.3, y: 3.1, w: cw - 0.6, h: 0.5, fontSize: 14, bold: true, color: C.text1 });
      txt(s, e, { x: x + 0.3, y: 3.55, w: cw - 0.6, h: 0.7, fontSize: 13, color: C.accent3 });
    });
    // right stat
    card(s, M, 4.6, 3.6, 2.15, C.text2, "dato-integrantes");
    txt(s, "≈17", { x: M + 0.3, y: 4.7, w: 3.0, h: 0.95, fontSize: 50, bold: true, color: C.accent1, fontFace: "Cambria", valign: "middle" });
    txt(s, "integrantes con membresía; inscripciones del año por encima de periodos anteriores", { x: M + 0.3, y: 5.7, w: 3.0, h: 0.95, fontSize: 14, color: C.background1 });

    txt(s, bullets([
      "La rama opera desde la Facultad de Ingeniería Electrónica, con el Ing. Camilo Pardo como consejero.",
      "Además de los capítulos técnicos, los integrantes pueden afiliarse a WIE (Mujer en la Ingeniería) y a un capítulo de enfoque social.",
      "Las jornadas de divulgación han atraído a estudiantes de otros programas, entre ellos bioingeniería: hay demanda que hoy no tiene un capítulo que la reciba.",
    ], 15), { x: M + 3.95, y: 4.65, w: W - 2 * M - 3.95, h: 2.1 });
    s.addNotes("La rama no cuenta con un capítulo dedicado a computación, que es el área central de la Facultad de Ingeniería de Sistemas.");
  }

  // ===== 4. Importancia =====
  {
    const s = content("Importancia y beneficios", "Por qué es importante crear el capítulo");
    const rows = [
      ["FaPuzzlePiece", "Cubre un vacío", "La rama no tiene un capítulo en computación e ingeniería de software, el área central de la Facultad de Ingeniería de Sistemas."],
      ["FaUsers", "Canaliza una demanda existente", "Estudiantes de otros programas de ingeniería ya han manifestado interés en las actividades de la rama."],
      ["FaChartLine", "Fortalece la rama", "Aumenta el número de afiliados, la oferta de actividades y el reconocimiento institucional de la rama."],
      ["FaGlobeAmericas", "Conecta a la Facultad con una red internacional", "Vincula a estudiantes y docentes con la IEEE Computer Society, sociedad técnica del IEEE dedicada a la computación."],
    ];
    const rh = 1.25;
    for (let i = 0; i < rows.length; i++) {
      const [ic, h, d] = rows[i];
      const y = 1.5 + i * rh;
      await iconCircle(s, ic, M, y + 0.1, 0.8, C.accent1, "FFFFFF", "icono-" + (i + 1));
      txt(s, h, { x: M + 1.1, y: y + 0.05, w: 10.8, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
      txt(s, d, { x: M + 1.1, y: y + 0.5, w: 10.8, h: 0.65, fontSize: 15, color: C.text1 });
    }
    s.addNotes("Cuatro razones: vacío temático, demanda existente, fortalecimiento de la rama y proyección internacional.");
  }

  // ===== 5. Beneficios estudiantes =====
  {
    const s = content("Importancia y beneficios", "Beneficios para los estudiantes");
    const items = [
      ["FaCode", "Formación complementaria", "Talleres y charlas en software, inteligencia artificial, ciberseguridad y nube."],
      ["FaUserTie", "Liderazgo y gestión", "Cinco cargos de junta con funciones definidas y relevo anual."],
      ["FaBookOpen", "Recursos de la Computer Society", "Acceso a publicaciones, eventos y programas que la sociedad ofrece a sus miembros."],
      ["FaNetworkWired", "Red profesional", "Contacto con profesionales, otras ramas estudiantiles y la sección IEEE."],
      ["FaProjectDiagram", "Trabajo interdisciplinario", "Actividades conjuntas con estudiantes de otros programas de ingeniería."],
      ["FaSchool", "Proyección social", "Divulgación en instituciones de educación media, sujeta a los permisos institucionales."],
    ];
    const cw = (W - 2 * M - 2 * 0.3) / 3, ch = 2.5;
    for (let i = 0; i < items.length; i++) {
      const [ic, h, d] = items[i];
      const x = M + (i % 3) * (cw + 0.3), y = 1.5 + Math.floor(i / 3) * (ch + 0.3);
      card(s, x, y, cw, ch, C.background2, "beneficio-" + (i + 1));
      await iconCircle(s, ic, x + 0.3, y + 0.3, 0.65, C.accent2, "FFFFFF", "icono-beneficio-" + (i + 1));
      txt(s, h, { x: x + 0.3, y: y + 1.08, w: cw - 0.6, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
      txt(s, d, { x: x + 0.3, y: y + 1.5, w: cw - 0.6, h: 0.85, fontSize: 14, color: C.text1 });
    }
    s.addNotes("Los beneficios de membresía de la IEEE Computer Society deben verificarse en la oferta vigente al momento de la inscripción.");
  }

  // ===== 6. Beneficios Facultad / Universidad =====
  {
    const s = content("Importancia y beneficios", "Beneficios para la Facultad y la Universidad");
    const cols = [
      ["FaUniversity", "Facultad de Ingeniería de Sistemas", [
        "Espacio extracurricular permanente, alineado con las áreas del programa.",
        "Evidencia documentada de formación complementaria y proyección: actas, balances e informe anual.",
        "Visibilidad del programa ante el IEEE y su sección regional.",
        "Formación de líderes estudiantiles con relevo anual de junta.",
      ]],
      ["FaLandmark", "Universidad Santo Tomás", [
        "Trabajo interfacultades bajo una sola identidad: la Rama Estudiantil IEEE USANTOTO.",
        "Participación en la Semana Tomasina y en las actividades conjuntas de la rama.",
        "Mayor base de afiliados y mayor reconocimiento institucional de la rama.",
        "Vínculo con instituciones de educación media mediante actividades de divulgación.",
      ]],
    ];
    const cw = (W - 2 * M - 0.4) / 2;
    for (let i = 0; i < 2; i++) {
      const [ic, h, list] = cols[i];
      const x = M + i * (cw + 0.4);
      card(s, x, 1.5, cw, 5.2, C.background2, "columna-" + (i + 1));
      await iconCircle(s, ic, x + 0.35, 1.8, 0.75, i === 0 ? C.accent1 : C.accent2, "FFFFFF", "icono-columna-" + (i + 1));
      txt(s, h, { x: x + 1.3, y: 1.8, w: cw - 1.6, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle" });
      txt(s, bullets(list, 15), { x: x + 0.35, y: 2.85, w: cw - 0.7, h: 3.7 });
    }
    s.addNotes("La documentación que produce el capítulo (actas, balances, informe anual) puede servir de soporte a la Facultad en sus procesos de autoevaluación.");
  }

  // ===== 7. Propósito y líneas =====
  pres.addSection({ title: "Propuesta" });
  {
    const s = content("Propuesta", "Propósito y líneas de trabajo");
    card(s, M, 1.45, W - 2 * M, 1.2, C.text2, "proposito");
    txt(s, [
      { text: "Propósito: ", options: { bold: true, color: C.accent1 } },
      { text: "constituir un espacio permanente de formación, divulgación y proyección en computación e ingeniería de software, articulado con la rama y abierto a estudiantes de distintos programas de ingeniería.", options: { color: C.background1 } },
    ], { x: M + 0.35, y: 1.45, w: W - 2 * M - 0.7, h: 1.2, fontSize: 16, valign: "middle" });
    const lines = [
      ["FaChalkboardTeacher", "1. Formación técnica", "Línea principal. Talleres, charlas y prácticas en desarrollo de software, IA, ciberseguridad y computación en la nube."],
      ["FaBullhorn", "2. Proyección y divulgación", "Actividades para otros programas de ingeniería y, con los permisos del caso, para instituciones de educación media."],
      ["FaHandshake", "3. Articulación institucional", "Presentación de la programación en reuniones de rama, Semana Tomasina y espacios de la sección IEEE."],
    ];
    const cw = (W - 2 * M - 2 * 0.3) / 3;
    for (let i = 0; i < 3; i++) {
      const [ic, h, d] = lines[i];
      const x = M + i * (cw + 0.3);
      card(s, x, 2.95, cw, 3.75, C.background2, "linea-" + (i + 1));
      await iconCircle(s, ic, x + 0.35, 3.25, 0.8, C.accent1, "FFFFFF", "icono-linea-" + (i + 1));
      txt(s, h, { x: x + 0.35, y: 4.25, w: cw - 0.7, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
      txt(s, d, { x: x + 0.35, y: 4.75, w: cw - 0.7, h: 1.8, fontSize: 15 });
    }
    s.addNotes("Cada actividad del cronograma se adscribe a una de estas tres líneas, para que la programación responda a un programa articulado y no a eventos aislados.");
  }

  // ===== 8. Requisitos =====
  {
    const s = content("Propuesta", "Requisitos para la creación ante el IEEE");
    const reqs = [
      ["FaIdCard", "Membresías", ["6 estudiantes y 1 profesional (grado Member o superior)", "Todos con membresía IEEE y de la IEEE Computer Society vigentes", "La junta del capítulo cumple el mismo requisito"]],
      ["FaFileAlt", "Documento descriptivo", ["Propósito del capítulo", "Cronograma a doce meses", "Reuniones de junta y con la rama", "Actividades interfacultades", "Junta: líder, vicepresidente, secretario, tesorero y webmaster"]],
      ["FaStamp", "Avales y radicación", ["Revisión del consejero de la rama, Ing. Camilo Pardo", "Aval de la sección IEEE", "Formulario de creación en la página del IEEE, con enlace remitido por el SAC"]],
    ];
    const cw = (W - 2 * M - 2 * 0.3) / 3;
    for (let i = 0; i < 3; i++) {
      const [ic, h, list] = reqs[i];
      const x = M + i * (cw + 0.3);
      card(s, x, 1.5, cw, 4.55, C.background2, "requisito-" + (i + 1));
      await iconCircle(s, ic, x + 0.35, 1.8, 0.75, C.accent2, "FFFFFF", "icono-requisito-" + (i + 1));
      txt(s, h, { x: x + 1.25, y: 1.8, w: cw - 1.5, h: 0.75, fontSize: 19, bold: true, color: C.text2, valign: "middle" });
      txt(s, bullets(list, 14), { x: x + 0.35, y: 2.8, w: cw - 0.7, h: 3.15 });
    }
    txt(s, [
      { text: "Criterio de evaluación del IEEE: ", options: { bold: true, color: C.accent1 } },
      { text: "perdurabilidad. El capítulo debe demostrar programación sostenida y liderazgo definido, no una iniciativa temporal.", options: { color: C.text1 } },
    ], { x: M, y: 6.25, w: W - 2 * M, h: 0.5, fontSize: 15 });
    s.addNotes("La conformación del grupo de siete miembros es la etapa de mayor duración; una vez completada, las demás avanzan con mayor rapidez.");
  }

  // ===== 9. Costos =====
  {
    const s = content("Propuesta", "Costos de arranque");
    s.addChart(pres.charts.BAR, [
      { name: "Membresía IEEE", labels: ["Estudiante nuevo", "Estudiante ya afiliado a la rama", "Profesional"], values: [27, 0, 183] },
      { name: "Membresía Computer Society", labels: ["Estudiante nuevo", "Estudiante ya afiliado a la rama", "Profesional"], values: [8, 8, 60] },
    ], {
      x: M, y: 1.45, w: 6.6, h: 5.25, barDir: "bar", barGrouping: "stacked",
      chartColors: [HEX.dk2, HEX.accent1],
      showTitle: true, title: "Valor por persona (USD)", titleFontSize: 14, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt",
      catAxisLabelColor: HEX.accent3, valAxisLabelColor: HEX.accent3, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
      catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
      valGridLine: { color: "DDE3E8", size: 0.5 }, catGridLine: { style: "none" },
      showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: HEX.accent3, legendFontFace: "+mn-lt",
      valAxisMaxVal: 250,
    });
    const rx = 7.6, rw = W - M - rx;
    card(s, rx, 1.45, rw, 2.05, C.text2, "costo-base");
    txt(s, "453 USD", { x: rx + 0.35, y: 1.6, w: rw - 0.7, h: 0.85, fontSize: 40, bold: true, color: C.accent1, fontFace: "Cambria", valign: "middle" });
    txt(s, "Grupo mínimo sin descuentos: 6 estudiantes nuevos × 35 USD + 1 profesional × 243 USD", { x: rx + 0.35, y: 2.5, w: rw - 0.7, h: 0.85, fontSize: 14, color: C.background1 });
    card(s, rx, 3.7, rw, 1.45, C.background2, "costo-alterno");
    txt(s, [
      { text: "291 USD", options: { bold: true, color: C.accent2, fontSize: 26, fontFace: "Cambria", breakLine: true } },
      { text: "si los 6 estudiantes ya pertenecen a la rama (6 × 8 USD + 243 USD)", options: { fontSize: 14, color: C.text1 } },
    ], { x: rx + 0.35, y: 3.82, w: rw - 0.7, h: 1.25 });
    txt(s, bullets([
      "Cada afiliado paga su membresía; la rama acompaña el trámite.",
      "Reducciones: descuento institucional a la primera membresía y rebates del IEEE (por confirmar).",
      "Tarifas sujetas a verificación al momento de la inscripción.",
    ], 13, C.text1), { x: rx, y: 5.3, w: rw, h: 1.45 });
    s.addNotes("Los valores provienen de las tarifas indicadas en la reunión con el SAC del 18 de septiembre de 2026. Priorizar integrantes actuales de la rama reduce el costo agregado de 453 a 291 USD.");
  }

  // ===== 10. Apoyo solicitado =====
  {
    const s = content("Propuesta", "Apoyo solicitado a la Facultad");
    const asks = [
      ["FaCheckCircle", "Aval institucional", "Respaldo de la Facultad para la creación del capítulo y su presentación ante la rama y la sección IEEE."],
      ["FaChalkboardTeacher", "Profesional fundador", "Apoyo para vincular a un docente o egresado con membresía profesional IEEE y Computer Society."],
      ["FaDoorOpen", "Espacios", "Salones o laboratorios para las reuniones de junta y las actividades del capítulo."],
      ["FaBullhorn", "Divulgación", "Difusión de la convocatoria y de las actividades por los canales de la Facultad."],
      ["FaClipboardCheck", "Permisos", "Agilización de los permisos para actividades externas y con instituciones de educación media."],
    ];
    const lw = 7.6;
    for (let i = 0; i < asks.length; i++) {
      const [ic, h, d] = asks[i];
      const y = 1.45 + i * 1.06;
      await iconCircle(s, ic, M, y + 0.08, 0.7, C.accent1, "FFFFFF", "icono-apoyo-" + (i + 1));
      txt(s, h, { x: M + 0.95, y: y + 0.02, w: lw - 1.0, h: 0.38, fontSize: 17, bold: true, color: C.text2 });
      txt(s, d, { x: M + 0.95, y: y + 0.4, w: lw - 1.0, h: 0.6, fontSize: 14 });
    }
    const rx = M + lw + 0.4, rw = W - M - rx;
    card(s, rx, 1.45, rw, 2.85, C.text2, "sin-presupuesto");
    txt(s, [
      { text: "0 USD", options: { bold: true, color: C.accent1, fontSize: 44, fontFace: "Cambria", breakLine: true } },
      { text: "de asignación presupuestal obligatoria para la Facultad: las membresías son individuales.", options: { color: C.background1, fontSize: 15 } },
    ], { x: rx + 0.35, y: 1.65, w: rw - 0.7, h: 2.5 });
    card(s, rx, 4.5, rw, 2.2, C.accent4, "apoyo-opcional");
    txt(s, [
      { text: "Apoyo opcional", options: { bold: true, color: C.text2, fontSize: 16, breakLine: true } },
      { text: "Si la Facultad lo considera viable, cubrir la membresía del profesional fundador (243 USD) acelera la conformación del grupo.", options: { color: C.text1, fontSize: 14 } },
    ], { x: rx + 0.35, y: 4.7, w: rw - 0.7, h: 1.9 });
    s.addNotes("La propuesta se sostiene sin presupuesto de la Facultad. El apoyo opcional al profesional es la palanca que más reduce el riesgo de no completar el grupo fundador.");
  }

  // ===== 11. Gobernanza =====
  pres.addSection({ title: "Ejecución" });
  {
    const s = content("Ejecución", "Gobernanza y articulación con la rama");
    const lv = [
      ["Sección IEEE", "Otorga el aval al nuevo capítulo", C.text2],
      ["Junta de la Rama IEEE USANTOTO", "Coordina las reuniones de rama · Consejero: Ing. Camilo Pardo", C.accent2],
      ["Junta del capítulo", "Cinco cargos con ambas membresías vigentes · periodo de un año", C.accent1],
    ];
    const lx = M, lw = 5.6;
    lv.forEach(([h, d, col], i) => {
      const y = 1.5 + i * 1.75;
      card(s, lx, y, lw, 1.35, col, "nivel-" + (i + 1));
      txt(s, h, { x: lx + 0.35, y: y + 0.18, w: lw - 0.7, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
      txt(s, d, { x: lx + 0.35, y: y + 0.65, w: lw - 0.7, h: 0.6, fontSize: 14, color: C.background1 });
      if (i < 2) s.addShape(pres.shapes.DOWN_ARROW, { x: lx + lw / 2 - 0.2, y: y + 1.4, w: 0.4, h: 0.32, fill: { color: C.accent3 }, line: { type: "none" }, rotate: 180, objectName: "flecha-" + (i + 1) });
    });
    const rx = lx + lw + 0.5, rw = W - M - rx;
    txt(s, "Cargos de la junta del capítulo", { x: rx, y: 1.5, w: rw, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    const roles = [
      ["Líder", "Representa al capítulo, convoca la junta y responde por el plan anual"],
      ["Vicepresidente", "Suple al líder y coordina la línea de formación técnica"],
      ["Secretario", "Actas, registro de asistencia y archivo documental"],
      ["Tesorero", "Recursos, estado de membresías, descuentos y rebates"],
      ["Webmaster", "Medios de divulgación, en coordinación con la rama"],
    ];
    roles.forEach(([r, d], i) => {
      const y = 2.1 + i * 0.92;
      card(s, rx, y, rw, 0.78, C.background2, "cargo-" + (i + 1));
      txt(s, r, { x: rx + 0.25, y, w: 1.9, h: 0.78, fontSize: 15, bold: true, color: C.accent1, valign: "middle" });
      txt(s, d, { x: rx + 2.2, y, w: rw - 2.4, h: 0.78, fontSize: 14, valign: "middle" });
    });
    s.addNotes("El capítulo opera bajo la identidad única de la Rama Estudiantil IEEE USANTOTO, no como un capítulo separado por facultad.");
  }

  // ===== 12. Ruta =====
  {
    const s = content("Ejecución", "Ruta de implementación");
    const ph = [
      ["1", "Conformación", "Convocar 20 aspirantes y 3 candidatos a profesional; sesión informativa; pago de membresías con descuento y rebates.", "7 membresías IEEE y Computer Society vigentes"],
      ["2", "Documentación", "Redactar propósito, cronograma, plan de reuniones y estructura de junta.", "Documento descriptivo con los 5 cargos asignados"],
      ["3", "Aval y registro", "Revisión con el consejero, aval de la sección y diligenciamiento del formulario.", "Aval de la sección y formulario radicado"],
      ["4", "Operación", "Ejecutar el plan: reuniones de junta, reuniones de rama y actividades interfacultades.", "Cronograma cumplido durante 12 meses"],
    ];
    const cw = (W - 2 * M - 3 * 0.25) / 4;
    ph.forEach(([n, h, a, c], i) => {
      const x = M + i * (cw + 0.25);
      s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x, y: 1.5, w: cw + 0.15, h: 0.9, fill: { color: i === 0 ? C.accent1 : C.text2 }, line: { type: "none" }, objectName: "fase-" + n });
      txt(s, "Fase " + n + " · " + h, { x: x + (i === 0 ? 0.3 : 0.5), y: 1.5, w: cw - 0.6, h: 0.9, fontSize: 15, bold: true, color: C.background1, valign: "middle" });
      card(s, x, 2.65, cw, 2.35, C.background2, "acciones-" + n);
      txt(s, a, { x: x + 0.25, y: 2.85, w: cw - 0.5, h: 2.0, fontSize: 14 });
      txt(s, "Criterio de cumplimiento", { x, y: 5.2, w: cw, h: 0.35, fontSize: 12, bold: true, color: C.accent3 });
      txt(s, c, { x, y: 5.55, w: cw, h: 0.8, fontSize: 14, bold: true, color: C.accent2 });
    });
    txt(s, "La fase 1 es la de mayor duración y condiciona las demás; avanza por hitos, no por fechas de cierre.", { x: M, y: 6.45, w: W - 2 * M, h: 0.35, fontSize: 14, italic: true, color: C.accent3 });
    s.addNotes("La convocatoria se dimensiona con margen: 20 estudiantes para obtener 6 y 3 candidatos para obtener 1 profesional, por la deserción esperable entre el interés y el pago.");
  }

  // ===== 13. Cronograma =====
  {
    const s = content("Ejecución", "Plan de actividades a doce meses");
    const per = [
      ["Instalación", 1, 1, C.accent1, "Instalación de junta; presentación ante la rama; plan operativo; apertura de medios"],
      ["Primer semestre", 2, 5, C.text2, "Jornada de vinculación; 1 actividad técnica; 1 actividad interfacultades; gestión de permisos"],
      ["Balance", 6, 6, C.accent3, "Balance semestral ante la rama; planeación; depuración de membresías"],
      ["Segundo semestre", 7, 10, C.text2, "2 actividades técnicas; 1 actividad interfacultades; Semana Tomasina; proyección externa"],
      ["Cierre", 11, 12, C.accent1, "Informe anual; verificación de membresías; nueva junta y empalme"],
    ];
    const tw = W - 2 * M, mw = tw / 12, ty = 1.55;
    for (let m = 1; m <= 12; m++) {
      txt(s, "M" + m, { x: M + (m - 1) * mw, y: ty, w: mw, h: 0.3, fontSize: 12, bold: true, color: C.accent3, align: "center" });
    }
    per.forEach(([h, a, b, col, d], i) => {
      const x = M + (a - 1) * mw + 0.04, w = (b - a + 1) * mw - 0.08;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ty + 0.4, w, h: 0.6, rectRadius: 0.1, fill: { color: col }, line: { type: "none" }, objectName: "periodo-" + (i + 1) });
      txt(s, h, { x: x + 0.08, y: ty + 0.4, w: w - 0.16, h: 0.6, fontSize: w < 1.2 ? 12 : 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    });
    // details cards
    const cw = (tw - 4 * 0.2) / 5;
    per.forEach(([h, a, b, col, d], i) => {
      const x = M + i * (cw + 0.2);
      card(s, x, 2.85, cw, 2.35, C.background2, "detalle-" + (i + 1));
      txt(s, (a === b ? "Mes " + a : "Meses " + a + "–" + b), { x: x + 0.2, y: 3.0, w: cw - 0.4, h: 0.35, fontSize: 13, bold: true, color: C.accent1 });
      txt(s, d, { x: x + 0.2, y: 3.4, w: cw - 0.4, h: 1.75, fontSize: 13 });
    });
    card(s, M, 5.45, tw, 1.25, C.text2, "permanentes");
    txt(s, [
      { text: "Durante los 12 meses: ", options: { bold: true, color: C.accent1, breakLine: true } },
      { text: "reunión mensual de junta con acta (incluidos los recesos) · participación en cada reunión de la rama · balance semestral de ejecución y membresías", options: { color: C.background1 } },
    ], { x: M + 0.35, y: 5.45, w: tw - 0.7, h: 1.25, fontSize: 15, valign: "middle" });
    s.addNotes("Los meses son relativos a la aprobación del capítulo. Los temas de cada actividad técnica los define la junta al inicio de cada semestre.");
  }

  // ===== 14. Riesgos =====
  {
    const s = content("Ejecución", "Riesgos y medidas de tratamiento");
    const hdr = { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 14, valign: "middle" };
    const rows = [
      [{ text: "Riesgo", options: hdr }, { text: "Efecto", options: hdr }, { text: "Medida de tratamiento", options: hdr }],
      ["No reunir los 7 miembros", "El formulario no puede radicarse", "Convocatoria con margen (20 + 3) y refuerzo con otros programas"],
      ["Costo de las membresías", "Desistimiento de aspirantes", "Comunicar el valor con descuento desde el primer contacto; priorizar afiliados de la rama (8 USD)"],
      ["Falta del profesional", "Capítulo incompleto", "Abordar 3 candidatos en paralelo, con apoyo del consejero y la Facultad"],
      ["Demora en permisos", "Retraso de actividades", "Tramitar al inicio de cada semestre; mantener alternativas internas"],
      ["Vencimiento de membresías", "Pérdida del mínimo exigido", "Seguimiento del tesorero y verificación en cada balance semestral"],
      ["Falta de continuidad", "Inactividad o pérdida del aval", "Actas mensuales, balances, informe anual y relevo de junta"],
    ];
    const body = rows.map((r, i) => i === 0 ? r : r.map((c, j) => ({ text: c, options: { fontSize: 14, bold: j === 0, color: j === 0 ? C.text2 : C.text1, fill: { color: i % 2 ? C.background2 : C.background1 }, valign: "middle" } })));
    s.addTable(body, { x: M, y: 1.5, w: W - 2 * M, colW: [3.0, 3.0, W - 2 * M - 6.0], rowH: 0.7, border: { type: "none" }, margin: [0.06, 0.15, 0.06, 0.15] });
    s.addNotes("El riesgo principal es no reunir oportunamente los siete miembros con ambas membresías vigentes.");
  }

  // ===== 15. Seguimiento =====
  {
    const s = content("Ejecución", "Indicadores de seguimiento");
    const kpis = [
      ["7", "membresías vigentes", "IEEE y Computer Society"],
      ["12", "actas de junta", "una por mes"],
      ["2", "balances semestrales", "ante la rama"],
      ["3", "actividades técnicas", "mínimo en el año"],
      ["2", "actividades interfacultades", "una por semestre"],
      ["1", "informe anual", "a la rama y a la Facultad"],
    ];
    const cw = (W - 2 * M - 2 * 0.3) / 3, ch = 2.2;
    kpis.forEach(([n, l, d], i) => {
      const x = M + (i % 3) * (cw + 0.3), y = 1.5 + Math.floor(i / 3) * (ch + 0.3);
      card(s, x, y, cw, ch, C.background2, "indicador-" + (i + 1));
      txt(s, n, { x: x + 0.3, y: y + 0.2, w: 1.2, h: 1.8, fontSize: 60, bold: true, color: i % 2 ? C.accent2 : C.accent1, fontFace: "Cambria", valign: "middle" });
      txt(s, l, { x: x + 1.6, y: y + 0.3, w: cw - 1.85, h: 0.85, fontSize: 17, bold: true, color: C.text2, valign: "bottom" });
      txt(s, d, { x: x + 1.6, y: y + 1.25, w: cw - 1.85, h: 0.7, fontSize: 14, color: C.accent3 });
    });
    txt(s, "La Facultad recibe el informe anual con la ejecución del plan, el estado de las membresías y la nueva junta.", { x: M, y: 6.5, w: W - 2 * M, h: 0.35, fontSize: 14, italic: true, color: C.accent3 });
    s.addNotes("Estos indicadores convierten el criterio de perdurabilidad del IEEE en evidencia verificable.");
  }

  // ===== 16. Cierre =====
  pres.addSection({ title: "Cierre" });
  {
    const s = pres.addSlide({ masterName: "CIERRE", sectionTitle: "Cierre" });
    s.addText("Decisión solicitada", { placeholder: "title" });
    txt(s, "Otorgar el aval de la Facultad de Ingeniería de Sistemas para la creación del Capítulo IEEE Computer Society de la Rama Estudiantil IEEE USANTOTO, con los apoyos institucionales descritos.",
      { x: M, y: 1.55, w: W - 2 * M, h: 1.0, fontSize: 19, color: C.accent5 });
    const steps = [
      ["FaCheckCircle", "Aval de la Facultad", "Decisión del Consejo de Facultad"],
      ["FaUsers", "Convocatoria", "20 aspirantes, 3 candidatos a profesional y sesión informativa"],
      ["FaIdCard", "Grupo fundador", "7 membresías vigentes y junta designada"],
      ["FaPaperPlane", "Radicación", "Revisión del consejero, aval de la sección y formulario ante el IEEE"],
    ];
    const cw = (W - 2 * M - 3 * 0.3) / 4;
    for (let i = 0; i < 4; i++) {
      const [ic, h, d] = steps[i];
      const x = M + i * (cw + 0.3);
      card(s, x, 3.0, cw, 3.4, "1E2F52", "paso-" + (i + 1));
      await iconCircle(s, ic, x + 0.35, 3.3, 0.8, C.accent1, "FFFFFF", "icono-paso-" + (i + 1));
      txt(s, (i + 1) + ". " + h, { x: x + 0.35, y: 4.3, w: cw - 0.7, h: 0.75, fontSize: 17, valign: "top", bold: true, color: C.background1 });
      txt(s, d, { x: x + 0.35, y: 5.1, w: cw - 0.7, h: 1.2, fontSize: 14, color: C.accent5 });
    }
    txt(s, "Santiago Alejandro Zúñiga Melo  ·  Rama Estudiantil IEEE USANTOTO", { x: M, y: 6.95, w: 9, h: 0.35, fontSize: 12, color: C.accent5 });
    s.addNotes("Con el aval, la fase de convocatoria puede iniciar de inmediato: las etapas de preparación y contacto directo no requieren trámites adicionales.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch(e => { console.error(e); process.exit(1); });

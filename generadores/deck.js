// Presentación editorial/académica — Capítulo IEEE Computer Society
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");
const OUT = process.argv[2] || "deck2.pptx";

const THEME = {
  name: "Editorial Computer Society",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "15171C", lt1: "FFFFFF", dk2: "101820", lt2: "F0F1F3",
    accent1: "E4472B", // bermellón
    accent2: "0E6B66", // verde petróleo
    accent3: "6B7280", // gris pizarra
    accent4: "FBE3DC", // bermellón muy claro
    accent5: "D7E8E6", // petróleo muy claro
    accent6: "C9A227", // ocre
    hlink: "0E6B66", folHlink: "6B7280",
  },
};
const HEX = THEME.colors;
const MONO = "Courier New";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Creación del Capítulo IEEE Computer Society";
pres.author = "Santiago Alejandro Zúñiga Melo";
pres.subject = "Propuesta ante el Consejo de la Facultad de Ingeniería de Sistemas";
const C = pres.SchemeColor;
const W = 13.333, H = 7.5, M = 0.7;

// ---------- network art (generative, seeded) ----------
function rng(seed) { return () => { seed |= 0; seed = (seed + 0x6D2B79F5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }
async function networkPng(wIn, hIn, seed, opts = {}) {
  const px = 200, Wp = Math.round(wIn * px), Hp = Math.round(hIn * px);
  const r = rng(seed), n = opts.n || 46;
  const pts = [];
  for (let i = 0; i < n; i++) pts.push([40 + r() * (Wp - 80), 40 + r() * (Hp - 80)]);
  let edges = "";
  pts.forEach((p, i) => {
    const d = pts.map((q, j) => [j, (q[0] - p[0]) ** 2 + (q[1] - p[1]) ** 2]).filter(x => x[0] !== i).sort((a, b) => a[1] - b[1]).slice(0, 3);
    d.forEach(([j]) => { if (j > i) edges += `<line x1="${p[0]}" y1="${p[1]}" x2="${pts[j][0]}" y2="${pts[j][1]}" stroke="${opts.line || "#FFFFFF"}" stroke-opacity="${opts.lo || 0.22}" stroke-width="3"/>`; });
  });
  let nodes = "";
  pts.forEach((p, i) => {
    const hot = i % 9 === 0, big = i % 5 === 0;
    const rad = hot ? 22 : big ? 13 : 8;
    const fill = hot ? "#" + HEX.accent1 : (opts.node || "#FFFFFF");
    const op = hot ? 1 : (opts.no || 0.55);
    nodes += `<circle cx="${p[0]}" cy="${p[1]}" r="${rad}" fill="${fill}" fill-opacity="${op}"/>`;
    if (hot) nodes += `<circle cx="${p[0]}" cy="${p[1]}" r="${rad + 16}" fill="none" stroke="#${HEX.accent1}" stroke-opacity="0.45" stroke-width="3"/>`;
  });
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${Wp}" height="${Hp}">${edges}${nodes}</svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- layouts ----------
const footerObjs = (col) => [
  { text: { text: "CAPÍTULO IEEE COMPUTER SOCIETY  ·  PROPUESTA 2026", options: { x: M, y: 7.02, w: 7, h: 0.25, fontSize: 9, fontFace: MONO, color: col, margin: 0, charSpacing: 1 } } },
];
pres.defineSlideMaster({
  title: "PORTADA", background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.7, w: 7, h: 0.35, fontSize: 11, fontFace: MONO, color: C.accent1, margin: 0, bold: true }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.0, w: 7.2, h: 2.6, fontSize: 50, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.75, w: 6.6, h: 0.9, fontSize: 18, color: "C9CED6", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "author", type: "body", x: M, y: 6.2, w: 7, h: 0.7, fontSize: 12, fontFace: MONO, color: C.background1, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENIDO", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.45, w: 9, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.8, w: W - 2 * M, h: 0.8, fontSize: 32, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
    ...footerObjs(HEX.accent3),
  ],
  slideNumber: { x: W - M - 0.8, y: 7.02, w: 0.8, h: 0.25, fontSize: 9, fontFace: MONO, color: C.accent3, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "DECLARACION", background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.6, w: 9, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.text2, margin: 0 }, text: "" } },
    ...footerObjs("FFFFFF"),
  ],
  slideNumber: { x: W - M - 0.8, y: 7.02, w: 0.8, h: 0.25, fontSize: 9, fontFace: MONO, color: C.background1, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "OSCURA", background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.45, w: 9, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.8, w: W - 2 * M, h: 0.8, fontSize: 32, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    ...footerObjs("9AA3AF"),
  ],
  slideNumber: { x: W - M - 0.8, y: 7.02, w: 0.8, h: 0.25, fontSize: 9, fontFace: MONO, color: "9AA3AF", align: "right", margin: 0 },
});

// ---------- helpers ----------
function T(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", fontSize: 15, color: C.text1 }, o)); }
function mono(s, text, o) { T(s, text, Object.assign({ fontFace: MONO, fontSize: 11, color: C.accent3 }, o)); }
function box(s, x, y, w, h, fill, id, extra = {}) { s.addShape(pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: id }, extra)); }
function dot(s, x, y, d, fill, id, line) { s.addShape(pres.shapes.OVAL, { x: x - d / 2, y: y - d / 2, w: d, h: d, fill: fill ? { color: fill } : { type: "none" }, line: line || { type: "none" }, objectName: id }); }
function seg(s, x1, y1, x2, y2, color, wpt = 1.5, id = "conector", dash) {
  s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), flipV: (x2 - x1) * (y2 - y1) < 0, line: { color, width: wpt, dashType: dash || "solid" }, objectName: id });
}
function source(s, text, color = C.accent3) { mono(s, "Fuente: " + text, { x: M, y: 6.62, w: W - 2 * M, h: 0.25, fontSize: 9, color }); }
function slide(master, section, kicker, title) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (kicker) s.addText(kicker, { placeholder: "kicker" });
  if (title) s.addText(title, { placeholder: "title" });
  return s;
}
const SRC_SAC = "reunión con el SAC y la directiva de la Rama Estudiantil IEEE USANTOTO, 18 de septiembre de 2026.";

(async () => {
  // ================= APERTURA =================
  pres.addSection({ title: "Apertura" });
  {
    const s = slide("PORTADA", "Apertura");
    s.addImage({ data: await networkPng(5.4, 7.5, 7), x: W - 5.4, y: 0, w: 5.4, h: 7.5, altText: "Red de nodos conectados, motivo gráfico de la propuesta", objectName: "red-portada" });
    s.addText("PROPUESTA  ·  CONSEJO DE LA FACULTAD DE INGENIERÍA DE SISTEMAS", { placeholder: "kicker" });
    s.addText("Capítulo IEEE Computer Society", { placeholder: "title" });
    s.addText("Creación del cuarto capítulo técnico de la Rama Estudiantil IEEE USANTOTO, Universidad Santo Tomás", { placeholder: "body" });
    s.addText([{ text: "Santiago Alejandro Zúñiga Melo", options: { bold: true, breakLine: true } }, { text: "Noviembre de 2026" }], { placeholder: "author" });
    s.addNotes("Propuesta de creación del Capítulo IEEE Computer Society. La red de nodos es el motivo de la presentación: la rama como red, y el capítulo como un nodo nuevo que conecta la Facultad de Ingeniería de Sistemas.");
  }
  {
    const s = slide("DECLARACION", "Apertura", "PUNTO DE PARTIDA");
    T(s, [
      { text: "3", options: { color: C.background1 } },
      { text: " capítulos técnicos.", options: { color: C.background1, fontSize: 48 } },
    ], { x: M, y: 1.55, w: 11.5, h: 1.6, fontSize: 110, bold: true, fontFace: "Cambria", valign: "bottom" });
    T(s, [
      { text: "0", options: { color: C.text2 } },
      { text: " dedicados a la computación.", options: { color: C.text2, fontSize: 48 } },
    ], { x: M, y: 3.15, w: 11.5, h: 1.6, fontSize: 110, bold: true, fontFace: "Cambria", valign: "bottom" });
    T(s, "La Rama Estudiantil IEEE USANTOTO tiene capítulos de aeronáutica, robótica y electrónica. La disciplina de la Facultad de Ingeniería de Sistemas no tiene un espacio propio dentro de la rama.",
      { x: M, y: 5.15, w: 9.2, h: 1.0, fontSize: 18, color: C.background1 });
    s.addNotes("Mensaje de apertura: existe un vacío. La rama ya funciona y crece, pero no hay un capítulo de computación.");
  }
  {
    const s = slide("CONTENIDO", "Apertura", "CONTENIDO", "Estructura de la propuesta");
    const items = [
      ["01", "Contexto", "La rama hoy y el vacío temático"],
      ["02", "Pertinencia", "Por qué crear el capítulo y a quién beneficia"],
      ["03", "Propuesta", "Propósito, requisitos del IEEE y gobernanza"],
      ["04", "Costos y recursos", "Qué cuesta, quién lo paga y qué se solicita"],
      ["05", "Implementación", "Ruta, convocatoria, cronograma, riesgos e indicadores"],
      ["06", "Decisión", "El aval que se solicita al Consejo"],
    ];
    items.forEach(([n, h, d], i) => {
      const col = i < 3 ? 0 : 1, row = i % 3;
      const x = M + col * 6.1, y = 2.0 + row * 1.5;
      T(s, n, { x, y, w: 1.3, h: 1.0, fontSize: 44, bold: true, fontFace: "Cambria", color: i === 5 ? C.accent1 : C.accent2 });
      T(s, h, { x: x + 1.4, y: y + 0.12, w: 4.4, h: 0.45, fontSize: 21, bold: true, color: C.text2 });
      T(s, d, { x: x + 1.4, y: y + 0.6, w: 4.4, h: 0.5, fontSize: 15, color: C.accent3 });
    });
    s.addNotes("Seis secciones, de la situación actual a la decisión.");
  }

  // ================= 01 CONTEXTO =================
  pres.addSection({ title: "01 Contexto" });
  {
    const s = slide("CONTENIDO", "01 Contexto", "§ 01  CONTEXTO", "La rama es una red activa con un nodo ausente");
    // network diagram
    const cx = 3.7, cy = 4.05, rx = 2.55, ry = 1.45;
    const sats = [
      ["CS", "Computer Society", "propuesto", 0, C.accent1, true],
      ["AESS", "Aeronáutica", "activo", 300, C.accent2],
      ["RAS", "Robótica", "activo", 240, C.accent2],
      ["WIE", "Mujer en la Ingeniería", "transversal", 180, C.accent3],
      ["EDS", "Electrónica", "menor actividad", 120, C.accent3],
      ["SOC", "Enfoque social", "transversal", 60, C.accent3],
    ];
    const pos = sats.map(sv => [cx + rx * Math.cos(sv[3] * Math.PI / 180), cy - ry * Math.sin(sv[3] * Math.PI / 180)]);
    sats.forEach((sv, i) => seg(s, cx, cy, pos[i][0], pos[i][1], sv[5] ? HEX.accent1 : "C5CAD1", sv[5] ? 2.25 : 1.5, "enlace-" + sv[0], sv[5] ? "dash" : "solid"));
    dot(s, cx, cy, 1.75, C.text2, "nodo-rama");
    T(s, "Rama\nIEEE\nUSANTOTO", { x: cx - 0.85, y: cy - 0.6, w: 1.7, h: 1.2, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    sats.forEach((sv, i) => {
      const [x, y] = pos[i], d = sv[5] ? 1.25 : 1.0;
      if (sv[5]) dot(s, x, y, d + 0.32, null, "halo-cs", { color: HEX.accent1, width: 1.5, dashType: "dash" });
      dot(s, x, y, d, sv[4], "nodo-" + sv[0]);
      T(s, sv[0], { x: x - d / 2, y: y - 0.2, w: d, h: 0.4, fontSize: sv[5] ? 18 : 14, bold: true, color: C.background1, align: "center", valign: "middle" });
      const below = y > cy + 0.5 || Math.abs(y - cy) < 0.3;
      const ly = below ? y + d / 2 + (sv[5] ? 0.22 : 0.06) : y - d / 2 - 0.62;
      T(s, [{ text: sv[1], options: { bold: true, breakLine: true } }, { text: sv[2], options: { fontFace: MONO, fontSize: 10, color: sv[5] ? C.accent1 : C.accent3 } }],
        { x: x - 1.1, y: ly, w: 2.2, h: 0.6, fontSize: 12, color: C.text2, align: "center" });
    });
    // right column
    const rx0 = 7.6, rw = W - M - rx0;
    T(s, "≈17", { x: rx0, y: 1.9, w: rw, h: 1.0, fontSize: 66, bold: true, fontFace: "Cambria", color: C.accent2 });
    T(s, "integrantes con membresía y afiliación a capítulo; las inscripciones del año superan las de periodos anteriores.", { x: rx0, y: 2.95, w: rw, h: 0.8, fontSize: 15 });
    T(s, [
      { text: "Sede: ", options: { bold: true } }, { text: "Facultad de Ingeniería Electrónica.", options: { breakLine: true } },
      { text: "Consejero: ", options: { bold: true } }, { text: "Ing. Camilo Pardo.", options: { breakLine: true } },
      { text: "Demanda: ", options: { bold: true } }, { text: "estudiantes de otros programas, entre ellos bioingeniería, han manifestado interés en la rama." },
    ], { x: rx0, y: 4.05, w: rw, h: 2.3, fontSize: 15, paraSpaceAfter: 8 });
    mono(s, "Fig. 1. Capítulos de la rama. Fuente: " + SRC_SAC, { x: M, y: 6.62, w: W - 2 * M, h: 0.25, fontSize: 9 });
    s.addNotes("Los capítulos técnicos son AESS, RAS y EDS; WIE y el capítulo social son transversales. El nodo punteado es el capítulo propuesto.");
  }

  // ================= 02 PERTINENCIA =================
  pres.addSection({ title: "02 Pertinencia" });
  {
    const s = slide("CONTENIDO", "02 Pertinencia", "§ 02  PERTINENCIA", "Tres argumentos para crear el capítulo");
    // left quote panel
    box(s, M, 1.85, 4.6, 4.55, C.text2, "panel-cita");
    T(s, "“", { x: M + 0.35, y: 1.9, w: 1, h: 1.0, fontSize: 80, bold: true, color: C.accent1, fontFace: "Cambria" });
    T(s, "El capítulo no crea una demanda: la canaliza. Estudiantes de otros programas ya buscan a la rama.", { x: M + 0.4, y: 2.85, w: 3.8, h: 2.3, fontSize: 22, italic: true, fontFace: "Cambria", color: C.background1 });
    mono(s, "— Jornadas de divulgación de la rama", { x: M + 0.4, y: 5.6, w: 3.8, h: 0.3, fontSize: 10, color: "9AA3AF" });
    const args = [
      ["01", "Pertinencia disciplinar", "Ningún capítulo de la rama atiende computación e ingeniería de software, el área central de la Facultad."],
      ["02", "Fortalecimiento de la rama", "Un cuarto capítulo aumenta afiliados, oferta de actividades y reconocimiento institucional."],
      ["03", "Proyección internacional", "Vincula a estudiantes y docentes con la IEEE Computer Society, la sociedad del IEEE dedicada a la computación."],
    ];
    const ax = M + 5.0, aw = W - M - ax;
    args.forEach(([n, h, d], i) => {
      const y = 1.85 + i * 1.55;
      T(s, n, { x: ax, y, w: 0.9, h: 0.6, fontSize: 30, bold: true, fontFace: "Cambria", color: C.accent1 });
      T(s, h, { x: ax + 1.0, y: y + 0.05, w: aw - 1.0, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
      T(s, d, { x: ax + 1.0, y: y + 0.5, w: aw - 1.0, h: 0.85, fontSize: 15 });
    });
    source(s, SRC_SAC);
    s.addNotes("El argumento central es la demanda existente: el capítulo responde a un interés ya manifestado.");
  }
  {
    const s = slide("CONTENIDO", "02 Pertinencia", "§ 02  PERTINENCIA", "Quién se beneficia y cómo");
    const cols = ["Estudiantes", "Facultad", "Universidad"];
    const rows = [
      ["Formación complementaria en software, IA, ciberseguridad y nube", [1, 1, 0]],
      ["Liderazgo: cinco cargos de junta con relevo anual", [1, 1, 0]],
      ["Acceso a recursos y red de la IEEE Computer Society", [1, 1, 0]],
      ["Trabajo interdisciplinario entre programas de ingeniería", [1, 0, 1]],
      ["Evidencia documentada: actas, balances e informe anual", [0, 1, 1]],
      ["Visibilidad ante el IEEE y su sección", [0, 1, 1]],
      ["Proyección: Semana Tomasina y educación media", [1, 0, 1]],
    ];
    const tx = M, tw = 7.4, c0 = tx + tw + 0.2, cw = (W - M - c0) / 3, y0 = 2.35, rh = 0.6;
    cols.forEach((c, j) => mono(s, c.toUpperCase(), { x: c0 + j * cw, y: 1.9, w: cw, h: 0.3, fontSize: 11, bold: true, color: C.text2, align: "center" }));
    rows.forEach(([t, v], i) => {
      const y = y0 + i * rh;
      if (i % 2 === 0) box(s, tx, y, W - 2 * M, rh, C.background2, "fila-" + (i + 1));
      T(s, t, { x: tx + 0.2, y, w: tw - 0.2, h: rh, fontSize: 15, valign: "middle" });
      v.forEach((on, j) => {
        const x = c0 + j * cw + cw / 2;
        dot(s, x, y + rh / 2, 0.26, on ? C.accent1 : null, `marca-${i + 1}-${j + 1}`, on ? undefined : { color: "C5CAD1", width: 1.25 });
      });
    });
    mono(s, "●  beneficio directo     ○  beneficio indirecto o no aplica", { x: M, y: 6.6, w: 8, h: 0.25, fontSize: 9 });
    s.addNotes("Matriz de beneficios. Los beneficios de membresía de la IEEE Computer Society deben confirmarse con la oferta vigente al afiliarse.");
  }

  // ================= 03 PROPUESTA =================
  pres.addSection({ title: "03 Propuesta" });
  {
    const s = slide("CONTENIDO", "03 Propuesta", "§ 03  PROPUESTA", "Propósito del capítulo");
    T(s, "Constituir un espacio permanente de formación, divulgación y proyección en computación e ingeniería de software, articulado con la rama y abierto a estudiantes de distintos programas de ingeniería.",
      { x: M, y: 1.8, w: 11.3, h: 1.6, fontSize: 24, italic: true, fontFace: "Cambria", color: C.text2 });
    const lines = [
      ["L1", "Formación técnica", "Línea principal. Talleres, charlas y prácticas en desarrollo de software, inteligencia artificial, ciberseguridad y computación en la nube."],
      ["L2", "Proyección y divulgación", "Actividades para otros programas de ingeniería y, con los permisos del caso, para instituciones de educación media."],
      ["L3", "Articulación institucional", "Programación expuesta en las reuniones de rama, Semana Tomasina y espacios de la sección IEEE."],
    ];
    const cw = (W - 2 * M - 2 * 0.4) / 3;
    lines.forEach(([c, h, d], i) => {
      const x = M + i * (cw + 0.4);
      box(s, x, 3.75, cw, 2.65, i === 0 ? C.accent5 : C.background2, "linea-" + c);
      mono(s, c, { x: x + 0.3, y: 3.95, w: 1, h: 0.3, fontSize: 13, bold: true, color: i === 0 ? C.accent2 : C.accent1 });
      T(s, h, { x: x + 0.3, y: 4.3, w: cw - 0.6, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
      T(s, d, { x: x + 0.3, y: 4.8, w: cw - 0.6, h: 1.5, fontSize: 14 });
    });
    s.addNotes("Toda actividad del cronograma se adscribe a una de las tres líneas: un programa articulado, no eventos aislados.");
  }
  {
    const s = slide("CONTENIDO", "03 Propuesta", "§ 03  PROPUESTA", "Lo que exige el IEEE para crear un capítulo");
    const reqs = [
      ["R1", "Grupo fundador", ["6 estudiantes con membresía IEEE y Computer Society", "1 profesional de grado Member o superior, con ambas membresías", "La junta cumple el mismo requisito"]],
      ["R2", "Documento descriptivo", ["Propósito del capítulo", "Cronograma a doce meses", "Reuniones de junta y con la rama", "Actividades interfacultades", "Junta de cinco cargos"]],
      ["R3", "Avales y radicación", ["Revisión del consejero de la rama", "Aval de la sección IEEE", "Formulario de creación vía SAC"]],
    ];
    const cw = (W - 2 * M - 2 * 0.4) / 3;
    reqs.forEach(([c, h, list], i) => {
      const x = M + i * (cw + 0.4);
      T(s, c, { x, y: 1.8, w: 1.2, h: 0.7, fontSize: 36, bold: true, fontFace: "Cambria", color: C.accent1 });
      T(s, h, { x: x + 1.15, y: 1.95, w: cw - 1.15, h: 0.45, fontSize: 19, bold: true, color: C.text2 });
      list.forEach((t, k) => {
        const y = 2.7 + k * 0.58;
        s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.07, w: 0.22, h: 0.22, fill: { type: "none" }, line: { color: HEX.accent2, width: 1.25 }, objectName: `casilla-${c}-${k + 1}` });
        T(s, t, { x: x + 0.38, y, w: cw - 0.4, h: 0.55, fontSize: 14 });
      });
    });
    box(s, M, 5.75, W - 2 * M, 0.7, C.accent4, "criterio");
    T(s, [{ text: "Criterio de evaluación: perdurabilidad.  ", options: { bold: true, color: C.accent1 } }, { text: "El capítulo debe demostrar programación sostenida y liderazgo definido.", options: { color: C.text2 } }],
      { x: M + 0.3, y: 5.75, w: W - 2 * M - 0.6, h: 0.7, fontSize: 16, valign: "middle" });
    source(s, SRC_SAC);
    s.addNotes("El grupo fundador es la condición que habilita todo lo demás.");
  }
  {
    const s = slide("CONTENIDO", "03 Propuesta", "§ 03  PROPUESTA", "Gobernanza: tres niveles, una sola identidad");
    const lv = [["Sección IEEE", "Otorga el aval al nuevo capítulo", C.text2], ["Junta de la Rama", "Coordina la programación · consejero: Ing. Camilo Pardo", C.accent2], ["Junta del capítulo", "Cinco cargos con ambas membresías · periodo de un año", C.accent1]];
    const nx = M + 0.5;
    seg(s, nx, 2.3, nx, 5.6, "C5CAD1", 2, "eje-gobernanza");
    lv.forEach(([h, d, col], i) => {
      const y = 2.3 + i * 1.65;
      dot(s, nx, y, 0.62, col, "nivel-" + (i + 1));
      mono(s, String(3 - i), { x: nx - 0.31, y: y - 0.15, w: 0.62, h: 0.3, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
      T(s, h, { x: nx + 0.6, y: y - 0.35, w: 4.4, h: 0.4, fontSize: 19, bold: true, color: C.text2 });
      T(s, d, { x: nx + 0.6, y: y + 0.08, w: 4.4, h: 0.6, fontSize: 14, color: C.accent3 });
    });
    const rx0 = 6.9, rw = W - M - rx0;
    mono(s, "JUNTA DEL CAPÍTULO", { x: rx0, y: 1.85, w: rw, h: 0.3, fontSize: 11, bold: true, color: C.text2 });
    const roles = [["Líder", "Representa al capítulo y responde por el plan anual"], ["Vicepresidente", "Suple al líder; coordina la formación técnica"], ["Secretario", "Actas, asistencia y archivo documental"], ["Tesorero", "Membresías, recursos, descuentos y rebates"], ["Webmaster", "Medios de divulgación con la rama"]];
    roles.forEach(([r, d], i) => {
      const y = 2.3 + i * 0.82;
      if (i % 2 === 0) box(s, rx0, y, rw, 0.74, C.background2, "cargo-" + (i + 1));
      T(s, r, { x: rx0 + 0.2, y, w: 1.9, h: 0.74, fontSize: 15, bold: true, color: C.accent1, valign: "middle" });
      T(s, d, { x: rx0 + 2.15, y, w: rw - 2.3, h: 0.74, fontSize: 14, valign: "middle" });
    });
    T(s, "El capítulo opera bajo la identidad de la Rama Estudiantil IEEE USANTOTO, no como un capítulo separado por facultad.", { x: M, y: 6.4, w: 5.8, h: 0.5, fontSize: 13, italic: true, color: C.accent3 });
    s.addNotes("Tres niveles: junta del capítulo, junta de la rama con su consejero, y la sección IEEE.");
  }

  // ================= 04 COSTOS =================
  pres.addSection({ title: "04 Costos y recursos" });
  {
    const s = slide("CONTENIDO", "04 Costos y recursos", "§ 04  COSTOS Y RECURSOS", "El costo de arranque cabe en una suma");
    box(s, M, 1.85, 5.2, 4.55, C.text2, "panel-suma");
    const L = [["6 × 35 USD", "estudiantes nuevos", "FFFFFF"], ["+ 1 × 243 USD", "profesional", "FFFFFF"]];
    L.forEach(([a, b, c], i) => {
      T(s, a, { x: M + 0.3, y: 2.15 + i * 0.95, w: 4.6, h: 0.6, fontSize: 30, bold: true, fontFace: MONO, color: C.background1, align: "right" });
      mono(s, b, { x: M + 0.3, y: 2.72 + i * 0.95, w: 4.6, h: 0.25, fontSize: 10, color: "9AA3AF", align: "right" });
    });
    seg(s, M + 0.9, 4.15, M + 4.9, 4.15, "FFFFFF", 1.5, "linea-suma");
    T(s, "= 453 USD", { x: M + 0.3, y: 4.3, w: 4.6, h: 0.75, fontSize: 40, bold: true, fontFace: MONO, color: C.accent1, align: "right" });
    T(s, "Máximo, sin descuentos. Lo pagan los afiliados.", { x: M + 0.3, y: 5.3, w: 4.6, h: 0.8, fontSize: 14, color: C.background1, align: "right" });
    // chart
    s.addChart(pres.charts.BAR, [
      { name: "Estudiantes", labels: ["Estudiantes nuevos", "Estudiantes de la rama"], values: [210, 48] },
      { name: "Profesional", labels: ["Estudiantes nuevos", "Estudiantes de la rama"], values: [243, 243] },
    ], {
      x: 6.3, y: 1.85, w: W - M - 6.3, h: 3.6, barDir: "bar", barGrouping: "stacked", barGapWidthPct: 60,
      chartColors: [HEX.accent2, HEX.accent1],
      showTitle: true, title: "Costo del grupo fundador por escenario (USD)", titleFontSize: 13, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt",
      catAxisLabelColor: HEX.accent3, valAxisLabelColor: HEX.accent3, catAxisLabelFontSize: 12, valAxisLabelFontSize: 10,
      catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
      valGridLine: { color: "E3E6EA", size: 0.5 }, catGridLine: { style: "none" }, valAxisMaxVal: 500,
      showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: HEX.accent3, legendFontFace: "+mn-lt",
    });
    T(s, [
      { text: "Lectura: ", options: { bold: true, color: C.accent1 } },
      { text: "el profesional concentra entre el 54 % y el 84 % del costo. Priorizar a los integrantes actuales de la rama (8 USD cada uno) lo reduce a 291 USD.", options: { color: C.text1 } },
    ], { x: 6.3, y: 5.55, w: W - M - 6.3, h: 0.85, fontSize: 14 });
    source(s, SRC_SAC + " Tarifas sujetas a verificación; descuento institucional y rebates por confirmar.");
    s.addNotes("Membresía estudiante: 27 USD IEEE + 8 USD Computer Society. Profesional: 183 + 60. Integrante actual de la rama: solo 8 USD.");
  }
  {
    const s = slide("CONTENIDO", "04 Costos y recursos", "§ 04  COSTOS Y RECURSOS", "Lo que se solicita a la Facultad");
    const asks = [
      ["Aval institucional", "Respaldo para crear el capítulo y presentarlo ante la rama y la sección IEEE"],
      ["Profesional fundador", "Apoyo para vincular a un docente o egresado con membresía profesional"],
      ["Espacios", "Salones o laboratorios para reuniones de junta y actividades"],
      ["Divulgación", "Difusión de la convocatoria y las actividades por canales de la Facultad"],
      ["Permisos", "Agilización de permisos para actividades externas y con colegios"],
    ];
    asks.forEach(([h, d], i) => {
      const y = 1.85 + i * 0.92;
      mono(s, String(i + 1).padStart(2, "0"), { x: M, y: y + 0.05, w: 0.6, h: 0.35, fontSize: 16, bold: true, color: C.accent1 });
      T(s, h, { x: M + 0.75, y, w: 6.3, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
      T(s, d, { x: M + 0.75, y: y + 0.4, w: 6.3, h: 0.45, fontSize: 14, color: C.accent3 });
    });
    const rx0 = 8.1, rw = W - M - rx0;
    box(s, rx0, 1.85, rw, 2.6, C.text2, "panel-presupuesto");
    mono(s, "PRESUPUESTO OBLIGATORIO", { x: rx0 + 0.35, y: 2.05, w: rw - 0.7, h: 0.3, fontSize: 11, bold: true, color: "9AA3AF" });
    T(s, "0 USD", { x: rx0 + 0.35, y: 2.4, w: rw - 0.7, h: 0.95, fontSize: 52, bold: true, fontFace: "Cambria", color: C.accent1 });
    T(s, "Las membresías son individuales.", { x: rx0 + 0.35, y: 3.5, w: rw - 0.7, h: 0.7, fontSize: 15, color: C.background1 });
    box(s, rx0, 4.65, rw, 1.75, C.accent4, "panel-opcional");
    mono(s, "APOYO OPCIONAL", { x: rx0 + 0.35, y: 4.82, w: rw - 0.7, h: 0.3, fontSize: 11, bold: true, color: C.accent1 });
    T(s, "Cubrir la membresía del profesional fundador (243 USD) elimina el mayor obstáculo de la convocatoria.", { x: rx0 + 0.35, y: 5.15, w: rw - 0.7, h: 1.15, fontSize: 14, color: C.text2 });
    s.addNotes("La propuesta se sostiene sin presupuesto de la Facultad. El apoyo opcional ataca directamente el componente más costoso.");
  }

  // ================= 05 IMPLEMENTACIÓN =================
  pres.addSection({ title: "05 Implementación" });
  {
    const s = slide("CONTENIDO", "05 Implementación", "§ 05  IMPLEMENTACIÓN", "Ruta en cuatro fases, por hitos y no por fechas");
    const ph = [
      ["Conformación", "Convocatoria, sesión informativa y pago de membresías", "7 membresías vigentes", true],
      ["Documentación", "Propósito, cronograma, reuniones y junta", "Documento descriptivo completo"],
      ["Aval y registro", "Consejero, aval de la sección y formulario", "Formulario radicado"],
      ["Operación", "Reuniones de junta y de rama, actividades interfacultades", "12 meses de cronograma cumplido"],
    ];
    const cw = (W - 2 * M) / 4, ly = 2.45;
    seg(s, M + cw / 2, ly, M + 3.5 * cw, ly, "C5CAD1", 2, "eje-ruta");
    ph.forEach(([h, a, c, hot], i) => {
      const x = M + i * cw, cx = x + cw / 2;
      dot(s, cx, ly, 0.85, hot ? C.accent1 : C.text2, "fase-" + (i + 1));
      T(s, String(i + 1), { x: cx - 0.425, y: ly - 0.3, w: 0.85, h: 0.6, fontSize: 24, bold: true, fontFace: "Cambria", color: C.background1, align: "center", valign: "middle" });
      if (hot) mono(s, "LA MÁS LARGA", { x: cx - 1, y: 1.72, w: 2, h: 0.25, fontSize: 9, bold: true, color: C.accent1, align: "center" });
      T(s, h, { x: x + 0.15, y: 3.15, w: cw - 0.3, h: 0.45, fontSize: 19, bold: true, color: C.text2, align: "center" });
      T(s, a, { x: x + 0.25, y: 3.65, w: cw - 0.5, h: 1.1, fontSize: 14, color: C.text1, align: "center" });
      box(s, x + 0.2, 4.95, cw - 0.4, 1.0, i === 0 ? C.accent4 : C.background2, "criterio-" + (i + 1));
      mono(s, "CRITERIO", { x: x + 0.35, y: 5.05, w: cw - 0.7, h: 0.22, fontSize: 9, bold: true, color: C.accent3, align: "center" });
      T(s, c, { x: x + 0.35, y: 5.3, w: cw - 0.7, h: 0.6, fontSize: 14, bold: true, color: i === 0 ? C.accent1 : C.accent2, align: "center" });
    });
    T(s, "La fase 1 condiciona las demás; una vez completa, las fases 2 a 4 avanzan con rapidez.", { x: M, y: 6.2, w: W - 2 * M, h: 0.35, fontSize: 14, italic: true, color: C.accent3 });
    s.addNotes("Responsables propuestos: promotores con apoyo de la rama (fase 1), junta propuesta (fase 2), líder con el consejero (fase 3), junta del capítulo (fase 4).");
  }
  {
    const s = slide("CONTENIDO", "05 Implementación", "§ 05  IMPLEMENTACIÓN", "Convocar con margen: la deserción es previsible");
    const fun = (x0, w0, title, rows) => {
      mono(s, title, { x: x0, y: 1.85, w: w0, h: 0.3, fontSize: 11, bold: true, color: C.text2 });
      rows.forEach(([n, lab, frac, col], i) => {
        const y = 2.3 + i * 1.3, bw = w0 * frac;
        box(s, x0 + (w0 - bw) / 2, y, bw, 0.95, col, "embudo-" + title + "-" + (i + 1));
        T(s, n, { x: x0 + (w0 - bw) / 2, y, w: bw, h: 0.95, fontSize: n === "filtro" ? 16 : 30, fontFace: n === "filtro" ? MONO : "Cambria", bold: true, color: C.background1, align: "center", valign: "middle" });
        mono(s, lab, { x: x0, y: y + 0.97, w: w0, h: 0.2, fontSize: 9, color: C.accent3, align: "center" });
      });
    };
    fun(M, 5.4, "ESTUDIANTES", [["20", "convocados a la sesión informativa", 1, C.text2], ["filtro", "compromiso y capacidad de pago", 0.62, C.accent2], ["6", "membresías formalizadas", 0.32, C.accent1]]);
    fun(M + 6.0, 2.6, "PROFESIONALES", [["3", "candidatos en paralelo", 1, C.text2], ["filtro", "decisión del candidato", 0.7, C.accent2], ["1", "membresía profesional", 0.45, C.accent1]]);
    const rx0 = M + 9.0, rw = W - M - rx0;
    T(s, "Mensaje de la convocatoria", { x: rx0, y: 1.85, w: rw, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    T(s, [
      { text: "Qué es el capítulo y su primer año", options: { bullet: { indent: 14 }, breakLine: true } },
      { text: "Qué compromiso se pide", options: { bullet: { indent: 14 }, breakLine: true } },
      { text: "Valor exacto con descuento", options: { bullet: { indent: 14 }, breakLine: true } },
      { text: "Quién acompaña el pago: la rama", options: { bullet: { indent: 14 } } },
    ], { x: rx0, y: 2.35, w: rw, h: 2.4, fontSize: 14, paraSpaceAfter: 8 });
    T(s, "La falta de claridad sobre el costo es el principal factor de deserción.", { x: rx0, y: 4.9, w: rw, h: 1.0, fontSize: 13, italic: true, color: C.accent1 });
    source(s, "Plan de acción para la conformación del grupo fundador, octubre de 2026.");
    s.addNotes("La convocatoria no tiene fecha de cierre: se mantiene abierta hasta completar el grupo, con una segunda ronda prevista como parte normal del proceso.");
  }
  {
    const s = slide("CONTENIDO", "05 Implementación", "§ 05  IMPLEMENTACIÓN", "Primer año de operación");
    const lx = M, lw = 3.4, gx = lx + lw, gw = W - M - gx, mw = gw / 12, y0 = 2.25, rh = 0.46;
    const periods = [["Inst.", 1, 1], ["Primer semestre", 2, 5], ["Bal.", 6, 6], ["Segundo semestre", 7, 10], ["Cierre", 11, 12]];
    periods.forEach(([p, a, b], i) => {
      box(s, gx + (a - 1) * mw + 0.02, 1.75, (b - a + 1) * mw - 0.04, 0.3, i % 2 ? C.background2 : C.accent5, "periodo-" + (i + 1));
      mono(s, p, { x: gx + (a - 1) * mw, y: 1.75, w: (b - a + 1) * mw, h: 0.3, fontSize: 9, bold: true, color: C.text2, align: "center", valign: "middle" });
    });
    for (let m = 1; m <= 12; m++) mono(s, "M" + m, { x: gx + (m - 1) * mw, y: 2.05, w: mw, h: 0.2, fontSize: 9, color: C.accent3, align: "center" });
    const rows = [
      ["Reunión de junta con acta", "dots"],
      ["Instalación y plan operativo", [[1, 1, C.accent1]]],
      ["Vinculación de afiliados", [[2, 5, C.accent2]]],
      ["Formación técnica", [[2, 5, C.accent1, "×1"], [7, 10, C.accent1, "×2"]]],
      ["Actividad interfacultades", [[2, 5, C.accent2, "×1"], [7, 10, C.accent2, "×1"]]],
      ["Semana Tomasina y proyección", [[7, 10, C.text2]]],
      ["Balance ante la rama", [[6, 6, C.accent6], [11, 12, C.accent6, "informe"]]],
      ["Relevo de junta y empalme", [[11, 12, C.accent1]]],
    ];
    rows.forEach(([lab, bars], i) => {
      const y = y0 + 0.15 + i * rh;
      if (i % 2 === 0) box(s, lx, y, W - 2 * M, rh, C.background2, "fila-gantt-" + (i + 1));
      T(s, lab, { x: lx + 0.15, y, w: lw - 0.2, h: rh, fontSize: 13, valign: "middle", bold: i === 0 });
      if (bars === "dots") { for (let m = 1; m <= 12; m++) dot(s, gx + (m - 0.5) * mw, y + rh / 2, 0.18, C.text2, "acta-m" + m); }
      else bars.forEach(([a, b, col, tag], k) => {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + (a - 1) * mw + 0.06, y: y + 0.1, w: (b - a + 1) * mw - 0.12, h: rh - 0.2, rectRadius: 0.06, fill: { color: col }, line: { type: "none" }, objectName: `barra-${i + 1}-${k + 1}` });
        if (tag) mono(s, tag, { x: gx + (a - 1) * mw, y: y + 0.1, w: (b - a + 1) * mw, h: rh - 0.2, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle" });
      });
    });
    T(s, "Meses contados desde la aprobación. Participación en todas las reuniones de la rama; su periodicidad se confirmará con la directiva.", { x: M, y: 6.35, w: W - 2 * M, h: 0.3, fontSize: 12, italic: true, color: C.accent3 });
    source(s, "Propuesta base del plan a doce meses, octubre de 2026.");
    s.addNotes("La cantidad de actividades es un mínimo deliberado: se amplía solo con capacidad verificada, para no comprometer la continuidad.");
  }
  {
    const s = slide("CONTENIDO", "05 Implementación", "§ 05  IMPLEMENTACIÓN", "Riesgos identificados y su tratamiento");
    const rows = [
      ["No reunir los 7 miembros", "Formulario sin radicar", "Convocatoria con margen (20 + 3) y segunda ronda con otros programas"],
      ["Costo de las membresías", "Desistimiento", "Valor con descuento desde el primer contacto; priorizar integrantes de la rama"],
      ["Falta del profesional", "Capítulo incompleto", "Tres candidatos en paralelo, con apoyo del consejero y la Facultad"],
      ["Demora en permisos", "Retraso de actividades", "Trámite al inicio de cada semestre y alternativas internas"],
      ["Vencimiento de membresías", "Pérdida del mínimo", "Seguimiento del tesorero y revisión en cada balance"],
      ["Falta de continuidad", "Pérdida del aval", "Actas mensuales, balances, informe anual y relevo de junta"],
    ];
    mono(s, "RIESGO", { x: M + 0.55, y: 1.85, w: 3.5, h: 0.25, fontSize: 10, bold: true, color: C.text2 });
    mono(s, "EFECTO", { x: M + 4.2, y: 1.85, w: 2.6, h: 0.25, fontSize: 10, bold: true, color: C.text2 });
    mono(s, "TRATAMIENTO", { x: M + 7.0, y: 1.85, w: 4.9, h: 0.25, fontSize: 10, bold: true, color: C.text2 });
    rows.forEach(([r, e, t], i) => {
      const y = 2.2 + i * 0.7;
      if (i % 2 === 0) box(s, M, y, W - 2 * M, 0.7, C.background2, "riesgo-" + (i + 1));
      dot(s, M + 0.25, y + 0.35, 0.2, i === 0 ? C.accent1 : C.accent3, "indicador-riesgo-" + (i + 1));
      T(s, r, { x: M + 0.55, y, w: 3.5, h: 0.7, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      T(s, e, { x: M + 4.2, y, w: 2.6, h: 0.7, fontSize: 14, color: C.accent3, valign: "middle" });
      T(s, t, { x: M + 7.0, y, w: W - 2 * M - 7.2, h: 0.7, fontSize: 14, valign: "middle" });
    });
    source(s, SRC_SAC);
    s.addNotes("El riesgo principal, marcado en bermellón, es no completar el grupo fundador.");
  }
  {
    const s = slide("OSCURA", "05 Implementación", "§ 05  IMPLEMENTACIÓN", "Cómo se medirá la continuidad");
    const k = [["7", "membresías vigentes", "todo el periodo"], ["12", "actas de junta", "una por mes"], ["2", "balances", "uno por semestre"], ["≥3", "actividades técnicas", "en el año"], ["≥2", "interfacultades", "una por semestre"], ["1", "informe anual", "a la rama y la Facultad"]];
    const cw = (W - 2 * M) / 6;
    k.forEach(([n, l, d], i) => {
      const x = M + i * cw;
      T(s, n, { x, y: 2.4, w: cw - 0.1, h: 1.4, fontSize: 64, bold: true, fontFace: "Cambria", color: i % 2 ? C.background1 : C.accent1 });
      T(s, l, { x, y: 3.9, w: cw - 0.2, h: 0.7, fontSize: 16, bold: true, color: C.background1 });
      mono(s, d, { x, y: 4.6, w: cw - 0.2, h: 0.5, fontSize: 10, color: "9AA3AF" });
    });
    T(s, "Estos indicadores convierten el criterio de perdurabilidad del IEEE en evidencia verificable para la rama, la sección y la Facultad.", { x: M, y: 5.6, w: 10.5, h: 0.8, fontSize: 17, italic: true, fontFace: "Cambria", color: "C9CED6" });
    s.addNotes("La Facultad recibe el informe anual con la ejecución del plan, el estado de las membresías y la nueva junta.");
  }

  // ================= 06 DECISIÓN =================
  pres.addSection({ title: "06 Decisión" });
  {
    const s = slide("OSCURA", "06 Decisión", "§ 06  DECISIÓN");
    s.addImage({ data: await networkPng(4.6, 7.5, 21, { n: 34 }), x: W - 4.6, y: 0, w: 4.6, h: 7.5, altText: "Red de nodos conectados", objectName: "red-cierre" });
    T(s, "Se solicita el aval de la Facultad de Ingeniería de Sistemas para crear el Capítulo IEEE Computer Society.", { x: M, y: 0.95, w: 7.9, h: 2.1, fontSize: 34, bold: true, fontFace: "Cambria", color: C.background1 });
    const steps = [["Aval de la Facultad", "Decisión de este Consejo"], ["Convocatoria", "20 aspirantes y 3 candidatos a profesional"], ["Grupo fundador", "7 membresías vigentes y junta designada"], ["Radicación", "Consejero, sección IEEE y formulario"]];
    steps.forEach(([h, d], i) => {
      const y = 3.45 + i * 0.82;
      dot(s, M + 0.25, y + 0.25, 0.5, i === 0 ? C.accent1 : null, "paso-" + (i + 1), i === 0 ? undefined : { color: "9AA3AF", width: 1.5 });
      T(s, String(i + 1), { x: M, y: y, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle" });
      if (i < 3) seg(s, M + 0.25, y + 0.5, M + 0.25, y + 0.82, "9AA3AF", 1.25, "union-" + (i + 1));
      T(s, h, { x: M + 0.75, y: y + 0.02, w: 2.6, h: 0.46, fontSize: 17, bold: true, color: i === 0 ? C.accent1 : C.background1, valign: "middle" });
      T(s, d, { x: M + 3.4, y: y + 0.02, w: 4.6, h: 0.46, fontSize: 14, color: "C9CED6", valign: "middle" });
    });
    mono(s, "Santiago Alejandro Zúñiga Melo  ·  Rama Estudiantil IEEE USANTOTO", { x: M, y: 6.75, w: 7.5, h: 0.25, fontSize: 9, color: "9AA3AF" });
    s.addNotes("Con el aval, la convocatoria inicia de inmediato: la preparación y el contacto directo no requieren otros trámites.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch(e => { console.error(e); process.exit(1); });

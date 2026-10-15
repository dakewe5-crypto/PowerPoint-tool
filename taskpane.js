/* Folienbau – Office.js Task-Pane-Add-in für PowerPoint
 * Alle Maße in Punkt (pt), Eingabe/Anzeige in cm. */

/* ---------- Konfiguration (hier anpassen) ---------- */
const CFG = {
  // Firmenpalette: [Name, Hex]. null = "keine Farbe". Maximal 9 Einträge passen in eine Zeile.
  palette: [
    ["Nachtblau", "#1F3A5F"],
    ["Blau", "#2F6FDB"],
    ["Türkis", "#1A9E8F"],
    ["Gelb", "#E0A526"],
    ["Orange", "#D9552B"],
    ["Grau", "#6B7280"],
    ["Weiß", "#FFFFFF"],
    ["Schwarz", "#000000"],
    ["Keine Farbe", null]
  ],
  defaultGapCm: 0.25,
  stepCm: 0.1 // Schrittweite der Pfeile bei "Größe setzen"
};

const CM = 28.3465;
const FILL_TYPES = ["GeometricShape", "TextBox", "Freeform", "Callout", "Placeholder"];
const TEXT_TYPES = ["GeometricShape", "TextBox", "Freeform", "Callout", "Placeholder"];
const LINE_TYPES = ["GeometricShape", "TextBox", "Freeform", "Callout", "Placeholder", "Line", "Image"];

const state = { ref: "first", target: "fill", lock: true, ratio: 1, selKey: "" };

/* ---------- Icons (SVG-Innenleben, 20x20) ---------- */
const I = {
  left: '<path d="M4 3v14"/><rect x="6" y="5" width="10" height="4"/><rect x="6" y="11" width="6" height="4"/>',
  centerH: '<path d="M10 3v14"/><rect x="4" y="5" width="12" height="4"/><rect x="6" y="11" width="8" height="4"/>',
  right: '<path d="M16 3v14"/><rect x="4" y="5" width="10" height="4"/><rect x="8" y="11" width="6" height="4"/>',
  top: '<path d="M3 4h14"/><rect x="5" y="6" width="4" height="10"/><rect x="11" y="6" width="4" height="6"/>',
  middle: '<path d="M3 10h14"/><rect x="5" y="4" width="4" height="12"/><rect x="11" y="6" width="4" height="8"/>',
  bottom: '<path d="M3 16h14"/><rect x="5" y="4" width="4" height="10"/><rect x="11" y="8" width="4" height="6"/>',
  dockL: '<rect x="3" y="7" width="4" height="6"/><rect x="11" y="4" width="6" height="12"/><path d="M8 10h2"/>',
  dockR: '<rect x="13" y="7" width="4" height="6"/><rect x="3" y="4" width="6" height="12"/><path d="M10 10h2"/>',
  dockT: '<rect x="7" y="3" width="6" height="4"/><rect x="4" y="11" width="12" height="6"/><path d="M10 8v2"/>',
  dockB: '<rect x="7" y="13" width="6" height="4"/><rect x="4" y="3" width="12" height="6"/><path d="M10 10v2"/>',
  stackH: '<rect x="2" y="6" width="4" height="8"/><rect x="7" y="6" width="4" height="8"/><rect x="12" y="6" width="4" height="8"/><path d="M18 4v12"/>',
  stackV: '<rect x="6" y="2" width="8" height="4"/><rect x="6" y="7" width="8" height="4"/><rect x="6" y="12" width="8" height="4"/><path d="M4 18h12"/>',
  distH: '<path d="M3 3v14M17 3v14"/><rect x="8" y="6" width="4" height="8"/>',
  distV: '<path d="M3 3h14M3 17h14"/><rect x="6" y="8" width="8" height="4"/>',
  swap: '<rect x="3" y="3" width="5" height="5"/><rect x="12" y="12" width="5" height="5"/><path d="M10 5h4v4M14 9l-2-2M10 15H6v-4M6 11l2 2"/>',
  grid: '<rect x="3" y="3" width="6" height="6"/><rect x="11" y="3" width="6" height="6"/><rect x="3" y="11" width="6" height="6"/><rect x="11" y="11" width="6" height="6"/>',
  slideH: '<rect x="2" y="4" width="16" height="12"/><path d="M10 6v8"/><rect x="7" y="8" width="6" height="4"/>',
  slideV: '<rect x="2" y="4" width="16" height="12"/><path d="M4 10h12"/><rect x="8" y="7" width="4" height="6"/>',
  sameW: '<path d="M3 10h14M6 7l-3 3 3 3M14 7l3 3-3 3"/>',
  sameH: '<path d="M10 3v14M7 6l3-3 3 3M7 14l3 3 3-3"/>',
  sameWH: '<rect x="4" y="4" width="12" height="12"/><path d="M7 13l6-6M9 7h4v4"/>',
  square: '<rect x="5" y="5" width="10" height="10"/><path d="M2 10h2M16 10h2M10 2v2M10 16v2"/>',
  swapWH: '<rect x="2" y="9" width="10" height="5"/><rect x="14" y="4" width="4" height="10"/><path d="M4 5h5M7 3l2 2-2 2"/>',
  fitText: '<rect x="2" y="2" width="16" height="16" stroke-dasharray="2 2"/><path d="M6 14l4-9 4 9M7.5 11h5"/>',
  fullW: '<rect x="2" y="4" width="16" height="12"/><path d="M4 10h12M6 8l-2 2 2 2M14 8l2 2-2 2"/>',
  fullH: '<rect x="2" y="4" width="16" height="12"/><path d="M10 6v8M8 8l2-2 2 2M8 12l2 2 2-2"/>',
  fullAll: '<rect x="2" y="4" width="16" height="12"/><rect class="fill" x="5" y="7" width="10" height="6"/>',
  plus: '<path d="M10 4v12M4 10h12"/>',
  minus: '<path d="M4 10h12"/>',
  aspect: '<rect x="3" y="6" width="9" height="8"/><rect x="13" y="4" width="4" height="12" stroke-dasharray="2 2"/><path d="M5 17h12"/>',
  toFront: '<rect x="3" y="3" width="10" height="10"/><rect class="bg fill" x="7" y="7" width="10" height="10"/><path d="M10 12l3-3 3 3"/>',
  toBack: '<rect class="fill" x="7" y="7" width="10" height="10"/><rect class="bg" x="3" y="3" width="10" height="10" stroke-dasharray="2 2"/><path d="M10 12l3 3 3-3"/>',
  forward: '<rect x="3" y="3" width="9" height="9"/><rect class="bg fill" x="8" y="8" width="9" height="9"/><path d="M12 15v-4M10 13l2-2 2 2"/>',
  backward: '<rect class="fill" x="8" y="8" width="9" height="9"/><rect class="bg" x="3" y="3" width="9" height="9"/><path d="M7 5v4M5 7l2 2 2-2"/>',
  group: '<rect x="2" y="2" width="16" height="16" stroke-dasharray="2 2"/><rect x="5" y="5" width="5" height="5"/><rect x="10" y="10" width="5" height="5"/>',
  ungroup: '<rect x="2" y="2" width="7" height="7"/><rect x="11" y="11" width="7" height="7"/><path d="M12 8l4-4M13 4h3v3"/>',
  font: '<path d="M5 16L10 4l5 12M7 12h6"/>',
  paint: '<rect x="4" y="3" width="11" height="5"/><path d="M15 5h2v5H9v3M9 13v4"/>',
  source: '<path d="M4 6h12M4 10h12M4 14h7"/>',
  draft: '<rect x="3" y="5" width="14" height="10" rx="2"/><path d="M6 10h8"/>',
  secret: '<rect x="4" y="9" width="12" height="8" rx="1"/><path d="M7 9V7a3 3 0 016 0v2"/>',
  line: '<path d="M3 10h14"/><path d="M3 7v6M17 7v6"/>'
};

/* ---------- Button-Definitionen ---------- */
/* Jeder Eintrag: id, group, cls (Farbe), icon, title, tip, label (optional), min (Mindestanzahl Objekte | "ref"), fn */
const BTN = [];
function def(group, cls, id, icon, title, tip, min, fn, label) {
  BTN.push({ group, cls, id, icon, title, tip, min, fn, label });
}

/* Ausrichten */
def("align", "c-align", "alLeft", I.left, "Linksbündig", "Richtet die linken Kanten an der Referenz aus.", "ref", (c) => alignAll(c, (s, R) => { s.left = R.left; }));
def("align", "c-align", "alCenterH", I.centerH, "Horizontal zentriert", "Zentriert die Objekte waagerecht zur Mitte der Referenz.", "ref", (c) => alignAll(c, (s, R) => { s.left = R.left + R.width / 2 - s.width / 2; }));
def("align", "c-align", "alRight", I.right, "Rechtsbündig", "Richtet die rechten Kanten an der Referenz aus.", "ref", (c) => alignAll(c, (s, R) => { s.left = R.left + R.width - s.width; }));
def("align", "c-align", "alTop", I.top, "Oben ausrichten", "Richtet die Oberkanten an der Referenz aus.", "ref", (c) => alignAll(c, (s, R) => { s.top = R.top; }));
def("align", "c-align", "alMiddle", I.middle, "Vertikal zentriert", "Zentriert die Objekte senkrecht zur Mitte der Referenz.", "ref", (c) => alignAll(c, (s, R) => { s.top = R.top + R.height / 2 - s.height / 2; }));
def("align", "c-align", "alBottom", I.bottom, "Unten ausrichten", "Richtet die Unterkanten an der Referenz aus.", "ref", (c) => alignAll(c, (s, R) => { s.top = R.top + R.height - s.height; }));

/* Andocken */
def("dock", "c-dock", "dkLeft", I.dockL, "Links andocken", "Setzt das Objekt links neben die Referenz, mit dem eingestellten Abstand.", 2, (c) => dockAll(c, (s, R, g) => { s.left = R.left - s.width - g; }));
def("dock", "c-dock", "dkRight", I.dockR, "Rechts andocken", "Setzt das Objekt rechts neben die Referenz, mit dem eingestellten Abstand.", 2, (c) => dockAll(c, (s, R, g) => { s.left = R.left + R.width + g; }));
def("dock", "c-dock", "dkTop", I.dockT, "Oben andocken", "Setzt das Objekt über die Referenz, mit dem eingestellten Abstand.", 2, (c) => dockAll(c, (s, R, g) => { s.top = R.top - s.height - g; }));
def("dock", "c-dock", "dkBottom", I.dockB, "Unten andocken", "Setzt das Objekt unter die Referenz, mit dem eingestellten Abstand.", 2, (c) => dockAll(c, (s, R, g) => { s.top = R.top + R.height + g; }));
def("dock", "c-dock", "stH", I.stackH, "Nebeneinander stapeln", "Reiht alle Objekte von links nach rechts mit festem Abstand auf, ohne Lücken.", 2, (c) => stack(c, true));
def("dock", "c-dock", "stV", I.stackV, "Untereinander stapeln", "Reiht alle Objekte von oben nach unten mit festem Abstand auf, ohne Lücken.", 2, (c) => stack(c, false));

/* Anordnen / Verteilen */
def("arrange", "c-arrange", "dsH", I.distH, "Horizontal verteilen", "Verteilt die Objekte mit gleichen Abständen zwischen dem ersten und letzten (mindestens 3).", 3, (c) => distribute(c, true));
def("arrange", "c-arrange", "dsV", I.distV, "Vertikal verteilen", "Verteilt die Objekte mit gleichen Abständen zwischen dem obersten und untersten (mindestens 3).", 3, (c) => distribute(c, false));
def("arrange", "c-arrange", "swap", I.swap, "Positionen tauschen", "Vertauscht die Positionen von genau zwei Objekten.", 2, (c) => swapTwo(c));
def("arrange", "c-arrange", "grid", I.grid, "Als Raster anordnen", "Ordnet alle Objekte in einem gleichmäßigen Raster an, gleiche Zellen, eingestellter Abstand.", 2, (c) => gridAll(c));
def("arrange", "c-arrange", "ctH", I.slideH, "Horizontal auf Folie zentrieren", "Zentriert die Auswahl als Block waagerecht auf der Folie.", 1, (c) => centerOnSlide(c, true, false));
def("arrange", "c-arrange", "ctV", I.slideV, "Vertikal auf Folie zentrieren", "Zentriert die Auswahl als Block senkrecht auf der Folie.", 1, (c) => centerOnSlide(c, false, true));

/* Größe angleichen */
def("size", "c-size", "szW", I.sameW, "Gleiche Breite", "Setzt alle Objekte auf die Breite der Referenz.", "ref", (c) => sizeAll(c, (s, R) => { s.width = R.width; }));
def("size", "c-size", "szH", I.sameH, "Gleiche Höhe", "Setzt alle Objekte auf die Höhe der Referenz.", "ref", (c) => sizeAll(c, (s, R) => { s.height = R.height; }));
def("size", "c-size", "szWH", I.sameWH, "Gleiche Breite und Höhe", "Setzt alle Objekte auf Breite und Höhe der Referenz.", "ref", (c) => sizeAll(c, (s, R) => { s.width = R.width; s.height = R.height; }));
def("size", "c-size", "szSq", I.square, "Quadratisch machen", "Setzt bei jedem Objekt die Höhe gleich der Breite.", 1, (c) => sizeAll(c, (s) => { s.height = s.width; }, true));
def("size", "c-size", "szSwap", I.swapWH, "Breite und Höhe tauschen", "Vertauscht Breite und Höhe, die Mitte bleibt an ihrer Stelle.", 1, (c) => sizeAll(c, (s) => { const cx = s.left + s.width / 2, cy = s.top + s.height / 2, w = s.width; s.width = s.height; s.height = w; s.left = cx - s.width / 2; s.top = cy - s.height / 2; }, true));
def("size", "c-size", "szFit", I.fitText, "Größe an Text anpassen", "Passt die Form automatisch an die Textmenge an.", 1, (c) => fitText(c));

def("sizeSlide", "c-slide", "slW", I.fullW, "Auf Folienbreite", "Streckt das Objekt über die ganze Folienbreite.", 1, (c) => slideSize(c, (s, S) => { s.left = 0; s.width = S.width; }));
def("sizeSlide", "c-slide", "slH", I.fullH, "Auf Folienhöhe", "Streckt das Objekt über die ganze Folienhöhe.", 1, (c) => slideSize(c, (s, S) => { s.top = 0; s.height = S.height; }));
def("sizeSlide", "c-slide", "slAll", I.fullAll, "Auf ganze Folie", "Lässt das Objekt die gesamte Folie ausfüllen.", 1, (c) => slideSize(c, (s, S) => { s.left = 0; s.top = 0; s.width = S.width; s.height = S.height; }));
def("sizeSlide", "c-slide", "asp", I.aspect, "Seitenverhältnis der Referenz", "Behält die Breite und übernimmt das Seitenverhältnis der Referenz.", "ref", (c) => sizeAll(c, (s, R) => { s.height = s.width * R.height / R.width; }));

/* Ebenen */
def("order", "c-order", "zFront", I.toFront, "In den Vordergrund", "Schiebt das Objekt ganz nach vorne, vor alle anderen.", 1, (c) => zorder(c, "BringToFront"));
def("order", "c-order", "zBack", I.toBack, "In den Hintergrund", "Schiebt das Objekt ganz nach hinten, hinter alle anderen.", 1, (c) => zorder(c, "SendToBack"));
def("order", "c-order", "zFwd", I.forward, "Ebene nach vorne", "Schiebt das Objekt eine Ebene nach vorne.", 1, (c) => zorder(c, "BringForward"));
def("order", "c-order", "zBwd", I.backward, "Ebene nach hinten", "Schiebt das Objekt eine Ebene nach hinten.", 1, (c) => zorder(c, "SendBackward"));
def("order", "c-order", "grp", I.group, "Objekte gruppieren", "Fasst die markierten Objekte zu einer Gruppe zusammen.", 2, (c) => groupSel(c));
def("order", "c-order", "ungrp", I.ungroup, "Gruppierung aufheben", "Löst die markierten Gruppen in einzelne Objekte auf.", 1, (c) => ungroupSel(c));

/* Format */
def("format", "", "fFont", I.font, "Schrift angleichen", "Übernimmt Schriftart, -größe, -farbe und Fett/Kursiv von der Referenz.", 2, (c) => copyFont(c), "Schrift angleichen");
def("format", "", "fFmt", I.paint, "Format übertragen", "Übernimmt Füllung und Linie von der Referenz auf die anderen Objekte.", 2, (c) => copyFormat(c), "Format übertragen");

/* Einfügen */
def("insert", "", "inSrc", I.source, "Quellenzeile", "Fügt unten links ein Textfeld „Quelle: …“ in kleiner Schrift ein.", 0, (c) => insertSource(c), "Quelle");
def("insert", "", "inDraft", I.draft, "Entwurf-Stempel", "Fügt oben rechts einen Stempel „ENTWURF“ ein.", 0, (c) => insertStamp(c, "ENTWURF", "#D9552B"), "Entwurf");
def("insert", "", "inConf", I.secret, "Vertraulich-Stempel", "Fügt oben rechts einen Stempel „VERTRAULICH“ ein.", 0, (c) => insertStamp(c, "VERTRAULICH", "#1F3A5F"), "Vertraulich");
def("insert", "", "inLine", I.line, "Trennlinie", "Fügt eine waagerechte Linie in Folienbreite ein.", 0, (c) => insertLine(c), "Linie");

/* ---------- Hilfsfunktionen ---------- */
class UserError extends Error {}

function svg(inner) {
  return '<svg viewBox="0 0 20 20" aria-hidden="true">' + inner + "</svg>";
}

function fmtCm(pt) {
  return (pt / CM).toFixed(2).replace(".", ",") + " cm";
}

function parseLen(text) {
  const m = String(text).trim().replace(",", ".").match(/^(-?\d+(?:\.\d+)?)\s*(cm|mm|pt|in|")?$/i);
  if (!m) return NaN;
  const v = parseFloat(m[1]);
  switch ((m[2] || "cm").toLowerCase()) {
    case "mm": return v / 10 * CM;
    case "pt": return v;
    case "in": case '"': return v * 72;
    default: return v * CM;
  }
}

function gapPt() {
  const v = parseLen(document.getElementById("gap").value);
  return Number.isFinite(v) && v >= 0 ? v : CFG.defaultGapCm * CM;
}

let toastTimer = null;
function toast(text, kind) {
  const el = document.getElementById("msg");
  el.className = kind || "";
  el.textContent = text;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.textContent = ""; el.className = ""; }, 3500);
}

function save(key, val) { try { localStorage.setItem("fb." + key, val); } catch (e) { /* ignorieren */ } }
function load(key, fallback) { try { const v = localStorage.getItem("fb." + key); return v === null ? fallback : v; } catch (e) { return fallback; } }

const BASE_PROPS = "items/id,items/name,items/type,items/left,items/top,items/width,items/height";

async function getSel(ctx) {
  const shapes = ctx.presentation.getSelectedShapes();
  shapes.load(BASE_PROPS);
  await ctx.sync();
  return shapes.items;
}

async function slideBox(ctx) {
  try {
    const ps = ctx.presentation.pageSetup;
    ps.load("slideWidth,slideHeight");
    await ctx.sync();
    return { left: 0, top: 0, width: ps.slideWidth, height: ps.slideHeight };
  } catch (e) {
    toast("Folienmaß nicht lesbar – 16:9 (33,87 × 19,05 cm) angenommen.", "warn");
    return { left: 0, top: 0, width: 960, height: 540 };
  }
}

/* Referenz bestimmen: { box, shape } */
async function getRef(ctx, items) {
  if (state.ref === "slide") return { box: await slideBox(ctx), shape: null };
  if (state.ref === "largest") {
    const s = items.reduce((a, b) => (b.width * b.height > a.width * a.height ? b : a));
    return { box: s, shape: s };
  }
  return { box: items[0], shape: items[0] };
}

function needRefShape(ref) {
  if (!ref.shape) throw new UserError("Diese Funktion braucht ein Objekt als Referenz – bitte „Erstes“ oder „Größtes“ wählen.");
}

/* ---------- Operationen ---------- */
async function alignAll(ctx, fn) {
  const items = await getSel(ctx);
  const ref = await getRef(ctx, items);
  items.forEach((s) => { if (s !== ref.shape) fn(s, ref.box); });
}

async function dockAll(ctx, fn) {
  const items = await getSel(ctx);
  const ref = await getRef(ctx, items);
  needRefShape(ref);
  const g = gapPt();
  items.forEach((s) => { if (s !== ref.shape) fn(s, ref.box, g); });
}

async function stack(ctx, horizontal) {
  const items = (await getSel(ctx)).slice().sort((a, b) => (horizontal ? a.left - b.left : a.top - b.top));
  const g = gapPt();
  let pos = horizontal ? items[0].left + items[0].width + g : items[0].top + items[0].height + g;
  for (let i = 1; i < items.length; i++) {
    const s = items[i];
    if (horizontal) { s.left = pos; pos += s.width + g; }
    else { s.top = pos; pos += s.height + g; }
  }
}

async function distribute(ctx, horizontal) {
  const items = (await getSel(ctx)).slice().sort((a, b) => (horizontal ? a.left - b.left : a.top - b.top));
  const first = items[0], last = items[items.length - 1];
  const start = horizontal ? first.left : first.top;
  const end = horizontal ? last.left + last.width : last.top + last.height;
  const sum = items.reduce((t, s) => t + (horizontal ? s.width : s.height), 0);
  const gap = (end - start - sum) / (items.length - 1);
  let pos = start;
  items.forEach((s) => {
    if (horizontal) { s.left = pos; pos += s.width + gap; }
    else { s.top = pos; pos += s.height + gap; }
  });
}

async function swapTwo(ctx) {
  const items = await getSel(ctx);
  if (items.length !== 2) throw new UserError("Zum Tauschen genau zwei Objekte markieren.");
  const [a, b] = items;
  const ac = { x: a.left + a.width / 2, y: a.top + a.height / 2 };
  const bc = { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  a.left = bc.x - a.width / 2; a.top = bc.y - a.height / 2;
  b.left = ac.x - b.width / 2; b.top = ac.y - b.height / 2;
}

async function gridAll(ctx) {
  const items = (await getSel(ctx)).slice().sort((a, b) => (Math.abs(a.top - b.top) > 10 ? a.top - b.top : a.left - b.left));
  const g = gapPt();
  const cols = Math.ceil(Math.sqrt(items.length));
  const cw = Math.max(...items.map((s) => s.width));
  const ch = Math.max(...items.map((s) => s.height));
  const ox = Math.min(...items.map((s) => s.left));
  const oy = Math.min(...items.map((s) => s.top));
  items.forEach((s, i) => {
    const x = ox + (i % cols) * (cw + g);
    const y = oy + Math.floor(i / cols) * (ch + g);
    s.left = x + (cw - s.width) / 2;
    s.top = y + (ch - s.height) / 2;
  });
}

async function centerOnSlide(ctx, h, v) {
  const items = await getSel(ctx);
  const S = await slideBox(ctx);
  const l = Math.min(...items.map((s) => s.left));
  const t = Math.min(...items.map((s) => s.top));
  const r = Math.max(...items.map((s) => s.left + s.width));
  const b = Math.max(...items.map((s) => s.top + s.height));
  const dx = h ? S.width / 2 - (l + r) / 2 : 0;
  const dy = v ? S.height / 2 - (t + b) / 2 : 0;
  items.forEach((s) => { s.left += dx; s.top += dy; });
}

async function sizeAll(ctx, fn, noRef) {
  const items = await getSel(ctx);
  if (noRef) { items.forEach((s) => fn(s)); return; }
  const ref = await getRef(ctx, items);
  const R = { left: ref.box.left, top: ref.box.top, width: ref.box.width, height: ref.box.height };
  items.forEach((s) => { if (s !== ref.shape) fn(s, R); });
}

async function slideSize(ctx, fn) {
  const items = await getSel(ctx);
  const S = await slideBox(ctx);
  items.forEach((s) => fn(s, S));
}

async function fitText(ctx) {
  const items = (await getSel(ctx)).filter((s) => TEXT_TYPES.includes(s.type));
  if (!items.length) throw new UserError("Keine Form mit Text markiert.");
  items.forEach((s) => { s.textFrame.autoSizeSetting = "AutoSizeShapeToFitText"; });
}

async function zorder(ctx, mode) {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new UserError("Ebenen brauchen PowerPoint mit API 1.8 (aktuelles Microsoft 365).");
  }
  const items = await getSel(ctx);
  const list = mode === "BringForward" || mode === "BringToFront" ? items : items.slice().reverse();
  list.forEach((s) => s.setZOrder(mode));
}

async function groupSel(ctx) {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new UserError("Gruppieren braucht PowerPoint mit API 1.8 (aktuelles Microsoft 365).");
  }
  const items = await getSel(ctx);
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  slide.shapes.addGroup(items.map((s) => s.id));
}

async function ungroupSel(ctx) {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new UserError("Gruppierung aufheben braucht PowerPoint mit API 1.8.");
  }
  const items = (await getSel(ctx)).filter((s) => s.type === "Group");
  if (!items.length) throw new UserError("Keine Gruppe markiert.");
  items.forEach((s) => s.group.ungroup());
}

async function copyFont(ctx) {
  const items = await getSel(ctx);
  const ref = await getRef(ctx, items);
  needRefShape(ref);
  const f = ref.shape.textFrame.textRange.font;
  f.load("name,size,color,bold,italic");
  await ctx.sync();
  items.forEach((s) => {
    if (s === ref.shape || !TEXT_TYPES.includes(s.type)) return;
    const t = s.textFrame.textRange.font;
    t.name = f.name; t.size = f.size; t.color = f.color; t.bold = f.bold; t.italic = f.italic;
  });
}

async function copyFormat(ctx) {
  const items = await getSel(ctx);
  const ref = await getRef(ctx, items);
  needRefShape(ref);
  ref.shape.fill.load("type,foregroundColor");
  ref.shape.lineFormat.load("color,weight,visible");
  await ctx.sync();
  const fill = ref.shape.fill, line = ref.shape.lineFormat;
  items.forEach((s) => {
    if (s === ref.shape) return;
    if (FILL_TYPES.includes(s.type)) {
      if (fill.type === "Solid") s.fill.setSolidColor(fill.foregroundColor); else s.fill.clear();
    }
    if (LINE_TYPES.includes(s.type)) {
      s.lineFormat.visible = line.visible;
      if (line.visible) { s.lineFormat.color = line.color; s.lineFormat.weight = line.weight; }
    }
  });
}

async function applyColor(hex) {
  try {
    await PowerPoint.run(async (ctx) => {
      const items = await getSel(ctx);
      if (!items.length) throw new UserError("Zuerst Objekte markieren.");
      let used = 0;
      items.forEach((s) => {
        if (state.target === "fill" && FILL_TYPES.includes(s.type)) {
          if (hex) s.fill.setSolidColor(hex); else s.fill.clear();
          used++;
        } else if (state.target === "line" && LINE_TYPES.includes(s.type)) {
          if (hex) { s.lineFormat.visible = true; s.lineFormat.color = hex; } else { s.lineFormat.visible = false; }
          used++;
        } else if (state.target === "text" && TEXT_TYPES.includes(s.type) && hex) {
          s.textFrame.textRange.font.color = hex;
          used++;
        }
      });
      if (!used) throw new UserError("Für diese Auswahl gibt es hier nichts zu färben.");
      await ctx.sync();
    });
  } catch (e) {
    toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
  }
}

/* Einfügen */
async function insertSource(ctx) {
  const S = await slideBox(ctx);
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const tb = slide.shapes.addTextBox("Quelle: [Quelle einfügen]", { left: 40, top: S.height - 34, width: S.width - 80, height: 20 });
  tb.name = "Quelle";
  tb.textFrame.textRange.font.size = 10;
  tb.textFrame.textRange.font.color = "#595959";
}

async function insertStamp(ctx, text, color) {
  const S = await slideBox(ctx);
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const w = 110, h = 26;
  const sh = slide.shapes.addGeometricShape("RoundRectangle", { left: S.width - w - 24, top: 18, width: w, height: h });
  sh.name = text;
  sh.fill.setSolidColor(color);
  sh.lineFormat.visible = false;
  sh.textFrame.textRange.text = text;
  sh.textFrame.textRange.font.size = 11;
  sh.textFrame.textRange.font.bold = true;
  sh.textFrame.textRange.font.color = "#FFFFFF";
  sh.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
}

async function insertLine(ctx) {
  const S = await slideBox(ctx);
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const ln = slide.shapes.addLine("Straight", { left: 40, top: S.height / 2, width: S.width - 80, height: 0 });
  ln.name = "Trennlinie";
  ln.lineFormat.color = "#6B7280";
  ln.lineFormat.weight = 1;
}

/* ---------- Ausführung mit Prüfung ---------- */
function makeHandler(b) {
  return async () => {
    try {
      await PowerPoint.run(async (ctx) => {
        if (b.min !== 0) {
          const shapes = ctx.presentation.getSelectedShapes();
          shapes.load("items/id");
          await ctx.sync();
          const n = shapes.items.length;
          const need = b.min === "ref" ? (state.ref === "slide" ? 1 : 2) : b.min;
          if (n === 0) throw new UserError("Zuerst Objekte auf der Folie markieren.");
          if (n < need) throw new UserError(need === 2 && b.min === "ref"
            ? "Mindestens 2 Objekte markieren (oder Referenz „Folie“ wählen)."
            : "Mindestens " + need + " Objekt" + (need > 1 ? "e" : "") + " markieren.");
        }
        await b.fn(ctx);
        await ctx.sync();
      });
    } catch (e) {
      toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
    }
  };
}

/* ---------- Oberfläche aufbauen ---------- */
function buildButtons() {
  BTN.forEach((b) => {
    const host = document.querySelector('[data-group="' + b.group + '"]');
    if (!host) return;
    const el = document.createElement("button");
    el.type = "button";
    el.className = "ib " + b.cls + (b.label ? " tall" : "");
    el.setAttribute("aria-label", b.title);
    el.dataset.tipTitle = b.title;
    el.dataset.tip = b.tip;
    el.innerHTML = svg(b.icon) + (b.label ? '<span style="color:var(--text)">' + b.label + "</span>" : "");
    el.addEventListener("click", makeHandler(b));
    host.appendChild(el);
  });

  const pal = document.querySelector('[data-group="palette"]');
  CFG.palette.forEach(([name, hex]) => {
    const el = document.createElement("button");
    el.type = "button";
    el.className = "sw" + (hex ? "" : " none");
    el.setAttribute("aria-label", name);
    el.dataset.tipTitle = name;
    el.dataset.tip = hex ? "Wendet " + hex + " auf die markierten Objekte an." : "Entfernt Füllung bzw. Linie der markierten Objekte.";
    if (hex) el.style.background = hex;
    el.addEventListener("click", () => applyColor(hex));
    pal.appendChild(el);
  });
}

function setupSegments() {
  const refBtns = document.querySelectorAll("[data-ref]");
  const setRef = (v) => {
    state.ref = v; save("ref", v);
    refBtns.forEach((b) => b.setAttribute("aria-checked", String(b.dataset.ref === v)));
    refreshSelection();
  };
  refBtns.forEach((b) => b.addEventListener("click", () => setRef(b.dataset.ref)));
  setRef(load("ref", "first"));

  const tgtBtns = document.querySelectorAll("[data-target]");
  const setTgt = (v) => {
    state.target = v; save("target", v);
    tgtBtns.forEach((b) => b.setAttribute("aria-checked", String(b.dataset.target === v)));
  };
  tgtBtns.forEach((b) => b.addEventListener("click", () => setTgt(b.dataset.target)));
  setTgt(load("target", "fill"));
}

/* ---------- Größe setzen ---------- */
function setupSizeFields() {
  const w = document.getElementById("w"), h = document.getElementById("h"), lock = document.getElementById("lock");
  lock.addEventListener("click", () => {
    state.lock = !state.lock;
    lock.setAttribute("aria-pressed", String(state.lock));
  });
  w.addEventListener("input", () => {
    const v = parseLen(w.value);
    if (state.lock && Number.isFinite(v) && state.ratio > 0) h.value = fmtCm(v * state.ratio);
  });
  h.addEventListener("input", () => {
    const v = parseLen(h.value);
    if (state.lock && Number.isFinite(v) && state.ratio > 0) w.value = fmtCm(v / state.ratio);
  });
  [w, h].forEach((el) => el.addEventListener("blur", () => {
    const v = parseLen(el.value);
    if (Number.isFinite(v)) el.value = fmtCm(v);
  }));
  document.querySelectorAll("[data-step]").forEach((b) => b.addEventListener("click", () => {
    stepSize(b.dataset.step, Number(b.dataset.dir));
  }));
  document.getElementById("gap").addEventListener("blur", (e) => {
    const v = parseLen(e.target.value);
    e.target.value = Number.isFinite(v) ? fmtCm(v) : fmtCm(CFG.defaultGapCm * CM);
    save("gap", e.target.value);
  });

  document.getElementById("applySize").addEventListener("click", async () => {
    try {
      const wv = w.value.trim() ? parseLen(w.value) : null;
      const hv = h.value.trim() ? parseLen(h.value) : null;
      if ((wv !== null && !(wv > 0)) || (hv !== null && !(hv > 0))) throw new UserError("Bitte Maße wie „12,5 cm“ eingeben.");
      if (wv === null && hv === null) throw new UserError("Breite und/oder Höhe eingeben.");
      await PowerPoint.run(async (ctx) => {
        const items = await getSel(ctx);
        if (!items.length) throw new UserError("Zuerst Objekte auf der Folie markieren.");
        items.forEach((s) => {
          const ratio = s.height / s.width;
          if (wv !== null && hv !== null) { s.width = wv; s.height = hv; }
          else if (wv !== null) { s.width = wv; if (state.lock) s.height = wv * ratio; }
          else { s.height = hv; if (state.lock) s.width = hv / ratio; }
        });
        await ctx.sync();
      });
    } catch (e) {
      toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
    }
  });
}

/* Pfeile: Breite/Höhe der markierten Objekte pro Klick um stepCm ändern */
async function stepSize(dim, dir) {
  const d = dir * CFG.stepCm * CM;
  try {
    await PowerPoint.run(async (ctx) => {
      const items = await getSel(ctx);
      if (!items.length) throw new UserError("Zuerst Objekte auf der Folie markieren.");
      items.forEach((s) => {
        const ratio = s.height / s.width;
        if (dim === "w") {
          const nw = Math.max(1, s.width + d);
          s.width = nw;
          if (state.lock) s.height = nw * ratio;
        } else {
          const nh = Math.max(1, s.height + d);
          s.height = nh;
          if (state.lock) s.width = nh / ratio;
        }
      });
      await ctx.sync();
    });
    state.selKey = "";
    await refreshSelection();
  } catch (e) {
    toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
  }
}

/* ---------- Auswahlanzeige ---------- */
let selTimer = null;
function scheduleRefresh() { clearTimeout(selTimer); selTimer = setTimeout(refreshSelection, 120); }

async function refreshSelection() {
  const count = document.getElementById("selCount");
  const hint = document.getElementById("selHint");
  try {
    await PowerPoint.run(async (ctx) => {
      const items = await getSel(ctx);
      const n = items.length;
      if (!n) {
        count.textContent = "Nichts markiert";
        hint.textContent = "Objekte auf der Folie markieren.";
        return;
      }
      count.textContent = n + (n === 1 ? " Objekt markiert" : " Objekte markiert");
      let refName;
      let box;
      if (state.ref === "slide") { refName = "Folie"; box = null; }
      else {
        const r = await getRef(ctx, items);
        refName = r.shape.name; box = r.box;
      }
      hint.textContent = "Referenz: " + refName;
      const first = box || items[0];
      state.ratio = first.height / first.width;
      const key = items.map((s) => s.id).join(",") + "|" + state.ref;
      if (key !== state.selKey) {
        state.selKey = key;
        const w = document.getElementById("w"), h = document.getElementById("h");
        if (document.activeElement !== w && document.activeElement !== h) {
          w.value = fmtCm(first.width);
          h.value = fmtCm(first.height);
        }
      }
    });
  } catch (e) {
    count.textContent = "Auswahl nicht lesbar";
    hint.textContent = "Bitte Folienansicht öffnen.";
  }
}

/* ---------- Tooltips ---------- */
function setupTooltips() {
  const tip = document.getElementById("tip");
  let timer = null;
  const hide = () => { clearTimeout(timer); tip.hidden = true; };
  const show = (el) => {
    tip.innerHTML = "";
    const t = document.createElement("strong");
    t.textContent = el.dataset.tipTitle || "";
    const p = document.createElement("span");
    p.textContent = el.dataset.tip;
    tip.append(t, p);
    tip.hidden = false;
    const r = el.getBoundingClientRect();
    const tw = tip.offsetWidth, th = tip.offsetHeight;
    let x = r.left + r.width / 2 - tw / 2;
    x = Math.max(6, Math.min(x, window.innerWidth - tw - 6));
    let y = r.bottom + 6;
    if (y + th > window.innerHeight - 6) y = Math.max(6, r.top - th - 6);
    tip.style.left = x + "px";
    tip.style.top = y + "px";
  };
  const target = (e) => (e.target.closest ? e.target.closest("[data-tip]") : null);
  document.addEventListener("mouseover", (e) => {
    const el = target(e);
    if (!el) return;
    hide();
    timer = setTimeout(() => show(el), 250);
  });
  document.addEventListener("mouseout", (e) => { if (target(e)) hide(); });
  document.addEventListener("focusin", (e) => { const el = target(e); if (el) { hide(); show(el); } });
  document.addEventListener("focusout", hide);
  document.addEventListener("click", hide);
  window.addEventListener("scroll", hide, true);
}

/* ---------- Abschnitte merken ---------- */
function setupSections() {
  document.querySelectorAll("details[data-sec]").forEach((d) => {
    const v = load("open." + d.dataset.sec, null);
    if (v !== null) d.open = v === "1";
    d.addEventListener("toggle", () => save("open." + d.dataset.sec, d.open ? "1" : "0"));
  });
}

/* ---------- Start ---------- */
Office.onReady((info) => {
  buildButtons();
  setupSegments();
  setupSizeFields();
  setupTooltips();
  setupSections();
  document.getElementById("gap").value = load("gap", fmtCm(CFG.defaultGapCm * CM));

  if (info.host !== Office.HostType.PowerPoint) {
    toast("Dieses Add-in läuft nur in PowerPoint.", "warn");
    return;
  }
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRefresh);
  refreshSelection();
});

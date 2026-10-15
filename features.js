/* Folienbau – Zusatzfunktionen: Textfeld-Optionen, Form angleichen, Icon-Bibliothek.
 * Nutzt die Hilfsfunktionen aus taskpane.js (getSel, getRef, toast, parseLen, fmtCm, ...). */

/* ================= Textfeld ================= */
const MARGIN_FIELDS = { mL: "leftMargin", mR: "rightMargin", mT: "topMargin", mB: "bottomMargin" };

/* Wendet fn(textFrame) auf alle markierten Formen mit Textfeld an. */
async function textOp(fn) {
  try {
    await PowerPoint.run(async (ctx) => {
      const all = await getSel(ctx);
      if (!all.length) throw new UserError("Zuerst Objekte auf der Folie markieren.");
      const items = all.filter((s) => TEXT_TYPES.includes(s.type));
      if (!items.length) throw new UserError("In der Auswahl gibt es keine Form mit Textfeld.");
      items.forEach((s) => fn(s.textFrame));
      await ctx.sync();
    });
    return true;
  } catch (e) {
    toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
    return false;
  }
}

/* Zeigt die Textfeld-Einstellungen der Referenzform an. */
async function syncTextFields(ctx, shape) {
  if (!shape || !TEXT_TYPES.includes(shape.type)) return;
  const tf = shape.textFrame;
  tf.load("autoSizeSetting,wordWrap,verticalAlignment,leftMargin,rightMargin,topMargin,bottomMargin");
  await ctx.sync();
  const v = String(tf.verticalAlignment || "");
  const va = v.startsWith("Top") ? "Top" : v.startsWith("Bottom") ? "Bottom" : v.startsWith("Middle") ? "Middle" : "";
  document.querySelectorAll("[data-valign]").forEach((b) => b.setAttribute("aria-checked", String(b.dataset.valign === va)));
  document.querySelectorAll('input[name="autosize"]').forEach((r) => { r.checked = r.value === tf.autoSizeSetting; });
  const wrap = document.getElementById("wrap");
  if (document.activeElement !== wrap) wrap.checked = !!tf.wordWrap;
  Object.entries(MARGIN_FIELDS).forEach(([id, prop]) => {
    const el = document.getElementById(id);
    if (document.activeElement !== el && Number.isFinite(tf[prop])) el.value = fmtCm(tf[prop]);
  });
}

/* Pfeile an den Texträndern: Rand aller markierten Formen pro Klick um stepCm ändern (nie unter 0). */
async function stepMargin(id, dir) {
  const prop = MARGIN_FIELDS[id];
  const d = dir * CFG.stepCm * CM;
  try {
    await PowerPoint.run(async (ctx) => {
      const all = await getSel(ctx);
      if (!all.length) throw new UserError("Zuerst Objekte auf der Folie markieren.");
      const items = all.filter((s) => TEXT_TYPES.includes(s.type));
      if (!items.length) throw new UserError("In der Auswahl gibt es keine Form mit Textfeld.");
      items.forEach((s) => s.textFrame.load(prop));
      await ctx.sync();
      items.forEach((s) => { s.textFrame[prop] = Math.max(0, s.textFrame[prop] + d); });
      await ctx.sync();
    });
    refreshSelection();
  } catch (e) {
    toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
  }
}

function initTextOptions() {
  document.querySelectorAll("[data-mstep]").forEach((b) => b.addEventListener("click", () => {
    stepMargin(b.dataset.mstep, Number(b.dataset.dir));
  }));
  document.querySelectorAll("[data-valign]").forEach((b) => b.addEventListener("click", async () => {
    const ok = await textOp((tf) => { tf.verticalAlignment = b.dataset.valign; });
    if (ok) document.querySelectorAll("[data-valign]").forEach((x) => x.setAttribute("aria-checked", String(x === b)));
  }));
  document.querySelectorAll('input[name="autosize"]').forEach((r) => r.addEventListener("change", () => {
    textOp((tf) => { tf.autoSizeSetting = r.value; });
  }));
  document.getElementById("wrap").addEventListener("change", (e) => {
    const on = e.target.checked;
    textOp((tf) => { tf.wordWrap = on; });
  });
  Object.entries(MARGIN_FIELDS).forEach(([id, prop]) => {
    const el = document.getElementById(id);
    el.addEventListener("change", async () => {
      const v = parseLen(el.value);
      if (!(v >= 0)) { toast("Bitte Maße wie „0,25 cm“ eingeben.", "warn"); return; }
      el.value = fmtCm(v);
      await textOp((tf) => { tf[prop] = v; });
    });
  });
}

/* ================= Form angleichen ================= */
const SHAPE_CHOICES = [
  ["Rectangle", "Rechteck"], ["RoundRectangle", "Abgerundetes Rechteck"], ["Ellipse", "Ellipse / Kreis"],
  ["Diamond", "Raute"], ["Triangle", "Dreieck"], ["RightArrow", "Pfeil nach rechts"], ["LeftArrow", "Pfeil nach links"],
  ["UpArrow", "Pfeil nach oben"], ["DownArrow", "Pfeil nach unten"], ["LeftRightArrow", "Pfeil beidseitig"],
  ["Chevron", "Chevron"], ["HomePlate", "Fünfeck-Pfeil"], ["Pentagon", "Fünfeck"], ["Hexagon", "Sechseck"],
  ["Octagon", "Achteck"], ["Trapezoid", "Trapez"], ["Parallelogram", "Parallelogramm"],
  ["Snip1Rectangle", "Rechteck, eine Ecke abgeschnitten"], ["Round2SameRectangle", "Rechteck, obere Ecken rund"],
  ["WedgeRectCallout", "Sprechblase"], ["Can", "Zylinder"], ["Cube", "Quader"], ["Cloud", "Wolke"],
  ["Plus", "Kreuz (Plus)"], ["Star5", "Stern (5 Zacken)"], ["Heart", "Herz"], ["Frame", "Rahmen"], ["Donut", "Ring"],
  ["FlowChartDecision", "Flussdiagramm: Entscheidung"], ["FlowChartTerminator", "Flussdiagramm: Start/Ende"]
];

/* OOXML-Formname (prstGeom) -> PowerPoint.GeometricShapeType; sonst genügt großes Anfangszeichen. */
const PRST_EXCEPT = {
  rect: "Rectangle", roundRect: "RoundRectangle", rtTriangle: "RightTriangle", wedgeRoundRectCallout: "WedgeRRectCallout",
  snip1Rect: "Snip1Rectangle", snip2SameRect: "Snip2SameRectangle", snip2DiagRect: "Snip2DiagonalRectangle",
  snipRoundRect: "SnipRoundRectangle", round1Rect: "Round1Rectangle", round2SameRect: "Round2SameRectangle",
  round2DiagRect: "Round2DiagonalRectangle", diagStripe: "DiagonalStripe"
};
function prstToEnum(prst) {
  return PRST_EXCEPT[prst] || prst.charAt(0).toUpperCase() + prst.slice(1);
}

const CONVERTIBLE = ["GeometricShape", "TextBox"];

let zipPromise = null;
function ensureJSZip() {
  if (window.JSZip) return Promise.resolve();
  if (!zipPromise) {
    zipPromise = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "jszip.min.js";
      s.onload = resolve;
      s.onerror = () => { zipPromise = null; reject(new UserError("jszip.min.js fehlt im Web-Ordner – bitte mit hochladen.")); };
      document.head.appendChild(s);
    });
  }
  return zipPromise;
}

/* Liest die Formarten (prstGeom) der aktuellen Folie aus einer exportierten Kopie der Folie. */
async function readGeometries(ctx) {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new UserError("Formart lesen braucht PowerPoint mit API 1.8 (aktuelles Microsoft 365).");
  }
  await ensureJSZip();
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const b64 = slide.exportAsBase64();
  await ctx.sync();
  const zip = await JSZip.loadAsync(b64.value, { base64: true });
  const entry = Object.keys(zip.files).find((n) => /^ppt\/slides\/slide\d+\.xml$/i.test(n));
  if (!entry) throw new UserError("Die Folie konnte nicht ausgelesen werden.");
  const xml = await zip.file(entry).async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const byId = {}, byName = {};
  Array.from(doc.getElementsByTagName("p:sp")).forEach((sp) => {
    const c = sp.getElementsByTagName("p:cNvPr")[0];
    if (!c) return;
    const g = sp.getElementsByTagName("a:prstGeom")[0];
    const prst = g ? g.getAttribute("prst") : (sp.getElementsByTagName("a:custGeom")[0] ? "custom" : null);
    if (!prst) return;
    byId[c.getAttribute("id")] = prst;
    byName[c.getAttribute("name")] = prst;
  });
  return { byId, byName };
}

/* Ersetzt jede Zielform durch eine neue Form der Art enumName und übernimmt Position, Größe, Füllung, Linie,
 * Text und Textfeld-Einstellungen. Reihenfolge (Ebene) bleibt erhalten. */
async function convertShapes(ctx, targets, enumName, keepIds) {
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const canZ = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
  const canRot = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
  const newIds = [];
  let lost = 0;

  for (const t of targets) {
    const s = slide.shapes.getItem(t.id);
    const all = slide.shapes;
    all.load("items/id");
    s.load("name,left,top,width,height" + (canZ ? ",zOrderPosition" : "") + (canRot ? ",rotation" : ""));
    s.fill.load("type,foregroundColor,transparency");
    s.lineFormat.load("visible,color,weight,dashStyle,transparency");
    const tf = s.textFrame;
    tf.load("autoSizeSetting,wordWrap,verticalAlignment,leftMargin,rightMargin,topMargin,bottomMargin");
    const tr = tf.textRange;
    tr.load("text");
    tr.font.load("name,size,color,bold,italic,underline");
    tr.paragraphFormat.load("horizontalAlignment");
    await ctx.sync();
    const total = all.items.length;

    const n = slide.shapes.addGeometricShape(enumName, { left: s.left, top: s.top, width: s.width, height: s.height });
    n.name = s.name;

    const f = s.fill;
    if (f.type === "Solid") {
      n.fill.setSolidColor(f.foregroundColor);
      if (f.transparency) n.fill.transparency = f.transparency;
    } else if (f.type === "NoFill") {
      n.fill.clear();
    } else {
      if (f.foregroundColor) n.fill.setSolidColor(f.foregroundColor);
      lost++;
    }

    const l = s.lineFormat;
    n.lineFormat.visible = l.visible;
    if (l.visible) {
      n.lineFormat.color = l.color;
      n.lineFormat.weight = l.weight;
      if (l.dashStyle) n.lineFormat.dashStyle = l.dashStyle;
      if (l.transparency) n.lineFormat.transparency = l.transparency;
    }

    const ntf = n.textFrame;
    ntf.verticalAlignment = tf.verticalAlignment;
    ntf.wordWrap = tf.wordWrap;
    ntf.leftMargin = tf.leftMargin; ntf.rightMargin = tf.rightMargin;
    ntf.topMargin = tf.topMargin; ntf.bottomMargin = tf.bottomMargin;
    if (tr.text) {
      const ntr = ntf.textRange;
      ntr.text = tr.text;
      const fo = tr.font, nf = ntr.font;
      if (fo.name) nf.name = fo.name;
      if (fo.size) nf.size = fo.size;
      if (fo.color) nf.color = fo.color;
      if (fo.bold !== null && fo.bold !== undefined) nf.bold = fo.bold;
      if (fo.italic !== null && fo.italic !== undefined) nf.italic = fo.italic;
      if (fo.underline) nf.underline = fo.underline;
      if (tr.paragraphFormat.horizontalAlignment) ntr.paragraphFormat.horizontalAlignment = tr.paragraphFormat.horizontalAlignment;
    }
    if (tf.autoSizeSetting && tf.autoSizeSetting !== "AutoSizeMixed") ntf.autoSizeSetting = tf.autoSizeSetting;
    if (canRot && s.rotation) n.rotation = s.rotation;

    s.delete();
    if (canZ) {
      const steps = (total - 1) - s.zOrderPosition;
      for (let i = 0; i < steps; i++) n.setZOrder("SendBackward");
    }
    n.load("id");
    await ctx.sync();
    newIds.push(n.id);
  }

  const ids = newIds.concat(keepIds);
  slide.setSelectedShapes(ids);
  await ctx.sync();
  state.order = ids.slice();
  state.selKey = "";
  return { count: newIds.length, lost };
}

function convertMessage(res) {
  let m = res.count + (res.count === 1 ? " Form umgewandelt." : " Formen umgewandelt.");
  if (res.lost) m += " Verlauf-/Bildfüllungen wurden durch eine Volltonfarbe ersetzt.";
  toast(m, "ok");
}

/* Alle anderen markierten Formen bekommen die Formart der Referenz. */
async function matchShape() {
  try {
    await PowerPoint.run(async (ctx) => {
      const items = await getSel(ctx);
      if (items.length < 2) throw new UserError("Mindestens 2 Formen markieren: zuerst die zu ändernden, zuletzt die Vorlage.");
      const ref = await getRef(ctx, items);
      needRefShape(ref);
      const geoms = await readGeometries(ctx);
      const geomOf = (s) => geoms.byId[s.id] || geoms.byName[s.name];
      const prst = geomOf(ref.shape);
      if (!prst) throw new UserError("Die Formart der Referenz konnte nicht ermittelt werden.");
      if (prst === "custom") throw new UserError("Die Referenz ist eine Freiform – diese Formart lässt sich nicht übernehmen.");
      let targets = items.filter((s) => s !== ref.shape && CONVERTIBLE.includes(s.type));
      if (!targets.length) throw new UserError("Die übrigen Objekte lassen sich nicht umwandeln (nur Formen und Textfelder).");
      targets = targets.filter((s) => geomOf(s) !== prst);
      if (!targets.length) { toast("Alle markierten Formen haben schon diese Formart.", "ok"); return; }
      convertMessage(await convertShapes(ctx, targets, prstToEnum(prst), [ref.shape.id]));
    });
  } catch (e) {
    toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
  }
}

/* Alle markierten Formen werden in die gewählte Formart umgewandelt. */
async function applyShapeType() {
  const enumName = document.getElementById("shapeType").value;
  save("shapeType", enumName);
  try {
    await PowerPoint.run(async (ctx) => {
      const items = await getSel(ctx);
      if (!items.length) throw new UserError("Zuerst Formen auf der Folie markieren.");
      const targets = items.filter((s) => CONVERTIBLE.includes(s.type));
      if (!targets.length) throw new UserError("In der Auswahl gibt es nichts, was sich umwandeln lässt (nur Formen und Textfelder).");
      convertMessage(await convertShapes(ctx, targets, enumName, []));
    });
  } catch (e) {
    toast(e instanceof UserError ? e.message : "Fehler: " + (e.message || e), "warn");
  }
}

function initShapes() {
  const sel = document.getElementById("shapeType");
  SHAPE_CHOICES.forEach(([val, label]) => {
    const o = document.createElement("option");
    o.value = val; o.textContent = label;
    sel.appendChild(o);
  });
  const saved = load("shapeType", "Rectangle");
  if (SHAPE_CHOICES.some(([v]) => v === saved)) sel.value = saved;
  document.getElementById("matchShape").addEventListener("click", matchShape);
  document.getElementById("applyShapeType").addEventListener("click", applyShapeType);
}

/* ================= Icons ================= */
const POPULAR = ["check", "x", "plus", "minus", "arrow-right", "arrow-left", "arrow-up", "arrow-down", "arrow-up-right", "chevron-right",
  "circle-check", "circle-x", "triangle-alert", "info", "circle-help", "lightbulb", "target", "flag", "star", "heart", "thumbs-up", "thumbs-down",
  "user", "users", "building-2", "briefcase", "calendar", "clock", "mail", "phone", "map-pin", "globe", "search", "settings", "shield-check", "lock",
  "key", "chart-bar", "chart-line", "chart-pie", "trending-up", "trending-down", "euro", "database", "server", "cloud", "cpu", "code", "laptop",
  "smartphone", "file-text", "folder", "link", "rocket", "puzzle", "handshake", "megaphone", "award", "trophy", "layers", "workflow", "refresh-cw",
  "zap", "wrench", "list-checks", "presentation", "brain", "sparkles", "factory", "truck", "leaf", "scale", "gauge"];
const ICON_PAGE = 60;
const iconState = { list: null, loading: null, results: [], shown: 0, query: "", color: "#1F3A5F", stroke: "2" };

function norm(s) {
  return String(s).toLowerCase().replace(/ß/g, "ss").normalize("NFD").replace(/[̀-ͯ]/g, "");
}

function loadIcons() {
  if (iconState.list) return Promise.resolve();
  if (!iconState.loading) {
    iconState.loading = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "icons.js";
      s.onload = () => {
        iconState.list = window.ICONS.map(([name, hay, inner]) => ({
          name, inner, hay: norm(name.replace(/-/g, " ") + " " + hay)
        }));
        resolve();
      };
      s.onerror = () => { iconState.loading = null; reject(new Error("icons.js nicht gefunden – bitte mit hochladen.")); };
      document.head.appendChild(s);
    });
  }
  return iconState.loading;
}

function searchIcons(q) {
  const words = norm(q).split(/\s+/).filter(Boolean);
  if (!words.length) {
    const byName = new Map(iconState.list.map((ic) => [ic.name, ic]));
    return POPULAR.map((n) => byName.get(n)).filter(Boolean);
  }
  const w0 = words[0];
  const score = (ic) => (ic.name.split("-").includes(w0) ? 0 : ic.name.includes(w0) ? 1 : 2);
  return iconState.list
    .filter((ic) => words.every((w) => ic.hay.includes(w)))
    .sort((a, b) => score(a) - score(b) || a.name.length - b.name.length);
}

function renderIcons(reset) {
  const grid = document.getElementById("iconGrid");
  const more = document.getElementById("iconMore");
  const info = document.getElementById("iconInfo");
  if (reset) { grid.textContent = ""; iconState.shown = 0; }
  grid.style.setProperty("--sw", iconState.stroke);
  const slice = iconState.results.slice(iconState.shown, iconState.shown + ICON_PAGE);
  const frag = document.createDocumentFragment();
  slice.forEach((ic) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "ic";
    b.setAttribute("aria-label", ic.name + " einfügen");
    b.dataset.tipTitle = ic.name;
    b.dataset.tip = "Fügt das Icon in der gewählten Farbe und Größe mittig in die Folie ein.";
    b.style.color = iconState.color;
    b.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true">' + ic.inner + "</svg>";
    b.addEventListener("click", () => insertIcon(ic));
    frag.appendChild(b);
  });
  grid.appendChild(frag);
  iconState.shown += slice.length;
  const rest = iconState.results.length - iconState.shown;
  more.hidden = rest <= 0;
  more.textContent = "Mehr anzeigen (" + rest + " weitere)";
  info.textContent = iconState.query
    ? (iconState.results.length ? iconState.results.length + " Treffer" : "Keine Treffer – andere Schreibweise oder englisches Wort probieren.")
    : "Beliebte Icons – oder oben suchen (" + iconState.list.length + " insgesamt).";
}

function applyIconSearch() {
  if (!iconState.list) return;
  iconState.results = searchIcons(iconState.query);
  renderIcons(true);
}

async function ensureIcons() {
  const info = document.getElementById("iconInfo");
  if (iconState.list) { if (!document.getElementById("iconGrid").children.length) applyIconSearch(); return; }
  info.textContent = "Lade Icons …";
  try {
    await loadIcons();
    applyIconSearch();
  } catch (e) {
    info.textContent = e.message;
  }
}

function setIconColor(hex, fromSwatch) {
  iconState.color = hex;
  save("icon.color", hex);
  document.querySelectorAll("#iconSw .sw").forEach((b) => b.setAttribute("aria-checked", String(fromSwatch && b.dataset.hex === hex)));
  document.querySelectorAll("#iconGrid .ic").forEach((b) => { b.style.color = hex; });
  const c = document.getElementById("iconColor");
  if (c.value.toLowerCase() !== hex.toLowerCase()) c.value = hex.toLowerCase();
}

async function insertIcon(ic) {
  try {
    if (!Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) {
      throw new UserError("SVG-Einfügen braucht ein aktuelles Microsoft 365 (ImageCoercion 1.2).");
    }
    let size = parseLen(document.getElementById("iconSize").value);
    if (!Number.isFinite(size) || size <= 0) size = 2 * CM;
    size = Math.min(Math.max(size, 8), 720);
    let S = { width: 960, height: 540 };
    try { await PowerPoint.run(async (ctx) => { S = await slideBox(ctx); }); } catch (e) { /* Standardmaß */ }
    const svgText = '<svg xmlns="http://www.w3.org/2000/svg" width="' + size + '" height="' + size + '" viewBox="0 0 24 24" fill="none" stroke="' +
      iconState.color + '" stroke-width="' + iconState.stroke + '" stroke-linecap="round" stroke-linejoin="round">' + ic.inner + "</svg>";
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(svgText, {
        coercionType: Office.CoercionType.XmlSvg || "xmlSvg",
        imageLeft: (S.width - size) / 2, imageTop: (S.height - size) / 2, imageWidth: size, imageHeight: size
      }, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error && r.error.message))));
    });
    toast("Icon „" + ic.name + "“ eingefügt.", "ok");
  } catch (e) {
    toast(e instanceof UserError ? e.message : "Icon konnte nicht eingefügt werden: " + (e.message || e), "warn");
  }
}

function initIcons() {
  const sw = document.getElementById("iconSw");
  CFG.palette.filter(([, hex]) => hex).forEach(([name, hex]) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sw";
    b.setAttribute("role", "radio");
    b.setAttribute("aria-checked", "false");
    b.setAttribute("aria-label", name);
    b.dataset.hex = hex;
    b.dataset.tipTitle = name;
    b.dataset.tip = "Icon in dieser Farbe einfügen.";
    b.style.background = hex;
    b.addEventListener("click", () => setIconColor(hex, true));
    sw.appendChild(b);
  });
  const savedColor = load("icon.color", iconState.color);
  setIconColor(savedColor, CFG.palette.some(([, h]) => h && h.toLowerCase() === savedColor.toLowerCase()));
  document.querySelectorAll("#iconSw .sw").forEach((b) => b.setAttribute("aria-checked", String(b.dataset.hex.toLowerCase() === savedColor.toLowerCase())));
  document.getElementById("iconColor").addEventListener("input", (e) => setIconColor(e.target.value, false));

  const size = document.getElementById("iconSize");
  size.value = load("icon.size", size.value);
  size.addEventListener("blur", () => {
    const v = parseLen(size.value);
    size.value = Number.isFinite(v) && v > 0 ? fmtCm(v) : "2,00 cm";
    save("icon.size", size.value);
  });

  const stroke = document.getElementById("iconStroke");
  stroke.value = load("icon.stroke", "2");
  iconState.stroke = stroke.value;
  stroke.addEventListener("change", () => {
    iconState.stroke = stroke.value;
    save("icon.stroke", stroke.value);
    document.getElementById("iconGrid").style.setProperty("--sw", stroke.value);
  });

  let t = null;
  document.getElementById("iconSearch").addEventListener("input", (e) => {
    clearTimeout(t);
    t = setTimeout(() => { iconState.query = e.target.value.trim(); applyIconSearch(); }, 150);
  });
  document.getElementById("iconMore").addEventListener("click", () => renderIcons(false));
  document.querySelector('details[data-sec="icons"]').addEventListener("toggle", (e) => { if (e.target.open) ensureIcons(); });
}

/* ================= Start ================= */
function initFeatures() {
  initTextOptions();
  initShapes();
  initIcons();
}

Office.onReady(boot);

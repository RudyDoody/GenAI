// Shared theme, layouts and helpers for the two decks (P5 sales proposal, board pitch).
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const THEME = {
  name: "WeaveWorks Group 4",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "151515", lt1: "FFFFFF", dk2: "1F2A5A", lt2: "F3F4F8",
    accent1: "0F7C78", accent2: "E9A23B", accent3: "2A78D6", accent4: "EB6834",
    accent5: "1BAF7A", accent6: "6250D6", hlink: "0F7C78", folHlink: "6250D6",
  },
};
const HEX = { indigo: "1F2A5A", teal: "0F7C78", marigold: "E9A23B", ink: "151515", ink2: "52514E",
  grid: "E4E3DE", light: "F3F4F8", blue: "2A78D6", orange: "EB6834", aqua: "1BAF7A", violet: "6250D6", red: "E34948", white: "FFFFFF" };

const TEAM = require("../data/facts.json").TEAM;
// 'Member 1' -> full name, '{M1}' -> first name (same rule as facts.sub_names in Python)
function subNames(x) {
  if (typeof x === "string") {
    for (const t of TEAM) {
      const k = t.id[1];
      x = x.split(`Member ${k}`).join(t.name).split(`{${t.id}}`).join(t.name.split(" ")[0]);
    }
    return x;
  }
  if (Array.isArray(x)) return x.map(subNames);
  if (x && typeof x === "object" && "text" in x) return Object.assign({}, x, { text: subNames(x.text) });
  return x;
}

function newDeck(title) {
  const pres = new pptxgen();
  const addSlide = pres.addSlide.bind(pres);
  pres.addSlide = (opts) => {
    const s = addSlide(opts);
    const [t, n, tb] = [s.addText.bind(s), s.addNotes.bind(s), s.addTable.bind(s)];
    s.addText = (txt, o) => t(subNames(txt), o);
    s.addNotes = (txt) => n(subNames(txt));
    s.addTable = (rows, o) => tb(subNames(rows), o);
    return s;
  };
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = title;
  pres.author = "Group 4";
  pres.company = "Group 4 - AI Transformation Consultancy";
  const C = pres.SchemeColor;
  const footer = (dark) => ({ text: { text: "Group 4  |  WeaveWorks Exports  |  synthetic / assumed data unless cited",
    options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 9, color: dark ? C.background2 : C.text2, isTextBox: true } } });
  pres.defineSlideMaster({
    title: "TITLE", background: { color: HEX.indigo },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 9.6, h: 1.9, fontSize: 40, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "bottom", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.25, w: 9.6, h: 1.2, fontSize: 18, color: C.background2, valign: "top", align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "SECTION", background: { color: HEX.teal },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 11.5, h: 1.4, fontSize: 36, bold: true, color: C.background1, fontFace: THEME.headFontFace, align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 11.5, h: 1.0, fontSize: 18, color: C.background1, align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT", background: { color: HEX.white },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 28, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "middle", align: "left" }, text: "" } },
      footer(false),
    ],
    slideNumber: { x: 12.3, y: 7.02, w: 0.6, h: 0.3, fontSize: 9, color: HEX.ink2 },
  });
  pres.defineSlideMaster({
    title: "CLOSING", background: { color: HEX.indigo },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 0.6, w: 11.7, h: 1.0, fontSize: 32, bold: true, color: C.background1, fontFace: THEME.headFontFace, align: "left" }, text: "" } },
      footer(true),
    ],
  });
  return { pres, C };
}

async function icon(IconComp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(IconComp, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// icon inside a filled circle
async function circleIcon(slide, IconComp, x, y, d, fill, name) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: `${name}-circle` });
  const pad = d * 0.24;
  slide.addImage({ data: await icon(IconComp), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: `${name}-icon` });
}

function card(slide, x, y, w, h, fill = HEX.light, name = "card") {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill }, objectName: name });
}

function text(slide, t, o) {
  slide.addText(t, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: HEX.ink, valign: "top" }, o));
}

module.exports = { THEME, HEX, newDeck, icon, circleIcon, card, text, subNames };

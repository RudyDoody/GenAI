// P1 - the three campaign creatives as editable PowerPoint files at exact pixel sizes, for import into Canva
// (each imports as an editable Canva design: real text boxes and shapes, not flat images).
// Same copy, colours and layout as the HTML mock-ups in creatives.py and the Marketing Pack spec.
const path = require("path");
const pptxgen = require("pptxgenjs");

const OUT = path.join(__dirname, "..", "02_Build", "P1_Marketing", "creatives", "canva_import");
const px = (v) => v / 96; // pptx works in inches; Canva reads 96 px per inch
const C = { indigo: "1F2A5A", teal: "0F7C78", marigold: "E9A23B", cotton: "F6F4EF", ink: "151515", grey: "52514E", mute: "8A8984", white: "FFFFFF", sub: "DFE3F1" };
const FONT = "Inter";

function deck(name, w, h) {
  const pres = new pptxgen();
  pres.defineLayout({ name, width: px(w), height: px(h) });
  pres.layout = name;
  pres.author = "Group 4";
  pres.title = name;
  return pres;
}
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: FONT, valign: "top" }, o));
const knit = (s, w, h, color) => { // subtle diagonal knit texture as thin lines
  for (let x = -h; x < w; x += 18) {
    s.addShape("line", { x: px(Math.max(x, 0)), y: px(Math.max(-x, 0)), w: px(Math.min(h, w - x, h + x, w)), h: px(Math.min(h, w - x, h + x, w)),
      line: { color, width: 0.75, transparency: 88 } });
  }
};

async function c1() {
  const W = 1200, H = 627, pres = deck("C1_LinkedIn_1200x627", W, H), s = pres.addSlide();
  s.background = { color: C.indigo };
  T(s, "WEAVEWORKS EXPORTS - TIRUPPUR", { x: px(72), y: px(64), w: px(700), h: px(30), fontSize: 15, bold: true, color: C.marigold, charSpacing: 2 });
  T(s, "Knit in Tiruppur.\nTraceable to the fibre.", { x: px(72), y: px(108), w: px(820), h: px(220), fontSize: 57, bold: true, color: C.white, lineSpacingMultiple: 0.95 });
  T(s, "Organic-cotton knitwear with farm-to-finish data your EU compliance team can actually use.", { x: px(72), y: px(350), w: px(720), h: px(90), fontSize: 21, color: C.sub });
  s.addShape("roundRect", { x: px(72), y: px(500), w: px(420), h: px(64), rectRadius: px(32), fill: { color: C.marigold }, line: { color: C.marigold } });
  T(s, "Get the traceability sample pack", { x: px(72), y: px(500), w: px(420), h: px(64), fontSize: 18, bold: true, color: C.ink, align: "center", valign: "middle" });
  const badges = [["GOTS", 92], ["OEKO-TEX", 130], ["DPP-ready data", 180]];
  let bx = W - 72 - badges.reduce((a, b) => a + b[1] + 10, -10);
  for (const [t, bw] of badges) {
    s.addShape("roundRect", { x: px(bx), y: px(512), w: px(bw), h: px(44), rectRadius: px(22), fill: { color: C.indigo }, line: { color: C.white, width: 1.5, transparency: 40 } });
    T(s, t, { x: px(bx), y: px(512), w: px(bw), h: px(44), fontSize: 14, color: C.white, align: "center", valign: "middle" });
    bx += bw + 10;
  }
  s.addShape("ellipse", { x: px(W - 72 - 170), y: px(64), w: px(170), h: px(170), fill: { color: C.indigo }, line: { color: C.teal, width: 7.5 } });
  s.addShape("ellipse", { x: px(W - 112 - 90), y: px(104), w: px(90), h: px(90), fill: { color: C.marigold }, line: { color: C.marigold } });
  return pres.writeFile({ fileName: path.join(OUT, "C1_LinkedIn_Traceable_by_Design.pptx") });
}

async function c2() {
  const W = 1080, H = 1080, pres = deck("C2_Instagram_1080x1080", W, H), s = pres.addSlide();
  s.background = { color: C.cotton };
  s.addShape("ellipse", { x: px(W - 340), y: px(H - 340), w: px(460), h: px(460), fill: { color: C.marigold }, line: { color: C.marigold } });
  s.addShape("ellipse", { x: px(W - 250), y: px(H - 250), w: px(280), h: px(280), fill: { color: C.indigo }, line: { color: C.indigo } });
  T(s, "FOR FAST-MOVING LABELS", { x: px(80), y: px(80), w: px(700), h: px(34), fontSize: 15, bold: true, color: C.teal, charSpacing: 2 });
  T(s, [{ text: "10", options: { fontSize: 188 } }, { text: " days", options: { fontSize: 68 } }],
    { x: px(80), y: px(120), w: px(900), h: px(260), bold: true, color: C.indigo, valign: "bottom" });
  T(s, "from tech-pack to\napproved sample.", { x: px(80), y: px(395), w: px(900), h: px(150), fontSize: 45, bold: true, color: C.ink, lineSpacingMultiple: 0.95 });
  T(s, "21 days", { x: px(80), y: px(600), w: px(140), h: px(50), fontSize: 23, color: C.mute, strike: "sngStrike", valign: "middle" });
  T(s, "→", { x: px(225), y: px(600), w: px(50), h: px(50), fontSize: 26, color: C.ink, valign: "middle", align: "center" });
  s.addShape("roundRect", { x: px(285), y: px(598), w: px(160), h: px(56), rectRadius: px(12), fill: { color: C.teal }, line: { color: C.teal } });
  T(s, "10 days", { x: px(285), y: px(598), w: px(160), h: px(56), fontSize: 23, bold: true, color: C.white, align: "center", valign: "middle" });
  T(s, "AI-assisted 3D previews first, so the physical sample you receive is the one you approve.", { x: px(80), y: px(700), w: px(700), h: px(110), fontSize: 21, color: "3B3B3B" });
  T(s, "weaveworks.example/studio", { x: px(80), y: px(970), w: px(600), h: px(40), fontSize: 20, bold: true, color: C.indigo });
  return pres.writeFile({ fileName: path.join(OUT, "C2_Instagram_10_Day_Sampling.pptx") });
}

async function c3() {
  const W = 1200, H = 627, pres = deck("C3_LinkedIn_1200x627", W, H), s = pres.addSlide();
  s.background = { color: C.teal };
  T(s, "VALUE PROGRAMMES - KIDS & BASICS", { x: px(72), y: px(60), w: px(700), h: px(30), fontSize: 15, bold: true, color: C.white, charSpacing: 2 });
  T(s, "Price you can plan on.\nQuality you never chase.", { x: px(72), y: px(100), w: px(720), h: px(150), fontSize: 41, bold: true, color: C.white, lineSpacingMultiple: 1.0 });
  T(s, ["Firm costing in 24 hours, assumptions shown", "OEKO-TEX 100 certified, amfori BSCI audited", "On-time shipment tracked every week"].map((t, i, a) => (
    { text: "✓  " + t, options: { breakLine: i < a.length - 1 } })),
    { x: px(72), y: px(300), w: px(700), h: px(150), fontSize: 18, color: C.white, paraSpaceAfter: 8 });
  const cx = W - 72 - 300;
  s.addShape("roundRect", { x: px(cx), y: px(103), w: px(300), h: px(420), rectRadius: px(24), fill: { color: C.white }, line: { color: C.white } });
  T(s, "YOUR QUOTE", { x: px(cx + 30), y: px(139), w: px(240), h: px(26), fontSize: 15, bold: true, color: C.teal, charSpacing: 1.5 });
  T(s, "Kids crew-neck tee\n180 gsm, 100% cotton", { x: px(cx + 30), y: px(183), w: px(240), h: px(60), fontSize: 17, color: C.grey });
  T(s, "24h", { x: px(cx + 30), y: px(263), w: px(240), h: px(80), fontSize: 48, bold: true, color: C.indigo });
  T(s, "from tech-pack to firm price", { x: px(cx + 30), y: px(348), w: px(240), h: px(50), fontSize: 15, color: C.grey });
  s.addShape("roundRect", { x: px(cx + 30), y: px(423), w: px(240), h: px(54), rectRadius: px(27), fill: { color: C.marigold }, line: { color: C.marigold } });
  T(s, "Request a quote", { x: px(cx + 30), y: px(423), w: px(240), h: px(54), fontSize: 17, bold: true, color: C.ink, align: "center", valign: "middle" });
  return pres.writeFile({ fileName: path.join(OUT, "C3_LinkedIn_Value_Programmes.pptx") });
}

(async () => {
  require("fs").mkdirSync(OUT, { recursive: true });
  for (const f of [c1, c2, c3]) console.log(await f());
})();

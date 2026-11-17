// P5 - B2B proposal deck for the largest customer segment (A: Sustainable mid-market brands).
// Prospect "Nordlicht Basics GmbH" and contact are FICTIONAL.
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const { FaLeaf, FaBolt, FaShieldAlt, FaRobot, FaUserCheck, FaCertificate, FaIndustry, FaGlobeEurope, FaClipboardCheck,
  FaHandshake, FaSearch, FaCube, FaShippingFast, FaFileAlt } = require("react-icons/fa");
const { THEME, HEX, newDeck, circleIcon, card, text } = require("./deck_common");
const F = require("../data/facts.json");

const OUT = path.join(__dirname, "..", "02_Build", "P5_Sales", "B2B_Proposal_Deck_Nordlicht.pptx");

(async () => {
  const { pres, C } = newDeck("Traceable by Design - proposal for Nordlicht Basics");

  // 1 Title
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Opening" });
  s.addText("Traceable by Design", { placeholder: "title" });
  s.addText("A knitwear partnership proposal for Nordlicht Basics GmbH, Hamburg\nPrepared for: Katrin A., Head of Sourcing  |  WeaveWorks Exports, Tiruppur  |  January 2027",
    { placeholder: "body" });
  await circleIcon(s, FaLeaf, 10.9, 2.5, 1.6, HEX.teal, "title-leaf");
  s.addNotes("Open with her problem, not our factory. Nordlicht and Katrin are fictional, built from our Segment A persona. Everything in this deck about Nordlicht is illustrative.");

  // 2 What we heard
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  s.addText("What we heard from your sourcing team", { placeholder: "title" });
  const heard = [
    [FaClipboardCheck, "Prove due diligence", "Your retail partners and board want evidence of where each garment comes from, not just certificates."],
    [FaLeaf, "Grow organic to 70%", "Organic and recycled fibres are a range target; supply must keep pace without quality drift."],
    [FaBolt, "Protect the calendar", "Late sample rounds compress your season; three rounds is the norm, two is the goal."],
    [FaFileAlt, "End audit fatigue", "The same questionnaire is answered five ways by five suppliers; data arrives as scattered PDFs."],
  ];
  for (let i = 0; i < heard.length; i++) {
    const x = 0.6 + i * 3.05;
    card(s, x, 1.6, 2.8, 4.6, HEX.light, `heard-card-${i}`);
    await circleIcon(s, heard[i][0], x + 0.3, 1.9, 0.8, i % 2 ? HEX.teal : HEX.indigo, `heard-${i}`);
    text(s, heard[i][1], { x: x + 0.3, y: 2.9, w: 2.25, h: 0.8, fontSize: 18, bold: true, color: HEX.indigo, fontFace: THEME.headFontFace, objectName: `heard-h-${i}` });
    text(s, heard[i][2], { x: x + 0.3, y: 3.75, w: 2.25, h: 2.2, fontSize: 14, color: HEX.ink2, objectName: `heard-b-${i}` });
  }
  text(s, "Source: discovery call notes (illustrative, fictional prospect).", { x: 0.6, y: 6.45, w: 8, h: 0.3, fontSize: 10, color: HEX.ink2, italic: true });
  s.addNotes("Confirm these four priorities with Katrin before going further. If one is wrong, skip the matching pillar later.");

  // 3 Who we are
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  s.addText("WeaveWorks at a glance", { placeholder: "title" });
  const stats = [
    [`${F.CO.pieces_per_year_mn}M`, "pieces shipped in the last 12 months"],
    ["20+", "years exporting knitwear from Tiruppur"],
    [`${Math.round(F.CO.export_share * 100)}%`, "of revenue from exports, mostly EU and UK"],
    [`${F.CO.active_buyers}`, "active brand and retail partners"],
  ];
  stats.forEach((st, i) => {
    const x = 0.6 + i * 3.05;
    text(s, st[0], { x, y: 1.7, w: 2.8, h: 1.1, fontSize: 54, bold: true, color: i % 2 ? HEX.teal : HEX.indigo, fontFace: THEME.headFontFace });
    text(s, st[1], { x, y: 2.85, w: 2.6, h: 0.8, fontSize: 14, color: HEX.ink2 });
  });
  card(s, 0.6, 4.1, 12.1, 2.2, HEX.light, "certs-card");
  await circleIcon(s, FaCertificate, 0.9, 4.45, 0.75, HEX.marigold, "certs");
  text(s, "Certified and audited", { x: 1.9, y: 4.45, w: 4, h: 0.45, fontSize: 18, bold: true, color: HEX.indigo });
  text(s, "GOTS  |  OEKO-TEX Standard 100  |  GRS  |  amfori BSCI audited  |  RSL aligned to EU REACH Annex XVII",
    { x: 1.9, y: 4.95, w: 10.5, h: 0.4, fontSize: 15, color: HEX.ink });
  text(s, "Products: basic and fashion T-shirts, polos, kidswear, loungewear in 140-240 gsm cotton, organic and recycled blends.",
    { x: 1.9, y: 5.45, w: 10.5, h: 0.6, fontSize: 14, color: HEX.ink2 });
  s.addNotes("All WeaveWorks figures are assumed for the capstone (fictional company). Do not add capacity claims we cannot show on a factory visit.");

  // 4 Proposal pillars
  pres.addSection({ title: "Proposal" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Proposal" });
  s.addText("Our proposal: three commitments, each one measurable", { placeholder: "title" });
  const pillars = [
    [FaSearch, "Traceable", HEX.indigo, ["One-page traceability sheet per style: fibre, spinner, knitter, dye house, stitching unit",
      "Certificates and test reports in one shared folder", "Data fields ready for the EU Digital Product Passport"], "100% of styles with a complete sheet"],
    [FaBolt, "Fast", HEX.teal, ["AI-assisted 3D preview within 5 working days", "Physical sample in 10 working days (today: 21)",
      "Firm costing within 24 hours, with yarn assumptions shown"], "Sample lead time <= 10 working days"],
    [FaShieldAlt, "Reliable", HEX.marigold, ["Capacity booked from a rolling demand forecast", "Early-warning alerts when an order is at risk",
      "24x7 support assistant for status and documents; people own every commitment"], "On-time-in-full >= 95%"],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.55, 3.85, 5.0, HEX.light, `pillar-${i}`);
    await circleIcon(s, pillars[i][0], x + 0.3, 1.85, 0.8, pillars[i][2], `pillar-${i}`);
    text(s, pillars[i][1], { x: x + 1.25, y: 1.95, w: 2.4, h: 0.6, fontSize: 24, bold: true, color: HEX.indigo, fontFace: THEME.headFontFace });
    s.addText(pillars[i][3].map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < 2 } })),
      { x: x + 0.3, y: 2.85, w: 3.3, h: 2.6, fontSize: 14, color: HEX.ink, paraSpaceAfter: 6, valign: "top", isTextBox: true, margin: 0 });
    text(s, "KPI: " + pillars[i][4], { x: x + 0.3, y: 5.75, w: 3.3, h: 0.5, fontSize: 13, bold: true, color: HEX.teal });
  }
  s.addNotes("Each pillar has a KPI we will report monthly. The 10-day and 24-hour promises start once our virtual sampling and costing tools are live (roadmap M7 and M4); for a January signing, the pilot starts with the 24-hour costing and 21-day sampling, moving to 10 days from Q3 2027.");

  // 5 How it works - process
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Proposal" });
  s.addText("From your tech-pack to your warehouse", { placeholder: "title" });
  const steps = [
    [FaFileAlt, "Day 0", "Tech-pack received; costing in 24 h"], [FaCube, "Day 5", "3D preview to approve or comment"],
    [FaClipboardCheck, "Day 10", "Physical sample + traceability sheet"], [FaIndustry, "Week 3-12", "Bulk production, weekly status"],
    [FaShippingFast, "Week 12-17", "Shipment; documents in 3 days; ETA alerts"]];
  for (let i = 0; i < steps.length; i++) {
    const x = 0.6 + i * 2.48;
    await circleIcon(s, steps[i][0], x + 0.65, 1.9, 0.95, i % 2 ? HEX.teal : HEX.indigo, `step-${i}`);
    if (i < steps.length - 1) s.addShape("line", { x: x + 1.75, y: 2.37, w: 1.3, h: 0, line: { color: HEX.grid, width: 2, endArrowType: "triangle" } });
    text(s, steps[i][1], { x: x + 0.1, y: 3.05, w: 2.1, h: 0.45, fontSize: 18, bold: true, color: HEX.indigo, align: "center" });
    text(s, steps[i][2], { x: x + 0.1, y: 3.55, w: 2.1, h: 1.0, fontSize: 14, color: HEX.ink2, align: "center" });
  }
  card(s, 0.6, 4.95, 12.1, 1.55, HEX.light, "ai-card");
  await circleIcon(s, FaRobot, 0.9, 5.3, 0.8, HEX.violet, "ai");
  text(s, "Where AI helps, and where people decide", { x: 2.0, y: 5.2, w: 9, h: 0.4, fontSize: 17, bold: true, color: HEX.indigo });
  text(s, "AI drafts 3D previews, costing and status updates. A named merchandiser approves every price, date, claim and compliance statement, and our AI policy forbids your data being used to train public models.",
    { x: 2.0, y: 5.65, w: 10.4, h: 0.8, fontSize: 14, color: HEX.ink2 });
  s.addNotes("Timelines are standard targets for core fabrics. Lead time depends on fabric development and approvals.");

  // 6 Value not just price (shock response)
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Proposal" });
  s.addText("Compare the total cost of a piece, not the quote", { placeholder: "title" });
  const labels = ["Extra sample rounds", "Compliance admin & data gaps", "Late-delivery cost"];
  const comp = [
    { name: "WeaveWorks", labels, values: [0.06, 0.05, 0.04] },
    { name: "Typical low-price offer", labels, values: [0.20, 0.28, 0.18] },
  ];
  s.addChart(pres.charts.BAR, comp, {
    x: 0.6, y: 1.5, w: 7.6, h: 4.9, barDir: "bar", barGrouping: "clustered", chartColors: [HEX.teal, HEX.orange], barGapWidthPct: 60,
    showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.ink, dataLabelFormatCode: "0.00", dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valGridLine: { color: HEX.grid, size: 1 }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 0.35,
    showTitle: true, title: "Hidden costs per piece on top of the FOB price, EUR (illustrative)", titleFontSize: 13, titleColor: HEX.ink, titleFontFace: "+mn-lt",
    valAxisLabelFormatCode: "0.00",
  });
  card(s, 8.6, 1.5, 4.1, 4.9, HEX.light, "value-card");
  text(s, "FOB quote", { x: 8.9, y: 1.75, w: 3.5, h: 0.35, fontSize: 13, color: HEX.ink2 });
  text(s, [{ text: "4.60", options: { color: HEX.teal } }, { text: "  vs  ", options: { fontSize: 16, color: HEX.ink2 } }, { text: "4.14", options: { color: HEX.orange } }],
    { x: 8.9, y: 2.05, w: 3.6, h: 0.6, fontSize: 30, bold: true, fontFace: THEME.headFontFace });
  text(s, "All-in cost per piece", { x: 8.9, y: 2.75, w: 3.5, h: 0.35, fontSize: 13, color: HEX.ink2 });
  text(s, [{ text: "4.75", options: { color: HEX.teal } }, { text: "  vs  ", options: { fontSize: 16, color: HEX.ink2 } }, { text: "4.80", options: { color: HEX.orange } }],
    { x: 8.9, y: 3.05, w: 3.6, h: 0.6, fontSize: 30, bold: true, fontFace: THEME.headFontFace });
  s.addText([
    { text: "Fewer sample rounds: 3D preview first", options: { bullet: true, breakLine: true } },
    { text: "Compliance data supplied, not chased", options: { bullet: true, breakLine: true } },
    { text: "Forecast-led capacity: fewer late shipments", options: { bullet: true } },
  ], { x: 8.9, y: 3.95, w: 3.6, h: 2.2, fontSize: 14, color: HEX.ink, paraSpaceAfter: 6, isTextBox: true, margin: 0, valign: "top" });
  text(s, "Illustrative, assumed figures. We will rebuild this with your own data in the pilot.", { x: 0.6, y: 6.5, w: 11, h: 0.3, fontSize: 10, italic: true, color: HEX.ink2 });
  s.addNotes("Shock E response: a rival now offers AI personalisation and 10% lower prices. We do not match blindly. Show that hidden costs close the gap, then offer the pilot to prove it. All numbers are assumptions, so invite Katrin to plug in her own.");

  // 7 Commercial offer & pilot
  pres.addSection({ title: "Offer" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Offer" });
  s.addText("Start with a 90-day pilot, then scale", { placeholder: "title" });
  const rows = [
    [{ text: "Pilot (90 days)", options: { bold: true, color: "FFFFFF", fill: { color: HEX.indigo } } }, { text: "Details", options: { bold: true, color: "FFFFFF", fill: { color: HEX.indigo } } }],
    ["Scope", "2 styles: organic crew tee 180 gsm + organic pique polo 220 gsm; about 20,000 pieces"],
    ["Indicative FOB", "EUR 4.40-4.80 (tee), EUR 6.10-6.60 (polo) depending on colours and trims (assumed)"],
    ["Included", "Traceability sheet per style; 3D preview; first proto sample free; weekly status; support assistant access"],
    ["Pilot KPIs", "Sample lead time; on-time-in-full; traceability sheet completeness; claims rate"],
    ["Exit clause", "If two of four KPIs miss target, either side can stop after the pilot at no cost"],
    ["Scale-up", "Annual programme about 180,000 pieces; 60-day payment terms; volume rebate above 200,000 pieces (to be agreed)"],
  ];
  s.addTable(rows, { x: 0.6, y: 1.5, w: 12.1, colW: [2.6, 9.5], fontSize: 14, color: HEX.ink, border: { type: "solid", color: HEX.grid, pt: 1 },
    fill: { color: "FFFFFF" }, rowH: 0.62, valign: "middle", fontFace: "Calibri" });
  text(s, "All commercial terms indicative and subject to a signed agreement; figures assumed for the capstone.", { x: 0.6, y: 6.5, w: 11, h: 0.3, fontSize: 10, italic: true, color: HEX.ink2 });
  s.addNotes("Prices are indicative and assumed. The merchandiser and CFO approve any final quote (AI policy: no AI-generated price goes to a buyer without human sign-off).");

  // 8 Why us comparison
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Offer" });
  s.addText("Why WeaveWorks for Nordlicht", { placeholder: "title" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.indigo }, align: "center" } });
  const yes = { text: "Yes", options: { color: HEX.teal, bold: true, align: "center" } };
  const part = { text: "Partly", options: { color: HEX.ink2, align: "center" } };
  const no = { text: "No", options: { color: HEX.red, bold: true, align: "center" } };
  s.addTable([
    [hdr("What you need"), hdr("WeaveWorks"), hdr("Typical low-price offer"), hdr("Large integrated mill")],
    ["Per-style traceability data, DPP-ready", yes, no, part],
    ["Sample in 10 working days", yes, part, part],
    ["Firm costing in 24 hours", yes, part, no],
    ["MOQ under 3,000 pieces per style", yes, no, no],
    ["GOTS + OEKO-TEX + GRS", yes, part, yes],
    ["Named merchandiser + 24x7 status assistant", yes, part, part],
  ], { x: 0.6, y: 1.5, w: 12.1, colW: [5.0, 2.37, 2.37, 2.36], fontSize: 15, color: HEX.ink, border: { type: "solid", color: HEX.grid, pt: 1 },
    rowH: 0.62, valign: "middle", fontFace: "Calibri" });
  text(s, "Comparison is our assessment of typical alternatives in the market, not of any named company.", { x: 0.6, y: 6.2, w: 11, h: 0.3, fontSize: 10, italic: true, color: HEX.ink2 });
  s.addNotes("Never name or disparage a competitor. 'Partly' means it varies by supplier.");

  // 9 Next steps
  s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Offer" });
  s.addText("Next steps", { placeholder: "title" });
  const ns = [["This week", "Share two tech-packs; we return costing in 24 hours"], ["Week 2", "3D previews and traceability sheets for review"],
    ["Week 3", "Pilot agreement signed; samples dispatched"], ["Day 90", "Pilot review against the four KPIs; scale decision"]];
  for (let i = 0; i < ns.length; i++) {
    const x = 0.8 + i * 3.0;
    s.addShape("roundRect", { x, y: 2.0, w: 2.75, h: 3.2, rectRadius: 0.12, fill: { color: "2B3870" }, line: { color: "2B3870" }, objectName: `ns-${i}` });
    text(s, ns[i][0], { x: x + 0.25, y: 2.3, w: 2.3, h: 0.6, fontSize: 22, bold: true, color: HEX.marigold, fontFace: THEME.headFontFace });
    text(s, ns[i][1], { x: x + 0.25, y: 3.0, w: 2.3, h: 2.0, fontSize: 16, color: "FFFFFF" });
  }
  await circleIcon(s, FaHandshake, 0.8, 5.65, 0.8, HEX.teal, "close");
  text(s, "Contact: Member 1 (Sales lead), WeaveWorks Exports  |  sales@weaveworks.example", { x: 1.85, y: 5.85, w: 10, h: 0.4, fontSize: 16, color: "FFFFFF" });
  s.addNotes("Ask for the two tech-packs before leaving the meeting.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(OUT);
})();

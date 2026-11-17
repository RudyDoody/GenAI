// Phase 5 - Board pitch (15 min): recommendation, live demo (Agent + forecast), ROI and risks.
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const fa = require("react-icons/fa");
const { THEME, HEX, newDeck, circleIcon, card, text } = require("./deck_common");
const F = require("../data/facts.json");

const ROOT = path.join(__dirname, "..");
const OUT = path.join(ROOT, "05_Defend", "Board_Pitch_Deck.pptx");
const FIG = path.join(ROOT, "data", "figures");
const B = F.BUDGET;
const inc = F.ROI.filter((r) => r.included);
const ben = [0, 1, 2].map((k) => inc.reduce((a, r) => a + r.benefit[k], 0));
const roi = ben.map((b) => (b - B.total) / B.total);
const pct = (x) => `${Math.round(x * 100)}%`;
const L = (x) => x.toFixed(1);
const ebitda = F.CO.revenue_cr * F.CO.ebitda_margin;

(async () => {
  const { pres } = newDeck("WeaveWorks Exports - 18-month AI plan (Board pitch, Group 4)");
  const notes = (s, t) => s.addNotes(t);

  // 1 Title
  pres.addSection({ title: "Recommendation" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Recommendation" });
  s.addText("Faster, traceable, trusted: WeaveWorks' 18-month AI plan", { placeholder: "title" });
  s.addText(`Board of WeaveWorks Exports  |  Group 4 AI Transformation Consultancy  |  Rs 50 lakh, Jan 2027 - Jun 2028`, { placeholder: "body" });
  await circleIcon(s, fa.FaCut, 10.9, 2.5, 1.6, HEX.teal, "title");
  notes(s, "[Member 1, 0:00-0:30] Good morning. In 15 minutes: our recommendation, two live demos, the numbers, the risks, and what we need from you today.");

  // 2 Recommendation
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Recommendation" });
  s.addText("Our recommendation: approve Rs " + L(B.total) + " lakh for 10 AI use cases in 3 stage-gated horizons", { placeholder: "title", fontSize: 24 });
  const recs = [
    [fa.FaCheckCircle, "Approve", `Rs ${L(B.total)} lakh committed, incl. Rs ${L(B.hidden)} lakh training, oversight & legal and Rs ${L(B.contingency)} lakh contingency. Rs ${L(B.headroom)} lakh held back until Stage-gate 1.`],
    [fa.FaChessKnight, "Respond to KnitSphere with value, not price", `Matching a 10% cut costs about Rs ${Math.round(F.CO.revenue_cr * 0.1)} crore a year, more than our EBITDA (about Rs ${L(ebitda)} crore, assumed). We protect at-risk accounts with value-engineering and launch our own personalisation, with guardrails.`],
    [fa.FaShieldAlt, "Govern from day one", "AI never promises money, dates or compliance. A named person owns every output. DPDP-ready before any customer-facing launch."],
  ];
  for (let i = 0; i < recs.length; i++) {
    const y = 1.55 + i * 1.75;
    card(s, 0.6, y, 8.2, 1.55, HEX.light, `rec-${i}`);
    await circleIcon(s, recs[i][0], 0.85, y + 0.35, 0.85, [HEX.teal, HEX.indigo, HEX.marigold][i], `rec-${i}`);
    text(s, recs[i][1], { x: 1.95, y: y + 0.18, w: 6.6, h: 0.45, fontSize: 18, bold: true, color: HEX.indigo });
    text(s, recs[i][2], { x: 1.95, y: y + 0.62, w: 6.6, h: 0.85, fontSize: 13, color: HEX.ink2 });
  }
  card(s, 9.1, 1.55, 3.6, 5.05, HEX.indigo, "roi-card");
  text(s, "Base-case 18-month ROI", { x: 9.4, y: 1.85, w: 3.1, h: 0.4, fontSize: 14, color: "FFFFFF" });
  text(s, pct(roi[1]), { x: 9.4, y: 2.25, w: 3.1, h: 1.1, fontSize: 60, bold: true, color: HEX.marigold, fontFace: THEME.headFontFace });
  text(s, `Pessimistic ${pct(roi[0])}  |  Optimistic ${pct(roi[2])}`, { x: 9.4, y: 3.45, w: 3.1, h: 0.4, fontSize: 13, color: "FFFFFF" });
  text(s, `Break-even if we deliver ${pct(B.total / ben[1])} of base-case benefits`, { x: 9.4, y: 4.2, w: 3.0, h: 0.9, fontSize: 15, bold: true, color: "FFFFFF" });
  text(s, "All benefit drivers are assumptions; see ROI_Model.xlsx.", { x: 9.4, y: 5.6, w: 3.0, h: 0.7, fontSize: 11, color: "CADCFC", italic: true });
  notes(s, "[Member 1, 0:30-1:45] Lead with the ask. Stress that the pessimistic case loses money inside 18 months; that is why we stage-gate. The sceptical CFO will ask about this, so say it before they do.");

  // 3 Situation
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Recommendation" });
  s.addText("The situation: compliance pressure, low digital skills, and a rival that just moved", { placeholder: "title", fontSize: 24 });
  const sit = [["Rs 120 cr", "revenue; ~1,150 people; 3.8M pieces a year (assumed)"], [`${F.CO.sampling_days_now} days`, "to a sample today; buyers expect about 10"],
    [`${pct(F.CO.repetitive_query_share)}`, `of ${F.CO.buyer_queries_per_month.toLocaleString()} monthly buyer queries are repetitive (assumed)`], ["-10%", "KnitSphere's price cut, plus an AI personalisation portal"]];
  sit.forEach((x, i) => {
    const xx = 0.6 + i * 3.05;
    text(s, x[0], { x: xx, y: 1.7, w: 2.85, h: 1.0, fontSize: 44, bold: true, color: i === 3 ? HEX.orange : HEX.indigo, fontFace: THEME.headFontFace });
    text(s, x[1], { x: xx, y: 2.75, w: 2.7, h: 0.9, fontSize: 14, color: HEX.ink2 });
  });
  card(s, 0.6, 4.0, 5.9, 2.5, HEX.light, "t1");
  await circleIcon(s, fa.FaGlobeEurope, 0.85, 4.25, 0.75, HEX.teal, "t1");
  text(s, "EU buyers demand proof", { x: 1.8, y: 4.3, w: 4.5, h: 0.45, fontSize: 18, bold: true, color: HEX.indigo });
  text(s, "Due-diligence questionnaires, Digital Product Passport data, forced-labour rules from Dec 2027. Traceability is becoming a condition of the order.", { x: 1.8, y: 4.8, w: 4.5, h: 1.5, fontSize: 14, color: HEX.ink2 });
  card(s, 6.8, 4.0, 5.9, 2.5, HEX.light, "t2");
  await circleIcon(s, fa.FaUsers, 7.05, 4.25, 0.75, HEX.marigold, "t2");
  text(s, "A workforce new to digital", { x: 8.0, y: 4.3, w: 4.5, h: 0.45, fontSize: 18, bold: true, color: HEX.indigo });
  text(s, "Adoption, not software, is the constraint. A third of the budget goes to training, oversight and legal, on purpose.", { x: 8.0, y: 4.8, w: 4.5, h: 1.5, fontSize: 14, color: HEX.ink2 });
  notes(s, "[Member 1, 1:45-2:45] Company figures are assumptions for the fictional client. Regulations are real and cited in the governance pack.");

  // 4 Diagnose
  pres.addSection({ title: "What we built" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "What we built" });
  s.addText("We scored 17 use cases and prototyped the five that pay back fastest", { placeholder: "title", fontSize: 24 });
  s.addImage({ path: path.join(FIG, "value_feasibility_2x2.png"), x: 0.6, y: 1.45, w: 10.4, h: 5.49, objectName: "two-by-two", altText: "Value x feasibility 2x2 of 17 AI use cases" });
  text(s, "Blue = prototyped\nViolet = added after the rival's move\nGrey = later or parked", { x: 11.15, y: 1.6, w: 1.75, h: 2.0, fontSize: 12, color: HEX.ink2 });
  notes(s, "[Member 4, 2:45-3:30] Value = revenue, margin, compliance, risk. Feasibility = data, cost, skills, time to value. Scores are team judgements.");

  // 5 Prototypes
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "What we built" });
  s.addText("Five working prototypes, each with a human owner", { placeholder: "title" });
  const pr = [
    [fa.FaBullhorn, "P1 Marketing", "3 buyer segments from 300 synthetic accounts; 2 personas; 'Traceable by Design' campaign, 3 creatives, 3-email flow, ad copy", "Member 1"],
    [fa.FaRobot, "P2 Support Agent", "No-code Agent on 9 knowledge files; 7 escalation triggers; zero money authority; 25-test suite", "Member 2"],
    [fa.FaCoins, "P3 Finance", "AI summary of S.P. Apparels' FY25 annual report: 12 claims checked, 5 corrected; 3-scenario budget", "Member 3"],
    [fa.FaChartLine, "P4 Operations", "Forecast in Excel and Power BI with a DAX variance measure; supply-disruption decision matrix", "Member 4"],
    [fa.FaHandshake, "P5 Sales", "B2B proposal deck for our largest segment, with a total-cost answer to the 10% price cut", "Member 1"],
  ];
  for (let i = 0; i < pr.length; i++) {
    const x = 0.6 + i * 2.46;
    card(s, x, 1.55, 2.3, 4.95, HEX.light, `p-${i}`);
    await circleIcon(s, pr[i][0], x + 0.25, 1.8, 0.75, i % 2 ? HEX.teal : HEX.indigo, `p-${i}`);
    text(s, pr[i][1], { x: x + 0.25, y: 2.7, w: 1.9, h: 0.45, fontSize: 16, bold: true, color: HEX.indigo });
    text(s, pr[i][2], { x: x + 0.25, y: 3.2, w: 1.85, h: 2.6, fontSize: 13, color: HEX.ink2 });
    text(s, "Owner: " + pr[i][3], { x: x + 0.25, y: 5.95, w: 1.9, h: 0.35, fontSize: 12, bold: true, color: HEX.teal });
  }
  notes(s, "[Member 4, 3:30-4:15] One line each. The finance prototype's best result was catching the AI's mistakes: a 40% error rate in the raw draft.");

  // 6 Demo
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "What we built" });
  s.addText("Live demo: the support Agent, then the forecast", { placeholder: "title" });
  s.addText("1  Order status (grounded)   2  At-risk order (escalates)   3  'Refund or I post to 50k followers' (holds the line)   4  Prompt injection (refuses)   5  Power BI variance page", { placeholder: "body" });
  notes(s, "[Member 2, 4:15-7:15 Agent; Member 4, 7:15-8:15 forecast] Run tests T01, T02, T11, T10 from Agent_Test_Suite.xlsx in a fresh chat each. Then the Power BI page: MAPE card, variance by product, capacity gap. BACKUP: screen recordings saved on the desktop and in the team drive. Do not improvise new prompts live.");

  // 7 Shock
  pres.addSection({ title: "Decide" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Decide" });
  s.addText("KnitSphere cut prices 10%. Matching it would cost more than we earn", { placeholder: "title", fontSize: 26 });
  s.addChart(pres.charts.BAR, [{ name: "Rs crore / year", labels: ["Cost of matching 10% for all buyers", "WeaveWorks EBITDA (assumed)"], values: [F.CO.revenue_cr * 0.1, ebitda] }], {
    x: 0.6, y: 1.5, w: 5.6, h: 4.6, barDir: "col", chartColors: [HEX.orange, HEX.indigo], invertedColors: [HEX.indigo], showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 14, dataLabelColor: HEX.ink, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: HEX.grid, size: 1 }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 14,
    showTitle: true, title: "Rs crore per year", titleFontSize: 13, titleColor: HEX.ink, titleFontFace: "+mn-lt",
  });
  const moves = [[fa.FaBalanceScale, "1. Value, not price", "Value-engineered offers for at-risk value retailers, costed in 24 hours. Show total cost per piece, not the quote."],
    [fa.FaMagic, "2. Personalisation with guardrails", "'WeaveWorks Studio': AI design variants, instant indicative quotes, traceability. No buyer designs used to train shared models."],
    [fa.FaCertificate, "3. Compliance as the moat", "Per-style traceability sheets and faster due-diligence answers for sustainable brands (48% of revenue)."]];
  for (let i = 0; i < moves.length; i++) {
    const y = 1.5 + i * 1.6;
    await circleIcon(s, moves[i][0], 6.6, y + 0.1, 0.7, [HEX.teal, HEX.violet, HEX.indigo][i], `mv-${i}`);
    text(s, moves[i][1], { x: 7.5, y: y + 0.05, w: 5.2, h: 0.4, fontSize: 17, bold: true, color: HEX.indigo });
    text(s, moves[i][2], { x: 7.5, y: y + 0.5, w: 5.2, h: 0.95, fontSize: 13, color: HEX.ink2 });
  }
  text(s, "Roadmap changes: UC17 Studio added (+Rs 3.8 lakh); costing and 3D sampling pulled forward; computer-vision QC deferred (-Rs 3.9 lakh).", { x: 0.6, y: 6.35, w: 12, h: 0.4, fontSize: 12, italic: true, color: HEX.ink2 });
  notes(s, "[Member 1, 8:15-9:30] Speed vs risk: Studio goes to 5 friendly buyers in a closed beta at M8, general launch at M9 only after the test suite passes. We will not ship a rushed bot.");

  // 8 Roadmap
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Decide" });
  s.addText("Three horizons, two stage-gates, one budget", { placeholder: "title" });
  const hz = [["H1  Prove & protect", "Jan-Jun 2027", HEX.blue, ["Governance + AI literacy for all", "Support Agent (M3)", "Segmentation & campaigns (M3)", "RFQ generator, forecast, costing (M4)"]],
    ["H2  Scale & differentiate", "Jul-Dec 2027", HEX.aqua, ["Compliance assistant (M7)", "Virtual 3D sampling (M7)", "WeaveWorks Studio portal (M9)", "Tamil voice SOP assistant (M9)"]],
    ["H3  Embed & extend", "Jan-Jun 2028", HEX.violet, ["HR: bias-checked JDs & screening (M13)", "Export-docs checker business case", "CV quality-check data; go/no-go at M18", "Annual AI audit; next roadmap"]]];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    s.addShape("roundRect", { x, y: 1.5, w: 3.85, h: 0.9, rectRadius: 0.1, fill: { color: hz[i][2] }, line: { color: hz[i][2] }, objectName: `hz-h-${i}` });
    text(s, hz[i][0], { x: x + 0.25, y: 1.58, w: 3.4, h: 0.45, fontSize: 18, bold: true, color: "FFFFFF" });
    text(s, hz[i][1], { x: x + 0.25, y: 2.02, w: 3.4, h: 0.32, fontSize: 13, color: "FFFFFF" });
    card(s, x, 2.55, 3.85, 2.9, HEX.light, `hz-b-${i}`);
    s.addText(hz[i][3].map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < 3 } })), { x: x + 0.25, y: 2.75, w: 3.4, h: 2.6, fontSize: 14, color: HEX.ink, paraSpaceAfter: 8, isTextBox: true, margin: 0, valign: "top" });
  }
  const gates = [["Gate 1 (M6)", "Agent >= 35% resolution, 0 breaches; quotes < 24 h; release reserve"], ["Gate 2 (M12)", "Scale or stop each use case on ROI evidence"]];
  gates.forEach((g, i) => {
    const x = 0.6 + i * 6.15;
    s.addShape("diamond", { x, y: 5.75, w: 0.45, h: 0.45, fill: { color: HEX.ink2 }, line: { color: HEX.ink2 }, objectName: `gate-${i}` });
    text(s, g[0], { x: x + 0.6, y: 5.72, w: 1.6, h: 0.5, fontSize: 15, bold: true, color: HEX.indigo, valign: "middle" });
    text(s, g[1], { x: x + 2.2, y: 5.72, w: 3.7, h: 0.6, fontSize: 13, color: HEX.ink2, valign: "middle" });
  });
  notes(s, "[Member 4, 9:30-10:30] Full Gantt in Roadmap_18_Month.docx. The HR stream we did not prototype sits in H3, as the brief requires.");

  // 9 Budget
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Decide" });
  s.addText(`Where the Rs ${F.BUDGET_LAKH} lakh goes, including the hidden costs`, { placeholder: "title" });
  const bl = [["Use cases H1", 0], ["Use cases H2", 0], ["Use cases H3", 0]];
  // spend by horizon of use cases (one-time in build month, run from go-live) - same rule as the roadmap doc
  const HZ = [[1, 6], [7, 12], [13, 18]];
  inc.forEach((r) => {
    const bm = Math.max(1, r.golive - 1);
    HZ.forEach(([a, b], k) => {
      if (bm >= a && bm <= b) bl[k][1] += r.one_time;
      bl[k][1] += r.run_pm * Math.max(0, Math.min(b, 18) - Math.max(a, r.golive) + 1);
    });
  });
  const items = [...bl, ["Training & change", F.HIDDEN_COSTS[0][1]], ["Human oversight", F.HIDDEN_COSTS[1][1]], ["Security & legal", F.HIDDEN_COSTS[2][1]],
    ["Contingency", F.CONTINGENCY_LAKH], ["Reserve (Gate 1)", B.headroom]];
  s.addChart(pres.charts.BAR, [{ name: "Rs lakh", labels: items.map((x) => x[0]), values: items.map((x) => Math.round(x[1] * 100) / 100) }], {
    x: 0.6, y: 1.45, w: 7.8, h: 5.1, barDir: "bar", catAxisOrientation: "maxMin", chartColors: [HEX.indigo, HEX.indigo, HEX.indigo, HEX.teal, HEX.teal, HEX.teal, HEX.marigold, "B5B3AC"],
    showLegend: false, showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 13, dataLabelColor: HEX.ink, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: HEX.grid, size: 1 }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 25, barGapWidthPct: 50,
    showTitle: true, title: "Rs lakh over 18 months", titleFontSize: 13, titleColor: HEX.ink, titleFontFace: "+mn-lt",
  });
  card(s, 8.8, 1.45, 3.9, 5.1, HEX.light, "budget-card");
  text(s, `Rs ${L(B.hidden + B.contingency)} lakh`, { x: 9.1, y: 1.75, w: 3.4, h: 0.8, fontSize: 34, bold: true, color: HEX.teal, fontFace: THEME.headFontFace });
  text(s, `(${pct((B.hidden + B.contingency) / B.total)} of committed spend) goes to training, oversight, legal and contingency`, { x: 9.1, y: 2.55, w: 3.4, h: 1.0, fontSize: 14, color: HEX.ink2 });
  text(s, "If costs overrun (pessimistic: +25% implementation, +20% licences) we defer the Tamil voice assistant first. We never cut training, oversight or security.", { x: 9.1, y: 3.85, w: 3.4, h: 2.4, fontSize: 14, color: HEX.ink });
  notes(s, "[Member 3, 10:30-11:30] Indigo = use cases by horizon; teal = hidden costs; marigold = contingency; grey = unallocated reserve. Detail: Finance_Budget_and_Peer.xlsx (3 scenarios).");

  // 10 ROI
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Decide" });
  s.addText("Every use case has a break-even point we can track", { placeholder: "title" });
  const sorted = [...inc].sort((a, b) => b.roi[1] - a.roi[1]);
  s.addChart(pres.charts.BAR, [{ name: "Base ROI", labels: sorted.map((r) => `${r.id} ${r.name.split(" (")[0]}`), values: sorted.map((r) => Math.round(r.roi[1] * 100)) }], {
    x: 0.6, y: 1.45, w: 7.8, h: 5.2, barDir: "bar", catAxisOrientation: "maxMin", chartColors: [HEX.teal], showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', dataLabelFontSize: 12, dataLabelColor: HEX.ink, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: HEX.grid, size: 1 }, catGridLine: { style: "none" }, valAxisLabelFormatCode: '0"%"', valAxisMinVal: -100, valAxisMaxVal: 500, catAxisLabelPos: "low",
    showTitle: true, title: "Base-case 18-month ROI on direct cost, %", titleFontSize: 13, titleColor: HEX.ink, titleFontFace: "+mn-lt",
  });
  const ag = F.ROI.find((r) => r.id === "UC04");
  const st = F.ROI.find((r) => r.id === "UC17");
  const ex = [["Programme ROI", `${pct(roi[0])} / ${pct(roi[1])} / ${pct(roi[2])}`, "pessimistic / base / optimistic"],
    ["Support Agent breaks even at", `${Math.round(ag.breakeven_driver).toLocaleString()} hrs/yr`, `of merchandiser time saved (base ${ag.vols[1].toLocaleString()})`],
    ["Studio breaks even at", `Rs ${L(st.breakeven_driver * 1)} cr/yr`, `of at-risk revenue retained (base Rs ${st.vols[1]} cr)`]];
  ex.forEach((e, i) => {
    const y = 1.45 + i * 1.75;
    card(s, 8.8, y, 3.9, 1.55, HEX.light, `ex-${i}`);
    text(s, e[0], { x: 9.05, y: y + 0.15, w: 3.5, h: 0.35, fontSize: 13, color: HEX.ink2 });
    text(s, e[1], { x: 9.05, y: y + 0.48, w: 3.5, h: 0.55, fontSize: 24, bold: true, color: HEX.indigo, fontFace: THEME.headFontFace });
    text(s, e[2], { x: 9.05, y: y + 1.05, w: 3.5, h: 0.4, fontSize: 12, color: HEX.ink2 });
  });
  notes(s, "[Member 3, 11:30-12:30] Weakest: UC13 HR screening is negative in 18 months; we keep it because of compliance and quality of hire, and it is tiny (Rs 0.7 lakh). ROI excludes shared hidden costs; the programme ROI includes them.");

  // 11 Risks
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Decide" });
  s.addText("The risks a regulator would ask about, and our controls", { placeholder: "title" });
  const rk = [[fa.FaCommentSlash, "Chatbot makes an out-of-policy promise", "Zero money/date authority; 7 critical tests; humans own commitments"],
    [fa.FaUserSecret, "Staff paste buyer data into free AI tools", "Approved tools with no-training terms; data classes; 1-hour reporting; training"],
    [fa.FaBalanceScale, "Hiring AI discriminates", "No auto-reject; anonymised inputs; planted-pair bias test; Tier-3 approval"],
    [fa.FaFileSignature, "Wrong compliance statement to an EU buyer", "Source-linked answers; Compliance Manager signs every outgoing answer"]];
  for (let i = 0; i < rk.length; i++) {
    const x = 0.6 + (i % 2) * 6.15, y = 1.5 + Math.floor(i / 2) * 1.85;
    card(s, x, y, 5.95, 1.65, HEX.light, `rk-${i}`);
    await circleIcon(s, rk[i][0], x + 0.25, y + 0.4, 0.8, HEX.red, `rk-${i}`);
    text(s, rk[i][1], { x: x + 1.3, y: y + 0.22, w: 4.4, h: 0.45, fontSize: 16, bold: true, color: HEX.indigo });
    text(s, rk[i][2], { x: x + 1.3, y: y + 0.7, w: 4.4, h: 0.8, fontSize: 13, color: HEX.ink2 });
  }
  card(s, 0.6, 5.3, 12.1, 1.3, HEX.indigo, "law-card");
  text(s, "Compliance map:  DPDP Act 2023 + Rules 2025 (notice, security, breach intimation, erasure)  |  CERT-In 6-hour reporting  |  EU GDPR for buyer contacts  |  EU AI Act Art. 50 chatbot disclosure  |  CSDDD, DPP, Forced Labour Regulation via buyers", { x: 0.9, y: 5.45, w: 11.5, h: 1.0, fontSize: 14, color: "FFFFFF", valign: "middle" });
  notes(s, "[Member 2, 12:30-13:30] 17 risks in Risk_Register.xlsx; approval workflow in 3 tiers; AI Board monthly. If asked: the DPDP Act applies to B2B, since buyer staff and employees are data principals.");

  // 12 People
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Decide" });
  s.addText("Our 28 merchandisers will feel this most, so we start with them", { placeholder: "title" });
  const ad = [["A", "Awareness", "Town-hall in Tamil + English: why, and a no-AI-job-loss pledge for 18 months"], ["D", "Desire", "Each person picks the first task AI takes off them; peer champions"],
    ["K", "Knowledge", "2-hour AI literacy for all; role labs; WhatsApp help group"], ["A", "Ability", "Go-live outside peak; buddy system; human sign-off in every flow"], ["R", "Reinforcement", "Buyer-time KPIs, not AI-usage scores; fix complaints in 2 weeks"]];
  ad.forEach((a, i) => {
    const x = 0.6 + i * 2.46;
    s.addShape("ellipse", { x: x + 0.6, y: 1.6, w: 1.0, h: 1.0, fill: { color: i % 2 ? HEX.teal : HEX.indigo }, line: { color: "FFFFFF" }, objectName: `ad-${i}` });
    text(s, a[0], { x: x + 0.6, y: 1.6, w: 1.0, h: 1.0, fontSize: 32, bold: true, color: "FFFFFF", align: "center", valign: "middle", fontFace: THEME.headFontFace });
    text(s, a[1], { x: x + 0.1, y: 2.8, w: 2.2, h: 0.45, fontSize: 17, bold: true, color: HEX.indigo, align: "center" });
    text(s, a[2], { x: x + 0.1, y: 3.3, w: 2.2, h: 1.6, fontSize: 13, color: HEX.ink2, align: "center" });
  });
  card(s, 0.6, 5.2, 12.1, 1.35, HEX.light, "people-card");
  text(s, "What we will not do: score staff on AI usage, monitor workers with AI without an approved purpose, or launch the floor voice assistant without consulting worker representatives.", { x: 0.9, y: 5.35, w: 11.5, h: 1.1, fontSize: 15, color: HEX.ink, valign: "middle" });
  notes(s, "[Member 2, 13:30-14:15] Full plan in Change_Plan_ADKAR.docx.");

  // 13 Ask
  s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Decide" });
  s.addText("What we ask the board today", { placeholder: "title" });
  const ask = [["1", `Approve Rs ${L(B.total)} lakh`, "with Rs " + L(B.headroom) + " lakh held for Stage-gate 1"], ["2", "Approve the AI policy", "and the AI Board (CEO, CFO, CARO, function head)"],
    ["3", "Endorse 'value, not price'", "no across-the-board 10% cut"], ["4", "Commit to the people pledge", "no AI-driven job losses for 18 months"]];
  ask.forEach((a, i) => {
    const x = 0.8 + i * 3.0;
    s.addShape("roundRect", { x, y: 1.9, w: 2.75, h: 3.0, rectRadius: 0.12, fill: { color: "2B3870" }, line: { color: "2B3870" }, objectName: `ask-${i}` });
    text(s, a[0], { x: x + 0.25, y: 2.1, w: 1, h: 0.8, fontSize: 40, bold: true, color: HEX.marigold, fontFace: THEME.headFontFace });
    text(s, a[1], { x: x + 0.25, y: 2.95, w: 2.3, h: 0.9, fontSize: 18, bold: true, color: "FFFFFF" });
    text(s, a[2], { x: x + 0.25, y: 3.85, w: 2.3, h: 0.9, fontSize: 14, color: "CADCFC" });
  });
  text(s, "First 90 days: AI policy and literacy live (M1-2); Agent and campaigns live (M3); costing assistant answering the rival (M4).", { x: 0.8, y: 5.4, w: 11.7, h: 0.6, fontSize: 16, color: "FFFFFF" });
  notes(s, "[Shared close, 14:15-15:00] Each member says one sentence on what they own. Then: 'We are happy to take your questions.'");

  // Appendix
  pres.addSection({ title: "Appendix" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
  s.addText("Appendix: team, roles and where to find the evidence", { placeholder: "title" });
  const rows = [[{ text: "Member", options: { bold: true, color: "FFFFFF", fill: { color: HEX.indigo } } }, { text: "Role", options: { bold: true, color: "FFFFFF", fill: { color: HEX.indigo } } }, { text: "Owns (evidence file)", options: { bold: true, color: "FFFFFF", fill: { color: HEX.indigo } } }]];
  F.TEAM.forEach((t) => rows.push([t.name, t.role, t.owns]));
  rows.push(["All", "CEO (shared)", "Roadmap_18_Month.docx, this deck, Shock_E_Response_Competitor_Move.docx"]);
  s.addTable(rows, { x: 0.6, y: 1.5, w: 12.1, colW: [1.6, 4.4, 6.1], fontSize: 13, color: HEX.ink, border: { type: "solid", color: HEX.grid, pt: 1 }, rowH: 0.7, valign: "middle", fontFace: "Calibri" });
  notes(s, "Use only if asked who did what.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(OUT);
})();

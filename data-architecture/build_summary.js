// Builds "Conversion Baseline & Data Architecture" summary deck. It links:
//  - the course conversion baseline work agreed with Sarah, Lyndsay and Lesley
//  - the student system data architecture map (Lucid working file, 1 Oct 2026)
//  - early findings from exploring the RA5 Power BI report
// Run: NODE_PATH=<dir with pptxgenjs> PPTX_SKILL_DIR=<pptx skill> node build_summary.js
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL_DIR + "/scripts/apply_theme.js");

const THEME = {
  name: "FVC Conversion Baseline",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2233", lt1: "FFFFFF", dk2: "2B2770", lt2: "F3F3FB",
    accent1: "5B57E8", accent2: "1F9D6A", accent3: "E8912D", accent4: "3A414A",
    accent5: "DEDEFF", accent6: "C3F0D6", hlink: "5B57E8", folHlink: "2B2770",
  },
};
const AMBER_TINT = "FDF0DF";
const MUTED = "6B6F80";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Course Conversion Baseline and Data Architecture";
pres.author = "Thanaporn Wattanamanon";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.7, w: 11.9, h: 2.0, fontSize: 42, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.9, w: 11.9, h: 1.5, fontSize: 20, color: C.accent5, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.1, w: 12.1, h: 0.45, fontSize: 15, color: C.text1, valign: "top", margin: 0 }, text: "" } },
    { text: { text: "Course Conversion Baseline  |  Forth Valley College", options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
});

// ---------- helpers ----------
function card(slide, name, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name, x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } });
}
function box(slide, name, x, y, w, h, title, sub, o = {}) {
  const fill = o.fill || C.accent5;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    objectName: name, x, y, w, h, rectRadius: 0.08, fill: { color: fill },
    line: { color: o.line || fill, width: o.lineW || 1, dashType: o.dash ? "dash" : "solid" },
  });
  const tc = o.color || C.text1;
  const runs = [{ text: title, options: { bold: true, fontSize: o.fs || 13, color: tc, breakLine: !!sub } }];
  if (sub) runs.push({ text: sub, options: { fontSize: o.sfs || 10.5, color: tc } });
  slide.addText(runs, { isTextBox: true, objectName: name + " text", x: x + 0.05, y, w: w - 0.1, h, align: o.align || "center", valign: "middle", margin: 2 });
}
function arrow(slide, x1, y1, x2, y2, o = {}) {
  slide.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: o.color || C.text2, width: o.width || 1.75, dashType: o.dash ? "dash" : "solid", endArrowType: "triangle" },
  });
}
function numDot(slide, n, x, y, d, fill) {
  slide.addShape(pres.shapes.OVAL, { objectName: "dot " + n, x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  slide.addText(String(n), { isTextBox: true, x, y, w: d, h: d, fontSize: d > 0.6 ? 18 : 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
}

// ================= 1. Title =================
pres.addSection({ title: "Why" });
let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Why" });
s.addText("Course Conversion Baseline and Data Architecture", { placeholder: "title" });
s.addText("Summary of the meeting with Sarah, Lyndsay and Lesley, linked to the student data map and first findings from the RA5 report", { placeholder: "body" });
s.addText("Thanaporn (Sanook) Wattanamanon  |  October 2026", { isTextBox: true, x: 0.7, y: 6.4, w: 8, h: 0.4, fontSize: 14, color: C.background1, margin: 0 });
s.addNotes("This deck brings together two pieces of work: the course conversion baseline Sarah asked for, and the data architecture map Darren assigned. The map explains where the conversion numbers come from.");

// ================= 2. Why this matters =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Why" });
s.addText("Why this matters", { placeholder: "title" });
s.addText("\"It's ultimately about financial sustainability and reducing the workload and using the data we have.\"  (Sarah)", { placeholder: "body" });
const stats = [
  ["79,000", "credits the college must deliver each year", C.accent1, C.background2],
  ["~45,000", "of those credits come from full-time courses", C.accent1, C.background2],
  ["10%", "of credits lost through under-enrolment", C.accent3, AMBER_TINT],
];
stats.forEach((st, i) => {
  const x = 0.6 + i * 4.1;
  card(s, "stat " + (i + 1), x, 1.85, 3.85, 2.2, st[3]);
  s.addText(st[0], { isTextBox: true, x, y: 1.95, w: 3.85, h: 1.2, fontSize: 54, bold: true, color: st[2], align: "center", valign: "middle", margin: 0 });
  s.addText(st[1], { isTextBox: true, x: x + 0.3, y: 3.15, w: 3.25, h: 0.75, fontSize: 14, color: C.text1, align: "center", valign: "top", margin: 0 });
});
const goals = [
  ["Financial sustainability", "Lost credits must be made up through January starts and credit shortfall plans. The aim is 20 students starting a course, not 14."],
  ["Less manual workload", "Recruitment is \"very onerous… all very manual\". Student admin keep opening and closing applications and sending waiting-list offers."],
];
goals.forEach((g, i) => {
  const x = 0.6 + i * 6.2;
  card(s, "goal " + (i + 1), x, 4.4, 5.9, 2.3, C.accent5);
  numDot(s, i + 1, x + 0.3, 4.65, 0.5, C.text2);
  s.addText(g[0], { isTextBox: true, x: x + 1.0, y: 4.62, w: 4.7, h: 0.55, fontSize: 18, bold: true, color: C.text2, valign: "middle", margin: 0 });
  s.addText(g[1], { isTextBox: true, x: x + 1.0, y: 5.25, w: 4.7, h: 1.3, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
});
s.addNotes("CTF is the College Transformation Framework, Sarah's three-year plan. This work supports it.");

// ================= 3. The problem =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Why" });
s.addText("One rule for every course does not fit", { placeholder: "title" });
s.addText("Every course aims for 200% applications and 120% offers, but courses convert very differently", { placeholder: "body" });
// Rule vs reality
card(s, "rule", 0.6, 1.85, 5.9, 2.6, C.background2);
s.addText("Today's blanket rule (20 places)", { isTextBox: true, x: 0.9, y: 2.0, w: 5.3, h: 0.45, fontSize: 17, bold: true, color: C.text2, margin: 0 });
[["40", "applications"], ["24", "offers"], ["20", "enrolments expected"]].forEach((r, i) => {
  const x = 0.9 + i * 1.8;
  s.addText(r[0], { isTextBox: true, x, y: 2.6, w: 1.6, h: 0.9, fontSize: 40, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0 });
  s.addText(r[1], { isTextBox: true, x, y: 3.5, w: 1.6, h: 0.7, fontSize: 13, color: C.text1, align: "center", valign: "top", margin: 0 });
});
card(s, "reality", 6.8, 1.85, 5.9, 2.6, AMBER_TINT);
s.addText("What actually happens", { isTextBox: true, x: 7.1, y: 2.0, w: 5.3, h: 0.45, fontSize: 17, bold: true, color: C.text2, margin: 0 });
s.addText([
  { text: "\"We put out 24 offers and sometimes 12 people can turn up.\"", options: { bullet: true, breakLine: true } },
  { text: "Construction: 14 places, 16–17 offers, 9 people turned up.", options: { bullet: true, breakLine: true } },
  { text: "Lecturers add waiting-list offers to people unlikely to come.", options: { bullet: true } },
], { isTextBox: true, x: 7.1, y: 2.55, w: 5.35, h: 1.75, fontSize: 14, color: C.text1, paraSpaceAfter: 6, valign: "top", margin: 0 });
// Now vs later
card(s, "now", 0.6, 4.75, 8.3, 1.95, C.accent6);
s.addText([
  { text: "What Sarah wants now", options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
  { text: "An Excel baseline: every course with its own conversion rates, target applications and target offers, so we have the best chance of full enrolment on day one.", options: { fontSize: 14, color: C.text1 } },
], { isTextBox: true, x: 0.9, y: 4.85, w: 7.8, h: 1.75, valign: "middle", margin: 0 });
card(s, "later", 9.2, 4.75, 3.5, 1.95, C.background2);
s.addText([
  { text: "Not yet", options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
  { text: "Automating targets in Enquirer or UnitE: \"way, way, way down the line\".", options: { fontSize: 14, color: C.text1 } },
], { isTextBox: true, x: 9.45, y: 4.85, w: 3.05, h: 1.75, valign: "middle", margin: 0 });

// ================= 4. Where the numbers come from (architecture link) =================
pres.addSection({ title: "Data" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Data" });
s.addText("Where the conversion numbers come from", { placeholder: "title" });
s.addText("The recruitment path through our systems, taken from the data architecture map", { placeholder: "body" });
const path5 = [
  ["Curriculum planning", "Target enrolments per course", C.background2, C.text1],
  ["Enquirer", "Applications, statuses, offers, curriculum planning", C.accent5, C.text1],
  ["UnitE", "Student, application, enrolment, course", C.accent1, C.background1],
  ["SQL Server connector", "Feeds from UnitE, Enquirer, TripleS, Bursary, iTrent", C.accent4, C.background1],
  ["Power BI: RA5", "Total and live applications, offers, conversion", C.accent5, C.text1],
  ["Excel baseline", "Course-specific rates and targets (new)", C.accent6, C.text1],
];
const pw = 1.85, pg = 0.2;
path5.forEach((p, i) => {
  const x = 0.6 + i * (pw + pg);
  box(s, "path " + p[0], x, 2.0, pw, 1.75, p[0], p[1], { fill: p[2], color: p[3], fs: 14, sfs: 11, dash: i === 5, line: i === 5 ? C.accent2 : undefined, lineW: 1.5 });
  if (i < 5) arrow(s, x + pw + 0.02, 2.875, x + pw + pg - 0.02, 2.875);
});
s.addText("Later: course targets go back into Enquirer", { isTextBox: true, x: 6.8, y: 3.85, w: 5.9, h: 0.35, fontSize: 12, italic: true, color: MUTED, align: "right", margin: 0 });
const watch = [
  ["Manual step inside the path", "Applications move from Enquirer into UnitE by hand, and so do Student Portal and FVC Portal updates. Re-keying can change counts."],
  ["Hidden middle layer", "RA5 reads from a Power BI semantic model. We can see the model but not yet which tables feed it."],
  ["Status history is not kept", "Enquirer keeps only each application's current status, so past offer counts cannot be rebuilt."],
];
watch.forEach((w, i) => {
  const x = 0.6 + i * 4.1;
  card(s, "watch " + (i + 1), x, 4.45, 3.85, 2.25, AMBER_TINT);
  numDot(s, i + 1, x + 0.25, 4.65, 0.45, C.accent3);
  s.addText(w[0], { isTextBox: true, x: x + 0.85, y: 4.62, w: 2.85, h: 0.5, fontSize: 15, bold: true, color: C.text2, valign: "middle", margin: 0 });
  s.addText(w[1], { isTextBox: true, x: x + 0.25, y: 5.25, w: 3.4, h: 1.35, fontSize: 13, color: C.text1, valign: "top", margin: 0 });
});
s.addNotes("This is the link between the two pieces of work. The data map shows that conversion numbers travel from curriculum planning and Enquirer through UnitE and the SQL Server connector into Power BI. The amber boxes are the points where the numbers can change or be lost along the way. Whether the RA5 semantic model reads from the SQL Server connector still needs confirming with Kevin.");

// ================= 5. Funnel =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Data" });
s.addText("Where students drop out", { placeholder: "title" });
s.addText("Once the rates are clean, we can see the stage where each course loses people", { placeholder: "body" });
const funnel = [
  ["Application", "Applied for the course", "People don't turn up to the skills test (construction)"],
  ["Interview / skills test", "Seen by the department", "Lecturers hold back, or make waiting-list offers"],
  ["Offer", "Offer made", "Offers turned down; waiting-list offers go to people unlikely to come"],
  ["Enrolment (day one)", "Starts the course", "Early dropout before credits are claimed"],
  ["Enrolled at ~6 weeks", "Credits claimed", "The point that counts for funding"],
];
funnel.forEach((f, i) => {
  const y = 1.8 + i * 1.0;
  const inset = i * 0.35;
  const fill = i === 4 ? C.accent2 : C.accent1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "funnel " + (i + 1), x: 0.6 + inset, y, w: 5.6 - inset * 2, h: 0.82, rectRadius: 0.08, fill: { color: fill, transparency: i * 8 }, line: { color: fill, transparency: i * 8 } });
  s.addText([{ text: f[0], options: { bold: true, fontSize: 14, breakLine: true } }, { text: f[1], options: { fontSize: 11 } }], { isTextBox: true, x: 0.6 + inset, y, w: 5.6 - inset * 2, h: 0.82, color: C.background1, align: "center", valign: "middle", margin: 0 });
  card(s, "cause " + (i + 1), 6.6, y, 6.1, 0.82, i === 4 ? C.accent6 : C.background2);
  s.addText(f[2], { isTextBox: true, x: 6.85, y, w: 5.7, h: 0.82, fontSize: 13, color: C.text1, valign: "middle", margin: 0 });
});
s.addText("Possible cause of drop-off (from the meeting)", { isTextBox: true, x: 6.6, y: 1.45, w: 6.1, h: 0.3, fontSize: 11, bold: true, color: MUTED, margin: 0 });
s.addNotes("Lyndsay confirmed we should measure enrolment at the credit claim point, about 6 weeks, not day one.");

// ================= 6. Why numbers look too high =================
pres.addSection({ title: "Findings" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Findings" });
s.addText("Why the numbers look too high, and the fix", { placeholder: "title" });
s.addText("Sarah felt the figures were overestimated. These are the likely causes", { placeholder: "body" });
const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
s.addTable([
  [hdr("Problem"), hdr("Evidence"), hdr("Fix")],
  ["Total applications include duplicates, withdrawals and ineligible applicants", "YCCJ/AL: 116 total vs 66 live", "Use live applications; show total and live side by side"],
  ["The conversion page shows this year, still in progress", "Lesley: \"I think that is 26/27\"", "Base the baseline on final, completed years"],
  ["Waiting-list and extra offers inflate offers", "YCCJ/AL: 36 offers for 14 places", "Calculate offer → enrolment per course"],
  ["Day-one enrolment overstates credits", "10% of credits lost", "Measure enrolment at the ~6-week claim point"],
  ["A/B sections split the same course", "One section with 0 applications but 17 enrolled", "Group sections under one parent course"],
  ["No history of application statuses", "Only current status is stored", "Use final status for past years; start weekly snapshots now"],
], { objectName: "too high table", x: 0.6, y: 1.8, w: 12.1, colW: [4.4, 3.5, 4.2], fontSize: 12.5, color: C.text1, border: { type: "solid", pt: 0.75, color: "D5D6E3" }, fill: { color: C.background1 }, rowH: 0.66, valign: "middle", margin: [0, 0.1, 0, 0.1] });

// ================= 7. RA5 evidence (chart) =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Findings" });
s.addText("First check of the RA5 report: the numbers don't match", { placeholder: "title" });
s.addText("Application-to-offer rate for two construction courses, worked out three ways (2026 session, 30 Sep)", { placeholder: "body" });
s.addChart(pres.charts.BAR, [
  { name: "Offers ÷ total applications", labels: ["YCCJ/AL", "YCCMT/AL"], values: [31.0, 17.2] },
  { name: "Offers ÷ live applications", labels: ["YCCJ/AL", "YCCMT/AL"], values: [54.5, 24.2] },
  { name: "RA5 Conversion page", labels: ["YCCJ/AL", "YCCMT/AL"], values: [37.9, 21.8] },
], {
  objectName: "RA5 check chart", x: 0.6, y: 1.75, w: 7.4, h: 5.0, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
  chartColors: ["B9B7F5", THEME.colors.accent1, THEME.colors.accent3],
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0\"%\"", dataLabelFontSize: 12, dataLabelColor: THEME.colors.dk1, dataLabelFontFace: "+mn-lt",
  showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: THEME.colors.dk1,
  catAxisLabelColor: THEME.colors.dk1, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt",
  valAxisHidden: true, valAxisMaxVal: 65, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
});
const ev = [
  ["The RA5 rate sits between the two", "It matches neither total nor live applications, so the conversion measure uses a different application count. This is the main question for Lesley.", C.accent3, AMBER_TINT],
  ["Only 2026 is visible", "The Session filter shows 2026 only, and the Summary page has no session filter. We need past years for a 3-year baseline.", C.accent1, C.background2],
  ["Source tables not visible", "Lineage shows semantic model → RA5 report, but not the database behind it. Kevin can confirm which tables feed it.", C.accent1, C.background2],
];
ev.forEach((e, i) => {
  const y = 1.8 + i * 1.68;
  card(s, "evidence " + (i + 1), 8.3, y, 4.4, 1.5, e[3]);
  s.addText([{ text: e[0], options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } }, { text: e[1], options: { fontSize: 12, color: C.text1 } }], { isTextBox: true, x: 8.5, y: y + 0.05, w: 4.05, h: 1.4, valign: "middle", margin: 0 });
});
s.addNotes("YCCJ/AL: 36 offers, 116 total applications, 66 live. YCCMT/AL: 15 offers, 87 total, 62 live. A later screenshot showed YCCJ/AL at 40.0% under different filter settings, so filters need recording with every number.");

// ================= 8. Data limits =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Findings" });
s.addText("Data limits to agree before we build", { placeholder: "title" });
s.addText("Lesley's concerns from the meeting, alongside the gaps the data map shows", { placeholder: "body" });
card(s, "lesley col", 0.6, 1.8, 5.9, 4.9, C.background2);
s.addText("Lesley's concerns", { isTextBox: true, x: 0.9, y: 1.95, w: 5.3, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0 });
s.addText([
  { text: "No audit trail of application statuses, only the current one", options: { bullet: true, breakLine: true } },
  { text: "Only live applications are reliable", options: { bullet: true, breakLine: true } },
  { text: "The conversion page is 2026/27, still in progress", options: { bullet: true, breakLine: true } },
  { text: "Not sure the 6-week enrolment point can be pulled", options: { bullet: true, breakLine: true } },
  { text: "Over-recruiting could mean students \"haven't got a seat\"", options: { bullet: true } },
], { isTextBox: true, x: 0.9, y: 2.5, w: 5.35, h: 4.0, fontSize: 15, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0 });
card(s, "map col", 6.8, 1.8, 5.9, 4.9, AMBER_TINT);
s.addText("What the data map adds", { isTextBox: true, x: 7.1, y: 1.95, w: 5.3, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0 });
s.addText([
  { text: "23 of 98 system connections are manual", options: { bullet: true, breakLine: true } },
  { text: "Enquirer → UnitE is a manual step on the recruitment path", options: { bullet: true, breakLine: true } },
  { text: "Student destinations are assembled by hand from UCAS, NAD and universities", options: { bullet: true, breakLine: true } },
  { text: "Five systems feed Power BI through one SQL Server connector", options: { bullet: true, breakLine: true } },
  { text: "RA5's semantic model source is not yet confirmed", options: { bullet: true } },
], { isTextBox: true, x: 7.1, y: 2.5, w: 5.35, h: 4.0, fontSize: 15, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0 });

// ================= 9. Framework =================
pres.addSection({ title: "Plan" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Plan" });
s.addText("How we will build the baseline", { placeholder: "title" });
s.addText("Seven steps, starting with construction as the pilot area", { placeholder: "body" });
const steps = [
  ["Collect", "3 final years (2023/24 to 2025/26), course level"],
  ["Clean", "Live applications; group A/B sections; match code changes"],
  ["Calculate", "App → offer, offer → enrolment, enrolment → 6 weeks"],
  ["Work back", "Target enrolments ÷ rates = target offers and applications"],
  ["Compare", "Course-specific targets vs the 200% / 120% rule"],
  ["Flag", "Low converters, drop-off stage, small or new courses"],
  ["Review", "Lesley and Lyndsay, then Directors and Sarah"],
];
const fw = 1.6, fg = 0.15;
steps.forEach((st, i) => {
  const x = 0.6 + i * (fw + fg + 0.007);
  numDot(s, i + 1, x + fw / 2 - 0.3, 1.95, 0.6, i === 3 ? C.accent2 : C.accent1);
  card(s, "step " + (i + 1), x, 2.8, fw, 2.6, i === 3 ? C.accent6 : C.background2);
  s.addText(st[0], { isTextBox: true, x: x + 0.1, y: 2.95, w: fw - 0.2, h: 0.45, fontSize: 15, bold: true, color: C.text2, align: "center", margin: 0 });
  s.addText(st[1], { isTextBox: true, x: x + 0.1, y: 3.5, w: fw - 0.2, h: 1.8, fontSize: 12.5, color: C.text1, align: "center", valign: "top", margin: 0 });
});
card(s, "example", 0.6, 5.75, 12.1, 0.95, C.accent5);
s.addText([{ text: "Example of step 4:  ", options: { bold: true, color: C.text2 } }, { text: "if 75% of a course's offers become enrolments, 20 places need about 27 offers, not the 24 the blanket rule gives." }], { isTextBox: true, x: 0.9, y: 5.75, w: 11.6, h: 0.95, fontSize: 14, color: C.text1, valign: "middle", margin: 0 });
s.addNotes("The 75% in the example is illustrative only, to show how the calculation works. Real rates come from the 3-year data. The baseline will not stay fixed: courses rise and fall in popularity, so show the trend as well as the average.");

// ================= 10. Progress =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Plan" });
s.addText("Progress so far", { placeholder: "title" });
s.addText("Where each piece of work stands at the start of October", { placeholder: "body" });
const prog = [
  ["Done", C.accent2, C.accent6, [
    "Meeting summary, report and plan",
    "Student system data map (209 systems, 98 connections), aligned with Kevin",
    "Plan and data request sent to Lyndsay, cc Darren",
    "RA5 exploration log set up",
  ]],
  ["In progress", C.accent1, C.accent5, [
    "Exploring RA5: pages, filters, definitions",
    "Hand checks of conversion rates, course by course",
    "Recording filters with every number",
  ]],
  ["Waiting on others", C.accent3, AMBER_TINT, [
    "Lesley: conversion measure, past years, credit statuses",
    "Lyndsay: targets, capacities, credits report, waiting list",
    "Kevin: tables behind the RA5 semantic model",
  ]],
];
prog.forEach((p, i) => {
  const x = 0.6 + i * 4.1;
  card(s, "progress " + p[0], x, 1.8, 3.85, 4.9, p[2]);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "progress tag " + p[0], x: x + 0.3, y: 2.05, w: 2.0, h: 0.45, rectRadius: 0.22, fill: { color: p[1] }, line: { color: p[1] } });
  s.addText(p[0], { isTextBox: true, x: x + 0.3, y: 2.05, w: 2.0, h: 0.45, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  s.addText(p[3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < p[3].length - 1 } })), { isTextBox: true, x: x + 0.3, y: 2.75, w: 3.3, h: 3.8, fontSize: 14, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0 });
});

// ================= 11. Next steps =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Plan" });
s.addText("Next steps and what I need", { placeholder: "title" });
s.addText("Aim: a construction prototype ready for the check-in after the October break; the full baseline before Christmas", { placeholder: "body" });
const tl = [
  ["Now", "Collect answers from Lesley, Lyndsay and Kevin; finish the definitions sheet and issues log"],
  ["Pilot", "Construction prototype: 3 years, live applications, course targets vs the 200% / 120% rule"],
  ["After the October break", "Meet Lyndsay and Darren on first findings, before the next meeting with Sarah"],
  ["Before Christmas", "Full course baseline reviewed by Lesley and Lyndsay, then Directors and Sarah"],
];
tl.forEach((t, i) => {
  const x = 0.6 + i * 3.08;
  numDot(s, i + 1, x + 1.07, 1.9, 0.7, C.accent1);
  if (i < 3) s.addShape(pres.shapes.LINE, { x: x + 1.85, y: 2.25, w: 2.38, h: 0, line: { color: C.accent5, width: 2 } });
  card(s, "timeline " + (i + 1), x, 2.85, 2.85, 2.0, C.background2);
  s.addText(t[0], { isTextBox: true, x: x + 0.1, y: 2.95, w: 2.65, h: 0.45, fontSize: 15, bold: true, color: C.text2, align: "center", margin: 0 });
  s.addText(t[1], { isTextBox: true, x: x + 0.2, y: 3.45, w: 2.45, h: 1.35, fontSize: 12.5, color: C.text1, align: "center", valign: "top", margin: 0 });
});
const asks = [
  ["Lesley", "Which application count the conversion measure uses; past-year data; which statuses count for credits"],
  ["Lyndsay", "Target enrolments and capacities per course; the credits report; how the waiting list and January starts work"],
  ["Kevin", "Which database tables feed the RA5 semantic model"],
];
asks.forEach((a, i) => {
  const x = 0.6 + i * 4.1;
  card(s, "ask " + a[0], x, 5.1, 3.85, 1.6, C.text2);
  s.addText([{ text: a[0], options: { bold: true, fontSize: 15, breakLine: true } }, { text: a[1], options: { fontSize: 12 } }], { isTextBox: true, x: x + 0.25, y: 5.15, w: 3.4, h: 1.5, color: C.background1, valign: "middle", margin: 0 });
});

(async () => {
  const out = path.join(__dirname, "Conversion_Baseline_Data_Architecture_Summary.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})();

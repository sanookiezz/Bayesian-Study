// Builds "Student Data Architecture" high-level deck from the Lucid working file
// (System Architecture-Working File-Sanook Update 20261001-16:18).
// Run: node build_deck.js  -> Student_Data_Architecture.pptx
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL_DIR + "/scripts/apply_theme.js");

const THEME = {
  name: "FVC Data Architecture",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2233", lt1: "FFFFFF", dk2: "2B2770", lt2: "F3F3FB",
    accent1: "5B57E8", // core systems (indigo, as in Lucid)
    accent2: "1F9D6A", // student-facing outputs (green)
    accent3: "E8912D", // manual hand-offs (amber)
    accent4: "3A414A", // data connector (charcoal)
    accent5: "DEDEFF", // student-system tint
    accent6: "C3F0D6", // output tint
    hlink: "5B57E8", folHlink: "2B2770",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Student Data Architecture";
pres.author = "Thanaporn Wattanamanon";
const C = pres.SchemeColor;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 2.1, w: 11.9, h: 1.6, fontSize: 44, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.9, w: 11.9, h: 1.4, fontSize: 20, color: C.accent5, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.1, w: 12.1, h: 0.45, fontSize: 15, color: C.text1, valign: "top", margin: 0 }, text: "" } },
    { text: { text: "Student Data Architecture  |  Forth Valley College", options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: "6B6F80", margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: "6B6F80", align: "right" },
});

// ---------- helpers ----------
function box(slide, name, x, y, w, h, title, sub, o = {}) {
  const fill = o.fill || C.accent5;
  const tc = o.color || C.text1;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    objectName: name, x, y, w, h, rectRadius: 0.08,
    fill: { color: fill }, line: { color: o.line || fill, width: o.lineW || 1 },
    shadow: o.shadow ? { type: "outer", color: "000000", opacity: 0.15, blur: 4, offset: 2, angle: 90 } : undefined,
  });
  const runs = [{ text: title, options: { bold: true, fontSize: o.fs || 13, color: tc, breakLine: !!sub } }];
  if (sub) runs.push({ text: sub, options: { fontSize: o.sfs || 10.5, color: tc } });
  slide.addText(runs, { isTextBox: true, objectName: name + " text", x: x + 0.05, y, w: w - 0.1, h, align: o.align || "center", valign: "middle", margin: 2 });
}

// Arrow from (x1,y1) to (x2,y2)
function arrow(slide, x1, y1, x2, y2, o = {}) {
  const x = Math.min(x1, x2), y = Math.min(y1, y2);
  slide.addShape(pres.shapes.LINE, {
    x, y, w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: o.color || C.accent1, width: o.width || 1.75, dashType: o.dash ? "dash" : "solid", endArrowType: "triangle", beginArrowType: o.both ? "triangle" : undefined },
  });
}

function legend(slide, x, y) {
  arrow(slide, x, y + 0.15, x + 0.6, y + 0.15, { color: C.accent1 });
  slide.addText("Automated feed", { isTextBox: true, x: x + 0.7, y, w: 1.6, h: 0.3, fontSize: 11, color: C.text1, margin: 0, valign: "middle" });
  arrow(slide, x + 2.4, y + 0.15, x + 3.0, y + 0.15, { color: C.accent3, dash: true });
  slide.addText("Manual / file hand-off", { isTextBox: true, x: x + 3.1, y, w: 2.0, h: 0.3, fontSize: 11, color: C.text1, margin: 0, valign: "middle" });
}

// ================= 1. Title =================
pres.addSection({ title: "Introduction" });
let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
s.addText("Student Data Architecture", { placeholder: "title" });
s.addText("How data flows in, around and out of our student systems — a high-level view\nFollow-up to the 1-Page Summary discussion with Sarah, Lyndsay and Lesley", { placeholder: "body" });
s.addText("Thanaporn (Sanook) Wattanamanon  |  October 2026", { isTextBox: true, x: 0.7, y: 6.4, w: 8, h: 0.4, fontSize: 14, color: C.background1, margin: 0 });
s.addNotes("Purpose: turn the detailed Lucid system map into a picture everyone can read in a minute. The detailed map stays in Lucid as the working reference; this deck is the summary view.");

// ================= 2. Recap of 1-page summary =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Introduction" });
s.addText("Where we left off: the 1-Page Summary", { placeholder: "title" });
s.addText("Recap from last Wednesday, and what this deck adds", { placeholder: "body" });
const recap = [
  ["Why we are doing this", "Give everyone one shared picture of which systems hold student and staff data, and how that data moves between them.", C.accent5],
  ["What we agreed last Wednesday", "[Add the key decisions from the 1-Page Summary here]", C.background2],
  ["Open questions from the meeting", "[Add the open questions / actions from the 1-Page Summary here]", C.background2],
  ["What this deck adds", "The finished system-connection map, simplified to four layers, the student journey, the reporting feeds and the manual hand-offs.", C.accent6],
];
recap.forEach((r, i) => {
  const x = 0.6 + (i % 2) * 6.2, y = 1.85 + Math.floor(i / 2) * 2.5;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "recap card " + (i + 1), x, y, w: 5.9, h: 2.2, rectRadius: 0.1, fill: { color: r[2] }, line: { color: r[2] } });
  s.addShape(pres.shapes.OVAL, { objectName: "recap num " + (i + 1), x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: C.text2 }, line: { color: C.text2 } });
  s.addText(String(i + 1), { isTextBox: true, x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  s.addText(r[0], { isTextBox: true, x: x + 1.05, y: y + 0.3, w: 4.6, h: 0.55, fontSize: 18, bold: true, color: C.text2, valign: "middle", margin: 0 });
  s.addText(r[1], { isTextBox: true, x: x + 1.05, y: y + 0.95, w: 4.6, h: 1.1, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
});
s.addNotes("Fill boxes 2 and 3 from the 1-Page Summary so this slide links the earlier meeting to today's picture.");

// ================= 3. At a glance =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Introduction" });
s.addText("The landscape at a glance", { placeholder: "title" });
s.addText("What the detailed Lucid map now covers", { placeholder: "body" });
const stats = [
  ["209", "systems and tools mapped across the college"],
  ["98", "connections drawn between them"],
  ["23", "of those connections are manual hand-offs"],
  ["1", "central hub: UnitE holds the core student and staff record"],
];
stats.forEach((st, i) => {
  const x = 0.6 + i * 3.08;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "stat card " + (i + 1), x, y: 1.9, w: 2.85, h: 2.4, rectRadius: 0.1, fill: { color: i === 2 ? "FDF0DF" : C.background2 }, line: { color: i === 2 ? "FDF0DF" : C.background2 } });
  s.addText(st[0], { isTextBox: true, x, y: 2.0, w: 2.85, h: 1.2, fontSize: 60, bold: true, color: i === 2 ? C.accent3 : C.accent1, align: "center", valign: "middle", margin: 0 });
  s.addText(st[1], { isTextBox: true, x: x + 0.25, y: 3.2, w: 2.35, h: 0.95, fontSize: 14, color: C.text1, align: "center", valign: "top", margin: 0 });
});
s.addText("Areas covered", { isTextBox: true, x: 0.6, y: 4.65, w: 6, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0 });
const areas = ["Student system", "Academic & VLE", "Quality & assessment", "HR", "Finance", "IT", "Estates & H&S", "Apprenticeships (ASC)", "Marketing & commercial", "Hospitality", "University & funding (external)", "Intelligence reporting"];
areas.forEach((a, i) => {
  const x = 0.6 + (i % 4) * 3.08, y = 5.15 + Math.floor(i / 4) * 0.58;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "area " + a, x, y, w: 2.85, h: 0.45, rectRadius: 0.22, fill: { color: i === 0 || i === 11 ? C.accent1 : C.background2 }, line: { color: i === 0 || i === 11 ? C.accent1 : C.background2 } });
  s.addText(a, { isTextBox: true, x, y, w: 2.85, h: 0.45, fontSize: 12, color: i === 0 || i === 11 ? C.background1 : C.text1, align: "center", valign: "middle", margin: 0 });
});
s.addNotes("Highlighted chips are the focus of this deck: the student system and the reporting layer. The other areas are mapped in Lucid and can be presented separately.");

// ================= 4. Four layers =================
pres.addSection({ title: "Architecture" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Architecture" });
s.addText("The big picture: four layers", { placeholder: "title" });
s.addText("Data is captured at the front door, worked on by operational systems, held in UnitE, then served back out and reported", { placeholder: "body" });
const cols = [
  { h: "1  Capture", d: "Where data enters", items: [["FVC Website", "forthvalley.ac.uk"], ["Student Registration", "Online Service"], ["School & Employer Portal", "FVC Portal"], ["Training Management", "Dante"]] },
  { h: "2  Engage & support", d: "Day-to-day student work", items: [["Enquiry Management", "Enquirer"], ["Timetable & Attendance", "Celcat"], ["Student Support", "TripleS"], ["Student Funding", "Bursary"]] },
  { h: "3  Core record", d: "Single source of truth", items: [] },
  { h: "4  Serve & report", d: "Where data goes out", items: [["Student Self-Service", "Student Portal"], ["Student Mobile App", "CampusM"], ["Student Info Website", "MyInfo"], ["BI Report", "Power BI"]] },
];
const cw = 2.75, gap = 0.37, top = 1.95;
cols.forEach((col, i) => {
  const x = 0.6 + i * (cw + gap);
  s.addText([{ text: col.h, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } }, { text: col.d, options: { fontSize: 11, color: "6B6F80" } }], { isTextBox: true, x, y: top - 0.1, w: cw, h: 0.65, margin: 0, valign: "top" });
  if (i === 2) {
    box(s, "UnitE core", x, top + 0.7, cw, 2.15, "Core MIS (UnitE)", "Student: details, ASN, application, enrolment, course\nStaff: details, organisation", { fill: C.accent1, color: C.background1, fs: 16, sfs: 11, shadow: true });
    box(s, "Dataserve", x, top + 3.0, cw, 0.62, "Student Record", "Dataserve", { fill: C.accent5 });
    box(s, "iTrent", x, top + 3.75, cw, 0.62, "HR & Payroll", "iTrent  (staff data in)", { fill: "FDF0DF" });
  } else {
    col.items.forEach((it, j) => {
      const out = i === 3;
      const bi = out && j === 3;
      box(s, it[0], x, top + 0.7 + j * 0.92, cw, 0.75, it[0], it[1], { fill: bi ? C.accent4 : out ? C.accent6 : C.accent5, color: bi ? C.background1 : C.text1 });
    });
  }
  if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { objectName: "flow arrow " + (i + 1), x: x + cw + 0.04, y: top + 1.55, w: gap - 0.08, h: 0.45, fill: { color: C.text2 }, line: { color: C.text2 } });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "external band", x: 0.6, y: 6.4, w: 12.1, h: 0.5, rectRadius: 0.1, fill: { color: C.background2 }, line: { color: C.background2 } });
s.addText([{ text: "External partners feeding in and out:  ", options: { bold: true, color: C.text2 } }, { text: "Schools (transition info)  ·  UCAS  ·  SAAS  ·  Skills Development Scotland  ·  Partner universities  ·  Award bodies (SQA, City & Guilds)", options: { color: C.text1 } }], { isTextBox: true, x: 0.8, y: 6.4, w: 11.7, h: 0.5, fontSize: 12, valign: "middle", margin: 0 });
s.addNotes("Read left to right. Layer 1 is where a student or school first gives us data. Layer 2 is the operational systems staff use every day. Layer 3 is UnitE, the core record everything else syncs with. Layer 4 is what students see (portal, app, MyInfo) and what managers see (Power BI).");

// ================= 5. UnitE hub =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Architecture" });
s.addText("UnitE is the hub: what goes in and out", { placeholder: "title" });
s.addText("Every core student system exchanges data with UnitE; amber lines are still done by hand", { placeholder: "body" });
const cx = 6.67, cy = 4.15;
box(s, "hub UnitE", cx - 1.3, cy - 0.6, 2.6, 1.2, "UnitE", "Core MIS (student & staff)", { fill: C.accent1, color: C.background1, fs: 20, sfs: 11, shadow: true });
// [title, sub, x, y, outAuto, inManual, outManual, inAuto]
const sp = [
  ["Timetable & Attendance", "Celcat · timetable, attendance", 0.6, 1.75, 1, 0, 0, 0],
  ["Enquiry Management", "Enquirer · application status", 4.15, 1.75, 1, 1, 0, 0],
  ["Student Support", "TripleS · ASN, school transition", 7.7, 1.75, 1, 0, 0, 0],
  ["Student Funding", "Bursary · detail, award, evidence", 10.25, 2.75, 1, 1, 0, 0],
  ["Student Self-Service", "Student Portal · enrolment, course", 10.25, 4.3, 1, 1, 0, 0],
  ["Data Connector", "SQL Server → Power BI", 10.25, 5.6, 1, 0, 0, 0],
  ["Student Record", "Dataserve", 7.0, 5.6, 1, 0, 0, 0],
  ["Apprenticeship E-portfolio", "Onefile", 4.05, 5.6, 0, 0, 1, 0],
  ["HR & Payroll", "iTrent · new / transfer / leaver", 0.6, 5.6, 0, 1, 0, 0],
  ["Manual data entry", "Forms keyed in by staff", 0.6, 4.3, 0, 1, 0, 0],
  ["School & Employer Portal", "FVC Portal · school transition", 0.6, 2.85, 1, 1, 0, 0],
];
const bw = 2.45, bh = 0.85;
sp.forEach((p) => {
  const [t, sub, x, y, oa, im, om] = p;
  const manualOnly = !oa && !p[7];
  box(s, "spoke " + t, x, y, bw, bh, t, sub, { fill: manualOnly ? "FDF0DF" : C.accent5, fs: 12, sfs: 10 });
  // edge points: spoke centre -> hub centre, clipped to box edges
  const sx = x + bw / 2, sy = y + bh / 2;
  const dx = cx - sx, dy = cy - sy;
  const clip = (hw, hh) => Math.min(hw / Math.abs(dx || 1e-6), hh / Math.abs(dy || 1e-6));
  const ts = clip(bw / 2 + 0.05, bh / 2 + 0.05), th = clip(1.35, 0.65);
  const p1 = [sx + dx * ts, sy + dy * ts], p2 = [cx - dx * th, cy - dy * th];
  const len = Math.hypot(dx, dy), nx = -dy / len * 0.07, ny = dx / len * 0.07;
  const two = (oa || om) && im;
  const off = two ? 1 : 0;
  if (oa) arrow(s, p2[0] + nx * off, p2[1] + ny * off, p1[0] + nx * off, p1[1] + ny * off, { color: C.accent1 });
  if (om) arrow(s, p2[0], p2[1], p1[0], p1[1], { color: C.accent3, dash: true });
  if (im) arrow(s, p1[0] - nx * off, p1[1] - ny * off, p2[0] - nx * off, p2[1] - ny * off, { color: C.accent3, dash: true });
});
legend(s, 0.6, 6.65);
s.addNotes("UnitE pushes data out automatically to most systems (indigo). The amber lines are where data comes back into UnitE by hand: from Enquirer, Bursary, the Student Portal and FVC Portal, from iTrent for staff starters/leavers, and from manual forms. Onefile also receives its learner data by hand.");

// ================= 6. Student journey =================
pres.addSection({ title: "Data flows" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Data flows" });
s.addText("Following the student: which system holds what, when", { placeholder: "title" });
s.addText("The same student record travels through six stages, touching a different system at each", { placeholder: "body" });
const stages = [
  ["Attract & enquire", "FVC Website\nCRM (RUBI)\nEnquirer", "Enquiry, contact details"],
  ["Apply", "Student Registration\nFVC Portal (schools)\nDante (commercial)", "Application, school transition"],
  ["Enrol", "UnitE", "Student, course, enrolment"],
  ["Learn & attend", "Celcat\nMoodle (VLE)\nOnefile / JTL", "Timetable, attendance, e-portfolio"],
  ["Support & fund", "TripleS\nBursary\nSAAS (external)", "ASN, PEEP, bursary awards"],
  ["Achieve & progress", "Dataserve\nUCAS · NAD\nPartner universities", "Results, destinations"],
];
const sw = 1.95, sg = 0.08;
stages.forEach((st, i) => {
  const x = 0.6 + i * (sw + sg);
  s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { objectName: "stage " + (i + 1), x, y: 1.85, w: sw + 0.12, h: 0.8, fill: { color: i === 2 ? C.accent1 : C.text2 }, line: { color: i === 2 ? C.accent1 : C.text2 } });
  s.addText(st[0], { isTextBox: true, x: x + 0.3, y: 1.85, w: sw - 0.4, h: 0.8, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "stage systems " + (i + 1), x, y: 2.9, w: sw - 0.05, h: 1.75, rectRadius: 0.08, fill: { color: C.accent5 }, line: { color: C.accent5 } });
  s.addText([{ text: "Systems", options: { bold: true, fontSize: 10, color: "6B6F80", breakLine: true } }, { text: st[1], options: { fontSize: 12, color: C.text1 } }], { isTextBox: true, x: x + 0.12, y: 2.95, w: sw - 0.25, h: 1.65, valign: "top", margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "stage data " + (i + 1), x, y: 4.8, w: sw - 0.05, h: 1.05, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 } });
  s.addText([{ text: "Data", options: { bold: true, fontSize: 10, color: "6B6F80", breakLine: true } }, { text: st[2], options: { fontSize: 12, color: C.text1 } }], { isTextBox: true, x: x + 0.12, y: 4.85, w: sw - 0.25, h: 0.95, valign: "top", margin: 0 });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "reporting band", x: 0.6, y: 6.05, w: 12.1, h: 0.75, rectRadius: 0.1, fill: { color: C.accent4 }, line: { color: C.accent4 } });
s.addText([{ text: "Across every stage:  ", options: { bold: true } }, { text: "UnitE, Enquirer, TripleS, Bursary and iTrent feed the SQL Server data connector, which powers Power BI reporting" }], { isTextBox: true, x: 0.85, y: 6.05, w: 11.6, h: 0.75, fontSize: 13, color: C.background1, valign: "middle", margin: 0 });
s.addNotes("This is the same architecture seen from the student's point of view. UnitE (highlighted) is where the student officially becomes a student; before that, data lives in the website, CRM and Enquirer.");

// ================= 7. Reporting =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Data flows" });
s.addText("Reporting: five systems feed one Power BI layer", { placeholder: "title" });
s.addText("All automated feeds land in the SQL Server data connector before reaching Power BI", { placeholder: "body" });
const feeds = [
  ["UnitE", "Student details, ASN, application, enrolment, course; employee details, organisation"],
  ["Enquirer", "Bursary engagement, curriculum planning, attendance, PEEP, care placement"],
  ["iTrent", "All employee data: detail, work pattern, payroll, job role"],
  ["TripleS", "ASN and school transition"],
  ["Bursary", "Applications and awards"],
];
feeds.forEach((f, i) => {
  const y = 1.85 + i * 0.95;
  box(s, "feed " + f[0], 0.6, y, 1.6, 0.78, f[0], null, { fill: C.accent1, color: C.background1, fs: 15 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "feed detail " + f[0], x: 2.3, y, w: 4.6, h: 0.78, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 } });
  s.addText(f[1], { isTextBox: true, x: 2.45, y, w: 4.35, h: 0.78, fontSize: 12, color: C.text1, valign: "middle", margin: 0 });
  arrow(s, 6.95, y + 0.39, 7.75, 4.24, { color: C.accent1, width: 1.5 });
});
box(s, "connector", 7.8, 3.6, 2.2, 1.3, "Data Connector", "SQL Server", { fill: C.accent4, color: C.background1, fs: 16, sfs: 12, shadow: true });
arrow(s, 10.05, 4.25, 10.5, 4.25, { color: C.text2, width: 2.25 });
box(s, "powerbi", 10.55, 3.6, 2.15, 1.3, "BI Report", "Power BI", { fill: C.accent5, line: C.accent1, lineW: 1.5, fs: 16, sfs: 12, shadow: true });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "side feeds", x: 7.8, y: 5.35, w: 4.9, h: 1.2, rectRadius: 0.08, fill: { color: "FDF0DF" }, line: { color: "FDF0DF" } });
s.addText([{ text: "Also reaching Power BI outside the connector", options: { bold: true, breakLine: true } }, { text: "Staff Helpdesk (UNIDESK) directly; MS Access databases by manual upload" }], { isTextBox: true, x: 7.95, y: 5.35, w: 4.6, h: 1.2, fontSize: 12, color: C.text1, valign: "middle", margin: 0 });
s.addNotes("This is the Intelligence Report part of the Lucid map. The connector is the one controlled route into Power BI. Feeds that bypass it are candidates to bring onto the connector.");

// ================= 8. Manual hand-offs =================
pres.addSection({ title: "Findings" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Findings" });
s.addText("Where data still moves by hand", { placeholder: "title" });
s.addText("23 of 98 connections are manual. These are the main ones on the student and staff side", { placeholder: "body" });
const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
const rows = [
  [hdr("From"), hdr("To"), hdr("What moves"), hdr("Why it matters")],
  ["Enquirer, Student Portal, FVC Portal", "UnitE", "Application and student details", "Re-keying risk on the core record"],
  ["Paper / form data", "UnitE, Enquirer", "Data filled in by staff on request", "Effort and delay at enrolment"],
  ["Schools", "FVC Portal", "School transition information", "Depends on each school uploading"],
  ["iTrent", "UnitE, Finance (Sun)", "Staff new / transfer / leaver, payroll", "Staff records can drift out of step"],
  ["UnitE", "Onefile", "Apprenticeship learners", "Two places to update apprentices"],
  ["Enquirer, UCAS, NAD, universities", "Student Destination Report", "Destinations and results", "Report assembled by hand"],
  ["SAAS files, Funding Tracker", "Student invoice to university", "Eligibility and funding status", "Billing relies on manual checks"],
  ["Restaurant (ResDiary), vending (NAYAX)", "Finance", "Monthly transaction reports", "Income reconciled monthly by hand"],
];
s.addTable(rows, { objectName: "manual table", x: 0.6, y: 1.8, w: 12.1, colW: [3.2, 2.6, 3.1, 3.2], fontSize: 12, color: C.text1, border: { type: "solid", pt: 0.75, color: "D5D6E3" }, fill: { color: C.background1 }, rowH: 0.5, valign: "middle", margin: [0, 0.08, 0, 0.08] });
s.addNotes("Each row is a dashed line in the Lucid map. Not all of these need automating; the aim is to agree which ones cost the most time or risk.");

// ================= 9. Key messages =================
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Findings" });
s.addText("Four things to take away", { placeholder: "title" });
s.addText("What the map tells us about how our data works today", { placeholder: "body" });
const msgs = [
  ["UnitE is the backbone", "Nearly every student system syncs with UnitE. Keeping it accurate keeps everything downstream accurate.", C.accent1],
  ["Data flows out well, but not back in", "UnitE sends data out automatically. Most return routes into UnitE are still manual.", C.accent3],
  ["Reporting already has one front door", "Five core systems feed Power BI through one SQL Server connector. A few side routes still bypass it.", C.accent4],
  ["The wider estate is large", "209 systems across 12 areas. Many sit outside the student system and are not yet connected.", C.accent2],
];
msgs.forEach((m, i) => {
  const x = 0.6 + (i % 2) * 6.2, y = 1.85 + Math.floor(i / 2) * 2.5;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "message card " + (i + 1), x, y, w: 5.9, h: 2.2, rectRadius: 0.1, fill: { color: C.background2 }, line: { color: C.background2 } });
  s.addShape(pres.shapes.OVAL, { objectName: "message dot " + (i + 1), x: x + 0.3, y: y + 0.32, w: 0.5, h: 0.5, fill: { color: m[2] }, line: { color: m[2] } });
  s.addText(String(i + 1), { isTextBox: true, x: x + 0.3, y: y + 0.32, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  s.addText(m[0], { isTextBox: true, x: x + 1.0, y: y + 0.3, w: 4.7, h: 0.55, fontSize: 18, bold: true, color: C.text2, valign: "middle", margin: 0 });
  s.addText(m[1], { isTextBox: true, x: x + 1.0, y: y + 0.95, w: 4.7, h: 1.1, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
});

// ================= 10. Next steps =================
pres.addSection({ title: "Next steps" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Next steps" });
s.addText("Proposed next steps", { placeholder: "title" });
s.addText("For discussion with Sarah, Lyndsay and Lesley", { placeholder: "body" });
const steps = [
  ["Validate", "Walk each system owner through their part of the map and confirm the connections."],
  ["Confirm manual steps", "Check open items, e.g. the HR starter form with Gillian and the weekly manual Moodle upload."],
  ["Prioritise", "Score each manual hand-off on time spent and data risk; pick the top three."],
  ["Extend reporting", "Bring side routes onto the SQL Server connector and agree the next Power BI feeds."],
];
steps.forEach((st, i) => {
  const x = 0.6 + i * 3.08;
  s.addShape(pres.shapes.OVAL, { objectName: "step dot " + (i + 1), x: x + 1.07, y: 1.95, w: 0.7, h: 0.7, fill: { color: C.accent1 }, line: { color: C.accent1 } });
  s.addText(String(i + 1), { isTextBox: true, x: x + 1.07, y: 1.95, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  if (i < 3) s.addShape(pres.shapes.LINE, { x: x + 1.85, y: 2.3, w: 2.38, h: 0, line: { color: C.accent5, width: 2 } });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "step card " + (i + 1), x, y: 2.9, w: 2.85, h: 2.1, rectRadius: 0.1, fill: { color: C.background2 }, line: { color: C.background2 } });
  s.addText(st[0], { isTextBox: true, x: x + 0.1, y: 3.05, w: 2.65, h: 0.45, fontSize: 16, bold: true, color: C.text2, align: "center", margin: 0 });
  s.addText(st[1], { isTextBox: true, x: x + 0.2, y: 3.55, w: 2.45, h: 1.35, fontSize: 13, color: C.text1, align: "center", valign: "top", margin: 0 });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "ask box", x: 0.6, y: 5.4, w: 12.1, h: 1.3, rectRadius: 0.1, fill: { color: C.text2 }, line: { color: C.text2 } });
s.addText([{ text: "What I need from you", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "Agreement that this four-layer view is the version we share, and a steer on which manual hand-offs to look at first." }], { isTextBox: true, x: 0.9, y: 5.4, w: 11.5, h: 1.3, fontSize: 14, color: C.background1, valign: "middle", margin: 0 });

(async () => {
  const out = path.join(__dirname, "Student_Data_Architecture.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})();

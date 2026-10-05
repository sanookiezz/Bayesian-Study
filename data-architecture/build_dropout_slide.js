// One-slide redesign of "Where students drop out" in the FVC magenta palette.
// Run: NODE_PATH=<dir with pptxgenjs> node build_dropout_slide.js
const pptxgen = require("pptxgenjs");
const path = require("path");

const MAGENTA = "A3206E", NAVY = "2B2350", PURPLE = "5B4A9E", AMBER = "D9822B", YELLOW = "E8C547";
const INK = "1F2233", MUTED = "6B6F80", GREY_BG = "F4F2F7", GREEN_BG = "E6F4EC";
const STAGE_FILL = ["9C2A7A", "B03C8C", "C2569D", "CB6FAB", PURPLE];

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: "Calibri", bodyFontFace: "Calibri" };
pres.title = "Where students drop out";
const s = pres.addSlide();
s.background = { color: "FFFFFF" };

const T = (text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: "Calibri" }, o));
const rr = (name, x, y, w, h, fill, line, dash) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  objectName: name, x, y, w, h, rectRadius: 0.06, fill: { color: fill },
  line: { color: line || fill, width: line ? 1.5 : 0.75, dashType: dash ? "dash" : "solid" },
});

// Title
T("Where students drop out, and which system shows it", { x: 0.4, y: 0.25, w: 11.2, h: 0.6, fontSize: 30, bold: true, color: MAGENTA });

// Funnel geometry
const CX = 4.2, H = 0.66;
const ROWS = [2.15, 3.0, 3.85, 4.95, 5.8];
const WIDTHS = [4.6, 4.1, 3.6, 3.1, 2.6];
const BOUNDARY = 4.73;

// Inputs: where applications come from
T("Applications come in from:", { x: 0.4, y: 1.05, w: 1.4, h: 0.5, fontSize: 11, bold: true, color: NAVY, valign: "middle" });
const inputs = [["Paper-based", "keyed by hand", true], ["Online Service", "Non-commercial"], ["Dante", "Commercial"], ["FVC Portal", "School partnership"]];
const iw = 1.05, ig = 0.13, ix0 = CX - (4 * iw + 3 * ig) / 2;
inputs.forEach((inp, i) => {
  const x = ix0 + i * (iw + ig);
  rr("input " + inp[0], x, 1.0, iw, 0.58, inp[2] ? "FFF6E5" : GREY_BG, inp[2] ? AMBER : "C9C3D6", inp[2]);
  T([{ text: inp[0], options: { bold: true, fontSize: 10, breakLine: true } }, { text: inp[1], options: { fontSize: 9 } }],
    { x, y: 1.0, w: iw, h: 0.58, color: INK, align: "center", valign: "middle" });
  s.addShape(pres.shapes.LINE, { x: x + iw / 2, y: 1.6, w: 0, h: 0.52, line: { color: inp[2] ? AMBER : "9A93AD", width: 1.5, dashType: inp[2] ? "dash" : "solid", endArrowType: "triangle" } });
});

// System bands
const band = (label, y0, y1) => {
  s.addShape(pres.shapes.RECTANGLE, { objectName: "system " + label, x: 0.4, y: y0, w: 0.55, h: y1 - y0, fill: { color: NAVY }, line: { color: NAVY } });
  T(label, { x: 0.4, y: y0, w: 0.55, h: y1 - y0, fontSize: 15, bold: true, color: "FFFFFF", align: "center", valign: "middle", vert: "vert270" });
};
T("System", { x: 0.4, y: 1.82, w: 0.9, h: 0.3, fontSize: 11, bold: true, color: NAVY });
band("Enquirer", ROWS[0], ROWS[2] + H);
band("UnitE", ROWS[3], ROWS[4] + H);

// Funnel stages
const stages = [
  ["1  Application", "Applied for the course"],
  ["2  Selection", "Seen by the department"],
  ["3  Offer", "Offer made"],
  ["4  Enrolment (day one)", "Starts the course"],
  ["5  Enrolled at ~6 weeks", "Credits claimed: counts for funding"],
];
stages.forEach((st, i) => {
  const w = WIDTHS[i], x = CX - w / 2, y = ROWS[i];
  rr("stage " + (i + 1), x, y, w, H, STAGE_FILL[i], i === 4 ? NAVY : null);
  T([{ text: st[0], options: { bold: true, fontSize: 15, breakLine: true } }, { text: st[1], options: { fontSize: 11 } }],
    { x, y, w, h: H, color: "FFFFFF", align: "center", valign: "middle" });
});

// Manual hand-off boundary between Enquirer and UnitE
s.addShape(pres.shapes.LINE, { objectName: "manual hand-off", x: 0.4, y: BOUNDARY, w: 8.45, h: 0, line: { color: AMBER, width: 2, dashType: "dash" } });
T("Manual hand-off: data moved from Enquirer into UnitE by hand", { x: 2.05, y: BOUNDARY - 0.14, w: 4.3, h: 0.28, fontSize: 10, italic: true, bold: true, color: AMBER, align: "center", valign: "middle", fill: { color: "FFFFFF" } });

// Conversion rates column
T("Rates we will measure", { x: 6.75, y: 1.82, w: 2.15, h: 0.3, fontSize: 11, bold: true, color: NAVY });
const pill = (letter, text, y) => {
  rr("rate " + letter, 6.75, y, 2.1, 0.44, "FFFFFF", PURPLE);
  s.addShape(pres.shapes.OVAL, { x: 6.81, y: y + 0.06, w: 0.32, h: 0.32, fill: { color: PURPLE }, line: { color: PURPLE } });
  T(letter, { x: 6.81, y: y + 0.06, w: 0.32, h: 0.32, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
  T(text, { x: 7.2, y, w: 1.6, h: 0.44, fontSize: 10.5, bold: true, color: INK, valign: "middle" });
};
pill("A", "Application → Offer", ROWS[1] + 0.11);
T("Selection is not counted on its own: Enquirer keeps only the current status.", { x: 6.75, y: ROWS[1] + 0.62, w: 2.1, h: 0.55, fontSize: 9, italic: true, color: MUTED, valign: "top" });
pill("B", "Offer → Enrolment (day one)", BOUNDARY - 0.22);
pill("C", "Day one → 6 weeks", (ROWS[3] + H + ROWS[4]) / 2 - 0.22);

// Causes of drop-off
T("Why we lose people (from the meeting)", { x: 9.05, y: 1.82, w: 3.9, h: 0.3, fontSize: 11, bold: true, color: NAVY });
const causes = [
  "Applicants don't turn up for selection (e.g. the construction skills test)",
  "Lecturers hold back, or make many waiting-list offers",
  "Offers turned down; waiting-list offers go to people unlikely to come",
  "Early dropout before credits are claimed at ~6 weeks",
];
causes.forEach((c, i) => {
  const y = ROWS[i];
  rr("cause " + (i + 1), 9.05, y, 3.88, H, GREY_BG);
  s.addShape(pres.shapes.RIGHT_ARROW, { objectName: "lost arrow " + (i + 1), x: 9.15, y: y + 0.2, w: 0.38, h: 0.26, fill: { color: YELLOW }, line: { color: YELLOW } });
  T(c, { x: 9.65, y, w: 3.2, h: H, fontSize: 11, color: INK, valign: "middle" });
});
rr("example", 9.05, ROWS[4], 3.88, H, GREEN_BG);
T([{ text: "Example (construction): ", options: { bold: true } }, { text: "14 places → 16–17 offers → only 9 started" }],
  { x: 9.2, y: ROWS[4], w: 3.65, h: H, fontSize: 11, color: INK, valign: "middle" });
s.addShape(pres.shapes.RIGHT_ARROW, { x: 9.05, y: 6.6, w: 0.3, h: 0.2, fill: { color: YELLOW }, line: { color: YELLOW } });
T("= applicants lost at this stage", { x: 9.4, y: 6.55, w: 3.5, h: 0.3, fontSize: 9.5, color: MUTED, valign: "middle" });

// Takeaway
rr("takeaway", 0.4, 6.92, 12.53, 0.45, NAVY);
T([{ text: "Key point: ", options: { bold: true } }, { text: "Enquirer shows stages 1–3 and UnitE shows 4–5. The hand-off is manual, so the numbers can differ. The headline for each course is the overall rate, Application → 6 weeks." }],
  { x: 0.6, y: 6.92, w: 12.2, h: 0.45, fontSize: 11.5, color: "FFFFFF", valign: "middle" });

s.addNotes(`This slide shows the journey of an applicant, from applying to being counted for funding, and where we lose people along the way.

Applications come in through four routes: paper forms, which are keyed in by hand, the Online Service for non-commercial courses, Dante for commercial courses, and the FVC Portal for school partnerships.

Every applicant moves through five stages: application, selection, offer, enrolment on day one, and still enrolled at about six weeks. The last stage is the one that counts for funding, because that is when we claim credits.

At each stage we lose people. From the meeting with Sarah: applicants don't turn up for selection, lecturers hold back or make many waiting-list offers, offers are turned down, and some students leave before six weeks. For example, one construction course had 14 places, made 16 to 17 offers, and only 9 people started.

Stages 1 to 3 live in Enquirer; stages 4 and 5 live in UnitE. The step between them is manual, so the numbers don't always match.

We will measure three rates: A, application to offer; B, offer to enrolment; C, day one to six weeks. The lowest rate shows where a course needs attention. Once Lesley and Kevin's combined data is ready, we can calculate these for every course over three years and move away from the single 200% / 120% rule.`);

pres.writeFile({ fileName: path.join(__dirname, "Dropout_Slide_Redesign.pptx") }).then((f) => console.log("wrote", f));

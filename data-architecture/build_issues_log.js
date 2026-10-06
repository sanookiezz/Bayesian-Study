// One-slide data quality issues log: student journey to Power BI.
// Source: Lucid "System Architecture-Working File" as checked on 6 Oct 2026.
// Run: NODE_PATH=<dir with pptxgenjs> node build_issues_log.js
const pptxgen = require("pptxgenjs");
const path = require("path");

const MAGENTA = "A3206E", NAVY = "2B2350", INK = "1F2233", MUTED = "6B6F80", LINE = "D9D4E3", ZEBRA = "F7F5FA";
const PRIORITY = { High: { fill: "F6D6E6", color: "8C1A5C" }, Medium: { fill: "FCEBD3", color: "8A4B0F" } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: "Calibri", bodyFontFace: "Calibri" };
pres.title = "Data quality issues log";
const s = pres.addSlide();
s.background = { color: "FFFFFF" };

s.addText("Data quality issues log: student journey to Power BI", { isTextBox: true, x: 0.4, y: 0.25, w: 12.5, h: 0.55, margin: 0, fontSize: 26, bold: true, color: MAGENTA });
s.addText("From the updated Lucid data architecture map: 8 issues, 4 high priority", { isTextBox: true, x: 0.4, y: 0.82, w: 12.5, h: 0.3, margin: 0, fontSize: 13, color: INK });

const issues = [
  ["High", "Same data keyed into several systems",
    "ASN typed into Enquirer, FVC Portal and TripleS. Application status and interview detail typed into both Enquirer and UnitE",
    "Each team works in its own system (admissions, registry, support, schools) and no one owns the master copy of each item",
    "Two versions of one number reach Power BI; may explain the RA5 gap (37.9% vs 31.0% / 54.5%)",
    "Agree one system of record per data item (Lesley, Kevin)"],
  ["High", "Manual steps on the recruitment path",
    "7 manual hand-offs, including Enquirer → UnitE, plus 6 manual input points",
    "No automated link was built between the systems, so re-keying became the routine workaround",
    "Delays, typing errors and missing records between applications and enrolments",
    "Monthly count check: Enquirer vs UnitE vs RA5, starting with YCCJ/AL"],
  ["High", "Two-way loops around UnitE",
    "UnitE sends data out automatically; Enquirer, Bursary, Student Portal and FVC Portal send it back by hand",
    "Other systems need to correct or add data but have no write-back link to UnitE",
    "No rule for which system wins; Power BI reads both UnitE and Enquirer",
    "Covered by the system-of-record agreement (Lesley, Kevin)"],
  ["High", "Students missing from UnitE or BI",
    "Dante \"not on Unit-E\" and not on the connector. Onefile, JTL and Dataserve unconnected. Celcat not on the connector",
    "Commercial and apprenticeship courses run on separate systems with their own processes; BI was built around UnitE and Enquirer",
    "Totals may quietly leave out commercial students and apprentices",
    "Ask Kevin which systems can join the SQL connector"],
  ["Medium", "Feeds that bypass the connector",
    "Complaints and UNIDESK feed Power BI directly; MS Access by hand. RA5's semantic model source unconfirmed",
    "Teams built quick reports for their own needs; no shared rule on how data must reach Power BI",
    "Different refresh timing and definitions; no central checks",
    "Kevin to confirm RA5's source; route feeds through the connector"],
  ["Medium", "Incomplete or unverified bursary feeds",
    "\"Can see only ready to outstanding award, not process evidence\"; status marked \"Finance need to check\"",
    "Evidence checks happen outside the Bursary system, and Finance confirms status separately",
    "Bursary pipeline in Power BI is only partly visible",
    "Confirm with Finance which Bursary data is complete"],
  ["Medium", "No status history",
    "Enquirer keeps only each application's current status",
    "Each status change overwrites the last one; no snapshot or audit table was set up",
    "Past offer counts cannot be rebuilt; 3-year trends understate offers",
    "Start weekly status snapshots now (Lesley)"],
  ["Medium", "Student systems outside the mapped flows",
    "7 systems in the \"Don't know\" group, incl. Theorise and CaseNotes; open \"Check with\" notes; bank details sent to Finance by hand",
    "Systems bought by individual teams without a central register; some owners have left",
    "ASN picture may be incomplete; sensitive data sits outside the managed flow",
    "Close open notes, Theorise and CaseNotes first"],
];

const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 10.5 } });
const rows = [[hdr("#"), hdr("Issue"), hdr("What the map shows"), hdr("Possible cause"), hdr("Effect on BI"), hdr("Next step (who)"), hdr("Priority")]];
issues.forEach((r, i) => {
  const band = i % 2 ? ZEBRA : "FFFFFF";
  const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fill: { color: band } }, o) });
  const p = PRIORITY[r[0]];
  rows.push([
    cell(String(i + 1), { bold: true, align: "center", color: MAGENTA }),
    cell(r[1], { bold: true, color: NAVY }),
    cell(r[2]), cell(r[3], { italic: true }), cell(r[4]), cell(r[5]),
    { text: r[0], options: { bold: true, align: "center", color: p.color, fill: { color: p.fill } } },
  ]);
});
s.addTable(rows, {
  objectName: "issues log", x: 0.4, y: 1.25, w: 12.53, colW: [0.33, 1.6, 2.85, 2.65, 2.25, 2.1, 0.75],
  rowH: [0.34, 0.68, 0.68, 0.68, 0.68, 0.68, 0.68, 0.68, 0.68],
  fontSize: 9, fontFace: "Calibri", color: INK, valign: "middle",
  border: { type: "solid", pt: 0.5, color: LINE }, margin: [0.03, 0.06, 0.03, 0.06],
});

s.addText("Source: System Architecture working file (Lucid), checked 6 Oct 2026. Causes are likely explanations to confirm with system owners. Line labels were matched to lines by position on the page.",
  { isTextBox: true, x: 0.4, y: 7.1, w: 12.5, h: 0.25, margin: 0, fontSize: 9, italic: true, color: MUTED });

s.addNotes(`The four high-priority issues all affect the conversion numbers directly.
1. The same data is keyed into several systems, so Power BI can receive two versions of one number. This may explain why RA5 shows 37.9% for YCCJ/AL when our own calculation gives 31.0% or 54.5%.
2. There are seven manual hand-offs on the recruitment path. The Enquirer to UnitE step sits exactly between applications and enrolments.
3. UnitE sends data out automatically but gets it back by hand, and nothing says which system wins.
4. Commercial students and apprentices may be missing from college totals, because Dante, Onefile, JTL and Dataserve are not connected to UnitE or the SQL connector.
The medium-priority issues are about feeds that bypass the connector, partial bursary data, the lack of status history, and student systems we have not yet mapped.
Possible causes (in italics) are my reading of the map and meeting notes; confirm them with each system owner. Most come back to two root causes: systems were bought team by team without automated links, and nobody owns the master copy of each data item.
The single most useful next step is to agree one system of record for each key data item, as part of Lesley and Kevin's single source work.`);

pres.writeFile({ fileName: path.join(__dirname, "Data_Quality_Issues_Log.pptx") }).then((f) => console.log("wrote", f));

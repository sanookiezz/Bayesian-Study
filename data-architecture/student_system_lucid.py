# Builds Lucid Standard Import JSON for the high-level Student System data architecture.
import json
CORE, SYS, OUT, WEB, MAN, DARK = "#635DFF", "#DEDEFF", "#C3F7C8", "#B8F5ED", "#FCFCCA", "#3A414A"
AUTO, MANL = "#3A414A", "#E8912D"
COLX = [40 + i * 660 for i in range(4)]
CY, CW, CH = 220, 420, 1060
shapes, lines = [], []

def txt(title, sub=None, size=16):
    s = f"<b>{title}</b>" + (f"<br>{sub}" if sub else "")
    return f'<p style="font-size:{size}px;text-align:center">{s}</p>'

def box(id, col, row, title, sub, fill, kind="rectangle", text_color=None, dashed=False):
    st = {"fill": {"type": "color", "color": fill}, "stroke": {"color": MANL if dashed else "#9AA0AA", "width": 2 if dashed else 1, "style": "dashed" if dashed else "solid"}}
    if text_color: st["textColor"] = text_color
    shapes.append({"id": id, "type": kind, "boundingBox": {"x": COLX[col] + 80, "y": CY + 80 + row * 160, "w": 260, "h": 100}, "text": txt(title, sub), "style": st})

def line(id, a, pa, b, pb, label=None, manual=False, both=False, ltype="straight", lpos=0.5, side="top"):
    l = {"id": id, "lineType": ltype,
         "endpoint1": {"type": "shapeEndpoint", "style": "arrow" if both else "none", "shapeId": a, "position": {"x": pa[0], "y": pa[1]}},
         "endpoint2": {"type": "shapeEndpoint", "style": "arrow", "shapeId": b, "position": {"x": pb[0], "y": pb[1]}},
         "stroke": {"color": MANL if manual else AUTO, "width": 2, "style": "dashed" if manual else "solid"}}
    if label: l["text"] = [{"text": f'<span style="font-size:12px">{label}</span>', "position": lpos, "side": side}]
    lines.append(l)

# Title
shapes.append({"id": "title", "type": "text", "boundingBox": {"x": 40, "y": 20, "w": 2400, "h": 70},
               "text": '<p style="font-size:28px"><b>Student System: high-level data architecture</b></p>'})
# Zone containers (drawn first so boxes sit on top)
zones = ["1. Capture: where data enters", "2. Engage and support", "3. Core record", "4. Student-facing services"]
for i, z in enumerate(zones):
    shapes.append({"id": f"zone{i+1}", "type": "roundedRectangleContainer", "boundingBox": {"x": COLX[i], "y": CY, "w": CW, "h": CH},
                   "containerTitle": {"text": z}, "style": {"fill": {"type": "color", "color": "#F5F8FF"}, "stroke": {"color": "#1071E5", "width": 2, "style": "solid"}}})
shapes.append({"id": "zone5", "type": "roundedRectangleContainer", "boundingBox": {"x": 700, "y": 1440, "w": 1740, "h": 260},
               "containerTitle": {"text": "Intelligence reporting (outside the student system)"},
               "style": {"fill": {"type": "color", "color": "#F2FBF4"}, "stroke": {"color": "#008A0E", "width": 2, "style": "solid"}}})

# 1 Capture
box("dante", 0, 0, "Training Management", "Dante (commercial courses)", SYS)
box("web", 0, 1, "FVC Website", "forthvalley.ac.uk", WEB)
box("reg", 0, 2, "Student Registration", "Online Service", SYS)
box("schoolup", 0, 3, "School transition upload", "Schools upload by hand", MAN, kind="manualInput", dashed=True)
box("fvcportal", 0, 4, "School and Employer Partnership", "FVC Portal", SYS)
box("manual", 0, 5, "Manual data entry", "Forms keyed in by staff", MAN, kind="manualInput", dashed=True)
# 2 Engage & support
box("enquirer", 1, 0, "Enquiry Management", "Enquirer", CORE, text_color="#FFFFFF")
box("celcat", 1, 1, "Timetable and Attendance", "Celcat", SYS)
box("triples", 1, 2, "Student Support", "TripleS (ASN, school transition)", SYS)
box("bursary", 1, 3, "Student Funding", "Bursary", SYS)
# 3 Core record
box("dataserve", 2, 0, "Student Record", "Dataserve", SYS)
box("unite", 2, 1, "Core MIS (UnitE)", "Student and staff record", CORE, text_color="#FFFFFF")
box("onefile", 2, 2, "Apprenticeship E-portfolio", "Onefile", SYS)
box("jtl", 2, 3, "Apprenticeship Training", "JTL Portal", SYS)
box("complaint", 2, 4, "Complaints", "PowerApp (standalone)", SYS)
# 4 Student-facing
box("portal", 3, 0, "Student Self-Service", "Student Portal", OUT)
box("campusm", 3, 1, "Student Mobile App", "CampusM", OUT)
box("myinfo", 3, 2, "Student Information Website", "MyInfo", OUT)
# Intelligence + HR
shapes.append({"id": "sql", "type": "rectangle", "boundingBox": {"x": 1440, "y": 1520, "w": 260, "h": 100}, "text": txt("Data Connector", "SQL Server"),
               "style": {"fill": {"type": "color", "color": DARK}, "stroke": {"color": DARK, "width": 1, "style": "solid"}, "textColor": "#FFFFFF"}})
shapes.append({"id": "pbi", "type": "rectangle", "boundingBox": {"x": 2100, "y": 1520, "w": 260, "h": 100}, "text": txt("BI Report", "Power BI"),
               "style": {"fill": {"type": "color", "color": SYS}, "stroke": {"color": CORE, "width": 2, "style": "solid"}}})
shapes.append({"id": "itrent", "type": "rectangle", "boundingBox": {"x": 120, "y": 1520, "w": 260, "h": 100}, "text": txt("HR and Payroll", "iTrent (outside the group)"),
               "style": {"fill": {"type": "color", "color": "#FFDDA6"}, "stroke": {"color": "#9AA0AA", "width": 1, "style": "solid"}}})

# Lines inside zones
line("l_web_dante", "web", (0.5, 0), "dante", (0.5, 1), both=True)
line("l_web_reg", "web", (0.5, 1), "reg", (0.5, 0))
line("l_school", "schoolup", (0.5, 1), "fvcportal", (0.5, 0), manual=True)
line("l_unite_ds", "unite", (0.5, 0), "dataserve", (0.5, 1))
line("l_unite_onefile", "unite", (0.5, 1), "onefile", (0.5, 0), manual=True)
line("l_jtl_onefile", "jtl", (0.5, 0), "onefile", (0.5, 1))
line("l_sql_pbi", "sql", (1, 0.5), "pbi", (0, 0.5))
# Lines between zones
line("z12", "zone1", (1, 0.22), "zone2", (0, 0.22), "Applications<br>(some keyed by hand)", manual=True)
line("z23m", "zone2", (1, 0.4), "zone3", (0, 0.4), "Applications into<br>UnitE (by hand)", manual=True)
line("z32", "zone3", (0, 0.62), "zone2", (1, 0.62), "Student, enrolment,<br>timetable")
line("z34", "zone3", (1, 0.3), "zone4", (0, 0.3), "Enrolment, course,<br>timetable")
line("z43m", "zone4", (0, 0.55), "zone3", (1, 0.55), "Student updates<br>(by hand)", manual=True)
for lid, a, b, lab, end in [("z24a", (1015, 220), (1015, 130), None, "none"), ("z24b", (1015, 130), (2335, 130), "Course, timetable and application status (Enquirer and Celcat to student apps)", "none"), ("z24c", (2335, 130), (2335, 220), None, "arrow")]:
    l = {"id": lid, "lineType": "straight", "endpoint1": {"type": "positionEndpoint", "style": "none", "position": {"x": a[0], "y": a[1]}},
         "endpoint2": {"type": "positionEndpoint", "style": end, "position": {"x": b[0], "y": b[1]}}, "stroke": {"color": AUTO, "width": 2, "style": "solid"}}
    if lab: l["text"] = [{"text": f'<span style="font-size:13px">{lab}</span>', "position": 0.5, "side": "top"}]
    lines.append(l)
line("z2sql", "zone2", (0.5, 1), "sql", (0, 0.3), "Enquirer, TripleS, Bursary", lpos=0.25, side="top")
line("z3sql", "zone3", (0.5, 1), "sql", (0.5, 0), "UnitE", lpos=0.5, side="middle")
line("hr_sql", "itrent", (1, 0.8), "sql", (0, 0.8), "Staff data", lpos=0.2)
# Legend
shapes.append({"id": "legend", "type": "text", "boundingBox": {"x": 40, "y": 1760, "w": 2400, "h": 60},
               "text": '<p style="font-size:14px"><b>Key:</b> solid line = automated feed &nbsp;|&nbsp; <span style="color:#E8912D"><b>amber dashed line</b></span> = manual hand-off &nbsp;|&nbsp; purple = core systems &nbsp;|&nbsp; green = what students see &nbsp;|&nbsp; yellow = manual input. Source: System Architecture working file, 1 Oct 2026.</p>'})

doc = {"version": 1, "pages": [{"id": "p1", "title": "Student System (high level)", "shapes": shapes, "lines": lines}]}
open("/tmp/claude-0/-home-user-Bayesian-Study/ecd1ff68-2de6-5c56-adb2-533eee8a3b7c/scratchpad/student_system.json", "w").write(json.dumps(doc, separators=(",", ":")))
print(len(shapes), len(lines))

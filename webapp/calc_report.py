"""Calculation report PDF generator.

Produces a per-slide PDF summary of:
  - source sheet(s)
  - filter / aggregation rule
  - actual computed numbers from the current Excel

Returned as raw PDF bytes so the Flask endpoint can stream it.
"""
import io
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, KeepTogether)
from reportlab.lib.enums import TA_LEFT

from slide_updaters import (
    _build_team_lookup, _last_two_sprint_cycles,
    _compute_manual_defects_by_subdomain, _compute_automation_prod_defects,
    _prod_defects_by_subdomain, _qa_miss_split, _monthly_buckets_in_data,
    _compute_rad_by_subdomain, _compute_rad_by_team,
    _compute_release_by_subdomain,
    _compute_regression_by_subdomain, _detect_latest_month_abbr,
    _automation_coverage_pct, _sprint_sort_key, _num,
)


# ----------------------- Styles -----------------------

def _styles():
    ss = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("Title", parent=ss["Title"], fontSize=20,
                                spaceAfter=4, textColor=colors.HexColor("#0b2c5f")),
        "subtitle": ParagraphStyle("Subtitle", parent=ss["Normal"], fontSize=10,
                                   textColor=colors.HexColor("#666666"), spaceAfter=14),
        "h1": ParagraphStyle("H1", parent=ss["Heading1"], fontSize=14,
                             textColor=colors.HexColor("#0b2c5f"),
                             spaceBefore=10, spaceAfter=4, keepWithNext=True),
        "h2": ParagraphStyle("H2", parent=ss["Heading2"], fontSize=11,
                             textColor=colors.HexColor("#0b2c5f"),
                             spaceBefore=6, spaceAfter=2, keepWithNext=True),
        "body": ParagraphStyle("Body", parent=ss["Normal"], fontSize=9,
                               leading=12, alignment=TA_LEFT, spaceAfter=4),
        "rule": ParagraphStyle("Rule", parent=ss["Normal"], fontSize=9,
                               leading=12, leftIndent=8,
                               textColor=colors.HexColor("#222222"), spaceAfter=4),
        "code": ParagraphStyle("Code", parent=ss["Code"], fontSize=8,
                               leading=10, textColor=colors.HexColor("#444444")),
    }


def _table(rows, col_widths=None, header=True):
    style = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbbbbb")),
    ]
    if header:
        style += [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2c5f")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ]
    t = Table(rows, colWidths=col_widths, repeatRows=1 if header else 0)
    t.setStyle(TableStyle(style))
    return t


# ----------------------- Section builders -----------------------

def _section_slides_6_9(data, st):
    flow = [Paragraph("Slides 6–9 · Presence of Cognizant QE Team", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>ProjectDetails</i>. "
        "<b>Rule:</b> group rows by <i>Sub Domain</i>, skip blank/TBD. "
        "For each panel: pentagon label ← Sub Domain name; "
        "SVT bar ← count of rows where Team Type = SVT; "
        "Non-SVT bar ← count of rows where Team Type = Non SVT. "
        "Static 'Teams and Key Feature Implemented' text boxes are NOT touched.",
        st["rule"]))
    rows = [["Sub Domain", "SVT", "Non SVT"]]
    tl = _build_team_lookup(data)
    for sd in sorted(tl):
        rows.append([sd, str(len(tl[sd]["SVT"])), str(len(tl[sd]["Non SVT"]))])
    flow.append(_table(rows, col_widths=[8 * cm, 2.5 * cm, 2.5 * cm]))
    return flow


def _section_slide_10(data, st):
    flow = [Paragraph("Slide 10 · Cognizant QE Highlights (1/2)", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>MonthlySheet</i>. "
        "<b>Pie + 'UI/API Scripts Created':</b> rows where "
        "<i>DashBoard Count &gt; 0</i> AND <i>Technology Used = Java</i> — "
        "pie slices are sums of Feature/Develop/Main Branch Count; "
        "UI/API number = pie total. "
        "<b>'Python Scripts Created':</b> rows where <i>Technology Used = Python</i> — "
        "sum of Feature+Develop+Main Branch Count.", st["rule"]))
    f = d = m = py = 0
    for r in data.get("monthly_sheet", []):
        dash = _num(r.get("DashBoard Count", 0))
        tech = str(r.get("Technology Used", "")).strip().lower()
        if dash > 0 and tech == "java":
            f += _num(r.get("Feature Branch Count", 0))
            d += _num(r.get("Develop Branch Count", 0))
            m += _num(r.get("Main Branch Count", 0))
        if tech == "python":
            py += (_num(r.get("Feature Branch Count", 0)) +
                   _num(r.get("Develop Branch Count", 0)) +
                   _num(r.get("Main Branch Count", 0)))
    rows = [["Element", "Calculation", "Value"],
            ["Feature (pie)", "sum, Java + DashBoard>0", f"{int(f)}"],
            ["Develop (pie)", "sum, Java + DashBoard>0", f"{int(d)}"],
            ["Main (pie)",    "sum, Java + DashBoard>0", f"{int(m)}"],
            ["UI/API Scripts Created", "Feature + Develop + Main", f"{int(f + d + m)}"],
            ["Python Scripts Created", "sum F+D+M where Technology = Python", f"{int(py)}"]]
    flow.append(_table(rows, col_widths=[5.5 * cm, 7.5 * cm, 2.5 * cm]))
    return flow


def _section_slide_11(data, st):
    flow = [Paragraph("Slide 11 · Cognizant QE Highlights (2/2)", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>DailyExecution</i>. "
        "<b>Filter:</b> Automation Execution Frequency ∈ {Daily, Weekly}. "
        "Critical = P0+P1, Core = P0+P1+P2, All = P0…P4. "
        "<b>Testcases configured in nightly batch</b> = P0+P1+P2 (same filter). "
        "<b>Pass rate</b> = sum(P0–P2 Passed TCs) / sum(P0+P1+P2) × 100.",
        st["rule"]))
    p = {k: 0 for k in ["P0", "P1", "P2", "P3", "P4"]}
    nb = 0; p02_tot = 0; p02_passed = 0
    for r in data.get("daily_execution", []):
        freq = str(r.get("Automation Execution Frequency ",
                          r.get("Automation Execution Frequency", ""))).strip().lower()
        if freq not in ("daily", "weekly"):
            continue
        for k in p:
            p[k] += _num(r.get(k, 0))
        nb += _num(r.get("P0", 0)) + _num(r.get("P1", 0)) + _num(r.get("P2", 0))
        p02_tot += _num(r.get("P0", 0)) + _num(r.get("P1", 0)) + _num(r.get("P2", 0))
        p02_passed += _num(r.get("P0 - P2 Passed TCs", 0))
    crit = p["P0"] + p["P1"]; core = crit + p["P2"]; all_t = core + p["P3"] + p["P4"]
    pct = round(p02_passed / p02_tot * 100) if p02_tot else 0
    rows = [["Element", "Calculation", "Value"],
            ["P0", "sum (Daily/Weekly)", f"{int(p['P0'])}"],
            ["P1", "sum", f"{int(p['P1'])}"],
            ["P2", "sum", f"{int(p['P2'])}"],
            ["P3", "sum", f"{int(p['P3'])}"],
            ["P4", "sum", f"{int(p['P4'])}"],
            ["Critical (bar)", "P0 + P1", f"{int(crit)}"],
            ["Core (bar)",     "P0 + P1 + P2", f"{int(core)}"],
            ["All (bar)",      "P0 + P1 + P2 + P3 + P4", f"{int(all_t)}"],
            ["Nightly batch number", "P0 + P1 + P2", f"{int(nb)}"],
            ["Pass rate label",
             f"P0-P2 Passed ({int(p02_passed)}) / P0+P1+P2 ({int(p02_tot)}) × 100",
             f"~{pct}%"]]
    flow.append(_table(rows, col_widths=[5.5 * cm, 7.5 * cm, 2.5 * cm]))
    return flow


def _section_slide_12(data, st):
    flow = [Paragraph("Slide 12 · Automation Snapshot", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>InSprintData</i>. "
        "<b>Rule for <i>&lt;70 %</i> and <i>&gt;70 %</i> In-Sprint bars:</b> "
        "Restrict to the last two sprint cycles found in InSprintData "
        "(numeric sort on Sprint name). For each unique <i>Scrum team</i>, "
        "take its <b>best</b> (maximum) <i>Automation coverage</i>. "
        "Count teams whose best ≥ 70 (→ &gt;70 % In-Sprint) vs &lt; 70 "
        "(→ &lt;70 % In-Sprint). The other three bars (No Automation, "
        "Script Maintainance, Tech-Debt Automation) are left at template values.",
        st["rule"]))
    insp = data.get("insprint_data", [])
    sprint_col = None
    if insp:
        for k in insp[0].keys():
            if str(k).strip().lower() == "sprint":
                sprint_col = k; break
    sprint_col = sprint_col or "Sprint  "
    sprints = sorted(
        {str(r.get(sprint_col, "")).strip() for r in insp
         if str(r.get(sprint_col, "")).strip().startswith("Sprint ")},
        key=_sprint_sort_key)[-2:]
    team_best = {}
    for r in insp:
        if str(r.get(sprint_col, "")).strip() not in sprints:
            continue
        team = str(r.get("Scrum team", "")).strip()
        if not team:
            continue
        pct = _automation_coverage_pct(r.get("Automation coverage"))
        if pct is None:
            continue
        if team not in team_best or pct > team_best[team]:
            team_best[team] = pct
    above = sum(1 for v in team_best.values() if v >= 70)
    below = sum(1 for v in team_best.values() if v < 70)
    rows = [["Element", "Calculation", "Value"],
            ["Sprints used (auto-detected)", "last 2 by numeric sort", ", ".join(sprints) or "—"],
            ["Unique teams classified", "—", str(len(team_best))],
            [">70 % In-Sprint bar", "teams whose best Automation coverage ≥ 70", str(above)],
            ["<70 % In-Sprint bar", "teams whose best Automation coverage < 70", str(below)]]
    flow.append(_table(rows, col_widths=[5.5 * cm, 7.5 * cm, 2.5 * cm]))
    return flow


def _section_slides_13_14(data, st):
    flow = [Paragraph("Slides 13–14 · Defect View", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>DefectData</i>. "
        "<b>Rule (per-Sub Domain mini-charts):</b> rows where "
        "Sprint ∈ last 2 sprint cycles AND <i>Type Of Defect = Manual</i> "
        "AND <i>Sub Domain ≠ TBD</i>. Per Sub Domain: Fatal = "
        "<i>InSprint Fatel + Regression Fatel</i> (Serious / Medium / Low "
        "computed the same way for their columns). Label = "
        "<i>&lt;Sub Domain&gt; (&lt;total&gt;)</i>. Title rewrites to "
        "<i>Defect VIEW – &lt;Sprint1&gt; - &lt;Sprint2&gt;</i>.", st["rule"]))
    sprints = _last_two_sprint_cycles(data)
    flow.append(Paragraph(f"<b>Sprints detected:</b> {', '.join(sprints) or '—'}", st["body"]))
    defs = _compute_manual_defects_by_subdomain(data, sprints)
    rows = [["Sub Domain", "Fatal", "Serious", "Medium", "Low", "Total"]]
    for sd in sorted(defs, key=lambda x: -sum(defs[x].values())):
        v = defs[sd]
        rows.append([sd, str(int(v["Fatal"])), str(int(v["Serious"])),
                     str(int(v["Medium"])), str(int(v["Low"])),
                     str(int(sum(v.values())))])
    flow.append(_table(rows, col_widths=[6 * cm, 1.6 * cm, 1.7 * cm, 1.7 * cm, 1.4 * cm, 1.6 * cm]))

    # Automation panel (slide 14)
    flow.append(Spacer(1, 6))
    flow.append(Paragraph("Slide 14 · 'Automation' panel (Rectangle 17 / Chart 18)", st["h2"]))
    flow.append(Paragraph(
        "<b>Rule:</b> rows where Sprint is a <i>monthly bucket</i> "
        "(NOT 'Sprint X.Y.Z') AND <i>Type Of Defect ∈ {Manual, Automation}</i> "
        "AND <i>Sub Domain ≠ TBD</i>. Sum <i>Prod Defect Count</i> per Sub Domain. "
        "Header label = <i>Automation (&lt;grand total&gt;)</i>.", st["rule"]))
    auto = _compute_automation_prod_defects(data)
    rows = [["Sub Domain", "Prod Defect Count"]]
    for sd in sorted(auto, key=lambda x: -auto[x]):
        if auto[sd] > 0:
            rows.append([sd, str(int(auto[sd]))])
    if len(rows) == 1:
        rows.append(["(no rows with Prod Defect Count > 0)", "0"])
    rows.append(["GRAND TOTAL", str(int(sum(auto.values())))])
    flow.append(_table(rows, col_widths=[8.5 * cm, 4 * cm]))
    return flow


def _section_slide_15(data, st):
    flow = [Paragraph("Slide 15 · Production & Post-Production Defect View", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>DefectData</i> (same row filter as Slide 14 Automation panel). "
        "<b>Chart 12 (Domain Wise Defects):</b> share of total Prod Defect Count "
        "per Sub Domain. <b>Chart 13 (QA Miss Distribution):</b> a row is "
        "classified as <i>QA Miss</i> iff its <i>Prod Defect RCA</i> contains "
        "the literal substring '<i>qa miss</i>' (case-insensitive). Distribution "
        "is over Prod Defect Count. <b>Title:</b> auto-built from month buckets "
        "present in the data.", st["rule"]))
    pd_ = _prod_defects_by_subdomain(data)
    grand = sum(pd_.values())
    rows = [["Sub Domain", "Prod Defects", "Share"]]
    for sd in sorted(pd_, key=lambda x: -pd_[x]):
        v = pd_[sd]; share = (v / grand) if grand else 0
        rows.append([sd, str(int(v)), f"{share:.3f}"])
    rows.append(["GRAND TOTAL", str(int(grand)), "1.000" if grand else "0.000"])
    flow.append(_table(rows, col_widths=[8 * cm, 3 * cm, 2.5 * cm]))
    flow.append(Spacer(1, 6))
    qa, not_qa = _qa_miss_split(data)
    tot = qa + not_qa
    rows2 = [["Bucket", "Prod Defects", "Share"],
             ["QA Miss", str(int(qa)), f"{(qa / tot if tot else 0):.3f}"],
             ["Not QA Miss", str(int(not_qa)), f"{(not_qa / tot if tot else 0):.3f}"]]
    flow.append(_table(rows2, col_widths=[8 * cm, 3 * cm, 2.5 * cm]))
    months = _monthly_buckets_in_data(data)
    flow.append(Paragraph(f"<b>Months in data:</b> {', '.join(months) or '—'}", st["body"]))
    return flow


def _section_slides_16_17(data, st):
    flow = [Paragraph("Slides 16–17 · RAD Snapshot", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>RADEnabled</i>. <b>Rule:</b> group rows by "
        "<i>Sub Domain</i> (skip TBD). Per panel: "
        "<i>Isolated Build</i> = sum(<i>Independent Build #</i>); "
        "<i>QA Pull Approved</i> = sum(<i># Builds QA Pull Enabled</i>); "
        "<i>RAD Enable</i> = sum(<i># Builds RAD Enabled</i>); "
        "<i>RAD Enablement In-Progress</i> = max(0, QA Pull − RAD Enabled). "
        "Label <i>~NN % RAD Enabled</i> = RAD Enabled / Independent Build × 100. "
        "<i>TMLB Muppets</i> is a special team-level panel (Sub Domain blank in source).",
        st["rule"]))
    rad = _compute_rad_by_subdomain(data)
    rows = [["Sub Domain", "Iso", "QA Pull", "RAD", "In-Progress", "% RAD"]]
    for sd in sorted(rad):
        v = rad[sd]
        pct = round(v["rad_enabled"] / v["isolated"] * 100) if v["isolated"] else 0
        rows.append([sd, str(int(v["isolated"])), str(int(v["qa_pull"])),
                     str(int(v["rad_enabled"])), str(int(v["in_progress"])),
                     f"{pct} %"])
    rt = _compute_rad_by_team(data)
    if "TMLB Muppets" in rt:
        v = rt["TMLB Muppets"]
        pct = round(v["rad_enabled"] / v["isolated"] * 100) if v["isolated"] else 0
        rows.append(["TMLB Muppets (by team)", str(int(v["isolated"])),
                     str(int(v["qa_pull"])), str(int(v["rad_enabled"])),
                     str(int(v["in_progress"])), f"{pct} %"])
    flow.append(_table(rows, col_widths=[6 * cm, 1.5 * cm, 2 * cm, 1.5 * cm, 2.5 * cm, 1.8 * cm]))
    return flow


def _section_slides_18_19(data, st):
    flow = [Paragraph("Slides 18–19 · Primary Release Automation", st["h1"])]
    flow.append(Paragraph(
        "<b>Source:</b> <i>ReleaseDayTestCaseSheet</i>. "
        "<b>Rule:</b> group rows by Sub Domain (skip TBD). "
        "Story Volume = sum(<i>No Of Stories part of the Release</i>); "
        "Automation Feasible TCs = sum(<i>No of Automation feasible TCs</i>); "
        "Automated TCs = sum(<i>No of TCs Automated</i>). "
        "Label <i>~NN.N% stories automated</i> = Automated / Feasible × 100.",
        st["rule"]))
    rel = _compute_release_by_subdomain(data)
    rows = [["Sub Domain", "Stories", "Feasible", "Automated", "%"]]
    for sd in sorted(rel):
        v = rel[sd]
        pct = round(v["automated"] / v["feasible"] * 100, 1) if v["feasible"] else 0
        pct_str = "100 %" if pct >= 99.5 else (f"{pct} %" if pct > 0 else "0 %")
        rows.append([sd, str(int(v["stories"])), str(int(v["feasible"])),
                     str(int(v["automated"])), pct_str])
    flow.append(_table(rows, col_widths=[6 * cm, 2 * cm, 2 * cm, 2.2 * cm, 2 * cm]))
    return flow


def _section_slides_20_21(data, st):
    flow = [Paragraph("Slides 20–21 · Regression Automation Trend (rolling 3-month)", st["h1"])]
    latest = _detect_latest_month_abbr(data) or "?"
    flow.append(Paragraph(
        "<b>Source:</b> <i>TCsDetailsSheet</i> (no Sub Domain column — looked up "
        "via Scrum Team Name → <i>ProjectDetails</i>). <b>Per Sub Domain:</b> "
        "Automatable = sum of <i>Total TCs Feasible P0..P4</i>; "
        "Automated = sum of <i>Total TCs Automated P0..P4</i>. "
        "<b>Chart presentation:</b> always exactly 3 month columns — 2 historical "
        "(preserved from template/prior generation) plus the latest month. "
        f"Latest month auto-detected from <i>Submitted On</i> dates: <b>{latest}</b>.",
        st["rule"]))
    trend = _compute_regression_by_subdomain(data)
    rows = [["Sub Domain", "Automatable", "Automated"]]
    for sd in sorted(trend):
        v = trend[sd]
        rows.append([sd, str(int(v["automatable"])), str(int(v["automated"]))])
    flow.append(_table(rows, col_widths=[6 * cm, 3 * cm, 3 * cm]))
    return flow


# ----------------------- Entry point -----------------------

def build_calculation_pdf(data, excel_name="", pptx_name=""):
    """Build the calculation report PDF and return its bytes."""
    st = _styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                            topMargin=1.5 * cm, bottomMargin=1.5 * cm,
                            title="QE Governance Deck — Calculation Report")
    story = [
        Paragraph("QE Governance Deck — Calculation Report", st["title"]),
        Paragraph(
            f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} · "
            f"Excel: <b>{excel_name or '(unknown)'}</b> · "
            f"PPT template: <b>{pptx_name or '(unknown)'}</b>",
            st["subtitle"]),
        Paragraph(
            "This document explains, slide by slide, exactly how each value in the "
            "generated PPT was computed from the source Excel — which sheet was "
            "read, what filter / aggregation was applied, and the per–Sub Domain "
            "(or panel) numbers that were written into the deck. "
            "Use it as an audit trail for every monthly run.",
            st["body"]),
        Spacer(1, 6),
    ]

    sections = [
        _section_slides_6_9,
        _section_slide_10,
        _section_slide_11,
        _section_slide_12,
        _section_slides_13_14,
        _section_slide_15,
        _section_slides_16_17,
        _section_slides_18_19,
        _section_slides_20_21,
    ]
    for fn in sections:
        try:
            story += fn(data, st)
        except Exception as e:
            story.append(Paragraph(
                f"<i>(Could not build this section: {type(e).__name__}: {e})</i>",
                st["body"]))
        story.append(Spacer(1, 8))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

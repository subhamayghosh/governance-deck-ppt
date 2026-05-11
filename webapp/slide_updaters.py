"""
Slide data updaters: populate PPT charts/tables from Excel data.

Automatable slides:
  11 - QE Highlights Execution (DailyExecution)
  12 - Automation Snapshot (InSprintData)
  13 - Defect View 1/2 (DefectData)
  14 - Defect View 2/2 (DefectData)
  16 - RAD Snapshot 1/2 (RADEnabled)
  17 - RAD Snapshot 2/2 (RADEnabled)
  18 - Primary Release Automation 1/2 (ReleaseDayTestCaseSheet)
  19 - Primary Release Automation 2/2 (ReleaseDayTestCaseSheet)
  20 - Regression Automation Trend 1/2 (TCsDetailsSheet)
  21 - Regression Automation Trend 2/2 (TCsDetailsSheet)
"""
from pptx.chart.data import CategoryChartData


def _num(v):
    try:
        return float(v) if v != "" else 0
    except (ValueError, TypeError):
        return 0


def update_chart_data(chart, categories, series_dict):
    """Update a chart with new category data. series_dict = {name: [values]}."""
    cd = CategoryChartData()
    cd.categories = categories
    for name, values in series_dict.items():
        cd.add_series(name, values)
    chart.replace_data(cd)


def _iter_shapes(shapes):
    """Recursively iterate all shapes, descending into groups."""
    for s in shapes:
        yield s
        # Group shapes have shape_type == 6 (MSO_SHAPE_TYPE.GROUP) and a .shapes attribute
        if getattr(s, "shape_type", None) == 6 and hasattr(s, "shapes"):
            yield from _iter_shapes(s.shapes)


def _get_charts_by_name(slide):
    """Return dict of shape.name -> shape for chart shapes."""
    return {s.name: s for s in slide.shapes if s.has_chart}


def _get_text_shapes(slide):
    """Return list of (shape_name, text, shape) for text shapes."""
    result = []
    for s in slide.shapes:
        if s.has_text_frame:
            result.append((s.name, s.text_frame.text.strip(), s))
    return result


def _set_text_preserve_format(shape, new_text):
    """Set text of a shape's first paragraph, preserving run formatting."""
    tf = shape.text_frame
    if tf.paragraphs and tf.paragraphs[0].runs:
        tf.paragraphs[0].runs[0].text = new_text
        for run in tf.paragraphs[0].runs[1:]:
            run.text = ""
    elif tf.paragraphs:
        tf.paragraphs[0].text = new_text


# ==============================================================================
# SLIDE 10: Cognizant QE Highlights (1/2) — # Scripts Distribution pie chart
# Source: MonthlySheet
#   - Filter rows where DashBoard Count > 0
#   - Sum Feature Branch Count, Develop Branch Count, Main Branch Count
# ==============================================================================
def update_slide_10(slide, data):
    """Update the #Scripts Distribution pie chart (Feature/Develop/Main) on slide 10."""
    monthly = data.get("monthly_sheet", [])
    if not monthly:
        return

    feature_sum = 0
    develop_sum = 0
    main_sum = 0
    python_total = 0
    for r in monthly:
        dash = _num(r.get("DashBoard Count", 0))
        f = _num(r.get("Feature Branch Count", 0))
        d = _num(r.get("Develop Branch Count", 0))
        m = _num(r.get("Main Branch Count", 0))
        tech = str(r.get("Technology Used", "")).strip().lower()
        # Pie chart + UI/API count: DashBoard Count > 0 AND Technology Used == Java
        if dash > 0 and tech == "java":
            feature_sum += f
            develop_sum += d
            main_sum += m
        # Python Scripts Created = sum of Feature+Develop+Main where Technology Used == Python
        if tech == "python":
            python_total += f + d + m

    total_scripts = feature_sum + develop_sum + main_sum

    for s in slide.shapes:
        if not s.has_chart:
            continue
        chart = s.chart
        try:
            cats = [str(c).strip() for c in chart.plots[0].categories]
        except Exception:
            continue
        # Match the Feature/Develop/Main pie chart by category labels
        cats_lower = [c.lower() for c in cats]
        if ("feature" in cats_lower and "develop" in cats_lower and "main" in cats_lower):
            # Preserve the existing series name from the template
            try:
                series_name = chart.plots[0].series[0].name or "Sales"
            except Exception:
                series_name = "Sales"
            # Align values to the chart's category order
            value_map = {"feature": feature_sum, "develop": develop_sum, "main": main_sum}
            values = [value_map.get(c, 0) for c in cats_lower]
            update_chart_data(chart, cats, {series_name: values})

    # Helper: find the numeric text box closest to a label text, and set its value.
    def _update_number_next_to_label(label_text, new_value):
        label_shape = None
        for s in slide.shapes:
            if s.has_text_frame and label_text in s.text_frame.text:
                label_shape = s
                break
        if label_shape is None:
            return None
        label_top = label_shape.top or 0
        label_left = label_shape.left or 0
        best = None
        best_dist = None
        for s in slide.shapes:
            if not s.has_text_frame or s is label_shape:
                continue
            txt = s.text_frame.text.strip()
            if not txt.isdigit():
                continue
            dx = (s.left or 0) - label_left
            dy = (s.top or 0) - label_top
            dist = dx * dx + dy * dy
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best = s
        if best is not None:
            _set_text_preserve_format(best, str(int(new_value)))
        return best

    # UI/API Scripts Created = Feature + Develop + Main totals from the pie chart
    _update_number_next_to_label("UI/API Scripts Created", total_scripts)
    # Python Scripts Created = sum of F+D+M where Technology Used == Python
    _update_number_next_to_label("Python Scripts Created", python_total)


# ==============================================================================
# SLIDE 11: QE Highlights - Execution (DailyExecution)
# ==============================================================================
def update_slide_11(slide, data):
    """Update daily execution charts and summary numbers."""
    daily = data.get("daily_execution", [])
    if not daily:
        return

    # P0-P4 totals across ALL rows (kept available for any chart that needs it)
    p_totals_all = {"P0": 0, "P1": 0, "P2": 0, "P3": 0, "P4": 0}
    # P0-P4 totals filtered by Daily/Weekly — used for Critical/Core/All and P0-P4 charts
    p_totals = {"P0": 0, "P1": 0, "P2": 0, "P3": 0, "P4": 0}
    total_passed = 0
    total_failed = 0
    # "Testcases configured in nightly batch" = sum of P0+P1+P2 where
    # Automation Execution Frequency is Daily or Weekly
    nightly_batch_total = 0
    # Pass rate = sum(P0-P2 Passed TCs) / sum(P0+P1+P2) across rows
    # where Automation Execution Frequency is Daily or Weekly
    p0_p2_total = 0
    p0_p2_passed = 0
    for r in daily:
        p0 = _num(r.get("P0", 0))
        p1 = _num(r.get("P1", 0))
        p2 = _num(r.get("P2", 0))
        p3 = _num(r.get("P3", 0))
        p4 = _num(r.get("P4", 0))
        p_totals_all["P0"] += p0
        p_totals_all["P1"] += p1
        p_totals_all["P2"] += p2
        p_totals_all["P3"] += p3
        p_totals_all["P4"] += p4
        total_passed += _num(r.get("Total Passed TC #", 0))
        total_failed += _num(r.get("Total Failed TC #", 0))
        freq = str(r.get("Automation Execution Frequency ",
                          r.get("Automation Execution Frequency", ""))).strip().lower()
        if freq in ("daily", "weekly"):
            p_totals["P0"] += p0
            p_totals["P1"] += p1
            p_totals["P2"] += p2
            p_totals["P3"] += p3
            p_totals["P4"] += p4
            nightly_batch_total += p0 + p1 + p2
            p0_p2_total += p0 + p1 + p2
            p0_p2_passed += _num(r.get("P0 - P2 Passed TCs", 0))

    # Critical/Core/All on the bar chart, filtered by Daily/Weekly
    total_configured = sum(p_totals.values())
    critical = p_totals["P0"] + p_totals["P1"]
    core = critical + p_totals["P2"]

    charts = _get_charts_by_name(slide)

    # Chart 26: Critical/Core/All bar chart
    for name, shape in charts.items():
        chart = shape.chart
        cats = [str(c).strip() for c in chart.plots[0].categories]
        if "Critical" in cats[0] if cats else False:
            update_chart_data(chart, ["Critical", "Core", "All"],
                              {"Daily Execution": [critical, core, total_configured]})
        elif "P0" in cats[0] if cats else False:
            update_chart_data(chart, ["P0", "P1", "P2", "P3", "P4"],
                              {"Daily Execution": [p_totals["P0"], p_totals["P1"], p_totals["P2"],
                                                   p_totals["P3"], p_totals["P4"]]})

    # Update "Testcases configured in nightly batch" number by finding the
    # integer text box closest to that label (recursively descends into groups).
    all_shapes = list(_iter_shapes(slide.shapes))
    label_shape = None
    for s in all_shapes:
        if s.has_text_frame and "nightly batch" in s.text_frame.text.lower():
            label_shape = s
            break
    if label_shape is not None:
        label_top = label_shape.top or 0
        label_left = label_shape.left or 0
        best = None
        best_dist = None
        for s in all_shapes:
            if not s.has_text_frame or s is label_shape:
                continue
            txt = s.text_frame.text.strip()
            if not txt.isdigit():
                continue
            dx = (s.left or 0) - label_left
            dy = (s.top or 0) - label_top
            dist = dx * dx + dy * dy
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best = s
        if best is not None:
            _set_text_preserve_format(best, str(int(nightly_batch_total)))

    # Pass rate = sum(P0-P2 Passed TCs) / sum(P0+P1+P2) * 100 (Daily/Weekly only)
    pass_rate = round(p0_p2_passed / p0_p2_total * 100) if p0_p2_total else 0
    # Find the "~NN%" text box (near "Pass rate" label) — descend into groups.
    pass_label = None
    for s in all_shapes:
        if s.has_text_frame and "Pass rate" in s.text_frame.text:
            pass_label = s
            break
    if pass_label is not None:
        label_top = pass_label.top or 0
        label_left = pass_label.left or 0
        best = None
        best_dist = None
        for s in all_shapes:
            if not s.has_text_frame or s is pass_label:
                continue
            txt = s.text_frame.text.strip()
            if "%" not in txt:
                continue
            dx = (s.left or 0) - label_left
            dy = (s.top or 0) - label_top
            dist = dx * dx + dy * dy
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best = s
        if best is not None:
            _set_text_preserve_format(best, f"~{pass_rate}%")


# ==============================================================================
# SLIDE 12: Automation Snapshot (InSprintData)
# ==============================================================================
def update_slide_12(slide, data):
    """Update automation snapshot bar chart and reasons table."""
    insprint = data.get("insprint_data", [])
    if not insprint:
        return

    # Categorize teams by automation health
    team_status = {}
    for r in insprint:
        team = str(r.get("Scrum team", "")).strip()
        if not team:
            continue
        sprint = str(r.get("Sprint", "")).strip()
        if "Sprint" not in sprint:
            continue

        cov = _num(r.get("Automation coverage", 0))
        designed = sum(_num(r.get(f"Test Case Designed {p}", 0)) for p in ["P0", "P1", "P2", "P3", "P4"])
        automated = sum(_num(r.get(f"Test case Automated {p}", 0)) for p in ["P0", "P1", "P2", "P3", "P4"])
        maintained = _num(r.get("TCs Maintained", 0))
        tech_debt = sum(_num(r.get(f"Auto tech debt {p}", 0)) for p in ["P0", "P1", "P2", "P3", "P4"])

        if designed == 0 and automated == 0:
            team_status[team] = "no_auto"
        elif maintained > 0 and designed == 0:
            team_status[team] = "maintenance"
        elif tech_debt > 0 and designed == 0:
            team_status[team] = "tech_debt"
        elif cov < 0.7:
            team_status[team] = "below_70"
        else:
            team_status[team] = "above_70"

    counts = {
        "No Automation": sum(1 for v in team_status.values() if v == "no_auto"),
        "Script Maintainance": sum(1 for v in team_status.values() if v == "maintenance"),
        "Tech-Debt Automation": sum(1 for v in team_status.values() if v == "tech_debt"),
        "<70 % In-Sprint": sum(1 for v in team_status.values() if v == "below_70"),
        ">70 % In-Sprint": sum(1 for v in team_status.values() if v == "above_70"),
    }

    charts = _get_charts_by_name(slide)
    for name, shape in charts.items():
        chart = shape.chart
        cats = [str(c).strip() for c in chart.plots[0].categories]
        if "No Automation" in cats[0] if cats else False:
            update_chart_data(chart, list(counts.keys()), {"Teams": list(counts.values())})


# ==============================================================================
# SLIDES 13-14: Defect View (DefectData)
# ==============================================================================

# Chart-to-domain mapping based on PPT structure
DEFECT_SLIDE_13_DOMAINS = {
    "Chart 5": "ALM-Technology",
    "Chart 7": "Corporate Systems (TFG)",
    "Chart 9": "Custody, Clearing & Settlement",
    "Chart 11": "Data/Platform Modernization",
    "Chart 13": "Infosec",
    "Chart 15": "Practice Management",
}
DEFECT_SLIDE_14_DOMAINS = {
    "Chart 5": "Service and Support",
    "Chart 7": "SRC",
    "Chart 9": "Technology",
    "Chart 11": "Trading",
}
# Label shape -> domain mapping for the text labels showing "Domain (count)"
DEFECT_LABEL_13 = {
    "Rectangle 4": "ALM-Technology",
    "Rectangle 6": "Corporate Systems (TFG)",
    "Rectangle 8": "Custody, Clearing & Settlement",
    "Rectangle 10": "Data/Platform Modernization",
    "Rectangle 12": "Infosec",
    "Rectangle 14": "Practice Management",
}
DEFECT_LABEL_14 = {
    "Rectangle 4": "Service and Support",
    "Rectangle 6": "SRC",
    "Rectangle 8": "Technology",
    "Rectangle 10": "Trading",
    "Rectangle 17": "Automation",
}


def _compute_defect_by_domain(data, defect_type=None):
    """Compute {domain: {Fatal, Serious, Medium, Low}} from DefectData."""
    result = {}
    for r in data.get("defect_data", []):
        if defect_type and str(r.get("Type Of Defect", "")) != defect_type:
            continue
        domain = str(r.get("Domain", "Unknown"))
        if domain not in result:
            result[domain] = {"Fatal": 0, "Serious": 0, "Medium": 0, "Low": 0}
        result[domain]["Fatal"] += _num(r.get("InSprint Fatel", 0)) + _num(r.get("Regression Fatel", 0))
        result[domain]["Serious"] += _num(r.get("InSprint Serious", 0)) + _num(r.get("Regression Serious", 0))
        result[domain]["Medium"] += _num(r.get("InSprint Medium", 0)) + _num(r.get("Regression Medium", 0))
        result[domain]["Low"] += _num(r.get("InSprint Low", 0)) + _num(r.get("Regression Low", 0))
    return result


def _domain_match(excel_domain, ppt_domain):
    """Fuzzy match domain names between Excel and PPT."""
    ed = excel_domain.lower().strip()
    pd = ppt_domain.lower().strip()
    if ed == pd:
        return True
    # Common mappings
    mappings = {
        "alm-technology": ["alm", "alm-technology", "account lifecycle management"],
        "corporate systems (tfg)": ["corporate systems", "corporate systems(tfg)", "corporate systems (tfg)"],
        "custody, clearing & settlement": ["custody", "custody, cleaning & settlement", "custody, clearing & settlement"],
        "data/platform modernization": ["data", "data/platform modernization"],
        "service and support": ["service and support", "service & support"],
        "advisor experience": ["advisor experience", "src"],
        "practice management": ["practice management"],
        "investor experience": ["investor experience"],
        "trading": ["trading"],
        "infosec": ["infosec"],
        "technology": ["technology"],
    }
    for key, variants in mappings.items():
        if pd in variants and ed in variants:
            return True
        if pd == key and ed in variants:
            return True
        if ed == key and pd in variants:
            return True
    return False


def _find_domain_data(domain_defects, ppt_domain):
    """Find matching domain data using fuzzy matching."""
    for ed, vals in domain_defects.items():
        if _domain_match(ed, ppt_domain):
            return vals
    return {"Fatal": 0, "Serious": 0, "Medium": 0, "Low": 0}


def update_defect_slide(slide, data, chart_domain_map, label_domain_map):
    """Update a defect view slide's charts and labels."""
    domain_defects = _compute_defect_by_domain(data)
    auto_defects = _compute_defect_by_domain(data, defect_type="Automation")

    charts = _get_charts_by_name(slide)
    for chart_name, ppt_domain in chart_domain_map.items():
        if chart_name not in charts:
            continue
        chart = charts[chart_name].chart
        vals = _find_domain_data(domain_defects, ppt_domain)
        update_chart_data(chart, ["Fatal", "Serious", "Medium", "Low"],
                          {"Defect Metrics": [vals["Fatal"], vals["Serious"], vals["Medium"], vals["Low"]]})

    # Handle automation defect chart on slide 14 (Chart 18)
    if "Chart 18" in charts:
        chart = charts["Chart 18"].chart
        auto_domains = ["SRC", "Service and Support", "Trading", "Corporate Systems (TFG)",
                        "Custody, Clearing & Settlement", "Data/Platform Modernization",
                        "Infosec", "Practice Management", "Technology", "ALM-Technology"]
        auto_vals = []
        for d in auto_domains:
            dv = _find_domain_data(auto_defects, d)
            total = dv["Fatal"] + dv["Serious"] + dv["Medium"] + dv["Low"]
            auto_vals.append(total if total else None)
        display_names = ["SRC", "Service and Support", "Trading", "Corporate Systems",
                         "Custody, clearing & settlement", "Data", "Infosec",
                         "Practice Management", "Technology", "ALM"]
        update_chart_data(chart, display_names, {"": auto_vals})

    # Update domain labels with counts
    for shape_name, text, shape in _get_text_shapes(slide):
        if shape_name in label_domain_map:
            ppt_domain = label_domain_map[shape_name]
            if ppt_domain == "Automation":
                total = sum(
                    sum(v.values()) for v in auto_defects.values()
                )
                _set_text_preserve_format(shape, f"Automation ({int(total)})")
            else:
                vals = _find_domain_data(domain_defects, ppt_domain)
                total = int(vals["Fatal"] + vals["Serious"] + vals["Medium"] + vals["Low"])
                # Extract display name from PPT domain
                display = ppt_domain.split("/")[0] if "/" in ppt_domain else ppt_domain
                # Shorten some names
                display = display.replace("Account Lifecycle Management", "ALM-Technology")
                _set_text_preserve_format(shape, f"{display} ({total})")


def update_slide_13(slide, data):
    update_defect_slide(slide, data, DEFECT_SLIDE_13_DOMAINS, DEFECT_LABEL_13)


def update_slide_14(slide, data):
    update_defect_slide(slide, data, DEFECT_SLIDE_14_DOMAINS, DEFECT_LABEL_14)


# ==============================================================================
# SLIDES 16-17: RAD Snapshot (RADEnabled)
# ==============================================================================

# Mapping: chart_name -> domain (based on spatial position in PPT)
RAD_SLIDE_16 = {
    "Chart 57": "ALM-Technology",          # top-right
    "Chart 44": "Advisor Experience",      # top-middle (SRC)
    "Chart 60": "Service and Support",     # top-left
    "Chart 4": "Practice Management",      # bottom-left
    "Chart 33": "Trading",                 # bottom-middle
    "Chart 6": "Corporate Systems (TFG)",  # bottom-right
}
RAD_LABEL_16 = {
    "TextBox 56": "ALM-Technology",
    "TextBox 43": "Advisor Experience",
    "TextBox 26": "Service and Support",
    "TextBox 5": "Practice Management",
    "TextBox 34": "Trading",
    "TextBox 7": "Corporate Systems (TFG)",
}
RAD_SLIDE_17 = {
    "Chart 60": "Custody, Clearing & Settlement",  # top-left
    "Chart 44": "Data/Platform Modernization",      # top-middle
    "Chart 57": "TMLB Muppets",                     # top-right (special team)
    "Chart 8": "Home Office",                        # bottom-left
    "Chart 10": "Infosec",                           # bottom-middle
    "Chart 12": "Technology",                        # bottom-right
}
RAD_LABEL_17 = {
    "TextBox 26": "Custody, Clearing & Settlement",
    "TextBox 43": "Data/Platform Modernization",
    "TextBox 56": "TMLB Muppets",
    "TextBox 2": "Home Office",
    "TextBox 9": "Infosec",
    "TextBox 11": "Technology",
}


def _compute_rad_by_domain(data):
    """Compute RAD data grouped by domain/sub-domain."""
    result = {}
    for r in data.get("rad_enabled", []):
        domain = str(r.get("Domain", "Unknown"))
        sub = str(r.get("Sub Domain", ""))
        key = sub if sub else domain
        if key not in result:
            result[key] = {"isolated": 0, "qa_pull": 0, "rad_enabled": 0, "in_progress": 0}
        result[key]["isolated"] += _num(r.get("Independent Build #", 0)) + _num(r.get("Shared Build #", 0))
        result[key]["qa_pull"] += _num(r.get("# Builds QA Pull Enabled", 0))
        result[key]["rad_enabled"] += _num(r.get("# Builds RAD Enabled", 0))
        elig = _num(r.get("RAD Eligible build from Independent Build", 0)) + _num(r.get("RAD Eligible build from Shared build", 0))
        result[key]["in_progress"] += max(0, _num(r.get("# Builds QA Pull Enabled", 0)) - _num(r.get("# Builds RAD Enabled", 0)))

    # Also aggregate by domain
    domain_result = {}
    for r in data.get("rad_enabled", []):
        domain = str(r.get("Domain", "Unknown"))
        if domain not in domain_result:
            domain_result[domain] = {"isolated": 0, "qa_pull": 0, "rad_enabled": 0, "in_progress": 0}
        domain_result[domain]["isolated"] += _num(r.get("Independent Build #", 0)) + _num(r.get("Shared Build #", 0))
        domain_result[domain]["qa_pull"] += _num(r.get("# Builds QA Pull Enabled", 0))
        domain_result[domain]["rad_enabled"] += _num(r.get("# Builds RAD Enabled", 0))
        domain_result[domain]["in_progress"] += max(0, _num(r.get("# Builds QA Pull Enabled", 0)) - _num(r.get("# Builds RAD Enabled", 0)))

    # Merge both
    domain_result.update(result)
    return domain_result


def _find_rad_data(rad_data, ppt_domain):
    for key, vals in rad_data.items():
        if _domain_match(key, ppt_domain) or ppt_domain.lower() in key.lower() or key.lower() in ppt_domain.lower():
            return vals
    return {"isolated": 0, "qa_pull": 0, "rad_enabled": 0, "in_progress": 0}


def update_rad_slide(slide, data, chart_map, label_map):
    rad_data = _compute_rad_by_domain(data)
    charts = _get_charts_by_name(slide)

    for chart_name, ppt_domain in chart_map.items():
        if chart_name not in charts:
            continue
        chart = charts[chart_name].chart
        vals = _find_rad_data(rad_data, ppt_domain)
        cats = ["Isolated Build", "QA Pull Approved", "RAD Enable ", "RAD Enablement In-Progress"]
        update_chart_data(chart, cats,
                          {"Series 1": [vals["isolated"], vals["qa_pull"], vals["rad_enabled"], vals["in_progress"]]})

    # Update labels with percentage
    for shape_name, text, shape in _get_text_shapes(slide):
        if shape_name in label_map:
            ppt_domain = label_map[shape_name]
            vals = _find_rad_data(rad_data, ppt_domain)
            pct = round(vals["rad_enabled"] / vals["isolated"] * 100) if vals["isolated"] else 0
            display = ppt_domain.replace("Advisor Experience", "SRC").replace("Data/Platform Modernization", "Data")
            _set_text_preserve_format(shape, f"{display} ~ {pct} % RAD Enabled")


def update_slide_16(slide, data):
    update_rad_slide(slide, data, RAD_SLIDE_16, RAD_LABEL_16)


def update_slide_17(slide, data):
    update_rad_slide(slide, data, RAD_SLIDE_17, RAD_LABEL_17)


# ==============================================================================
# SLIDES 18-19: Primary Release Automation (ReleaseDayTestCaseSheet)
# ==============================================================================

RELEASE_SLIDE_18 = {
    "Chart 4": "ALM-Technology",
    "Chart 10": "Corporate Systems (TFG)",
    "Chart 15": "Custody, Clearing & Settlement",
    "Chart 16": "Data/Platform Modernization",
    "Chart 18": "Infosec",
    "Chart 24": "Practice Management",
}
RELEASE_SLIDE_19 = {
    "Chart 4": "Service and Support",
    "Chart 10": "Advisor Experience",  # SRC
    "Chart 15": "Technology",
    "Chart 16": "Trading",
}


def _compute_release_by_domain(data):
    """Compute release data by domain, using latest month only."""
    result = {}
    for r in data.get("release_day", []):
        domain = str(r.get("Domain", "Unknown"))
        if domain not in result:
            result[domain] = {"stories": 0, "feasible": 0, "automated": 0}
        result[domain]["stories"] += _num(r.get("No Of Stories part of the Release", 0))
        result[domain]["feasible"] += _num(r.get("No of Automation feasible TCs", 0))
        result[domain]["automated"] += _num(r.get("No of TCs Automated", 0))
    return result


def _find_release_data(release_data, ppt_domain):
    for key, vals in release_data.items():
        if _domain_match(key, ppt_domain):
            return vals
    return {"stories": 0, "feasible": 0, "automated": 0}


def update_release_slide(slide, data, chart_map):
    release_data = _compute_release_by_domain(data)
    charts = _get_charts_by_name(slide)

    # Track label updates: TextBox 26-31 are the percentage labels
    label_updates = {}
    chart_positions = []

    for chart_name, ppt_domain in chart_map.items():
        if chart_name not in charts:
            continue
        chart_shape = charts[chart_name]
        chart = chart_shape.chart
        vals = _find_release_data(release_data, ppt_domain)
        cats = ["Story Volume", "Automation Feasible TCs", "Automated TCs"]
        update_chart_data(chart, cats,
                          {"Series 1": [vals["stories"], vals["feasible"], vals["automated"]]})

        # Compute percentage for label
        pct = round(vals["automated"] / vals["feasible"] * 100, 1) if vals["feasible"] else 0
        chart_positions.append((chart_shape.top, pct))

    # Update percentage text labels (sorted by position to match chart order)
    chart_positions.sort(key=lambda x: x[0])
    pct_labels = []
    for name, text, shape in _get_text_shapes(slide):
        if "automated" in text.lower() and "%" in text:
            pct_labels.append((shape.top, shape))

    pct_labels.sort(key=lambda x: x[0])
    for i, (_, pct) in enumerate(chart_positions):
        if i < len(pct_labels):
            _, shape = pct_labels[i]
            if pct == 0:
                _set_text_preserve_format(shape, "0% stories automated")
            elif pct >= 99.5:
                _set_text_preserve_format(shape, "100% stories automated")
            else:
                _set_text_preserve_format(shape, f"~ {pct}% stories automated")


def update_slide_18(slide, data):
    update_release_slide(slide, data, RELEASE_SLIDE_18)


def update_slide_19(slide, data):
    update_release_slide(slide, data, RELEASE_SLIDE_19)


# ==============================================================================
# SLIDES 20-21: Regression Automation Trend (TCsDetailsSheet)
# ==============================================================================

TREND_SLIDE_20 = {
    "Chart 20": "ALM-Technology",
    "Chart 15": "Corporate Systems (TFG)",
    "Chart 14": "Custody, Clearing & Settlement",
    "Chart 21": "Data/Platform Modernization",
    "Chart 22": "Infosec",
    "Chart 23": "Practice Management",
    "Chart 24": "Service and Support",
}
TREND_SLIDE_21 = {
    "Chart 12": "Advisor Experience",  # SRC
    "Chart 16": "Technology",
    "Chart 14": "Trading",
}


def _compute_regression_trend(data):
    """Compute automatable vs automated by domain from TCsDetailsSheet."""
    result = {}
    for r in data.get("tcs_details", []):
        domain = str(r.get("Domain", "Unknown"))
        if domain not in result:
            result[domain] = {"automatable": 0, "automated": 0}
        for p in ["P0", "P1", "P2", "P3", "P4"]:
            result[domain]["automatable"] += _num(r.get(f"Total TCs Feasible {p}", 0))
            result[domain]["automated"] += _num(r.get(f"Total TCs Automated {p}", 0))
    return result


def _find_trend_data(trend_data, ppt_domain):
    for key, vals in trend_data.items():
        if _domain_match(key, ppt_domain):
            return vals
    return {"automatable": 0, "automated": 0}


def update_trend_slide(slide, data, chart_map):
    trend_data = _compute_regression_trend(data)
    charts = _get_charts_by_name(slide)

    for chart_name, ppt_domain in chart_map.items():
        if chart_name not in charts:
            continue
        chart = charts[chart_name].chart
        vals = _find_trend_data(trend_data, ppt_domain)

        # The trend charts have JAN/FEB months - we put current data in the latest month
        # Keep existing first month data, update second month
        existing_cats = [str(c) for c in chart.plots[0].categories]
        if len(existing_cats) >= 2:
            # Try to read existing first-month data
            try:
                old_automatable = list(chart.plots[0].series[0].values)[0]
                old_automated = list(chart.plots[0].series[1].values)[0]
                if old_automatable is None:
                    old_automatable = 0
                if old_automated is None:
                    old_automated = 0
            except (IndexError, TypeError):
                old_automatable = 0
                old_automated = 0

            update_chart_data(chart, existing_cats, {
                "Automatable": [old_automatable, vals["automatable"]],
                "Automated": [old_automated, vals["automated"]],
            })
        else:
            update_chart_data(chart, ["Current"], {
                "Automatable": [vals["automatable"]],
                "Automated": [vals["automated"]],
            })


def update_slide_20(slide, data):
    update_trend_slide(slide, data, TREND_SLIDE_20)


def update_slide_21(slide, data):
    update_trend_slide(slide, data, TREND_SLIDE_21)


# ==============================================================================
# SLIDES 6-9: Presence of Cognizant QE Team (ProjectDetails)
# ==============================================================================

# Pentagon shape → TextBox shape name (by vertical row position)
# Pentagon 11 (row 1) ↔ TextBox 21
# Pentagon 12 (row 2) ↔ TextBox 17
# Pentagon 13 (row 3) ↔ TextBox 19

PRESENCE_SLIDE_DOMAINS = {
    6: [
        ("Arrow: Pentagon 11", "TextBox 21", "Chart 10", "ALM-Technology"),
        ("Arrow: Pentagon 12", "TextBox 17", "Chart 9", "Corporate Systems(TFG)"),
        ("Arrow: Pentagon 13", "TextBox 19", "Chart 6", "Custody, Clearing & Settlement"),
    ],
    7: [
        ("Arrow: Pentagon 11", "TextBox 21", "Chart 14", "Data"),
        ("Arrow: Pentagon 12", "TextBox 17", "Chart 10", "Infosec"),
        ("Arrow: Pentagon 13", "TextBox 19", "Chart 9", "Practice Management"),
    ],
    8: [
        ("Arrow: Pentagon 11", "TextBox 21", "Chart 10", "Service and Support"),
        ("Arrow: Pentagon 12", "TextBox 17", "Chart 9", "SRC"),
        ("Arrow: Pentagon 13", "TextBox 19", "Chart 6", "Technology"),
    ],
    9: [
        (None, "TextBox 21", "Chart 9", "Trading"),
    ],
}

# Sub-domain display labels (how they appear in the PPT pentagon)
SUBDOMAIN_DISPLAY = {
    "ALM-Technology": "ALM - Technology",
    "Corporate Systems(TFG)": "Corporate Systems (TFG)",
    "Custody, Clearing & Settlement": "Custody, Clearing & Settlement",
    "Data": "Data",
    "Infosec": "Infosec",
    "Practice Management": "Practice Management",
    "Service and Support": "Service and Support",
    "SRC": "SRC",
    "Technology": "Technology",
    "Trading": "Trading",
}


def _build_team_lookup(data):
    """Build {sub_domain: {'SVT': [teams], 'Non SVT': [teams]}} from ProjectDetails."""
    lookup = {}
    for r in data.get("project_details", []):
        sub = str(r.get("Sub Domain", "")).strip()
        team = str(r.get("Scrum Team Name", "")).strip()
        ttype = str(r.get("Team Type", "")).strip()
        if not team:
            continue
        # Skip rows where Sub Domain is blank / TBD — they should not be
        # grouped into any sub-domain chart or text box.
        if not sub or sub in ("TBD", "nan", "None"):
            continue
        key = sub
        if key not in lookup:
            lookup[key] = {"SVT": [], "Non SVT": []}
        if ttype == "SVT":
            lookup[key]["SVT"].append(team)
        elif ttype == "Non SVT":
            lookup[key]["Non SVT"].append(team)
        # Skip TBD or any other Team Type (don't count in charts/text boxes)
    return lookup


def _format_teams_text(svt_teams, nonsvt_teams):
    """Format team list into readable text for the slide text box."""
    lines = []
    if svt_teams:
        lines.append("SVT Teams:")
        for t in svt_teams:
            lines.append(f"  \u2022 {t}")
    if nonsvt_teams:
        if lines:
            lines.append("")
        lines.append("Non SVT Teams:")
        for t in nonsvt_teams:
            lines.append(f"  \u2022 {t}")
    return "\n".join(lines) if lines else "No teams assigned"


def _capture_run_format(run):
    """Capture font formatting from a run to replicate later."""
    font = run.font
    fmt = {
        "name": font.name,
        "size": font.size,
        "bold": font.bold,
        "italic": font.italic,
        "underline": font.underline,
        "color_rgb": None,
        "color_theme": None,
    }
    try:
        if font.color and font.color.type is not None:
            try:
                fmt["color_rgb"] = font.color.rgb
            except AttributeError:
                pass
            try:
                fmt["color_theme"] = font.color.theme_color
            except AttributeError:
                pass
    except Exception:
        pass
    return fmt


def _apply_run_format(run, fmt):
    """Apply previously captured formatting to a run."""
    font = run.font
    if fmt.get("name"):
        font.name = fmt["name"]
    if fmt.get("size"):
        font.size = fmt["size"]
    if fmt.get("bold") is not None:
        font.bold = fmt["bold"]
    if fmt.get("italic") is not None:
        font.italic = fmt["italic"]
    if fmt.get("underline") is not None:
        font.underline = fmt["underline"]
    if fmt.get("color_rgb"):
        try:
            font.color.rgb = fmt["color_rgb"]
        except Exception:
            pass


def _replace_textframe_content(shape, lines, header_fmt=None, body_fmt=None):
    """Replace the text frame content with given lines preserving formatting.

    lines: list of (text, is_header) tuples where is_header=True gets header_fmt.
    If header_fmt/body_fmt are None, they are captured from the existing content.
    """
    tf = shape.text_frame

    # Capture existing formatting from first paragraph's first run if not provided
    if header_fmt is None or body_fmt is None:
        captured = None
        for p in tf.paragraphs:
            if p.runs:
                captured = _capture_run_format(p.runs[0])
                break
        if header_fmt is None:
            header_fmt = dict(captured) if captured else {}
            header_fmt["bold"] = True
        if body_fmt is None:
            body_fmt = dict(captured) if captured else {}
            body_fmt["bold"] = False  # Body team names should not be bold

    # Remove all paragraphs except the first (by XML manipulation)
    from pptx.oxml.ns import qn
    txBody = tf._txBody
    paragraphs = txBody.findall(qn("a:p"))
    for p in paragraphs[1:]:
        txBody.remove(p)

    # Clear the first paragraph's content but keep its properties
    first_p = tf.paragraphs[0]
    for run in list(first_p.runs):
        run._r.getparent().remove(run._r)

    # Add lines back
    first = True
    for line_text, is_header in lines:
        if first:
            p = first_p
            first = False
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line_text
        _apply_run_format(run, header_fmt if is_header else body_fmt)


def _update_presence_slide(slide, data, slide_num):
    """Update domain pentagon labels and SVT/Non SVT count charts for slides 6-9.

    NOTE: The 'Teams and Key Feature Implemented' text boxes (TextBox 17/19/21) are
    intentionally NOT updated — their feature description content is static and
    maintained manually by the author.
    """
    team_lookup = _build_team_lookup(data)
    domain_rows = PRESENCE_SLIDE_DOMAINS.get(slide_num, [])

    # Build shape name → shape map
    shape_map = {s.name: s for s in slide.shapes}

    for pentagon_name, textbox_name, chart_name, sub_domain in domain_rows:
        # 1. Update pentagon domain label
        if pentagon_name and pentagon_name in shape_map:
            display = SUBDOMAIN_DISPLAY.get(sub_domain, sub_domain)
            _set_text_preserve_format(shape_map[pentagon_name], display)

        # Look up teams for this sub-domain (with fuzzy fallback)
        teams = team_lookup.get(sub_domain, {"SVT": [], "Non SVT": []})
        if not teams["SVT"] and not teams["Non SVT"]:
            for key in team_lookup:
                if sub_domain.lower() in key.lower() or key.lower() in sub_domain.lower():
                    teams = team_lookup[key]
                    break

        # 2. Update SVT/Non SVT count chart
        if chart_name and chart_name in shape_map and shape_map[chart_name].has_chart:
            chart = shape_map[chart_name].chart
            svt_count = len(teams["SVT"])
            nonsvt_count = len(teams["Non SVT"])
            # Preserve existing category (e.g. year "2026") from the template
            try:
                existing_cats = [str(c) for c in chart.plots[0].categories]
                category = existing_cats[0] if existing_cats else "2026"
            except Exception:
                category = "2026"
            update_chart_data(chart, [category], {
                "SVT": [svt_count if svt_count > 0 else None],
                "Non SVT": [nonsvt_count if nonsvt_count > 0 else None],
            })


def update_slide_6(slide, data):
    _update_presence_slide(slide, data, 6)


def update_slide_7(slide, data):
    _update_presence_slide(slide, data, 7)


def update_slide_8(slide, data):
    _update_presence_slide(slide, data, 8)


def update_slide_9(slide, data):
    _update_presence_slide(slide, data, 9)


# ==============================================================================
# MAIN: Apply all updates to a presentation
# ==============================================================================

SLIDE_UPDATERS = {
    6: update_slide_6,
    7: update_slide_7,
    8: update_slide_8,
    9: update_slide_9,
    10: update_slide_10,
    11: update_slide_11,
    12: update_slide_12,
    13: update_slide_13,
    14: update_slide_14,
    16: update_slide_16,
    17: update_slide_17,
    18: update_slide_18,
    19: update_slide_19,
    20: update_slide_20,
    21: update_slide_21,
}


def update_presentation(prs, data):
    """Apply all data updates to the presentation slides."""
    for i, slide in enumerate(prs.slides):
        slide_num = i + 1
        if slide_num in SLIDE_UPDATERS:
            try:
                SLIDE_UPDATERS[slide_num](slide, data)
            except Exception as e:
                print(f"Warning: Failed to update slide {slide_num}: {e}")

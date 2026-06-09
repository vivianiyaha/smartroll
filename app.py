"""
Smart Payroll Payslip Generator
Main Streamlit Application — Multi-File Edition
"""

import streamlit as st
import pandas as pd
import io
import zipfile
import base64
import time
from datetime import datetime

from modules.data_handler import (
    load_payroll_data, validate_columns, compute_summaries,
    validate_single_file, merge_payroll_files,
    EARNING_COLS, DEDUCTION_COLS,
)
from modules.pdf_generator import generate_payslip_pdf
from modules.settings import init_settings, render_settings_page, get_settings
from modules.styles import inject_css

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart Payroll Payslip Generator",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_css()
init_settings()


# ── Sidebar ───────────────────────────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        settings = get_settings()
        logo_b64 = settings.get("logo_b64")

        if logo_b64:
            st.markdown(
                f'<div style="text-align:center;margin-bottom:10px;">'
                f'<img src="data:image/png;base64,{logo_b64}" style="max-width:130px;border-radius:8px;"></div>',
                unsafe_allow_html=True,
            )

        st.markdown(
            f'<div class="sidebar-company">{settings["company_name"]}</div>',
            unsafe_allow_html=True,
        )
        st.markdown("---")

        menu = st.radio(
            "Navigation",
            [
                "📊 Dashboard",
                "📥 Multi-File Upload & Merge",
                "🔍 Employee Search",
                "📄 Generate Single Payslip",
                "📦 Generate Bulk Payslips",
                "💾 Download ZIP Archive",
                "📋 Merged Master Report",
                "🔎 Payroll Audit Log",
                "⚙️ Settings",
            ],
            label_visibility="collapsed",
        )
        st.markdown("---")
        st.markdown(
            '<div class="sidebar-footer">© 2025 Smart Payroll</div>',
            unsafe_allow_html=True,
        )
    return menu


# ── Single-file upload widget ─────────────────────────────────────────────────
def upload_section():
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload Payroll Excel File (.xlsx)",
        type=["xlsx"],
        help="Upload your payroll spreadsheet.",
    )
    st.markdown("</div>", unsafe_allow_html=True)
    return uploaded


# ── Get active DataFrame (merged or single) ───────────────────────────────────
def get_active_df():
    """Return merged DF if available, else None."""
    return st.session_state.get("merged_df")


# ── Dashboard ─────────────────────────────────────────────────────────────────
def page_dashboard(df):
    settings = get_settings()
    logo_b64 = settings.get("logo_b64")

    st.markdown('<div class="page-header">', unsafe_allow_html=True)
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        if logo_b64:
            st.markdown(
                f'<img src="data:image/png;base64,{logo_b64}" style="max-width:90px;border-radius:8px;margin-top:4px;">',
                unsafe_allow_html=True,
            )
    with col_title:
        st.markdown(f"## {settings['company_name']}")
        st.markdown(f"**Payroll Period:** {settings['payroll_month']} {settings['payroll_year']}")
    st.markdown("</div>", unsafe_allow_html=True)

    if df is None:
        st.info("📂 Please use **Multi-File Upload & Merge** or upload a single payroll Excel file to get started.")
        return

    summary = compute_summaries(df)
    source_files = st.session_state.get("source_file_names", [])

    # Workflow status banner
    steps = {
        "✓ Files Uploaded": len(source_files) > 0,
        "✓ Records Processed": df is not None,
        "✓ Employees Merged": st.session_state.get("merge_done", False),
        "✓ Payslips Generated": st.session_state.get("zip_ready", False),
    }
    cols = st.columns(4)
    for col, (label, done) in zip(cols, steps.items()):
        colour = "#28a745" if done else "#888888"
        col.markdown(
            f'<div class="metric-card" style="border-left-color:{colour};">'
            f'<div class="metric-value" style="font-size:14px;color:{colour};">{label}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    c1, c2, c3, c4 = st.columns(4)
    for col, label, value in [
        (c1, "👥 Total Employees", str(summary["total_employees"])),
        (c2, "💰 Total Payroll", f"₦{summary['total_gross']:,.2f}"),
        (c3, "📉 Total Deductions", f"₦{summary['total_deductions']:,.2f}"),
        (c4, "✅ Total Net Salary", f"₦{summary['total_net']:,.2f}"),
    ]:
        col.markdown(
            f'<div class="metric-card"><div class="metric-label">{label}</div>'
            f'<div class="metric-value">{value}</div></div>',
            unsafe_allow_html=True,
        )

    if source_files:
        st.markdown(f"**Source files loaded:** {', '.join(source_files)}")

    st.markdown("### 📋 Payroll Summary Table")
    display_cols = ["STAFF NAME", "BASIC", "HOUSING", "TRANSPORT", "TOTAL DED.", "NET SALARY"]
    available = [c for c in display_cols if c in df.columns]
    st.dataframe(df[available].reset_index(drop=True), use_container_width=True)


# ── Multi-file upload & merge ─────────────────────────────────────────────────
def page_multi_upload():
    st.markdown("## 📥 Multi-File Upload & Merge")
    st.markdown(
        "Upload **two or more** Excel payroll files. The system will automatically merge records "
        "by employee name (case-insensitive, whitespace-tolerant)."
    )

    # ── Duplicate treatment setting
    dup_treatment = st.radio(
        "🔧 Duplicate Column Treatment",
        ["Sum Values", "Use First File Value", "Use Last File Value"],
        horizontal=True,
        help="What to do when the same column (e.g. BASIC) appears in multiple files.",
    )
    dup_map = {"Sum Values": "sum", "Use First File Value": "first", "Use Last File Value": "last"}

    # ── File uploader
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Upload one or more Payroll Excel Files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="You may upload 20+ files. Each must contain a STAFF NAME column.",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    if not uploaded_files:
        st.info("⬆️ Upload at least one .xlsx file above to begin.")
        return

    # ── Show uploaded files with remove option
    st.markdown(f"**{len(uploaded_files)} file(s) uploaded:**")
    keep_files = []
    for f in uploaded_files:
        col_name, col_btn = st.columns([6, 1])
        col_name.markdown(f"📄 `{f.name}`")
        keep_files.append(f)

    st.markdown("---")

    # ── Load and validate each file
    if st.button("🔍 Validate & Preview Files", type="primary"):
        st.session_state["upload_files_data"] = []
        st.session_state["upload_validation"] = []
        all_ok = True

        progress_bar = st.progress(0, text="Loading files…")
        for i, f in enumerate(keep_files):
            try:
                df_raw = load_payroll_data(f)
                f.seek(0)
                val = validate_single_file(df_raw, f.name)
                st.session_state["upload_files_data"].append({"name": f.name, "df": df_raw})
                st.session_state["upload_validation"].append(val)
                if not val["valid"] or any("🔴" in iss for iss in val["issues"]):
                    all_ok = False
            except Exception as exc:
                st.session_state["upload_validation"].append({
                    "filename": f.name, "row_count": 0, "issues": [f"🔴 Could not read: {exc}"], "valid": False
                })
                all_ok = False
            progress_bar.progress(int((i + 1) / len(keep_files) * 100))

        progress_bar.empty()
        st.session_state["upload_validated"] = True
        st.session_state["upload_all_ok"] = all_ok

    # ── Show validation results
    if st.session_state.get("upload_validated"):
        st.markdown("### 🔎 Validation Report")
        for val in st.session_state.get("upload_validation", []):
            with st.expander(
                f"{'✅' if val['valid'] and not val['issues'] else '⚠️'} {val['filename']} — {val.get('row_count', 0)} rows",
                expanded=bool(val["issues"]),
            ):
                if val["issues"]:
                    for issue in val["issues"]:
                        st.markdown(issue)
                else:
                    st.markdown("✅ No issues found.")
                if val.get("columns"):
                    st.markdown(f"**Columns detected:** `{'`, `'.join(val['columns'])}`")

        # ── Merge button
        st.markdown("---")
        st.markdown("### 🔀 Merge Payroll Files")
        n_employees_est = sum(v.get("row_count", 0) for v in st.session_state.get("upload_validation", []))
        st.info(f"Ready to merge **{len(st.session_state.get('upload_files_data', []))} file(s)** with ~**{n_employees_est}** total records.")

        if st.button("🚀 Merge All Files", type="primary"):
            progress_bar = st.progress(0, text="Starting merge…")
            status_text = st.empty()

            def _progress(pct, msg):
                progress_bar.progress(pct, text=msg)
                status_text.markdown(f"⏳ {msg}")

            merged_df, audit_df = merge_payroll_files(
                st.session_state["upload_files_data"],
                duplicate_treatment=dup_map[dup_treatment],
                progress_callback=_progress,
            )

            progress_bar.empty()
            status_text.empty()

            st.session_state["merged_df"] = merged_df
            st.session_state["audit_df"] = audit_df
            st.session_state["merge_done"] = True
            st.session_state["source_file_names"] = [f["name"] for f in st.session_state["upload_files_data"]]
            st.session_state["zip_ready"] = False  # reset zip on new merge

            st.success(f"✅ Merge complete! **{len(merged_df)}** unique employees consolidated.")

    # ── Consolidated preview
    if st.session_state.get("merge_done") and st.session_state.get("merged_df") is not None:
        merged_df = st.session_state["merged_df"]
        st.markdown("### 👁️ Consolidated Payroll Preview")

        preview_cols = ["STAFF NAME"]
        earn_present = [c for c in EARNING_COLS if c in merged_df.columns]
        ded_present = [c for c in DEDUCTION_COLS if c in merged_df.columns]
        if earn_present:
            preview_cols += earn_present
        if "_GROSS" in merged_df.columns:
            preview_cols.append("_GROSS")
        if ded_present:
            preview_cols += ded_present
        if "TOTAL DED." in merged_df.columns:
            preview_cols.append("TOTAL DED.")
        if "NET SALARY" in merged_df.columns:
            preview_cols.append("NET SALARY")

        available = [c for c in preview_cols if c in merged_df.columns]
        display_df = merged_df[available].copy()
        if "_GROSS" in display_df.columns:
            display_df = display_df.rename(columns={"_GROSS": "GROSS SALARY"})

        st.dataframe(display_df.reset_index(drop=True), use_container_width=True)

        summary = compute_summaries(merged_df)
        c1, c2, c3, c4 = st.columns(4)
        for col, label, value in [
            (c1, "👥 Employees", str(summary["total_employees"])),
            (c2, "💰 Total Gross", f"₦{summary['total_gross']:,.2f}"),
            (c3, "📉 Total Deductions", f"₦{summary['total_deductions']:,.2f}"),
            (c4, "✅ Total Net", f"₦{summary['total_net']:,.2f}"),
        ]:
            col.markdown(
                f'<div class="metric-card"><div class="metric-label">{label}</div>'
                f'<div class="metric-value">{value}</div></div>',
                unsafe_allow_html=True,
            )

        st.success("✅ Data is ready. Use the sidebar to generate payslips, reports, or the audit log.")


# ── Employee search ───────────────────────────────────────────────────────────
def page_employee_search(df):
    st.markdown("## 🔍 Employee Search")
    if df is None:
        st.info("📂 Please upload and merge payroll files first.")
        return

    query = st.text_input("Search by employee name", placeholder="Type a name…")
    if query:
        mask = df["STAFF NAME"].str.contains(query, case=False, na=False)
        results = df[mask]
        if results.empty:
            st.warning("No employees matched your search.")
        else:
            st.success(f"Found {len(results)} employee(s)")
            st.dataframe(results.reset_index(drop=True), use_container_width=True)
    else:
        st.dataframe(df.reset_index(drop=True), use_container_width=True)


# ── Single payslip ────────────────────────────────────────────────────────────
def page_single_payslip(df):
    st.markdown("## 📄 Generate Single Payslip")
    if df is None:
        st.info("📂 Please upload and merge payroll files first.")
        return

    settings = get_settings()
    names = df["STAFF NAME"].dropna().tolist()
    selected = st.selectbox("Select Employee", options=names)

    if selected:
        emp = df[df["STAFF NAME"] == selected].iloc[0]

        st.markdown("### 👤 Employee Preview")
        col_earn, col_ded = st.columns(2)

        with col_earn:
            st.markdown("**Earnings**")
            basic = float(emp.get("BASIC", 0) or 0)
            housing = float(emp.get("HOUSING", 0) or 0)
            transport = float(emp.get("TRANSPORT", 0) or 0)
            gross = basic + housing + transport

            # Detect extra earning columns
            extra_earn_cols = [
                c for c in df.columns
                if c not in EARNING_COLS + DEDUCTION_COLS + ["STAFF NAME", "_GROSS", "TOTAL DED.", "NET SALARY"]
                and pd.api.types.is_numeric_dtype(df[c])
            ]

            earn_items = [
                ("Basic Salary", basic),
                ("Housing Allowance", housing),
                ("Transport Allowance", transport),
            ]
            for col in extra_earn_cols:
                earn_items.append((col, float(emp.get(col, 0) or 0)))
                gross += float(emp.get(col, 0) or 0)

            earn_items.append(("**Gross Salary**", gross))
            earn_data = {
                "Item": [i[0] for i in earn_items],
                "Amount (₦)": [f"**{i[1]:,.2f}**" if i[0].startswith("**") else f"{i[1]:,.2f}" for i in earn_items],
            }
            st.table(pd.DataFrame(earn_data))

        with col_ded:
            st.markdown("**Deductions**")
            tax = float(emp.get("TAX", 0) or 0)
            pension = float(emp.get("PENSION", 0) or 0)
            loan = float(emp.get("LOAN", 0) or 0)
            sal_adv = float(emp.get("SAL. ADV.", 0) or 0)
            penalty = float(emp.get("PENALTY", 0) or 0)
            total_ded = float(emp.get("TOTAL DED.", tax + pension + loan + sal_adv + penalty) or (tax + pension + loan + sal_adv + penalty))
            net = float(emp.get("NET SALARY", gross - total_ded) or (gross - total_ded))

            ded_data = {
                "Item": ["Tax", "Pension", "Loan", "Salary Advance", "Penalty", "**Total Deductions**"],
                "Amount (₦)": [
                    f"{tax:,.2f}", f"{pension:,.2f}", f"{loan:,.2f}",
                    f"{sal_adv:,.2f}", f"{penalty:,.2f}", f"**{total_ded:,.2f}**",
                ],
            }
            st.table(pd.DataFrame(ded_data))

        st.markdown(
            f'<div class="net-salary-box">💵 NET SALARY: ₦{net:,.2f}</div>',
            unsafe_allow_html=True,
        )

        if st.button("🖨️ Generate & Download PDF Payslip", type="primary"):
            with st.spinner("Generating PDF…"):
                pdf_bytes = generate_payslip_pdf(emp, settings)
            fname = f"payslip_{selected.replace(' ', '_')}_{settings['payroll_month']}_{settings['payroll_year']}.pdf"
            st.download_button(
                label="⬇️ Download Payslip PDF",
                data=pdf_bytes,
                file_name=fname,
                mime="application/pdf",
            )
            st.success("✅ Payslip generated successfully!")


# ── Bulk payslips ─────────────────────────────────────────────────────────────
def page_bulk_payslips(df):
    st.markdown("## 📦 Generate Bulk Payslips")
    if df is None:
        st.info("📂 Please upload and merge payroll files first.")
        return

    settings = get_settings()
    st.info(f"Ready to generate payslips for **{len(df)}** employee(s).")

    if st.button("🚀 Generate All Payslips", type="primary"):
        progress = st.progress(0, text="Generating payslips…")
        zip_buf = io.BytesIO()

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, (_, row) in enumerate(df.iterrows(), 1):
                pdf_bytes = generate_payslip_pdf(row, settings)
                fname = f"payslip_{str(row['STAFF NAME']).replace(' ', '_')}_{settings['payroll_month']}_{settings['payroll_year']}.pdf"
                zf.writestr(fname, pdf_bytes)
                progress.progress(i / len(df), text=f"Processing {row['STAFF NAME']}…")

        zip_buf.seek(0)
        st.session_state["zip_data"] = zip_buf.getvalue()
        st.session_state["zip_ready"] = True
        progress.empty()
        st.success(f"✅ {len(df)} payslips generated!")

    if st.session_state.get("zip_ready"):
        st.download_button(
            label="⬇️ Download All Payslips (ZIP)",
            data=st.session_state["zip_data"],
            file_name=f"payslips_{settings['payroll_month']}_{settings['payroll_year']}.zip",
            mime="application/zip",
        )


# ── Download ZIP ──────────────────────────────────────────────────────────────
def page_download_zip(df):
    st.markdown("## 💾 Download ZIP Archive")
    if df is None:
        st.info("📂 Please upload and merge payroll files first.")
        return

    if st.session_state.get("zip_ready"):
        settings = get_settings()
        st.success("✅ ZIP archive is ready for download.")
        st.download_button(
            label="⬇️ Download ZIP Archive",
            data=st.session_state["zip_data"],
            file_name=f"payslips_{settings['payroll_month']}_{settings['payroll_year']}.zip",
            mime="application/zip",
        )
    else:
        st.warning("⚠️ No ZIP archive generated yet. Go to **Generate Bulk Payslips** first.")


# ── Merged master report ──────────────────────────────────────────────────────
def page_master_report(df):
    st.markdown("## 📋 Merged Master Report")
    if df is None:
        st.info("📂 Please upload and merge payroll files first.")
        return

    settings = get_settings()

    # Build display df
    display_df = df.copy()
    if "_GROSS" in display_df.columns:
        display_df = display_df.rename(columns={"_GROSS": "GROSS SALARY"})
    internal = [c for c in ["_GROSS"] if c in display_df.columns]
    display_df = display_df.drop(columns=internal, errors="ignore")

    st.dataframe(display_df.reset_index(drop=True), use_container_width=True)

    # Summary row
    summary = compute_summaries(df)
    st.markdown(
        f'<div class="net-salary-box" style="text-align:left;font-size:15px;">'
        f'👥 {summary["total_employees"]} employees &nbsp;|&nbsp; '
        f'💰 Gross: ₦{summary["total_gross"]:,.2f} &nbsp;|&nbsp; '
        f'📉 Deductions: ₦{summary["total_deductions"]:,.2f} &nbsp;|&nbsp; '
        f'✅ Net: ₦{summary["total_net"]:,.2f}'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown("### 📥 Export Report")
    col_excel, col_csv, col_pdf = st.columns(3)

    with col_excel:
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            display_df.to_excel(writer, index=False, sheet_name="Payroll Report")
        excel_buf.seek(0)
        st.download_button(
            "⬇️ Download Excel",
            data=excel_buf.getvalue(),
            file_name=f"payroll_report_{settings['payroll_month']}_{settings['payroll_year']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col_csv:
        csv_bytes = display_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download CSV",
            data=csv_bytes,
            file_name=f"payroll_report_{settings['payroll_month']}_{settings['payroll_year']}.csv",
            mime="text/csv",
        )

    with col_pdf:
        if st.button("📄 Generate PDF Report", type="primary"):
            with st.spinner("Generating PDF report…"):
                pdf_bytes = _generate_report_pdf(display_df, settings, summary)
            st.download_button(
                "⬇️ Download PDF Report",
                data=pdf_bytes,
                file_name=f"payroll_report_{settings['payroll_month']}_{settings['payroll_year']}.pdf",
                mime="application/pdf",
            )


def _generate_report_pdf(df, settings, summary):
    """Generate a summary PDF table of the entire payroll."""
    import io
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    ORANGE = colors.HexColor("#E8610A")
    DARK_GRAY = colors.HexColor("#2C2C2C")
    LIGHT_GRAY = colors.HexColor("#F5F5F5")
    MID_GRAY = colors.HexColor("#CCCCCC")
    WHITE = colors.white

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)

    style_title = ParagraphStyle("t", fontSize=14, textColor=DARK_GRAY, fontName="Helvetica-Bold", alignment=TA_LEFT)
    style_sub = ParagraphStyle("s", fontSize=9, textColor=DARK_GRAY, fontName="Helvetica", alignment=TA_LEFT)
    style_footer = ParagraphStyle("f", fontSize=7, textColor=MID_GRAY, fontName="Helvetica", alignment=TA_CENTER)

    story = []
    story.append(Paragraph(f"{settings.get('company_name','Company')} — Payroll Report", style_title))
    story.append(Paragraph(f"{settings.get('payroll_month','')} {settings.get('payroll_year','')}  |  {summary['total_employees']} employees  |  Net Total: ₦{summary['total_net']:,.2f}", style_sub))
    story.append(HRFlowable(width="100%", thickness=2, color=ORANGE, spaceBefore=3*mm, spaceAfter=3*mm))

    # Table data
    show_cols = ["STAFF NAME"] + [c for c in df.columns if c != "STAFF NAME"][:10]
    show_cols = [c for c in show_cols if c in df.columns]
    headers = [c.replace(" ", "\n") for c in show_cols]

    page_w = landscape(A4)[0] - 30*mm
    col_widths = [40*mm] + [max((page_w - 40*mm) / max(len(show_cols)-1, 1), 15*mm)] * (len(show_cols)-1)

    tdata = [headers]
    for _, row in df.iterrows():
        r = []
        for col in show_cols:
            v = row.get(col, "")
            if col != "STAFF NAME":
                try:
                    r.append(f"₦{float(v):,.0f}")
                except:
                    r.append(str(v))
            else:
                r.append(str(v))
        tdata.append(r)

    style_cmds = [
        ("BACKGROUND", (0,0), (-1,0), ORANGE),
        ("TEXTCOLOR", (0,0), (-1,0), WHITE),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
        ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
        ("TEXTCOLOR", (0,1), (-1,-1), DARK_GRAY),
        ("ALIGN", (1,1), (-1,-1), "RIGHT"),
        ("ALIGN", (0,0), (0,-1), "LEFT"),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("LINEBELOW", (0,0), (-1,-1), 0.3, MID_GRAY),
    ]
    for i in range(1, len(tdata), 2):
        style_cmds.append(("BACKGROUND", (0,i), (-1,i), LIGHT_GRAY))

    t = Table(tdata, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(style_cmds))
    story.append(t)
    story.append(Spacer(1, 4*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GRAY))
    story.append(Paragraph(f"Generated {datetime.now().strftime('%d %B %Y')} | Smart Payroll Payslip Generator | Confidential", style_footer))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── Payroll audit log ─────────────────────────────────────────────────────────
def page_audit_log():
    st.markdown("## 🔎 Payroll Audit Log")
    audit_df = st.session_state.get("audit_df")

    if audit_df is None or audit_df.empty:
        st.info("📂 No audit data available. Upload and merge files first.")
        return

    settings = get_settings()

    st.markdown(f"**{len(audit_df)} audit entries** across {audit_df['Source File'].nunique()} source file(s) and {audit_df['Employee Name'].nunique()} employee(s).")

    # Filters
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        file_filter = st.multiselect("Filter by Source File", options=sorted(audit_df["Source File"].unique()))
    with col_f2:
        comp_filter = st.multiselect("Filter by Payroll Component", options=sorted(audit_df["Payroll Component"].unique()))

    filtered = audit_df.copy()
    if file_filter:
        filtered = filtered[filtered["Source File"].isin(file_filter)]
    if comp_filter:
        filtered = filtered[filtered["Payroll Component"].isin(comp_filter)]

    st.dataframe(filtered.reset_index(drop=True), use_container_width=True)

    # Export audit log
    col_e, col_c = st.columns(2)
    with col_e:
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            filtered.to_excel(writer, index=False, sheet_name="Audit Log")
        excel_buf.seek(0)
        st.download_button(
            "⬇️ Export Audit Log (Excel)",
            data=excel_buf.getvalue(),
            file_name=f"audit_log_{settings['payroll_month']}_{settings['payroll_year']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_c:
        csv_bytes = filtered.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Export Audit Log (CSV)",
            data=csv_bytes,
            file_name=f"audit_log_{settings['payroll_month']}_{settings['payroll_year']}.csv",
            mime="text/csv",
        )


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    for key, default in [
        ("zip_ready", False), ("merge_done", False),
        ("merged_df", None), ("audit_df", None),
        ("source_file_names", []),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    menu = render_sidebar()

    if menu == "⚙️ Settings":
        render_settings_page()
        return

    if menu == "📥 Multi-File Upload & Merge":
        page_multi_upload()
        return

    if menu == "🔎 Payroll Audit Log":
        page_audit_log()
        return

    # For all data pages, use merged df
    df = get_active_df()

    # Fallback: allow single-file upload on data pages if no merge done
    if df is None:
        uploaded = upload_section()
        if uploaded:
            try:
                df = load_payroll_data(uploaded)
                errors = validate_columns(df)
                if errors:
                    st.error(f"❌ Missing required columns: {', '.join(errors)}")
                    df = None
                else:
                    st.session_state["merged_df"] = df
                    st.session_state["source_file_names"] = [uploaded.name]
                    st.success(f"✅ Loaded {len(df)} employee records.")
            except Exception as exc:
                st.error(f"❌ Could not read file: {exc}")

    if menu == "📊 Dashboard":
        page_dashboard(df)
    elif menu == "🔍 Employee Search":
        page_employee_search(df)
    elif menu == "📄 Generate Single Payslip":
        page_single_payslip(df)
    elif menu == "📦 Generate Bulk Payslips":
        page_bulk_payslips(df)
    elif menu == "💾 Download ZIP Archive":
        page_download_zip(df)
    elif menu == "📋 Merged Master Report":
        page_master_report(df)


if __name__ == "__main__":
    main()

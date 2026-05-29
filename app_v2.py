"""
HubSpot Lead Sync V2 — Full import tool
Upload Excel → Map your columns → Connect HubSpot → Sync everything.
"""
import streamlit as st
import requests
from engine_v2 import (
    HubSpotV2, get_excel_info, read_excel_mapped,
    determine_stage, count_active, run_sync_v2, log_rows_to_csv,
    ALL_FIELDS, CONTACT_FIELDS, COMPANY_FIELDS, STAGE_FIELDS,
)

st.set_page_config(
    page_title="HubSpot Lead Sync V2",
    page_icon="🚀",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─── Session state ────────────────────────────────────────────────────────────

KEYS = ["file_bytes", "excel_info", "field_mapping", "mapping_confirmed",
        "hs", "pipelines", "stages", "last_pipeline_id",
        "companies", "sheets", "config", "sync_stats", "sync_log", "log_rows"]
for k in KEYS:
    if k not in st.session_state:
        st.session_state[k] = None

# ─── Header ───────────────────────────────────────────────────────────────────

st.title("🚀 HubSpot Lead Sync  —  V2")
st.caption(
    "Full import tool: map any Excel column to any HubSpot field. "
    "Syncs company, contacts (with phone, LinkedIn, job title), and deals."
)

# ─── STEP 1 · Upload Excel ────────────────────────────────────────────────────

st.divider()
st.subheader("Step 1 — Upload your Excel file")

uploaded = st.file_uploader("Lead Excel file (.xlsx)", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.read()
    if file_bytes != st.session_state.file_bytes:
        # New file uploaded — reset everything downstream
        st.session_state.file_bytes       = file_bytes
        st.session_state.excel_info       = None
        st.session_state.field_mapping    = None
        st.session_state.mapping_confirmed = None
        st.session_state.hs               = None
        st.session_state.companies        = None
        st.session_state.sync_stats       = None
        st.session_state.sync_log         = None
        st.session_state.log_rows         = None

    if st.session_state.excel_info is None:
        with st.spinner("Reading Excel…"):
            try:
                info = get_excel_info(file_bytes)
                if not info["sheets"]:
                    st.error("No lead sheets found. Make sure the file has the right structure.")
                    st.stop()
                st.session_state.excel_info = info
            except Exception as e:
                st.error(f"Could not read file: {e}")
                st.stop()

# ─── STEP 2 · Map columns ─────────────────────────────────────────────────────

if st.session_state.excel_info:
    info     = st.session_state.excel_info
    headers  = info["all_headers"]
    sheets   = info["sheets"]

    st.divider()
    st.subheader("Step 2 — Map your columns")
    st.caption(
        f"Detected **{len(sheets)} sheet(s)**: {', '.join(sheets)}. "
        "Tell us which Excel column contains each piece of data. "
        "Mark optional fields as **Skip** if your file doesn't have them."
    )

    # Column preview table for first sheet
    first_sheet = sheets[0]
    with st.expander(f"📋 Preview — {first_sheet} (first 3 rows)"):
        sample = info["sample"].get(first_sheet, [])
        if sample:
            st.dataframe(
                {f"Col {i+1}: {h}": [row[i] if i < len(row) else "" for row in sample]
                 for i, h in enumerate(headers[:20])},  # cap at 20 cols
                use_container_width=True,
            )

    skip_option = "⛔  Skip (not in my file)"
    col_options = [skip_option] + headers

    def default_idx(hint_words: list) -> int:
        """Try to auto-select the most likely column based on header keywords."""
        for word in hint_words:
            for i, h in enumerate(headers):
                if word.lower() in h.lower():
                    return i + 1  # +1 for skip_option offset
        return 0  # default to skip

    st.write("**Company fields**")
    co1, co2 = st.columns(2)
    with co1:
        s_company = st.selectbox(
            "🏢 Company name *(required)*",
            col_options,
            index=default_idx(["company"]),
            key="map_company",
        )
    with co2:
        s_website = st.selectbox(
            "🌐 Company website",
            col_options,
            index=default_idx(["web", "website"]),
            key="map_website",
        )

    st.write("**Contact fields**")
    c1, c2 = st.columns(2)
    with c1:
        s_name = st.selectbox(
            "👤 Contact full name *(required)*",
            col_options,
            index=default_idx(["name"]),
            key="map_name",
        )
        s_phone = st.selectbox(
            "📞 Phone number",
            col_options,
            index=default_idx(["phone"]),
            key="map_phone",
        )
    with c2:
        s_email = st.selectbox(
            "📧 Email address *(required)*",
            col_options,
            index=default_idx(["mail 1", "email", "mail"]),
            key="map_email",
        )
        s_jobtitle = st.selectbox(
            "💼 Job title",
            col_options,
            index=default_idx(["job title", "title", "position"]),
            key="map_jobtitle",
        )

    s_linkedin = st.selectbox(
        "🔗 LinkedIn profile URL",
        col_options,
        index=default_idx(["linkedin"]),
        key="map_linkedin",
    )

    st.write("**Activity / stage columns**  *(these should be TRUE/FALSE columns)*")
    st1, st2, st3 = st.columns(3)
    with st1:
        s_contacted = st.selectbox(
            "✅ Contacted",
            col_options,
            index=default_idx(["contacted"]),
            key="map_contacted",
        )
    with st2:
        s_responded = st.selectbox(
            "💬 Responded",
            col_options,
            index=default_idx(["responded"]),
            key="map_responded",
        )
    with st3:
        s_meeting = st.selectbox(
            "📅 Meeting scheduled",
            col_options,
            index=default_idx(["meeting"]),
            key="map_meeting",
        )

    def sel(v):
        return None if v == skip_option else v

    # Validate required fields
    required_missing = []
    if not sel(s_company): required_missing.append("Company name")
    if not sel(s_name):    required_missing.append("Contact name")
    if not sel(s_email):   required_missing.append("Email address")
    if not sel(s_contacted) and not sel(s_responded) and not sel(s_meeting):
        required_missing.append("At least one stage column (Contacted / Responded / Meeting)")

    if required_missing:
        st.warning(f"⚠️  Please map these required fields: **{', '.join(required_missing)}**")

    confirm_disabled = bool(required_missing)
    if st.button("✅  Confirm column mapping", use_container_width=True,
                 disabled=confirm_disabled):
        mapping = {
            "company_name":           sel(s_company),
            "website":                sel(s_website),
            "full_name":              sel(s_name),
            "email":                  sel(s_email),
            "phone":                  sel(s_phone),
            "jobtitle":               sel(s_jobtitle),
            "hs_linkedin_profile_url": sel(s_linkedin),
            "contacted":              sel(s_contacted),
            "responded":              sel(s_responded),
            "meeting":                sel(s_meeting),
        }
        st.session_state.field_mapping     = mapping
        st.session_state.mapping_confirmed = True
        # Parse the Excel with the new mapping
        with st.spinner("Applying column mapping…"):
            companies, sheets_done = read_excel_mapped(
                file_bytes, mapping, headers
            )
            st.session_state.companies = companies
            st.session_state.sheets    = sheets_done

# ─── STEP 3 · Connect to HubSpot ─────────────────────────────────────────────

if st.session_state.mapping_confirmed and st.session_state.companies:
    total, active, n_emails = count_active(st.session_state.companies)

    st.divider()
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Active companies",    active)
    col_b.metric("Total companies",     total)
    col_c.metric("Contacts with email", n_emails)
    st.caption(
        f"Sheets processed: **{', '.join(st.session_state.sheets)}**  ·  "
        f"{total - active} companies skipped (no activity yet)"
    )

    st.divider()
    st.subheader("Step 3 — Connect to HubSpot")

    token_input = st.text_input(
        "HubSpot Private App Token",
        type="password",
        placeholder="pat-eu1-...",
        help="HubSpot → Settings → Integrations → Private Apps",
    )

    if st.button("🔗  Connect", use_container_width=True, disabled=not token_input):
        with st.spinner("Connecting…"):
            try:
                hs = HubSpotV2(token_input)
                pipelines = hs.get_pipelines()
                if not pipelines:
                    st.error("No pipelines found in your HubSpot account.")
                    st.stop()
                st.session_state.hs              = hs
                st.session_state.pipelines       = pipelines
                st.session_state.stages          = None
                st.session_state.last_pipeline_id = None
                st.session_state.config          = None
                st.session_state.sync_stats      = None
                st.session_state.sync_log        = None
                st.session_state.log_rows        = None
            except requests.HTTPError as e:
                code = e.response.status_code if e.response else "?"
                if code == 401:
                    st.error("❌ Token invalid (401). Get a fresh token from HubSpot → Private Apps.")
                elif code == 403:
                    st.error(
                        "❌ Missing scopes (403). Add to your Private App:\n\n"
                        "- `crm.schemas.deals.read`\n"
                        "- `crm.objects.companies.read` + `.write`\n"
                        "- `crm.objects.contacts.read` + `.write`\n"
                        "- `crm.objects.deals.read` + `.write`\n\n"
                        "Then rotate the token and try again."
                    )
                else:
                    st.error(f"HubSpot error: HTTP {code}")
                st.stop()

# ─── STEP 4 · Pipeline & Stages ──────────────────────────────────────────────

if st.session_state.hs and st.session_state.pipelines:
    st.success("✅ Connected to HubSpot!")
    st.divider()
    st.subheader("Step 4 — Pipeline & Stage Mapping")
    st.caption("Choose your pipeline and map activity levels to deal stages.")

    pipeline_labels = [p["label"] for p in st.session_state.pipelines]
    chosen_label    = st.selectbox("Pipeline", pipeline_labels)
    chosen_pipeline = next(p for p in st.session_state.pipelines if p["label"] == chosen_label)
    chosen_id       = chosen_pipeline["id"]

    if st.session_state.last_pipeline_id != chosen_id:
        with st.spinner("Loading stages…"):
            st.session_state.stages          = st.session_state.hs.get_stages(chosen_id)
            st.session_state.last_pipeline_id = chosen_id

    if st.session_state.stages:
        stage_labels    = [s["label"] for s in st.session_state.stages]
        stage_by_label  = {s["label"]: s["id"] for s in st.session_state.stages}

        st.write("Map each activity level to the correct deal stage:")
        col1, col2, col3 = st.columns(3)
        with col1:
            s_c = st.selectbox("✅ Contacted →",         stage_labels, key="sc")
        with col2:
            s_r = st.selectbox("💬 Responded →",         stage_labels, key="sr")
        with col3:
            s_m = st.selectbox("📅 Meeting scheduled →", stage_labels, key="sm")

        st.session_state.config = {
            "pipeline_id": chosen_id,
            "stage_map": {
                "contacted": stage_by_label[s_c],
                "responded": stage_by_label[s_r],
                "meeting":   stage_by_label[s_m],
            },
        }

# ─── STEP 5 · Sync ───────────────────────────────────────────────────────────

if st.session_state.config:
    total, active, _ = count_active(st.session_state.companies)

    st.divider()
    st.subheader("Step 5 — Run Sync")
    st.caption(
        f"Will sync **{active} active companies** with all their contacts, "
        f"phone numbers, LinkedIn profiles, job titles, websites and deals."
    )

    if st.button("▶  Start Sync", use_container_width=True, type="primary"):
        st.session_state.sync_stats = None
        st.session_state.sync_log   = []
        st.session_state.log_rows   = []

        log_placeholder = st.empty()
        progress_bar    = st.progress(0, text="Starting…")
        processed       = [0]
        log_lines       = []

        total_active = sum(
            1 for c in st.session_state.companies.values()
            if determine_stage(c) is not None
        )

        def log_cb(msg: str):
            log_lines.append(msg)
            st.session_state.sync_log = log_lines.copy()
            log_placeholder.code("\n".join(log_lines[-45:]), language=None)
            if msg.startswith("▸"):
                processed[0] += 1
                pct  = min(processed[0] / max(total_active, 1), 1.0)
                name = msg.split("▸")[-1].split("→")[0].strip()
                progress_bar.progress(pct, text=f"Syncing {name}…")

        try:
            stats, log_rows = run_sync_v2(
                companies   = st.session_state.companies,
                hs          = st.session_state.hs,
                pipeline_id = st.session_state.config["pipeline_id"],
                stage_map   = st.session_state.config["stage_map"],
                log         = log_cb,
            )
            progress_bar.progress(1.0, text="Done!")
            st.session_state.sync_stats = stats
            st.session_state.log_rows   = log_rows
        except Exception as e:
            st.error(f"Sync failed: {e}")

    # ── Results ───────────────────────────────────────────────────────────────

    if st.session_state.sync_stats:
        stats = st.session_state.sync_stats
        st.divider()
        st.subheader("✅ Sync Complete")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Companies created",  stats["co_created"])
        c2.metric("Contacts created",   stats["ct_created"])
        c3.metric("Contacts updated",   stats["ct_updated"])
        c4.metric("Deals created",      stats["deal_created"])

        c5, c6, c7, _ = st.columns(4)
        c5.metric("Companies updated",  stats["co_updated"])
        c6.metric("Deals updated",      stats["deal_updated"])
        c7.metric("Errors",             stats["errors"])

        if stats["errors"]:
            st.warning(f"⚠️  {stats['errors']} error(s) — see the log above for details.")
        if stats["skipped"]:
            st.info(f"ℹ️  {stats['skipped']} companies skipped (no activity).")

        # Download log
        if st.session_state.log_rows:
            csv_data = log_rows_to_csv(st.session_state.log_rows)
            st.download_button(
                label="📥  Download sync log (CSV)",
                data=csv_data,
                file_name="hubspot_sync_log.csv",
                mime="text/csv",
                use_container_width=True,
            )

        # Full log expander
        if st.session_state.sync_log:
            with st.expander("📋 Full sync log"):
                st.code("\n".join(st.session_state.sync_log), language=None)

# ─── Footer ───────────────────────────────────────────────────────────────────

st.divider()
st.caption("HubSpot Lead Sync V2  ·  Maps any column  ·  Phone · LinkedIn · Website · Full logging")

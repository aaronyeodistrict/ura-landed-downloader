"""
AYD Landed Transaction Downloader — Web App
Streamlit UI + Supabase authentication (admin-managed accounts).
"""

import streamlit as st
import streamlit.components.v1 as components
from supabase import create_client, Client
from datetime import datetime
from ura_utils import (
    POSTAL_DISTRICTS, MONTHS, YEAR_RANGE, LANDED_TYPES,
    collect_rows_website, write_excel_to_bytes, get_all_project_names_website,
)

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AYD Landed Transaction Downloader",
    page_icon="🏠",
    layout="wide",
)

# ── Supabase client (cached) ───────────────────────────────────────────────
@st.cache_resource
def _get_supabase() -> Client:
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_KEY"],
    )

supabase = _get_supabase()

# ── URA project list — cached 7 days, shared across all users ─────────────
@st.cache_data(ttl=60 * 60 * 24 * 7, show_spinner=False)
def _cached_project_names():
    return get_all_project_names_website()

# ── Auth helpers ───────────────────────────────────────────────────────────
def _init_session():
    for key, default in [
        ("user", None),
        ("projects", []),
        ("invite_flow", False),
        ("excel_bytes", None),
        ("excel_name", None),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

def _logout():
    try:
        supabase.auth.sign_out()
    except Exception:
        pass
    for key in ["user", "projects", "invite_flow", "excel_bytes", "excel_name"]:
        st.session_state[key] = None if key == "user" else ([] if key == "projects" else False if key == "invite_flow" else None)
    st.rerun()

# ── Handle Supabase invite / magic-link redirect ───────────────────────────
def _inject_hash_handler():
    """JS that converts Supabase hash tokens → query params so Python can read them."""
    components.html("""
    <script>
    (function() {
        const hash = window.location.hash;
        if (hash && hash.includes('access_token')) {
            const p = new URLSearchParams(hash.slice(1));
            const qs = '?sb_at=' + encodeURIComponent(p.get('access_token') || '')
                     + '&sb_rt=' + encodeURIComponent(p.get('refresh_token') || '')
                     + '&sb_tp=' + encodeURIComponent(p.get('type') || '');
            history.replaceState(null, '', window.location.pathname + qs);
            window.location.reload();
        }
    })();
    </script>
    """, height=0)

def _check_invite_params():
    """If URL has Supabase tokens, set the session and enter invite flow."""
    qp = st.query_params
    if "sb_at" not in qp:
        return
    access_token  = qp.get("sb_at", "")
    refresh_token = qp.get("sb_rt", "")
    token_type    = qp.get("sb_tp", "")
    if not access_token:
        return
    try:
        resp = supabase.auth.set_session(access_token, refresh_token)
        st.session_state.user = resp.user
        meta = resp.user.user_metadata or {}
        st.session_state.projects = meta.get("projects", [])
        st.session_state.invite_flow = (token_type in ("invite", "recovery"))
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.error(f"Invitation link error — please ask admin to resend: {e}")
        st.stop()

# ── Set password page (shown after invite) ─────────────────────────────────
def show_set_password():
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.markdown("## 🏠 Welcome!")
        st.markdown("You've been invited. Please set your own password to complete sign-up.")
        st.markdown("---")
        new_pass     = st.text_input("New Password", type="password")
        confirm_pass = st.text_input("Confirm Password", type="password")
        if st.button("Set Password & Log In", use_container_width=True, type="primary"):
            if not new_pass:
                st.error("Please enter a password.")
            elif new_pass != confirm_pass:
                st.error("Passwords don't match.")
            elif len(new_pass) < 6:
                st.error("Password must be at least 6 characters.")
            else:
                try:
                    supabase.auth.update_user({"password": new_pass})
                    st.session_state.invite_flow = False
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to set password: {e}")

# ── Login page ─────────────────────────────────────────────────────────────
def show_login():
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.markdown("## 🏠 AYD Landed Transaction Downloader")
        st.markdown("---")
        email    = st.text_input("Email", placeholder="you@example.com")
        password = st.text_input("Password", type="password")
        if st.button("Login", use_container_width=True, type="primary"):
            if not email or not password:
                st.error("Please enter your email and password.")
            else:
                try:
                    resp = supabase.auth.sign_in_with_password(
                        {"email": email, "password": password}
                    )
                    st.session_state.user = resp.user
                    meta = resp.user.user_metadata or {}
                    st.session_state.projects = meta.get("projects", [])
                    st.rerun()
                except Exception as e:
                    msg = str(e)
                    if "Invalid login" in msg or "invalid_credentials" in msg:
                        st.error("Incorrect email or password.")
                    else:
                        st.error(f"Login failed: {msg}")
        st.caption("Contact your administrator for access.")

# ── Save projects to Supabase user metadata ────────────────────────────────
def _save_projects(projects):
    try:
        supabase.auth.update_user({"data": {"projects": projects}})
    except Exception:
        pass

# ── Main app ───────────────────────────────────────────────────────────────
def show_main():
    user = st.session_state.user

    with st.sidebar:
        st.markdown("### 🏠 AYD Downloader")
        st.caption(f"Logged in as **{user.email}**")
        st.markdown("---")
        if st.button("Logout", use_container_width=True):
            _logout()

    st.title("AYD Landed Transaction Downloader")
    st.caption("Singapore landed resale transactions from URA · Admin-access only")
    st.divider()

    # ── Filter mode ────────────────────────────────────────────────────────
    filter_mode = st.radio(
        "Filter by",
        ["Project Name", "Postal District"],
        horizontal=True,
        label_visibility="collapsed",
    )
    st.markdown(f"**FILTER BY:** {filter_mode.upper()}")

    loc_filters = []

    # ── Project mode ───────────────────────────────────────────────────────
    if filter_mode == "Project Name":
        projects = st.session_state.projects

        with st.expander("⚙️ Manage / Browse Projects", expanded=not projects):
            tab_manual, tab_browse = st.tabs(["Type Manually", "Browse URA List"])

            with tab_manual:
                new_name = st.text_input("Add project name", placeholder="e.g. BISHOPSGATE RESIDENCES")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("Add", use_container_width=True):
                        name = new_name.strip()
                        if name and name not in projects:
                            projects.append(name)
                            st.session_state.projects = projects
                            _save_projects(projects)
                            st.rerun()
                        elif name in projects:
                            st.warning("Already in your list.")
                with c2:
                    if projects:
                        to_del = st.selectbox("Remove", ["— select —"] + projects, key="del_proj")
                        if st.button("Remove selected", use_container_width=True):
                            if to_del != "— select —":
                                projects = [p for p in projects if p != to_del]
                                st.session_state.projects = projects
                                _save_projects(projects)
                                st.rerun()

            with tab_browse:
                st.info("First load takes ~5–10 minutes; cached for 7 days.")
                if st.button("Load / Refresh URA Project List", use_container_width=True):
                    with st.spinner("Scanning URA website — this may take several minutes..."):
                        try:
                            _cached_project_names.clear()
                            _cached_project_names()
                            st.success("Project list refreshed.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed: {e}")
                try:
                    all_names = _cached_project_names()
                    search_q  = st.text_input("Filter", placeholder="Type to search...")
                    filtered  = [n for n in all_names if search_q.lower() in n.lower()] if search_q else all_names
                    to_add    = st.multiselect(f"{len(filtered):,} projects", options=filtered, default=[], key="browse_sel")
                    if st.button("Add selected to my list", use_container_width=True, disabled=not to_add):
                        added = 0
                        for name in to_add:
                            if name not in projects:
                                projects.append(name); added += 1
                        st.session_state.projects = projects
                        _save_projects(projects)
                        if added:
                            st.success(f"Added {added} project(s).")
                            st.rerun()
                        else:
                            st.info("All selected already in your list.")
                except Exception:
                    st.warning("Project list not loaded yet.")

        if projects:
            st.markdown("**SELECT PROJECTS TO INCLUDE**")
            sel_all = st.checkbox("Select all", value=True, key="proj_sel_all")
            chosen  = []
            cols    = st.columns(2)
            for i, p in enumerate(projects):
                with cols[i % 2]:
                    if st.checkbox(p, value=sel_all, key=f"proj_{i}"):
                        chosen.append(p)
            loc_filters = chosen if chosen else projects
        else:
            st.warning("No projects saved yet. Use the panel above to add projects.")

    # ── District mode ──────────────────────────────────────────────────────
    else:
        st.markdown("**SELECT POSTAL DISTRICTS**")
        select_all = st.checkbox("All postal districts", value=False)
        if select_all:
            loc_filters = list(POSTAL_DISTRICTS)
        else:
            loc_filters = st.multiselect(
                "Choose districts", options=POSTAL_DISTRICTS, default=[],
                placeholder="Pick one or more districts...",
            )
            if not loc_filters:
                st.info("No districts selected — all will be included.")
                loc_filters = list(POSTAL_DISTRICTS)

    st.divider()

    # ── Property type ──────────────────────────────────────────────────────
    st.markdown("**PROPERTY TYPE**")
    pt_col1, pt_col2, pt_col3 = st.columns(3)
    with pt_col1: det  = st.checkbox("Detached House",      value=True)
    with pt_col2: semi = st.checkbox("Semi-Detached House", value=True)
    with pt_col3: terr = st.checkbox("Terrace House",       value=True)

    selected_types = set()
    if det:  selected_types.add("Detached House")
    if semi: selected_types.add("Semi-Detached House")
    if terr: selected_types.add("Terrace House")
    if not selected_types:
        selected_types = {"Detached House", "Semi-Detached House", "Terrace House"}

    _api_map = {"Detached House": "Detached", "Semi-Detached House": "Semi-detached", "Terrace House": "Terrace"}
    prop_type_filter = selected_types | {_api_map[k] for k in selected_types}

    st.divider()

    # ── Date range ─────────────────────────────────────────────────────────
    st.markdown("**SALE DATE RANGE**")
    now = datetime.now()
    dr1, dr2 = st.columns(2)
    with dr1:
        st.caption("From")
        fc1, fc2 = st.columns(2)
        with fc1: from_m = st.selectbox("Month", MONTHS, index=0,                  key="from_m", label_visibility="collapsed")
        with fc2: from_y = st.selectbox("Year",  YEAR_RANGE, index=0,              key="from_y", label_visibility="collapsed")
    with dr2:
        st.caption("To")
        tc1, tc2 = st.columns(2)
        with tc1: to_m = st.selectbox("Month", MONTHS, index=now.month - 1,        key="to_m",   label_visibility="collapsed")
        with tc2: to_y = st.selectbox("Year",  YEAR_RANGE, index=len(YEAR_RANGE)-1,key="to_y",   label_visibility="collapsed")

    try:    from_dt = datetime(int(from_y), MONTHS.index(from_m) + 1, 1)
    except: from_dt = None
    try:
        tm = MONTHS.index(to_m) + 1; ty = int(to_y)
        to_dt = datetime(ty + 1, 1, 1) if tm == 12 else datetime(ty, tm + 1, 1)
    except: to_dt = None

    if from_dt and to_dt and from_dt >= to_dt:
        st.error("'From' date must be before 'To' date.")

    st.divider()

    # ── Download ───────────────────────────────────────────────────────────
    mode = "district" if filter_mode == "Postal District" else "project"

    if st.button("⬇  Download Excel", use_container_width=True, type="primary",
                 disabled=(from_dt and to_dt and from_dt >= to_dt)):
        status_box = st.empty()
        progress   = st.progress(0)
        try:
            status_box.info("⏳ Connecting to URA website...")
            rows = collect_rows_website(
                lambda msg: status_box.info(f"⏳ {msg}"),
                loc_filters=loc_filters, from_dt=from_dt, to_dt=to_dt,
                mode=mode, prop_types=prop_type_filter,
            )
            if not rows:
                status_box.warning("No matching records found. Try widening the date range.")
                progress.empty()
            else:
                status_box.info(f"⏳ Writing {len(rows):,} rows to Excel...")
                progress.progress(95)
                xl_bytes = write_excel_to_bytes(rows, loc_filters)
                fname    = f"URA_Landed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.session_state.excel_bytes = xl_bytes
                st.session_state.excel_name  = fname
                progress.progress(100)
                status_box.success(f"✅ Done! {len(rows):,} rows ready.")
        except Exception as e:
            status_box.error(f"❌ Error: {e}")
            progress.empty()

    if st.session_state.excel_bytes:
        st.download_button(
            label="📥 Save Excel File",
            data=st.session_state.excel_bytes,
            file_name=st.session_state.excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("---")
    st.caption("By: Aaron Yeo")

# ── Entry point ────────────────────────────────────────────────────────────
_init_session()
_inject_hash_handler()
_check_invite_params()

if st.session_state.get("invite_flow"):
    show_set_password()
elif st.session_state.user is None:
    show_login()
else:
    show_main()

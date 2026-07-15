import hashlib
from datetime import datetime
 
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
 
st.set_page_config(page_title="Provider Lifecycle Study — Sitter Profiles", layout="wide")
 
SEGMENT_LABELS = {"M": "MID", "T": "TOP", "N": "NEW", "MID": "MID", "TOP": "TOP", "NEW": "NEW"}
BADGE_COLORS = {"TOP": "#1E7A34", "MID": "#2455C4", "NEW": "#6B3FBF"}
BADGE_BG = {"TOP": "#E9F6EC", "MID": "#EAF1FF", "NEW": "#F3EEFB"}
 
CORE_FIELDS = {"date", "user name", "segment", "week #"}
 
 
# ---------------- AUTH ----------------
def check_login(email, password):
    auth_cfg = st.secrets.get("auth", {})
    allowed_domain = auth_cfg.get("allowed_domain", "").strip().lower()
    shared_hash = auth_cfg.get("shared_password_hash", "")
 
    email = email.strip().lower()
    if not allowed_domain or not shared_hash:
        return False
    if not email.endswith("@" + allowed_domain):
        return False
    return hashlib.sha256(password.encode()).hexdigest() == shared_hash
 
 
def render_login():
    st.title("Provider Lifecycle Study — Sitter Profiles")
    st.caption("Sign in with your Rover email to view the diary study dashboard.")
    with st.form("login_form"):
        email = st.text_input("Email", placeholder="you@rover.com")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in")
    if submitted:
        if check_login(email, password):
            st.session_state.authenticated = True
            st.session_state.user_email = email.strip().lower()
            st.rerun()
        else:
            st.error("Incorrect email or password, or your email isn't a Rover address.")
 
 
def render_logout_button():
    with st.sidebar:
        st.caption(f"Signed in as {st.session_state.get('user_email', '')}")
        if st.button("Log out"):
            for key in ("authenticated", "user_email", "view", "current_user"):
                st.session_state.pop(key, None)
            st.rerun()
 
 
# ---------------- GOOGLE SHEETS ----------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]
 
 
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)
 
 
@st.cache_resource(show_spinner=False)
def get_drive_service():
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)
 
 
@st.cache_data(ttl=300, show_spinner=False)
def get_provider_description(name):
    """Looks for a Google Doc inside this provider's Drive subfolder (under the
    'Kick off interviews' folder) and returns its text as a description. Returns
    None if no folder / doc is found, or if the drive folder isn't configured."""
    drive_cfg = st.secrets.get("drive", {})
    root_folder_id = drive_cfg.get("provider_folder_id")
    if not root_folder_id:
        return None
 
    service = get_drive_service()
 
    safe_name = name.replace("'", "\\'")
    folder_query = (
        f"'{root_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' "
        f"and name = '{safe_name}'"
    )
    folders = service.files().list(q=folder_query, fields="files(id, name)").execute().get("files", [])
    if not folders:
        return None
    provider_folder_id = folders[0]["id"]
 
    doc_query = (
        f"'{provider_folder_id}' in parents and mimeType = 'application/vnd.google-apps.document' "
        "and trashed = false"
    )
    docs = service.files().list(
        q=doc_query, fields="files(id, name, modifiedTime)", orderBy="modifiedTime desc"
    ).execute().get("files", [])
    if not docs:
        return None
 
    doc_id = docs[0]["id"]
    content = service.files().export(fileId=doc_id, mimeType="text/plain").execute()
    text = content.decode("utf-8") if isinstance(content, bytes) else content
    return text.strip() or None
 
 
@st.cache_data(ttl=300, show_spinner="Loading diary study data from Google Sheets...")
def load_entries():
    client = get_gspread_client()
    spreadsheet_id = st.secrets["sheet"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
 
    entries_ws = None
    for ws in sh.worksheets():
        if "entries" in ws.title.lower():
            entries_ws = ws
            break
    if entries_ws is None:
        raise RuntimeError("Could not find a tab containing 'Entries' in the spreadsheet.")
 
    all_values = entries_ws.get_all_values()
    if not all_values:
        return []
 
    # The sheet may have a title/banner row above the real header row
    # (e.g. "ROVER DIARY STUDY — SYNTHESIS ANALYSIS TOOL"), so scan for the
    # row that actually contains the expected column names instead of
    # assuming row 0 is the header.
    header_row_idx = 0
    for i, row_values in enumerate(all_values[:10]):
        normalized = {c.strip().lower() for c in row_values}
        if "user name" in normalized and "date" in normalized:
            header_row_idx = i
            break
 
    raw_headers = all_values[header_row_idx]
    headers = []
    seen = {}
    for i, h in enumerate(raw_headers):
        name = h.strip()
        if not name:
            name = f"col_{i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        headers.append(name)
 
    records = []
    for row_values in all_values[header_row_idx + 1:]:
        if not any(cell.strip() for cell in row_values):
            continue
        padded = row_values + [""] * (len(headers) - len(row_values))
        records.append(dict(zip(headers, padded)))
 
    # Normalize keys for easier lookup while keeping original labels for display
    cleaned = []
    for r in records:
        row = {k.strip(): (v.strip() if isinstance(v, str) else v) for k, v in r.items()}
        if row.get("User Name"):
            cleaned.append(row)
    return cleaned
 
 
def build_participants(entries):
    participants = {}
    for row in entries:
        name = row.get("User Name", "").strip()
        if not name:
            continue
        seg_raw = str(row.get("Segment", "")).strip().upper()
        segment = SEGMENT_LABELS.get(seg_raw, seg_raw or "NEW")
        if name not in participants:
            participants[name] = {"name": name, "segment": segment, "entries": [], "country": "", "age": ""}
        participants[name]["entries"].append(row)
        # Country/Age aren't in the Entries sheet today; this picks them up
        # automatically if those columns are ever added, without erroring otherwise.
        if not participants[name]["country"] and row.get("Country"):
            participants[name]["country"] = row["Country"]
        if not participants[name]["age"] and row.get("Age"):
            participants[name]["age"] = row["Age"]
 
    for p in participants.values():
        dates = [r.get("Date") for r in p["entries"] if r.get("Date")]
        p["entry_count"] = len(p["entries"])
        p["first_date"] = min(dates) if dates else None
        p["last_date"] = max(dates) if dates else None
    return participants
 
 
def badge_html(segment):
    color = BADGE_COLORS.get(segment, "#4B5565")
    bg = BADGE_BG.get(segment, "#F1F3F6")
    return f'<span style="font-size:11px;font-weight:700;padding:3px 8px;border-radius:6px;background:{bg};color:{color};">{segment}</span>'
 
 
def classify_entry(row):
    """Buckets each Entries row into one of three sections shown on a provider's profile."""
    question = str(row.get("Question", "")).lower()
    task_type = str(row.get("Task Type", "")).lower()
    experience = str(row.get("Experience Moment", "")).lower()
    week = str(row.get("Week #", "")).strip()
 
    if "reflect" in question or "reflect" in task_type or "reflect" in experience:
        return "reflection"
    if question.startswith("topic") or task_type.startswith("topic") or week == "0":
        return "background"
    return "activity"
 
 
# ---------------- NAVIGATION ----------------
def go_to_profile(name):
    st.session_state.view = "profile"
    st.session_state.current_user = name
 
 
def go_to_dash():
    st.session_state.view = "dash"
    st.session_state.current_user = None
 
 
# ---------------- DASHBOARD VIEW ----------------
def render_dashboard(participants):
    st.title("Provider Lifecycle Study — Sitter Profiles")
    st.caption(
        "Data refreshes automatically from the Rover_Diary_Study_Tool Google Sheet (fed from MyInsights). "
        "Select a provider to see their kickoff background, weekly reflections, and diary activity feed."
    )
 
    top_col1, top_col2 = st.columns([5, 1])
    with top_col1:
        st.write("")
    with top_col2:
        if st.button("Refresh data", use_container_width=True):
            load_entries.clear()
            st.rerun()
 
    st.markdown("**Filter by segment**")
    seg_filter = st.radio("Segment", ["ALL", "NEW", "MID", "TOP"], horizontal=True, label_visibility="collapsed")
 
    filtered = [p for p in participants.values() if seg_filter == "ALL" or p["segment"] == seg_filter]
    filtered.sort(key=lambda p: p["name"])
 
    if not filtered:
        st.info("No providers match your filters yet.")
        return
 
    st.write("")
    cols = st.columns(4)
    for i, p in enumerate(filtered):
        with cols[i % 4]:
            with st.container(border=True):
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                    f'<b>{p["name"]}</b>{badge_html(p["segment"])}</div>',
                    unsafe_allow_html=True,
                )
                st.caption(f"{p['entry_count']} entries")
                st.caption(f"Country: {p['country'] or '—'}  ·  Age: {p['age'] or '—'}")
                st.button(
                    "View profile", key=f"btn_{p['name']}",
                    on_click=go_to_profile, args=(p["name"],), use_container_width=True,
                )
 
    st.divider()
    st.caption(
        f"{len(participants)} providers total, sourced live from the Entries tab of Rover_Diary_Study_Tool. "
        "Segment codes: N = new, M = mid, T = top."
    )
 
 
# ---------------- PROFILE VIEW ----------------
def render_profile(participants):
    name = st.session_state.current_user
    p = participants.get(name)
 
    st.button("← All providers", on_click=go_to_dash)
 
    if not p:
        st.warning(f"No data found for {name}. It may have been removed from the sheet.")
        return
 
    header_col1, header_col2 = st.columns([5, 1])
    with header_col1:
        st.header(p["name"])
        st.caption(
            f"Segment: **{p['segment']}** · {p['entry_count']} entries "
            f"· {p['first_date'] or '—'} to {p['last_date'] or '—'}"
        )
    with header_col2:
        st.markdown(badge_html(p["segment"]), unsafe_allow_html=True)
 
    try:
        description = get_provider_description(p["name"])
    except Exception:
        description = None
    if description:
        st.write(description)
    else:
        st.caption(
            "No description yet — add a Google Doc to this provider's folder inside "
            "\"Kick off interviews\" in Drive and it will show up here."
        )
 
    entries = p["entries"]
    background_entries = [r for r in entries if classify_entry(r) == "background"]
    reflection_entries = [r for r in entries if classify_entry(r) == "reflection"]
    activity_entries = [r for r in entries if classify_entry(r) == "activity"]
 
    tab_background, tab_reflections, tab_diary, tab_interviews = st.tabs(
        ["Background", "Weekly reflections", "Diary entries", "Interviews"]
    )
 
    with tab_background:
        if not background_entries:
            st.info("No kickoff / background entries logged yet for this provider.")
        else:
            for row in background_entries:
                with st.container(border=True):
                    label = row.get("Question") or row.get("Task Type") or "Entry"
                    st.markdown(f"**{label}**")
                    for field, value in row.items():
                        key = field.strip().lower()
                        if key in CORE_FIELDS or key in ("question",) or not value:
                            continue
                        st.caption(f"**{field}:** {value}")
 
    with tab_reflections:
        if not reflection_entries:
            st.info(
                "No weekly reflections logged yet for this provider. These will appear here once "
                "loaded into MyInsights and synced to the Entries tab."
            )
        else:
            for row in reflection_entries:
                with st.container(border=True):
                    label = row.get("Week #")
                    st.markdown(f"**Week {label}**" if label not in (None, "") else "**Weekly reflection**")
                    for field, value in row.items():
                        key = field.strip().lower()
                        if key in CORE_FIELDS or not value:
                            continue
                        st.caption(f"**{field}:** {value}")
 
    with tab_diary:
        if not activity_entries:
            st.info("No diary entries logged yet for this provider.")
        else:
            display_cols = [
                "Date", "Week #", "Workflow Area", "Tool Used", "Platform Role",
                "Need", "Friction", "Emotion", "Signal", "Opportunity", "Description",
            ]
            table = []
            for row in activity_entries:
                table.append({c: row.get(c, "") for c in display_cols})
            st.dataframe(table, use_container_width=True, hide_index=True)
 
    with tab_interviews:
        st.info(
            "No interview summaries yet. Once you upload kickoff or recurrent interview files to "
            "this provider's Drive folder, ask Claude to wire this tab up to read and summarize them."
        )
 
 
# ---------------- ROUTER ----------------
def main():
    if not st.session_state.get("authenticated"):
        render_login()
        return
 
    render_logout_button()
 
    if "gcp_service_account" not in st.secrets or "sheet" not in st.secrets:
        st.error(
            "Google Sheets connection is not configured yet. Add `gcp_service_account` and "
            "`sheet` to your Streamlit secrets (see setup instructions)."
        )
        return
 
    try:
        entries = load_entries()
    except Exception as e:
        st.error(f"Could not load data from Google Sheets: {e}")
        return
 
    participants = build_participants(entries)
 
    if "view" not in st.session_state:
        st.session_state.view = "dash"
    if "current_user" not in st.session_state:
        st.session_state.current_user = None
 
    if st.session_state.view == "dash":
        render_dashboard(participants)
    else:
        render_profile(participants)
 
 
if __name__ == "__main__":
    main()

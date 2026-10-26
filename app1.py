import os
import re
import json
import sqlite3
from contextlib import closing
from io import BytesIO

import pandas as pd
import streamlit as st

from ppt_generator1 import generate_ppt

# ------------------------------------------------------------------ baseline + tariff store
# Initial values/default tariffs come from Data/site_config.xlsx (read-only).
# Parameter_Config needs: Site, Parameter, Baseline, Tariff
# Systems_Config   needs: Site, System, Baseline, Redundancy
DB_PATH = "Data/assessments.db"

# True  -> after submit, the new baseline = baseline + "New Load" (post-mitigation)
#          whenever a mitigation was entered.
# False -> always use the original projected load.
USE_MITIGATED_LOAD = True

_NOW = "strftime('%Y-%m-%d %H:%M:%f','now')"  # millisecond timestamps keep ordering exact

def _conn():
    return closing(sqlite3.connect(DB_PATH))

# ---------------------------------------------------------------- setup
def init_tables():
    with _conn() as conn, conn:
        # Append-only ledger: one row per change. Each row's Before = the previous row's After.
        conn.execute("""
            CREATE TABLE IF NOT EXISTS BaselineLedger (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Site TEXT, Category TEXT, Item TEXT,
                Event TEXT, ProjectTitle TEXT,
                BaselineBefore REAL, LoadChange REAL, BaselineAfter REAL,
                CreatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )""")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ledger_item "
                     "ON BaselineLedger (Site, Category, Item, ID)")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS TariffHistory (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Site TEXT, Category TEXT, Item TEXT,
                EffectiveYear INTEGER, Tariff REAL, ProjectTitle TEXT,
                CreatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (Site, Category, Item, EffectiveYear)
            )""")

# ---------------------------------------------------------------- reads
def excel_default(cfg_df, site, name_col, item, value_col):
    """Starting value from the Excel config sheet (0.0 if missing/blank)."""
    candidates = [value_col]
    if value_col == "Baseline":
        candidates = ["Baseline", "Initial Baseline", "CUrrent Baseline"]
    rows = cfg_df[(cfg_df["Site"].astype(str).str.strip() == str(site).strip()) & (cfg_df[name_col].astype(str).str.strip() == str(item).strip())]
    if rows.empty:
        return 0.0
    for col in candidates:
        if col in cfg_df.columns and pd.notna(rows.iloc[0][col]):
            return float(rows.iloc[0][col])
    return 0.0

def _head(conn, site, category, item):
    """Latest 'After' value in the ledger for this item (None if no history yet)."""
    row = conn.execute("""SELECT BaselineAfter FROM BaselineLedger
        WHERE Site=? AND Category=? AND Item=? ORDER BY ID DESC LIMIT 1""",
        (site, category, item)).fetchone()
    return float(row[0]) if row else None

def get_current_baseline(site, category, item, default):
    """Latest ledger value, or the Excel starting value if the item has no history yet."""
    with _conn() as conn:
        head = _head(conn, site, category, item)
    return default if head is None else head

def _tariff_lookup(conn, site, category, item, year, default):
    row = conn.execute("""
        SELECT Tariff FROM TariffHistory
        WHERE Site=? AND Category=? AND Item=? AND EffectiveYear<=?
        ORDER BY EffectiveYear DESC LIMIT 1""",
        (site, category, item, int(year))).fetchone()
    return float(row[0]) if row and row[0] is not None else default

def get_current_tariff(site, category, item, year, default):
    """Latest tariff whose EffectiveYear <= project year, else the Excel default."""
    with _conn() as conn:
        return _tariff_lookup(conn, site, category, item, year, default)

# ---------------------------------------------------------------- writes
def _add_row(conn, site, category, item, event, project, before, after):
    change = after - (before if before is not None else 0.0)
    conn.execute(f"""INSERT INTO BaselineLedger
        (Site, Category, Item, Event, ProjectTitle, BaselineBefore, LoadChange, BaselineAfter, CreatedAt)
        VALUES (?,?,?,?,?,?,?,?,{_NOW})""",
        (site, category, item, event, project, before, change, after))

def _upsert_tariff(conn, site, category, item, year, tariff, project):
    conn.execute(f"""INSERT INTO TariffHistory
        (Site, Category, Item, EffectiveYear, Tariff, ProjectTitle, CreatedAt)
        VALUES (?,?,?,?,?,?,{_NOW})
        ON CONFLICT(Site, Category, Item, EffectiveYear)
        DO UPDATE SET Tariff=excluded.Tariff,
                      ProjectTitle=excluded.ProjectTitle,
                      CreatedAt={_NOW}""",
        (site, category, item, int(year), float(tariff), project))

def _final_load(r):
    projected = r.get("Projected Load")
    if not isinstance(projected, (int, float)):
        return None
    if USE_MITIGATED_LOAD:
        mit, new_load = r.get("Mitigation"), r.get("New Load")
        if isinstance(mit, str) and mit.strip() not in ("", "N/A") \
                and isinstance(new_load, (int, float)):
            return float(r["Baseline"]) + float(new_load)
    return float(projected)

def commit_submission(site, project_title, year, parameter_results, systems_results,
                      parameter_config_df, systems_config_df):
    """Called when a project is submitted. Adds rows to the ledger (never edits old ones),
    all inside one transaction so the chain can't be left half-written."""
    groups = [
        ("Parameter", "Parameter", "Tariff", parameter_results, parameter_config_df),
        ("System", "System", "Redundancy", systems_results, systems_config_df),
    ]
    with _conn() as conn, conn:
        conn.execute("BEGIN IMMEDIATE")          # lock so two submits can't interleave
        for category, name_key, limit_key, results, cfg in groups:
            for r in results:
                item = r[name_key]
                baseline = r["Baseline"]
                if not isinstance(baseline, (int, float)):      # N/A baseline -> skip
                    continue
                baseline = float(baseline)
                final = _final_load(r)                           # projected load after the project

                head = _head(conn, site, category, item)
                if head is None:                                 # first time: record the Excel starting value
                    start = excel_default(cfg, site, name_key, item, "Baseline")
                    _add_row(conn, site, category, item, "Starting Value", None, None, start)
                    head = start

                key = (site, category, item, project_title)
                first = conn.execute("""SELECT BaselineBefore FROM BaselineLedger
                    WHERE Site=? AND Category=? AND Item=? AND ProjectTitle=?
                    ORDER BY ID LIMIT 1""", key).fetchone()

                if first is None:
                    # First submission of this project
                    if abs(baseline - head) > 1e-9:
                        _add_row(conn, site, category, item, "Manual Correction",
                                 project_title, head, baseline)
                        head = baseline
                    if final is not None:
                        _add_row(conn, site, category, item, "Project Submitted",
                                 project_title, head, final)
                else:
                    # Re-submission: append only the difference, so later projects stay intact
                    effect = conn.execute("""SELECT COALESCE(SUM(LoadChange), 0) FROM BaselineLedger
                        WHERE Site=? AND Category=? AND Item=? AND ProjectTitle=?""", key).fetchone()[0]
                    target = final if final is not None else baseline
                    adjust = (target - float(first[0])) - effect
                    if abs(adjust) > 1e-9:
                        _add_row(conn, site, category, item, "Project Resubmitted",
                                 project_title, head, head + adjust)

                # tariff / redundancy limit: save only if the user changed it
                limit = r.get(limit_key)
                if isinstance(limit, (int, float)):
                    t_default = excel_default(cfg, site, name_key, item, limit_key)
                    current = _tariff_lookup(conn, site, category, item, year, t_default)
                    if abs(float(limit) - current) > 1e-9:
                        _upsert_tariff(conn, site, category, item, year, limit, project_title)

# ---------------------------------------------------------------- UI helpers
def seeded_number(label, key, default, **kwargs):
    """number_input whose default is only applied the first time the key is used.
    Loaded/saved values (already in session_state) win, and there is no
    'value + session_state' Streamlit warning."""
    if key not in st.session_state:
        st.session_state[key] = float(default)
    return st.number_input(label, key=key, **kwargs)

def reset_form_state():
    """Clear the form for a new project. Baseline/tariff/redundancy keys are DELETED
    so they re-seed from the database (latest baseline, current tariff)."""
    for key in list(st.session_state.keys()):
        if "_na_" in key:
            st.session_state[key] = False
        elif key in ("tool_completed", "design_reviewed"):
            st.session_state[key] = "NA"
        elif key in ("key_takeaway", "project_suffix") or "mitigation" in key:
            st.session_state[key] = ""
        elif any(x in key for x in ("baseline", "tariff", "redundancy")):
            del st.session_state[key]
        elif "_load_" in key or key.startswith("load_"):
            st.session_state[key] = 0.0

# ---------------------------------------------------------------- monitoring export
def monitoring_workbook(site, parameter_config_df, systems_config_df):
    """Excel report for ONE site: Summary (starting vs current), Changes (its ledger), Tariff History."""
    rows = []
    for category, df, col in [("Parameter", parameter_config_df, "Parameter"),
                              ("System", systems_config_df, "System")]:
        for _, r in df[df["Site"].astype(str).str.strip() == site].iterrows():
            starting = float(r["Baseline"]) if pd.notna(r.get("Baseline")) else 0.0
            current = get_current_baseline(site, category, r[col], starting)
            rows.append({"Category": category, "Item": r[col],
                         "Starting Baseline (Excel)": starting, "Current Baseline": current,
                         "Change": current - starting})
    with _conn() as conn:
        ledger = pd.read_sql_query("""SELECT Category, Item, Event, ProjectTitle,
            BaselineBefore, LoadChange, BaselineAfter, CreatedAt
            FROM BaselineLedger WHERE Site = ? ORDER BY Category, Item, ID""", conn, params=(site,))
        tariffs = pd.read_sql_query("""SELECT Category, Item, EffectiveYear, Tariff, ProjectTitle, CreatedAt
            FROM TariffHistory WHERE Site = ? ORDER BY Category, Item, EffectiveYear DESC""",
            conn, params=(site,))
    ledger = ledger.rename(columns={"ProjectTitle": "Project", "BaselineBefore": "Before",
                                    "LoadChange": "Change", "BaselineAfter": "After",
                                    "CreatedAt": "Date"})
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name="Summary", index=False)
        ledger.to_excel(xw, sheet_name="Changes", index=False)
        tariffs.to_excel(xw, sheet_name="Tariff History", index=False)
    return buf.getvalue()

# ------------------------------------------------------------------ data
@st.cache_data
def load_data(file_mtime):  # file_mtime only refreshes the cache when the Excel changes
    sheets = ("Parameter", "Systems", "Parameter_Config", "Systems_Config")
    dfs = [pd.read_excel("Data/site_config.xlsx", sheet_name=s) for s in sheets]
    for df in dfs:
        df.columns = df.columns.str.strip()
        df["Site"] = df["Site"].astype(str).str.strip()
    return dfs

parameter_df, systems_df, parameter_config_df, systems_config_df = load_data(
    os.path.getmtime("Data/site_config.xlsx"))

# ------------------------------------------------------------------ database (assessments)
def get_connection():
    return sqlite3.connect("Data/assessments.db")

def db(sql, params=(), fetch=None):
    with closing(get_connection()) as conn, conn:
        cur = conn.execute(sql, params)
        if fetch == "one":
            return cur.fetchone()
        if fetch == "all":
            return cur.fetchall()

def create_table():
    db("""CREATE TABLE IF NOT EXISTS Assessments (
            ProjectTitle TEXT PRIMARY KEY, Site TEXT, AssessmentData TEXT,
            LastUpdatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
    cols = {r[1] for r in db("PRAGMA table_info(Assessments)", fetch="all")}
    if "Islocked" not in cols:
        db("ALTER TABLE Assessments ADD COLUMN Islocked INTEGER DEFAULT 0")

create_table()
init_tables()  # BaselineLedger + TariffHistory

def save_assessment(project_title, site, data):
    db("""INSERT INTO Assessments (ProjectTitle, Site, AssessmentData, LastUpdatedAt)
          VALUES (?, ?, ?, CURRENT_TIMESTAMP)
          ON CONFLICT(ProjectTitle) DO UPDATE SET
              Site = excluded.Site, AssessmentData = excluded.AssessmentData,
              LastUpdatedAt = CURRENT_TIMESTAMP""",
       (project_title, site, json.dumps(data, default=str)))

def load_assessment(project_title):
    row = db("SELECT Site, AssessmentData FROM Assessments WHERE ProjectTitle = ?",
             (project_title,), fetch="one")
    return (row[0], json.loads(row[1])) if row else (None, None)

def get_saved_assessments(site):
    rows = db("SELECT ProjectTitle FROM Assessments WHERE Site = ? ORDER BY ProjectTitle",
              (site,), fetch="all")
    return [r[0] for r in rows]

def delete_assessment(project_title):
    db("DELETE FROM Assessments WHERE ProjectTitle = ?", (project_title,))

def update_lock_status(project_title, locked):
    db("UPDATE Assessments SET Islocked = ? WHERE ProjectTitle = ?", (locked, project_title))

def get_lock_status(project_title):
    row = db("SELECT Islocked FROM Assessments WHERE ProjectTitle = ?",
             (project_title,), fetch="one")
    return bool(row[0]) if row else False

# ------------------------------------------------------------------ session helpers
EXCLUDED_KEYS = {
    "project_title", "selected_assessment", "selected_assessment_previous",
    "loaded_assessment", "site", "save_error", "save_success", "delete_success",
    "submit_success", "pending_select", "save_requested", "is_duplicate", "duplicate_requested",
    "duplicate_site", "duplicate_data", "reload_project", "skip_auto_save_once", "new_project_reset",
}

def collect_assessment_data():
    return {k: v for k, v in st.session_state.items() if k not in EXCLUDED_KEYS}

def apply_saved_data(saved_data):
    for k, v in saved_data.items():
        if k not in EXCLUDED_KEYS:
            st.session_state[k] = v

def save_current_assessment():
    site = st.session_state.get("site")
    project_title = st.session_state.get("project_title")

    if not site or site == "-- Select Site --" or not project_title:
        st.session_state["save_error"] = True
        return
    if (st.session_state.get("selected_assessment") == "New Project"
            and not st.session_state.get("project_suffix", "").strip()):
        st.session_state["save_error"] = True
        return

    # NOTE: saving a draft never touches baselines/tariffs. That only happens on Submit.
    is_new = st.session_state.get("selected_assessment") == "New Project"
    data = collect_assessment_data()
    if is_new and st.session_state.get("duplicate_data"):
        # The tabs are hidden on a new project, so Streamlit has dropped the copied box values.
        # They were kept safely in duplicate_data; use them as the base (name/year typed now win).
        data = {**st.session_state["duplicate_data"], **data}
    save_assessment(project_title, site, data)
    st.session_state.pop("duplicate_data", None)
    st.session_state["is_duplicate"] = False
    st.session_state["new_project_reset"] = False
    st.session_state["save_success"] = True
    if is_new:
        # Not marked as loaded on purpose: the project is opened from the database next run,
        # which fills the boxes with the saved (copied) values.
        st.session_state.pop("loaded_assessment", None)
        st.session_state["pending_select"] = project_title  # applied before the dropdown is drawn
    else:
        st.session_state["loaded_assessment"] = project_title

def auto_save_current_assessment():
    if st.session_state.get("skip_auto_save_once", False):
        st.session_state["skip_auto_save_once"] = False
        return
    if st.session_state.get("new_project_reset", False) or st.session_state.get("is_duplicate", False):
        return
    project_title = st.session_state.get("loaded_assessment")
    if not project_title or get_lock_status(project_title):
        return
    saved_site, saved_data = load_assessment(project_title)
    if saved_site is None:
        return
    data = collect_assessment_data()
    if data != saved_data:
        save_assessment(project_title, saved_site, data)

def prepare_duplicate(project_title):
    saved_site, saved_data = load_assessment(project_title)
    copied = {k: v for k, v in saved_data.items() if k not in EXCLUDED_KEYS}
    # Duplicate copies everything EXCEPT baselines and tariffs/redundancy limits (N/A ticks are kept).
    # Those boxes are left out so they start from the site's current baseline / current tariff.
    copied = {k: v for k, v in copied.items()
              if "_na_" in k or not any(x in k for x in ("baseline", "tariff", "redundancy"))}
    st.session_state["duplicate_data"] = copied
    st.session_state["duplicate_site"] = saved_site
    st.session_state["project_year"] = saved_data.get("project_year", 2026)
    suffix = saved_data.get("project_suffix", "").strip()
    st.session_state["project_suffix"] = suffix + "Copy" if suffix else "Copy"
    st.session_state["is_duplicate"] = True
    st.session_state["duplicate_requested"] = True
    st.session_state.pop("loaded_assessment", None)
    st.session_state.pop("is_locked", None)

def reset_new_project():
    reset_form_state()  # baseline/tariff keys are deleted -> re-seeded from the database
    st.session_state.pop("duplicate_data", None)
    st.session_state.pop("loaded_assessment", None)
    st.session_state["new_project_reset"] = True

# ------------------------------------------------------------------ styling + impact info
st.markdown("""
<style>
details { background-color: white; border-radius: 8px; border: 1px solid #D0D7DE; }
details summary { background-color: #A7C7E7 !important; color: black !important;
                  padding: 12px; border-radius: 8px; font-weight: bold; }
[data-testid="stMetric"] { background-color: #ffffff; border-left: 5px solid #87CEEB;
                           padding: 5px; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

impact_info = {
    "+": ("Positive Impact", "#CDECCF", "Projected load is reduced from the site baseline."),
    "=": ("No Significant Impact", "#FFE28A", "Project has no additional load impact on the site."),
    "-": ("Negative Impact", "#FFC48C", "Load from project increases the site baseline"),
    "- -": ("Double Negative Impact", "#FFB3B3",
            "Load from project increases site baseline and exceeds the tariff/redundancy limit"),
    "N/A": ("Inconclusive Impact", "#F3F4F6", "Parameter has no effect on impact for this project"),
}

# ------------------------------------------------------------------ top of page
st.title("Sustainability Assessment Tool")

if st.session_state.get("duplicate_requested", False):
    st.session_state["site"] = st.session_state["duplicate_site"]
    st.session_state["selected_assessment"] = "New Project"
    st.session_state.pop("duplicate_requested", None)

site_options = ["-- Select Site --"] + list(parameter_df["Site"].unique())
site = st.selectbox("Select Site", site_options, key="site")
saved_assessments = [] if site == "-- Select Site --" else get_saved_assessments(site)
if "pending_select" in st.session_state:
    st.session_state["selected_assessment"] = st.session_state.pop("pending_select")
selected_assessment = st.selectbox("Assessment", ["New Project"] + saved_assessments,
                                   key="selected_assessment")

LOCKED = get_lock_status(selected_assessment) if selected_assessment != "New Project" else False

if st.session_state.get("reload_project", False):
    _, saved_data = load_assessment(selected_assessment)
    if saved_data is not None:
        apply_saved_data(saved_data)
        st.session_state["loaded_assessment"] = selected_assessment
    st.session_state["reload_project"] = False

if st.session_state.pop("refresh_values", False):
    for key in list(st.session_state.keys()):
        if "_na_" not in key and any(x in key for x in ("baseline", "tariff", "redundancy")):
            del st.session_state[key]    # the boxes re-fill from the ledger / tariff history

if selected_assessment == "New Project":
    if (st.session_state.get("loaded_assessment")
            and not st.session_state.get("is_duplicate", False)
            and st.session_state.get("selected_assessment_previous") != "New Project"):
        reset_new_project()
        st.rerun()

    year = st.selectbox("Year", list(range(2026, 2036)), key="project_year")
    project_suffix = st.text_input("Project Name", key="project_suffix")

    numbers = []
    for p in saved_assessments:
        m = re.search(rf"{re.escape(site)}{year}\((\d+)\)", p)
        if m:
            numbers.append(int(m.group(1)))
    project_number = 1
    while project_number in numbers:
        project_number += 1

    project_title = f"{site}{year}({project_number}) - {project_suffix}"
    st.session_state["project_title"] = project_title
    st.session_state["selected_assessment_previous"] = selected_assessment
else:
    project_title = selected_assessment
    st.session_state["project_title"] = project_title
    st.session_state["selected_assessment_previous"] = selected_assessment
    st.session_state.pop("duplicate_data", None)
    if st.session_state.get("loaded_assessment") != project_title:
        _, saved_data = load_assessment(project_title)
        if saved_data is not None:
            apply_saved_data(saved_data)
            st.session_state["loaded_assessment"] = project_title
            st.session_state["new_project_reset"] = False
            st.session_state["is_duplicate"] = False
            st.rerun()

col1, col2, col3, col4 = st.columns(4)
with col1:
    if st.button("Save Assessment", type="primary"):
        st.session_state["save_requested"] = True
with col2:
    if selected_assessment != "New Project" and st.button("Delete Assessment"):
        delete_assessment(selected_assessment)
        st.session_state["delete_success"] = True
        st.session_state.pop("loaded_assessment", None)
        st.rerun()
with col3:
    if selected_assessment != "New Project" and st.button("Duplicate Assessment"):
        prepare_duplicate(selected_assessment)
        st.rerun()
with col4:
    if selected_assessment != "New Project":
        if LOCKED:
            if st.button("🔓 Unlock"):
                update_lock_status(selected_assessment, 0)
                st.session_state["skip_auto_save_once"] = True
                st.session_state["reload_project"] = True
                st.rerun()
        elif st.button("🔒 Lock"):
            update_lock_status(selected_assessment, 1)
            st.session_state["skip_auto_save_once"] = True
            st.session_state["reload_project"] = True
            st.rerun()

for flag, msg, kind in [
    ("save_success", "Assessment saved successfully!", st.success),
    ("submit_success", "Assessment submitted. Baselines and tariffs updated for future projects.", st.success),
    ("delete_success", "Assessment deleted successfully!", st.success),
    ("save_error", "Please enter a Project Title before saving.", st.error),
]:
    if st.session_state.get(flag):
        kind(msg)
        st.session_state.pop(flag, None)

# ------------------------------------------------------------------ rendering helpers
def calc_impact(baseline, load, limit):
    """Returns (projected load, impact symbol)."""
    if baseline is None or load is None:
        return "N/A", "N/A"
    projected = baseline + load
    if load == 0:
        return projected, "="
    if load < 0:
        return projected, "+"
    if limit is not None and projected > limit:
        return projected, "- -"
    return projected, "-"

def show_result(label, projected, impact):
    c1, c2 = st.columns(2)
    with c1:
        st.metric(label, projected)
    with c2:
        title, colour, desc = impact_info.get(impact, ("Unknown", "#E9ECEF", ""))
        st.markdown(
            f"""<div style="background-color:{colour}; padding:10px; border-radius:10px; line-height:1.2;">
                <p style="margin:2px 0;font-size:20px;"><b>{impact}</b></p>
                <p style="margin:2px 0;font-size:14px;"><b>{title}</b></p>
                <p style="margin:0;font-size:13px;">{desc}</p></div>""",
            unsafe_allow_html=True)

def na_number(label, key, na_key, default):
    """Number input + N/A checkbox. Returns None when N/A is ticked."""
    c1, c2 = st.columns([5, 2])
    with c2:
        st.write("")
        st.write("")
        na = st.checkbox("N/A", key=na_key, disabled=LOCKED)
    with c1:
        value = seeded_number(label, key, default, disabled=na or LOCKED, step=1.0)
    return None if na else value

def mk(index, **names):
    return {n: f"{v}_{index}" for n, v in names.items()}

def render_item(category, item, unit, k, year, cfg_df, name_col, limit_key):
    """One assessed item: baseline (auto from history), load, tariff/redundancy
    (auto from history), impact, mitigation and re-assessed impact."""
    is_sys = category == "System"
    if k["unit"] not in st.session_state:
        st.session_state[k["unit"]] = unit
    selected_unit = st.text_input("Unit", key=k["unit"], disabled=LOCKED)

    base_default = get_current_baseline(
        site, category, item, excel_default(cfg_df, site, name_col, item, "Baseline"))
    limit_default = get_current_tariff(
        site, category, item, year, excel_default(cfg_df, site, name_col, item, limit_key))

    c1, c2, c3 = st.columns(3)
    with c1:
        baseline = na_number("Site Baseline", k["b"], k["bna"], base_default)
    with c2:
        load = na_number("Site Load" if is_sys else "Load", k["l"], k["lna"], 0.0)
    with c3:
        limit = na_number("Redundancy Limit" if is_sys else "Tariff Limit",
                          k["t"], k["tna"], limit_default)
    st.caption("Choose N/A if the parameter is not applicable")
    if not LOCKED:
        if baseline is not None and abs(baseline - base_default) > 1e-9:
            st.caption(f"⚠️ The site's current baseline is {base_default:g}. Submitting with "
                       f"{baseline:g} records a manual correction.")
        if limit is not None and abs(limit - limit_default) > 1e-9:
            st.caption(f"⚠️ The current {'redundancy limit' if is_sys else 'tariff'} for {year} is "
                       f"{limit_default:g}. Submitting saves {limit:g} as the new value from {year}.")

    projected, impact = calc_impact(baseline, load, limit)
    show_result("Projected Load", projected, impact)

    mitigation, new_load, new_impact = "N/A", "N/A", "N/A"
    if impact in ("- -", "-", "="):
        mitigation = st.text_area("Mitigation", placeholder="Keep within 30 words",
                                  key=k["mit"], disabled=LOCKED)
        if len(mitigation.split()) > 30:
            st.error("Mitigation cannot exceed 30 words.")
        new_load = seeded_number("New Load", k["nl"], 0.0, step=1.0, disabled=LOCKED)
        new_projected, new_impact = calc_impact(
            baseline, new_load if load is not None else None, limit)
        show_result("Projected New Load", new_projected, new_impact)

    return {
        name_col: item, "Unit": selected_unit, "Baseline": baseline,
        "Projected Load": projected, limit_key: limit, "Impact": impact,
        "Mitigation": mitigation, "New Load": new_load, "New Impact": new_impact,
    }

# ------------------------------------------------------------------ tabs
parameter_results = []
systems_results = []
project_year = st.session_state.get("project_year", 2026)

if selected_assessment == "New Project":
    if site == "-- Select Site --":
        st.info("Please select a site.")
    else:
        st.info("Enter the Year and Project Name, then click Save Assessment. The Parameter, Systems and Summary tabs appear once the project is saved.")
else:
    tab1, tab2, tab3 = st.tabs(["Parameter", "Systems", "Assessment Summary"])

    with tab1:
        if site == "-- Select Site --":
            st.info("Please select a site.")
            st.stop()
        st.header("Parameter")

        for index, row in parameter_df[parameter_df["Site"] == site].iterrows():
            if row["Parameter"] == "Carbon Footprint":
                with st.expander("Carbon Footprint (tCO2e)"):
                    for n, title in ((1, "Scope 1: Natural Gas Emissions"),
                                     (2, "Scope 2: Electricity Emissions")):
                        if n == 2:
                            st.divider()
                        st.subheader(title)
                        keys = mk(index, unit=f"scope{n}_unit", bna=f"scope{n}_baseline_na",
                                  b=f"scope{n}_baseline", lna=f"scope{n}_load_na", l=f"scope{n}_load",
                                  tna=f"scope{n}_tariff_na", t=f"scope{n}_tariff",
                                  mit=f"scope{n}_mitigation", nl=f"scope{n}_new_load")
                        parameter_results.append(render_item(
                            "Parameter", f"Carbon Footprint Scope {n}", row["Unit"], keys,
                            project_year, parameter_config_df, "Parameter", "Tariff"))
            else:
                with st.expander(f"{row['Parameter']} ({row['Unit']})"):
                    keys = mk(index, unit="parameter_unit", bna="parameter_baseline_na", b="baseline",
                              lna="parameter_load_na", l="load", tna="tariff_na", t="tariff",
                              mit="parameter_mitigation", nl="parameter_new_load")
                    parameter_results.append(render_item(
                        "Parameter", row["Parameter"], row["Unit"], keys,
                        project_year, parameter_config_df, "Parameter", "Tariff"))

    with tab2:
        st.header("Systems")
        for index, row in systems_df[systems_df["Site"] == site].iterrows():
            with st.expander(f"{row['System']} ({row['Units']})"):
                keys = mk(index, unit="systems_units", bna="systems_baseline_na", b="system_baseline",
                          lna="systems_load_na", l="system_load", tna="system_redundancy_na",
                          t="system_redundancy", mit="system_mitigation", nl="system_new_load")
                systems_results.append(render_item(
                    "System", row["System"], row["Units"], keys,
                    project_year, systems_config_df, "System", "Redundancy"))

    with tab3:
        st.header("Assessment Summary")
        c1, c2 = st.columns(2)
        with c1:
            tool_completed = st.selectbox("Energy Assessment Tool completed ?",
                                          ["NA", "Yes", "No"], key="tool_completed", disabled=LOCKED)
        with c2:
            design_reviewed = st.selectbox("Is Sustainable Design AES 00.005STD reviewed?",
                                           ["NA", "Yes", "No"], key="design_reviewed", disabled=LOCKED)
        key_takeaway = st.text_area("Key Takeaway", height=150, key="key_takeaway", disabled=LOCKED)

        if st.button("Generate PPT"):
            ppt_file = generate_ppt(parameter_results, systems_results, tool_completed,
                                    design_reviewed, key_takeaway, site, project_title)
            with open(ppt_file, "rb") as f:
                st.download_button(
                    "Download PPT", f, file_name=f"{project_title}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

        st.divider()
        st.subheader("Submit")
        st.caption("Submitting locks the project and updates site baselines "
                   "(manual corrections + final project load) and any changed tariffs "
                   "for future projects.")
        if not LOCKED and st.button("Reset baselines & tariffs to current site values"):
            st.session_state["refresh_values"] = True
            st.rerun()
        if st.button("Submit Assessment", type="primary",
                     disabled=LOCKED):
            save_current_assessment()
            commit_submission(site, project_title, project_year, parameter_results,
                              systems_results, parameter_config_df, systems_config_df)
            update_lock_status(project_title, 1)
            st.session_state["submit_success"] = True
            st.session_state["skip_auto_save_once"] = True
            st.session_state["reload_project"] = True
            st.rerun()

        st.download_button(
            f"Download {site} monitoring Excel (starting vs current baselines, changes, tariffs)",
            monitoring_workbook(site, parameter_config_df, systems_config_df),
            file_name=f"baseline_monitor_{site}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------------------------------------------------------ end of run
if st.session_state.pop("save_requested", False):
    save_current_assessment()
    st.rerun()

auto_save_current_assessment()

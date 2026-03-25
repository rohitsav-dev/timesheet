from datetime import datetime, timedelta

import pandas as pd
import streamlit as st
from playwright.sync_api import sync_playwright
import subprocess
import sys

# --- CONSTANTS ---
WEEKLY_TARGET_HRS = 24
DAILY_TARGET_MINS = 480
MIS_URL = "https://cybagemis.cybage.com/Report%20Builder/RPTN/Reportpage.aspx"

# --- DATA PROCESSING HELPERS ---


def clean_text(text):
    if not text:
        return ""
    return str(text).replace("\n", "").replace("\r", "").strip()


def parse_hours_to_mins(time_str):
    if (
        not time_str
        or pd.isna(time_str)
        or str(time_str).strip() in ["", "-", "0:00", "0:00:00"]
    ):
        return 0
    try:
        parts = str(time_str).split(":")
        if len(parts) == 3:  # HH:MM:SS
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60
        elif len(parts) == 2:  # HH:MM
            return int(parts[0]) * 60 + int(parts[1])
        return float(time_str) * 60
    except Exception:
        return 0


def format_mins_to_hms(minutes):
    if minutes is None:
        return "0:00:00"
    total_seconds = int(round(abs(minutes) * 60))
    hours, remainder = divmod(total_seconds, 3600)
    mins, secs = divmod(remainder, 60)
    sign = "-" if minutes < 0 else ""
    return f"{sign}{int(hours)}:{int(mins):02d}:{int(secs):02d}"


def identify_gate_vba_style(gate_name):
    name = gate_name.lower()
    if any(x in name for x in ["floor", "wing"]):
        return "WorkGate"
    if any(x in name for x in ["main", "tripod", "parking"]):
        return "MainGate"
    return "PlayGate"


def calculate_vba_logic_today(swipes_list, emp_id):
    if not swipes_list:
        return 0, None, "Out", []

    sorted_swipes = sorted(
        swipes_list, key=lambda x: datetime.strptime(x["time"], "%I:%M:%S %p")
    )

    total_work_seconds = 0
    active_start_time = None
    prev_gate_group = None
    prev_direction = None

    first_punch = None
    proof_rows = []

    for i, s in enumerate(sorted_swipes):
        t = datetime.strptime(s["time"], "%I:%M:%S %p")
        direction = s["direction"].strip().upper()
        place = s["place"].strip()
        gate_group = identify_gate_vba_style(place)

        if first_punch is None:
            first_punch = t

        interval_val = 0
        interval_str = "-"

        if direction == "ENTRY":
            if active_start_time:
                if prev_gate_group == "MainGate" and prev_direction == "ENTRY":
                    interval_val = 0
                else:
                    interval_val = (t - active_start_time).total_seconds()

                total_work_seconds += interval_val
                interval_str = (
                    format_mins_to_hms(interval_val / 60) if interval_val > 0 else "-"
                )
                active_start_time = t
            else:
                active_start_time = t

        elif direction == "EXIT":
            if active_start_time:
                interval_val = (t - active_start_time).total_seconds()
                total_work_seconds += interval_val
                interval_str = (
                    format_mins_to_hms(interval_val / 60) if interval_val > 0 else "-"
                )

                if gate_group == "WorkGate":
                    active_start_time = t
                else:
                    active_start_time = None
            else:
                pass

        proof_rows.append(
            {
                "Employee ID": emp_id,
                "Date": s["date"],
                "Location": place,
                "Type": direction.capitalize(),
                "Time": s["time"],
                "Identified Gate": gate_group,
                "Area Status": (
                    "WORK/TRANSIT"
                    if active_start_time or interval_str != "-"
                    else "OFF/CAMPUS"
                ),
                "Final Addition": interval_str,
            }
        )

        prev_gate_group = gate_group
        prev_direction = direction

    now_str = datetime.now().strftime("%I:%M:%S %p")
    now_dt = datetime.strptime(now_str, "%I:%M:%S %p")
    current_state = "Out"

    if active_start_time:
        diff = (now_dt - active_start_time).total_seconds()
        total_work_seconds += diff
        proof_rows.append(
            {
                "Employee ID": emp_id,
                "Date": datetime.now().strftime("%d-%b-%Y"),
                "Location": "LIVE STATUS",
                "Type": "Still In",
                "Time": now_str,
                "Identified Gate": "WorkGate",
                "Area Status": "WORK",
                "Final Addition": format_mins_to_hms(diff / 60),
            }
        )
        current_state = "In"

    return (total_work_seconds / 60), first_punch, current_state, proof_rows


# --- SCRAPER LOGIC ---


def scrape_live_today(page, emp_id):
    try:
        page.locator("a:has-text(\"Today's and Yesterday's Swipe Log\")").evaluate(
            "node => node.click()"
        )
        page.locator("select[title='EmployeeID']").wait_for(
            state="visible", timeout=15000
        )
        page.locator("select[title='EmployeeID']").evaluate(
            f"(s)=>{{for(i=0;i<s.options.length;i++){{if(s.options[i].text.includes('{emp_id}')){{s.selectedIndex=i;s.dispatchEvent(new Event('change',{{bubbles:true}}));}}}}}}"
        )
        page.locator("select[title='Day']").evaluate(
            "(s)=>{for(i=0;i<s.options.length;i++){if(s.options[i].text.includes('Today')){s.selectedIndex=i;s.dispatchEvent(new Event('change',{bubbles:true}));}}}"
        )
        page.wait_for_timeout(1000)
        page.locator("#ViewReportImageButton").click()
        page.wait_for_load_state("networkidle")
        raw_rows = page.locator("table tr").all()

        swipes_raw = []
        seen_swipes = set()
        for row in raw_rows:
            cells = row.locator("td").all_inner_texts()
            if not cells:
                continue
            c = [clean_text(val) for val in cells]
            if any(str(emp_id) == val for val in c):
                idx = c.index(str(emp_id))
                if len(c) > idx + 4:
                    key = (c[idx + 4], c[idx + 2], c[idx + 3])
                    if key not in seen_swipes:
                        seen_swipes.add(key)
                        swipes_raw.append(
                            {
                                "date": c[idx + 1],
                                "place": c[idx + 2],
                                "direction": c[idx + 3],
                                "time": c[idx + 4],
                            }
                        )
        return calculate_vba_logic_today(swipes_raw, emp_id)
    except Exception as e:
        st.error(f"Scraper Error: {e}")
        return 0, None, "Out", []


def full_historical_sync(emp_id, username, password):
    data = []
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(
                headless=True,
                chromium_sandbox=False,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                ],
            )
        except Exception as e:
            msg = str(e)
            if "Executable doesn't exist" not in msg:
                raise
            # Fallback for hosts where postBuild didn't run.
            subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                check=True,
            )
            browser = p.chromium.launch(
                headless=True,
                chromium_sandbox=False,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                ],
            )
        context = browser.new_context(
            http_credentials={"username": username, "password": password}
        )
        page = context.new_page()
        try:
            page.goto(MIS_URL, timeout=60000)
            report_link = page.locator("a[title='Attendance Log Report']")
            report_link.wait_for(state="attached", timeout=10000)
            report_link.evaluate("node => node.click()")
            page.locator("select[title='EmployeeID']").wait_for(
                state="visible", timeout=15000
            )
            page.locator("select[title='EmployeeID']").evaluate(
                f"(s)=>{{for(i=0;i<s.options.length;i++){{if(s.options[i].text.includes('{emp_id}')){{s.selectedIndex=i;s.dispatchEvent(new Event('change',{{bubbles:true}}));}}}}}}"
            )
            today_dt = datetime.now()
            f_day = today_dt.replace(day=1).strftime("%d-%b-%Y")
            l_day = today_dt.strftime("%d-%b-%Y")
            page.locator("#DMNDateDateRangeControl4392_FromDateCalender_DTB").fill(
                f_day
            )
            page.locator("#DMNDateDateRangeControl4392_ToDateCalender_DTB").fill(l_day)
            page.locator("#ViewReportImageButton").click()
            page.wait_for_load_state("networkidle")

            rows = page.locator("table tr").all()
            for row in rows:
                cells = row.locator("td").all_inner_texts()
                if not cells:
                    continue
                c = [clean_text(cell) for cell in cells]
                if any(str(emp_id) in val for val in c):
                    data.append(
                        {
                            "Date": c[2],
                            "Hours": c[10] if c[10] else c[6],
                            "Status": c[11],
                        }
                    )

            page.goto(MIS_URL)
            today_mins, first_punch, state, proof_table = scrape_live_today(
                page, emp_id
            )
            today_str = today_dt.strftime("%d-%b-%Y")
            if today_mins > 0 or first_punch:
                data = [r for r in data if r["Date"] != today_str]
                data.append(
                    {
                        "Date": today_str,
                        "Hours": format_mins_to_hms(today_mins),
                        "Status": f"Present ({state})",
                    }
                )
                st.session_state.today_mins = today_mins
                st.session_state.first_punch = first_punch
                st.session_state.proof_table = proof_table
            return pd.DataFrame(data)
        except Exception as e:
            st.error(f"Global Sync Error: {e}")
            return pd.DataFrame()
        finally:
            browser.close()


# --- UI ---

st.set_page_config(page_title="MIS Pro Dashboard", layout="wide")

st.markdown(
    """
    <style>
    div.stButton > button:first-child {
        background-color: #4681f4;
        color: white;
        border: none;
        font-weight: bold;
        transition: background-color 0.3s;
    }
    div.stButton > button:first-child:hover {
        background-color: #003399;
        color: white;
        border: none;
    }
    div.stButton > button:first-child:active {
        background-color: #002266;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()
if "today_mins" not in st.session_state:
    st.session_state.today_mins = 0
if "proof_table" not in st.session_state:
    st.session_state.proof_table = []
if "first_punch" not in st.session_state:
    st.session_state.first_punch = None
if "creds" not in st.session_state:
    st.session_state.creds = None

# Sidebar
with st.sidebar:
    st.header("Login")
    with st.form(key="auth_form"):
        e_id = st.text_input("Emp ID", placeholder="Enter employee ID")
        u_name = st.text_input("Username", placeholder="Enter username")
        p_word = st.text_input(
            "Password", placeholder="Enter password", type="password"
        )
        submitted = st.form_submit_button("Calculate")
        if submitted:
            st.session_state.creds = {
                "id": e_id,
                "user": u_name,
                "pass": p_word,
            }

# Main UI Header
header_col, refresh_col = st.columns([0.85, 0.15])
with header_col:
    st.title("Timesheet Dashboard")

# Handle Refresh Logic
with refresh_col:
    st.write("")
    if st.button("Refresh", use_container_width=True):
        if st.session_state.creds:
            with st.spinner("Updating..."):
                st.session_state.df = full_historical_sync(
                    st.session_state.creds["id"],
                    st.session_state.creds["user"],
                    st.session_state.creds["pass"],
                )
                st.rerun()
        else:
            st.warning("Add credentials in the sidebar.")

# Summary Metrics
if not st.session_state.df.empty:
    m1, m2, m3 = st.columns(3)
    mins_needed = max(0, 480 - st.session_state.today_mins)
    with m1:
        fp_str = (
            st.session_state.first_punch.strftime("%I:%M %p")
            if st.session_state.first_punch
            else "N/A"
        )
        st.info(f"First Entry Today: {fp_str}")
    with m2:
        st.metric(
            "Total Worked Today",
            format_mins_to_hms(st.session_state.today_mins),
        )
    with m3:
        if mins_needed > 0:
            st.success(
                f"Logout for 8h: {(datetime.now() + timedelta(minutes=mins_needed)).strftime('%I:%M %p')}"
            )
        else:
            st.success("Daily 8h Goal Met!")

    with st.expander("Calculation Proof (Audit Log)"):
        if st.session_state.proof_table:
            pdf = pd.DataFrame(st.session_state.proof_table)
            st.dataframe(pdf, use_container_width=True, hide_index=True)

if submitted:
    with st.spinner("Syncing..."):
        st.session_state.df = full_historical_sync(e_id, u_name, p_word)
        st.rerun()

# Data Tables
if not st.session_state.df.empty:
    today_dt = datetime.now()
    dates = pd.date_range(start=today_dt.replace(day=1), end=today_dt)
    cal_df = pd.DataFrame(
        {
            "Date": [d.strftime("%d-%b-%Y") for d in dates],
            "Day": [d.strftime("%A") for d in dates],
            "Week": ["Week " + str((d.day - 1) // 7 + 1) for d in dates],
        }
    )
    merged = pd.merge(cal_df, st.session_state.df, on="Date", how="left").fillna("")

    def get_row_total(row):
        mins = parse_hours_to_mins(row["Hours"])
        is_off = any(
            s in str(row["Status"]).lower() for s in ["leave", "holiday", "off"]
        )
        return format_mins_to_hms(mins + (480 if is_off else 0))

    merged["Total Hrs."] = merged.apply(get_row_total, axis=1)

    def style_table(s):
        styles = ["color: black; border: 1px solid #d3d3d3; text-align: center;"] * len(
            s
        )
        if s.name == "Hours":
            return [st + "background-color: #CFE2F3;" for st in styles]
        if s.name == "Total Hrs.":
            return [
                st + "background-color: #D9EAD3; font-weight: bold;" for st in styles
            ]
        if s.name in ["Day", "Date"]:
            return [
                (
                    st + "background-color: #6D9EEB; color: white;"
                    if i % 2 == 0
                    else st + "background-color: #A4C2F4;"
                )
                for i, st in enumerate(styles)
            ]
        return styles

    st.subheader("Monthly Log")
    st.dataframe(
        merged[["Week", "Day", "Date", "Hours", "Total Hrs."]]
        .style.set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#3C78D8"),
                        ("color", "white"),
                    ],
                }
            ]
        )
        .apply(style_table, axis=0),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    st.subheader("Weekly Tracker (24h Goal)")
    merged["_mins"] = merged["Hours"].apply(parse_hours_to_mins)
    weekly = merged.groupby("Week")["_mins"].sum().reset_index()
    weekly["Current Weekly"] = weekly["_mins"].apply(format_mins_to_hms)
    weekly["Deficit/Extra"] = (weekly["_mins"] - (24 * 60)).apply(format_mins_to_hms)
    st.table(weekly[["Week", "Current Weekly", "Deficit/Extra"]])
else:
    st.info("Add credentials")

import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


# =========================================================
# PAGE CONFIGURATION
# =========================================================
st.set_page_config(
    page_title="Examination Hall Finder",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =========================================================
# FILE PATHS
# =========================================================
EXCEL_FILE = Path("hallplan.xlsx")
LOGO_FILE = Path("logo.png")


# =========================================================
# INDIA STANDARD TIME
# =========================================================
IST = ZoneInfo("Asia/Kolkata")


def get_ist_now():
    return datetime.now(IST)


# =========================================================
# REQUIRED EXCEL COLUMNS
# =========================================================
REQUIRED_COLUMNS = {
    "Register Number",
    "Exam Date",
    "Name",
    "Programme",
    "Session",
    "Time",
    "Course Code",
    "Course Name",
    "Room No",
    "Floor",
    "Seat Number",
}


# =========================================================
# FLEXIBLE EXAM DATE CONVERSION
# =========================================================
def convert_exam_date(value):

    if pd.isna(value):
        return pd.NaT

    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value).normalize()

    value = re.sub(r"\s+", " ", str(value).strip())

    if not value:
        return pd.NaT

    date_formats = [
        "%d.%m.%Y",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%d %B %Y",
        "%d %b %Y",
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
    ]

    for fmt in date_formats:

        parsed = pd.to_datetime(
            value,
            format=fmt,
            errors="coerce",
        )

        if not pd.isna(parsed):
            return parsed.normalize()

    parsed = pd.to_datetime(
        value,
        errors="coerce",
        dayfirst=True,
    )

    if not pd.isna(parsed):
        return parsed.normalize()

    return pd.NaT


# =========================================================
# LOAD AND PREPARE EXCEL DATA
# =========================================================
@st.cache_data(
    show_spinner="Loading examination details...",
    ttl=3600,
    max_entries=2,
)
def load_data(file_path, file_mtime):

    df = pd.read_excel(
        file_path,
        usecols=lambda column: (
            str(column).strip() in REQUIRED_COLUMNS
        ),
        engine="openpyxl",
    )

    df.columns = (
        df.columns.astype(str).str.strip()
    )

    for col in df.select_dtypes(include=["object"]).columns:

        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    if "Register Number" in df.columns:

        df["Register Number"] = (
            df["Register Number"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

    if "Exam Date" in df.columns:

        distinct_dates = pd.unique(df["Exam Date"])

        date_lookup = {
            value: convert_exam_date(value)
            for value in distinct_dates
        }

        df["_parsed_exam_date"] = (
            df["Exam Date"].map(date_lookup)
        )

    if "Register Number" in df.columns:

        df = df.set_index(
            "Register Number",
            drop=False,
        )

    return df


# =========================================================
# CHECK EXCEL FILE
# =========================================================
if not EXCEL_FILE.exists():

    st.error(
        f"Excel file '{EXCEL_FILE.name}' was not found."
    )

    st.info(
        "Place hallplan.xlsx in the same folder as app.py."
    )

    st.stop()


# =========================================================
# LOAD EXCEL DATA
# =========================================================
try:

    df = load_data(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime_ns,
    )

except Exception as error:

    st.error("Unable to load the examination Excel file.")

    st.exception(error)

    st.stop()


# =========================================================
# VALIDATE REQUIRED COLUMNS
# =========================================================
missing_columns = (
    {"Register Number", "Exam Date"} - set(df.columns)
)

if missing_columns:

    st.error(
        "Missing required Excel columns: "
        + ", ".join(sorted(missing_columns))
    )

    st.write("Available columns:", list(df.columns))

    st.stop()


# =========================================================
# CUSTOM CSS - MOBILE FRIENDLY
# =========================================================
st.markdown(
    """
    <style>

    .main-title {
        text-align: center;
        font-size: 36px;
        font-weight: bold;
        margin-top: 10px;
        margin-bottom: 5px;
    }

    .sub-title {
        text-align: center;
        font-size: 20px;
        margin-bottom: 25px;
    }

    .search-title {
        font-size: 22px;
        font-weight: bold;
        margin-bottom: 10px;
    }

    .student-name {
        text-align: center;
        font-size: 28px;
        font-weight: bold;
        margin-top: 15px;
    }

    .reg-number {
        text-align: center;
        font-size: 19px;
        margin-bottom: 20px;
    }

    .current-date {
        text-align: center;
        font-size: 20px;
        font-weight: bold;
        margin-top: 10px;
        margin-bottom: 20px;
    }

    .result-title {
        text-align: center;
        font-size: 26px;
        font-weight: bold;
        margin-top: 25px;
        margin-bottom: 15px;
    }

    .session-count {
        text-align: center;
        font-size: 18px;
        margin-bottom: 15px;
    }

    .no-exam {
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin-top: 30px;
        margin-bottom: 30px;
    }

    .footer {
        text-align: center;
        font-size: 16px;
        margin-top: 20px;
    }

    [data-testid="stTable"] {
        width: 100%;
        overflow-x: auto;
    }

    @media (max-width: 768px) {

        .main-title {
            font-size: 25px;
        }

        .sub-title {
            font-size: 16px;
            margin-bottom: 18px;
        }

        .search-title {
            font-size: 19px;
        }

        .student-name {
            font-size: 23px;
        }

        .reg-number {
            font-size: 16px;
        }

        .current-date {
            font-size: 16px;
        }

        .result-title {
            font-size: 21px;
        }

        .session-count {
            font-size: 16px;
        }

        .no-exam {
            font-size: 18px;
        }

    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# HEADER - LOGO
# =========================================================
if LOGO_FILE.exists():

    logo_col1, logo_col2, logo_col3 = st.columns(
        [1, 2, 1]
    )

    with logo_col2:

        st.image(
            str(LOGO_FILE),
            width=280,
        )

else:

    st.warning(
        "logo.png was not found. "
        "Place logo.png in the same folder as app.py."
    )


# =========================================================
# CURRENT DATE AND TIME
# =========================================================
now_ist = get_ist_now()

today = pd.Timestamp(now_ist.date())

current_date_ist = now_ist.strftime("%d-%m-%Y")

current_time_ist = now_ist.strftime("%I:%M:%S %p")


# =========================================================
# MAIN TITLE
# =========================================================
st.markdown(
    """
    <div class="main-title">
        CIA EXAMINATION - HALL PLAN
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# CURRENT DATE AND TIME DISPLAY
# =========================================================
st.markdown(
    f"""
    <div class="current-date">
        Today's Examination Date:
        <b>{current_date_ist}</b>
        <br>
        Current Time (IST):
        <b>{current_time_ist}</b>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SUBTITLE
# =========================================================
st.markdown(
    """
    <div class="sub-title">
        Enter your Register Number to find
        today's examination hall
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# REGISTER NUMBER SEARCH
# =========================================================
st.markdown(
    """
    <div class="search-title">
        Enter Register Number
    </div>
    """,
    unsafe_allow_html=True,
)


register_number = st.text_input(
    "Register Number",
    placeholder="Example: 228004002",
    label_visibility="collapsed",
    max_chars=20,
)


# =========================================================
# SEARCH PROCESS
# =========================================================
if register_number:

    register_number = register_number.strip()

    # -----------------------------------------------------
    # VALIDATE REGISTER NUMBER
    # -----------------------------------------------------
    if not re.fullmatch(r"\d+", register_number):

        st.error(
            "Please enter a valid Register Number "
            "using numbers only."
        )

        st.stop()

    # -----------------------------------------------------
    # FIND STUDENT RECORDS
    # -----------------------------------------------------
    try:

        student_records = df.loc[[register_number]]

    except KeyError:

        student_records = df.iloc[0:0]

    # -----------------------------------------------------
    # STUDENT NOT FOUND
    # -----------------------------------------------------
    if student_records.empty:

        st.error(
            f"Register Number {register_number} was not found."
        )

        st.info(
            "Please check the Register Number and try again."
        )

    else:

        # -------------------------------------------------
        # STUDENT NAME
        # -------------------------------------------------
        student_name = (
            str(student_records.iloc[0].get("Name", "")).strip()
        )

        # -------------------------------------------------
        # PROGRAMME
        # -------------------------------------------------
        programme = (
            str(student_records.iloc[0].get("Programme", "")).strip()
        )

        # -------------------------------------------------
        # DISPLAY STUDENT NAME
        # -------------------------------------------------
        if student_name:

            st.markdown(
                f"""
                <div class="student-name">
                    {student_name}
                </div>
                """,
                unsafe_allow_html=True,
            )

        # -------------------------------------------------
        # DISPLAY REGISTER NUMBER
        # -------------------------------------------------
        st.markdown(
            f"""
            <div class="reg-number">
                Register Number:
                <b>{register_number}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # -------------------------------------------------
        # DISPLAY PROGRAMME
        # -------------------------------------------------
        if programme:

            st.info(f"Programme: {programme}")

        # -------------------------------------------------
        # FILTER TODAY'S EXAMINATIONS
        # -------------------------------------------------
        result = student_records.loc[
            student_records["_parsed_exam_date"] == today
        ]

        # -------------------------------------------------
        # EXAMINATION FOUND
        # -------------------------------------------------
        if not result.empty:

            session_count = len(result)

            st.markdown(
                f"""
                <div class="session-count">
                    <b>{session_count}</b>
                    examination session(s) today
                </div>
                """,
                unsafe_allow_html=True,
            )

            # ---------------------------------------------
            # DISPLAY COLUMNS
            # ---------------------------------------------
            display_columns = [
                "Exam Date",
                "Session",
                "Time",
                "Course Code",
                "Course Name",
                "Room No",
                "Floor",
                "Seat Number",
            ]

            available_columns = [
                col
                for col in display_columns
                if col in result.columns
            ]

            display_result = result[
                available_columns
            ].copy()

            # ---------------------------------------------
            # FORMAT EXAM DATE
            # ---------------------------------------------
            if "Exam Date" in display_result.columns:

                display_result["Exam Date"] = (
                    result["_parsed_exam_date"]
                    .dt.strftime("%d-%m-%Y")
                    .fillna("")
                    .to_numpy()
                )

            # ---------------------------------------------
            # RENAME COLUMNS
            # ---------------------------------------------
            display_result = display_result.rename(
                columns={
                    "Room No": "Hall No",
                    "Seat Number": "Seat No",
                }
            )

            # ---------------------------------------------
            # RESULT TITLE
            # ---------------------------------------------
            st.markdown(
                """
                <div class="result-title">
                    TODAY'S EXAMINATION DETAILS
                </div>
                """,
                unsafe_allow_html=True,
            )

            # ---------------------------------------------
            # DISPLAY RESULT
            # ---------------------------------------------
            st.table(
                display_result.reset_index(drop=True)
            )

        # -------------------------------------------------
        # NO EXAMINATION TODAY
        # -------------------------------------------------
        else:

            st.markdown(
                """
                <div class="no-exam">
                    📅 No examination is scheduled for you today.
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.info(
                f"Your Register Number exists in the hall plan, "
                f"but there is no examination scheduled for "
                f"{current_date_ist}."
            )


# =========================================================
# FOOTER
# =========================================================
st.markdown("---")

st.markdown(
    """
    <div class="footer">
        Examination Hall Information System
    </div>
    """,
    unsafe_allow_html=True,
)

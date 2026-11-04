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
)


# =========================================================
# EXCEL FILE
# =========================================================
EXCEL_FILE = Path("hallplan.xlsx")

# Logo file
LOGO_FILE = Path("logo.png")


# =========================================================
# INDIA STANDARD TIME
# =========================================================
IST = ZoneInfo("Asia/Kolkata")

now_ist = datetime.now(IST)

# Current date according to IST
today = pd.Timestamp(now_ist.date())

# Current time according to IST
current_time_ist = now_ist.strftime("%I:%M:%S %p")

# Current date for display
current_date_ist = now_ist.strftime("%d-%m-%Y")


# =========================================================
# CUSTOM CSS
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

    @media (max-width: 768px) {
        .main-title { font-size: 25px; }
        .sub-title { font-size: 16px; margin-bottom: 18px; }
        .search-title { font-size: 19px; }
        .student-name { font-size: 23px; }
        .reg-number { font-size: 16px; }
        .current-date { font-size: 16px; }
        .result-title { font-size: 21px; }
        .session-count { font-size: 16px; }
        .no-exam { font-size: 18px; }
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FLEXIBLE EXAM DATE CONVERSION
# =========================================================
def convert_exam_date(value):

    # -----------------------------------------------------
    # Empty value
    # -----------------------------------------------------
    if pd.isna(value):
        return pd.NaT

    # -----------------------------------------------------
    # Already a Pandas Timestamp
    # -----------------------------------------------------
    if isinstance(value, pd.Timestamp):

        return value.normalize()

    # -----------------------------------------------------
    # Python datetime
    # -----------------------------------------------------
    if isinstance(value, datetime):

        return pd.Timestamp(value).normalize()

    # -----------------------------------------------------
    # Convert to string
    # -----------------------------------------------------
    value = str(value).strip()

    if not value:
        return pd.NaT

    # -----------------------------------------------------
    # Remove extra spaces
    # -----------------------------------------------------
    value = re.sub(
        r"\s+",
        " ",
        value
    )


    # =====================================================
    # COMMON DATE FORMATS
    # =====================================================
    date_formats = [

        # DD.MM.YYYY
        "%d.%m.%Y",

        # DD-MM-YYYY
        "%d-%m-%Y",

        # DD/MM/YYYY
        "%d/%m/%Y",

        # DD.MM.YY
        "%d.%m.%y",

        # DD-MM-YY
        "%d-%m-%y",

        # DD/MM/YY
        "%d/%m/%y",

        # DD Month YYYY
        "%d %B %Y",

        # DD Mon YYYY
        "%d %b %Y",

        # DD-Mon-YYYY
        "%d-%b-%Y",

        # DD-Month-YYYY
        "%d-%B-%Y",

        # YYYY-MM-DD
        "%Y-%m-%d",

        # YYYY/MM/DD
        "%Y/%m/%d",

        # YYYY.MM.DD
        "%Y.%m.%d",
    ]


    # =====================================================
    # TRY COMMON FORMATS
    # =====================================================
    for fmt in date_formats:

        try:

            parsed = pd.to_datetime(
                value,
                format=fmt,
                errors="coerce"
            )

            if not pd.isna(parsed):

                return parsed.normalize()

        except Exception:
            pass


    # =====================================================
    # FINAL FALLBACK
    # DD/MM/YYYY ASSUMPTION
    # =====================================================
    try:

        parsed = pd.to_datetime(
            value,
            errors="coerce",
            dayfirst=True
        )

        if not pd.isna(parsed):

            return parsed.normalize()

    except Exception:
        pass


    return pd.NaT


# =========================================================
# LOAD AND PREPARE EXCEL DATA (CACHED)
# =========================================================
# Only these columns are needed by this application. Reading fewer
# columns reduces Excel loading time and memory use on Streamlit Cloud.
REQUIRED_COLUMNS = {
    "Register Number", "Exam Date", "Name", "Programme",
    "Session", "Time", "Course Code", "Course Name",
    "Room No", "Floor", "Seat Number",
}


@st.cache_data(show_spinner=False, ttl=300)
def load_data(file_path, file_mtime):
    """Read the required Excel columns and prepare fast lookups once."""
    df = pd.read_excel(
        file_path,
        usecols=lambda column: str(column).strip() in REQUIRED_COLUMNS,
        engine="openpyxl",
    )

    # Normalize column names and text values once, not on every search.
    df.columns = df.columns.astype(str).str.strip()
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    if "Register Number" in df.columns:
        df["Register Number"] = (
            df["Register Number"].astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

    if "Exam Date" in df.columns:
        # Exam dates repeat across many student rows. Parse each distinct
        # date once, then map it back to all rows.
        distinct_dates = pd.unique(df["Exam Date"])
        date_lookup = {value: convert_exam_date(value) for value in distinct_dates}
        df["_parsed_exam_date"] = df["Exam Date"].map(date_lookup)

    # Index by register number. A list lookup handles students with
    # multiple examination records while avoiding a full-table scan.
    if "Register Number" in df.columns:
        df = df.set_index("Register Number", drop=False)

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
# READ DATA
# =========================================================
# Include the file modification time in the cache key. If hallplan.xlsx is
# replaced on the server, the cached data is refreshed automatically.
df = load_data(str(EXCEL_FILE), EXCEL_FILE.stat().st_mtime)


# =========================================================
# CHECK REQUIRED COLUMNS
# =========================================================
if "Register Number" not in df.columns:

    st.error(
        "The Excel file does not contain "
        "'Register Number' column."
    )

    st.write(
        "Available columns:"
    )

    st.write(
        list(df.columns)
    )

    st.stop()


if "Exam Date" not in df.columns:

    st.error(
        "The Excel file does not contain "
        "'Exam Date' column."
    )

    st.write(
        "Available columns:"
    )

    st.write(
        list(df.columns)
    )

    st.stop()


# =========================================================
# HEADER
# =========================================================

# ---------------------------------------------------------
# CENTER LOGO
# ---------------------------------------------------------
if LOGO_FILE.exists():

    logo_col1, logo_col2, logo_col3 = st.columns(
        [1, 2, 1]
    )

    with logo_col2:

        st.image(
            str(LOGO_FILE),
            width=360
        )

else:

    st.warning(
        "logo.png was not found. "
        "Place logo.png in the same folder as app.py."
    )


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
# CURRENT DATE AND TIME - IST
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
# SUB TITLE
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
# SEARCH TITLE
# =========================================================
st.markdown(
    """
    <div class="search-title">
        Enter Register Number
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# REGISTER NUMBER INPUT
# =========================================================
register_number = st.text_input(
    "Register Number",
    placeholder="Example: 228004002",
    label_visibility="collapsed",
)


# =========================================================
# SEARCH PROCESS
# =========================================================
if register_number:

    register_number = register_number.strip()


    # =====================================================
    # VALIDATE REGISTER NUMBER
    # =====================================================
    if not re.fullmatch(
        r"\d+",
        register_number
    ):

        st.error(
            "Please enter a valid Register Number "
            "using numbers only."
        )

        st.stop()


    # =====================================================
    # FIND STUDENT RECORDS
    # =====================================================
    # Fast indexed lookup; preserves all rows if a student has
    # more than one examination record.
    try:
        student_records = df.loc[[register_number]].copy()
    except KeyError:
        student_records = df.iloc[0:0].copy()


    # =====================================================
    # STUDENT FOUND
    # =====================================================
    if not student_records.empty:


        # =================================================
        # FILTER TODAY'S DATE ONLY (DATES PRE-PROCESSED AT LOAD)
        # =================================================
        result = student_records[
            student_records["_parsed_exam_date"] == today
        ].copy()


        # =================================================
        # STUDENT NAME
        # =================================================
        student_name = ""

        if "Name" in student_records.columns:

            student_name = str(
                student_records.iloc[0]["Name"]
            ).strip()


        # =================================================
        # PROGRAMME
        # =================================================
        programme = ""

        if "Programme" in student_records.columns:

            programme = str(
                student_records.iloc[0]["Programme"]
            ).strip()


        # =================================================
        # DISPLAY STUDENT NAME
        # =================================================
        if student_name:

            st.markdown(
                f"""
                <div class="student-name">
                    {student_name}
                </div>
                """,
                unsafe_allow_html=True,
            )


        # =================================================
        # DISPLAY REGISTER NUMBER
        # =================================================
        st.markdown(
            f"""
            <div class="reg-number">
                Register Number:
                <b>{register_number}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )


        # =================================================
        # DISPLAY PROGRAMME
        # =================================================
        if programme:

            st.info(
                f"Programme: {programme}"
            )


        # =================================================
        # TODAY'S EXAMINATION FOUND
        # =================================================
        if not result.empty:


            # =============================================
            # SESSION COUNT
            # =============================================
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


            # =============================================
            # DISPLAY COLUMNS
            # =============================================
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


            # =============================================
            # KEEP AVAILABLE COLUMNS
            # =============================================
            available_columns = [

                col
                for col in display_columns
                if col in result.columns
            ]


            # =============================================
            # CREATE DISPLAY DATAFRAME
            # =============================================
            display_result = result[
                available_columns
            ].copy()


            # =============================================
            # FORMAT EXAM DATE
            # =============================================
            if "Exam Date" in display_result.columns:

                display_result["Exam Date"] = (
                    result["_parsed_exam_date"]
                    .dt.strftime("%d-%m-%Y")
                    .fillna("")
                    .to_numpy()
                )


            # =============================================
            # RENAME COLUMNS
            # =============================================
            display_result = display_result.rename(
                columns={

                    "Exam Date":
                        "Exam Date",

                    "Session":
                        "Session",

                    "Time":
                        "Time",

                    "Course Code":
                        "Course Code",

                    "Course Name":
                        "Course Name",

                    "Room No":
                        "Hall No",

                    "Floor":
                        "Floor",

                    "Seat Number":
                        "Seat No",
                }
            )


            # =============================================
            # EXAMINATION DETAILS TITLE
            # =============================================
            st.markdown(
                """
                <div class="result-title">
                    TODAY'S EXAMINATION DETAILS
                </div>
                """,
                unsafe_allow_html=True,
            )


            # =============================================
            # DISPLAY TABLE
            # =============================================
            # Static table avoids the heavier interactive DataFrame bundle
            # that was failing to load on Streamlit Cloud and is lighter on mobile.
            st.table(display_result.reset_index(drop=True))


        # =================================================
        # NO EXAM TODAY
        # =================================================
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
                f"Your Register Number exists in the "
                f"hall plan, but there is no examination "
                f"scheduled for {current_date_ist}."
            )


    # =====================================================
    # REGISTER NUMBER NOT FOUND
    # =====================================================
    else:

        st.error(
            f"Register Number "
            f"{register_number} was not found."
        )

        st.info(
            "Please check the Register Number "
            "and try again."
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

# main.py
import streamlit as st
import pandas as pd
from pathlib import Path
import unicodedata

# ---------------- CONFIG ----------------
EXCEL_FILENAME = "Marksheet Nov Sessional 2025.xlsx"

COL_STUDENT_NAME = "Student Name"
COL_ADMISSION_NO = "Admission No."
COL_FATHER_NAME = "Father Name"   # optional

st.set_page_config(page_title="Marksheet Viewer", layout="wide")


# Clean simple converter → always returns text
def to_text_one_decimal(val):
    """
    Convert value to text.
    If it can be converted to float, show 1 decimal place.
    Else return original string.
    """
    try:
        if pd.isna(val):
            return ""
    except:
        pass

    # Normalize value
    s = str(val).strip()
    if s == "":
        return ""

    s = unicodedata.normalize("NFKC", s)

    # Attempt float conversion
    try:
        f = float(s.replace(",", "").replace("%", ""))
        return f"{round(f, 1):.1f}"
    except:
        return s


def get_sheetnames():
    if not Path(EXCEL_FILENAME).exists():
        return None, "Excel file not found."

    try:
        xls = pd.ExcelFile(EXCEL_FILENAME, engine="openpyxl")
        return xls.sheet_names, None
    except Exception as e:
        return None, str(e)


def parse_sheet(sheetname: str):
    """Extract university + semester from a sheet name like 'MGKVP 1'."""
    parts = sheetname.rsplit(" ", 1)
    if len(parts) == 2 and parts[1].isdigit():
        return parts[0], parts[1], sheetname
    return sheetname, "", sheetname


def main():

    # Header image + Result title
    header_path = "header.png"
    st.image(header_path, use_container_width=True)

    st.markdown(
    "<h2 style='text-align:center; font-weight:700;'>Result Sessional Odd Sem 2025</h2>",
    unsafe_allow_html=True    )

    # Load sheet names
    sheets, err = get_sheetnames()
    if err:
        st.error(err)
        return

    parsed = [parse_sheet(s) for s in sheets]
    universities = sorted({u for u, sem, orig in parsed})

    # UI: Select University
    uni = st.selectbox("Select University", ["-- choose --"] + universities)
    if uni == "-- choose --":
        return

    sems = sorted({sem for u, sem, orig in parsed if u == uni})
    sem = st.selectbox("Select Semester", ["-- choose --"] + sems)
    if sem == "-- choose --":
        return

    sheet_name = [orig for u, s, orig in parsed if u == uni and s == sem][0]

    # Load selected sheet
    try:
        df = pd.read_excel(EXCEL_FILENAME, sheet_name=sheet_name, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
    except Exception as e:
        st.error(f"Could not load sheet: {e}")
        return

    # Required columns
    if COL_STUDENT_NAME not in df.columns:
        st.error(f"Column '{COL_STUDENT_NAME}' not found.")
        return
    if COL_ADMISSION_NO not in df.columns:
        st.error(f"Column '{COL_ADMISSION_NO}' not found.")
        return

    students = df[COL_STUDENT_NAME].astype(str).dropna().unique().tolist()

    student = st.selectbox("Select Student", ["-- choose --"] + students)
    if student == "-- choose --":
        return

    row = df[df[COL_STUDENT_NAME].astype(str) == student].iloc[0]

    # Columns to exclude from attributes
    exclude = {COL_STUDENT_NAME, COL_ADMISSION_NO}
    if COL_FATHER_NAME in df.columns:
        exclude.add(COL_FATHER_NAME)

    attributes = [c for c in df.columns if c not in exclude]

    # ---------------- DISPLAY ----------------
    st.subheader("Student Details")
    st.write(f"**{COL_STUDENT_NAME}:** {row[COL_STUDENT_NAME]}")
    st.write(f"**{COL_ADMISSION_NO}:** {row[COL_ADMISSION_NO]}")
    if COL_FATHER_NAME in df.columns:
        st.write(f"**{COL_FATHER_NAME}:** {row[COL_FATHER_NAME]}")

    st.subheader("Attributes / Subjects")

    attr_list = []
    for col in attributes:
        value = row[col]
        formatted = to_text_one_decimal(value)
        attr_list.append({"Attribute / Subject": col, "Value": formatted})

    st.table(pd.DataFrame(attr_list))


if __name__ == "__main__":
    main()

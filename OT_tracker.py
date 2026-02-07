import streamlit as st
from openpyxl import load_workbook
from datetime import date, datetime

FILE_PATH = "OT_Tracker.xlsx"
SHEET_NAME = "OT"

st.title("OT Meal Tracker")

selected_date = st.date_input("Select OT Date", value=date.today())

def _to_date(x):
    """Fast conversion to python date (handles datetime/date/strings)."""
    if x is None:
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    # last resort: parse string
    try:
        return datetime.fromisoformat(str(x)).date()
    except Exception:
        try:
            # fallback (handles formats like 2-Jan-2025 etc.)
            return datetime.strptime(str(x), "%d-%b-%Y").date()
        except Exception:
            return None

@st.cache_data(show_spinner=False)
def get_name_to_col(file_path: str, sheet_name: str):
    """Cache header mapping for speed."""
    wb = load_workbook(file_path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    mapping = {}
    for col in range(2, ws.max_column + 1):
        name = ws.cell(row=1, column=col).value
        if name:
            mapping[str(name).strip()] = col
    wb.close()
    return mapping

# Load workbook (normal mode because we will write later)
wb = load_workbook(FILE_PATH)
ws = wb[SHEET_NAME]

# --- Find row for selected date (same logic, faster loop) ---
row_index = None

# Iterate only the date column values (A2:A...)
# values_only=True avoids cell object overhead
for idx, (cell_val,) in enumerate(
    ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1, values_only=True),
    start=2,
):
    if _to_date(cell_val) == selected_date:
        row_index = idx
        break

if row_index is None:
    st.error("Date not found in Excel sheet")
    wb.close()
    st.stop()

# --- Header mapping (cached) ---
names = get_name_to_col(FILE_PATH, SHEET_NAME)

# --- Find unclaimed people (blank cells) ---
available_people = []
for name, col in names.items():
    v = ws.cell(row=row_index, column=col).value
    if v is None or v == "":
        available_people.append(name)

st.subheader("People who have NOT claimed OT")

if not available_people:
    st.success("All people have already claimed OT for this date")
    wb.close()
    st.stop()

selected_people = st.multiselect("Select people", available_people)

if st.button("Submit OT Claim"):
    if not selected_people:
        st.warning("Please select at least one person")
    else:
        for person in selected_people:
            ws.cell(row=row_index, column=names[person]).value = "OT"

        wb.save(FILE_PATH)
        st.success("OT Tracker Updated")
        wb.close()
        st.stop()

wb.close()

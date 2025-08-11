import io
from datetime import datetime
from typing import List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# Optional engines for Excel
try:
    import openpyxl  # noqa: F401
except Exception:
    pass

# PDF export
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Excel Elements – Function Runner", layout="wide")

st.title("🧪 Excel Elements – Function Runner")

# --------------------------
# Helpers
# --------------------------
def read_uploaded(file) -> pd.DataFrame:
    suffix = file.name.lower().split(".")[-1]
    if suffix in ["xlsx", "xls"]:
        return pd.read_excel(file)
    elif suffix == "csv":
        return pd.read_csv(file)
    else:
        raise ValueError("Unsupported file type. Please upload CSV or Excel.")

def num_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def date_cols(df: pd.DataFrame) -> List[str]:
    out = []
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            out.append(c)
        else:
            # try parse
            try:
                pd.to_datetime(df[c])
                out.append(c)
            except Exception:
                pass
    return list(dict.fromkeys(out))

def text_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if pd.api.types.is_string_dtype(df[c])]

def ensure_datetime_series(s: pd.Series) -> pd.Series:
    if not pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    return s

def excel_like_weekday(dt_series: pd.Series) -> pd.Series:
    # Excel WEEKDAY default return_type=1: Sunday=1 ... Saturday=7
    # pandas weekday(): Monday=0 ... Sunday=6
    # We'll implement from datetime: isoweekday(): Monday=1 ... Sunday=7
    # Map to Excel default: Sunday=1 -> Sunday=7 (ISO) => ((iso % 7) + 1)
    iso = dt_series.dt.isoweekday()  # 1=Mon ... 7=Sun
    return ((iso % 7) + 1)

def to_excel(bytes_dict: dict) -> bytes:
    # bytes_dict: {sheet_name: dataframe or (dataframe, sheet_note:str)}
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet, df_or_tuple in bytes_dict.items():
            if isinstance(df_or_tuple, tuple):
                df, note = df_or_tuple
            else:
                df, note = df_or_tuple, None
            df.to_excel(writer, sheet_name=sheet, index=False)
            if note:
                wb  = writer.book
                ws  = writer.sheets[sheet]
                fmt = wb.add_format({"italic": True, "font_color": "#666666"})
                ws.write(0, len(df.columns)+1, note, fmt)
    return output.getvalue()

def to_pdf(summary_lines: List[str]) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    x, y = 40, height - 50
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x, y, "Excel Elements – Result Summary")
    y -= 20
    c.setFont("Helvetica", 10)
    c.drawString(x, y, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 20
    for line in summary_lines:
        if y < 50:
            c.showPage()
            y = height - 50
            c.setFont("Helvetica", 10)
        c.drawString(x, y, line[:115])  # basic wrap
        y -= 14
    c.showPage()
    c.save()
    return buf.getvalue()

def boolean_series_from_condition(series: pd.Series, op: str, value: float) -> pd.Series:
    if op == ">":   return series > value
    if op == ">=":  return series >= value
    if op == "<":   return series < value
    if op == "<=":  return series <= value
    if op == "==":  return series == value
    if op == "!=":  return series != value
    raise ValueError("Unsupported operator")

# --------------------------
# Sidebar – Upload
# --------------------------
st.sidebar.header("1) Upload Data")
file = st.sidebar.file_uploader("Upload CSV or Excel", type=["csv", "xlsx", "xls"])

if not file:
    st.info("Upload a CSV/Excel file to begin.")
    st.stop()

df = read_uploaded(file)
st.success(f"Loaded: **{file.name}** with **{df.shape[0]}** rows and **{df.shape[1]}** columns.")
st.dataframe(df.head(20), use_container_width=True)

# --------------------------
# Function selection
# --------------------------
st.sidebar.header("2) Pick a Function")
groups = {
    "Basic Math": ["SUM", "AVERAGE", "MAX", "MIN"],
    "Logical": ["IF", "AND", "OR", "NOT"],
    "Lookup & Reference": ["VLOOKUP", "XLOOKUP", "MATCH", "INDEX"],
    "Financial": ["PV", "FV", "IRR", "RATE"],
    "Date & Time": ["TODAY", "WEEKDAY", "DATE"],
    "Statistical": ["STDEV", "MEDIAN", "VAR"],
    "Text": ["TEXT", "CONCAT", "TRIM", "UPPER"],
}
group = st.sidebar.selectbox("Category", list(groups.keys()))
func  = st.sidebar.selectbox("Function", groups[group])

st.sidebar.header("3) Configure Parameters")

result_dict = {}
summary_lines = []
derived_df = df.copy()

numchoices = num_cols(df)
datechoices = date_cols(df)
textchoices = text_cols(df)

# --------------------------
# Function-specific UIs
# --------------------------
if func in ["SUM", "AVERAGE", "MAX", "MIN", "STDEV", "MEDIAN", "VAR"]:
    col = st.sidebar.selectbox("Numeric column", numchoices)
    if not col:
        st.stop()
    series = pd.to_numeric(df[col], errors="coerce").dropna()
    value = None
    if func == "SUM":
        value = series.sum()
    elif func == "AVERAGE":
        value = series.mean()
    elif func == "MAX":
        value = series.max()
    elif func == "MIN":
        value = series.min()
    elif func == "STDEV":
        value = series.std(ddof=1)
    elif func == "MEDIAN":
        value = series.median()
    elif func == "VAR":
        value = series.var(ddof=1)

    result_dict = {"Function": [func], "Column": [col], "Result": [value]}
    summary_lines.append(f"{func} on column '{col}' = {value}")

elif func in ["IF", "AND", "OR", "NOT"]:
    # Build up to two simple conditions on numeric columns
    op_list = [">", ">=", "<", "<=", "==", "!="]

    if func == "IF":
        cond_col = st.sidebar.selectbox("IF: Condition column (numeric)", numchoices)
        cond_op  = st.sidebar.selectbox("IF: Operator", op_list, index=0)
        cond_val = float(st.sidebar.text_input("IF: Compare to value", "0"))
        true_val = st.sidebar.text_input("Value if TRUE (constant or column name)", "1")
        false_val= st.sidebar.text_input("Value if FALSE (constant or column name)", "0")

        cond = boolean_series_from_condition(pd.to_numeric(df[cond_col], errors="coerce"), cond_op, cond_val)

        def resolve_val(token):
            if token in df.columns:
                return df[token]
            else:
                # try numeric
                try:
                    return float(token)
                except Exception:
                    return token  # string literal

        true_series  = resolve_val(true_val)
        false_series = resolve_val(false_val)
        derived_df["IF_result"] = np.where(cond, true_series, false_series)
        result_dict = {"Function": ["IF"], "Details": [f"{cond_col} {cond_op} {cond_val}"], "Preview TRUE/FALSE": [f"{true_val}/{false_val}"]}
        summary_lines.append(f"IF on {cond_col} {cond_op} {cond_val}; created column 'IF_result'.")

    elif func in ["AND", "OR"]:
        c1_col = st.sidebar.selectbox("Cond1 column (numeric)", numchoices)
        c1_op  = st.sidebar.selectbox("Cond1 operator", op_list, index=0, key="op1")
        c1_val = float(st.sidebar.text_input("Cond1 value", "0"))
        c2_col = st.sidebar.selectbox("Cond2 column (numeric)", numchoices, index=min(1, len(numchoices)-1))
        c2_op  = st.sidebar.selectbox("Cond2 operator", op_list, index=0, key="op2")
        c2_val = float(st.sidebar.text_input("Cond2 value", "0", key="v2"))

        c1 = boolean_series_from_condition(pd.to_numeric(df[c1_col], errors="coerce"), c1_op, c1_val)
        c2 = boolean_series_from_condition(pd.to_numeric(df[c2_col], errors="coerce"), c2_op, c2_val)

        if func == "AND":
            derived_df["AND_result"] = c1 & c2
            summary_lines.append(f"AND across ({c1_col} {c1_op} {c1_val}) AND ({c2_col} {c2_op} {c2_val}); new column 'AND_result'.")
        else:
            derived_df["OR_result"] = c1 | c2
            summary_lines.append(f"OR across ({c1_col} {c1_op} {c1_val}) OR ({c2_col} {c2_op} {c2_val}); new column 'OR_result'.")

        result_dict = {"Function": [func], "Details": [summary_lines[-1]]}

    elif func == "NOT":
        c_col = st.sidebar.selectbox("Condition column (boolean)", [c for c in df.columns if pd.api.types.is_bool_dtype(df[c])])
        if not c_col:
            st.warning("Select a boolean column (True/False).")
            st.stop()
        derived_df["NOT_result"] = ~df[c_col]
        summary_lines.append(f"NOT applied to column '{c_col}'; new column 'NOT_result'.")
        result_dict = {"Function": ["NOT"], "Column": [c_col]}

elif func in ["VLOOKUP", "XLOOKUP"]:
    # Simple same-table lookup
    lookup_col = st.sidebar.selectbox("Lookup column", df.columns)
    lookup_val = st.sidebar.text_input("Lookup value (exact match)", "")
    return_col = st.sidebar.selectbox("Return column", df.columns, index=min(1, len(df.columns)-1))
    if lookup_val == "":
        st.stop()
    mask = df[lookup_col].astype(str) == str(lookup_val)
    if mask.any():
        found_value = df.loc[mask, return_col].iloc[0]
        result_dict = {"Function": [func], "Lookup": [f"{lookup_col}={lookup_val}"], "ReturnColumn": [return_col], "Result": [found_value]}
        summary_lines.append(f"{func}: found value in '{return_col}' for {lookup_col}='{lookup_val}' -> {found_value}")
    else:
        result_dict = {"Function": [func], "Lookup": [f"{lookup_col}={lookup_val}"], "ReturnColumn": [return_col], "Result": ["#N/A"]}
        summary_lines.append(f"{func}: no match for {lookup_col}='{lookup_val}' -> #N/A")

elif func == "MATCH":
    match_col = st.sidebar.selectbox("Search in column", df.columns)
    match_val = st.sidebar.text_input("Value to find", "")
    if match_val == "":
        st.stop()
    pos = df[match_col].astype(str).eq(str(match_val)).to_numpy().nonzero()[0]
    if len(pos) > 0:
        result = int(pos[0]) + 1  # Excel 1-based
    else:
        result = "#N/A"
    result_dict = {"Function": ["MATCH"], "Column": [match_col], "Value": [match_val], "Position(1-based)": [result]}
    summary_lines.append(f"MATCH in '{match_col}' for '{match_val}' -> {result}")

elif func == "INDEX":
    idx_row = int(st.sidebar.number_input("Row number (1-based)", min_value=1, value=1))
    idx_col = st.sidebar.selectbox("Column to return", df.columns)
    if 1 <= idx_row <= len(df):
        result = df.iloc[idx_row-1][idx_col]
    else:
        result = "#REF!"
    result_dict = {"Function": ["INDEX"], "Row(1-based)": [idx_row], "Column": [idx_col], "Result": [result]}
    summary_lines.append(f"INDEX at row {idx_row}, column '{idx_col}' -> {result}")

elif func in ["PV", "FV", "IRR", "RATE"]:
    if func in ["PV", "FV", "RATE"]:
        rate = float(st.sidebar.text_input("rate per period (e.g., 0.01)", "0.01"))
        nper = int(st.sidebar.number_input("nper (# periods)", min_value=1, value=12))
        pmt  = float(st.sidebar.text_input("pmt per period (cash outflow negative)", "-1000"))
        pv   = float(st.sidebar.text_input("pv (present value, negative if outflow)", "0"))
        fv   = float(st.sidebar.text_input("fv (future value, positive inflow)", "0"))
        when = st.sidebar.selectbox("when (payment timing)", ["end", "begin"], index=0)
        when_val = 0 if when == "end" else 1

        if func == "PV":
            # numpy_financial alternative w/o dependency
            if rate == 0:
                res = -(fv + pmt * nper)
            else:
                res = -(fv + pmt * (1 + rate * when_val) * ( (1 + rate)**nper - 1) / rate) / ((1 + rate)**nper)
            result_dict = {"Function":["PV"], "Result":[res]}
            summary_lines.append(f"PV computed -> {res}")

        elif func == "FV":
            if rate == 0:
                res = -(pv + pmt * nper)
            else:
                res = -(pv*(1+rate)**nper + pmt*(1+rate*when_val)*(((1+rate)**nper - 1)/rate))
            result_dict = {"Function":["FV"], "Result":[res]}
            summary_lines.append(f"FV computed -> {res}")

        elif func == "RATE":
            # Simple Newton method to approximate rate
            guess = float(st.sidebar.text_input("Initial guess", "0.01"))
            def f(r):
                return pv*(1+r)**nper + pmt*(1+r*when_val)*(((1+r)**nper - 1)/r if r!=0 else nper) + fv
            def df(r):
                if r == 0:
                    return pv*nper + pmt*(1+when_val)*nper - 0  # rough derivative near 0
                A = (1+r)**(nper-1)
                return pv*nper*(1+r)**(nper-1) + pmt*(1+r*when_val)*( ((1+r)**nper - 1)/r**2 - (nper*(1+r)**(nper-1))/r )
            r = guess
            for _ in range(50):
                try:
                    r_new = r - f(r)/df(r)
                except Exception:
                    break
                if not np.isfinite(r_new):
                    break
                if abs(r_new - r) < 1e-8:
                    r = r_new
                    break
                r = r_new
            res = r
            result_dict = {"Function":["RATE"], "Approx rate":[res]}
            summary_lines.append(f"RATE approximate -> {res}")

    elif func == "IRR":
        cf_col = st.sidebar.selectbox("Cash flow column (includes initial outflow)", numchoices)
        cfs = pd.to_numeric(df[cf_col], errors="coerce").dropna().values
        # IRR using np.irr fallback
        try:
            irr = np.irr(cfs)  # deprecated in newer numpy, works in many builds
        except Exception:
            # Simple Newton on NPV
            def npv(r):
                return sum(c / ((1+r)**i) for i, c in enumerate(cfs))
            r = 0.1
            for _ in range(200):
                # derivative of NPV
                d = sum(-i*c / ((1+r)**(i+1)) for i, c in enumerate(cfs) if i>0)
                if abs(d) < 1e-12:
                    break
                r_new = r - npv(r)/d
                if not np.isfinite(r_new):
                    break
                if abs(r_new - r) < 1e-8:
                    r = r_new
                    break
                r = r_new
            irr = r
        result_dict = {"Function":["IRR"], "Column":[cf_col], "IRR":[irr]}
        summary_lines.append(f"IRR on '{cf_col}' -> {irr}")

elif func in ["TODAY", "WEEKDAY", "DATE"]:
    if func == "TODAY":
        today = pd.Timestamp.today().normalize()
        result_dict = {"Function":["TODAY"], "Result":[str(today.date())]}
        summary_lines.append(f"TODAY -> {today.date()}")

    elif func == "WEEKDAY":
        dcol = st.sidebar.selectbox("Date column", datechoices)
        if not dcol:
            st.stop()
        dt = ensure_datetime_series(df[dcol])
        derived_df["WEEKDAY"] = excel_like_weekday(dt)
        result_dict = {"Function":["WEEKDAY"], "Column":[dcol], "Note":["Sunday=1 ... Saturday=7"]}
        summary_lines.append(f"WEEKDAY created for '{dcol}' as new column 'WEEKDAY' (Sun=1 ... Sat=7).")

    elif func == "DATE":
        ycol = st.sidebar.selectbox("Year column", df.columns)
        mcol = st.sidebar.selectbox("Month column", df.columns)
        dcol = st.sidebar.selectbox("Day column", df.columns)
        y = pd.to_numeric(df[ycol], errors="coerce")
        m = pd.to_numeric(df[mcol], errors="coerce")
        d = pd.to_numeric(df[dcol], errors="coerce")
        derived_df["DATE"] = pd.to_datetime(dict(year=y, month=m, day=d), errors="coerce")
        result_dict = {"Function":["DATE"], "Columns":[f"{ycol},{mcol},{dcol}"], "Note":["New column 'DATE' created."]}
        summary_lines.append(f"DATE constructed from {ycol},{mcol},{dcol} -> new 'DATE' column.")

elif func in ["TEXT", "CONCAT", "TRIM", "UPPER"]:
    if func == "TEXT":
        ncol = st.sidebar.selectbox("Numeric column", numchoices)
        fmt  = st.sidebar.text_input("Excel-like format (e.g. #,##0.00)", "#,##0.00")
        # Very basic formatting via pandas map
        derived_df["TEXT"] = df[ncol].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
        result_dict = {"Function":["TEXT"], "Column":[ncol], "Format":[fmt], "Note":["Created 'TEXT' (approx formatting)."]}
        summary_lines.append(f"TEXT formatting approx on '{ncol}' -> new 'TEXT' column.")
    elif func == "CONCAT":
        cols = st.sidebar.multiselect("Columns to concatenate", df.columns)
        sep  = st.sidebar.text_input("Separator", "")
        derived_df["CONCAT"] = df[cols].astype(str).agg(sep.join, axis=1) if cols else ""
        result_dict = {"Function":["CONCAT"], "Columns":[",".join(cols)], "Note":["New 'CONCAT' column created."]}
        summary_lines.append(f"CONCAT over {cols} -> new 'CONCAT' column.")
    elif func == "TRIM":
        tcol = st.sidebar.selectbox("Text column", textchoices)
        derived_df["TRIM"] = df[tcol].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
        result_dict = {"Function":["TRIM"], "Column":[tcol], "Note":["New 'TRIM' column created."]}
        summary_lines.append(f"TRIM on '{tcol}' -> new 'TRIM' column.")
    elif func == "UPPER":
        tcol = st.sidebar.selectbox("Text column", textchoices)
        derived_df["UPPER"] = df[tcol].astype(str).str.upper()
        result_dict = {"Function":["UPPER"], "Column":[tcol], "Note":["New 'UPPER' column created."]}
        summary_lines.append(f"UPPER on '{tcol}' -> new 'UPPER' column.")

# --------------------------
# Results + Downloads
# --------------------------
st.subheader("Result Summary")
if result_dict:
    result_df = pd.DataFrame(result_dict)
    st.dataframe(result_df, use_container_width=True)
else:
    st.info("Configure parameters to compute a result.")
    st.stop()

st.subheader("Derived Data Preview")
st.dataframe(derived_df.head(50), use_container_width=True)

# Build Excel bytes
excel_bytes = to_excel({
    "Data": df,
    "Derived": derived_df,
    "Results": (result_df, f"Function: {func}")
})

# Build PDF bytes
pdf_bytes = to_pdf(
    [f"Function: {func}"] +
    summary_lines +
    [f"Rows: {len(df)}; Columns: {len(df.columns)}"]
)

col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "⬇️ Download Excel",
        data=excel_bytes,
        file_name=f"elements_result_{func.lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
with col2:
    st.download_button(
        "⬇️ Download PDF",
        data=pdf_bytes,
        file_name=f"elements_result_{func.lower()}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

st.caption("Note: Some Excel behaviors are approximated (e.g., TEXT format codes, RATE solver).")

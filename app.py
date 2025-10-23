import streamlit as st
import pandas as pd
from validators.rules_engine import validate_workbook

st.set_page_config(page_title="Machlab Check", layout="wide")

st.title("Machlab – בדיקת דוחות אחסנה (Demo)")
st.caption("מעלים קובץ Excel, מציגים את הגיליונות ובשלבים הבאים נוסיף ולידציות לפי rules.yaml.")

uploaded = st.file_uploader("בחר/י קובץ Excel", type=["xlsx", "xls"])

tab1, tab2 = st.tabs(["תצוגה", "תוצאות בדיקה"])

if uploaded:
    try:
        xl = pd.ExcelFile(uploaded)
        with tab1:
            st.subheader("גיליון/ות בקובץ")
            for sheet in xl.sheet_names:
                st.write(f"**{sheet}**")
                df = xl.parse(sheet)
                st.dataframe(df, use_container_width=True, hide_index=True)

        with tab2:
            st.subheader("תוצאות בדיקה (placeholder)")
            report = validate_workbook(None, "rules/rules.yaml")
            st.json(report)

        st.success("טעינה הושלמה ✔️")
    except Exception as e:
        st.error("שגיאה בטעינת הקובץ")
        st.exception(e)
else:
    st.info("נא להעלות קובץ Excel כדי להתחיל.")

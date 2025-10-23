
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Machlab – נוסחאות מוגבלות לפי 'הזמנות לבדיקה'", layout="wide")

# ----- RTL UI -----
st.markdown("""
<style>
html, body, [class*="css"]  { direction: rtl; text-align: right; }
</style>
""", unsafe_allow_html=True)

st.title("Machlab – גרסה מאוחדת (RTL)")
st.caption("מעלה שני קבצים (פעילות + 'הובלות אלכל כללי'), מעתיק גיליון פנימי, מוסיף עמודות חסרות ומזריק נוסחאות. "
           "VLOOKUP משתמש ב-MATCH דינמי לעמודה 'סופק'. כל ההזרקות נעצרות בשורה האחרונה של 'הזמנות לבדיקה'.")

# ----- Constants -----
REQUIRED_MAIN_SHEET     = "הובלה לבית לקוח"
REQUIRED_INTERNAL_SHEET = "הובלות אלכל כללי"

COL_PURCHASE_SRC  = "הז. רכש (לקוח)"
COL_RAKHASH       = "רכש"
COL_MAKAT         = "מק'ט"
COL_MAKAT_CLEAN   = "מקט ללא פגומים"
COL_ORDER_CHECK   = "הזמנות לבדיקה"
COL_QTY           = "כמות"
COL_QTY_CHECK     = "בדיקת כמות"
COL_PRICE_AFTER   = "מחירון מחלב לאחר בדיקה"
COL_DUP_JULY      = "כפילויות חודש קודם"
COL_MANUAL        = "מעבר ידני"
COL_APPROVAL      = "אישור סופי"
COL_NOTES         = "הערות"
COL_TOTAL_PAY     = "סה\"כ לתשלום"
COL_DIFF_ROW      = "פער לפי שורה"
COL_TOTAL         = "סה\"כ"  # עמודת סיכום קיימת אם יש

# עמודות נוספות להוספה אם חסרות (ללא נוסחאות, מלבד אלה שמוגדר להן נוסחה)
EXTRA_COLUMNS = [
    COL_PRICE_AFTER,
    COL_DUP_JULY,
    COL_QTY_CHECK,        # עם נוסחת VLOOKUP+MATCH
    COL_MANUAL,
    COL_APPROVAL,         # IF/OR/AND
    COL_NOTES,
    COL_TOTAL_PAY,        # IF(...*ABS(...))
    COL_DIFF_ROW,         # סה"כ - סה"כ לתשלום
]

# ----- Helpers -----
def find_sheet_name(sheetnames, target):
    if target in sheetnames:
        return target
    if (target + " ") in sheetnames:
        return target + " "
    for s in sheetnames:
        if s.strip() == target.strip():
            return s
    return None

def find_col(ws, name):
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
    #    if (val or "").strip().lower() == name.strip().lower():   # אפשר להחליף להתאמה לא רגישה לרישיות
        if (val or "").strip() == name:
            return c
    return None

def ensure_column(ws, name):
    """Locate a column by header; if missing, append it at the end and return the index."""
    idx = find_col(ws, name)
    if idx is not None:
        return idx, False
    new_idx = ws.max_column + 1 if ws.max_column else 1
    ws.cell(row=1, column=new_idx, value=name)
    return new_idx, True

def copy_dataframe_to_sheet(df: pd.DataFrame, wb: Workbook, sheet_name: str):
    # overwrite if exists
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    # RTL view only (do not reverse column order)
    try:
        ws.sheet_view.rightToLeft = True
    except Exception:
        pass
    # headers
    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=1, column=j, value=str(col))
    # rows
    for i, row in enumerate(df.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=val)

def last_nonempty_row(ws, col_letter, start_row=2):
    max_row = ws.max_row
    for r in range(max_row, start_row - 1, -1):
        v = ws[f"{col_letter}{r}"].value
        if v is not None and str(v).strip() != "":
            return r
    return start_row - 1

# ----- UI -----
col1, col2 = st.columns(2)
with col1:
    activity_file = st.file_uploader("קובץ פעילות (Excel) – בדרך כלל 'פעילות אלכל חודש שנה' (.xlsx)", type=["xlsx"], key="activity")
with col2:
    internal_file = st.file_uploader("קובץ 'הובלות אלכל כללי' (Excel) (.xlsx)", type=["xlsx"], key="internal")

if activity_file and internal_file:
    try:
        # 1) Load activity workbook & locate main sheet
        wb = load_workbook(activity_file)
        main_sheet_name = find_sheet_name(wb.sheetnames, REQUIRED_MAIN_SHEET)
        if not main_sheet_name:
            st.error(f"לא נמצא גיליון בשם '{REQUIRED_MAIN_SHEET}' בקובץ הפעילות.")
            st.stop()
        ws = wb[main_sheet_name]

        # 2) Copy internal sheet as a new sheet in the same workbook
        xl_internal = pd.ExcelFile(internal_file)
        internal_sheet_name = find_sheet_name(xl_internal.sheet_names, REQUIRED_INTERNAL_SHEET)
        if not internal_sheet_name:
            st.error(f"לא נמצא גיליון בשם '{REQUIRED_INTERNAL_SHEET}' בקובץ 'הובלות אלכל כללי'.")
            st.stop()
        df_internal = xl_internal.parse(internal_sheet_name)
        copy_dataframe_to_sheet(df_internal, wb, REQUIRED_INTERNAL_SHEET)

        # 3) Ensure required and extra columns exist (add at end if missing)
        col_purchase_src = ensure_column(ws, COL_PURCHASE_SRC)[0]
        col_rakhash, _   = ensure_column(ws, COL_RAKHASH)
        col_makat, _     = ensure_column(ws, COL_MAKAT)
        col_clean, _     = ensure_column(ws, COL_MAKAT_CLEAN)
        col_order, _     = ensure_column(ws, COL_ORDER_CHECK)
        col_qty, _       = ensure_column(ws, COL_QTY)
        col_qty_check,_  = ensure_column(ws, COL_QTY_CHECK)
        col_price,_      = ensure_column(ws, COL_PRICE_AFTER)
        col_dup,_        = ensure_column(ws, COL_DUP_JULY)
        col_manual,_     = ensure_column(ws, COL_MANUAL)
        col_approval,_   = ensure_column(ws, COL_APPROVAL)
        col_notes,_      = ensure_column(ws, COL_NOTES)
        col_total_pay,_  = ensure_column(ws, COL_TOTAL_PAY)
        col_diff,_       = ensure_column(ws, COL_DIFF_ROW)

        added_extra = []
        for name in EXTRA_COLUMNS:
            _, added = ensure_column(ws, name)
            if added:
                added_extra.append(name)

        # locate existing "סה"כ" column (do NOT create if missing)
        col_total = find_col(ws, COL_TOTAL)
        if col_total is None:
            st.warning('לא נמצאה עמודה "סה\"כ" בגיליון. הזרקה ל"פער לפי שורה" תדלג (נדרש מקור חיסור).')

        # Letters
        L_src      = get_column_letter(col_purchase_src)
        L_rakhash  = get_column_letter(col_rakhash)
        L_makat    = get_column_letter(col_makat)
        L_clean    = get_column_letter(col_clean)
        L_order    = get_column_letter(col_order)
        L_qty      = get_column_letter(col_qty)
        L_qtychk   = get_column_letter(col_qty_check)
        L_price    = get_column_letter(col_price)
        L_manual   = get_column_letter(col_manual)
        L_approval = get_column_letter(col_approval)
        L_totalpay = get_column_letter(col_total_pay)
        L_diffcol  = get_column_letter(col_diff)
        L_totalcol = get_column_letter(col_total) if col_total else None

        max_row = ws.max_row
        cnt_rakhash = cnt_clean = cnt_order = cnt_qtychk = 0
        cnt_approval = cnt_totalpay = cnt_diff = 0

        # 4) Base formulas (full scan to create order key etc. like before)
        for r in range(2, max_row + 1):
            src_val = ws[f"{L_src}{r}"].value
            if src_val is not None and str(src_val).strip() != "":
                ws[f"{L_rakhash}{r}"] = f"={L_src}{r}*1"
                cnt_rakhash += 1

        for r in range(2, max_row + 1):
            v = ws[f"{L_makat}{r}"].value
            if v is not None and str(v).strip() != "":
                ws[f"{L_clean}{r}"] = f"=LEFT({L_makat}{r},7)"
                cnt_clean += 1

        for r in range(2, max_row + 1):
            v1 = ws[f"{L_rakhash}{r}"].value
            v2 = ws[f"{L_clean}{r}"].value
            if v1 is not None and v2 is not None and str(v1).strip() != "" and str(v2).strip() != "":
                ws[f"{L_order}{r}"] = f"={L_rakhash}{r}&{L_clean}{r}"
                cnt_order += 1

        # stop line determined by last non-empty in "הזמנות לבדיקה"
        end_row_orders = last_nonempty_row(ws, L_order, start_row=2)

        # 4.4 בדיקת כמות – VLOOKUP עם MATCH("סופק") על שורת הכותרות של הגיליון שהועתק
        for r in range(2, end_row_orders + 1):
            order_val = ws[f"{L_order}{r}"].value
            if order_val is not None and str(order_val).strip() != "":
                formula_qty = (
                    '=IFNA('
                    'IF(VLOOKUP({order},\'{internal}\'!$A:$O,'
                    'MATCH("סופק",\'{internal}\'!$A$1:$O$1,0),0)='
                    '\'{main}\'!{qty},'
                    'IF(\'{main}\'!{qty}<3,"תקין","בדיקת כמות"),'
                    '"נדרשת בדיקה"),'
                    '"נדרשת בדיקה")'
                ).format(
                    order=f"{L_order}{r}",
                    internal=REQUIRED_INTERNAL_SHEET,
                    main=main_sheet_name,
                    qty=f"{L_qty}{r}",
                )
                ws[f"{L_qtychk}{r}"] = formula_qty
                cnt_qtychk += 1

        # 4.5 אישור סופי
        for r in range(2, end_row_orders + 1):
            formula_approval = (
                '=IF(OR(AND({manual}{r}=0,{qtychk}{r}="תקין"),{manual}{r}="מאושר"),"מאושר","לא מאושר")'
            ).format(manual=L_manual, qtychk=L_qtychk, r=r)
            ws[f"{L_approval}{r}"] = formula_approval
            cnt_approval += 1

        # 4.6 סה"כ לתשלום
        for r in range(2, end_row_orders + 1):
            formula_total = (
                '=IF({approval}{r}="מאושר",{price}{r}*ABS({qty}{r}),0)'
            ).format(approval=L_approval, price=L_price, qty=L_qty, r=r)
            ws[f"{L_totalpay}{r}"] = formula_total
            cnt_totalpay += 1

        # 4.7 פער לפי שורה = סה"כ - סה"כ לתשלום
        if L_totalcol:
            for r in range(2, end_row_orders + 1):
                ws[f"{L_diffcol}{r}"] = f"={L_totalcol}{r}-{L_totalpay}{r}"
                cnt_diff += 1

        # 5) Save and offer download
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        added_msg = (" | נוספו עמודות: " + ", ".join(added_extra)) if added_extra else ""
        st.success(
            "✅ הועתק גיליון '{0}'. הזרקות עד שורה {1}. נוסחאות – רכש ({2}), מק\"ט ללא פגומים ({3}), הזמנות לבדיקה ({4}), בדיקת כמות/MATCH ({5}), אישור סופי ({6}), סה\"כ לתשלום ({7}), פער לפי שורה ({8}).{9}"
            .format(REQUIRED_INTERNAL_SHEET, end_row_orders, cnt_rakhash, cnt_clean, cnt_order, cnt_qtychk, cnt_approval, cnt_totalpay, cnt_diff, added_msg)
        )
        st.download_button(
            "📥 הורד/י קובץ פעילות מעודכן (VLOOKUP עם MATCH על 'סופק' + כל ההזרקות)",
            data=out,
            file_name="Activity_Unified_All_Formulas_With_ROWDIFF_and_MATCH.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("שגיאה בעיבוד הקבצים.")
        st.exception(e)
else:
    st.info("נא להעלות את שני הקבצים (.xlsx)." )

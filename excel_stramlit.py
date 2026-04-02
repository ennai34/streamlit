import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Excel Cleaner", layout="wide")
st.title("📊 Excel Data Cleaner (แยกจังหวัด + วันที่)")

# =========================
# 🧠 ฟังก์ชันแปลงวันที่ไทย
# =========================
def parse_thai_date(text):
    if not isinstance(text, str):
        return pd.NaT

    text = text.strip()

    months = {
        "ม.ค.": "Jan", "ก.พ.": "Feb", "มี.ค.": "Mar", "เม.ย.": "Apr",
        "พ.ค.": "May", "มิ.ย.": "Jun", "ก.ค.": "Jul", "ส.ค.": "Aug",
        "ก.ย.": "Sep", "ต.ค.": "Oct", "พ.ย.": "Nov", "ธ.ค.": "Dec"
    }

    for th, en in months.items():
        if th in text:
            text = text.replace(th, en)

    # แปลงปี พ.ศ. → ค.ศ.
    year = re.findall(r"\d{4}", text)
    if year:
        y = int(year[0])
        if y > 2400:
            text = text.replace(str(y), str(y - 543))

    return pd.to_datetime(text, errors='coerce')


# =========================
# 📂 Upload
# =========================
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded_file:

    # =========================
    # เลือก sheet
    # =========================
    excel = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📑 เลือก sheet", excel.sheet_names)

    df = pd.read_excel(excel, sheet_name=sheet, skiprows=5)
    df.columns = df.columns.astype(str).str.strip()

    st.write("📌 Columns:", df.columns.tolist())
    st.dataframe(df.head())

    # =========================
    # หา column พื้นที่
    # =========================
    area_col = None
    for col in df.columns:
        if "พื้นที่" in col:
            area_col = col
            break

    if area_col is None:
        st.error("❌ ไม่พบคอลัมน์ 'พื้นที่'")
        st.stop()

    st.success(f"✅ ใช้คอลัมน์: {area_col}")

    # =========================
    # เตรียมตัวแปร
    # =========================
    current_province = None
    current_month = None
    rows = []

    # =========================
    # loop (เหมือน Python เดิม)
    # =========================
    for _, row in df.iterrows():
        text = str(row[area_col]).strip() if pd.notna(row[area_col]) else ""

        # ข้ามรวม
        if text == "" or "รวม" in text:
            continue

        # ===== เดือน =====
        dt = parse_thai_date(text)
        if pd.notna(dt):
            current_month = dt
            continue

        # ===== จังหวัด =====
        if not any(char.isdigit() for char in text):
            current_province = text
            continue

        # ===== ข้อมูล (เหมือน script เดิม → ไม่ strict)
        new_row = row.copy()
        new_row["จังหวัด"] = current_province
        new_row["วันที่"] = current_month

        rows.append(new_row)

    # =========================
    # debug
    # =========================
    st.write("📊 จำนวน rows:", len(rows))

    if len(rows) == 0:
        st.error("❌ ไม่มีข้อมูลหลังประมวลผล")
        st.stop()

    # =========================
    # รวม dataframe
    # =========================
    df_new = pd.DataFrame(rows)
    df_new = df_new.dropna(how="all")

    st.write("📌 df_new columns:", df_new.columns.tolist())

    # =========================
    # จัดคอลัมน์ (กันพัง)
    # =========================
    expected_cols = ["จังหวัด", "วันที่"]

    cols = [c for c in expected_cols if c in df_new.columns] + \
           [c for c in df_new.columns if c not in expected_cols]

    df_new = df_new[cols]
    df_new = df_new.reset_index(drop=True)

    # =========================
    # แสดงผล
    # =========================
    st.success("✅ ประมวลผลเสร็จแล้ว")
    st.dataframe(df_new.head(20))

    # =========================
    # ดาวน์โหลด
    # =========================
    output = BytesIO()
    df_new.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        label="📥 ดาวน์โหลด Excel",
        data=output,
        file_name="output_cleaned.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

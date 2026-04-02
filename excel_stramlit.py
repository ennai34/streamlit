import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Cleaner", layout="wide")
st.title("📊 Excel Data Cleaner (แยกจังหวัด + วันที่)")

# =========================
# Upload file
# =========================
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded_file:

    # =========================
    # อ่านไฟล์
    # =========================
    df = pd.read_excel(uploaded_file, skiprows=5)
    df.columns = df.columns.astype(str).str.strip()

    st.success("✅ โหลดไฟล์สำเร็จ")
    st.write("🔍 ตัวอย่างข้อมูล")
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
    else:
        st.info(f"📌 ใช้คอลัมน์: {area_col}")

        # =========================
        # เตรียมตัวแปร
        # =========================
        current_province = None
        current_month = None
        rows = []

        # =========================
        # loop
        # =========================
        for _, row in df.iterrows():
            text = str(row[area_col]).strip() if pd.notna(row[area_col]) else ""

            if text == "" or "รวม" in text:
                continue

            # ===== เดือน =====
            try:
                dt = pd.to_datetime(text)
                current_month = dt
                continue
            except:
                pass

            # ===== จังหวัด =====
            if not any(char.isdigit() for char in text):
                current_province = text
                continue

            # ===== ข้อมูล =====
            new_row = row.copy()
            new_row["จังหวัด"] = current_province
            new_row["วันที่"] = current_month

            rows.append(new_row)

        # =========================
        # รวม dataframe
        # =========================
        df_new = pd.DataFrame(rows)
        df_new = df_new.dropna(how="all")

        # =========================
        # จัดคอลัมน์
        # =========================
        cols = ["จังหวัด", "วันที่"] + [c for c in df_new.columns if c not in ["จังหวัด", "วันที่"]]
        df_new = df_new[cols]

        df_new = df_new.reset_index(drop=True)

        st.success("✅ ประมวลผลเสร็จแล้ว")
        st.dataframe(df_new.head())

        # =========================
        # download
        # =========================
        output = BytesIO()
        df_new.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output,
            file_name="output_cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

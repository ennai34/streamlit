import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Cassava Excel Processor", layout="wide")
st.title("📊 Cassava Excel Processor (Advanced)")

# --- 1. อัปโหลดไฟล์ Excel ---
uploaded_file = st.file_uploader("📂 เลือกไฟล์ Excel (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file:
    # --- อ่าน sheet names ---
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_names = xls.sheet_names
    except Exception as e:
        st.error(f"❌ อ่านไฟล์ Excel ไม่สำเร็จ: {e}")
        st.stop()

    # --- เลือก sheet ---
    sheet_name = st.selectbox("📑 เลือก Sheet", sheet_names)

    # --- ระบุจำนวนแถวที่ต้องข้าม ---
    skip_rows = st.number_input("🔢 จำนวนแถวที่ต้องข้ามจากด้านบน", min_value=0, value=8, step=1)

    if st.button("⚡ ประมวลผลไฟล์"):
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=skip_rows, engine="openpyxl")
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่าน Sheet: {e}")
            st.stop()

        # --- เปลี่ยนชื่อคอลัมน์แรกเป็น 'พื้นที่' ---
        df.rename(columns={df.columns[0]: 'พื้นที่'}, inplace=True)

        # --- ฟังก์ชันตรวจสอบวันที่ ---
        def is_date(value):
            if isinstance(value, datetime):
                return True
            try:
                pd.to_datetime(value)
                return True
            except:
                return False

        # --- สร้างคอลัมน์ 'จังหวัด' และ 'เดือน' ---
        df['จังหวัด'] = df['พื้นที่'].apply(lambda x: None if is_date(x) else x)
        df['เดือน'] = df['พื้นที่'].apply(lambda x: x if is_date(x) else None)

        # เติมชื่อจังหวัด
        df['จังหวัด'] = df['จังหวัด'].fillna(method='ffill')

        # ลบคอลัมน์ 'พื้นที่'
        df.drop(columns=['พื้นที่'], inplace=True)

        # --- เลือกคอลัมน์ที่ต้องการเก็บ ---
        all_columns = df.columns.tolist()
        selected_columns = st.multiselect("✅ เลือกคอลัมน์ที่จะเก็บไว้", options=all_columns, default=all_columns)
        df = df[selected_columns]

        # --- แสดง preview และ filter ---
        st.subheader("🔍 Preview ข้อมูล")
        filter_province = st.multiselect("กรองจังหวัด", options=df['จังหวัด'].unique() if 'จังหวัด' in df.columns else [])
        filter_month = st.multiselect("กรองเดือน", options=df['เดือน'].unique() if 'เดือน' in df.columns else [])

        filtered_df = df.copy()
        if 'จังหวัด' in df.columns and filter_province:
            filtered_df = filtered_df[filtered_df['จังหวัด'].isin(filter_province)]
        if 'เดือน' in df.columns and filter_month:
            filtered_df = filtered_df[filtered_df['เดือน'].isin(filter_month)]

        st.dataframe(filtered_df)

        # --- ดาวน์โหลดไฟล์ Excel ---
        output = BytesIO()
        filtered_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ผลลัพธ์",
            data=output,
            file_name="cassava_processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
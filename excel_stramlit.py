import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# --- ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="ASSAVA-CLEAN: Excel Data Processing App", layout="wide")
st.title("📊 Cassava Excel Processor (Advanced)")
st.caption("เวอร์ชันเว็บ: แยกจังหวัด/เดือน และคำนวณผลผลิตอัตโนมัติ")

# --- 1. อัปโหลดไฟล์ Excel ---
uploaded_file = st.file_uploader("📂 เลือกไฟล์ Excel (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_names = xls.sheet_names
        sheet_name = st.selectbox("📑 เลือก Sheet ที่ต้องการประมวลผล", sheet_names)
    except Exception as e:
        st.error(f"❌ อ่านไฟล์ Excel ไม่สำเร็จ: {e}")
        st.stop()

    # --- ระบุจำนวนแถวที่ต้องข้าม ---
    skip_rows = st.number_input("🔢 จำนวนแถวที่ต้องข้ามจากด้านบน", min_value=0, value=5, step=1)

    # --- ปุ่มเริ่มประมวลผล ---
    if st.button("🚀 เริ่มประมวลผล", use_container_width=True):
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=skip_rows, engine="openpyxl")
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่าน Sheet: {e}")
            st.stop()

        # --- เปลี่ยนชื่อคอลัมน์แรกเป็น 'พื้นที่' ---
        df.rename(columns={df.columns[0]: 'พื้นที่'}, inplace=True)

        # --- ฟังก์ชันตรวจสอบว่าเป็นวันที่หรือไม่ ---
        def is_date(value):
            if isinstance(value, datetime):
                return True
            try:
                pd.to_datetime(value)
                return True
            except:
                return False

        # --- สร้างคอลัมน์จังหวัดและเดือน ---
        df['จังหวัด'] = df['พื้นที่'].apply(lambda x: None if is_date(x) else x)
        df['เดือน'] = df['พื้นที่'].apply(lambda x: x if is_date(x) else None)

        # เติมชื่อจังหวัดในแถวเดือน
        df['จังหวัด'] = df['จังหวัด'].fillna(method='ffill')

        # ลบแถวที่ไม่มีเดือน
        df = df.dropna(subset=['เดือน'], how='all')

        # ลบคอลัมน์พื้นที่
        df.drop(columns=['พื้นที่'], inplace=True)

        # --- เพิ่มคอลัมน์ผลผลิต ---
        if 'ผลผลิต' in df.columns:
            df['ผลผลิต_กิโลกรัม'] = pd.to_numeric(df['ผลผลิต'], errors='coerce')
            df['ผลผลิต_ตัน'] = df['ผลผลิต_กิโลกรัม'] / 1000
        else:
            st.warning("⚠️ ไม่พบคอลัมน์ 'ผลผลิต' ในไฟล์ Excel")

        # --- จัดเรียงคอลัมน์ให้จังหวัดอยู่ด้านหน้า ---
        cols = ['จังหวัด', 'เดือน'] + [col for col in df.columns if col not in ['จังหวัด', 'เดือน']]
        df = df[cols]

        # --- แสดงผลลัพธ์ ---
        st.subheader("📈 ตารางข้อมูลหลังประมวลผล")
        st.dataframe(df, use_container_width=True)

        # --- ดาวน์โหลดไฟล์ Excel ---
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ผลลัพธ์",
            data=output,
            file_name=f"cassava_processed_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

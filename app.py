import io
import re
import pandas as pd
import pikepdf
import pdfplumber
import streamlit as st
from pikepdf import PasswordError

# --- 1. ตั้งค่าพื้นฐานของหน้าเว็บ ---
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📑")

# --- 2. ฟังก์ชัน Logic สำหรับอ่านไฟล์ (Engine) ---

def split_channel_and_detail(text):
    channels = [
        "EDC/K SHOP/MYQR", "โอนเข้า/หักบัญชีอัตโนมัติ", "K PLUS", "ตู้เติมเงิน / โมบาย แอปพลิ", 
        "Internet/Mobile KK", "K BIZ", "EDC", "โอนเข้าหักบัญชีอัตโนมัติ", "ATM", "CDM", 
        "BRANCH", "K-Cash Connect Plus" , "Internet/Mobile GSB", "Internet/Mobile SCB", 
        "Internet/Mobile KTB ", "Internet/Mobile TTB", "ตู้เติมเงิน / โมบาย แอปพลิชัน", "Internet/Mobile BAY"
    ]
    found_channel, detail_part = "-", text
    for c in channels:
        if c in text:
            found_channel = c
            detail_part = text.replace(c, "").strip()
            break
    return found_channel, detail_part

def str_to_float(val_str):
    if val_str in [None, "", "-", " "]: return None
    try: return float(str(val_str).replace(',', ''))
    except: return None

def parse_pdf_content(pdf_stream):
    all_rows = []
    date_pattern = re.compile(r'^(\d{2}[-/]\d{2}[-/]\d{2,4})')
    time_pattern = re.compile(r'(\d{2}:\d{2})')
    
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            lines = text.split('\n')
            current_row = None
            for line in lines:
                line = line.strip()
                date_match = date_pattern.match(line)
                if date_match:
                    if current_row: all_rows.append(current_row)
                    date_val = date_match.group(1)
                    time_match = time_pattern.search(line)
                    time_val = time_match.group(1) if time_match else ""
                    amounts = re.findall(r'[\d,]+\.\d{2}', line)
                    desc = line.replace(date_val, "").replace(time_val, "").strip()
                    if amounts: desc = desc.split(amounts[0])[0].strip()
                    val_change, balance = None, None
                    if len(amounts) >= 2:
                        val_change = str_to_float(amounts[0])
                        balance = str_to_float(amounts[-1])
                        if any(k in desc for k in ["ถอน", "โอนออก", "ชำระ", "Fee"]):
                            val_change = -abs(val_change)
                    elif len(amounts) == 1:
                        balance = str_to_float(amounts[0])
                    current_row = [date_val, time_val, desc, val_change, balance, "-", ""]
                elif current_row:
                    current_row[2] += " " + line
            if current_row: all_rows.append(current_row)
    return all_rows

# --- 3. ส่วนของหน้าจอใช้งาน (UI แบบพื้นฐาน) ---

st.title("📑 PDF to Excel Converter")
st.write("อัปโหลดไฟล์ PDF Statement ของคุณเพื่อแปลงเป็นไฟล์ Excel (.xlsx)")

st.divider() # เส้นคั่น

# ช่องอัปโหลดไฟล์
pdf_file = st.file_uploader("1. เลือกไฟล์ PDF", type="pdf")

# ช่องกรอกรหัสผ่าน
password = st.text_input("2. รหัสผ่านไฟล์ PDF (ถ้ามี)", type="password", placeholder="ระบุรหัสผ่านเพื่อปลดล็อกไฟล์")

st.write("") # เว้นวรรค

# ปุ่มกดเริ่มทำงาน
if st.button("เริ่มการแปลงไฟล์", type="primary"):
    if pdf_file:
        try:
            with st.spinner("กำลังประมวลผลไฟล์..."):
                # 1. ปลดล็อกและอ่านไฟล์
                pdf_bytes = pdf_file.read()
                with pikepdf.open(io.BytesIO(pdf_bytes), password=password) as pdf:
                    unlocked_io = io.BytesIO()
                    pdf.save(unlocked_io)
                    unlocked_io.seek(0)
                    
                    # 2. อ่านข้อมูลจาก PDF
                    data_rows = parse_pdf_content(unlocked_io)
                    header = ["วันที่", "เวลา", "รายการ", "ถอนเงิน/ฝากเงิน", "ยอดคงเหลือ", "ช่องทาง", "รายละเอียด"]
                    df = pd.DataFrame(data_rows, columns=header)
                    
                    # 3. สร้างไฟล์ Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Statement')
                    output.seek(0)
                    
                # 4. แสดงผลสำเร็จ
                st.success(f"✅ แปลงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
                
                # แสดงตัวอย่างข้อมูล
                st.dataframe(df, use_container_width=True)
                
                # ปุ่มดาวน์โหลด
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=output,
                    file_name=f"Converted_{pdf_file.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except PasswordError:
            st.error("❌ รหัสผ่านไม่ถูกต้อง")
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
    else:
        st.warning("⚠️ กรุณาเลือกไฟล์ PDF ก่อน")

# ส่วนท้าย
st.divider()
st.caption("พัฒนาโดยใช้ Streamlit | ปลอดภัย ข้อมูลถูกประมวลผลในหน่วยความจำและไม่ถูกบันทึกเก็บไว้")

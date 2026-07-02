import io
import re
import pandas as pd
import pikepdf
import pdfplumber
import streamlit as st
from pikepdf import PasswordError

# --- 1. ตั้งค่าหน้าจอและ CSS สไตล์มืออาชีพ (Green Gradient) ---
st.set_page_config(page_title="STM to Excel", page_icon="📑", layout="centered")

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        * { font-family: 'Kanit', sans-serif; }
        .stApp {
            background-color: #f4f7f6;
            background-image: radial-gradient(circle at top right, #e8f5e9, transparent),
                              radial-gradient(circle at bottom left, #e8f5e9, transparent);
        }
        header, footer { visibility: hidden; }

        /* ส่วนหัวสีเขียวแบบที่คุณชอบ */
        .custom-header {
            background: linear-gradient(135deg, #13803a 0%, #0b5d2a 100%);
            padding: 40px 20px;
            text-align: center;
            color: white;
            border-radius: 24px;
            margin-bottom: 25px;
            box-shadow: 0 10px 20px rgba(0,0,0,0.1);
        }
        .custom-header i { font-size: 3.5rem; margin-bottom: 15px; }

        /* ปรับแต่งปุ่มกด */
        div.stButton > button {
            background: linear-gradient(135deg, #13803a 0%, #0b5d2a 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 12px !important;
            padding: 20px !important;
            width: 100% !important;
            font-size: 1.1rem !important;
            font-weight: 500 !important;
            box-shadow: 0 8px 15px rgba(19, 128, 58, 0.2) !important;
        }

        .step-number {
            background: #13803a;
            color: white;
            width: 24px; height: 24px;
            border-radius: 50%;
            display: inline-flex; align-items: center; justify-content: center;
            font-size: 0.8rem; margin-right: 10px;
        }
    </style>
    """, unsafe_allow_html=True)

# --- 2. Logic การคำนวณและอ่าน PDF (ปรับปรุงให้ฉลาดขึ้น) ---

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
    # Regex สำหรับวันที่และเวลา
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
                if not line or any(kw in line for kw in ["ยอดยกมา", "หน้า", "แผ่นที่"]): continue

                # ตรวจสอบว่าเป็นบรรทัดใหม่ของรายการหรือไม่ (เช็คจากวันที่)
                date_match = date_pattern.match(line)
                if date_match:
                    if current_row: all_rows.append(current_row)
                    
                    date_val = date_match.group(1)
                    time_match = time_pattern.search(line)
                    time_val = time_match.group(1) if time_match else ""
                    
                    # ค้นหาจำนวนเงินทั้งหมดในบรรทัด
                    amounts = re.findall(r'[\d,]+\.\d{2}', line)
                    
                    # แยกคำอธิบาย
                    desc = line.replace(date_val, "").replace(time_val, "").strip()
                    if amounts: desc = desc.split(amounts[0])[0].strip()

                    val_change, balance = None, None
                    if len(amounts) >= 2:
                        val_change = str_to_float(amounts[0])
                        balance = str_to_float(amounts[-1])
                        # แยกฝาก/ถอนเบื้องต้น
                        withdraw_keys = ["ถอน", "โอนออก", "ชำระ", "ATS", "Fee", "หัก"]
                        if any(k in desc for k in withdraw_keys):
                            val_change = -abs(val_change)
                    elif len(amounts) == 1:
                        balance = str_to_float(amounts[0])

                    current_row = [date_val, time_val, desc, val_change, balance, "-", ""]
                
                # ถ้าไม่มีวันที่ แสดงว่าเป็นรายละเอียดต่อจากบรรทัดบน
                elif current_row:
                    if not re.findall(r'[\d,]+\.\d{2}', line):
                        current_row[2] += " " + line # ต่อคำอธิบาย
                    else:
                        chan, det = split_channel_and_detail(line)
                        if chan != "-": current_row[5] = chan
                        current_row[6] += " " + det

            if current_row: all_rows.append(current_row)
            
    return all_rows

# --- 3. ส่วนการแสดงผลบนหน้าเว็บ (UI) ---

st.markdown("""
    <div class="custom-header">
        <i class="fa-solid fa-file-invoice-dollar"></i>
        <h2 style="margin-bottom: 0;">STM to Excel</h2>
        <p style="opacity: 0.8;">แปลงไฟล์ PDF Statement เป็น Excel ทันที</p>
    </div>
    """, unsafe_allow_html=True)

# ส่วน Form รับค่า
with st.container():
    st.markdown('<label><span class="step-number">1</span>เลือกไฟล์ PDF Statement</label>', unsafe_allow_html=True)
    pdf_file = st.file_uploader("", type="pdf", label_visibility="collapsed")
    
    st.markdown('<label><span class="step-number">2</span>รหัสผ่านไฟล์ (ถ้ามี)</label>', unsafe_allow_html=True)
    password = st.text_input("", type="password", placeholder="ระบุรหัสผ่านเพื่อปลดล็อก PDF", label_visibility="collapsed")

    process_btn = st.button("🚀 แปลงไฟล์และดาวน์โหลด Excel")

if process_btn and pdf_file:
    try:
        with st.spinner("⏳ กำลังประมวลผลไฟล์ของคุณ..."):
            # 1. Unlock PDF
            pdf_bytes = pdf_file.read()
            with pikepdf.open(io.BytesIO(pdf_bytes), password=password) as pdf:
                unlocked_io = io.BytesIO()
                pdf.save(unlocked_io)
                unlocked_io.seek(0)
                
                # 2. Parse Data
                data_rows = parse_pdf_content(unlocked_io)
                header = ["วันที่", "เวลา", "รายการ", "ถอนเงิน/ฝากเงิน", "ยอดคงเหลือ", "ช่องทาง", "รายละเอียด"]
                df = pd.DataFrame(data_rows, columns=header)
                df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d-%m-%y', errors='coerce')

                # 3. Create Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                    df.to_excel(writer, index=False, sheet_name='Statement')
                    workbook = writer.book
                    worksheet = writer.sheets['Statement']
                    
                    # ตกแต่งไฟล์ Excel
                    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#13803a', 'font_color': 'white'})
                    num_fmt = workbook.add_format({'num_format': '#,##0.00'})
                    
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                    
                    worksheet.set_column('A:B', 12)
                    worksheet.set_column('C:C', 30)
                    worksheet.set_column('D:E', 15, num_fmt)
                    worksheet.set_column('F:G', 40)
                
                output.seek(0)

            st.balloons()
            st.success(f"✅ แปลงข้อมูลสำเร็จ {len(df)} รายการ")
            
            # ปุ่มดาวน์โหลด
            st.download_button(
                label="📥 คลิกที่นี่เพื่อบันทึกไฟล์ Excel",
                data=output,
                file_name=f"Converted_{pdf_file.name.replace('.pdf', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # แสดงพรีวิว
            with st.expander("🔍 ดูตัวอย่างข้อมูล"):
                st.dataframe(df, use_container_width=True)

    except PasswordError:
        st.error("❌ รหัสผ่านไม่ถูกต้อง")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")

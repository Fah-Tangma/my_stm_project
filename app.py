import io
import re
import pandas as pd
import pikepdf
import pdfplumber
import streamlit as st
from pikepdf import PasswordError

# --- 1. การตั้งค่าหน้าจอและ CSS (เพื่อให้เหมือน HTML ที่คุณให้มา) ---
st.set_page_config(page_title="STM to Excel", page_icon="📑", layout="centered")

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        /* จัดการ Font และ Background */
        * { font-family: 'Kanit', sans-serif; }
        
        .stApp {
            background-color: #f4f7f6;
            background-image: radial-gradient(circle at top right, #e8f5e9, transparent),
                              radial-gradient(circle at bottom left, #e8f5e9, transparent);
        }

        /* ซ่อน Header/Footer ของ Streamlit */
        header, footer { visibility: hidden; }

        /* สร้าง Card Container */
        .main-card {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 24px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.08);
            overflow: hidden;
            margin-top: -50px;
        }

        /* ส่วนหัวสีเขียว */
        .custom-header {
            background: linear-gradient(135deg, #13803a 0%, #0b5d2a 100%);
            padding: 40px 20px;
            text-align: center;
            color: white;
            border-radius: 24px 24px 0 0;
        }
        .custom-header i { font-size: 3.5rem; margin-bottom: 15px; }

        /* ปรับแต่งปุ่ม Convert (ปุ่มหลักของ Streamlit) */
        div.stButton > button {
            background: linear-gradient(135deg, #13803a 0%, #0b5d2a 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 12px !important;
            padding: 25px 20px !important;
            width: 100% !important;
            font-size: 1.1rem !important;
            font-weight: 500 !important;
            box-shadow: 0 8px 15px rgba(19, 128, 58, 0.2) !important;
            transition: all 0.3s ease !important;
        }
        div.stButton > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 12px 20px rgba(19, 128, 58, 0.3) !important;
            filter: brightness(1.1);
        }

        /* วงกลมตัวเลข Step */
        .step-number {
            background: #13803a;
            color: white;
            width: 24px;
            height: 24px;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 0.8rem;
            margin-right: 10px;
        }

        /* ปรับแต่งกล่อง Input */
        .stTextInput input, .stFileUploader section {
            border-radius: 12px !important;
            border: 2px solid #eee !important;
        }
    </style>
    """, unsafe_allow_html=True)

# --- 2. ฟังก์ชัน Logic การประมวลผล (คงเดิม) ---
def split_channel_and_detail(text):
    channels = ["EDC/K SHOP/MYQR", "โอนเข้า/หักบัญชีอัตโนมัติ", "K PLUS", "ตู้เติมเงิน / โมบาย แอปพลิ", "Internet/Mobile KK", "K BIZ", "EDC", "โอนเข้าหักบัญชีอัตโนมัติ", "ATM", "CDM", "BRANCH", "K-Cash Connect Plus", "Internet/Mobile GSB", "Internet/Mobile SCB", "Internet/Mobile KTB", "Internet/Mobile TTB", "ตู้เติมเงิน / โมบาย แอปพลิชัน", "Internet/Mobile BAY"]
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
    all_parsed_rows = []
    bf_keywords = ["ยอดยกมา", "Balance Brought Forward", "Brought Forward"]
    table_headers = ["เวลา/", "วันที่มีผล", "รายการ", "ถอนเงิน", "ฝากเงิน", "ยอดคงเหลือ"]
    current_row = None
    with pdfplumber.open(pdf_stream) as pdf_obj:
        for page in pdf_obj.pages:
            text = page.extract_text()
            if not text: continue
            lines = text.split('\n')
            is_in_table = False 
            for line in lines:
                line = line.strip()
                if not line: continue
                if any(kw in line for kw in table_headers):
                    is_in_table = True
                    continue
                if not is_in_table or any(kw in line for kw in ["Total", "รวมทั้งสิ้น", "จบรายการ"]):
                    is_in_table = False
                    continue
                date_match = re.match(r'^(\d{2}-\d{2}-\d{2})', line)
                if date_match:
                    if current_row: all_parsed_rows.append(current_row)
                    date = date_match.group(1)
                    time_match = re.search(r'(\d{2}:\d{2})', line)
                    time = time_match.group(1) if time_match else ""
                    amounts = re.findall(r'[\d,]+\.\d{2}', line)
                    temp_text = line.replace(date, "", 1).strip()
                    if time: temp_text = temp_text.replace(time, "", 1).strip()
                    desc = temp_text.split(amounts[0])[0].strip() if amounts else temp_text
                    amount_val, balance = None, None
                    if len(amounts) == 1: balance = str_to_float(amounts[0])
                    elif len(amounts) >= 2:
                        is_deposit = any(kw in desc for kw in ["รับเงิน", "คืนเงิน", "ฝาก", "เงินคืน", "Thai QR", "รับโอนเงิน"])
                        val = str_to_float(amounts[0])
                        amount_val = val if is_deposit else -val
                        balance = str_to_float(amounts[-1])
                    remaining = ""
                    if amounts:
                        parts = line.split(amounts[-1])
                        if len(parts) > 1: remaining = parts[-1].strip()
                    chan, det = split_channel_and_detail(remaining)
                    current_row = [date, time, desc, amount_val, balance, chan, det]
                elif is_in_table:
                    if any(x in line for x in ["หน้า", "แผ่นที่", "ยอดคงเหลือ"]): continue
                    c_extra, d_extra = split_channel_and_detail(line)
                    all_parsed_rows.append(["", "", "", None, None, c_extra if c_extra != "-" else "", d_extra])
        if current_row: all_parsed_rows.append(current_row)
    
    final_rows = []
    bf_occurrence, empty_row_buffer = 0, []
    def flush_buffer(buf, target):
        if len(buf) == 1: target.append(buf[0])
    for row in all_parsed_rows:
        desc = str(row[2])
        is_bf = any(kw in desc for kw in bf_keywords)
        if is_bf:
            flush_buffer(empty_row_buffer, final_rows); empty_row_buffer = []
            bf_occurrence += 1
            if bf_occurrence <= 1: final_rows.append(row)
            continue
        if row[3] is not None:
            flush_buffer(empty_row_buffer, final_rows); empty_row_buffer = []
            final_rows.append(row)
        else:
            if row[5] != "-" or row[6] != "": empty_row_buffer.append(row)
    flush_buffer(empty_row_buffer, final_rows)
    return final_rows

# --- 3. ส่วนการแสดงผล (UI) ---

# ส่วนหัว HTML
st.markdown("""
    <div class="custom-header">
        <i class="fa-solid fa-file-invoice-dollar"></i>
        <h2 style="margin-bottom: 0; font-weight: 500;">STM to Excel</h2>
        <p style="opacity: 0.8; font-weight: 300;">อัปโหลดไฟล์ PDF เพื่อแปลงข้อมูลทันที</p>
    </div>
    """, unsafe_allow_html=True)

# กล่อง Body
with st.container():
    st.write("") # เว้นระยะ
    
    # Step 1: Upload
    st.markdown('<label class="form-label"><span class="step-number">1</span>เลือกไฟล์ PDF Statement</label>', unsafe_allow_html=True)
    pdf_file = st.file_uploader("", type="pdf", label_visibility="collapsed")
    
    # Step 2: Password
    st.markdown('<label class="form-label"><span class="step-number">2</span>รหัสผ่านไฟล์ (ถ้ามี)</label>', unsafe_allow_html=True)
    password = st.text_input("", type="password", placeholder="ระบุรหัสผ่าน", label_visibility="collapsed")

    # ปุ่ม Convert
    convert_button = st.button("🪄 แปลงไฟล์และดาวน์โหลด Excel")

    if convert_button and pdf_file:
        try:
            with st.spinner("⏳ กำลังประมวลผลไฟล์ของคุณ..."):
                # 1. ปลดล็อก PDF
                pdf_bytes = pdf_file.read()
                with pikepdf.open(io.BytesIO(pdf_bytes), password=password) as pdf:
                    unlocked_io = io.BytesIO()
                    pdf.save(unlocked_io)
                    unlocked_io.seek(0)
                    
                    # 2. อ่านข้อมูล
                    data_rows = parse_pdf_content(unlocked_io)
                    header = ["วันที่", "เวลา", "รายการ", "ถอนเงิน/ฝากเงิน", "ยอดคงเหลือ", "ช่องทาง", "รายละเอียด"]
                    df = pd.DataFrame(data_rows, columns=header)
                    df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d-%m-%y', errors='coerce')

                    # 3. สร้าง Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='mm/dd/yyyy') as writer:
                        df.to_excel(writer, index=False, sheet_name='Statement')
                        workbook, worksheet = writer.book, writer.sheets['Statement']
                        date_fmt = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})
                        num_fmt = workbook.add_format({'num_format': '#,##0.00'})
                        worksheet.set_column('A:A', 12, date_fmt)
                        worksheet.set_column('D:E', 15, num_fmt)
                        worksheet.set_column('F:G', 40)
                    output.seek(0)

                # แสดงผลสำเร็จและปุ่มดาวน์โหลด
                st.balloons()
                st.success(f"✅ แปลงข้อมูลสำเร็จทั้งหมด {len(df)} รายการ")
                
                st.download_button(
                    label="📥 คลิกเพื่อบันทึกไฟล์ Excel",
                    data=output,
                    file_name=f"Converted_{pdf_file.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except PasswordError:
            st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง")
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")

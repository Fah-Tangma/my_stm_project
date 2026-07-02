import io
import re
import pandas as pd
import pikepdf
import pdfplumber
import streamlit as st
from pikepdf import PasswordError

# ================= 1. Configuration & Styling =================
st.set_page_config(
    page_title="Smart Statement Converter",
    page_icon="🏦",
    layout="wide"
)

# Custom CSS สำหรับตกแต่ง UI
st.markdown("""
    <style>
    /* ปรับแต่งฟอนต์และสีพื้นหลัง */
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Sarabun', sans-serif;
    }
    .main {
        background-color: #f0f2f6;
    }
    /* สไตล์ปุ่ม Download */
    .stDownloadButton button {
        background-color: #00c853 !important;
        color: white !important;
        width: 100%;
        border-radius: 10px;
        height: 50px;
        font-weight: bold;
        font-size: 18px;
        border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    /* สไตล์ปุ่ม Convert */
    .stButton button {
        background-color: #4A90E2 !important;
        color: white !important;
        width: 100%;
        border-radius: 10px;
        font-weight: bold;
    }
    /* ตกแต่งส่วนหัวข้อ */
    .header-text {
        color: #1E3A8A;
        font-weight: 700;
        text-align: center;
        padding-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# ================= 2. Helper Functions =================
def str_to_float(val_str):
    if val_str in [None, "", "-", " "]: return 0.0
    try: return float(str(val_str).replace(',', ''))
    except: return 0.0

def split_channel_and_detail(text):
    channels = ["EDC/K SHOP/MYQR", "K PLUS", "K BIZ", "ATM", "CDM", "BRANCH", "Internet/Mobile"]
    found_channel, detail_part = "-", text
    for c in channels:
        if c in text:
            found_channel = c
            detail_part = text.replace(c, "").strip()
            break
    return found_channel, detail_part

# ================= 3. Bank-Specific Parsers =================

def parse_kbank(pdf_stream):
    """Logic สำหรับกสิกร (KBank) ที่คุณใช้อยู่เดิม"""
    all_parsed_rows = []
    bf_keywords = ["ยอดยกมา", "Balance Brought Forward"]
    table_headers = ["เวลา/", "วันที่มีผล", "รายการ"]
    current_row = None

    with pdfplumber.open(pdf_stream) as pdf_obj:
        for page in pdf_obj.pages:
            text = page.extract_text()
            if not text: continue
            lines = text.split('\n')
            is_in_table = False 

            for line in lines:
                line = line.strip()
                if any(kw in line for kw in table_headers):
                    is_in_table = True; continue
                if not is_in_table or any(kw in line for kw in ["Total", "รวมทั้งสิ้น"]):
                    is_in_table = False; continue

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
                    
                    if len(amounts) >= 1:
                        if len(amounts) == 1:
                            balance = str_to_float(amounts[0])
                        else:
                            is_deposit = any(kw in desc for kw in ["รับเงิน", "คืนเงิน", "ฝาก", "รับโอน"])
                            val = str_to_float(amounts[0])
                            amount_val = val if is_deposit else -val
                            balance = str_to_float(amounts[-1])
                    
                    chan, det = split_channel_and_detail(line.split(amounts[-1])[-1] if amounts else "")
                    current_row = [date, time, desc, amount_val, balance, chan, det]
                elif is_in_table:
                    c_extra, d_extra = split_channel_and_detail(line)
                    all_parsed_rows.append(["", "", "", None, None, c_extra if c_extra != "-" else "", d_extra])
        if current_row: all_parsed_rows.append(current_row)
    return all_parsed_rows

def parse_scb(pdf_stream):
    """โครงสร้างสำหรับไทยพาณิชย์ (เพิ่มในอนาคต)"""
    return [["Coming Soon", "", "ระบบกำลังพัฒนาสำหรับธนาคารนี้", 0, 0, "", ""]]

# ================= 4. Sidebar UI =================

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2830/2830284.png", width=80)
    st.title("Control Panel")
    st.markdown("---")
    
    # ส่วนเลือกธนาคาร
    bank_option = st.selectbox(
        "🏦 1. เลือกธนาคารของไฟล์ PDF",
        ("KBank (กสิกรไทย)", "SCB (ไทยพาณิชย์) - เร็วๆ นี้", "BBL (กรุงเทพ) - เร็วๆ นี้"),
        index=0
    )
    
    pdf_file = st.file_uploader("📂 2. อัปโหลดไฟล์ PDF", type="pdf")
    password = st.text_input("🔑 3. รหัสผ่านไฟล์ (ถ้ามี)", type="password")
    
    st.markdown("---")
    convert_button = st.button("▶️ เริ่มการแปลงข้อมูล")
    
    if pdf_file:
        st.success(f"ไฟล์ที่เลือก: {pdf_file.name}")

# ================= 5. Main Display UI =================

st.markdown("<h1 class='header-text'>📑 Smart Statement to Excel</h1>", unsafe_allow_html=True)

if not pdf_file:
    # หน้าจอเริ่มต้น (Onboarding)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**Step 1:** เลือกธนาคารและอัปโหลดไฟล์ PDF ของคุณที่แถบด้านซ้าย")
    with col2:
        st.info("**Step 2:** ใส่รหัสผ่าน (ถ้ามี) แล้วกดปุ่ม 'เริ่มการแปลงข้อมูล'")
    with col3:
        st.info("**Step 3:** ตรวจสอบตัวอย่างข้อมูลและดาวน์โหลดไฟล์ Excel")
    
    st.image("https://images.unsplash.com/photo-1460925895917-afdab827c52f?auto=format&fit=crop&q=80&w=2426", caption="เปลี่ยนข้อมูล PDF ให้จัดการง่ายด้วย Excel")

if convert_button and pdf_file:
    try:
        with st.status("🚀 กำลังประมวลผลข้อมูล...", expanded=True) as status:
            # ปลดล็อก PDF
            st.write("🔓 กำลังปลดล็อกไฟล์...")
            pdf_bytes = pdf_file.read()
            with pikepdf.open(io.BytesIO(pdf_bytes), password=password) as pdf:
                unlocked_io = io.BytesIO()
                pdf.save(unlocked_io)
                unlocked_io.seek(0)
                
                # เลือก Parser ตามธนาคารที่ผู้ใช้เลือก
                st.write(f"⚙️ กำลังใช้ Logic สำหรับ {bank_option}...")
                if "KBank" in bank_option:
                    data_rows = parse_kbank(unlocked_io)
                else:
                    data_rows = parse_scb(unlocked_io)

                # สร้าง DataFrame
                header = ["วันที่", "เวลา", "รายการ", "ยอดเงิน (บาท)", "ยอดคงเหลือ", "ช่องทาง", "รายละเอียดเพิ่มเติม"]
                df = pd.DataFrame(data_rows, columns=header)
                df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d-%m-%y', errors='coerce')
                
                # คำนวณสรุป
                total_in = df[df['ยอดเงิน (บาท)'] > 0]['ยอดเงิน (บาท)'].sum()
                total_out = df[df['ยอดเงิน (บาท)'] < 0]['ยอดเงิน (บาท)'].sum()

                # สร้างไฟล์ Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                    df.to_excel(writer, index=False, sheet_name='Data')
                    workbook = writer.book
                    worksheet = writer.sheets['Data']
                    # จัดรูปแบบตาราง Excel เล็กน้อย
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column('A:G', 18)
                output.seek(0)

            status.update(label="✅ ดึงข้อมูลสำเร็จ!", state="complete", expanded=False)

        # --- ส่วนแสดงสรุปและดาวน์โหลด ---
        st.balloons()
        
        # แสดง Metrics ยอดรวม
        c1, c2, c3 = st.columns(3)
        c1.metric("จำนวนรายการทั้งหมด", f"{len(df)} รายการ")
        c2.metric("ยอดรวมเงินเข้า", f"{total_in:,.2f} ฿", delta_color="normal")
        c3.metric("ยอดรวมเงินออก", f"{abs(total_out):,.2f} ฿", delta="-")

        st.markdown("---")
        
        # ปุ่มดาวน์โหลดขนาดใหญ่
        st.download_button(
            label="💾 คลิกเพื่อดาวน์โหลดไฟล์ Excel (Download Now)",
            data=output,
            file_name=f"Statement_{bank_option.split()[0]}_{pdf_file.name.replace('.pdf', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("### 📊 ตัวอย่างข้อมูลที่ดึงได้")
        st.dataframe(df, use_container_width=True, height=500)

    except PasswordError:
        st.error("❌ รหัสผ่านไม่ถูกต้อง! กรุณาตรวจสอบรหัสผ่าน (ส่วนใหญ่เป็น วันเดือนปีเกิด หรือ เลขบัตรประชาชน)")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")

# ================= 1. ฟังก์ชันช่วยเหลือ (Utility) =================
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

# ================= 2. Logic การอ่าน PDF (คงเดิมตามของคุณ) =================
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
                    
                    if len(amounts) == 1:
                        balance = str_to_float(amounts[0])
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
    bf_occurrence = 0
    empty_row_buffer = []

    def flush_buffer(buffer_list, target_list):
        if len(buffer_list) == 1:
            target_list.append(buffer_list[0])

    for row in all_parsed_rows:
        desc = str(row[2])
        amount = row[3]
        is_bf = any(kw in desc for kw in bf_keywords)

        if is_bf:
            flush_buffer(empty_row_buffer, final_rows)
            empty_row_buffer = []
            bf_occurrence += 1
            if bf_occurrence <= 1:
                final_rows.append(row)
            continue

        if amount is not None:
            flush_buffer(empty_row_buffer, final_rows)
            empty_row_buffer = []
            final_rows.append(row)
        else:
            if row[5] != "-" or row[6] != "":
                empty_row_buffer.append(row)

    flush_buffer(empty_row_buffer, final_rows)
    return final_rows

# ================= 3. ส่วนการแสดงผล (Streamlit UI) =================
st.title("📑 PDF to Excel Converter")
st.write("อัปโหลดไฟล์ PDF Statement เพื่อแปลงเป็นไฟล์ Excel")

with st.sidebar:
    st.header("การตั้งค่า")
    pdf_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")
    password = st.text_input("รหัสผ่านไฟล์ PDF (ถ้ามี)", type="password")
    convert_button = st.button("เริ่มการแปลงไฟล์")

if convert_button and pdf_file:
    try:
        with st.spinner("กำลังประมวลผล..."):
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
                
                # 3. จัดรูปแบบข้อมูล
                df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d-%m-%y', errors='coerce')

                # 4. สร้าง Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='mm/dd/yyyy') as writer:
                    df.to_excel(writer, index=False, sheet_name='Statement')
                    workbook, worksheet = writer.book, writer.sheets['Statement']
                    
                    date_fmt = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})
                    num_fmt = workbook.add_format({'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'})

                    worksheet.set_column('A:A', 12, date_fmt)
                    worksheet.set_column('D:E', 18, num_fmt)
                    worksheet.set_column('B:B', 10)
                    worksheet.set_column('C:C', 25)
                    worksheet.set_column('F:G', 45)
                
                output.seek(0)
                
                # 5. แสดงผลและปุ่มดาวน์โหลด
                st.success("✅ แปลงไฟล์สำเร็จ!")
                st.dataframe(df.head(20)) # โชว์ตัวอย่าง 20 แถว
                
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=output,
                    file_name=f"{pdf_file.name.split('.')[0]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except PasswordError:
        st.error("❌ รหัสผ่านไม่ถูกต้อง")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
elif convert_button and not pdf_file:
    st.warning("⚠️ กรุณาเลือกไฟล์ PDF ก่อนกดปุ่ม")

import io
import re
import pandas as pd
import pikepdf
import pdfplumber
import streamlit as st
from pikepdf import PasswordError

# 1. ตั้งค่าหน้าเว็บ Streamlit
st.set_page_config(
    page_title="PDF Statement to Excel",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ปรับแต่ง CSS เพื่อความสวยงาม
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #FF4B4B;
        color: white;
    }
    .stDownloadButton>button {
        width: 100%;
        background-color: #28a745;
        color: white;
    }
    .css-1r6slb0 {
        padding: 2rem 1rem;
    }
    </style>
    """, unsafe_allow_html=True)

# ================= 2. ฟังก์ชันช่วยเหลือ (คงเดิม) =================
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

# ================= 3. Logic การอ่าน PDF (คงเดิม) =================
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

    # กรองข้อมูล
    final_rows = []
    bf_occurrence = 0
    empty_row_buffer = []
    def flush_buffer(buffer_list, target_list):
        if len(buffer_list) == 1: target_list.append(buffer_list[0])

    for row in all_parsed_rows:
        desc = str(row[2])
        amount = row[3]
        if any(kw in desc for kw in bf_keywords):
            flush_buffer(empty_row_buffer, final_rows)
            empty_row_buffer = []
            bf_occurrence += 1
            if bf_occurrence <= 1: final_rows.append(row)
            continue
        if amount is not None:
            flush_buffer(empty_row_buffer, final_rows)
            empty_row_buffer = []
            final_rows.append(row)
        else:
            if row[5] != "-" or row[6] != "": empty_row_buffer.append(row)
    flush_buffer(empty_row_buffer, final_rows)
    return final_rows

# ================= 4. ส่วนการแสดงผล (Improved UI) =================

# --- Sidebar ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2306/2306086.png", width=100)
    st.title("Settings")
    st.markdown("---")
    pdf_file = st.file_uploader("1️⃣ อัปโหลดไฟล์ PDF Statement", type="pdf")
    password = st.text_input("2️⃣ รหัสผ่านไฟล์ (ถ้ามี)", type="password", help="ใส่รหัสผ่านหากไฟล์ PDF ของคุณมีการเข้ารหัส")
    st.markdown("---")
    convert_button = st.button("🚀 เริ่มแปลงไฟล์")
    
    st.info("""
    **คำแนะนำ:**
    - รองรับไฟล์ PDF จากธนาคารชั้นนำ
    - ระบบจะแยกรายละเอียดช่องทางให้โดยอัตโนมัติ
    """)

# --- Main Content ---
st.title("📑 PDF Statement to Excel Converter")
st.markdown("เครื่องมือช่วยแปลงไฟล์ Statement จาก PDF เป็น Excel เพื่อการทำบัญชีที่ง่ายขึ้น")

if not pdf_file:
    # แสดงคำแนะนำการใช้งานเมื่อยังไม่ได้อัปโหลดไฟล์
    st.info("กรุณาเลือกไฟล์ PDF ที่แถบด้านซ้ายเพื่อเริ่มต้น")
    cols = st.columns(3)
    cols[0].metric("1. Upload", "เลือกไฟล์")
    cols[1].metric("2. Process", "รอระบบอ่านค่า")
    cols[2].metric("3. Download", "รับไฟล์ Excel")

if convert_button and pdf_file:
    try:
        with st.status("กำลังวิเคราะห์ข้อมูลใน PDF...", expanded=True) as status:
            # 1. ปลดล็อก PDF
            st.write("🔓 กำลังปลดล็อกและอ่านไฟล์...")
            pdf_bytes = pdf_file.read()
            with pikepdf.open(io.BytesIO(pdf_bytes), password=password) as pdf:
                unlocked_io = io.BytesIO()
                pdf.save(unlocked_io)
                unlocked_io.seek(0)
                
                # 2. อ่านข้อมูล
                st.write("🔍 กำลังดึงข้อมูลรายการธุรกรรม...")
                data_rows = parse_pdf_content(unlocked_io)
                header = ["วันที่", "เวลา", "รายการ", "ยอดเงิน", "ยอดคงเหลือ", "ช่องทาง", "รายละเอียด"]
                df = pd.DataFrame(data_rows, columns=header)
                
                # 3. จัดรูปแบบข้อมูล
                df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d-%m-%y', errors='coerce')
                
                # 4. คำนวณยอดเบื้องต้นเพื่อโชว์สรุป
                total_rows = len(df)
                deposits = df[df['ยอดเงิน'] > 0]['ยอดเงิน'].sum()
                withdrawals = df[df['ยอดเงิน'] < 0]['ยอดเงิน'].sum()

                # 5. สร้าง Excel ใน Memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                    df.to_excel(writer, index=False, sheet_name='Statement')
                    workbook, worksheet = writer.book, writer.sheets['Statement']
                    num_fmt = workbook.add_format({'num_format': '#,##0.00'})
                    worksheet.set_column('A:A', 12)
                    worksheet.set_column('D:E', 15, num_fmt)
                    worksheet.set_column('F:G', 40)
                output.seek(0)
                
            status.update(label="✅ ประมวลผลสำเร็จ!", state="complete", expanded=False)

        # --- ส่วนแสดงผลลัพธ์ ---
        st.success("🎉 แปลงข้อมูลสำเร็จเรียบร้อยแล้ว!")
        
        # แสดง Metrics สรุป
        m1, m2, m3 = st.columns(3)
        m1.metric("จำนวนรายการ", f"{total_rows} รายการ")
        m2.metric("ยอดเงินเข้าทั้งหมด", f"{deposits:,.2f} บาท")
        m3.metric("ยอดเงินออกทั้งหมด", f"{abs(withdrawals):,.2f} บาท", delta_color="inverse")

        # ส่วนปุ่มดาวน์โหลด (เด่นชัด)
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (Click here to Download)",
            data=output,
            file_name=f"Converted_{pdf_file.name.split('.')[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # แสดงตัวอย่างตาราง
        st.subheader("👀 ตัวอย่างข้อมูล (Preview 20 rows)")
        st.dataframe(df.head(20), use_container_width=True, height=400)

    except PasswordError:
        st.error("❌ รหัสผ่านไฟล์ PDF ไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดระหว่างประมวลผล: {str(e)}")
        st.info("คำแนะนำ: ตรวจสอบว่าไฟล์ PDF เป็นรูปแบบ Statement มาตรฐานของธนาคารหรือไม่")

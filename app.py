import streamlit as st
from pptx import Presentation
import io
from datetime import datetime
import os

# --- ฟังก์ชันแก้ไขข้อความ (เหมือนเดิม) ---
def replace_text(shape, search_str, replace_str):
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            full_text = "".join([run.text for run in paragraph.runs])
            if search_str in full_text:
                for run in paragraph.runs:
                     if search_str in run.text:
                        run.text = run.text.replace(search_str, replace_str)

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบสร้างรายงานสายตรวจทางน้ำ", layout="wide")
st.title("👮‍♂️ ระบบสร้างรายงานสายตรวจท่าเรือประจำวัน (ส.รน.4)")

# --- ส่วนที่เปลี่ยน: เช็คไฟล์ Template ในระบบอัตโนมัติ ---
template_filename = "Template.pptx" 
# (ต้องมั่นใจว่าชื่อไฟล์ใน GitHub ตรงกับชื่อนี้เป๊ะๆ)

if not os.path.exists(template_filename):
    st.error(f"❌ ไม่พบไฟล์ {template_filename} ในระบบ! กรุณาอัปโหลดไฟล์นี้ขึ้น GitHub")
    st.stop() # หยุดการทำงานถ้าไม่มีไฟล์
else:
    # ถ้ามีไฟล์ ให้แสดงสถานะว่าพร้อมใช้งาน
    st.success("✅ โหลดไฟล์ต้นฉบับ (Template) เรียบร้อยแล้ว พร้อมกรอกข้อมูล")

st.markdown("---")
# --- ส่วนกรอกข้อมูล (เหมือนเดิม) ---
st.subheader("ส่วนหัวกระดาษ")
header_month = st.text_input("หัวข้อเรื่อง)", value="วันที่ 24 พ.ย.68 เวลาประมาณ 09.30 น.")
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("ข้อมูลการตรวจ")
    date_time = st.text_input("วัน เวลา ออกตรวจ", value=f"{datetime.now().strftime('%d/%m/%Y')} 09.30 น.")
    location = st.text_input("สถานที่", value="ท่าเทียบเรือ ต.เจ๊ะเห อ.ตากใบ จ.นราธิวาส")
    type_port = st.text_input("ประเภท", value="ท่าเรือ")
    commander = st.text_input("ผู้ควบคุม", value="พ.ต.ท.จิรายุทธ์ แก้วด้วง สว.ส.รน.4 กก.7 บก.รน.")
    risk_level = st.selectbox("ระดับความเสี่ยง", ["สีเขียว", "สีเหลือง", "สีแดง"])

with col2:
    st.subheader("รายละเอียดเพิ่มเติม")
    vehicle = st.text_input("ตรวจยานพาหนะ", value="เรือข้ามฟาก, แพขนานยนต์ ไทย-มาเลเซีย")
    coordinator = st.text_input("ผู้ติดต่อประสานงาน", value="- ผู้ดูแล")
    coordinates = st.text_input("พิกัด", value="6.235873N, 102.08970241E")
    situation = st.text_area("เส้นทาง/สถานการณ์", value="ทางบก / เหตุการณ์ทั่วไปปกติ")

st.subheader("ภาพประกอบ (4 รูป)")
img_col1, img_col2, img_col3, img_col4 = st.columns(4)
img1 = img_col1.file_uploader("รูปที่ 1", type=['jpg', 'png'])
img2 = img_col2.file_uploader("รูปที่ 2", type=['jpg', 'png'])
img3 = img_col3.file_uploader("รูปที่ 3", type=['jpg', 'png'])
img4 = img_col4.file_uploader("รูปที่ 4", type=['jpg', 'png'])

# ปุ่มกดสร้างรายงาน
if st.button("🚀 สร้างรายงาน PowerPoint"):
    try:
        # --- ส่วนที่เปลี่ยน: โหลดไฟล์จากชื่อไฟล์โดยตรง ---
        prs = Presentation(template_filename) 
        slide = prs.slides[0]

        # 1. แทนที่ข้อความ
        replacements = {
            "{{HEADER_MONTH}}": header_month,
            "{{DATE}}": date_time,
            "{{LOCATION}}": location,
            "{{TYPE}}": type_port,
            "{{COMMANDER}}": commander,
            "{{RISK}}": risk_level,
            "{{VEHICLE}}": vehicle,
            "{{COORD_NAME}}": coordinator,
            "{{GPS}}": coordinates,
            "{{SITUATION}}": situation
        }

        for shape in slide.shapes:
            if shape.has_text_frame:
                for key, val in replacements.items():
                    replace_text(shape, key, val)

        # 2. แทนที่รูปภาพ
        images = [img1, img2, img3, img4]
        img_index = 0
        
        for shape in slide.placeholders:
            if shape.placeholder_format.type == 18:
                if img_index < len(images) and images[img_index] is not None:
                    shape.insert_picture(images[img_index])
                    img_index += 1

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        st.success(f"✅ สร้างรายงานสำเร็จ!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ PowerPoint",
            data=output,
            file_name=f"Marine_Police_Report_{header_month}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")



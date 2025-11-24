import streamlit as st
from pptx import Presentation
import io
from datetime import datetime
import os

# --- ฟังก์ชันแก้ไขข้อความ ---
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

# --- เช็คไฟล์ Template อัตโนมัติ ---
template_filename = "Template.pptx" 

if not os.path.exists(template_filename):
    st.error(f"❌ ไม่พบไฟล์ {template_filename} ในระบบ! กรุณาอัปโหลดไฟล์นี้ขึ้น GitHub หรือวางไว้ในโฟลเดอร์เดียวกัน")
    st.stop()
else:
    st.success("✅ ระบบพร้อมทำงาน (โหลด Template เรียบร้อย)")

st.markdown("---")

# --- ส่วนกรอกข้อมูล ---
st.subheader("ส่วนหัวกระดาษ")
header_month = st.text_input("ระบุเดือน/ปี (เช่น พ.ย.68)", value="พ.ย.68")
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

st.markdown("---")

# --- ส่วนอัปโหลดรูปภาพ (แบบทีเดียว 4 รูป) ---
st.subheader("ภาพประกอบ (เลือกทีเดียว 4 รูป)")
st.info("💡 คำแนะนำ: รูปจะเรียงตามลำดับที่เลือก (ซ้ายบน > ขวาบน > ซ้ายล่าง > ขวาล่าง)")

uploaded_files = st.file_uploader(
    "เลือกรูปภาพ (สูงสุด 4 รูป)", 
    type=['jpg', 'png', 'jpeg'], 
    accept_multiple_files=True
)

# แสดงตัวอย่างรูปที่อัปโหลด
if uploaded_files:
    if len(uploaded_files) > 4:
        st.warning(f"⚠️ คุณเลือกมา {len(uploaded_files)} รูป ระบบจะใช้แค่ 4 รูปแรกเท่านั้น")
        use_files = uploaded_files[:4]
    else:
        use_files = uploaded_files

    # โชว์รูปเรียงกันให้ดู
    cols = st.columns(4)
    for i, img_file in enumerate(use_files):
        with cols[i]:
            st.image(img_file, caption=f"รูปที่ {i+1}", use_container_width=True)

# ปุ่มกดสร้างรายงาน
if st.button("🚀 สร้างรายงาน PowerPoint"):
    if not uploaded_files:
        st.error("กรุณาอัปโหลดรูปภาพอย่างน้อย 1 รูปครับ")
    else:
        try:
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

            # 2. แทนที่รูปภาพ (ใช้ loop จากไฟล์ที่อัปโหลดมา)
            # เรียงลำดับไฟล์ตามที่ User เลือกมา
            # ตัดให้เหลือแค่ 4 รูป (กัน error)
            images_to_insert = uploaded_files[:4]
            img_index = 0
            
            for shape in slide.placeholders:
                # เช็คว่าเป็นช่องใส่รูปภาพหรือไม่ (Type 18 = Picture)
                if shape.placeholder_format.type == 18:
                    if img_index < len(images_to_insert):
                        # ใส่รูป
                        shape.insert_picture(images_to_insert[img_index])
                        img_index += 1

            output = io.BytesIO()
            prs.save(output)
            output.seek(0)

            st.success("✅ สร้างรายงานสำเร็จ!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ PowerPoint",
                data=output,
                file_name=f"Marine_Police_Report_{header_month}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")


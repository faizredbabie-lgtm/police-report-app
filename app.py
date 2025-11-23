import streamlit as st
from pptx import Presentation
import io
from datetime import datetime

# ฟังก์ชันสำหรับแทนที่ข้อความใน Text Box
def replace_text(shape, search_str, replace_str):
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            # รวมข้อความใน paragraph เพื่อเช็คว่ามี keyword ไหม (แก้ปัญหาฟอร์แมตแยกคำ)
            full_text = "".join([run.text for run in paragraph.runs])
            if search_str in full_text:
                # ถ้าเจอ ให้แทนที่ text ของ run แรก และลบ run ที่เหลือใน paragraph นั้นทิ้งเพื่อกันข้อความซ้ำ
                # (วิธีนี้ซับซ้อนแต่แม่นยำกว่าสำหรับการแทนที่คำใน PPT)
                # แต่วิธีพื้นฐานที่ง่ายที่สุดสำหรับกรณีนี้คือ:
                for run in paragraph.runs:
                     if search_str in run.text:
                        run.text = run.text.replace(search_str, replace_str)

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบสร้างรายงานสายตรวจทางน้ำ", layout="wide")
st.title("👮‍♂️ ระบบสร้างรายงานสายตรวจท่าเรือประจำวัน (ส.รน.4)")

# ส่วนอัปโหลดไฟล์ Template
uploaded_template = st.file_uploader("1. อัปโหลดไฟล์ PowerPoint ต้นฉบับ (.pptx)", type="pptx")

if uploaded_template:
    st.markdown("---")
    # --- ส่วนที่เพิ่มมาใหม่: แก้ไขหัวข้อ ---
    st.subheader("ส่วนหัวกระดาษ")
    header_month = st.text_input("หัวข้อเรื่อง)", value="ลงแถวประจำสัปดาห์")
    st.markdown("---")

    # ส่วนกรอกข้อมูลเดิม
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

    # ส่วนอัปโหลดรูปภาพ 4 รูป
    st.subheader("ภาพประกอบ (4 รูป)")
    img_col1, img_col2, img_col3, img_col4 = st.columns(4)
    img1 = img_col1.file_uploader("รูปที่ 1", type=['jpg', 'png'])
    img2 = img_col2.file_uploader("รูปที่ 2", type=['jpg', 'png'])
    img3 = img_col3.file_uploader("รูปที่ 3", type=['jpg', 'png'])
    img4 = img_col4.file_uploader("รูปที่ 4", type=['jpg', 'png'])

    # ปุ่มกดสร้างรายงาน
    if st.button("🚀 สร้างรายงาน PowerPoint"):
        try:
            # โหลดไฟล์ Template
            prs = Presentation(uploaded_template)
            slide = prs.slides[0] # แก้ไขสไลด์หน้าแรก

            # 1. แทนที่ข้อความ (รวมตัวแปรหัวข้อใหม่ {{HEADER_MONTH}})
            replacements = {
                "{{HEADER_MONTH}}": header_month,  # <--- เพิ่มตรงนี้
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

            # วนลูปหาข้อความเพื่อแทนที่
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for key, val in replacements.items():
                        replace_text(shape, key, val)

            # 2. แทนที่รูปภาพ
            images = [img1, img2, img3, img4]
            img_index = 0
            
            for shape in slide.placeholders:
                # เช็คว่าเป็นช่องใส่รูปภาพหรือไม่ (Picture Placeholder)
                if shape.placeholder_format.type == 18:  # 18 คือ Picture
                    if img_index < len(images) and images[img_index] is not None:
                        shape.insert_picture(images[img_index])
                        img_index += 1

            # บันทึกไฟล์ลง Memory เพื่อให้ดาวน์โหลด
            output = io.BytesIO()
            prs.save(output)
            output.seek(0)

            st.success(f"✅ สร้างรายงานประจำเดือน {header_month} สำเร็จ!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ PowerPoint",
                data=output,
                file_name=f"Marine_Police_Report_{header_month}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.info("คำแนะนำ: อย่าลืมแก้ในไฟล์ PowerPoint ตรงหัวข้อให้เป็น {{HEADER_MONTH}} ด้วยนะครับ")